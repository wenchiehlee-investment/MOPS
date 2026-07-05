#!/bin/bash
# Runs batch_convert.py --retry-ocr in small chunks, committing and pushing
# after each chunk so progress isn't lost if the job is interrupted partway
# through a long (potentially many-hour) run.
set -uo pipefail
cd "$(dirname "$0")/.."

BATCH_SIZE="${1:-30}"
export PYTHONIOENCODING=utf-8

batch_num=0
while true; do
  batch_num=$((batch_num + 1))
  echo "=================================================================="
  echo "Batch $batch_num (limit=$BATCH_SIZE)  $(date)"
  echo "=================================================================="

  # A prior iteration's merge-conflict may have been left unresolved (e.g.
  # this script's own auto-merge failed). Converting on top of a conflicted
  # tree just wastes OCR calls on commits that can never succeed -- stop
  # here instead of looping blindly.
  if git status --porcelain | grep -q '^UU'; then
    echo "ERROR: unresolved merge conflict from a previous batch -- stopping. Resolve manually (see data/reports/*.csv) and re-run."
    exit 1
  fi

  python scripts/batch_convert.py --retry-ocr --limit "$BATCH_SIZE"
  python scripts/generate_mops_health.py

  git add downloads data/reports/mops_health_summary.csv data/reports/mops_matrix_latest.csv data/reports/batch_convert_failures.log
  if git diff --staged --quiet; then
    echo "Nothing to commit this batch."
  else
    git commit -m "chore(ocr): retry-ocr batch $batch_num ($BATCH_SIZE files)"
    if ! git push; then
      echo "Push rejected (remote advanced) -- fetching and merging, then retrying push."
      git fetch origin
      if git merge origin/main -m "Merge origin/main into ocr-retry progress (batch $batch_num)"; then
        git push || echo "WARNING: push still failed for batch $batch_num after merge -- will retry next batch"
      else
        echo "ERROR: merge conflict for batch $batch_num -- aborting merge and stopping (continuing would waste OCR calls on commits that can't land). Resolve manually and re-run."
        git merge --abort
        exit 1
      fi
    fi
  fi

  ocr_remaining=$(tail -1 data/reports/mops_health_summary.csv | awk -F',' '{print $6}')
  echo "OCR remaining: $ocr_remaining"
  if [ "$ocr_remaining" -le 0 ] 2>/dev/null; then
    echo "All done -- ocr_needed_count is 0."
    break
  fi
done
