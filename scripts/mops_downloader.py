#!/usr/bin/env python3
"""Standalone CLI script for MOPS downloader."""

import sys
from pathlib import Path

# Add package to path
package_dir = Path(__file__).parent.parent
sys.path.insert(0, str(package_dir))

from mops_downloader.cli import main

if __name__ == '__main__':
    sys.exit(main())
