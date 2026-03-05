import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── 1. Historical quarterly actuals (NT$ M, brand consolidated) ───────────────
rev = {
    "2024Q1": 114_106, "2024Q2": 136_260, "2024Q3": 156_743, "2024Q4": 141_426,
    "2025Q1": 135_188, "2025Q2": 174_018, "2025Q3": 189_907,
}

# ── 2. Monthly MOPS data (NT$ 百萬 M) ─────────────────────────────────────────
# Source: https://wenchiehlee.github.io/mkdocs-investment/reports/
#         stage2-cleaning-revenue-report/stage2-cleaning-revenue-report-2357/
mops = {
    "2024Q4": (499.3 + 554.6 + 485.4) * 100,   # = 153_930M
    "2025Q1": (377.5 + 456.3 + 643.2) * 100,   # = 147_700M
    "2025Q2": (562.2 + 632.2 + 685.7) * 100,   # = 188_010M
    "2025Q3": (548.7 + 628.2 + 825.9) * 100,   # = 200_280M
    "2025Q4": (663.3 + 668.8 + 697.1) * 100,   # = 202_920M
}
# Monthly 2025 (NT$ 百萬)
m25 = {
    "Jan": 377.5e2, "Feb": 456.3e2, "Mar": 643.2e2,
    "Apr": 562.2e2, "May": 632.2e2, "Jun": 685.7e2,
    "Jul": 548.7e2, "Aug": 628.2e2, "Sep": 825.9e2,
    "Oct": 663.3e2, "Nov": 668.8e2, "Dec": 697.1e2,
}
jan26 = 679.4e2   # January 2026 actual monthly MOPS (NT$ M)

# ── 3. MOPS-to-IR ratio calibration ──────────────────────────────────────────
ratios = {q: mops[q] / rev[q] for q in ["2024Q4", "2025Q1", "2025Q2", "2025Q3"]}
avg_ratio = sum(ratios.values()) / len(ratios)
print("── MOPS / IR 品牌合併 校正比率 ─────────────────────────────────")
for q, r in ratios.items():
    print(f"  {q}: {r:.4f}")
print(f"  Average: {avg_ratio:.4f}")
print()

# ── 4. Q4 2025 actual IR estimate ──────────────────────────────────────────────
ir_q4_2025 = mops["2025Q4"] / avg_ratio
rev["2025Q4"] = ir_q4_2025

print("── Q4 2025 實際 IR 季度估計 ──────────────────────────────────")
print(f"  MOPS月加總:  NT${mops['2025Q4']/100:.1f}億")
print(f"  IR估計:      NT${ir_q4_2025/100:.0f}億 = NT${ir_q4_2025:,.0f}M")
print(f"  YoY vs Q4 2024: {(ir_q4_2025/rev['2024Q4']-1)*100:+.1f}%")
print(f"  QoQ vs Q3 2025: {(ir_q4_2025/rev['2025Q3']-1)*100:+.1f}%")
print()

# ── 5. Q1 2026 estimate (3 scenarios, Feb+Mar still unknown) ──────────────────
# Jan 2026 known: +80% YoY (679.4億 vs 377.5億)
# Feb/Mar 2026 estimated via YoY growth rate scenarios
q1_scenarios = {
    "Conservative": dict(feb_yoy=0.45, mar_yoy=0.40),
    "Base Case":    dict(feb_yoy=0.55, mar_yoy=0.50),
    "Optimistic":   dict(feb_yoy=0.65, mar_yoy=0.65),
}
q1_results = {}
print("── Q1 2026 估計（以Jan+80%為錨，Feb/Mar情境假設）────────────")
for name, s in q1_scenarios.items():
    feb26 = m25["Feb"] * (1 + s["feb_yoy"])
    mar26 = m25["Mar"] * (1 + s["mar_yoy"])
    mops_q1 = jan26 + feb26 + mar26
    ir_q1   = mops_q1 / avg_ratio
    yoy     = ir_q1 / rev["2025Q1"] - 1
    qoq     = ir_q1 / rev["2025Q4"] - 1
    q1_results[name] = dict(ir=ir_q1, yoy=yoy, qoq=qoq,
                             feb_yoy=s["feb_yoy"], mar_yoy=s["mar_yoy"])
    print(f"  {name:<14}  Feb YoY {s['feb_yoy']:+.0%}  Mar YoY {s['mar_yoy']:+.0%}"
          f"  → IR ~NT${ir_q1/100:.0f}億 (YoY {yoy:+.1%}  QoQ {qoq:+.1%})")
print()

# ── 6. Q2 2025 BG baseline (from IR YoY hints + total calibration) ────────────
# Q2 2025 IR guidance: System +15~20% YoY, OP +40~50% YoY, AIoT +10~15% YoY
# Q2 2024 split estimated: System ~52%, OP ~46%, AIoT ~2%
sys_q2_24 = rev["2024Q2"] * 0.52
op_q2_24  = rev["2024Q2"] * 0.46
aiot_q2_24= rev["2024Q2"] * 0.02

sys_q2_25  = sys_q2_24  * 1.175  # midpoint +17.5%
op_q2_25   = op_q2_24   * 1.45   # midpoint +45%
aiot_q2_25 = aiot_q2_24 * 1.125  # midpoint +12.5%
# Scale to actual total
scale25 = rev["2025Q2"] / (sys_q2_25 + op_q2_25 + aiot_q2_25)
sys_q2_25  *= scale25
op_q2_25   *= scale25
aiot_q2_25 *= scale25
# Server in Q2 2025: B300/GB300 not yet ramped (started late Q3 2025)
# Estimate from Q3 2024 server ~20,400M + growth: ~22,000M
server_q2_25    = 22_000.0
nonserver_q2_25 = op_q2_25 - server_q2_25

print("── Q2 2025 BG 估計（基期）─────────────────────────────────────")
print(f"  System BG:        {sys_q2_25:9,.0f} M  ({sys_q2_25/rev['2025Q2']*100:.1f}%)")
print(f"  Open Platform:    {op_q2_25:9,.0f} M  ({op_q2_25/rev['2025Q2']*100:.1f}%)")
print(f"    Server:         {server_q2_25:9,.0f} M  (B300/GB300 未量產)")
print(f"    OP ex-Server:   {nonserver_q2_25:9,.0f} M")
print(f"  AIoT BG:          {aiot_q2_25:9,.0f} M  ({aiot_q2_25/rev['2025Q2']*100:.1f}%)")
print(f"  Total Q2 2025:    {rev['2025Q2']:9,} M")
print()

# ── 7. Q4 2025 BG estimates (carry-forward base for QoQ) ─────────────────────
# From predict_q4_2025.py results (anchored to actual total):
op_q3_25     = 78_916.0   # Q3 2025 OP total (est)
server_q3_25 = 40_800.0   # Q3 2025 server (est, 51.7% of OP)
sys_q4_25_mid = 95_152.0  # Q4 2025 System BG (QoQ -12.5% from Q3)
aiot_q4_25   = 3_879.0    # Q4 2025 AIoT
op_q4_25     = ir_q4_2025 - sys_q4_25_mid - aiot_q4_25
server_q4_25 = op_q4_25 * (server_q3_25 / op_q3_25)   # ~51.7% ratio
nonserver_q4_25 = op_q4_25 - server_q4_25

print("── Q4 2025 BG 估計（QoQ 基期）─────────────────────────────────")
print(f"  System BG:        {sys_q4_25_mid:9,.0f} M  ({sys_q4_25_mid/ir_q4_2025*100:.1f}%)")
print(f"  Server BG:        {server_q4_25:9,.0f} M  ({server_q4_25/ir_q4_2025*100:.1f}%)")
print(f"  OP ex-Server:     {nonserver_q4_25:9,.0f} M  ({nonserver_q4_25/ir_q4_2025*100:.1f}%)")
print(f"  AIoT BG:          {aiot_q4_25:9,.0f} M  ({aiot_q4_25/ir_q4_2025*100:.1f}%)")
print(f"  Open Platform:    {op_q4_25:9,.0f} M  ({op_q4_25/ir_q4_2025*100:.1f}%)")
print(f"  Total Q4 2025:    {ir_q4_2025:9,.0f} M")
print()

# ── 8. Q2 2026 Prediction ─────────────────────────────────────────────────────
# Key drivers:
#   System BG:  Q2 is peak gaming season. Tariff front-loading + Copilot+ PC cycle.
#               YoY from Q2 2025: +30~50% (strong but normalising vs Jan +80%)
#   Server BG:  B300/GB300 fully ramped, GB200/NVL72 production expanding.
#               YoY from Q2 2025 server (pre-ramp ~22,000M): +130~180%
#   OP ex-Server: RTX 50/60 product cycle, stable growth.
#               YoY +30~50%
#   AIoT:       Steady +10~20% YoY
#
# Cross-check: Q2/Q1 seasonal QoQ (hist avg: +19~29%)
#   Using Base Q1 2026 → Q2 2026 QoQ +22% midpoint

print("=" * 66)
print("── Q2 2026 BG 營收預測 (NT$ M) ───────────────────────────────")
print(f"   參考 Q2 2025 實際: NT${rev['2025Q2']:,}M  (NT${rev['2025Q2']//100}億)")
print(f"   參考 Q4 2025 估計: NT${ir_q4_2025:,.0f}M  (NT${ir_q4_2025/100:.0f}億)")
print(f"   Jan 2026 月營收:   NT${jan26:,.0f}M  (YoY +80%，實際)")
print("=" * 66)
print()

q2_scenarios = {
    "Conservative": dict(sys_yoy=0.30, server_yoy=1.30, op_ex_yoy=0.30, aiot_yoy=0.10),
    "Base Case":    dict(sys_yoy=0.40, server_yoy=1.55, op_ex_yoy=0.40, aiot_yoy=0.15),
    "Optimistic":   dict(sys_yoy=0.50, server_yoy=1.80, op_ex_yoy=0.50, aiot_yoy=0.20),
}

all_q2 = {}
for name, s in q2_scenarios.items():
    sys_q2  = sys_q2_25   * (1 + s["sys_yoy"])
    srv_q2  = server_q2_25 * (1 + s["server_yoy"])
    nop_q2  = nonserver_q2_25 * (1 + s["op_ex_yoy"])
    aiot_q2 = aiot_q2_25  * (1 + s["aiot_yoy"])
    op_q2   = srv_q2 + nop_q2
    total   = sys_q2 + op_q2 + aiot_q2

    yoy  = (total / rev["2025Q2"] - 1) * 100
    q1_base = q1_results["Base Case"]["ir"]  # QoQ from base Q1 2026
    qoq  = (total / q1_base - 1) * 100
    srv_yoy_pct = s["server_yoy"] * 100
    sys_yoy_pct = s["sys_yoy"] * 100

    all_q2[name] = dict(sys=sys_q2, server=srv_q2, nonserver=nop_q2,
                        op=op_q2, aiot=aiot_q2, total=total,
                        yoy=yoy, qoq=qoq)

    print(f"── {name} ──────────────────────────────────────────────────")
    print(f"  1. System BG:        {sys_q2:9,.0f} M  YoY {sys_yoy_pct:+.0f}%")
    print(f"  2. Server BG:        {srv_q2:9,.0f} M  YoY {srv_yoy_pct:+.0f}%  (B300/GB300全速+GB200量產)")
    print(f"  3. OP BG ex-Server:  {nop_q2:9,.0f} M  YoY {s['op_ex_yoy']*100:+.0f}%")
    print(f"  4. AIoT BG:          {aiot_q2:9,.0f} M  YoY {s['aiot_yoy']*100:+.0f}%")
    print(f"  ─────────────────────────────────────────────────────────")
    print(f"  Open Platform total: {op_q2:9,.0f} M  ({op_q2/total*100:.1f}%)")
    print(f"  Total Q2 2026:       {total:9,.0f} M  (NT${total/100:.0f}億)")
    print(f"  vs Q2 2025: YoY {yoy:+.1f}%   vs Q1 2026 (Base): QoQ {qoq:+.1f}%")
    print()

# ── 9. Summary table ──────────────────────────────────────────────────────────
print("=" * 66)
print("── 預測摘要 (NT$億) ──────────────────────────────────────────────")
print(f"  {'BG':<22} {'Conservative':>12} {'Base Case':>12} {'Optimistic':>12}")
print(f"  {'-'*22} {'-'*12} {'-'*12} {'-'*12}")
bgs = [("1. System BG","sys"),("2. Server BG","server"),
       ("3. OP BG ex-Server","nonserver"),("4. AIoT BG","aiot"),("Total","total")]
for label, key in bgs:
    row = "  " + f"{label:<22}"
    for n in ["Conservative","Base Case","Optimistic"]:
        v = all_q2[n][key] / 100
        row += f" {v:>11.0f}億"
    print(row)
print()
print("  YoY vs Q2 2025:")
row = "  " + f"{'vs Q2 2025':<22}"
for n in ["Conservative","Base Case","Optimistic"]:
    row += f" {all_q2[n]['yoy']:>+10.1f}%"
print(row)
print()

# ── 10. Quarterly revenue trend ───────────────────────────────────────────────
print("── 季度營收趨勢 (NT$億) ───────────────────────────────────────")
trend = [
    ("2024Q1", rev["2024Q1"]),
    ("2024Q2", rev["2024Q2"]),
    ("2024Q3", rev["2024Q3"]),
    ("2024Q4", rev["2024Q4"]),
    ("2025Q1", rev["2025Q1"]),
    ("2025Q2", rev["2025Q2"]),
    ("2025Q3", rev["2025Q3"]),
    ("2025Q4*", ir_q4_2025),
    ("2026Q1 (Base)*", q1_results["Base Case"]["ir"]),
    ("2026Q2 (Base)*", all_q2["Base Case"]["total"]),
]
for label, val in trend:
    bar_len = int(val / 5000)
    marker = "▪" if "*" in label else "█"
    print(f"  {label:<20} NT${val/100:>6.0f}億  {marker*bar_len}")
print()
print("  * 帶 * 為估計值；2025Q4 用月加總÷MOPS/IR比率推算")
print()
print("  注意事項:")
print("  ・Q4 2025 IR 尚未發布，以月營收MOPS加總÷校正比率(1.079)估算")
print("  ・Q1 2026 僅知1月實際(+80%YoY)，2-3月依情境假設")
print("  ・Q2 2026 Server 假設 B300/GB300 維持量產 + GB200 NVL72 加速")
print("  ・最大風險：美國關稅政策若導致Q1/Q2前拉後鬆，Q2 2026實際可能偏低")
print("  ・最大上行：AI需求持續超預期，伺服器交貨期縮短，Q2連續創高")
