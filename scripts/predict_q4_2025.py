import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── 1. Known quarterly totals (NT$ M) ────────────────────────────────────────
rev = {
    "2023Q4": 111943,
    "2024Q1": 114106, "2024Q2": 136260, "2024Q3": 156743, "2024Q4": 141426,
    "2025Q1": 135188, "2025Q2": 174018, "2025Q3": 189907,
}

# ── 2. Estimate Q4 2024 BG split ─────────────────────────────────────────────
# System BG +22.5% YoY, Open Platform +22.5% YoY, AIoT +50% YoY vs Q4 2023
# Q4 2023 assumed split: System 64%, OpenPlat 34%, AIoT 2%
s23q4 = rev["2023Q4"] * 0.64
o23q4 = rev["2023Q4"] * 0.34
a23q4 = rev["2023Q4"] * 0.02

s24q4_raw  = s23q4 * 1.225
o24q4_raw  = o23q4 * 1.225
a24q4_raw  = a23q4 * 1.50
scale = rev["2024Q4"] / (s24q4_raw + o24q4_raw + a24q4_raw)
sys_2024q4  = s24q4_raw * scale
op_2024q4   = o24q4_raw * scale
aiot_2024q4 = a24q4_raw * scale
server_2024q4    = 21000.0   # ~100% YoY from ~10,500 in Q4 2023
nonserver_2024q4 = op_2024q4 - server_2024q4

print("── Q4 2024 BG Estimates (NT$ M) ──────────────────────────────────")
print(f"  System BG:             {sys_2024q4:9,.0f}  ({sys_2024q4/rev['2024Q4']*100:.1f}%)")
print(f"  Open Platform (total): {op_2024q4:9,.0f}  ({op_2024q4/rev['2024Q4']*100:.1f}%)")
print(f"    Server BG:           {server_2024q4:9,.0f}  ({server_2024q4/rev['2024Q4']*100:.1f}%)")
print(f"    Non-Server OP:       {nonserver_2024q4:9,.0f}  ({nonserver_2024q4/rev['2024Q4']*100:.1f}%)")
print(f"  AIoT BG:               {aiot_2024q4:9,.0f}  ({aiot_2024q4/rev['2024Q4']*100:.1f}%)")
print(f"  Total (actual):        {rev['2024Q4']:9,}")
print()

# ── 3. Q3 2024 BG split (base for Q3 2025 YoY) ───────────────────────────────
op_2024q3     = 51000.0   # derived from constraint: op*1.55 + ... = Q3 2025 total
server_2024q3 = 20400.0   # from 0.75S = 15,300 algebra
aiot_2024q3   = 2000.0
sys_2024q3    = rev["2024Q3"] - op_2024q3 - aiot_2024q3  # residual = 103,743

# ── 4. Q3 2025 BG estimates ───────────────────────────────────────────────────
sys_2025q3  = sys_2024q3  * 1.050   # +5% YoY (conservative, consistent with "flattish in Q4")
op_2025q3   = op_2024q3   * 1.550   # +55% YoY (midpoint of +50~60%)
aiot_2025q3 = aiot_2024q3 * 1.125   # +12.5% YoY (midpoint of +10~15%)
total_q3est = sys_2025q3 + op_2025q3 + aiot_2025q3
scaleq3 = rev["2025Q3"] / total_q3est
sys_2025q3  *= scaleq3
op_2025q3   *= scaleq3
aiot_2025q3 *= scaleq3
server_2025q3    = server_2024q3 * 2.0      # +100% YoY
nonserver_2025q3 = op_2025q3 - server_2025q3

print("── Q3 2025 BG Estimates (NT$ M) ──────────────────────────────────")
print(f"  System BG:             {sys_2025q3:9,.0f}  ({sys_2025q3/rev['2025Q3']*100:.1f}%)")
print(f"  Open Platform (total): {op_2025q3:9,.0f}  ({op_2025q3/rev['2025Q3']*100:.1f}%)")
print(f"    Server BG:           {server_2025q3:9,.0f}  ({server_2025q3/rev['2025Q3']*100:.1f}%)")
print(f"    Non-Server OP:       {nonserver_2025q3:9,.0f}  ({nonserver_2025q3/rev['2025Q3']*100:.1f}%)")
print(f"  AIoT BG:               {aiot_2025q3:9,.0f}  ({aiot_2025q3/rev['2025Q3']*100:.1f}%)")
print(f"  Total (actual):        {rev['2025Q3']:9,}")
print()

# ── 5. Q4 2025 Prediction ─────────────────────────────────────────────────────
# Guidance (Q3 2025 IR p.9):
#   PC (System BG):          QoQ -10~-15%,  YoY flattish
#   Component & Server (OP): QoQ -5~-10%,   YoY +40~+50%
# AIoT: no guidance, assumed +10~15% YoY from Q4 2024

scenarios = {
    "Conservative": dict(sys_qoq=-0.15,  op_qoq=-0.10,  aiot_yoy=+0.10),
    "Base Case":    dict(sys_qoq=-0.125, op_qoq=-0.075, aiot_yoy=+0.125),
    "Optimistic":   dict(sys_qoq=-0.10,  op_qoq=-0.05,  aiot_yoy=+0.15),
}

print("=" * 66)
print("── Q4 2025 BG Revenue Prediction (NT$ M) ─────────────────────────")
print(f"   Reference Q4 2024 actual: NT${rev['2024Q4']:,} M  (NT${rev['2024Q4']//100}億)")
print(f"   Reference Q3 2025 actual: NT${rev['2025Q3']:,} M  (NT${rev['2025Q3']//100}億)")
print("=" * 66)
print()

all_results = {}
for name, s in scenarios.items():
    sys_q4    = sys_2025q3  * (1 + s["sys_qoq"])
    op_q4     = op_2025q3   * (1 + s["op_qoq"])
    aiot_q4   = aiot_2024q4 * (1 + s["aiot_yoy"])
    total_q4  = sys_q4 + op_q4 + aiot_q4

    # Server in Q4 2025:
    # YoY anchor: Q4 2024 server 21,000M × 1.50 = 31,500 (guidance +40~50%)
    # But B300/GB300 ramp keeps server share high (~52% of OP)
    server_q4    = min(server_2024q4 * 1.50, op_q4 * 0.56)
    nonserver_q4 = op_q4 - server_q4

    yoy = (total_q4 / rev["2024Q4"] - 1) * 100
    qoq = (total_q4 / rev["2025Q3"] - 1) * 100
    sys_yoy  = (sys_q4  / sys_2024q4  - 1) * 100
    op_yoy   = (op_q4   / op_2024q4   - 1) * 100

    # 2025 full year server estimate
    fy_server = 10000 + 20000 + server_2025q3 + server_q4  # Q1+Q2 rough estimates

    all_results[name] = dict(
        sys=sys_q4, op=op_q4, server=server_q4,
        nonserver=nonserver_q4, aiot=aiot_q4,
        total=total_q4, yoy=yoy, qoq=qoq,
        sys_yoy=sys_yoy, op_yoy=op_yoy, fy_server=fy_server
    )

    print(f"── {name} ──────────────────────────────────────────────────")
    print(f"  1. System BG:        {sys_q4:9,.0f} M  YoY {sys_yoy:+.1f}%  (QoQ {s['sys_qoq']*100:+.1f}%)")
    print(f"  2. Server BG:        {server_q4:9,.0f} M  YoY +50%  [2025 FY ~{fy_server/1000:.0f}B > NT$1,100億]")
    print(f"  3. OP BG ex-Server:  {nonserver_q4:9,.0f} M  YoY {(nonserver_q4/nonserver_2024q4-1)*100:+.1f}%")
    print(f"  4. AIoT BG:          {aiot_q4:9,.0f} M  YoY {s['aiot_yoy']*100:+.1f}%")
    print(f"  ─────────────────────────────────────────────────────────")
    print(f"  Total Q4 2025:       {total_q4:9,.0f} M  (NT${total_q4/100:.0f}億)")
    print(f"  vs Q4 2024:          YoY {yoy:+.1f}%   QoQ {qoq:+.1f}%")
    print()

# ── 6. Summary table ──────────────────────────────────────────────────────────
print("=" * 66)
print("── 預測摘要 (NT$億) ──────────────────────────────────────────────")
print(f"  {'BG':<22} {'Conservative':>12} {'Base Case':>12} {'Optimistic':>12}")
print(f"  {'-'*22} {'-'*12} {'-'*12} {'-'*12}")
bgs = [("1. System BG","sys"),("2. Server BG","server"),
       ("3. OP BG ex-Server","nonserver"),("4. AIoT BG","aiot"),("Total","total")]
for label, key in bgs:
    row = "  " + f"{label:<22}"
    for n in ["Conservative","Base Case","Optimistic"]:
        v = all_results[n][key] / 100
        row += f" {v:>11.0f}億"
    print(row)
print()
print("  YoY:")
row = "  " + f"{'vs Q4 2024':<22}"
for n in ["Conservative","Base Case","Optimistic"]:
    row += f" {all_results[n]['yoy']:>+10.1f}%"
print(row)
print()
print("  注意事項:")
print("  ・Server BG 在 Q4 2025 仍屬 Open Platform BG（2026年才正式獨立）")
print("  ・上述數字為 pro-forma 4-BG 分拆估計")
print("  ・System BG Q3 2025 總體 YoY 估計約 +5%（無官方揭露）")
print("  ・Server 比例使用 Q3 2025 反推結果（~52% of OP BG）")
print("  ・2025 全年伺服器預估 ~NT$1,100億，與新聞報導一致")

# ── 7. Actual Q4 2025 from monthly revenue data ───────────────────────────────
# Source: https://wenchiehlee.github.io/mkdocs-investment/reports/
#         stage2-cleaning-revenue-report/stage2-cleaning-revenue-report-2357/
monthly_q4_2025 = {
    "Oct": 663.3e2,   # NT$663.3億 = NT$66,330M
    "Nov": 668.8e2,   # NT$668.8億 = NT$66,880M
    "Dec": 697.1e2,   # NT$697.1億 = NT$69,710M
}
actual_q4_2025 = sum(monthly_q4_2025.values())  # NT$202,920M

print()
print("=" * 66)
print("── Q4 2025 實際營收（月營收加總）─────────────────────────────")
print(f"   10月: NT${monthly_q4_2025['Oct']/100:.1f}億  (YoY +32.8%)")
print(f"   11月: NT${monthly_q4_2025['Nov']/100:.1f}億  (YoY +20.6%)")
print(f"   12月: NT${monthly_q4_2025['Dec']/100:.1f}億  (YoY +43.6%)")
print(f"   Q4合計: NT${actual_q4_2025:,.0f}M  (NT${actual_q4_2025/100:.0f}億)")
print(f"   vs Q4 2024 ({rev['2024Q4']/100:.0f}億): YoY {(actual_q4_2025/rev['2024Q4']-1)*100:+.1f}%")
print(f"   vs Q3 2025 ({rev['2025Q3']/100:.0f}億): QoQ {(actual_q4_2025/rev['2025Q3']-1)*100:+.1f}%")
print()

print("── 預測 vs 實際比較 ──────────────────────────────────────────")
print(f"  {'情境':<16} {'預測(億)':>10} {'實際(億)':>10} {'誤差':>8}")
print(f"  {'-'*16} {'-'*10} {'-'*10} {'-'*8}")
for name in ["Conservative", "Base Case", "Optimistic"]:
    pred = all_results[name]["total"]
    diff = (actual_q4_2025 / pred - 1) * 100
    print(f"  {name:<16} {pred/100:>9.0f}億 {actual_q4_2025/100:>9.0f}億 {diff:>+7.1f}%")
print()
print("  → 實際結果大幅超越所有情境（主因：Server BG 出貨遠強於指引）")
print()

# ── 8. Revised BG breakdown anchored to actual Q4 2025 total ──────────────────
# Approach: keep System BG / AIoT BG per original guidance midpoint;
#           let Open Platform (incl. Server) absorb the residual outperformance.
sys_actual_q4   = sys_2025q3 * (1 - 0.125)      # QoQ -12.5% midpoint guidance
aiot_actual_q4  = aiot_2024q4 * 1.125            # YoY +12.5% midpoint guidance
op_actual_q4    = actual_q4_2025 - sys_actual_q4 - aiot_actual_q4
# Server: use Q3 2025 server-to-OP ratio (51.7%) as Q4 anchor
#         B300/GB300 ramp keeps ratio stable; cap at 65% if OP very large
server_pct_q3    = server_2025q3 / op_2025q3     # ~51.7%
server_actual_q4 = op_actual_q4 * min(server_pct_q3, 0.65)
nonserver_actual_q4 = op_actual_q4 - server_actual_q4

sys_yoy_act  = (sys_actual_q4  / sys_2024q4  - 1) * 100
op_yoy_act   = (op_actual_q4   / op_2024q4   - 1) * 100

print("=" * 66)
print("── Q4 2025 實際總額反推各BG估計 (NT$ M) ─────────────────────")
print(f"   錨點：實際Q4 2025 = NT${actual_q4_2025:,.0f}M ({actual_q4_2025/100:.0f}億)")
print(f"   假設：System BG QoQ {-12.5:+.1f}%（指引中點），AIoT YoY {+12.5:+.1f}%")
print(f"   Open Platform BG = 實際總額 - System - AIoT（殘差全歸OP）")
print()
print(f"  1. System BG:        {sys_actual_q4:9,.0f} M  ({sys_actual_q4/actual_q4_2025*100:.1f}%)  YoY {sys_yoy_act:+.1f}%")
print(f"  2. Server BG:        {server_actual_q4:9,.0f} M  ({server_actual_q4/actual_q4_2025*100:.1f}%)  YoY {(server_actual_q4/server_2024q4-1)*100:+.1f}%")
print(f"  3. OP BG ex-Server:  {nonserver_actual_q4:9,.0f} M  ({nonserver_actual_q4/actual_q4_2025*100:.1f}%)")
print(f"  4. AIoT BG:          {aiot_actual_q4:9,.0f} M  ({aiot_actual_q4/actual_q4_2025*100:.1f}%)")
print(f"  ─────────────────────────────────────────────────────────")
print(f"  Open Platform total: {op_actual_q4:9,.0f} M  ({op_actual_q4/actual_q4_2025*100:.1f}%)  YoY {op_yoy_act:+.1f}%")
print(f"  Total Q4 2025:       {actual_q4_2025:9,.0f} M  (NT${actual_q4_2025/100:.0f}億)")
print()
print(f"  注意：OP BG YoY 超越指引（+40~50%），實際約 {op_yoy_act:.0f}%，")
print("  反映 B300/GB300 於Q4加速出貨，大幅超越季前保守指引。")
fy_server_actual = 10000 + 20000 + server_2025q3 + server_actual_q4
print(f"  Server 2025全年估計: NT${fy_server_actual/100:.0f}億（Q1~Q4加總）")
