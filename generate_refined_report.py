"""
================================================================
一万元理财方案 — 细化版（基于三大赛道 + 具体代码 + 精确时间节点）
生成日期: 2026-07-28
================================================================
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ── 样式常量 ──────────────────────────────────────────────
DARK_BLUE   = "1B3A5C"
MED_BLUE    = "2E6B9E"
LIGHT_BLUE  = "D6E4F0"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F5F5F5"
RED_ALERT   = "C0392B"
GREEN_BG    = "27AE60"
ORANGE_BG   = "E67E22"
DARK_GRAY   = "2C3E50"
SECTOR_AI   = "1A5276"
SECTOR_EN   = "1E8449"
SECTOR_RO   = "935116"

title_font    = Font(name="微软雅黑", size=18, bold=True, color=WHITE)
section_font  = Font(name="微软雅黑", size=14, bold=True, color=DARK_BLUE)
header_font   = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
body_font     = Font(name="微软雅黑", size=10.5, color="333333")
bold_font     = Font(name="微软雅黑", size=10.5, bold=True, color="333333")
warn_font     = Font(name="微软雅黑", size=10, bold=True, color=RED_ALERT)
small_font    = Font(name="微软雅黑", size=9, color="888888")
green_font    = Font(name="微软雅黑", size=10.5, bold=True, color=GREEN_BG)
blue_font     = Font(name="微软雅黑", size=10.5, bold=True, color=MED_BLUE)
fund_font     = Font(name="微软雅黑", size=10, color="555555")

title_fill    = PatternFill("solid", fgColor=DARK_BLUE)
header_fill   = PatternFill("solid", fgColor=MED_BLUE)
light_fill    = PatternFill("solid", fgColor=LIGHT_BLUE)
gray_fill     = PatternFill("solid", fgColor=LIGHT_GRAY)
white_fill    = PatternFill("solid", fgColor=WHITE)
red_fill      = PatternFill("solid", fgColor="FDEDEC")
green_fill    = PatternFill("solid", fgColor="EAFAF1")
dark_fill     = PatternFill("solid", fgColor=DARK_BLUE)

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left   = Alignment(horizontal="left", vertical="center", wrap_text=True)
wrap   = Alignment(horizontal="left", vertical="top", wrap_text=True)


def apply_cell(ws, row, col, value, font=body_font, fill=white_fill,
               alignment=left, border=thin_border):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font; c.fill = fill; c.alignment = alignment
    if border: c.border = border
    return c


def apply_header(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.font = header_font; c.fill = header_fill; c.alignment = center; c.border = thin_border
    return c


def set_col_widths(ws, widths: dict):
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def merge_title(ws, row, text, cols=9):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = title_font; c.fill = title_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    for i in range(1, cols + 1):
        ws.cell(row=row, column=i).fill = title_fill; ws.cell(row=row, column=i).border = Border()


# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ─── Sheet 1: 细化资产配置总览 ─────────────────────────────
ws1 = wb.active
ws1.title = "细化配置总览"
COLS = 10
set_col_widths(ws1, {1: 18, 2: 16, 3: 10, 4: 22, 5: 10, 6: 12, 7: 10, 8: 14, 9: 22, 10: 22})

merge_title(ws1, 1, "一万元理财细化方案 — 三大赛道 + 具体代码 + 时间节点", COLS)
ws1.row_dimensions[1].height = 48

r = 3
info_data = [
    ("委托资金", "¥10,000.00"), ("报告日期", "2026年7月28日"),
    ("投资期限", "短期6个月 / 中长期1-2年"), ("风险偏好", "稳健偏成长 (R2-R4混合)"),
    ("核心策略", "三大赛道进攻(65%) + 稳健防御(35%)"),
]
for i, (k, v) in enumerate(info_data):
    apply_cell(ws1, r + i, 1, k, font=bold_font, fill=light_fill)
    ws1.merge_cells(start_row=r + i, start_column=2, end_row=r + i, end_column=COLS)
    apply_cell(ws1, r + i, 2, v, fill=light_fill)
    ws1.row_dimensions[r + i].height = 24

r += len(info_data) + 1
apply_cell(ws1, r, 1, "▎配置总表", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
r += 1

h = ["序号","资产类别","所属赛道","具体基金/ETF","代码","金额(元)","占比","风险","预期年化","操作策略"]
for i, hv in enumerate(h):
    apply_header(ws1, r, i + 1, hv)
ws1.row_dimensions[r].height = 28

alloc = [
    # 三大赛道 — 进攻端
    ["1","AI算力核心","🥇 AI与半导体","芯片ETF景顺","159560","1,200","12%","R4","20-35%","分批逢低定投"],
    ["2","AI芯片设计","🥇 AI与半导体","科创芯片设计ETF天弘","589070","800","8%","R4","25-40%","回调加仓"],
    ["3","AI全产业链","🥇 AI与半导体","科创创业人工智能ETF景顺","159142","500","5%","R4","18-30%","定投配置"],
    ["4","新能源全链","🥈 新能源与储能","新能源ETF南方","516160","1,000","10%","R3","15-25%","分批定投"],
    ["5","储能电池","🥈 新能源与储能","新能源电池ETF华宝","159071","600","6%","R4","20-35%","高弹性进攻"],
    ["6","动力+储能电池","🥈 新能源与储能","电池ETF鹏华","560410","400","4%","R3","15-25%","均衡配置"],
    ["7","机器人核心","🥉 机器人与智能制造","机器人ETF华夏","562500","900","9%","R4","20-35%","核心长持"],
    ["8","高端装备","🥉 机器人与智能制造","高端制造ETF易方达","562910","400","4%","R3","12-20%","稳健仓位"],
    ["9","机器人场外","🥉 机器人与智能制造","天弘机器人ETF联接C","014881","200","2%","R4","20-35%","场外定投"],
    ["10","宽基底仓","— 宽基指数","沪深300ETF华泰柏瑞","510360","500","5%","R3","8-15%","3批定投"],
    # 防御端
    ["11","安全垫","🛡️ 防御资产","嘉实超短债债券A","012773","800","8%","R2","1.5-3%","一次性买入"],
    ["12","安全垫","🛡️ 防御资产","国泰利享安益短债C","—","700","7%","R2","1.5-2.5%","一次性买入"],
    ["13","稳健增强","🛡️ 防御资产","博时恒乐C","—","600","6%","R2-R3","4-8%","持有6个月+"],
    ["14","稳健增强","🛡️ 防御资产","景顺长城景颐双利C","—","400","4%","R2-R3","4-7%","持有6个月+"],
    ["15","对冲配置","🛡️ 防御资产","华安黄金ETF","518880","500","5%","R3","5-15%","Q3低点布局"],
    ["16","灵活现金","🛡️ 防御资产","余额宝/零钱通","—","500","5%","R1","0.5-1.5%","活期备用"],
]

for i, row_data in enumerate(alloc):
    rr = r + 1 + i
    ws1.row_dimensions[rr].height = 26
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws1, rr, j + 1, val, alignment=center)
        elif j == 3:
            apply_cell(ws1, rr, j + 1, val, font=bold_font)
        elif j == 4:
            apply_cell(ws1, rr, j + 1, val, font=blue_font, alignment=center)
        elif j == 8:
            apply_cell(ws1, rr, j + 1, val, font=green_font, alignment=center)
        else:
            apply_cell(ws1, rr, j + 1, val, alignment=center)

# 汇总
rr = r + 1 + len(alloc)
apply_cell(ws1, rr, 1, "合计", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
apply_cell(ws1, rr, 2, "三大赛道65% + 防御35%", font=Font(name="微软雅黑", size=10, color=WHITE), fill=dark_fill, alignment=center)
apply_cell(ws1, rr, 6, "10,000", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
apply_cell(ws1, rr, 7, "100%", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
for j in [3,4,5,8,9,10]:
    apply_cell(ws1, rr, j, "", fill=dark_fill)

# 赛道占比
rr += 2
apply_cell(ws1, rr, 1, "▎赛道占比汇总", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)
rr += 1

sector_summary = [
    ["🥇 AI与半导体", "¥2,500", "25%", "芯片ETF景顺 + 科创芯片设计 + 科创创业AI", "进攻核心，AI算力+国产替代双主线"],
    ["🥈 新能源与储能", "¥2,000", "20%", "新能源ETF南方 + 电池ETF华宝 + 电池ETF鹏华", "十五五政策驱动，储能高景气"],
    ["🥉 机器人与智能制造", "¥1,500", "15%", "机器人ETF华夏 + 高端制造ETF + 天弘机器人联接", "量产元年，高成长赛道"],
    ["📊 宽基指数", "¥500", "5%", "沪深300ETF华泰柏瑞 510360", "底仓配置，赚市场整体上行"],
    ["🛡️ 防御资产", "¥3,500", "35%", "短债 + 固收+ + 黄金 + 现金", "安全缓冲，对冲波动风险"],
]
for i, row_data in enumerate(sector_summary):
    rr2 = rr + i
    ws1.row_dimensions[rr2].height = 24
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws1, rr2, j + 1, val, font=bold_font)
        elif j == 2:
            apply_cell(ws1, rr2, j + 1, val, font=blue_font, alignment=center)
        else:
            apply_cell(ws1, rr2, j + 1, val)

# 免责
rr = rr + len(sector_summary) + 1
apply_cell(ws1, rr, 1, "⚠️ 免责声明：本报表基于2026年7月28日公开市场数据编制，仅供学习参考，不构成投资建议。市场有风险，投资需谨慎。",
           font=small_font, fill=red_fill, border=None)
ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)


# ─── Sheet 2: 买卖时间节点（核心） ──────────────────────────
ws2 = wb.create_sheet("买卖时间节点表")
COLS2 = 10
set_col_widths(ws2, {1: 5, 2: 20, 3: 10, 4: 28, 5: 28, 6: 26, 7: 26, 8: 24, 9: 24, 10: 28})

merge_title(ws2, 1, "精确买卖时间节点 — 含具体日期、条件、操作金额", COLS2)
ws2.row_dimensions[1].height = 48

r = 3
apply_cell(ws2, r, 1, "按时间线排列，标注每一笔买入/卖出的精确日期和触发条件", font=small_font, fill=light_fill, border=None)
ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS2)
r += 1

h2 = ["序号","基金名称(代码)","赛道","买入日期/触发条件","买入金额","买入方式","目标仓位","卖出日期/触发条件","止盈线","止损线/预案"]
for i, hv in enumerate(h2):
    apply_header(ws2, r, i + 1, hv)
ws2.row_dimensions[r].height = 30

buy_sell = [
    # === 立即执行 (2026年8月第1周) ===
    ["1","嘉实超短债A\n(012773)","防御",
     "2026年8月3日(周一)\n立即一次性买入",
     "¥800","支付宝/天天基金","¥800/¥800",
     "持有至2027年2月到期\n届时根据利率环境决定续持或赎回",
     "持有到期即可\n无需主动止盈","不设止损\n(R2低风险)"],
    ["2","国泰利享安益短债C","防御",
     "2026年8月3日(周一)\n立即一次性买入",
     "¥700","支付宝/天天基金","¥700/¥700",
     "同1号，持有至2027年2月\n30天以上免赎回费",
     "—","不设止损"],
    ["3","博时恒乐C","防御",
     "2026年8月3日(周一)\n一次性买入",
     "¥600","支付宝/天天基金","¥600/¥600",
     "持有≥6个月后评估\n若年化≥8%可赎回转投赛道",
     "年化≥8%赎回","跌破成本5%\n减仓50%"],
    ["4","景顺长城景颐双利C","防御",
     "2026年8月3日(周一)\n一次性买入",
     "¥400","支付宝/天天基金","¥400/¥400",
     "同上，持有≥6个月\n关注季度报告权益仓位变化",
     "年化≥7%赎回","跌破成本5%\n减仓50%"],
    ["5","余额宝/零钱通","防御",
     "2026年8月3日(周一)\n立即转入",
     "¥500","支付宝/微信","¥500/¥500",
     "随时可取，预留极端补仓\n日常不轻易动用",
     "—","不适用"],

    # === 8月分批建仓 ===
    ["6","沪深300ETF华泰柏瑞\n(510360)","宽基",
     "第1批: 8月3日(周一) ¥200\n第2批: 8月17日(周一) ¥150\n第3批: 9月1日(周一) ¥150",
     "¥500","券商APP场内\n或支付宝定投联接","¥500/¥500",
     "① 盈利≥15%止盈1/3(约¥166)\n② 盈利≥25%止盈剩余1/2\n③ 上证突破4500点逐步减仓",
     "盈利≥15%→止盈1/3\n盈利≥25%→止盈1/2","亏损≥10%\n减仓50%"],
    ["7","芯片ETF景顺\n(159560)","AI半导体",
     "第1批: 8月10日 ¥400\n(中报密集披露后)\n第2批: 9月1日 ¥400\n(8月中报利空出尽)\n第3批: 10月10日 ¥400\n(三季报预告窗口前)",
     "¥1,200","券商APP场内","¥1,200/¥1,200",
     "① 盈利≥20%止盈1/3\n② 盈利≥35%止盈1/2\n③ AI资本开支连续2季下滑→全部止盈",
     "盈利≥20%→止盈1/3\n盈利≥35%→止盈1/2","亏损≥15%\n(¥180)止损"],
    ["8","科创芯片设计ETF天弘\n(589070)","AI半导体",
     "第1批: 8月11日 ¥300\n第2批: 等科创板回调≥5% ¥300\n第3批: 10月三季报前 ¥200",
     "¥800","券商APP场内","¥800/¥800",
     "高弹性品种:\n① 盈利≥30%止盈1/2\n② 盈利≥50%全部止盈\n③ 国产芯片份额突破85%\n   分批兑现",
     "盈利≥30%→止盈1/2\n盈利≥50%→全出","亏损≥20%\n(¥160)止损"],
    ["9","科创创业AI ETF景顺\n(159142)","AI半导体",
     "2026年8月起每月定投\n每次¥100，共5次\n(8/9/10/11/12月各¥100)",
     "¥500","券商APP或支付宝\n场外联接定投","¥500/¥500",
     "覆盖面广，长期持有\n2027年7月评估是否续持",
     "盈利≥25%→止盈1/2","亏损≥15%\n减仓50%"],

    # === 新能源 ===
    ["10","新能源ETF南方\n(516160)","新能源储能",
     "第1批: 8月4日 ¥400\n(十五五规划刚落地)\n第2批: 10月20日 ¥350\n(Q3数据验证后)\n第3批: 12月15日 ¥250\n(中央经济工作会议前)",
     "¥1,000","券商APP场内","¥1,000/¥1,000",
     "① 十五五装机超预期→持有\n② 光伏供给出清确认→加仓\n③ 盈利≥50%止盈1/3\n④ 2028年目标装机完成\n   80%时逐步退出",
     "盈利≥50%→止盈1/3\n盈利≥80%→止盈1/2","亏损≥15%\n减仓50%"],
    ["11","新能源电池ETF华宝\n(159071)","新能源储能",
     "第1批: 8月11日 ¥250\n第2批: 9月储能招标旺季 ¥200\n第3批: 11月(美中期选举后\n不确定性消除) ¥150",
     "¥600","券商APP场内","¥600/¥600",
     "高弹性品种:\n① 盈利≥35%止盈1/2\n② 全球储能增速回落\n  →全部止盈\n③ 钠电/固态电池突破\n  →持有加仓",
     "盈利≥35%→止盈1/2","亏损≥18%\n止损"],
    ["12","电池ETF鹏华\n(560410)","新能源储能",
     "8月11日一次性买入 ¥400\n或分2批: 8月+10月各¥200",
     "¥400","券商APP场内","¥400/¥400",
     "覆盖动力+储能+消费电子\n盈利≥20%止盈1/3\n长期持有至2028年",
     "盈利≥20%→止盈1/3","亏损≥15%\n减仓50%"],

    # === 机器人 ===
    ["13","机器人ETF华夏\n(562500)","机器人智造",
     "第1批: 8月4日 ¥300\n(建底仓)\n第2批: 8月20日 ¥350\n(WRC大会催化后)\n第3批: 等Optimus V3量产\n消息确认 ¥250",
     "¥900","券商APP场内","¥900/¥900",
     "① Optimus V3量产确认→加仓\n② 盈利≥40%止盈1/3\n③ 盈利≥60%止盈1/2\n④ 人形机器人年出货\n  突破50万台→逐步退出",
     "盈利≥40%→止盈1/3\n盈利≥60%→止盈1/2","亏损≥20%\n(¥180)止损"],
    ["14","高端制造ETF易方达\n(562910)","机器人智造",
     "8月11日一次性买入 ¥400",
     "¥400","券商APP场内","¥400/¥400",
     "波动较小，长期持有\n盈利≥20%止盈1/2\n作为赛道压舱石",
     "盈利≥20%→止盈1/2","亏损≥12%\n减仓50%"],
    ["15","天弘机器人ETF联接C\n(014881)","机器人智造",
     "8月起每月定投 ¥50\n共4个月(8-11月)\n支付宝即可操作",
     "¥200","支付宝定投\n(最低10元起)","¥200/¥200",
     "场外定投方便\n2027年Q2评估\n若机器人量产确认→继续",
     "盈利≥35%→止盈","亏损≥20%\n止损"],

    # === 黄金 ===
    ["16","华安黄金ETF\n(518880)","防御",
     "第1批: 等国际金价\n<$4,000/oz(约9-10月)\n买入 ¥250\n第2批: 美联储12月议息后\n若鸽派→加仓 ¥250\n注: 若Q3金价未跌破$4,000\n则在10月底直接买入¥500",
     "¥500","券商APP场内\n或博时黄金联接C\n(002611)场外","¥500/¥500",
     "① 金价反弹至$4,400→\n   止盈50%(¥250)\n② 金价反弹至$4,600→\n   全部止盈\n③ Q4末未达目标→\n   持有至2027年Q1",
     "$4,400止盈50%\n$4,600全部止盈","金价跌破$3,700\n全部止损"],
]

for i, row_data in enumerate(buy_sell):
    rr = r + 1 + i
    ws2.row_dimensions[rr].height = 72
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws2, rr, j + 1, val, alignment=center)
        elif j == 1:
            apply_cell(ws2, rr, j + 1, val, font=bold_font)
        elif j == 6:
            apply_cell(ws2, rr, j + 1, val, font=blue_font, alignment=center)
        elif j in (8, 9):
            apply_cell(ws2, rr, j + 1, val, font=warn_font if j == 9 else Font(name="微软雅黑", size=10, bold=True, color=GREEN_BG), alignment=wrap)
        else:
            apply_cell(ws2, rr, j + 1, val, alignment=wrap)


# ─── Sheet 3: 时间线日历 ───────────────────────────────────
ws3 = wb.create_sheet("关键事件日历")
COLS3 = 7
set_col_widths(ws3, {1: 16, 2: 10, 3: 28, 4: 24, 5: 22, 6: 22, 7: 26})

merge_title(ws3, 1, "2026-2027 投资关键事件日历与操作节点", COLS3)
ws3.row_dimensions[1].height = 48

r = 3
h3 = ["时间段","日期","关键事件","影响赛道","操作建议","预期市场反应","加仓/减仓信号"]
for i, hv in enumerate(h3):
    apply_header(ws3, r, i + 1, hv)
ws3.row_dimensions[r].height = 28

calendar = [
    # 8月
    ["2026年8月","8/1","新能源汽车维修国标实施","新能源","观望","—","—"],
    ["2026年8月","8/3","建仓窗口开启：买入全部防御资产\n+ 沪深300第1批 + 机器人底仓","全部","全面建仓","—","🟢 加仓信号"],
    ["2026年8月","8/10-15","中报业绩预告密集披露期\nAI/半导体公司业绩验证","AI半导体","根据业绩决定第2批加仓节奏\n超预期→加速买入\n不及预期→延后至8月底","业绩超预期→板块涨3-5%\n不及预期→回调5-10%","业绩超预期: 🟢加速\n不及预期: 🟡观望"],
    ["2026年8月","8/17","7月国民经济数据发布","宽基/新能源","关注工业增加值、社零数据\n数据好→沪深300第2批","—","数据超预期: 🟢"],
    ["2026年8月","8/19-23","2026世界机器人大会(WRC)","机器人","机器人ETF第2批加仓\n大会催化后买入¥350","新品发布→板块涨5-8%","🟢 WRC催化加仓"],
    ["2026年8月","8/31","中报强制披露截止日","全部","评估持仓公司业绩\n汰弱留强","—","业绩不及预期: 🔴减仓"],

    # 9月
    ["2026年9月","9/1","沪深300第3批定投","宽基","¥150买入","—","🟢定投日"],
    ["2026年9月","9/15","8月经济数据发布","全部","—","—","—"],
    ["2026年9月","9/17","美联储9月议息会议","黄金/AI","若不加息→黄金加仓信号\n若加息→黄金观望、AI承压","不加息→黄金涨3-5%\n加息→科技回调5-8%","不加息: 🟢黄金\n加息: 🔴科技避险"],
    ["2026年9月","9月底","中美元首会晤(预期)","全部","风险偏好提升窗口\n若成果积极→全面加仓","积极→全市场涨3-5%","🟢 若成果超预期"],

    # 10月
    ["2026年10月","10/10","Q3季报预告窗口","AI半导体","芯片ETF第3批 ¥400\n科创芯片第3批 ¥200","—","🟢加仓"],
    ["2026年10月","10/19","Q3 GDP数据公布","全部","GDP>4.5%→加仓宽基\nGDP<4%→增防御","—","GDP>4.5%: 🟢\nGDP<4%: 🔴"],
    ["2026年10月","10/20","新能源ETF第2批","新能源","¥350买入","—","🟢加仓"],
    ["2026年10月","10/29","美联储10月议息","黄金","—","—","—"],
    ["2026年10月","10/31","三季报截止日","全部","评估所有持仓业绩","—","—"],

    # 11月
    ["2026年11月","11/3","美国中期选举","全部","选举结果揭晓后\n不确定性消除→加仓窗口\n电池ETF + 机器人定投","共和党赢→传统能源+\n民主党赢→新能源+","🟢 靴子落地加仓"],
    ["2026年11月","11/15","北交所五周年","—","—","—","—"],
    ["2026年11月","11/18-19","APEC(深圳)","全部","关注中美关系最新进展","—","—"],

    # 12月
    ["2026年12月","12/10","美联储年内最后一次议息","黄金/AI","不加息+鸽派→黄金第2批¥250","鸽派→黄金涨","鸽派: 🟢黄金"],
    ["2026年12月","12/15","新能源ETF第3批","新能源","¥250买入\n(中央经济工作会议前布局)","—","🟢加仓"],
    ["2026年12月","12月中旬","中央经济工作会议","全部","定调2027年经济政策\n若积极→全面加仓\n若稳健→维持仓位","积极定调→市场涨3-5%","积极: 🟢加仓\n稳健: 🟡维持"],
    ["2026年12月","12月下旬","年末资金面偏紧\n机构调仓锁定收益","全部","可能出现阶段性低点\n→定投加仓好时机\n→备好现金等待","年末效应→大盘股占优\n小盘成长回调","🟢 年末低点加仓"],

    # 2027年Q1
    ["2027年Q1","1月","春季躁动备战期\n关注年报预告","AI/机器人","1月若回调→为2月春季\n行情加仓储备弹药","1月通常偏弱","🟡 储备资金"],
    ["2027年Q1","2月","春节行情(全年最佳月份)\n上涨概率76%","AI/机器人/新能源","若持仓已有盈利→加仓\n重点加仓: 机器人+AI\n(小盘成长2月弹性最大)","2月平均涨幅+4.3%\n小盘成长弹性最强","🟢🟢 最佳加仓月"],
    ["2027年Q1","3月","两会(政策窗口)\n年报密集披露","全部","两会政策超预期→加仓\n年报验证业绩→持有\n年报不及预期→减仓","政策驱动行情","两会利好: 🟢"],
    ["2027年Q2","4月","一季报披露\n业绩决断期","全部","业绩验证后调整持仓\n汰弱留强","—","—"],
    ["2027年Q2","6月","半年末(全年最弱月份)\n资金面紧张","全部","控制仓位\n保留现金防御\n为下半年布局做准备","6月通常偏弱\n上涨概率最低","🔴 防御为主"],
    ["2027年7月","7月","投资一周年复盘","全部","全面评估：\n① 各标的收益率\n② 再平衡至目标权重\n③ 决定下半年策略","—","📊 全面复盘"],
    ["2028年7月","7月","投资两周年终评","全部","根据届时市场环境\n决定继续持有/减仓/清仓","—","🏁 终评"],
]

for i, row_data in enumerate(calendar):
    rr = r + 1 + i
    ws3.row_dimensions[rr].height = 50
    for j, val in enumerate(row_data):
        if j == 4:
            apply_cell(ws3, rr, j + 1, val, font=Font(name="微软雅黑", size=10, bold=True, color=DARK_BLUE), alignment=wrap)
        elif j == 6:
            # Color code the signals
            if "🟢" in val:
                apply_cell(ws3, rr, j + 1, val, font=Font(name="微软雅黑", size=9.5, bold=True, color=GREEN_BG), alignment=wrap)
            elif "🔴" in val:
                apply_cell(ws3, rr, j + 1, val, font=Font(name="微软雅黑", size=9.5, bold=True, color=RED_ALERT), alignment=wrap)
            else:
                apply_cell(ws3, rr, j + 1, val, alignment=wrap)
        else:
            apply_cell(ws3, rr, j + 1, val, alignment=wrap)


# ─── Sheet 4: 三大赛道详解 ─────────────────────────────────
ws4 = wb.create_sheet("三大赛道投资逻辑")
COLS4 = 8
set_col_widths(ws4, {1: 16, 2: 26, 3: 26, 4: 26, 5: 24, 6: 24, 7: 22, 8: 24})

merge_title(ws4, 1, "三大赛道核心投资逻辑与标的详解", COLS4)
ws4.row_dimensions[1].height = 48

r = 3
h4 = ["赛道","为什么值得投(5年逻辑)","推荐标的(代码)","标的定位","五年预期CAGR","关键催化剂","主要风险","与一万元方案的关系"]
for i, hv in enumerate(h4):
    apply_header(ws4, r, i + 1, hv)
ws4.row_dimensions[r].height = 30

logic = [
    ["🥇 AI与半导体\n(配置¥2,500/25%)",
     "1. 全球半导体2026冲击9750亿美元(+26%)\n"
     "2. DRAM供需紧张至2028年\n"
     "3. 国产AI芯片市占率15%→76%\n"
     "4. 数字芯片Q1利润+517%\n"
     "5. 云巨头年CAPEX超7000亿美元\n"
     "6. 十五五新质生产力核心",
     "① 芯片ETF景顺(159560)\n   — 数字芯片设计占49.55%\n"
     "② 科创芯片设计ETF天弘(589070)\n   — 100%科创板芯片设计\n"
     "③ 科创创业AI ETF景顺(159142)\n   — 科创板+创业板AI全链",
     "159560: 核心底仓\n589070: 高弹性进攻\n159142: 均衡分散",
     "25%-35%",
     "• 每季巨头CAPEX指引\n• 华为昇腾新品发布\n• 存储芯片价格走势\n• 国产设备验证突破\n• WAIC大会(每年7月)",
     "• AI变现不及预期\n• 估值分位94%偏高\n• 地缘限制先进制程\n• 技术路线切换",
     "进攻核心：\n占成长端的38%\n是本方案的\n第一重仓赛道"],
    ["🥈 新能源与储能\n(配置¥2,000/20%)",
     "1. 十五五规划：2030年风光装机\n   28亿千瓦，储能容量电价落地\n"
     "2. 全球储能出货+88%，大储招标\n   +193%，中美欧三重共振\n"
     "3. AI算电协同写入政府工作报告\n"
     "4. 电网十五五投资4万亿(+40%)\n"
     "5. 光伏供给出清+太空光伏新方向\n"
     "6. 钠电/固态电池产业化提速",
     "① 新能源ETF南方(516160)\n   — 光伏+风电+储能+电网全链\n"
     "② 新能源电池ETF华宝(159071)\n   — 储能电池核心(设备占比80%)\n"
     "③ 电池ETF鹏华(560410)\n   — 动力+储能+消费电子电池",
     "516160: 稳健底仓\n159071: 储能高弹性\n560410: 均衡配置",
     "20%-30%",
     "• 各省储能补贴落地\n• 碳酸锂价格企稳\n• 光伏组件招标价\n• 钠电/固态电池突破\n• 海外储能大单",
     "• 光伏装机短期承压\n• 锂电产能过剩\n• 技术路线替代\n• 海外贸易壁垒",
     "稳健成长：\n占成长端的31%\n政策确定性最高\n波动小于AI"],
    ["🥉 机器人与智能制造\n(配置¥1,500/15%)",
     "1. 2026人形机器人量产元年\n   全球出货1.86万→12万台\n"
     "2. 核心部件国产化率目标80%\n   工信部最高5000万研发补贴\n"
     "3. 特斯拉Optimus V3量产临近\n   比亚迪八月发布人形机器人\n"
     "4. 指数预期增速32%/25%/20%\n   (2026-2028连续三年20%+)\n"
     "5. 1.2万亿科创再贷款覆盖\n   具身智能十五五核心方向",
     "① 机器人ETF华夏(562500)\n   — 规模166亿最大，人形含量70%\n"
     "② 高端制造ETF易方达(562910)\n   — 工业母机+高端装备\n"
     "③ 天弘机器人ETF联接C(014881)\n   — 场外定投首选",
     "562500: 核心长持\n562910: 稳健压舱\n014881: 场外定投",
     "20%-35%",
     "• Optimus V3量产公告\n• 比亚迪机器人发布\n• WRC大会(每年8月)\n• 核心部件国产验证\n• 人形机器人出货数据",
     "• 量产进度delay\n• 技术路线不确定\n• 核心部件良率低\n• 板块累计涨幅已大",
     "高弹性卫星：\n占成长端的23%\n弹性最大、\n波动也最大"],
]

for i, row_data in enumerate(logic):
    rr = r + 1 + i
    ws4.row_dimensions[rr].height = 130
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws4, rr, j + 1, val, font=bold_font, alignment=wrap)
        elif j == 4:
            apply_cell(ws4, rr, j + 1, val, font=green_font, alignment=center)
        else:
            apply_cell(ws4, rr, j + 1, val, alignment=wrap)


# ─── Sheet 5: 风险与应对 ───────────────────────────────────
ws5 = wb.create_sheet("风险应对预案")
COLS5 = 7
set_col_widths(ws5, {1: 18, 2: 22, 3: 20, 4: 24, 5: 26, 6: 22, 7: 26})

merge_title(ws5, 1, "细化风险应对预案 — 针对每只持仓的止损/止盈/调仓", COLS5)
ws5.row_dimensions[1].height = 48

r = 3
h5 = ["持仓标的","最大仓位","最大可亏损","止损触发条件","止损后资金去向","止盈触发条件","止盈后资金去向"]
for i, hv in enumerate(h5):
    apply_header(ws5, r, i + 1, hv)
ws5.row_dimensions[r].height = 30

risk = [
    ["芯片ETF景顺(159560)","¥1,200","¥180(15%)","① 跌破成本15%\n② AI巨头CAPEX连降2季\n③ 存储芯片价格逆转","转入短债或货币基金\n等待半导体周期底部信号\n再分批抄底","① 盈利≥20%→卖¥400\n② 盈利≥35%→再卖¥400\n③ 剩余¥400长期持有","止盈资金：\n50%转固收+\n50%等回调再投"],
    ["科创芯片设计ETF(589070)","¥800","¥160(20%)","① 跌破成本20%\n② 国产芯片份额下降\n③ 科创板系统性下跌","转入沪深300ETF\n降低波动率\n保留AI敞口","① 盈利≥30%→卖¥400\n② 盈利≥50%→清仓","止盈资金：\n转机器人或新能源\n赛道间轮动"],
    ["科创创业AI ETF(159142)","¥500","¥75(15%)","① 跌破成本15%\n② AI应用变现证伪","定投不停，降低成本\n等AI产业回暖","① 盈利≥25%→卖¥250\n② 剩余长期持有","转固收+或短期债券"],
    ["新能源ETF南方(516160)","¥1,000","¥150(15%)","① 跌破成本15%\n② 十五五执行大幅不及预期\n③ 光伏组件持续暴跌","转入电网ETF(防御更强)\n或增配债券","① 盈利≥50%→卖¥333\n② 盈利≥80%→再卖¥333","止盈资金转机器人赛道\n或保留现金等下一轮"],
    ["新能源电池ETF(159071)","¥600","¥108(18%)","① 跌破成本18%\n② 碳酸锂持续暴跌\n③ 储能增速大幅回落","转入新能源ETF南方\n(更分散更稳健)","① 盈利≥35%→卖¥300\n② 钠电突破→持有加仓","止盈资金保留现金\n或追加机器人"],
    ["电池ETF鹏华(560410)","¥400","¥60(15%)","① 跌破成本15%\n② 电池技术路线颠覆","转入电网ETF","① 盈利≥20%→卖¥133\n② 长期持有剩余","转固收+"],
    ["机器人ETF华夏(562500)","¥900","¥180(20%)","① 跌破成本20%\n② Optimus量产delay超6月\n③ 人形机器人重大安全事故","转入高端制造ETF\n(波动更小更稳健)","① 盈利≥40%→卖¥300\n② 盈利≥60%→再卖¥300\n③ 出货达50万台→全出","止盈资金：\n50%转AI赛道\n50%转防御资产"],
    ["高端制造ETF(562910)","¥400","¥48(12%)","① 跌破成本12%\n② PMI连续3月<50\n③ 制造业投资大幅下滑","转入沪深300或债券","① 盈利≥20%→卖¥200\n② 剩余长期持有","转短债或固收+"],
    ["天弘机器人联接C(014881)","¥200","¥40(20%)","① 跌破成本20%\n② 人形机器人产业证伪","停止定投，赎回转货基","① 盈利≥35%→清仓\n② 若量产确认→继续定投","转入AI或新能源"],
    ["沪深300ETF(510360)","¥500","¥50(10%)","① 跌破成本10%\n② 上证跌破3400\n③ 经济硬着陆风险","转入短债避险\n等待市场企稳","① 盈利≥15%→卖¥166\n② 盈利≥25%→卖¥250","止盈资金转入\n三大赛道定投"],
    ["短债基金(012773)","¥800","不设止损","R2低风险，持有到期","—","持有至2027年2月\n到期评估续持或转投","到期后若股市低估\n→转入三大赛道"],
    ["固收+基金(博时+景顺)","¥1,000","¥50(5%)","① 跌破成本5%\n② 基金经理更换\n③ 连续2季跑输基准","转入纯短债避险","① 年化≥8%→赎回\n② 持有满6个月灵活调整","股市低估时转入权益\n股市高估时继续持有"],
    ["黄金ETF(518880)","¥500","¥500(金价<$3,700)","① 国际金价跌破$3,700/oz\n② 美联储超预期连续加息\n③ 美元指数暴涨","全部止损转入短债\n等美联储转鸽再配置","① $4,400→卖¥250\n② $4,600→全卖\n③ Q4末未达→持有至Q1","止盈资金转入\n三大赛道或现金"],
]

for i, row_data in enumerate(risk):
    rr = r + 1 + i
    ws5.row_dimensions[rr].height = 72
    for j, val in enumerate(row_data):
        if j in (1, 2):
            apply_cell(ws5, rr, j + 1, val, font=blue_font, alignment=center)
        elif j == 3:
            apply_cell(ws5, rr, j + 1, val, font=warn_font, alignment=wrap)
        else:
            apply_cell(ws5, rr, j + 1, val, alignment=wrap)


# ── 保存 ──────────────────────────────────────────────────
output_path = r"C:\Users\Lenovo\理财\一万元理财细化方案_三大赛道_20260728.xlsx"
wb.save(output_path)
print(f"Report saved to: {output_path}")
