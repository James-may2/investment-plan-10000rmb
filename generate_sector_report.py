"""
============================================================
未来五年(2026-2031)最具发展前景的三大赛道 — Excel 报表
生成日期: 2026-07-28
============================================================
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
ACCENT_GOLD = "D4A843"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F5F5F5"
RED_ALERT   = "C0392B"
GREEN_BG    = "27AE60"
ORANGE_BG   = "E67E22"
DARK_GRAY   = "2C3E50"

title_font    = Font(name="微软雅黑", size=18, bold=True, color=WHITE)
subtitle_font = Font(name="微软雅黑", size=12, bold=True, color=MED_BLUE)
header_font   = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
section_font  = Font(name="微软雅黑", size=14, bold=True, color=DARK_BLUE)
sector_font   = Font(name="微软雅黑", size=16, bold=True, color=WHITE)
body_font     = Font(name="微软雅黑", size=10.5, color="333333")
bold_font     = Font(name="微软雅黑", size=10.5, bold=True, color="333333")
rank_font     = Font(name="微软雅黑", size=24, bold=True, color=DARK_BLUE)
warn_font     = Font(name="微软雅黑", size=10, bold=True, color=RED_ALERT)
small_font    = Font(name="微软雅黑", size=9, color="888888")
green_font    = Font(name="微软雅黑", size=10.5, bold=True, color=GREEN_BG)

title_fill    = PatternFill("solid", fgColor=DARK_BLUE)
header_fill   = PatternFill("solid", fgColor=MED_BLUE)
light_fill    = PatternFill("solid", fgColor=LIGHT_BLUE)
gray_fill     = PatternFill("solid", fgColor=LIGHT_GRAY)
gold_fill     = PatternFill("solid", fgColor=ACCENT_GOLD)
white_fill    = PatternFill("solid", fgColor=WHITE)
red_fill      = PatternFill("solid", fgColor="FDEDEC")
green_fill    = PatternFill("solid", fgColor="EAFAF1")
dark_fill     = PatternFill("solid", fgColor=DARK_BLUE)
sector1_fill  = PatternFill("solid", fgColor="1A5276")  # 深蓝
sector2_fill  = PatternFill("solid", fgColor="1E8449")  # 深绿
sector3_fill  = PatternFill("solid", fgColor="935116")  # 深棕

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
    c.font = font
    c.fill = fill
    c.alignment = alignment
    if border:
        c.border = border
    return c


def apply_header(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.font = header_font
    c.fill = header_fill
    c.alignment = center
    c.border = thin_border
    return c


def set_col_widths(ws, widths: dict):
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w


def merge_title(ws, row, text, cols=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = title_font
    c.fill = title_fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    for i in range(1, cols + 1):
        ws.cell(row=row, column=i).fill = title_fill
        ws.cell(row=row, column=i).border = Border()


def sector_banner(ws, row, text, fill, cols=8):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = sector_font
    c.fill = fill
    c.alignment = Alignment(horizontal="center", vertical="center")
    for i in range(1, cols + 1):
        ws.cell(row=row, column=i).fill = fill


# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ─── Sheet 1: 三大赛道总览 ────────────────────────────────
ws1 = wb.active
ws1.title = "三大赛道总览"
COLS = 8
set_col_widths(ws1, {1: 16, 2: 24, 3: 22, 4: 24, 5: 18, 6: 18, 7: 18, 8: 22})

# 标题
merge_title(ws1, 1, "未来五年(2026-2031)最具发展前景的三大赛道", COLS)
ws1.row_dimensions[1].height = 45

r = 3
info_text = [
    "报告日期: 2026-07-28 | 展望周期: 2026-2031年 | 数据来源: 券商研报、行业白皮书、公开市场数据",
]
for text in info_text:
    apply_cell(ws1, r, 1, text, font=small_font, fill=light_fill, border=None)
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
    r += 1

# 排名总览
r += 1
apply_cell(ws1, r, 1, "▎五年赛道排名", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
r += 1

overview_headers = ["排名", "赛道名称", "核心驱动力", "五年预期增速(CAGR)", "政策支持力度", "推荐配置权重", "风险等级", "适合投资者"]
for i, h in enumerate(overview_headers):
    apply_header(ws1, r, i + 1, h)
ws1.row_dimensions[r].height = 28

overview_data = [
    ["🥇 第1名", "AI与半导体",
     "算力军备竞赛 + 国产替代 + 存储超级周期",
     "25%-35%",
     "极强（新质生产力核心）",
     "40%-50%",
     "R4（中高风险）",
     "进取型/成长型"],
    ["🥈 第2名", "新能源与储能",
     "十五五规划 + 算电协同 + 全球能源转型",
     "20%-30%",
     "极强（28亿千瓦装机目标）",
     "25%-35%",
     "R3-R4（中风险）",
     "稳健型/成长型"],
    ["🥉 第3名", "机器人与智能制造",
     "量产元年 + 具身智能 + 核心部件国产化",
     "20%-30%",
     "强（1.2万亿科创再贷款）",
     "15%-25%",
     "R4（中高风险）",
     "进取型投资者"],
]

for i, row_data in enumerate(overview_data):
    rr = r + 1 + i
    ws1.row_dimensions[rr].height = 42
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws1, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 6:
            apply_cell(ws1, rr, j + 1, val, font=warn_font, alignment=center)
        else:
            apply_cell(ws1, rr, j + 1, val, alignment=center)

# 关键判断
rr = r + 1 + len(overview_data) + 1
apply_cell(ws1, rr, 1, "▎为什么是这三个赛道？", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)
rr += 1

key_points = [
    "1. 政策共振：三大赛道均被写入「十五五」规划和新质生产力战略，享受国家最高级别产业政策支持，政策确定性极强。",
    "2. 技术奇点：AI大模型推理需求每6个月翻倍、人形机器人2026年量产元年、固态电池/钠电池产业化加速——技术突破正在集中兑现为商业价值。",
    "3. 资金共识：全球科技巨头2026年资本开支超7,000亿美元押注AI；国家电网十五五投资4万亿；科创再贷款1.2万亿覆盖机器人与半导体。",
    "4. 业绩兑现：AI芯片Q1利润同比+211%、储能出货量同比+88%、工业机器人产量同比+28%——三大赛道已从概念炒作进入业绩驱动阶段。",
    "5. 国产替代窗口：地缘博弈加速半导体/高端装备自主可控，国产AI芯片市占率从15%跃升至76%，这一趋势在未来五年不可逆转且加速深化。",
]
for point in key_points:
    apply_cell(ws1, rr, 1, point, font=body_font, fill=white_fill, border=None)
    ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)
    ws1.row_dimensions[rr].height = 24
    rr += 1

# 免责声明
rr += 1
disclaimer = (
    "⚠️ 免责声明：本报表基于2026年7月公开市场数据及机构研报编制，仅供学习参考，不构成投资建议。"
    "市场有风险，投资需谨慎。过往业绩不代表未来表现。五大机构一致看好不代表一定上涨。"
)
apply_cell(ws1, rr, 1, disclaimer, font=small_font, fill=red_fill, border=None)
ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)
ws1.row_dimensions[rr].height = 32


# ─── Sheet 2: 三大赛道详细分析 ────────────────────────────
ws2 = wb.create_sheet("三大赛道详细分析")
COLS2 = 7
set_col_widths(ws2, {1: 18, 2: 30, 3: 30, 4: 22, 5: 22, 6: 22, 7: 30})

merge_title(ws2, 1, "三大赛道深度分析 — 为什么值得投资？", COLS2)
ws2.row_dimensions[1].height = 45

# ─── 赛道1: AI与半导体 ───
r = 3
sector_banner(ws2, r, "🥇 赛道一：AI与半导体 — 未来五年最强主线", sector1_fill, COLS2)
ws2.row_dimensions[r].height = 36
r += 1

s1_headers = ["分析维度", "核心内容", "关键数据", "确定性", "时间窗口", "催化剂", "重点关注"]
for i, h in enumerate(s1_headers):
    apply_header(ws2, r, i + 1, h)
ws2.row_dimensions[r].height = 28

s1_data = [
    ["产业趋势",
     "AI正从训练走向大规模推理部署，推理算力需求每6个月翻倍。存储芯片进入超级周期，DRAM供需紧张至少持续至2028年。",
     "全球半导体市场2026年冲击9750亿美元（+26%）；云巨头年资本开支超7000亿美元",
     "极高",
     "2026-2030持续",
     "每季财报、WAIC大会、巨头CAPEX指引",
     "算力芯片、HBM存储、先进封装"],
    ["国产替代",
     "华为「韬定律」V2突破无EUV约束实现先进制程跨越。国产AI芯片市占率从15%跃升至76%，从备胎变主力。",
     "数字芯片设计Q1利润同比+517%；国产AI加速卡占比76%（2025年）",
     "极高",
     "2026-2028加速",
     "华为新品发布、国产设备验证突破",
     "国产GPU/昇腾链、半导体设备"],
    ["政策支持",
     "「新质生产力」核心载体；大基金三期规模超预期；「十五五」将AI列为首要突破方向。",
     "科创再贷款1.2万亿；大基金三期注册资本3440亿",
     "极高",
     "长期持续",
     "大基金投资进展、财税优惠政策",
     "被投企业、设备材料龙头"],
    ["业绩兑现",
     "板块利润连续多季度高速增长，已从概念炒作进入业绩驱动阶段。",
     "AI芯片2026Q1利润同比+211%；半导体设备市场增速上调至23.5%",
     "高",
     "2026H2-2027",
     "中报/年报业绩披露",
     "业绩超预期个股"],
    ["估值水平",
     "板块估值分位数处于历史高位(94%-97%)，短期需注意回调风险，但高增速可消化高估值。",
     "半导体ETF PE 114倍，估值分位94.65%",
     "中等（短期偏高）",
     "逢回调布局",
     "市场回调、业绩miss后的黄金坑",
     "定投平滑成本"],
    ["主要风险",
     "研发进度不及预期；AI变现能力存疑；先进制程面临外部地缘限制；估值回调风险。",
     "—",
     "需持续跟踪",
     "—",
     "地缘政治变化、技术路线切换",
     "分散配置、设置止损"],
]

for i, row_data in enumerate(s1_data):
    rr = r + 1 + i
    ws2.row_dimensions[rr].height = 60
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws2, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws2, rr, j + 1, val, font=Font(name="微软雅黑", size=10.5, bold=True, color=GREEN_BG), alignment=center)
        else:
            apply_cell(ws2, rr, j + 1, val, alignment=wrap)

# ─── 赛道2: 新能源与储能 ───
r = r + 1 + len(s1_data) + 1
sector_banner(ws2, r, "🥈 赛道二：新能源与储能 — 算电协同+十五五最强政策驱动", sector2_fill, COLS2)
ws2.row_dimensions[r].height = 36
r += 1

for i, h in enumerate(s1_headers):
    apply_header(ws2, r, i + 1, h)
ws2.row_dimensions[r].height = 28

s2_data = [
    ["产业趋势",
     "「AI的尽头是电力」— 算电协同写入政府工作报告。全球储能出货量同比增长88%，大储全球开花、户储工商储方兴未艾。",
     "2026H1全球储能出货328GWh(+88%)；国内大储EPC招标305GWh(+193%)",
     "极高",
     "2026-2030持续",
     "十五五规划细则落地、储能IRR提升",
     "大储系统、钠电池、构网型储能"],
    ["政策驱动",
     "「十五五」可再生能源规划：2030年风电+光伏装机28亿千瓦以上。储能容量电价机制落地元年，IRR可达6%-10%。",
     "风光装机目标28亿千瓦；电网投资4万亿(+40% vs十四五)；储能补贴165-370元/kW·年",
     "极高",
     "2026-2030",
     "各省电价政策、电网招标",
     "电网设备、储能逆变器、PCS"],
    ["技术变革",
     "固态电池十五五加速推进；钠离子电池产业化提速(宁德2GWh欧洲大单)；光伏N型+太空光伏新技术方向。",
     "钠电池全年产量冲击7GWh；光伏HJT/BC渗透率提升；太空光伏写入规划",
     "高",
     "2027-2030",
     "固态电池量产、钠电成本突破",
     "电池龙头、设备厂商"],
    ["供给出清",
     "光伏强制国标2027年实施，落后产能出清；硅料减产检修，供需缺口收窄；锂电排产加速+涨价共振。",
     "光伏招标价回升至1元/W以上；电解液/六氟磷酸锂率先涨价",
     "高",
     "2026H2-2027",
     "硅料价格企稳、组件涨价",
     "光伏龙头、电解液/隔膜"],
    ["全球需求",
     "北美AI缺电引发电网升级；欧洲户储去库完成+工商储崛起；亚非拉缺电驱动；澳洲73亿澳元、英国150亿英镑补贴。",
     "电力设备出口额同比+36%；变压器出口+25%",
     "极高",
     "2026-2028",
     "海外订单公告、出海进展",
     "电力设备出口、变压器"],
    ["主要风险",
     "光伏国内装机短期承压(H1同比-66%)；锂电产能过剩尾部风险；技术路线切换(固态/钠电替代)；海外贸易壁垒。",
     "—",
     "需持续跟踪",
     "—",
     "光伏需求复苏信号、海外关税变化",
     "优选龙头、分散子行业"],
]

for i, row_data in enumerate(s2_data):
    rr = r + 1 + i
    ws2.row_dimensions[rr].height = 60
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws2, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws2, rr, j + 1, val, font=Font(name="微软雅黑", size=10.5, bold=True, color=GREEN_BG), alignment=center)
        else:
            apply_cell(ws2, rr, j + 1, val, alignment=wrap)

# ─── 赛道3: 机器人与智能制造 ───
r = r + 1 + len(s2_data) + 1
sector_banner(ws2, r, "🥉 赛道三：机器人与智能制造 — 量产元年，从0到1向1到N跨越", sector3_fill, COLS2)
ws2.row_dimensions[r].height = 36
r += 1

for i, h in enumerate(s1_headers):
    apply_header(ws2, r, i + 1, h)
ws2.row_dimensions[r].height = 28

s3_data = [
    ["产业趋势",
     "2026年是人形机器人量产元年，全球出货量从1.86万台跃升至12万台。竞争重心从技术展示转向产能建设与供应链整合。",
     "全球出货量2026年12万台（中国占>50%）；工业机器人产量1-5月同比+28.1%",
     "极高",
     "2026-2028爆发",
     "特斯拉Optimus V3量产、比亚迪机器人发布",
     "减速器、丝杠、传感器、灵巧手"],
    ["国产替代",
     "工信部目标2026年底核心零部件国产化率80%。减速器、伺服电机、传感器等核心部件国产厂商加速验证导入。",
     "核心部件厂商订单同比+70-80%；国产化率目标80%",
     "高",
     "2026-2028",
     "国产零部件通过特斯拉/华为验证",
     "谐波减速器、行星滚柱丝杠、六维力传感器"],
    ["政策+资金",
     "「十五五」具身智能列为未来产业核心；工信部最高5000万元研发补贴；1.2万亿科创再贷款覆盖机器人核心部件。",
     "北京、深圳等10+省市出台机器人专项政策",
     "极高",
     "长期持续",
     "各地产业基金落地、补贴细则",
     "政策受益区域龙头"],
    ["应用场景",
     "从工业制造向医疗、家庭服务、危险环境作业等场景拓展。低空经济+机器人+AI融合——具身智能是大模型终极应用形态。",
     "特斯拉Optimus V3已在工厂实测；宇树机甲完成全球首秀",
     "高",
     "2027-2030",
     "应用场景落地案例、大客户订单",
     "本体厂商、系统集成商"],
    ["业绩预期",
     "人形机器人指数2026-2028年预期增速32%/25%/20%，连续三年保持20%+高增速。",
     "Wind一致预期增速32.27%(2026)、24.97%(2027)、20.31%(2028)",
     "高",
     "2026-2028",
     "每季财报、订单公告",
     "特斯拉链、华为链核心标的"],
    ["主要风险",
     "量产进度不及预期；技术路线不确定(液压vs电驱)；核心部件良率爬坡慢；估值泡沫(板块累计涨幅已大)。",
     "—",
     "需持续跟踪",
     "—",
     "量产里程碑兑现情况",
     "ETF分散个股风险"],
]

for i, row_data in enumerate(s3_data):
    rr = r + 1 + i
    ws2.row_dimensions[rr].height = 60
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws2, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws2, rr, j + 1, val, font=Font(name="微软雅黑", size=10.5, bold=True, color=GREEN_BG), alignment=center)
        else:
            apply_cell(ws2, rr, j + 1, val, alignment=wrap)


# ─── Sheet 3: 投资标的清单 ─────────────────────────────────
ws3 = wb.create_sheet("投资标的清单")
COLS3 = 12
set_col_widths(ws3, {1: 14, 2: 24, 3: 10, 4: 10, 5: 16, 6: 12, 7: 10, 8: 10, 9: 12, 10: 22, 11: 16, 12: 28})

merge_title(ws3, 1, "具体投资标的清单 — 每赛道至少2个推荐项目", COLS3)
ws3.row_dimensions[1].height = 45

# ── 赛道1: AI与半导体 ──
r = 3
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS3)
c = ws3.cell(row=r, column=1, value="🥇 赛道一：AI与半导体 — 推荐 5 个投资标的")
c.font = Font(name="微软雅黑", size=13, bold=True, color=WHITE)
c.fill = sector1_fill
c.alignment = center
for j in range(1, COLS3 + 1):
    ws3.cell(row=r, column=j).fill = sector1_fill
ws3.row_dimensions[r].height = 32
r += 1

fund_headers = [
    "赛道", "基金/ETF名称", "代码", "类型", "跟踪指数", "规模",
    "管理费", "风险等级", "近1年表现", "投资理由", "适合策略", "关注要点"
]
for i, h in enumerate(fund_headers):
    apply_header(ws3, r, i + 1, h)
ws3.row_dimensions[r].height = 28

fund_data_s1 = [
    ["AI与半导体",
     "芯片ETF景顺", "159560", "场内ETF",
     "中证芯片产业指数",
     "—", "0.50%", "R4",
     "数字芯片设计占比49.55%，精准卡位AI算力芯片核心标的；国产替代+AI双击逻辑",
     "核心底仓(占赛道40%)，定投买入",
     "半导体周期、国产替代进展、AI资本开支"],
    ["AI与半导体",
     "科创芯片设计ETF天弘", "589070", "场内ETF",
     "上证科创板芯片设计指数",
     "20.62亿", "0.50%", "R4",
     "100%科创板芯片设计公司，同类规模第一；前十大权重65.3%，锐度高",
     "卫星仓位(占赛道25%)，进攻性强",
     "科创板波动大，注意仓位控制"],
    ["AI与半导体",
     "科创创业人工智能ETF景顺", "159142", "场内ETF",
     "中证科创创业人工智能指数",
     "—", "0.50%", "R4",
     "科创板+创业板双板覆盖，半导体占39.5%+AI应用+机器人；赛道覆盖面广",
     "均衡配置(占赛道20%)，一键布局AI全链",
     "双创板块估值波动，适合长期持有"],
    ["AI与半导体",
     "科创半导体设备ETF鹏华", "589020", "场内ETF",
     "上证科创板半导体材料设备指数",
     "—", "0.50%", "R4",
     "聚焦半导体设备和材料，前十大权重71.12%；半导体设备处于超级周期早期",
     "卫星仓位(占赛道10%)，国产设备链",
     "设备订单、国产验证进展"],
    ["AI与半导体",
     "天弘中证人工智能主题A", "011839", "场外基金",
     "中证人工智能主题指数",
     "—", "1.00%", "R4",
     "场外投资者首选，无需证券账户，支付宝/天天基金可买；覆盖AI全产业链",
     "场外定投(占赛道5%)，每月定投¥200-500",
     "主动基金费率较高，关注基金经理稳定性"],
]

for i, row_data in enumerate(fund_data_s1):
    rr = r + 1 + i
    ws3.row_dimensions[rr].height = 52
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws3, rr, j + 1, val, alignment=center)
        elif j == 9:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=wrap)
        else:
            apply_cell(ws3, rr, j + 1, val, alignment=wrap)

# ── 赛道2: 新能源与储能 ──
r = r + 1 + len(fund_data_s1) + 1
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS3)
c = ws3.cell(row=r, column=1, value="🥈 赛道二：新能源与储能 — 推荐 5 个投资标的")
c.font = Font(name="微软雅黑", size=13, bold=True, color=WHITE)
c.fill = sector2_fill
c.alignment = center
for j in range(1, COLS3 + 1):
    ws3.cell(row=r, column=j).fill = sector2_fill
ws3.row_dimensions[r].height = 32
r += 1

for i, h in enumerate(fund_headers):
    apply_header(ws3, r, i + 1, h)
ws3.row_dimensions[r].height = 28

fund_data_s2 = [
    ["新能源与储能",
     "新能源ETF南方", "516160", "场内ETF",
     "中证新能源指数(399808)",
     "—", "0.50%", "R3",
     "覆盖光伏+风电+储能+电网设备全产业链；一键布局十五五新能源28亿千瓦装机目标",
     "核心底仓(占赛道40%)，分批买入",
     "十五五规划落地、风光装机数据、电价政策"],
    ["新能源与储能",
     "新能源电池ETF华宝", "159071", "场内ETF",
     "国证新能源电池指数",
     "—", "0.50%", "R4",
     "储能电池核心标的，电力设备占比超80%；储能出货量同比+88%，全球大储高景气",
     "进攻仓位(占赛道30%)，储能高弹性",
     "储能招标/中标数据、碳酸锂价格、海外订单"],
    ["新能源与储能",
     "电池ETF鹏华", "560410", "场内ETF",
     "中证电池主题指数",
     "—", "0.50%", "R3",
     "动力电池+储能电池+消费电子电池全覆盖；钠电池+固态电池新技术催化密集",
     "均衡配置(占赛道15%)，电池全链条",
     "钠电/固态电池突破、电池排产数据"],
    ["新能源与储能",
     "光伏ETF国泰", "159864", "场内ETF",
     "中证光伏产业指数",
     "—", "0.50%", "R4",
     "光伏全产业链（硅料/硅片/电池/组件）；供给出清+太空光伏双主线，困境反转逻辑",
     "卫星仓位(占赛道10%)，博弈周期反转",
     "硅料价格、产能出清进度、组件招标价"],
    ["新能源与储能",
     "电网ETF", "561380", "场内ETF",
     "中证电网设备主题指数",
     "—", "0.50%", "R3",
     "受益十五五电网4万亿投资(+40%)；电力设备出口+36%，变压器+高压开关出海景气度高",
     "稳健仓位(占赛道5%)，电网投资确定性高",
     "国网招标、海外订单、特高压进展"],
]

for i, row_data in enumerate(fund_data_s2):
    rr = r + 1 + i
    ws3.row_dimensions[rr].height = 52
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws3, rr, j + 1, val, alignment=center)
        elif j == 9:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=wrap)
        else:
            apply_cell(ws3, rr, j + 1, val, alignment=wrap)

# ── 赛道3: 机器人与智能制造 ──
r = r + 1 + len(fund_data_s2) + 1
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS3)
c = ws3.cell(row=r, column=1, value="🥉 赛道三：机器人与智能制造 — 推荐 4 个投资标的")
c.font = Font(name="微软雅黑", size=13, bold=True, color=WHITE)
c.fill = sector3_fill
c.alignment = center
for j in range(1, COLS3 + 1):
    ws3.cell(row=r, column=j).fill = sector3_fill
ws3.row_dimensions[r].height = 32
r += 1

for i, h in enumerate(fund_headers):
    apply_header(ws3, r, i + 1, h)
ws3.row_dimensions[r].height = 28

fund_data_s3 = [
    ["机器人与智能制造",
     "机器人ETF华夏", "562500", "场内ETF",
     "中证机器人指数(H30590)",
     "166.48亿", "0.50%", "R4",
     "规模最大、流动性最好的机器人ETF，人形机器人含量近70%；2026量产元年核心配置工具",
     "核心底仓(占赛道50%)，逢低定投",
     "特斯拉Optimus进展、比亚迪机器人发布"],
    ["机器人与智能制造",
     "机器人ETF广发", "159050", "场内ETF",
     "国证机器人产业指数(980022)",
     "—", "0.50%", "R4",
     "最新份额创成立以来新高；估值处于近2年相对低位，性价比突出",
     "卫星仓位(占赛道25%)，与华夏互补",
     "指数成分差异、折溢价率"],
    ["机器人与智能制造",
     "高端制造ETF易方达", "562910", "场内ETF",
     "装备产业指数(H11054)",
     "—", "0.50%", "R3",
     "覆盖机器人工控+工业母机+高端装备，人形机器人与高端制造双轮驱动",
     "均衡仓位(占赛道15%)，波动相对小",
     "PMI、工业机器人产量、制造业投资"],
    ["机器人与智能制造",
     "天弘中证机器人ETF联接C", "014881", "场外基金",
     "中证机器人指数(H30590)",
     "—", "0.50%", "R4",
     "场外投资者首选，C类免申购费；支付宝/天天基金可定投，最低10元起投",
     "场外定投(占赛道10%)，每月定投¥200-300",
     "联接基金跟踪误差、赎回费规则"],
]

for i, row_data in enumerate(fund_data_s3):
    rr = r + 1 + i
    ws3.row_dimensions[rr].height = 52
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 3:
            apply_cell(ws3, rr, j + 1, val, alignment=center)
        elif j == 9:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=wrap)
        else:
            apply_cell(ws3, rr, j + 1, val, alignment=wrap)


# ─── Sheet 4: 五年资金配置方案 ──────────────────────────────
ws4 = wb.create_sheet("五年资金配置方案")
COLS4 = 8
set_col_widths(ws4, {1: 16, 2: 18, 3: 14, 4: 14, 5: 20, 6: 20, 7: 18, 8: 28})

merge_title(ws4, 1, "五年投资规划 — 假设本金 ¥10,000 的三赛道配置", COLS4)
ws4.row_dimensions[1].height = 45

r = 3
apply_cell(ws4, r, 1, "▎整体配置框架", font=section_font, fill=white_fill, border=None)
ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS4)
r += 1

alloc_headers = ["赛道", "配置金额", "占比", "投资标的", "买入策略", "持有周期", "预期年化收益", "退出条件"]
for i, h in enumerate(alloc_headers):
    apply_header(ws4, r, i + 1, h)
ws4.row_dimensions[r].height = 28

alloc_data = [
    ["AI与半导体", "¥4,000", "40%",
     "芯片ETF景顺(159560) ¥1,500\n科创芯片设计ETF天弘(589070) ¥1,000\n科创创业AI ETF景顺(159142) ¥800\n科创半导体设备ETF鹏华(589020) ¥400\n天弘人工智能主题A(011839) ¥300",
     "分4批定投，每次回调3-5%加仓；首批2026年8月建仓40%，后续每季度定投15%",
     "3-5年\n（核心长持）",
     "20%-35%",
     "① 行业景气度出现拐点信号（如AI资本开支连续2季下滑）\n② 标的盈利≥80%可止盈50%\n③ 跌破成本20%止损"],
    ["新能源与储能", "¥3,500", "35%",
     "新能源ETF南方(516160) ¥1,400\n新能源电池ETF华宝(159071) ¥1,000\n电池ETF鹏华(560410) ¥500\n光伏ETF国泰(159864) ¥400\n电网ETF(561380) ¥200",
     "分3批买入，2026年8月建仓30%，Q4加仓30%，2027年Q1加仓40%",
     "2-5年\n（中期持有+\n动态调整）",
     "15%-25%",
     "① 十五五政策落地不及预期\n② 碳酸锂价格持续暴跌\n③ 盈利≥50%可止盈1/3\n④ 储能增速大幅放缓时减仓"],
    ["机器人与智能制造", "¥2,500", "25%",
     "机器人ETF华夏(562500) ¥1,200\n机器人ETF广发(159050) ¥600\n高端制造ETF易方达(562910) ¥400\n天弘机器人ETF联接C(014881) ¥300",
     "2026年8月建仓30%，等待Optimus V3量产节点加仓40%，2027年H1视出货数据加仓30%",
     "3-5年\n（高成长长持）",
     "20%-35%",
     "① 人形机器人量产严重低于预期\n② 特斯拉/Optimus链重大负面\n③ 核心标的盈利≥100%止盈50%\n④ 跌破成本25%止损"],
]

for i, row_data in enumerate(alloc_data):
    rr = r + 1 + i
    ws4.row_dimensions[rr].height = 100
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws4, rr, j + 1, val, font=bold_font, alignment=center)
        elif j == 2:
            apply_cell(ws4, rr, j + 1, val, font=Font(name="微软雅黑", size=10.5, bold=True, color=MED_BLUE), alignment=center)
        elif j == 6:
            apply_cell(ws4, rr, j + 1, val, font=green_font, alignment=center)
        else:
            apply_cell(ws4, rr, j + 1, val, alignment=wrap)

# 汇总行
rr = r + 1 + len(alloc_data)
apply_cell(ws4, rr, 1, "合计", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
apply_cell(ws4, rr, 2, "¥10,000", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
apply_cell(ws4, rr, 3, "100%", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=dark_fill, alignment=center)
for j in range(4, COLS4 + 1):
    apply_cell(ws4, rr, j, "", fill=dark_fill)

# 五年预期收益模拟
rr += 2
apply_cell(ws4, rr, 1, "▎五年收益情景模拟（假设本金¥10,000一次性投入 + 后续每年追加¥5,000）", font=section_font, fill=white_fill, border=None)
ws4.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS4)
rr += 1

sim_headers = ["情景", "概率", "年化收益假设", "第1年末", "第2年末", "第3年末", "第4年末", "第5年末(2031)"]
for i, h in enumerate(sim_headers):
    apply_header(ws4, rr, i + 1, h)
ws4.row_dimensions[r].height = 28

sim_data = [
    ["🟢 乐观", "25%", "25%-30%",
     "¥12,500-13,000", "¥20,600-21,900", "¥30,800-34,500", "¥43,500-51,900", "¥59,400-76,500"],
    ["🟡 基准", "50%", "15%-20%",
     "¥11,500-12,000", "¥17,700-19,400", "¥25,400-29,300", "¥34,800-42,100", "¥46,500-58,600"],
    ["🔴 悲观", "25%", "0%-5%",
     "¥10,000-10,500", "¥15,000-16,100", "¥20,000-22,100", "¥25,000-28,700", "¥30,000-36,000"],
]

for i, row_data in enumerate(sim_data):
    rr2 = rr + 1 + i
    ws4.row_dimensions[rr2].height = 24
    for j, val in enumerate(row_data):
        apply_cell(ws4, rr2, j + 1, val, alignment=center)

# 说明
rr = rr + 1 + len(sim_data) + 1
note_text = (
    "注：以上为简化模拟，假设每年年末追加¥5,000投入，按年化复利计算，未扣除交易费用和税费。"
    "实际回报受市场波动、买入时机、费率等因素影响，可能与模拟结果存在较大差异。"
)
apply_cell(ws4, rr, 1, note_text, font=small_font, fill=light_fill, border=None)
ws4.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS4)
ws4.row_dimensions[rr].height = 30


# ─── Sheet 5: 风险提示 ─────────────────────────────────────
ws5 = wb.create_sheet("风险提示与应对")
COLS5 = 6
set_col_widths(ws5, {1: 20, 2: 24, 3: 22, 4: 22, 5: 24, 6: 28})

merge_title(ws5, 1, "三大赛道投资风险提示与应对策略", COLS5)
ws5.row_dimensions[1].height = 45

r = 3
risk_headers = ["风险类型", "影响赛道", "触发场景", "预估影响", "应对措施", "预警信号"]
for i, h in enumerate(risk_headers):
    apply_header(ws5, r, i + 1, h)
ws5.row_dimensions[r].height = 28

risk_data = [
    ["AI资本开支大幅缩减",
     "AI与半导体",
     "北美云巨头CAPEX指引连续2季下调；AI变现不及预期引发泡沫破裂",
     "板块可能回调30%-50%",
     "① 减仓至赛道配置的50%\n② 资金转向新能源和债券\n③ 等待业绩验证后再加仓",
     "微软/谷歌/亚马逊季度CAPEX数据、AI应用DAU增速"],
    ["半导体周期见顶",
     "AI与半导体",
     "存储芯片价格连续下跌；DRAM供给过剩；全球半导体销售增速转负",
     "半导体ETF可能下跌40%+",
     "① 半导体设备/芯片ETF止损\n② 保留AI应用端仓位\n③ 等待周期底部信号再抄底",
     "DRAM/NAND合约价、SEMI设备订单数据"],
    ["光伏/锂电产能过剩加剧",
     "新能源与储能",
     "光伏组件价格跌破成本线；锂电电芯价格战；行业亏损面扩大",
     "光伏锂电可能继续下跌20-30%",
     "① 减仓光伏ETF\n② 保留储能和电网ETF\n③ 关注产能出清进度",
     "硅料/碳酸锂价格、行业开工率、龙头企业毛利率"],
    ["储能政策推进不及预期",
     "新能源与储能",
     "容量电价补贴缩水；十五五配套细则迟迟不出；地方执行打折扣",
     "储能板块预期下调15-25%",
     "① 降仓储能至赛道50%\n② 增配电网(确定性更强)\n③ 转向海外储能标的",
     "各省储能补贴政策、电网侧储能IRR变化"],
    ["人形机器人量产delay",
     "机器人与智能制造",
     "特斯拉Optimus V3量产时间推迟超6个月；核心部件良率无法突破",
     "机器人板块可能回调25-35%",
     "① 减仓至赛道配置40%\n② 保留高端制造ETF(基本面更扎实)\n③ 等待量产里程碑确认",
     "特斯拉Optimus进展、核心零部件订单交付数据"],
    ["系统性市场下跌",
     "全部三个赛道",
     "全球金融危机、重大地缘冲突、国内经济硬着陆",
     "全部权益资产可能下跌30-50%",
     "① 权益仓位降至30%\n② 资金转入短债/货币避险\n③ 等待市场企稳后分批抄底",
     "VIX恐慌指数、上证跌破3400、美债收益率倒挂"],
    ["技术路线替代风险",
     "全部（尤其半导体+机器人）",
     "新架构替代现有技术（如量子计算替代GPU、液压替代电驱等）",
     "细分方向可能被颠覆",
     "① 投资ETF而非个股分散风险\n② 关注技术演进动态\n③ 及时调整持仓结构",
     "顶会论文/专利方向变化、巨头技术路线切换"],
]

for i, row_data in enumerate(risk_data):
    rr = r + 1 + i
    ws5.row_dimensions[rr].height = 70
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws5, rr, j + 1, val, font=bold_font)
        elif j == 3:
            apply_cell(ws5, rr, j + 1, val, font=warn_font, alignment=wrap)
        else:
            apply_cell(ws5, rr, j + 1, val, alignment=wrap)


# ── 最终保存 ──────────────────────────────────────────────
output_path = r"C:\Users\Lenovo\理财\未来五年三大赛道投资分析_20260728.xlsx"
wb.save(output_path)
print(f"Report saved to: {output_path}")
