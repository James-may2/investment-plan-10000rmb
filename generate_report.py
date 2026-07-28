"""
============================================================
一万人民币理财投资方案 — Excel 报表生成器
生成日期: 2026-07-28
============================================================
"""

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from datetime import datetime

# ── 样式常量 ──────────────────────────────────────────────
DARK_BLUE   = "1B3A5C"
MED_BLUE    = "2E6B9E"
LIGHT_BLUE  = "D6E4F0"
ACCENT_GOLD = "D4A843"
WHITE       = "FFFFFF"
LIGHT_GRAY  = "F5F5F5"
MED_GRAY    = "E0E0E0"
RED_ALERT   = "C0392B"
GREEN       = "27AE60"
ORANGE      = "E67E22"

title_font      = Font(name="微软雅黑", size=18, bold=True, color=WHITE)
subtitle_font   = Font(name="微软雅黑", size=12, bold=True, color=MED_BLUE)
header_font     = Font(name="微软雅黑", size=11, bold=True, color=WHITE)
section_font    = Font(name="微软雅黑", size=13, bold=True, color=DARK_BLUE)
body_font       = Font(name="微软雅黑", size=10.5, color="333333")
bold_font       = Font(name="微软雅黑", size=10.5, bold=True, color="333333")
warn_font       = Font(name="微软雅黑", size=10, bold=True, color=RED_ALERT)
small_font      = Font(name="微软雅黑", size=9, color="888888")
pct_font        = Font(name="微软雅黑", size=10.5, bold=True, color=MED_BLUE)

title_fill      = PatternFill("solid", fgColor=DARK_BLUE)
header_fill     = PatternFill("solid", fgColor=MED_BLUE)
light_fill      = PatternFill("solid", fgColor=LIGHT_BLUE)
gray_fill       = PatternFill("solid", fgColor=LIGHT_GRAY)
gold_fill       = PatternFill("solid", fgColor=ACCENT_GOLD)
white_fill      = PatternFill("solid", fgColor=WHITE)
red_fill        = PatternFill("solid", fgColor="FDEDEC")
green_fill      = PatternFill("solid", fgColor="EAFAF1")

thin_border = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)
bottom_border = Border(bottom=Side(style="medium", color=DARK_BLUE))

center  = Alignment(horizontal="center", vertical="center", wrap_text=True)
left    = Alignment(horizontal="left", vertical="center", wrap_text=True)
wrap    = Alignment(horizontal="left", vertical="top", wrap_text=True)


def apply_cell(ws, row, col, value, font=body_font, fill=white_fill,
               alignment=left, number_format=None, border=thin_border):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font
    c.fill = fill
    c.alignment = alignment
    if number_format:
        c.number_format = number_format
    if border:
        c.border = border
    return c


def apply_header(ws, row, col, value, width=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = header_font
    c.fill = header_fill
    c.alignment = center
    c.border = thin_border
    if width:
        ws.column_dimensions[get_column_letter(col)].width = width
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


# ═══════════════════════════════════════════════════════════
wb = openpyxl.Workbook()

# ─── Sheet 1: 投资总览 ────────────────────────────────────
ws1 = wb.active
ws1.title = "投资总览"
COLS = 7
set_col_widths(ws1, {1: 20, 2: 16, 3: 14, 4: 14, 5: 18, 6: 22, 7: 22})

# 标题
merge_title(ws1, 1, "📊 一万元人民币理财投资方案", COLS)
ws1.row_dimensions[1].height = 42

# 基本信息区
r = 3
info_data = [
    ("委托资金", "¥10,000.00（人民币壹万元整）"),
    ("报告日期", "2026年7月28日"),
    ("投资期限", "短期 6 个月（至2027年1月） / 长期 1-2 年（至2028年7月）"),
    ("风险偏好", "稳健偏成长（R2-R4 混合配置）"),
    ("预期年化收益", "短期 3%-8% / 长期 8%-18%"),
    ("最大可承受回撤", "短期 ≤3% / 长期 ≤15%"),
]
for i, (k, v) in enumerate(info_data):
    apply_cell(ws1, r + i, 1, k, font=bold_font, fill=light_fill)
    ws1.merge_cells(start_row=r + i, start_column=2, end_row=r + i, end_column=COLS)
    apply_cell(ws1, r + i, 2, v, fill=light_fill)
    ws1.row_dimensions[r + i].height = 26

# 市场环境摘要
r += len(info_data) + 1
apply_cell(ws1, r, 1, "▎当前市场环境（2026年7月）", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
r += 1
market_notes = [
    "• 上证指数约 3,800 点，经历 7 月调整后多家机构判断为「黄金买点」",
    "• 下半年政策基调：财政为主、货币护航，预计降准 50bp",
    "• AI/科技为全市场共识主线，盈利增长强劲，但需精选业绩确定标的",
    "• 红利资产进入「击球区」，股息率与国债利差达 3.5%",
    "• 黄金 Q3 寻底（支撑 $4,000/oz），Q4 有望修复反弹",
    "• 短债基金收益率 1.3%-3.3%，适合作为流动安全垫",
]
for note in market_notes:
    apply_cell(ws1, r, 1, note, font=body_font, fill=white_fill, border=None)
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
    ws1.row_dimensions[r].height = 22
    r += 1

# 核心策略
r += 1
apply_cell(ws1, r, 1, "▎核心配置策略：哑铃型结构", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
r += 1
strategy_text = (
    "一端配置「科技成长」（科创50 + 半导体/AI 主题 ETF）捕捉产业红利，"
    "另一端配置「稳健防御」（短债 + 红利 + 黄金）对冲波动。"
    "中段以「沪深300 宽基」作为底仓，享受市场整体上行。"
)
apply_cell(ws1, r, 1, strategy_text, font=body_font, fill=light_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
ws1.row_dimensions[r].height = 30

# 配置饼图数据（文字版）
r += 2
apply_cell(ws1, r, 1, "▎资产配置一览", font=section_font, fill=white_fill, border=None)
ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
r += 1
alloc_headers = ["资产类别", "投资金额(元)", "占比(%)", "风险等级", "投资期限", "预期年化", "核心逻辑"]
for i, h in enumerate(alloc_headers):
    apply_header(ws1, r, i + 1, h)
ws1.row_dimensions[r].height = 28

alloc_data = [
    ["💰 短债基金",       2000, 20, "R2（低风险）", "短期 6 个月", "1.5%-3%",   "流动安全垫，随时可取"],
    ["📈 沪深300 ETF",    2500, 25, "R3（中风险）", "中长期 1-2 年","8%-15%",   "宽基底仓，赚市场整体上行"],
    ["🚀 科创50 ETF",     2000, 20, "R4（中高风险）","中长期 1-2 年","12%-25%", "AI主线核心进攻仓位"],
    ["🛡️ 固收+基金",      1500, 15, "R2-R3",        "中期 6-12 个月","3%-8%",  "稳健增强，债券打底+少量权益"],
    ["🥇 黄金ETF",        1000, 10, "R3（中风险）", "中长期 1-2 年","5%-15%",  "对冲通胀与地缘风险"],
    ["🏭 红利低波ETF",     500,  5, "R2-R3",        "中长期 1-2 年","4%-10%",  "防御收息，低波动压舱"],
    ["💵 现金/货币基金",    500,  5, "R1（低风险）", "活期/随时",    "0.5%-1.5%","极端情况补仓备用"],
]
for i, row_data in enumerate(alloc_data):
    rr = r + 1 + i
    ws1.row_dimensions[rr].height = 26
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws1, rr, j + 1, val, font=bold_font)
        elif j == 2:
            apply_cell(ws1, rr, j + 1, val, font=pct_font)
        else:
            apply_cell(ws1, rr, j + 1, val)

# 汇总行
rr = r + 1 + len(alloc_data)
ws1.row_dimensions[rr].height = 26
apply_cell(ws1, rr, 1, "📌 合计", font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=DARK_BLUE))
apply_cell(ws1, rr, 2, 10000, font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=DARK_BLUE))
apply_cell(ws1, rr, 3, 100, font=Font(name="微软雅黑", size=11, bold=True, color=WHITE), fill=PatternFill("solid", fgColor=DARK_BLUE))
for j in range(4, COLS + 1):
    apply_cell(ws1, rr, j, "", fill=PatternFill("solid", fgColor=DARK_BLUE))

# 免责声明
rr += 2
disclaimer = (
    "⚠️ 免责声明：本报表基于2026年7月28日公开市场数据及机构研报编制，仅供学习参考，不构成投资建议。"
    "市场有风险，投资需谨慎。过往业绩不代表未来表现。投资者应根据自身风险承受能力独立决策。"
)
apply_cell(ws1, rr, 1, disclaimer, font=small_font, fill=red_fill, border=None)
ws1.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS)
ws1.row_dimensions[rr].height = 36


# ─── Sheet 2: 具体投资标的 ─────────────────────────────────
ws2 = wb.create_sheet("具体投资标的")
COLS2 = 10
set_col_widths(ws2, {1: 18, 2: 22, 3: 12, 4: 10, 5: 12, 6: 12, 7: 12, 8: 16, 9: 18, 10: 22})

merge_title(ws2, 1, "📋 具体投资标的明细", COLS2)
ws2.row_dimensions[1].height = 42

r = 3
apply_cell(ws2, r, 1, "▎推荐基金/ETF 清单（截至2026年7月）", font=section_font, fill=white_fill, border=None)
ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS2)
r += 1

fund_headers = [
    "资产类别", "基金/ETF 名称", "代码", "类型", "配置金额",
    "风险等级", "近1年收益", "管理费率", "买入策略", "推荐理由"
]
for i, h in enumerate(fund_headers):
    apply_header(ws2, r, i + 1, h)
ws2.row_dimensions[r].height = 28

fund_data = [
    # 短债基金
    ["短债基金", "嘉实超短债债券A", "012773", "场外债基", "¥1,000",
     "R2", "1.41%", "0.30%", "一次性买入，持有30天以上免赎回费",
     "规模48亿，回撤极低，持有30天盈利概率超98%"],
    ["短债基金", "国泰利享安益短债C", "—", "场外债基", "¥1,000",
     "R2", "1.41%", "0.30%", "一次性买入，持有7天以上",
     "近1年最大回撤仅0.04%，低波王者"],
    # 沪深300
    ["沪深300 ETF", "沪深300ETF华泰柏瑞", "510360", "场内ETF", "¥1,500",
     "R3", "—", "0.50%", "分3批定投，每两周一次，每次¥500",
     "规模最大流动性最好，7月净流入超200亿，估值历史偏低"],
    ["沪深300 ETF", "易方达沪深300ETF联接A", "110020", "场外联接", "¥1,000",
     "R3", "—", "0.15%", "一次性买入，或与场内搭配",
     "场外方便定投，费率低至0.15%"],
    # 科创50
    ["科创50 ETF", "科创50ETF华夏", "588000", "场内ETF", "¥1,500",
     "R4", "—", "0.50%", "分3批逢低买入，每次回调3%-5%加仓¥500",
     "7月净流入超300亿，成交额超千亿，AI核心标的"],
    ["科创50 ETF", "易方达科创50ETF联接C", "011609", "场外联接", "¥500",
     "R4", "—", "0.15%", "定投备用，场外方便小额投入",
     "C类免申购费，适合短期持有"],
    # 固收+
    ["固收+基金", "博时恒乐C", "—", "场外混合", "¥1,000",
     "R2-R3", "6.86%", "0.80%", "一次性买入，持有6个月以上",
     "固收+人气榜第1，债券打底+少量权益"],
    ["固收+基金", "景顺长城景颐双利C", "—", "场外混合", "¥500",
     "R2-R3", "6.19%", "0.80%", "一次性买入，持有6个月以上",
     "人气榜前列，成立时间长，业绩稳健"],
    # 黄金ETF
    ["黄金ETF", "华安黄金ETF", "518880", "场内ETF", "¥500",
     "R3", "—", "0.50%", "Q3逢回调（金价<$4,000/oz时）分批买入",
     "国内规模最大的黄金ETF，流动性最好"],
    ["黄金ETF", "博时黄金ETF联接C", "002611", "场外联接", "¥500",
     "R3", "—", "0.50%", "与场内同步，定投方式",
     "场外方便小额投入，C类免申购费"],
    # 红利低波
    ["红利低波ETF", "港股红利低波ETF景顺", "159569", "场内ETF", "¥300",
     "R2-R3", "—", "0.50%", "一次性买入或分2批",
     "港股红利估值更低，股息率吸引力强"],
    ["红利低波ETF", "中证红利ETF易方达", "515180", "场内ETF", "¥200",
     "R2-R3", "—", "0.15%", "一次性买入",
     "A股红利标杆，费率极低"],
    # 现金
    ["货币基金", "余额宝/零钱通", "—", "货币基金", "¥500",
     "R1", "~1.2%", "0%", "随时转入",
     "日常流动性管理，极端情况补仓资金"],
]

for i, row_data in enumerate(fund_data):
    rr = r + 1 + i
    ws2.row_dimensions[rr].height = 32
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws2, rr, j + 1, val, font=bold_font)
        elif j == 4:
            apply_cell(ws2, rr, j + 1, val, font=pct_font)
        else:
            apply_cell(ws2, rr, j + 1, val)

# 费用说明
rr = r + 1 + len(fund_data) + 1
apply_cell(ws2, rr, 1, "▎费用说明", font=section_font, fill=white_fill, border=None)
ws2.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS2)
rr += 1
fee_notes = [
    "• 场内ETF交易佣金：约万2.5（券商收取），单笔最低5元。建议单笔交易金额≥2,000元以摊薄佣金占比。",
    "• 场外基金申购费：多数C类份额免申购费；A类申购费约0.1%-0.15%（1折后），持有超30天/180天免赎回费。",
    "• 管理费+托管费：已包含在基金净值中，投资者无需额外支付。",
    "• 综合年化持有成本：约0.3%-0.8%，对一万元本金即年费30-80元。",
]
for note in fee_notes:
    apply_cell(ws2, rr, 1, note, font=body_font, fill=white_fill, border=None)
    ws2.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS2)
    ws2.row_dimensions[rr].height = 22
    rr += 1


# ─── Sheet 3: 买卖操作指南 ─────────────────────────────────
ws3 = wb.create_sheet("买卖操作指南")
COLS3 = 7
set_col_widths(ws3, {1: 18, 2: 28, 3: 26, 4: 26, 5: 24, 6: 22, 7: 22})

merge_title(ws3, 1, "📅 买入/卖出操作时间表", COLS3)
ws3.row_dimensions[1].height = 42

r = 3
apply_cell(ws3, r, 1, "▎第一阶段：建仓期（2026年8月）", font=section_font, fill=white_fill, border=None)
ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS3)
r += 1

op_headers = ["资产类别", "买入时机/条件", "买入方式", "卖出条件/止盈", "止损线", "目标持有期", "注意事项"]
for i, h in enumerate(op_headers):
    apply_header(ws3, r, i + 1, h)
ws3.row_dimensions[r].height = 28

op_data = [
    ["短债基金\n(¥2,000)",
     "立即一次性买入\n（8月第1周）",
     "支付宝/天天基金/\n银行App申购",
     "持有到期即可，\n无需主动卖出",
     "不设止损\n（低风险品种）",
     "6个月",
     "选择C类份额，\n免申购费"],
    ["沪深300 ETF\n(¥2,500)",
     "8月开始定投，\n分3批买入：\n第1批 8/1\n第2批 8/15\n第3批 9/1",
     "券商APP场内买入\n或支付宝定投\n场外联接基金",
     "① 盈利≥15% 止盈1/3\n② 盈利≥25% 止盈1/2\n③ 指数破4500点\n  逐步减仓",
     "亏损≥10%\n减仓50%",
     "1-2年",
     "定投平滑成本，\n避免单点追高"],
    ["科创50 ETF\n(¥2,000)",
     "8-9月逢低分批：\n回调3%加仓¥500\n回调5%再加仓¥500\n最多分4批",
     "券商APP场内买入\n(代码588000)",
     "① 盈利≥20% 止盈1/3\n② 盈利≥35% 止盈1/2\n③ AI概念过热时\n  分批减仓",
     "亏损≥15%\n止损离场",
     "1-2年",
     "高波动品种，\n控制仓位上限\n不超过总资金30%"],
    ["固收+基金\n(¥1,500)",
     "8月一次性买入\n或分2批\n（8月初+9月初）",
     "支付宝/天天基金/\n银行App申购",
     "① 年化收益≥8%\n  考虑赎回\n② 持有满6个月后\n  灵活调整",
     "亏损≥5%\n减仓50%",
     "6-12个月",
     "关注季度报告，\n权益仓位变化"],
    ["黄金ETF\n(¥1,000)",
     "Q3逢低布局：\n国际金价<$4,000/oz\n时买入首批¥500\n10月再买¥500",
     "券商APP场内买入\n(代码518880)\n或场外联接",
     "① 金价反弹至\n  $4,400+ 止盈50%\n② $4,600+ 全部止盈\n③ Q4末逐步减仓",
     "金价跌破\n$3,700止损",
     "1-2年\n(波段操作)",
     "等待Q3低点，\n勿追高；\n关注美联储议息"],
    ["红利低波ETF\n(¥500)",
     "8月一次性买入\n或分2批",
     "券商APP场内买入\n(代码159569或\n515180)",
     "① 股息率回落至\n  2.5%以下考虑卖出\n② 替代：长期持有\n  享受分红",
     "亏损≥10%\n减仓50%",
     "1-2年\n(可长期持有)",
     "防御品种，\n波动较小"],
    ["现金/货币基金\n(¥500)",
     "立即转入",
     "支付宝余额宝/\n微信零钱通",
     "随时可取",
     "不适用",
     "活期",
     "极端补仓储备金，\n日常不轻易动用"],
]

for i, row_data in enumerate(op_data):
    rr = r + 1 + i
    ws3.row_dimensions[rr].height = 95
    for j, val in enumerate(row_data):
        if j == 0:
            apply_cell(ws3, rr, j + 1, val, font=bold_font, alignment=center)
        else:
            apply_cell(ws3, rr, j + 1, val, alignment=wrap)

# 时间线
rr = r + 1 + len(op_data) + 1
apply_cell(ws3, rr, 1, "▎关键时间节点", font=section_font, fill=white_fill, border=None)
ws3.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS3)
rr += 1
timeline = [
    ["2026年8月",  "建仓月：买入全部短债、固收+、红利ETF；启动沪深300和科创50定投"],
    ["2026年9月",  "定投第2批沪深300；视回调情况加仓科创50"],
    ["2026年10月", "黄金Q3低点布局窗口；关注三季报业绩验证"],
    ["2026年11月", "美国中期选举前后关注市场波动，可能带来加仓机会"],
    ["2026年12月", "年末资金面偏紧，可能迎来阶段性低点，定投加仓"],
    ["2027年1月",  "短期复盘：评估短债+固收+表现，决定是否到期转换"],
    ["2027年Q1",   "春季躁动窗口，关注科技板块是否超预期"],
    ["2027年Q2",   "年报+一季报披露期，业绩验证后调整持仓结构"],
    ["2027年7月",  "长期投资一周年：全面复盘，重新平衡配置"],
    ["2028年7月",  "长期投资两周年：最终评估，根据当时市场决定去留"],
]
for i, (date, event) in enumerate(timeline):
    apply_cell(ws3, rr + i, 1, date, font=bold_font, fill=light_fill)
    ws3.merge_cells(start_row=rr + i, start_column=2, end_row=rr + i, end_column=COLS3)
    apply_cell(ws3, rr + i, 2, event, fill=light_fill)
    ws3.row_dimensions[rr + i].height = 22


# ─── Sheet 4: 风险分析与应对 ───────────────────────────────
ws4 = wb.create_sheet("风险分析与应对")
COLS4 = 6
set_col_widths(ws4, {1: 22, 2: 26, 3: 16, 4: 26, 5: 26, 6: 30})

merge_title(ws4, 1, "⚠️ 风险分析与应对策略", COLS4)
ws4.row_dimensions[1].height = 42

r = 3
risk_headers = ["风险类型", "触发场景", "影响程度", "影响资产", "应对措施", "预案"]
for i, h in enumerate(risk_headers):
    apply_header(ws4, r, i + 1, h)
ws4.row_dimensions[r].height = 28

risks = [
    ["市场系统性下跌",
     "上证跌破3400点\n或出现连续大跌",
     "★★★★★\n（高）",
     "全部权益类资产\n（沪深300、科创50\n红利ETF）",
     "① 暂停定投，观望\n② 红利和债券做防守\n③ 现金储备补仓",
     "最大容忍总亏损\n15%（¥1,500），\n触及即全面减仓"],
    ["AI赛道泡沫破裂",
     "AI龙头业绩大幅miss\n或海外AI支出骤降",
     "★★★★\n（较高）",
     "科创50 ETF\n（占比20%）",
     "① 科创50跌破成本\n  15%即止损\n② 资金转回债基避险",
     "科创50单品种\n最大亏损不超过\n¥400（20%）"],
    ["美债利率飙升/\n美联储超预期加息",
     "10Y美债破5%\n美联储再加息2次+",
     "★★★★\n（较高）",
     "黄金ETF\n科创50 ETF",
     "① 黄金加速下跌则\n  暂缓加仓\n② 转向短债避险",
     "黄金跌破$3,700\n全部止损"],
    ["人民币大幅升值",
     "USD/CNY跌破6.8",
     "★★★\n（中等）",
     "港股红利ETF\n黄金ETF",
     "① 减仓港股红利\n② 黄金国际定价\n  影响有限",
     "港股仓位仅5%\n影响可控"],
    ["国内经济超预期下滑",
     "GDP增速低于4%\n消费持续低迷",
     "★★★\n（中等）",
     "沪深300 ETF\n固收+基金",
     "① 增配债券/货币\n② 减仓权益\n③ 等待政策对冲",
     "降息降准预期下\n债市反而受益"],
    ["流动性危机/\n黑天鹅事件",
     "类似疫情级别\n突发事件",
     "★★\n（较低概率）",
     "全部资产",
     "① 立即赎回短债\n② 权益减仓至30%\n③ 持有现金等待",
     "极端情况保留\n至少¥3,000现金"],
    ["通胀超预期上行",
     "CPI持续>3%\n央行收紧货币",
     "★★★\n（中等）",
     "固收+基金\n债券类资产",
     "① 缩短债基久期\n② 黄金对冲通胀\n  反而受益",
     "配置黄金10%\n已部分对冲"],
]

for i, row_data in enumerate(risks):
    rr = r + 1 + i
    ws4.row_dimensions[rr].height = 70
    for j, val in enumerate(row_data):
        if j == 2:
            apply_cell(ws4, rr, j + 1, val, font=warn_font, alignment=center)
        elif j == 0:
            apply_cell(ws4, rr, j + 1, val, font=bold_font)
        else:
            apply_cell(ws4, rr, j + 1, val, alignment=wrap)

# 风险预算总结
rr = r + 1 + len(risks) + 1
apply_cell(ws4, rr, 1, "▎风险预算总结", font=section_font, fill=white_fill, border=None)
ws4.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS4)
rr += 1
risk_summary = [
    "• 极端情景（全部触发）：最大可能亏损 ≈ ¥3,000-4,000（总资金30%-40%）",
    "• 正常波动范围：月度波动 ±¥200-500（总资金2%-5%）",
    "• 配置中 R1-R2 低风险资产占比 40%（短债+固收++货币+红利），提供安全缓冲",
    "• VaR(95%, 1个月)：约 ¥350，即95%概率下月度亏损不超过350元",
    "• 最大回撤控制目标：整体组合≤15%（¥1,500），单个品种≤20%",
]
for note in risk_summary:
    apply_cell(ws4, rr, 1, note, font=body_font, fill=white_fill, border=None)
    ws4.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS4)
    ws4.row_dimensions[rr].height = 22
    rr += 1


# ─── Sheet 5: 学习笔记 ─────────────────────────────────────
ws5 = wb.create_sheet("投资学习笔记")
COLS5 = 5
set_col_widths(ws5, {1: 18, 2: 30, 3: 26, 4: 24, 5: 30})

merge_title(ws5, 1, "📖 投资入门知识学习笔记", COLS5)
ws5.row_dimensions[1].height = 42

r = 3
apply_cell(ws5, r, 1, "▎基金/ETF 基础知识", font=section_font, fill=white_fill, border=None)
ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS5)
r += 1

knowledge = [
    ["概念", "解释", "关键要点", "适用人群", "学习来源"],
    ["ETF（交易型\n开放式指数基金）",
     "在交易所上市交易的\n指数基金，可以像买卖\n股票一样实时交易",
     "• 费率低（0.15%-0.5%）\n• 透明度高\n• 流动性好\n• 需要证券账户",
     "有一定投资经验\n的投资者",
     "上交所/深交所\n投资者教育专栏"],
    ["场外基金",
     "通过基金公司/银行/\n第三方平台（支付宝等）\n申购赎回的基金",
     "• 操作便捷\n• 可设置定投\n• 赎回T+1到账\n• 无需证券账户",
     "初学者\n上班族",
     "支付宝-理财-\n基金入门"],
    ["债券基金",
     "80%以上资产投资于\n债券的基金，收益来自\n票息+资本利得",
     "• 风险R1-R2\n• 年化1%-4%\n• 替代银行理财\n• 关注久期和信用等级",
     "保守型投资者\n短期资金管理",
     "中国债券信息网\n晨星中国"],
    ["固收+基金",
     "以债券等固收资产打底\n（≥70%），辅以少量\n权益/可转债增厚收益",
     "• 回撤通常<5%\n• 年化3%-8%\n• 攻守兼备\n• 注意权益仓位比例",
     "稳健偏进取\n投资者",
     "天天基金网\n基金定期报告"],
    ["指数基金\n（宽基指数）",
     "跟踪沪深300、科创50\n等宽基指数的基金，\n被动管理",
     "• 分散个股风险\n• 费率低于主动基金\n• 赚市场平均收益\n• 长期定投效果好",
     "所有投资者\n尤其新手",
     "中证指数官网\n约翰·博格《共同\n基金常识》"],
    ["定投\n（定期定额投资）",
     "固定时间、固定金额\n买入同一基金，\n摊平买入成本",
     "• 克服择时困难\n• 利用波动降低\n  平均成本\n• 纪律性强\n• 需长期坚持",
     "工薪族\n波动市投资者",
     "《聪明的投资者》\n本杰明·格雷厄姆"],
    ["资产配置",
     "根据风险偏好将资金\n分配在不同类别资产\n（股/债/黄金/现金）",
     "• 决定90%以上的\n  长期收益\n• 定期再平衡\n• 分散化是唯一\n  免费午餐",
     "所有投资者",
     "《漫步华尔街》\n伯顿·马尔基尔"],
    ["夏普比率/\n最大回撤",
     "夏普比率=超额收益/\n波动率，衡量风险调整\n后收益；最大回撤=\n从高点到低点最大跌幅",
     "• 夏普>1算不错\n• 最大回撤决定了\n  你能拿不拿得住\n• 选基金先看回撤",
     "进阶投资者",
     "晨星中国\n且慢/有知有行"],
]

for i, row_data in enumerate(knowledge):
    rr = r + i
    ws5.row_dimensions[rr].height = 85 if i > 0 else 28
    for j, val in enumerate(row_data):
        if i == 0:
            apply_header(ws5, rr, j + 1, val)
        elif j == 0:
            apply_cell(ws5, rr, j + 1, val, font=bold_font, alignment=center)
        else:
            apply_cell(ws5, rr, j + 1, val, alignment=wrap)

# 推荐书单
rr = r + len(knowledge) + 1
apply_cell(ws5, rr, 1, "▎推荐学习资源", font=section_font, fill=white_fill, border=None)
ws5.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS5)
rr += 1
books = [
    "📚 入门：《指数基金投资指南》（银行螺丝钉）、《小狗钱钱》（博多·舍费尔）",
    "📚 进阶：《聪明的投资者》（格雷厄姆）、《漫步华尔街》（马尔基尔）、《投资中最简单的事》（邱国鹭）",
    "📚 宏观：《置身事内》（兰小欢）、《周期》（霍华德·马克斯）",
    "🌐 网站：晨星中国(morningstar.cn)、天天基金网(1234567.com.cn)、中证指数(csindex.com.cn)",
    "🌐 工具：且慢(qieman.com)、有知有行、雪球(xueqiu.com)、韭圈儿",
    "📱 APP：支付宝-理财、天天基金、蛋卷基金、华泰涨乐财富通",
]
for note in books:
    apply_cell(ws5, rr, 1, note, font=body_font, fill=white_fill, border=None)
    ws5.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS5)
    ws5.row_dimensions[rr].height = 22
    rr += 1


# ─── Sheet 6: 复盘跟踪表 ───────────────────────────────────
ws6 = wb.create_sheet("复盘跟踪模板")
COLS6 = 10
set_col_widths(ws6, {1: 14, 2: 18, 3: 12, 4: 12, 5: 12, 6: 12, 7: 12, 8: 14, 9: 16, 10: 30})

merge_title(ws6, 1, "📝 月度投资复盘跟踪表（模板）", COLS6)
ws6.row_dimensions[1].height = 42

r = 3
apply_cell(ws6, r, 1, "说明：每月初填写上月数据，跟踪每项资产表现。本表为空白模板，请自行复印填写。",
           font=small_font, fill=light_fill, border=None)
ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS6)
r += 1

track_headers = [
    "日期", "资产类别", "持有金额(元)", "市值(元)", "浮盈/亏(元)",
    "收益率(%)", "累计收益(元)", "操作记录", "下月计划", "备注/市场观察"
]
for i, h in enumerate(track_headers):
    apply_header(ws6, r, i + 1, h)
ws6.row_dimensions[r].height = 28

# 预填8月初始数据
initial_data = [
    ["2026-08-01", "短债基金",   2000, "", "", "", "", "买入¥2,000", "持有不动", ""],
    ["2026-08-01", "沪深300 ETF",2500, "", "", "", "", "首批定投¥833", "继续定投", ""],
    ["2026-08-01", "科创50 ETF", 2000, "", "", "", "", "待回调买入", "等待回调信号", "关注AI板块走势"],
    ["2026-08-01", "固收+基金",  1500, "", "", "", "", "买入¥1,500", "持有不动", ""],
    ["2026-08-01", "黄金ETF",    1000, "", "", "", "", "暂未买入", "等待Q3低点", "金价观察"],
    ["2026-08-01", "红利低波ETF", 500, "", "", "", "", "买入¥500", "持有不动", ""],
    ["2026-08-01", "现金/货币",   500, "", "", "", "", "保留", "保留备用", ""],
]
for i, row_data in enumerate(initial_data):
    rr = r + 1 + i
    ws6.row_dimensions[rr].height = 22
    for j, val in enumerate(row_data):
        apply_cell(ws6, rr, j + 1, val)

# 汇总行
rr = r + 1 + len(initial_data)
apply_cell(ws6, rr, 1, "合计", font=bold_font, fill=light_fill)
apply_cell(ws6, rr, 3, 10000, font=bold_font, fill=light_fill)
for j in [1, 2] + list(range(4, COLS6 + 1)):
    if j != 3:
        apply_cell(ws6, rr, j, "", fill=light_fill)

# 收益情景模拟
rr += 2
apply_cell(ws6, rr, 1, "▎收益情景模拟", font=section_font, fill=white_fill, border=None)
ws6.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=COLS6)
rr += 1

scenario_headers = ["情景", "概率估计", "半年预期收益", "半年后总资产", "一年预期收益", "一年后总资产", "两年预期收益", "两年后总资产", "主要假设", ""]
# split into two rows for readability
for i, h in enumerate(scenario_headers[:9]):
    apply_header(ws6, rr, i + 1, h)

scenarios = [
    ["🟢 乐观", "30%", "+¥600~800", "¥10,600~10,800", "+¥1,200~1,800", "¥11,200~11,800", "+¥2,500~4,000", "¥12,500~14,000", "AI主线持续+大盘创新高+黄金反弹"],
    ["🟡 基准", "45%", "+¥250~500", "¥10,250~10,500", "+¥600~1,200", "¥10,600~11,200", "+¥1,200~2,500", "¥11,200~12,500", "市场温和上行+结构性行情"],
    ["🔴 悲观", "25%", "-¥150~+¥100","¥9,850~10,100","-¥300~+¥300", "¥9,700~10,300", "-¥500~+¥500", "¥9,500~10,500", "市场震荡下跌+AI估值回归"],
]
for i, row_data in enumerate(scenarios):
    rr2 = rr + 1 + i
    ws6.row_dimensions[rr2].height = 22
    for j, val in enumerate(row_data):
        apply_cell(ws6, rr2, j + 1, val)

# ── 最终保存 ──────────────────────────────────────────────
output_path = r"C:\Users\Lenovo\理财\一万元理财投资方案_20260728.xlsx"
wb.save(output_path)
print(f"Report saved to: {output_path}")
