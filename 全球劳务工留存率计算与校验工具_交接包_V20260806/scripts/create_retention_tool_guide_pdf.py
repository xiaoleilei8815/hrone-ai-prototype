from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "output" / "pdf" / "劳务工留存率计算与校验工具_区域使用说明.pdf"

INK = colors.HexColor("#24313A")
INK_SOFT = colors.HexColor("#52606A")
TEAL = colors.HexColor("#0F766E")
TEAL_DARK = colors.HexColor("#0B5E58")
TEAL_LIGHT = colors.HexColor("#E2F2EF")
BLUE = colors.HexColor("#35648B")
BLUE_LIGHT = colors.HexColor("#EAF1F6")
AMBER = colors.HexColor("#9A6217")
AMBER_LIGHT = colors.HexColor("#FFF4D6")
RED = colors.HexColor("#B6423C")
RED_LIGHT = colors.HexColor("#FBE9E7")
LINE = colors.HexColor("#D7DCE0")
SURFACE = colors.white
SURFACE_SOFT = colors.HexColor("#F5F7F8")


pdfmetrics.registerFont(
    TTFont("Heiti", "/System/Library/Fonts/STHeiti Medium.ttc", subfontIndex=0)
)
pdfmetrics.registerFont(
    TTFont("HeitiLight", "/System/Library/Fonts/STHeiti Light.ttc", subfontIndex=0)
)
pdfmetrics.registerFontFamily(
    "HeitiLight", normal="HeitiLight", bold="Heiti", italic="HeitiLight", boldItalic="Heiti"
)


BASE = getSampleStyleSheet()
BODY = ParagraphStyle(
    "BodyCN",
    parent=BASE["BodyText"],
    fontName="HeitiLight",
    fontSize=9.5,
    leading=15,
    textColor=INK,
    wordWrap="CJK",
    spaceAfter=5,
)
SMALL = ParagraphStyle(
    "SmallCN", parent=BODY, fontSize=8.2, leading=12.5, textColor=INK_SOFT
)
TITLE = ParagraphStyle(
    "TitleCN",
    parent=BODY,
    fontName="Heiti",
    fontSize=25,
    leading=34,
    textColor=SURFACE,
    alignment=TA_LEFT,
)
SUBTITLE = ParagraphStyle(
    "SubtitleCN",
    parent=BODY,
    fontSize=11,
    leading=18,
    textColor=colors.HexColor("#D7E7E5"),
)
H1 = ParagraphStyle(
    "H1CN",
    parent=BODY,
    fontName="Heiti",
    fontSize=17,
    leading=24,
    textColor=INK,
    spaceAfter=10,
)
H2 = ParagraphStyle(
    "H2CN",
    parent=BODY,
    fontName="Heiti",
    fontSize=12,
    leading=18,
    textColor=TEAL_DARK,
    spaceBefore=5,
    spaceAfter=6,
)
CALLOUT = ParagraphStyle(
    "CalloutCN",
    parent=BODY,
    fontName="Heiti",
    fontSize=12,
    leading=19,
    textColor=INK,
    alignment=TA_CENTER,
)
FORMULA = ParagraphStyle(
    "FormulaCN",
    parent=BODY,
    fontName="Heiti",
    fontSize=11,
    leading=18,
    textColor=TEAL_DARK,
    alignment=TA_CENTER,
)
TABLE_HEAD = ParagraphStyle(
    "TableHeadCN",
    parent=BODY,
    fontName="Heiti",
    fontSize=8.5,
    leading=12,
    textColor=SURFACE,
    alignment=TA_CENTER,
)
TABLE_BODY = ParagraphStyle(
    "TableBodyCN",
    parent=BODY,
    fontSize=8.2,
    leading=12.5,
    textColor=INK,
)


def p(text, style=BODY):
    return Paragraph(text, style)


def bullets(items, style=BODY):
    result = []
    for item in items:
        result.extend([p(f"•&nbsp;&nbsp;{item}", style), Spacer(1, 1.2 * mm)])
    return result


def section_title(number, title, kicker=None):
    label = f"{number:02d}"
    left = Table(
        [[p(label, ParagraphStyle("Num", parent=H2, textColor=SURFACE, alignment=TA_CENTER)), p(title, H1)]],
        colWidths=[13 * mm, 158 * mm],
    )
    left.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), TEAL),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (0, 0), 0),
                ("RIGHTPADDING", (0, 0), (0, 0), 0),
                ("TOPPADDING", (0, 0), (0, 0), 6),
                ("BOTTOMPADDING", (0, 0), (0, 0), 6),
                ("LEFTPADDING", (1, 0), (1, 0), 8),
                ("RIGHTPADDING", (1, 0), (1, 0), 0),
                ("TOPPADDING", (1, 0), (1, 0), 0),
                ("BOTTOMPADDING", (1, 0), (1, 0), 0),
                ("LINEBELOW", (1, 0), (1, 0), 1, LINE),
            ]
        )
    )
    result = [left, Spacer(1, 4 * mm)]
    if kicker:
        result.extend([p(kicker, SMALL), Spacer(1, 2 * mm)])
    return result


def info_box(title, body, background=TEAL_LIGHT, accent=TEAL, width=171 * mm):
    table = Table(
        [[p(title, H2), p(body, BODY)]],
        colWidths=[34 * mm, width - 34 * mm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), background),
                ("LINEBEFORE", (0, 0), (0, -1), 4, accent),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return table


def formula_box(text):
    table = Table([[p(text, FORMULA)]], colWidths=[171 * mm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), SURFACE_SOFT),
                ("BOX", (0, 0), (-1, -1), 0.8, LINE),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
            ]
        )
    )
    return table


def data_table(headers, rows, widths, header_color=TEAL_DARK):
    data = [[p(cell, TABLE_HEAD) for cell in headers]]
    for row in rows:
        data.append([p(str(cell), TABLE_BODY) for cell in row])
    table = Table(data, colWidths=widths, repeatRows=1, hAlign="LEFT")
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), header_color),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.45, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]
    for row_index in range(1, len(data)):
        if row_index % 2 == 0:
            style.append(("BACKGROUND", (0, row_index), (-1, row_index), SURFACE_SOFT))
    table.setStyle(TableStyle(style))
    return table


def step_block(number, title, body, tone=TEAL):
    num_style = ParagraphStyle(
        f"Step{number}", parent=H2, textColor=SURFACE, alignment=TA_CENTER, fontSize=11
    )
    table = Table(
        [[p(str(number), num_style), p(f"<b>{title}</b><br/>{body}", BODY)]],
        colWidths=[12 * mm, 159 * mm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, 0), tone),
                ("BACKGROUND", (1, 0), (1, 0), SURFACE_SOFT),
                ("BOX", (1, 0), (1, 0), 0.5, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (0, 0), 0),
                ("RIGHTPADDING", (0, 0), (0, 0), 0),
                ("LEFTPADDING", (1, 0), (1, 0), 9),
                ("RIGHTPADDING", (1, 0), (1, 0), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return table


class GuideDocTemplate(BaseDocTemplate):
    def __init__(self, filename):
        super().__init__(
            filename,
            pagesize=A4,
            rightMargin=19 * mm,
            leftMargin=19 * mm,
            topMargin=19 * mm,
            bottomMargin=18 * mm,
            title="劳务工留存率计算与校验工具 - 区域使用说明",
            author="FBU 海外交付前台",
        )
        frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height,
            id="body",
        )
        self.addPageTemplates([PageTemplate(id="guide", frames=[frame], onPage=self.draw_page)])

    @staticmethod
    def draw_page(canvas, doc):
        width, height = A4
        if doc.page > 1:
            canvas.saveState()
            canvas.setFillColor(INK)
            canvas.rect(0, height - 11 * mm, width, 11 * mm, fill=1, stroke=0)
            canvas.setFont("Heiti", 7.5)
            canvas.setFillColor(colors.HexColor("#C9D3D8"))
            canvas.drawString(19 * mm, height - 7.2 * mm, "OTWS / SUPPLIER PORTAL")
            canvas.setFillColor(INK_SOFT)
            canvas.setFont("HeitiLight", 7.5)
            canvas.drawString(19 * mm, 9 * mm, "劳务工留存率计算与校验工具  |  区域使用说明")
            canvas.drawRightString(width - 19 * mm, 9 * mm, f"{doc.page:02d}")
            canvas.setStrokeColor(LINE)
            canvas.line(19 * mm, 13 * mm, width - 19 * mm, 13 * mm)
            canvas.restoreState()


def build_story():
    story = []

    # Cover
    cover = Table(
        [[p("OTWS / SUPPLIER PORTAL", SUBTITLE)], [Spacer(1, 16 * mm)], [p("劳务工留存率<br/>计算与校验工具", TITLE)], [Spacer(1, 7 * mm)], [p("区域使用说明", ParagraphStyle("CoverSub", parent=SUBTITLE, fontName="Heiti", fontSize=15))]],
        colWidths=[171 * mm],
        rowHeights=[10 * mm, 22 * mm, 70 * mm, 8 * mm, 24 * mm],
    )
    cover.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), INK),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 15 * mm),
                ("RIGHTPADDING", (0, 0), (-1, -1), 15 * mm),
                ("TOPPADDING", (0, 0), (-1, -1), 7 * mm),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.extend([cover, Spacer(1, 10 * mm)])
    story.append(
        info_box(
            "一句话讲清楚",
            "这套工具会锁定同一批新入职劳务工，给每个人相同的观察时长，再结合OTWS派遣日期和真实考勤，判断他是否真正留下。",
        )
    )
    story.extend(
        [
            Spacer(1, 8 * mm),
            p("<b>不只能算90天留存率。</b>观察周期和考勤校验期可以独立选择，也可以输入任意天数。", CALLOUT),
            Spacer(1, 11 * mm),
            data_table(
                ["两个可调参数", "作用", "例子"],
                [
                    ["留存观察周期", "决定每个人经过多少天的考验", "45天、90天、180天"],
                    ["考勤校验期", "决定截止日前向前检查多少天考勤", "仅截止日、近14日、自定义21日"],
                ],
                [42 * mm, 82 * mm, 47 * mm],
            ),
            Spacer(1, 15 * mm),
            p("版本 1.0  |  区域HR、仓库与供应商管理人员使用  |  2026年7月", SMALL),
            PageBreak(),
        ]
    )

    # 1. Principle
    story.extend(section_title(1, "这套工具到底在算什么？"))
    story.extend(
        [
            info_box(
                "核心对象",
                "<b>统计同一批新入职人员，在分别经过相同天数后，还有多少人真正留存。</b>",
            ),
            Spacer(1, 6 * mm),
            p("工具不会统一拿某一天的出勤人数相减，而是给每个人计算自己的观察期截止日。"),
            Spacer(1, 3 * mm),
            formula_box("观察期截止日 = 首次派遣日期 + 留存观察天数"),
            Spacer(1, 6 * mm),
            data_table(
                ["员工", "首次派遣日期", "选择的观察周期", "个人观察期截止日"],
                [
                    ["张三", "1月10日", "90天", "4月10日"],
                    ["李四", "3月20日", "90天", "6月18日"],
                    ["王五", "2月1日", "自定义45天", "3月18日"],
                ],
                [28 * mm, 43 * mm, 48 * mm, 52 * mm],
            ),
            Spacer(1, 6 * mm),
            info_box(
                "为什么更公平？",
                "1月入职和3月入职的人都会经过完全相同的观察天数。不会出现月底入职只接受几天考验，却被当作已留存的问题。",
                BLUE_LIGHT,
                BLUE,
            ),
            Spacer(1, 7 * mm),
            p("<b>留存率公式</b>", H2),
            formula_box("留存率 = 留存人数 ÷（留存人数 + 未留存人数）"),
            Spacer(1, 4 * mm),
            p("“未成熟”“待核验”以及已确认豁免的人员，不进入正式计分分母。", SMALL),
            PageBreak(),
        ]
    )

    # 2. Inputs
    story.extend(section_title(2, "需要准备哪些OTWS数据？"))
    source_rows = [
        ["首次派遣队列", "必传", "确定本次统计的新入职人员", "工号、姓名、首次派遣日期、最晚派遣日期、供应商、区域"],
        ["最晚派遣交叉表", "选传", "检查系统日期和粗暴剔除误伤", "只做交叉校验，不直接决定留存"],
        ["考勤记录", "必传，可多选", "验证工人在现实中是否仍然到岗", "要覆盖所有人员的完整考勤校验期"],
    ]
    story.extend(
        [
            data_table(["数据源", "是否必需", "作用", "重点说明"], source_rows, [36 * mm, 27 * mm, 49 * mm, 59 * mm]),
            Spacer(1, 7 * mm),
            p("<b>有效出勤的统一口径</b>", H2),
        ]
    )
    story.extend(
        bullets(
            [
                "工时大于0。",
                "确认状态为已确认或已复核。",
                "派遣考勤和串岗考勤都可以纳入。",
                "同一工号在同一天有多条记录时，只按1个出勤日计算。",
            ]
        )
    )
    story.extend(
        [
            Spacer(1, 5 * mm),
            info_box(
                "为什么一定要看考勤？",
                "现实中可能出现工人已经不来，但区域没有及时在OTWS结束派遣的情况。因此，系统日期只能证明“账面上还有派遣”，考勤才是“现实中仍在工作”的辅助证据。",
                AMBER_LIGHT,
                AMBER,
            ),
            Spacer(1, 6 * mm),
            info_box(
                "数据安全",
                "工具在本机浏览器内解析和计算文件，不上传员工数据。",
                BLUE_LIGHT,
                BLUE,
            ),
            PageBreak(),
        ]
    )

    # 3. Custom period
    story.extend(section_title(3, "如何选择或自定义留存观察周期？"))
    period_rows = [
        ["7天", "首次派遣日 + 7天"],
        ["两周", "首次派遣日 + 14天"],
        ["一个月", "首次派遣日 + 30天"],
        ["两个月", "首次派遣日 + 60天"],
        ["三个月", "首次派遣日 + 90天"],
        ["半年", "首次派遣日 + 180天"],
        ["一年", "首次派遣日 + 365天"],
        ["自定义", "用户输入任意正整数天数"],
    ]
    story.extend(
        [
            p("页面提供常用周期快捷选项，也可以自定义任意天数。"),
            Spacer(1, 3 * mm),
            data_table(["页面选项", "实际计算方式"], period_rows, [52 * mm, 119 * mm]),
            Spacer(1, 6 * mm),
            info_box(
                "重要边界",
                "页面中的“一个月、三个月、半年、一年”分别按30、90、180和365个自然日计算，不是按自然月月底计算。",
                AMBER_LIGHT,
                AMBER,
            ),
            Spacer(1, 7 * mm),
            p("<b>例子：统计45天留存率</b>", H2),
            step_block(1, "点击“自定义”", "在留存观察周期选项中打开自定义输入框。"),
            Spacer(1, 3 * mm),
            step_block(2, "输入45", "表示每个人从首次派遣日起，经过45个自然日后进行留存检查。"),
            Spacer(1, 3 * mm),
            step_block(3, "应用口径并计算", "工具重新计算每个人的观察期截止日。"),
            Spacer(1, 5 * mm),
            formula_box("观察期截止日 = 首次派遣日 + 45天"),
            PageBreak(),
        ]
    )

    # 4. Custom attendance
    story.extend(section_title(4, "如何选择或自定义考勤校验期？"))
    attendance_rows = [
        ["仅截止日", "只检查观察期截止日当天"],
        ["含截止日近7日", "检查截止日及此前6天"],
        ["含截止日近14日", "检查截止日及此前13天"],
        ["含截止日近30日", "检查截止日及此前29天"],
        ["自定义", "用户输入任意正整数天数"],
    ]
    story.extend(
        [
            p("考勤校验期用来判断：员工在个人观察期截止日之前，最近是否仍有真实有效出勤。"),
            Spacer(1, 3 * mm),
            data_table(["页面选项", "检查范围"], attendance_rows, [63 * mm, 108 * mm]),
            Spacer(1, 6 * mm),
            p("<b>例子：观察期截止日为6月30日</b>", H2),
            data_table(
                ["校验口径", "实际检查日期"],
                [
                    ["仅截止日", "6月30日"],
                    ["近7日", "6月24日至6月30日"],
                    ["近14日", "6月17日至6月30日"],
                    ["近30日", "6月1日至6月30日"],
                    ["自定义21日", "6月10日至6月30日"],
                ],
                [55 * mm, 116 * mm],
            ),
            Spacer(1, 6 * mm),
            info_box(
                "计数方式",
                "校验期包含截止日当天。例如自定义21日，就是从“截止日前20天”检查到截止日当天。",
                BLUE_LIGHT,
                BLUE,
            ),
            Spacer(1, 7 * mm),
            p("<b>两个自定义选项可以自由组合</b>", H2),
            formula_box("例：留存观察周期 = 45天  |  考勤校验期 = 10天"),
            Spacer(1, 4 * mm),
            p("系统先用“首次派遣日+45天”算出每个人的截止日，再检查截止日前10个自然日内的有效出勤。"),
            Spacer(1, 5 * mm),
            info_box(
                "每次改参数后",
                "都要重新点击“应用口径并计算”。观察天数会改变截止日，校验天数则可能改变员工的留存判定。",
                AMBER_LIGHT,
                AMBER,
            ),
            PageBreak(),
        ]
    )

    # 5. Rules
    story.extend(section_title(5, "工具如何逐人判断留存？"))
    rule_rows = [
        ["1", "计算个人截止日", "首次派遣日+观察天数"],
        ["2", "检查数据是否成熟", "考勤数据必须覆盖完整校验期"],
        ["3", "检查最晚派遣日期", "早于个人截止日则判未留存"],
        ["4", "检查最近结束日期", "截止日前已结束，且结束后无有效出勤，判未留存"],
        ["5", "检查有效出勤", "校验期内至少1天有效出勤"],
        ["6", "给出结果", "留存、未留存、未成熟或待核验"],
    ]
    story.extend(
        [
            data_table(["顺序", "检查项", "大白话解释"], rule_rows, [18 * mm, 55 * mm, 98 * mm]),
            Spacer(1, 7 * mm),
            data_table(
                ["最终状态", "什么情况会进入"],
                [
                    ["留存", "派遣日期覆盖截止日，且校验期内有有效出勤"],
                    ["未留存", "最晚派遣早于截止日，或截止日前已结束，或校验期无有效出勤"],
                    ["未成熟", "员工的截止日或完整校验期超出已上传考勤的覆盖范围"],
                    ["待核验", "缺少关键日期，或数据之间存在无法自动消除的冲突"],
                    ["豁免", "业务核实后确认为我方减量、批准休假或其他非供应商责任情形"],
                ],
                [35 * mm, 136 * mm],
            ),
            Spacer(1, 7 * mm),
            info_box(
                "实务处理",
                "系统显示派遣覆盖截止日，但校验期内无有效出勤的人员，本工具默认暂按未留存计，同时自动进入异常复核。区域可以用休假、排班或派遣证据进行人工改判。",
                RED_LIGHT,
                RED,
            ),
            PageBreak(),
        ]
    )

    # 6. Coverage and workflow
    story.extend(section_title(6, "实际怎么操作？"))
    workflow = [
        (1, "确定统计队列", "例如首次派遣日期在1月1日至3月31日的员工。"),
        (2, "上传首次派遣队列", "这张表决定本次要计算哪些人。"),
        (3, "按需上传最晚派遣交叉表", "用来检查日期异常和粗暴剔除误伤。"),
        (4, "选择留存观察周期", "使用快捷选项，或点击自定义并输入天数。"),
        (5, "选择考勤校验期", "可选仅截止日、近7/14/30日，或自定义。"),
        (6, "查看所需考勤覆盖范围", "页面会根据两个参数自动计算需要的考勤日期。"),
        (7, "上传全部相关月度考勤", "可以一次多选，也可以按月追加。"),
        (8, "读取并校验数据", "工具检查字段、日期覆盖、重复工号和考勤质量。"),
        (9, "应用口径并计算", "生成国家、供应商和人员三个层级的结果。"),
        (10, "处理异常复核并导出", "保留人工调整结果和证据说明，再导出正式报表。"),
    ]
    for index, title, body in workflow:
        story.extend([step_block(index, title, body), Spacer(1, 2.5 * mm)])
    story.extend(
        [
            Spacer(1, 4 * mm),
            info_box(
                "最容易漏掉的一步",
                "先选参数，再根据页面提示准备考勤文件。不要在只有6月考勤的情况下，直接计算截止日到9月的180天留存率。",
                AMBER_LIGHT,
                AMBER,
            ),
            PageBreak(),
        ]
    )

    # 7. Review and talk track
    story.extend(section_title(7, "异常复核与区域宣导话术"))
    review_rows = [
        ["系统覆盖截止日，但校验期无有效出勤", "核实批准休假、未排班或系统漏结束"],
        ["最晚派遣和最近结束日期冲突", "回溯OTWS派遣删除、结束和重新派遣记录"],
        ["日期恰好落在观察期截止日", "确认当日是否实际工作，以及公司的日期边界口径"],
        ["校验期存在未确认正工时", "先完成考勤确认或复核，再重新计算"],
        ["校验期出现多个供应商", "核实是否发生雇主转换，确认原供应商的归责"],
        ["考勤供应商与入职供应商不一致", "确认供应商转换和是否需要豁免"],
    ]
    story.extend(
        [
            data_table(["自动进入异常复核的情况", "区域需要做什么"], review_rows, [86 * mm, 85 * mm]),
            Spacer(1, 7 * mm),
            p("<b>建议对区域同事这样介绍</b>", H2),
            info_box(
                "可直接念的话术",
                "“这套工具不是拿两天的出勤总人数相减。它会锁定同一批新入职劳务工，根据我们选择的观察天数，给每个人计算自己的截止日。然后再结合OTWS派遣日期和截止日前的真实考勤，判断他是否真正留下。观察周期和考勤校验期都可以自定义。”",
                TEAL_LIGHT,
                TEAL,
            ),
            Spacer(1, 7 * mm),
            p("<b>开始使用前的最后检查</b>", H2),
        ]
    )
    checklist = [
        ["□", "队列中的首次派遣日期范围是否正确？"],
        ["□", "是否已经确认本次要看7天、45天、90天还是其他周期？"],
        ["□", "考勤校验期是否符合区域排班和休假特点？"],
        ["□", "已上传考勤是否覆盖页面提示的完整日期范围？"],
        ["□", "异常复核的人工调整是否留下证据和说明？"],
    ]
    checklist_table = Table([[p(a, H2), p(b, BODY)] for a, b in checklist], colWidths=[12 * mm, 159 * mm])
    checklist_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), SURFACE_SOFT),
                ("BOX", (0, 0), (-1, -1), 0.5, LINE),
                ("INNERGRID", (0, 0), (-1, -1), 0.3, LINE),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 7),
                ("RIGHTPADDING", (0, 0), (-1, -1), 7),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.extend(
        [
            checklist_table,
            Spacer(1, 7 * mm),
            p("本工具的价值不在于“得到一个百分比”，而在于让每个人经过同样长的观察，并且让每一个判定都能回到个人、日期和考勤证据。", CALLOUT),
        ]
    )
    return story


def main():
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    document = GuideDocTemplate(str(OUTPUT))
    document.build(build_story())
    print(OUTPUT)


if __name__ == "__main__":
    main()
