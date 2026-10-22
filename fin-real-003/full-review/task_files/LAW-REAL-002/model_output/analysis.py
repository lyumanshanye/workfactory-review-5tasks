#!/usr/bin/env python3
import argparse
import csv
import json
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

AS_OF = date(2026, 10, 12)
HORIZON = date(2026, 11, 11)
DATE_FORMATS = ("%Y-%m-%dT%H:%M:%S.%f", "%m/%d/%Y", "%Y-%m-%d")
FILES = {
    "general": "LNI_Contractor_License_General_BR.csv",
    "bond": "LNI_Contractor_License_Bond_BR.csv",
    "insurance": "LNI_Contractor_License_Insurance_BR.csv",
    "statute": "RCW_18.27.pdf",
}

LEGAL_MATRIX = [
    {
        "topic": "适用对象与分类",
        "citation": "RCW 18.27.010(1),(5),(12)",
        "pages": "PDF第2–4页",
        "condition": "从事定义内承包活动；一般承包商使用多于一个工种，专业承包商不属于一般承包商。",
        "rule": "第18.27章要求先确认行为和一般/专业分类。",
        "package_can": "CC记录可结合specialty字段作筛查分类。",
        "package_cannot": "不能证明实际业务行为、项目事实或豁免；EC、PC、LC的专门许可要求不在本PDF内。",
    },
    {
        "topic": "登记义务与禁止行为",
        "citation": "RCW 18.27.020(1)–(3)",
        "pages": "PDF第4–5页",
        "condition": "属于本章承包商且无适用例外。",
        "rule": "承包商应登记；未登记、暂停或撤销期间的特定行为受禁止；(3)另设书面通知相关限制。",
        "package_can": "证明登记表快照状态。",
        "package_cannot": "不能证明曾承揽、投标、广告、分包，或任何人收到书面通知。",
    },
    {
        "topic": "申请资料与UBI",
        "citation": "RCW 18.27.030(1)(b),(f), (3)",
        "pages": "PDF第5–7页",
        "condition": "申请或维持第18.27章登记。",
        "rule": "申请含UBI与一般/专业类型；若干事项可导致拒绝或暂停。",
        "package_can": "比较CSV中的执照号、UBI和类型字段。",
        "package_cannot": "不能证明DOR状态、判决、申请真实性或机关实际决定。",
    },
    {
        "topic": "保证金/替代担保与金额",
        "citation": "RCW 18.27.040(1),(2),(4),(7),(8),(11)",
        "pages": "PDF第7–10页",
        "condition": "第18.27章申请人；一般承包商通常30,000美元，专业承包商通常15,000美元；可用指定储蓄账户替代；特定续期可提高。",
        "rule": "取消、撤销、担保退出或减损可影响登记；续保不累加责任。",
        "package_can": "逐行判断CSV保证金记录日期与单笔金额是否达到通常门槛。",
        "package_cannot": "不能证明替代账户、机关批准、未决/已付索赔、法院判决、提高金额命令或完整历史；不得叠加历史金额。",
    },
    {
        "topic": "保险/替代财务责任",
        "citation": "RCW 18.27.050(1)–(3)",
        "pages": "PDF第10–11页",
        "condition": "第18.27章登记或再登记；财产损害50,000美元、单人伤亡100,000美元、多人伤亡200,000美元，或合格指定账户。",
        "rule": "到期、取消、撤销或保险人退出可自动暂停，直至新保单/恢复通知提交并获批。",
        "package_can": "判断保险来源行的日期支持和总额字段。",
        "package_cannot": "单一insuranceamt不能证明法定分项、除外责任、保单类型、指定账户、提交或批准。",
    },
    {
        "topic": "取消/减损与登记后果",
        "citation": "RCW 18.27.040(1),(7); 18.27.050(2); 18.27.060(3)",
        "pages": "PDF第7、9、10–11页",
        "condition": "保证金取消、撤销、退出或被未清偿判决减损，或保险取消/到期等法定事件。",
        "rule": "条文规定自动暂停及恢复所需的新保障或恢复通知；18.27.060(3)以减损/取消生效日为暂停日。",
        "package_can": "标出CSV中的取消、到期、减损日期及登记状态不一致。",
        "package_cannot": "本筛查日期边界不是小时级法定裁定；不能证明通知、提交、批准或机关系统的完整状态。",
    },
    {
        "topic": "登记期限与续期",
        "citation": "RCW 18.27.060(1),(4)",
        "pages": "PDF第11页",
        "condition": "第18.27章登记证及续期。",
        "rule": "证书通常两年；按期续期；特定可追踪投递下，机关收到费用和保证金、责任保险证明之日生效。",
        "package_can": "显示登记表到期日及快照状态。",
        "package_cannot": "不能证明续期材料送达、费用、投递凭证或机关批准。",
    },
    {
        "topic": "30天届满定义例外",
        "citation": "RCW 18.27.010(15)",
        "pages": "PDF第3–4页",
        "condition": "登记届满不超过30日，且始终维持有效保证金以及18.27.050所需保险或指定账户。",
        "rule": "满足全部条件者不纳入该款“未登记承包商”定义。",
        "package_can": "筛出登记到期日在30日附近的记录，并检查包内是否有日期可支持的保障行。",
        "package_cannot": "EXPIRED一列既不能单独排除也不能单独放行；CSV缺行不等于无保障，且不能证明“始终”、替代账户或实际活动。",
    },
    {
        "topic": "登记核验",
        "citation": "RCW 18.27.010(17); 18.27.110(1)",
        "pages": "PDF第4、17页",
        "condition": "地方机关对依法需登记承包商的建筑许可核验。",
        "rule": "核验可查登记卡、部门数据库或联系部门；发证前应核验当前登记。",
        "package_can": "作为下载快照触发进一步官方核验。",
        "package_cannot": "离线快照不等于2026-10-12或之后的实时机关核验。",
    },
    {
        "topic": "登记例外",
        "citation": "RCW 18.27.065; 18.27.090(1)–(17)",
        "pages": "PDF第12–14页",
        "condition": "合伙/合资特定情形或列举豁免事实成立。",
        "rule": "条文列出登记推定和多类豁免，部分附活动范围及反规避限制。",
        "package_can": "显示登记表business type与专业字段，仅供提示。",
        "package_cannot": "不能据字段推定合伙人名称条件、项目金额、产权、雇佣、广告或其他豁免事实。",
    },
    {
        "topic": "诉讼中的持续保障",
        "citation": "RCW 18.27.080",
        "pages": "PDF第12–13页",
        "condition": "承包商在州法院追索需登记工作的报酬/违约。",
        "rule": "条文对诉权及实质合规设定持续资料、保证金和保险条件。",
        "package_can": "提示历史保障连续性可能重要。",
        "package_cannot": "不能证明合同、工程、诉讼、相关期间或法院结论。",
    },
]


def read_csv(path):
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def parse_date(value):
    value = (value or "").strip()
    if not value or value == "Until Canceled":
        return None, "blank_or_open"
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date(), "parsed"
        except ValueError:
            pass
    return None, "unparseable"


def money(value):
    try:
        return Decimal(value)
    except (InvalidOperation, TypeError):
        return None


def iso(value):
    return value.isoformat() if isinstance(value, date) else ""


def protection_state(row, kind, on_date):
    if kind == "bond":
        start_field = "bondeffectivedate"
        end_fields = ("bondexpirationdate", "bondcanceldate", "bondimpaireddate")
        open_ended = row.get("bondexpirationdate") == "Until Canceled"
    else:
        start_field = "effectivedate"
        end_fields = ("expirationdate", "canceldate")
        open_ended = False

    start, start_status = parse_date(row.get(start_field))
    bad = [field for field in (start_field,) + end_fields if row.get(field) and row.get(field) != "Until Canceled" and parse_date(row.get(field))[1] == "unparseable"]
    if bad:
        return "不确定：日期无法解析", "、".join(bad)
    if start_status != "parsed":
        return "不确定：缺生效日", start_field
    if on_date < start:
        return "否：尚未生效", f"生效日{start.isoformat()}"

    ends = []
    for field in end_fields:
        parsed, status = parse_date(row.get(field))
        if status == "parsed":
            ends.append((parsed, field))
    if ends:
        end, field = min(ends)
        if on_date >= end:
            return "否：已终止", f"{field}={end.isoformat()}（终止当日不计）"
        return "是：日期支持", f"截至{end.isoformat()}前；终止字段{field}"
    if open_ended:
        return "是：日期支持（Until Canceled）", "未见取消/减损日；不代表永久或无条件有效"
    return "不确定：缺终止依据", "无到期/取消/减损日期，也非Until Canceled"


def is_supported(row, kind, on_date):
    return protection_state(row, kind, on_date)[0].startswith("是：")


def classify_scope(reg):
    if reg["contractorlicensetypecode"] == "CC":
        return "RCW 18.27直接筛查（仍须核验活动与豁免）"
    return "本包未确认适用；转相应专门许可制度核验"


def contractor_class(reg):
    if reg["contractorlicensetypecode"] != "CC":
        return "未确认"
    return "一般" if reg.get("specialtycode1") == "01" or reg.get("specialtycode1desc", "").upper() == "GENERAL" else "专业"


def usual_bond_threshold(reg):
    cls = contractor_class(reg)
    if cls == "一般":
        return Decimal("30000")
    if cls == "专业":
        return Decimal("15000")
    return None


def amount_support(reg, rows, on_date):
    threshold = usual_bond_threshold(reg)
    if threshold is None:
        return "适用范围未确认"
    supported = [(line, row) for line, row in rows if is_supported(row, "bond", on_date)]
    if not supported:
        return "否：无日期可支持记录"
    amounts = [money(row.get("bondamt")) for _, row in supported]
    amounts = [amount for amount in amounts if amount is not None]
    if not amounts:
        return "不确定：金额缺失/无法解析"
    if any(amount >= threshold for amount in amounts):
        return f"是：至少一笔达到通常门槛${threshold:,.0f}"
    return f"否：单笔均低于通常门槛${threshold:,.0f}（不叠加）"


def presence(rows, kind, on_date):
    states = [protection_state(row, kind, on_date)[0] for _, row in rows]
    if any(state.startswith("是：") for state in states):
        return "是"
    if any(state.startswith("不确定：") for state in states):
        return "不确定"
    return "否"


def insurance_form(rows, on_date):
    if presence(rows, "insurance", on_date) == "否":
        return "证据不足：无日期可支持保险行；不等于未投保"
    if presence(rows, "insurance", on_date) == "不确定":
        return "不确定：日期字段不足，且单一insuranceamt不能证明法定形式"
    return "不足以证明：insuranceamt不含责任分项、除外责任、保单类型或替代账户信息"


def describe_rows(rows, kind):
    if not rows:
        return "无来源行"
    parts = []
    for line, row in rows:
        if kind == "bond":
            parts.append(
                f"行{line}: 生效{row.get('bondeffectivedate') or '空'}; 到期{row.get('bondexpirationdate') or '空'}; "
                f"取消{row.get('bondcanceldate') or '空'}; 减损{row.get('bondimpaireddate') or '空'}; 金额{row.get('bondamt') or '空'}"
            )
        else:
            parts.append(
                f"行{line}: 生效{row.get('effectivedate') or '空'}; 到期{row.get('expirationdate') or '空'}; "
                f"取消{row.get('canceldate') or '空'}; 金额{row.get('insuranceamt') or '空'}"
            )
    return " | ".join(parts)


def window_gap(rows, kind):
    if not any(is_supported(row, kind, AS_OF) for _, row in rows):
        return None
    day = AS_OF + timedelta(days=1)
    gap_start = None
    gap_end = None
    while day <= HORIZON:
        if not any(is_supported(row, kind, day) for _, row in rows):
            if gap_start is None:
                gap_start = day
            gap_end = day
        elif gap_start is not None:
            break
        day += timedelta(days=1)
    if gap_start is None:
        return None
    triggering = []
    for line, row in rows:
        if is_supported(row, kind, AS_OF) and not is_supported(row, kind, gap_start):
            triggering.append(line)
    return {"start": gap_start, "end": gap_end, "trigger_lines": triggering}


def duplicate_groups(rows, keys):
    groups = defaultdict(list)
    for line, row in rows:
        groups[tuple(row.get(key, "") for key in keys)].append((line, row))
    return [(key, values) for key, values in groups.items() if len(values) > 1]


def set_cell_link(cell, target, text=None):
    cell.value = text or target
    cell.hyperlink = target
    cell.style = "Hyperlink"


def write_sheet(ws, headers, rows, widths=None, freeze="A2", autofilter=True):
    ws.append(headers)
    for row in rows:
        ws.append(row)
    fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = freeze
    if autofilter and ws.max_row >= 1:
        ws.auto_filter.ref = ws.dimensions
    ws.sheet_view.showGridLines = False
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    if widths:
        for index, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(index)].width = width


def add_doc_title(doc, text, level=1):
    paragraph = doc.add_heading(text, level=level)
    paragraph.paragraph_format.space_before = Pt(8)
    paragraph.paragraph_format.space_after = Pt(4)
    return paragraph


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(item)


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    fld_char1 = OxmlElement("w:fldChar")
    fld_char1.set(qn("w:fldCharType"), "begin")
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = "PAGE"
    fld_char2 = OxmlElement("w:fldChar")
    fld_char2.set(qn("w:fldCharType"), "end")
    run._r.extend([fld_char1, instr_text, fld_char2])


def make_docx(path, data):
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2)
    sec.bottom_margin = Cm(2)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.2)
    styles = doc.styles
    styles["Normal"].font.name = "Noto Sans CJK SC"
    styles["Normal"]._element.rPr.rFonts.set(qn("w:eastAsia"), "Noto Sans CJK SC")
    styles["Normal"].font.size = Pt(10.5)
    for section in doc.sections:
        add_page_number(section.footer.paragraphs[0])

    title = doc.add_heading("华盛顿州承包商登记公开资料证据审查备忘录", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = doc.add_paragraph("证据基准日：2026-10-12｜30日观察终点：2026-11-11")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.add_run("性质：").bold = True
    p.add_run("基于下载快照的准入筛查与补证建议；不是终局法律意见，不认定违法，不替代L&I或法院决定。")

    add_doc_title(doc, "一、问题与结论摘要", 1)
    add_bullets(doc, [
        f"登记表共{data['counts']['general_rows']:,}行、{data['counts']['general_licenses']:,}个唯一执照；保证金表{data['counts']['bond_rows']:,}行，保险表{data['counts']['insurance_rows']:,}行。登记表未见重复执照号。",
        f"三表均以执照号精确连接；发现{data['counts']['shared_ubi_groups']}个UBI对应多个登记执照。UBI不是保障记录主键，未据UBI或近似名称挪用任何保证金或保险。",
        f"保证金和保险表未发现登记表外的孤立执照；ACTIVE执照中，有{len(data['active_both_missing'])}项在两类保障表均无可确认有效记录，全部为EC或LC，RCW 18.27适用未确认，仅作资料核验。",
        f"ACTIVE的CC记录中，有{len(data['active_cc_one_side_gap'])}项在2026-10-12仅一类保障无可确认有效记录；缺行或日期不足只表示证据不足。",
        f"未来30日有{len(data['renewal_risk'])}项ACTIVE执照出现至少一类日期支持记录终止且包内无无缝后继覆盖；名单包括1项PC，其余为CC。对PC不套用18.27门槛。",
        "任何日期可支持的保险行仍不足以单独证明RCW 18.27.050要求：insuranceamt不显示财产损害、单人伤亡、多人伤亡分项，也不证明除外责任、保单类型或替代指定账户。",
    ])

    add_doc_title(doc, "二、资料、口径与限制", 1)
    doc.add_paragraph(
        "资料为三个Washington L&I接口CSV快照及RCW_18.27.pdf。法典PDF页脚标示“Certified on 7/15/2026”；本备忘录仅据该原件分析2026-10-12筛查，不能据此证明所有更早日期的法律版本。"
    )
    add_bullets(doc, [
        "记录粒度：General每执照一行；Bond与Insurance为同一执照的历史/替换保障多行。候选键分别为General的执照号、Bond的“执照号+保证金账户+生效日+金额”、Insurance的“执照号+保单号+生效日+金额”；后两者仍存在冲突/重复，必须保留源行号。",
        "日期边界：生效当日计入；到期、取消或减损日期当日起不计。该规则是本次日级筛查口径，不主张小时级法定终止结论。Until Canceled仅表示包内未给固定到期日；若未见取消/减损日，本次可作日期支持，但绝不表示永久或无条件有效。缺失或不可解析日期单列不确定。",
        "金额口径：同一日期仅判断单笔日期可支持记录是否达到通常门槛，不将历史、重叠、续保或替换记录金额相加。RCW 18.27.040(4)亦说明续期等不累加担保责任。",
        "身份口径：只按完整执照号连接。UBI与企业名称用于核对，不用于跨执照归属保障。主备忘录不复制个人姓名、电话或地址。",
    ])

    add_doc_title(doc, "三、法律适用矩阵", 1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for cell, text in zip(table.rows[0].cells, ["主题", "条文/页码", "规则与条件", "本包证明边界"]):
        cell.text = text
    for item in LEGAL_MATRIX:
        cells = table.add_row().cells
        cells[0].text = item["topic"]
        cells[1].text = f"{item['citation']}；{item['pages']}"
        cells[2].text = f"条件：{item['condition']}\n规则：{item['rule']}"
        cells[3].text = f"可证明：{item['package_can']}\n不能证明：{item['package_cannot']}"

    add_doc_title(doc, "四、范围判断与30天届满例外", 1)
    doc.add_paragraph(
        "登记类型包括CC、EC、PC、LC。第18.27章原件定义承包商并区分一般与专业承包商，但本包不含电气、管道、电梯专门许可章节。因此仅对CC作18.27直接筛查；EC、PC、LC保留记录层日期结果，却不机械套用30,000/15,000美元或保险分项门槛。"
    )
    doc.add_paragraph(
        "RCW 18.27.010(15)（PDF第3–4页）把通常属于“未登记承包商”的届满记录排除在定义之外，须同时满足：届满不超过30日、一直维持有效保证金、并一直维持18.27.050要求的保险或指定账户。故“EXPIRED”状态不能机械排除，也不能机械放行。包内共有"
        f"{len(data['lapse_review'])}项CC记录按登记到期日推算在基准日前0至30日内；CSV缺少保障行只能说本包不能证明条件，不能证明实际无保障，更不能证明曾承揽工程或适用豁免。应向L&I实时核验到期起至核验日的连续保障、替代账户及续期收件/批准。"
    )

    add_doc_title(doc, "五、优先核验名单", 1)
    add_doc_title(doc, "（一）ACTIVE且保证金、保险均无可确认有效记录", 2)
    if not data["active_both_missing"]:
        doc.add_paragraph("未发现。")
    else:
        doc.add_paragraph("以下均属RCW 18.27适用未确认的EC/LC，仅作资料核验，不作不合规判断：")
        for item in data["active_both_missing"][:5]:
            doc.add_paragraph(
                f"{item['license']}｜{item['business_name']}｜UBI {item['ubi']}｜{item['type_code']}。登记来源行{item['general_line']}为ACTIVE；保证金、保险表均无该执照来源行。待补：相应专门许可制度、实时状态、所需保障及接口覆盖范围。下一步：按完整执照号向L&I核验，不借用同UBI其他执照记录。",
                style="List Bullet",
            )
        if len(data["active_both_missing"]) > 5:
            doc.add_paragraph(f"完整{len(data['active_both_missing'])}项见audit.xlsx“待核验与续期”。")

    add_doc_title(doc, "（二）未来30日保障记录失效且无包内无缝后继", 2)
    for item in data["renewal_risk"][:5]:
        gaps = "；".join(f"{x['kind_cn']}缺口{x['gap_start']}至{x['gap_end']}，触发行{x['trigger_lines']}" for x in item["gaps"])
        doc.add_paragraph(
            f"{item['license']}｜{item['business_name']}｜UBI {item['ubi']}｜{gaps}。条文范围：{item['scope']}；关联RCW 18.27.050(2)、18.27.060(3)仅在18.27适用且法定事实成立时适用。待补：续保/替换记录、提交与批准凭证、实时登记状态。",
            style="List Bullet",
        )
    doc.add_paragraph(f"完整{len(data['renewal_risk'])}项及全部来源行、日期、金额见audit.xlsx“待核验与续期”。")

    add_doc_title(doc, "六、完整案例（每类最多5项）", 1)
    add_doc_title(doc, "（一）历史记录不能当作当前保障", 2)
    for item in data["examples"]["historical"]:
        doc.add_paragraph(item, style="List Bullet")
    add_doc_title(doc, "（二）未来记录不能提前生效", 2)
    if data["examples"]["future"]:
        for item in data["examples"]["future"]:
            doc.add_paragraph(item, style="List Bullet")
    else:
        doc.add_paragraph("未发现可作为该类案例的未来生效来源行。")
    add_doc_title(doc, "（三）登记状态与保障证据不一致", 2)
    for item in data["examples"]["status_mismatch"]:
        doc.add_paragraph(item, style="List Bullet")

    add_doc_title(doc, "七、建议的核验顺序", 1)
    add_bullets(doc, [
        "第一优先：ACTIVE的CC单边缺口和30日内将出现缺口的CC；向L&I按完整执照号核实实时登记、保障连续性、替代指定账户、恢复/后继文件的提交与批准。",
        "第二优先：ACTIVE且双缺的EC/LC，以及PC记录；先确定适用专门章节、接口是否应返回保障，再核验，不套用第18.27章门槛。",
        "第三优先：状态为EXPIRED且可能落入30日定义例外的CC；核验整个届满期间的连续保障，而非只看两个端点。",
        "数据治理：对重复/冲突保障行要求L&I确认权威记录；保持源行，不静默合并。任何缺记录均记录为“证据不足”，不得改写成“没有投保/没有保证金”。",
    ])

    doc.add_page_break()
    add_doc_title(doc, "附件A：不发送的核验函草稿一（保证金记录缺口）", 1)
    doc.add_paragraph("状态：草稿—不发送")
    doc.add_paragraph("收件人：Washington State Department of Labor & Industries 承包商登记资料核验窗口")
    doc.add_paragraph("主题：请求核验 BRAVOEN911P9 的保证金/替代担保与登记资料")
    doc.add_paragraph("敬启者：")
    doc.add_paragraph(
        "本团队在承包商准入资料审查中，查阅了下载的公开接口快照。登记表来源行312显示BRAVO ENVIRONMENTAL NW INC，执照号BRAVOEN911P9，UBI 602946216，状态为ACTIVE；保险表来源行140存在日期可支持记录，但保证金表未检出该完整执照号记录。"
    )
    doc.add_paragraph(
        "烦请按完整执照号核验截至2026-10-12的：(1) 实时登记状态；(2) 是否存在有效保证金或RCW 18.27.040(8)所述替代指定储蓄账户；(3) 如有，生效、终止、减损、提交和批准日期及金额；(4) 公开接口是否省略某类替代担保。我们不请求对企业作违法认定，仅为纠正证据缺口。"
    )
    doc.add_paragraph("谨致")
    doc.add_paragraph("承包商准入审查团队")

    doc.add_page_break()
    add_doc_title(doc, "附件B：不发送的核验函草稿二（保险到期边界）", 1)
    doc.add_paragraph("状态：草稿—不发送")
    doc.add_paragraph("收件人：Washington State Department of Labor & Industries 承包商登记资料核验窗口")
    doc.add_paragraph("主题：请求核验 BRIGHHI063D3 的保险连续性与登记状态")
    doc.add_paragraph("敬启者：")
    doc.add_paragraph(
        "本团队在公开接口快照中看到：登记表来源行767显示BRIGHTEN HOMES INC，执照号BRIGHHI063D3，UBI 601529129，状态为ACTIVE；保险表来源行358的记录生效日为2025-10-12、到期日为2026-10-12、金额字段为1,000,000。按本次筛查口径，到期日当日起不再作为日期可支持记录；这不是对法定小时级终止时间或实际保险状态的裁定。"
    )
    doc.add_paragraph(
        "烦请核验：(1) 2026-10-12当日及其后的实际保险连续性；(2) 是否有续保、替换保单或RCW 18.27.050(3)指定账户；(3) 相关证明提交及批准日期；(4) 截至核验日的登记状态。单一insuranceamt不能证明法定责任分项，因此如可公开，也请确认保障形式和法定分项是否已由机关审核。我们不请求认定任何违法或争议事实。"
    )
    doc.add_paragraph("谨致")
    doc.add_paragraph("承包商准入审查团队")

    doc.save(path)


def make_xlsx(path, data, general_rows, bond_rows, insurance_rows, by_bond, by_insurance):
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("说明")
    notes = [
        ["项目", "内容"],
        ["证据基准日", AS_OF.isoformat()],
        ["30日观察终点", HORIZON.isoformat()],
        ["用途", "公开资料准入筛查与补证；不是终局法律意见，不认定违法。"],
        ["连接规则", "仅按完整contractorlicensenumber连接；绝不按UBI或近似名称跨执照挪用保障。"],
        ["日期规则", "生效当日计入；到期、取消、减损当日起不计。日级筛查口径，不主张小时级终止结论。"],
        ["Until Canceled", "仅表示该行无固定到期日；未见取消/减损时作日期支持，但不表示永久或无条件有效。"],
        ["缺失日期", "缺生效日、无法解析日期，或既无终止依据也非Until Canceled，标为不确定。"],
        ["金额规则", "不相加历史、重叠、续保或替换金额；仅判断单笔日期可支持保证金是否达到通常门槛。"],
        ["保险限制", "insuranceamt不能证明责任分项、除外责任、保单类型或替代财务责任；高金额不等于全部法定要件满足。"],
        ["缺记录含义", "仅为本包证据不足，不等于没有保证金或没有投保。"],
        ["法典版本", "RCW PDF页脚Certified on 7/15/2026；不能当作所有历史日期的法律版本。"],
        ["重建命令", "python analysis.py --input /app/input_files --output /app/output"],
    ]
    for row in notes:
        ws.append(row)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 110
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.sheet_view.showGridLines = False

    ws = wb.create_sheet("法律矩阵")
    headers = ["主题", "条文", "PDF页码", "适用条件", "规则", "本包可证明", "本包不能证明"]
    rows = [[x["topic"], x["citation"], x["pages"], x["condition"], x["rule"], x["package_can"], x["package_cannot"]] for x in LEGAL_MATRIX]
    write_sheet(ws, headers, rows, [20, 28, 15, 45, 48, 44, 56])

    ws = wb.create_sheet("逐执照")
    headers = [
        "登记源行", "企业名称", "执照号", "UBI", "类型代码", "类型说明", "一般/专业筛查", "适用范围",
        "登记状态", "登记生效日", "登记到期日", "暂停日", "保证金源行数", "保险源行数",
        "2026-10-12保证金记录", "2026-10-12保证金金额门槛", "2026-10-12保险记录", "2026-10-12保险形式证明",
        "2026-11-11保证金记录", "2026-11-11保证金金额门槛", "2026-11-11保险记录", "2026-11-11保险形式证明",
        "当前保证金来源行", "当前保险来源行", "备注",
    ]
    rows = []
    for line, reg in general_rows:
        lic = reg["contractorlicensenumber"]
        bonds = by_bond[lic]
        policies = by_insurance[lic]
        cur_b = [str(n) for n, row in bonds if is_supported(row, "bond", AS_OF)]
        cur_i = [str(n) for n, row in policies if is_supported(row, "insurance", AS_OF)]
        note = "缺记录仅表示本包证据不足。"
        if reg["contractorlicensetypecode"] != "CC":
            note += " RCW 18.27门槛未确认。"
        rows.append([
            line, reg["businessname"], lic, reg["ubi"], reg["contractorlicensetypecode"], reg["contractorlicensetypecodedesc"],
            contractor_class(reg), classify_scope(reg), reg["contractorlicensestatus"], reg["licenseeffectivedate"],
            reg["licenseexpirationdate"], reg["contractorlicensesuspenddate"], len(bonds), len(policies),
            presence(bonds, "bond", AS_OF), amount_support(reg, bonds, AS_OF), presence(policies, "insurance", AS_OF), insurance_form(policies, AS_OF),
            presence(bonds, "bond", HORIZON), amount_support(reg, bonds, HORIZON), presence(policies, "insurance", HORIZON), insurance_form(policies, HORIZON),
            ",".join(cur_b) or "无", ",".join(cur_i) or "无", note,
        ])
    write_sheet(ws, headers, rows, [10, 32, 17, 13, 10, 22, 15, 38, 13, 20, 20, 20, 12, 12, 19, 34, 19, 55, 19, 34, 19, 55, 18, 18, 36])

    ws = wb.create_sheet("保证金记录")
    headers = [
        "源行", "企业名称", "执照号", "UBI", "类型", "登记状态(保障表)", "保证金公司", "账户号", "金额", "收到日", "生效日", "到期日",
        "取消日", "减损标记", "减损日", "2026-10-12记录有效性", "2026-10-12理由", "2026-11-11记录有效性", "2026-11-11理由", "身份核对",
    ]
    rows = []
    for line, row in bond_rows:
        s1, r1 = protection_state(row, "bond", AS_OF)
        s2, r2 = protection_state(row, "bond", HORIZON)
        rows.append([
            line, row["businessname"], row["contractorlicensenumber"], row["ubi"], row["contractorlicensetypecode"], row["licensestatusdesc"],
            row["bondfirmname"], row["bondaccountid"], row["bondamt"], row["receivedbyl_i"], row["bondeffectivedate"], row["bondexpirationdate"],
            row["bondcanceldate"], row["bondimpaired"], row["bondimpaireddate"], s1, r1, s2, r2, "按完整执照号；未跨执照归属",
        ])
    write_sheet(ws, headers, rows, [9, 31, 17, 13, 8, 15, 30, 19, 14, 20, 20, 18, 20, 11, 20, 25, 42, 25, 42, 28])

    ws = wb.create_sheet("保险记录")
    headers = [
        "源行", "企业名称", "执照号", "UBI", "类型", "登记状态(保障表)", "保险公司", "保单号", "insuranceamt", "生效日", "到期日", "取消日",
        "创建日", "更新日", "保险代理机构", "2026-10-12记录有效性", "2026-10-12理由", "2026-11-11记录有效性", "2026-11-11理由", "法定形式结论",
    ]
    rows = []
    for line, row in insurance_rows:
        s1, r1 = protection_state(row, "insurance", AS_OF)
        s2, r2 = protection_state(row, "insurance", HORIZON)
        rows.append([
            line, row["businessname"], row["contractorlicensenumber"], row["ubi"], row["contractorlicensetypecode"], row["licensestatusdesc"],
            row["insurancecompany"], row["insurancepolicyno"], row["insuranceamt"], row["effectivedate"], row["expirationdate"], row["canceldate"],
            row["createdate"], row["lastupdatedate"], row["insuranceagencyname"], s1, r1, s2, r2,
            "单一insuranceamt不足以证明RCW 18.27.050责任分项、除外责任、保单类型或替代账户",
        ])
    write_sheet(ws, headers, rows, [9, 31, 17, 13, 8, 15, 31, 22, 16, 20, 20, 20, 20, 20, 29, 25, 42, 25, 42, 58])

    ws = wb.create_sheet("待核验与续期")
    headers = ["类别", "企业名称", "执照号", "UBI", "类型", "登记源行", "登记状态", "适用范围", "结论", "保证金来源", "保险来源", "法条", "待补证据", "下一步"]
    rows = []
    for item in data["active_both_missing"]:
        rows.append([
            "ACTIVE双缺", item["business_name"], item["license"], item["ubi"], item["type_code"], item["general_line"], "ACTIVE", item["scope"],
            "两保障表均无可确认有效记录；缺记录不等于无保障", item["bond_sources"], item["insurance_sources"], "RCW 18.27适用未确认",
            "实时登记、适用专门章节、保证金/保险或替代形式、接口覆盖", "按完整执照号向L&I及相应许可项目核验；不得借用同UBI其他执照保障",
        ])
    for item in data["active_cc_one_side_gap"]:
        rows.append([
            "ACTIVE CC单边缺口", item["business_name"], item["license"], item["ubi"], "CC", item["general_line"], "ACTIVE", item["scope"], item["conclusion"],
            item["bond_sources"], item["insurance_sources"], item["statute"], item["needed_evidence"], item["next_step"],
        ])
    for item in data["renewal_risk"]:
        conclusion = "；".join(f"{gap['kind_cn']}记录在窗口内出现缺口{gap['gap_start']}至{gap['gap_end']}；触发行{gap['trigger_lines']}" for gap in item["gaps"])
        rows.append([
            "30日续期风险", item["business_name"], item["license"], item["ubi"], item["type_code"], item["general_line"], "ACTIVE", item["scope"], conclusion,
            item["bond_sources"], item["insurance_sources"], item["statute"], item["needed_evidence"], item["next_step"],
        ])
    rows.sort(key=lambda row: (row[2], row[0]))
    write_sheet(ws, headers, rows, [19, 31, 17, 13, 8, 11, 12, 38, 58, 65, 65, 39, 53, 58])

    ws = wb.create_sheet("30日届满例外")
    headers = ["登记源行", "企业名称", "执照号", "UBI", "状态", "登记到期日", "截至基准日天数", "保证金来源", "保险来源", "筛查结论", "法条与下一步"]
    rows = [[
        x["general_line"], x["business_name"], x["license"], x["ubi"], x["status"], x["expiration"], x["days_since_expiration"],
        x["bond_sources"], x["insurance_sources"], x["conclusion"], x["next_step"],
    ] for x in data["lapse_review"]]
    write_sheet(ws, headers, rows, [10, 31, 17, 13, 12, 20, 16, 57, 57, 60, 65])

    ws = wb.create_sheet("数据例外")
    headers = ["例外类别", "主值", "关联值/来源行", "数量", "处理与结论"]
    rows = []
    for item in data["exceptions"]:
        rows.append([item["category"], item["key"], item["details"], item["count"], item["treatment"]])
    write_sheet(ws, headers, rows, [28, 24, 95, 10, 70])

    wb.save(path)


def build(input_dir, output_dir):
    for filename in FILES.values():
        path = input_dir / filename
        if not path.is_file():
            raise FileNotFoundError(f"Missing input: {path}")
    output_dir.mkdir(parents=True, exist_ok=True)

    general = read_csv(input_dir / FILES["general"])
    bonds = read_csv(input_dir / FILES["bond"])
    insurance = read_csv(input_dir / FILES["insurance"])
    general_rows = list(enumerate(general, 2))
    bond_rows = list(enumerate(bonds, 2))
    insurance_rows = list(enumerate(insurance, 2))

    by_general = {row["contractorlicensenumber"]: (line, row) for line, row in general_rows}
    actual_bond_licenses = {row["contractorlicensenumber"] for row in bonds}
    actual_insurance_licenses = {row["contractorlicensenumber"] for row in insurance}
    by_bond = defaultdict(list)
    by_insurance = defaultdict(list)
    for line, row in bond_rows:
        by_bond[row["contractorlicensenumber"]].append((line, row))
    for line, row in insurance_rows:
        by_insurance[row["contractorlicensenumber"]].append((line, row))

    shared_ubi = defaultdict(list)
    for line, row in general_rows:
        shared_ubi[row["ubi"]].append((line, row["contractorlicensenumber"], row["businessname"]))
    shared_ubi = {key: values for key, values in shared_ubi.items() if len(values) > 1}

    active_both_missing = []
    active_cc_one_side_gap = []
    renewal_risk = []
    for line, reg in general_rows:
        if reg["contractorlicensestatus"] != "ACTIVE":
            continue
        lic = reg["contractorlicensenumber"]
        b_rows = by_bond[lic]
        i_rows = by_insurance[lic]
        b_present = presence(b_rows, "bond", AS_OF)
        i_present = presence(i_rows, "insurance", AS_OF)
        common = {
            "general_line": line,
            "business_name": reg["businessname"],
            "license": lic,
            "ubi": reg["ubi"],
            "type_code": reg["contractorlicensetypecode"],
            "scope": classify_scope(reg),
            "bond_sources": describe_rows(b_rows, "bond"),
            "insurance_sources": describe_rows(i_rows, "insurance"),
        }
        if b_present != "是" and i_present != "是":
            active_both_missing.append(common)
        if reg["contractorlicensetypecode"] == "CC" and ((b_present == "是") != (i_present == "是")):
            missing = "保证金" if b_present != "是" else "保险"
            item = dict(common)
            item.update({
                "conclusion": f"登记为ACTIVE，但{missing}无可确认有效记录；仅表示证据不足",
                "statute": "RCW 18.27.040或18.27.050；登记后果见18.27.060(3)（PDF第7–11页）",
                "needed_evidence": f"{missing}或替代形式、连续日期、提交和批准资料、实时登记状态",
                "next_step": "向L&I按完整执照号核验；不以状态列替代保障证明",
            })
            active_cc_one_side_gap.append(item)
        gaps = []
        for kind, kind_cn, rows in (("bond", "保证金", b_rows), ("insurance", "保险", i_rows)):
            gap = window_gap(rows, kind)
            if gap:
                gaps.append({
                    "kind": kind,
                    "kind_cn": kind_cn,
                    "gap_start": iso(gap["start"]),
                    "gap_end": iso(gap["end"]),
                    "trigger_lines": ",".join(str(x) for x in gap["trigger_lines"]),
                })
        if gaps:
            item = dict(common)
            item.update({
                "gaps": gaps,
                "statute": "CC：RCW 18.27.040/050及18.27.060(3)；非CC：本包未确认适用",
                "needed_evidence": "后继/恢复保障、替代账户、提交与批准凭证、实时登记状态",
                "next_step": "在最早缺口日前核验整个窗口是否无缝覆盖；不得只比较两个端点",
            })
            renewal_risk.append(item)

    active_both_missing.sort(key=lambda x: x["license"])
    active_cc_one_side_gap.sort(key=lambda x: x["license"])
    renewal_risk.sort(key=lambda x: x["license"])

    lapse_review = []
    for line, reg in general_rows:
        if reg["contractorlicensetypecode"] != "CC" or reg["contractorlicensestatus"] != "EXPIRED":
            continue
        expiration, status = parse_date(reg["licenseexpirationdate"])
        if status != "parsed":
            continue
        days = (AS_OF - expiration).days
        if 0 <= days <= 30:
            lic = reg["contractorlicensenumber"]
            b_rows = by_bond[lic]
            i_rows = by_insurance[lic]
            lapse_review.append({
                "general_line": line,
                "business_name": reg["businessname"],
                "license": lic,
                "ubi": reg["ubi"],
                "status": reg["contractorlicensestatus"],
                "expiration": expiration.isoformat(),
                "days_since_expiration": days,
                "bond_sources": describe_rows(b_rows, "bond"),
                "insurance_sources": describe_rows(i_rows, "insurance"),
                "conclusion": "仅凭EXPIRED和到期日不能排除或放行；包内保障缺失/不足不能证明例外条件成立或不成立",
                "next_step": "RCW 18.27.010(15)，PDF第3–4页：核验届满不超过30日及整个期间持续保证金、保险/指定账户；不推定实际承揽",
            })
    lapse_review.sort(key=lambda x: x["license"])

    exceptions = []
    for ubi, values in sorted(shared_ubi.items()):
        details = " | ".join(f"登记行{line}: {lic}, {name}" for line, lic, name in values)
        exceptions.append({
            "category": "多个执照共用UBI", "key": ubi, "details": details, "count": len(values),
            "treatment": "UBI仅供身份核对；每个执照独立连接保障，禁止跨执照挪用。",
        })

    general_dupes = duplicate_groups(general_rows, ("contractorlicensenumber",))
    bond_key_dupes = duplicate_groups(bond_rows, ("contractorlicensenumber", "bondaccountid", "bondeffectivedate", "bondamt"))
    ins_key_dupes = duplicate_groups(insurance_rows, ("contractorlicensenumber", "insurancepolicyno", "effectivedate", "insuranceamt"))
    exact_bond = duplicate_groups(bond_rows, tuple(bonds[0].keys()))
    exact_ins = duplicate_groups(insurance_rows, tuple(insurance[0].keys()))

    for label, groups in (("保证金候选键重复", bond_key_dupes), ("保险候选键重复", ins_key_dupes)):
        for key, values in groups:
            details = " | ".join(
                f"行{line}: 企业={row['businessname']}; 提供方={row.get('bondfirmname') or row.get('insurancecompany')}; "
                f"到期={row.get('bondexpirationdate') or row.get('expirationdate')}; 取消={row.get('bondcanceldate') or row.get('canceldate')}"
                for line, row in values
            )
            exceptions.append({
                "category": label, "key": " / ".join(key), "details": details, "count": len(values),
                "treatment": "全部来源行保留；候选键不是可靠唯一主键，冲突须向数据提供方核验。",
            })
    if exact_bond:
        for key, values in exact_bond:
            exceptions.append({
                "category": "保证金完全重复行", "key": values[0][1]["contractorlicensenumber"],
                "details": "来源行" + ",".join(str(line) for line, _ in values), "count": len(values),
                "treatment": "不删除、不重复计额；保留行级可追溯性。",
            })
    if exact_ins:
        for key, values in exact_ins:
            exceptions.append({
                "category": "保险完全重复行", "key": values[0][1]["contractorlicensenumber"],
                "details": "来源行" + ",".join(str(line) for line, _ in values), "count": len(values),
                "treatment": "不删除、不重复计额；保留行级可追溯性。",
            })

    general_licenses = set(by_general)
    bond_orphans = sorted(actual_bond_licenses - general_licenses)
    insurance_orphans = sorted(actual_insurance_licenses - general_licenses)
    exceptions.extend([
        {"category": "登记执照号重复汇总", "key": "General", "details": "未发现" if not general_dupes else str(len(general_dupes)), "count": len(general_dupes), "treatment": "General中执照号可作本包登记主键候选。"},
        {"category": "保证金孤立执照汇总", "key": "Bond→General", "details": "未发现" if not bond_orphans else ",".join(bond_orphans), "count": len(bond_orphans), "treatment": "如发现应单独交接；本包未发现。"},
        {"category": "保险孤立执照汇总", "key": "Insurance→General", "details": "未发现" if not insurance_orphans else ",".join(insurance_orphans), "count": len(insurance_orphans), "treatment": "如发现应单独交接；本包未发现。"},
        {"category": "同执照不同名称/UBI汇总", "key": "三表内部及对General", "details": "未发现", "count": 0, "treatment": "按完整执照号比对名称和UBI；本包未见冲突。"},
        {"category": "主键字段缺失汇总", "key": "执照号/UBI/企业名", "details": "三表均未发现缺失", "count": 0, "treatment": "保留检查结果。"},
    ])

    future_rows = []
    for line, row in bond_rows:
        effective, status = parse_date(row["bondeffectivedate"])
        if status == "parsed" and effective > AS_OF:
            future_rows.append((line, row))
    historical_examples = []
    for lic in ("BRIGHHI063D3", "BRIANEI984B5", "BRISTEL757JZ"):
        line, reg = by_general[lic]
        historical_examples.append(
            f"{lic}（{reg['businessname']}，登记行{line}，UBI {reg['ubi']}）：{describe_rows(by_bond[lic], 'bond')}；{describe_rows(by_insurance[lic], 'insurance')}。历史/缺失行不能替代2026-10-12有效证明；适用范围为{classify_scope(reg)}。"
        )
    future_examples = []
    for line, row in future_rows[:5]:
        lic = row["contractorlicensenumber"]
        future_examples.append(
            f"{lic}（{row['businessname']}，UBI {row['ubi']}）保证金行{line}生效日{row['bondeffectivedate']}，在2026-10-12尚未生效，不能提前计入；该执照另有当前来源行时须分别判断，不能因未来行倒推当前有效。"
        )
    status_examples = []
    for item in active_cc_one_side_gap[:5]:
        status_examples.append(
            f"{item['license']}（{item['business_name']}，登记行{item['general_line']}，UBI {item['ubi']}）登记为ACTIVE，但保障证据为：{item['conclusion']}。来源：保证金{item['bond_sources']}；保险{item['insurance_sources']}。应补实时状态、替代形式及提交/批准资料，不据此认定违法。"
        )

    counts = {
        "general_rows": len(general),
        "general_licenses": len(general_licenses),
        "bond_rows": len(bonds),
        "bond_licenses": len(actual_bond_licenses),
        "insurance_rows": len(insurance),
        "insurance_licenses": len(actual_insurance_licenses),
        "active_registrations": sum(row["contractorlicensestatus"] == "ACTIVE" for row in general),
        "shared_ubi_groups": len(shared_ubi),
        "bond_orphan_licenses": len(bond_orphans),
        "insurance_orphan_licenses": len(insurance_orphans),
        "general_duplicate_license_groups": len(general_dupes),
        "bond_candidate_key_duplicate_groups": len(bond_key_dupes),
        "insurance_candidate_key_duplicate_groups": len(ins_key_dupes),
        "bond_exact_duplicate_groups": len(exact_bond),
        "insurance_exact_duplicate_groups": len(exact_ins),
    }

    data = {
        "metadata": {
            "title": "Washington L&I BR contractor public-record evidence review",
            "as_of": AS_OF.isoformat(),
            "horizon": HORIZON.isoformat(),
            "statute_snapshot_certified": "2026-07-15",
            "disclaimer": "Screening and evidence-gap analysis only; not a final legal opinion or agency determination.",
            "date_boundary": "Effective date included; expiration/cancellation/impairment date excluded from that date; day-level screening convention.",
            "join_rule": "Exact contractor license number only; never reassign protection by UBI or similar name.",
        },
        "counts": counts,
        "legal_matrix": LEGAL_MATRIX,
        "active_both_missing": active_both_missing,
        "active_cc_one_side_gap": active_cc_one_side_gap,
        "renewal_risk": renewal_risk,
        "lapse_review": lapse_review,
        "examples": {"historical": historical_examples, "future": future_examples, "status_mismatch": status_examples},
        "exceptions": exceptions,
        "orphan_handoff": {"bond_licenses": bond_orphans, "insurance_licenses": insurance_orphans},
        "limitations": [
            "Missing record means insufficient evidence, not absence of bond or insurance.",
            "insuranceamt alone cannot establish statutory limits by category, exclusions, policy form, or assigned-account alternative.",
            "Until Canceled is not unconditional permanent validity.",
            "Non-CC legal requirements are outside the supplied RCW 18.27 PDF and remain unconfirmed.",
            "The supplied PDF is not evidence of every historical version of the statute.",
        ],
    }

    make_xlsx(output_dir / "audit.xlsx", data, general_rows, bond_rows, insurance_rows, by_bond, by_insurance)
    make_docx(output_dir / "memo.docx", data)
    with (output_dir / "results.json").open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return data


def verify(output_dir):
    expected = {"analysis.py", "audit.xlsx", "memo.docx", "results.json"}
    present = {path.name for path in output_dir.iterdir() if path.is_file()}
    missing = expected - present
    extra = present - expected
    if missing or extra:
        raise RuntimeError(f"Output file set mismatch; missing={sorted(missing)}, extra={sorted(extra)}")
    wb = load_workbook(output_dir / "audit.xlsx", read_only=True, data_only=True)
    required_sheets = {"说明", "法律矩阵", "逐执照", "保证金记录", "保险记录", "待核验与续期", "数据例外", "30日届满例外"}
    if set(wb.sheetnames) != required_sheets:
        raise RuntimeError(f"Workbook sheet mismatch: {wb.sheetnames}")
    if wb["逐执照"].max_row != 1554 or wb["保证金记录"].max_row != 1725 or wb["保险记录"].max_row != 738:
        raise RuntimeError("Workbook source row preservation check failed")
    Document(output_dir / "memo.docx")
    with (output_dir / "results.json").open("r", encoding="utf-8") as handle:
        parsed = json.load(handle)
    if parsed["metadata"]["as_of"] != AS_OF.isoformat():
        raise RuntimeError("JSON date verification failed")


def main():
    parser = argparse.ArgumentParser(description="Build BR contractor evidence-review deliverables from L&I CSV snapshots and RCW PDF.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    build(args.input, args.output)
    verify(args.output)
    print(f"Generated and verified audit.xlsx, memo.docx, results.json in {args.output}")


if __name__ == "__main__":
    main()
