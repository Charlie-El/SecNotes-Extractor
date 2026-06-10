from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent if SCRIPT_DIR.name == "scripts" else SCRIPT_DIR
SOURCE_FILE = ROOT / "outputs" / "notes_coverage_extract.xlsx"
TARGET_FILE = ROOT / "notes_coverage_extract_pdf.xlsx"


@dataclass(frozen=True)
class FieldSpec:
    pdf_name: str
    source_name: str | None
    description: str
    status: str
    note: str = ""


TABLE1_FIELDS = [
    FieldSpec("CUSIP", "CUSIP", "证券代码（美股主键）", "已保留"),
    FieldSpec("stock_code", "stock_code", "证券代码（主键）", "已保留"),
    FieldSpec("company_name", "company_name", "公司名称", "已保留"),
    FieldSpec("report_period", "report_period", "报告期（如 2024-12-31）", "已保留"),
    FieldSpec("report_type", "report_type", "年报/半年报/季报", "已保留"),
    FieldSpec(
        "industry_code_l1",
        "industry_l1",
        "一级行业",
        "待确认",
        "原表字段名为 industry_l1，且当前值看起来是行业名称，不一定是代码。",
    ),
    FieldSpec(
        "industry_code_l2",
        "industry_l2",
        "二级行业",
        "待确认",
        "原表字段名为 industry_l2，且当前值看起来是行业名称，不一定是代码。",
    ),
    FieldSpec("main_business_desc", "main_business_desc", "主营业务描述原文（去噪后）", "已保留"),
    FieldSpec(
        "main_business_desc_summary",
        None,
        "标准化摘要（50-200 字）",
        "原表缺失",
        "原表没有可直接对应的摘要字段，本次不做自动生成。",
    ),
    FieldSpec("total_revenue", "total_revenue", "营业总收入（元）", "已保留"),
    FieldSpec("segment_disclosure_flag", "segment_disclosure_flag", "是否有业务分部披露", "已保留"),
    FieldSpec("segment_count", "segment_count", "业务分部数量", "已保留"),
    FieldSpec(
        "primary_disclosure_dim",
        "primary_disclosure_dim",
        "主要披露维度：segment/product/industry/narrative",
        "已保留",
    ),
    FieldSpec(
        "disclosure_dims_available",
        "disclosure_dims_available",
        '所有可用维度列表，如 ["product", "region"]',
        "已保留",
    ),
    FieldSpec("has_revenue_breakdown", "has_revenue_breakdown", "是否有任何形式的营收分解", "已保留"),
    FieldSpec("has_margin_breakdown", "has_margin_breakdown", "是否有毛利/毛利率分解", "已保留"),
]

TABLE2_FIELDS = [
    FieldSpec("segment_id", "segment_id", "主键（自增）", "已保留", "原表为组合标识，不是数据库自增主键。"),
    FieldSpec("stock_code", "stock_code", "外键", "已保留"),
    FieldSpec("report_period", "report_period", "外键", "已保留"),
    FieldSpec("segment_name", "segment_name", "分部名称（如“航天业务”）", "已保留"),
    FieldSpec("segment_type", "segment_type", "按业务/按产品/按地区", "已保留"),
    FieldSpec("segment_member_raw", "segment_member_raw", "原始 XBRL 维度成员名", "已保留"),
    FieldSpec(
        "segment_member_country_cn",
        "segment_member_country_cn",
        "若原始成员是完整英文国家名，则翻译为中文；缩写、简称、区域词不翻译",
        "已保留",
    ),
    FieldSpec("segment_revenue", "segment_revenue", "分部营业收入", "已保留"),
    FieldSpec("segment_revenue_ratio", "segment_revenue_ratio", "占总营收比例", "已保留"),
    FieldSpec("segment_gross_profit", "segment_gross_profit", "分部毛利", "已保留"),
    FieldSpec("segment_gross_margin", "segment_gross_margin", "毛利率", "已保留"),
    FieldSpec("segment_desc", "segment_desc", "分部业务描述", "已保留"),
    FieldSpec(
        "product_keywords",
        None,
        "提取的关键产品/服务词",
        "原表缺失",
        "原表没有可直接对应的关键词字段。",
    ),
    FieldSpec("source_section", "source_section", "business/MDA/notes", "已保留"),
    FieldSpec(
        "mention_sections",
        None,
        '出现的章节列表，如 ["business", "MDA"]',
        "原表缺失",
        "原表没有按章节数组单独存储。",
    ),
    FieldSpec(
        "is_in_company_intro",
        None,
        "是否出现在公司简介/主营业务定义句",
        "原表缺失",
        "原表没有可直接对应的布尔判断字段。",
    ),
    FieldSpec(
        "is_strategic_focus",
        None,
        "是否被列为战略发展方向",
        "原表缺失",
        "原表没有可直接对应的布尔判断字段。",
    ),
]


HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9EAF7")
DESC_FILL = PatternFill(fill_type="solid", fgColor="FFF2CC")
NOTE_FILL = PatternFill(fill_type="solid", fgColor="EDEDED")
MISSING_FILL = PatternFill(fill_type="solid", fgColor="FCE4D6")
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
FONT_NORMAL = Font(name="Arial", size=10)
FONT_BOLD = Font(name="Arial", size=10, bold=True)
FONT_RED_BOLD = Font(name="Arial", size=10, bold=True, color="FF0000")


def load_source_rows() -> tuple[list[str], list[list[object]]]:
    wb = load_workbook(SOURCE_FILE, read_only=True, data_only=False)
    ws = wb[wb.sheetnames[0]]
    headers = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    rows: list[list[object]] = []
    for row in ws.iter_rows(min_row=4, values_only=True):
        rows.append(list(row))
    return headers, rows


def autosize(ws) -> None:
    for col_idx, column in enumerate(ws.columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = 0
        for cell in column[:200]:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 40)


def write_table_sheet(ws, title_note: str, specs: list[FieldSpec], source_index: dict[str, int], rows: list[list[object]]) -> None:
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(specs))
    note_cell = ws.cell(1, 1, title_note)
    note_cell.fill = NOTE_FILL
    note_cell.font = FONT_BOLD
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for col_idx, spec in enumerate(specs, start=1):
        desc_cell = ws.cell(2, col_idx, spec.description)
        desc_cell.fill = DESC_FILL if spec.status == "已保留" else MISSING_FILL
        desc_cell.font = FONT_NORMAL
        desc_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        desc_cell.border = THIN_BORDER

        header_cell = ws.cell(3, col_idx, spec.pdf_name)
        header_cell.fill = HEADER_FILL
        header_cell.font = FONT_RED_BOLD if spec.status != "已保留" else FONT_BOLD
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_cell.border = THIN_BORDER

    for row_idx, source_row in enumerate(rows, start=4):
        for col_idx, spec in enumerate(specs, start=1):
            value = None
            if spec.source_name and spec.source_name in source_index:
                value = source_row[source_index[spec.source_name]]
            cell = ws.cell(row_idx, col_idx, value)
            cell.font = FONT_NORMAL
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 34
    ws.row_dimensions[3].height = 22
    ws.auto_filter.ref = f"A3:{ws.cell(ws.max_row, len(specs)).coordinate}"
    autosize(ws)


def write_mapping_sheet(ws, table1_count: int, table2_count: int) -> None:
    headers = ["范围", "PDF 字段名", "处理结果", "原表字段名", "说明"]
    top_note = (
        "本次仅整理《【数据】主题指数数据.pdf》2.1 基础数据表（表1、表2）。"
        "PDF 2.2 主题筛选用表（表3、表4）在当前原表中无直接对应数据，未纳入本次整理。"
        f" 表1输出 {table1_count} 行，表2输出 {table2_count} 行。"
    )
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    cell = ws.cell(1, 1, top_note)
    cell.fill = NOTE_FILL
    cell.font = FONT_BOLD
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for col_idx, header in enumerate(headers, start=1):
        header_cell = ws.cell(2, col_idx, header)
        header_cell.fill = HEADER_FILL
        header_cell.font = FONT_BOLD
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        header_cell.border = THIN_BORDER

    row_idx = 3
    for scope, specs in (
        ("表1_公司主营业务原文", TABLE1_FIELDS),
        ("表2_公司业务分部明细", TABLE2_FIELDS),
    ):
        for spec in specs:
            ws.cell(row_idx, 1, scope)
            field_cell = ws.cell(row_idx, 2, spec.pdf_name)
            status_cell = ws.cell(row_idx, 3, spec.status)
            ws.cell(row_idx, 4, spec.source_name or "")
            ws.cell(row_idx, 5, spec.note)
            if spec.status != "已保留":
                field_cell.font = FONT_RED_BOLD
                status_cell.font = FONT_RED_BOLD
                field_cell.fill = MISSING_FILL
                status_cell.fill = MISSING_FILL
            row_idx += 1

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            if cell.row == 2:
                cell.font = FONT_BOLD
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            elif cell.column not in (2, 3) or cell.row < 3 or ws.cell(cell.row, 3).value == "已保留":
                cell.font = FONT_NORMAL
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            else:
                cell.font = FONT_RED_BOLD
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = THIN_BORDER

    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 24
    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:E{ws.max_row}"
    autosize(ws)


def main() -> None:
    headers, data_rows = load_source_rows()
    source_index = {name: idx for idx, name in enumerate(headers)}

    row_type_idx = source_index["row_type"]
    company_rows = [row for row in data_rows if row[row_type_idx] == "company"]
    segment_rows = [row for row in data_rows if row[row_type_idx] == "segment"]

    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    ws1 = wb.create_sheet("表1_公司主营业务原文")
    ws2 = wb.create_sheet("表2_公司业务分部明细")
    ws3 = wb.create_sheet("字段映射说明")

    write_table_sheet(
        ws1,
        "按 PDF 表1 整理，仅保留 company 行；红色字段表示原表无直接对应字段或字段含义待确认。",
        TABLE1_FIELDS,
        source_index,
        company_rows,
    )
    write_table_sheet(
        ws2,
        "按 PDF 表2 整理，仅保留 segment 行；红色字段表示原表无直接对应字段或字段含义待确认。",
        TABLE2_FIELDS,
        source_index,
        segment_rows,
    )
    write_mapping_sheet(ws3, len(company_rows), len(segment_rows))

    wb.save(TARGET_FILE)


if __name__ == "__main__":
    main()
