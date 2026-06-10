from __future__ import annotations

import argparse
import json
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from babel import Locale
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from lxml import etree


SCRIPT_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPT_DIR.parent if SCRIPT_DIR.name == "scripts" else SCRIPT_DIR
INPUTS_DIR = ROOT_DIR / "inputs"
DATA_DIR = INPUTS_DIR / "data"
WORKBOOK_PATH = INPUTS_DIR / "美股数据提取筛选.xlsx"
if not WORKBOOK_PATH.exists():
    WORKBOOK_PATH = INPUTS_DIR / "workbook.xlsx"
EXTRA_STOCKBOOK_PATH = INPUTS_DIR / "补充美股（HALO）和港股清单.xlsx"
OUTPUT_DIR = ROOT_DIR / "outputs"
OUTPUT_XLSX = OUTPUT_DIR / "notes_coverage_extract.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "notes_coverage_extract.csv"
OUTPUT_SUMMARY = OUTPUT_DIR / "notes_coverage_summary.json"

NOTES_DIRS = sorted([p for p in DATA_DIR.iterdir() if p.is_dir() and p.name.endswith("_notes")]) if DATA_DIR.exists() else []

ANNUAL_FORMS = {"10-K", "20-F", "40-F"}
SEC_TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
USER_AGENT = "Codex notes coverage extract contact@openai.com"

REVENUE_TAG_PRIORITY = [
    "RevenueFromExternalCustomers",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "RevenueFromRenderingOfTelecommunicationServices",
    "RevenueFromContractsWithCustomers",
    "Revenues",
    "RevenuesAndOther",
    "RevenuesAndRealizedGainsLossesOnDerivativeInstruments",
    "RegulatedAndUnregulatedOperatingRevenue",
    "RegulatedOperatingRevenueGas",
    "SalesRevenueNet",
    "Revenue",
    "NetSales",
    "RevenueFromContractWithCustomerIncludingAssessedTax",
    "TotalRevenue",
    "RevenueAndOperatingIncome",
    "RevenueFromSaleOfGoods",
    "RevenueFromRenderingOfServices",
    "RevenueFromSaleOfGold",
]

METRIC_TAGS = [
    "Assets",
    "Liabilities",
    "StockholdersEquity",
    "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
    "CashAndCashEquivalentsAtCarryingValue",
    "GrossProfit",
    "GrossProfitLoss",
    "OperatingIncomeLoss",
    "NetIncomeLoss",
    "ResearchAndDevelopmentExpense",
    "DepreciationDepletionAndAmortization",
    "NetCashProvidedByUsedInOperatingActivities",
    "PaymentsToAcquirePropertyPlantAndEquipment",
]

TEXT_TAG_PRIORITIES = {
    "main_business_tag": [
        "BusinessDescriptionPolicyTextBlock",
        "BusinessDescriptionAndAccountingPoliciesTextBlock",
        "BusinessDescriptionAndBasisOfPresentationTextBlock",
        "OrganizationConsolidationBasisOfPresentationBusinessDescriptionAndAccountingPoliciesTextBlock",
        "BusinessOverviewPolicyTextBlock",
        "DescriptionOfBusinessTextBlock",
        "BusinessDescriptionTextBlock",
        "BusinessDescriptionAndOverviewTextBlock",
    ],
    "segment_text_tag": [
        "SegmentReportingDisclosureTextBlock",
        "ScheduleOfSegmentReportingInformationBySegmentTextBlock",
        "ScheduleOfSegmentReportingInformationBySegmentTableTextBlock",
        "DisclosureOfEntitysReportableSegmentsExplanatory",
        "DisclosureOfOperatingSegmentsExplanatory",
    ],
    "revenue_note_tag": [
        "DisaggregationOfRevenueTableTextBlock",
        "DisaggregationOfRevenueTextBlock",
        "DisclosureOfDisaggregationOfRevenueFromContractsWithCustomersExplanatory",
        "RevenueRecognitionPolicyTextBlock",
        "RevenueRecognitionDisclosureTextBlock",
    ],
}

EN_LOCALE = Locale.parse("en")
ZH_LOCALE = Locale.parse("zh")

ROW_COLUMNS = [
    "row_type",
    "stock_code",
    "company_name",
    "CUSIP",
    "ticker",
    "ticker_match_method",
    "cik",
    "company_name_filing",
    "report_type",
    "report_period",
    "filed_date",
    "accepted_datetime",
    "filing_month_batch",
    "adsh",
    "instance",
    "sic",
    "industry_l1",
    "industry_l2",
    "industry_l3",
    "industry_l4",
    "fiscal_year",
    "fiscal_period",
    "fiscal_year_end",
    "wksi",
    "filer_category",
    "sec_file_number",
    "tax_id",
    "local_phone",
    "public_float_usd",
    "float_date",
    "country_business_address",
    "state_business_address",
    "city_business_address",
    "country_incorporation",
    "state_incorporation",
    "main_business_desc",
    "main_business_tag",
    "segment_text_tag",
    "revenue_note_tag",
    "total_revenue",
    "total_revenue_tag",
    "total_revenue_version",
    "total_revenue_uom",
    "gross_profit",
    "gross_profit_tag",
    "operating_income",
    "operating_income_tag",
    "net_income",
    "net_income_tag",
    "assets",
    "assets_uom",
    "liabilities",
    "liabilities_uom",
    "equity",
    "equity_uom",
    "cash_and_equivalents",
    "cash_uom",
    "operating_cash_flow",
    "operating_cash_flow_uom",
    "capex",
    "capex_uom",
    "rd_expense",
    "rd_expense_uom",
    "depreciation_amortization",
    "depreciation_amortization_uom",
    "eps_basic",
    "eps_diluted",
    "shares_outstanding",
    "shares_outstanding_uom",
    "shares_issued",
    "shares_issued_uom",
    "shares_authorized",
    "shares_authorized_uom",
    "weighted_avg_diluted_shares",
    "weighted_avg_diluted_shares_uom",
    "segment_disclosure_flag",
    "segment_count",
    "primary_disclosure_dim",
    "disclosure_dims_available",
    "has_revenue_breakdown",
    "has_margin_breakdown",
    "segment_id",
    "segment_name",
    "segment_type",
    "segment_axis",
    "segment_member_raw",
    "segment_member_country_cn",
    "segment_revenue",
    "segment_revenue_uom",
    "segment_revenue_ratio",
    "segment_gross_profit",
    "segment_gross_profit_uom",
    "segment_gross_margin",
    "segment_desc",
    "source_section",
    "source_tag",
    "source_version",
    "source_group_tag",
    "source_context",
    "source_dim_hash",
    "source_segments_text",
    "source_report",
    "source_report_shortname",
    "source_report_longname",
    "evidence_text",
    "extraction_status",
    "notes_file_batch",
]

FIELD_DESCRIPTIONS = {}
FIELD_DESCRIPTIONS.update(
    {
        "row_type": "行类型；company 表示公司层汇总，segment 表示业务分部/产品/地区等明细行。",
        "stock_code": "股票代码，直接沿用股票池中的 SECU_CODE。",
        "company_name": "公司中文名称。",
        "CUSIP": "证券 CUSIP 标识。",
        "ticker": "用于匹配 filing 的标准化 ticker。",
        "ticker_match_method": "股票池与 filing 的匹配方式。",
        "cik": "SEC 公司 CIK 编号。",
        "company_name_filing": "年报申报中的英文公司名。",
        "report_type": "申报类型，如 10-K、20-F、40-F。",
        "report_period": "报告期截止日。",
        "filed_date": "年报提交日期。",
        "accepted_datetime": "SEC 接收时间。",
        "filing_month_batch": "notes 数据所在月度批次。",
        "adsh": "SEC filing 唯一编号 ADSH。",
        "instance": "申报实例文件名。",
        "sic": "SEC SIC 行业代码。",
        "industry_l1": "股票池行业字段第 1 层。",
        "industry_l2": "股票池行业字段第 2 层。",
        "industry_l3": "股票池行业字段第 3 层。",
        "industry_l4": "股票池行业字段第 4 层。",
        "main_business_desc": "主营业务或相关附注原文。",
        "main_business_tag": "主营业务原文对应的 XBRL 标签。",
        "segment_text_tag": "分部披露原文对应的 XBRL 标签。",
        "revenue_note_tag": "收入拆分/收入确认原文对应的 XBRL 标签。",
        "total_revenue": "公司总营业收入。",
        "segment_disclosure_flag": "是否存在分部或收入拆分披露。",
        "segment_count": "主口径下识别出的分部数量。",
        "primary_disclosure_dim": "主要分部披露维度类型。",
        "disclosure_dims_available": "该 filing 中识别到的全部维度类型列表。",
        "has_revenue_breakdown": "是否存在带维度的收入拆分事实。",
        "has_margin_breakdown": "是否存在分部毛利或毛利率。",
        "segment_id": "分部明细行唯一编号。",
        "segment_name": "将原始成员名拆词后的可读名称。",
        "segment_type": "分部类型，如 business_segment、product、region、customer_type、end_market 等。",
        "segment_axis": "原始 XBRL 维度轴名称。",
        "segment_member_raw": "原始 XBRL 维度成员名，未改写。",
        "segment_member_country_cn": "若 segment_member_raw 对应国家名、国家代码或常见地区/区域词，则给出中文名。",
        "segment_revenue": "该分部/维度成员对应收入。",
        "segment_revenue_uom": "分部收入计量单位。",
        "segment_revenue_ratio": "分部收入占公司总收入比重。",
        "segment_gross_profit": "分部毛利。",
        "segment_gross_profit_uom": "分部毛利单位。",
        "segment_gross_margin": "分部毛利率。",
        "segment_desc": "与该分部最相关的原文片段。",
        "source_section": "来源层级，当前主要为 notes 或 company。",
        "source_tag": "来源数值事实标签。",
        "source_version": "来源 taxonomy 版本。",
        "source_group_tag": "来源口径组标签。",
        "source_dim_hash": "来源维度哈希，可回连 dim 表。",
        "source_segments_text": "来源维度组合原文。",
        "source_report": "来源 report 编号。",
        "source_report_shortname": "来源报表简称。",
        "source_report_longname": "来源报表全名。",
        "evidence_text": "输出时保留的证据原文。",
        "extraction_status": "提取状态，ok 为较完整，partial 为部分字段缺失。",
        "notes_file_batch": "结果来自哪个 notes 批次目录。",
    }
)

TABLE_NOTE = (
    "本表保留股票池中已在 SEC notes 月度批次中覆盖的公司，并按同一公司只保留最新年报/年度申报；"
    "company 行是公司层汇总信息，segment 行是分部/产品/地区/客户等拆分明细。"
)

NOTES_CACHE: dict[str, dict[str, pd.DataFrame]] = {}

AM_ANNUAL_DIR = INPUTS_DIR / "Am_annual"
AM_SUPPLEMENT_TARGETS = {
    "NGD.A": "New Gold Inc",
    "AUGO.O": "Aura Minerals Inc",
    "QMMM.O": "QMMM Holdings Ltd-A",
}

COUNTRY_NAME_WRAPPER_PREFIXES = [
    "country of ",
    "country ",
    "segment ",
]

COUNTRY_NAME_WRAPPER_SUFFIXES = [
    " country",
    " region",
]

COUNTRY_NAME_ALIASES = {
    "bahamas": "BS",
    "bolivia": "BO",
    "brunei": "BN",
    "cape verde": "CV",
    "china mainland": "CN",
    "cote d ivoire": "CI",
    "czech republic": "CZ",
    "democratic republic of the congo": "CD",
    "gambia": "GM",
    "iran": "IR",
    "ivory coast": "CI",
    "laos": "LA",
    "macedonia": "MK",
    "moldova": "MD",
    "north korea": "KP",
    "palestine": "PS",
    "republic of the congo": "CG",
    "russia": "RU",
    "russian federation": "RU",
    "south korea": "KR",
    "syria": "SY",
    "tanzania": "TZ",
    "great britain": "GB",
    "united states": "US",
    "united states of america": "US",
    "venezuela": "VE",
    "vietnam": "VN",
}

COUNTRY_SKIP_VALUES = {
    "country of domicile",
    "domestic country",
    "foreign country",
    "other country",
    "others country",
}

REGION_NAME_MAP = {
    "americas": "美洲",
    "other americas": "其他美洲",
    "total americas": "美洲合计",
    "other america": "其他美洲",
    "other american countries": "其他美洲国家",
    "latin america": "拉丁美洲",
    "other latin america": "其他拉丁美洲",
    "latin america and caribbean": "拉丁美洲和加勒比地区",
    "region of latin america and caribbean": "拉丁美洲和加勒比地区",
    "north america": "北美",
    "other north america": "其他北美地区",
    "south america": "南美",
    "rest of north and south america": "其他南北美洲",
    "united states and canada": "美国和加拿大",
    "united states and territories": "美国及其属地",
    "united kingdom and ireland": "英国和爱尔兰",
    "europe": "欧洲",
    "other europe": "其他欧洲",
    "rest of europe": "其他欧洲",
    "rest of europe and middle east": "其他欧洲和中东地区",
    "western europe": "西欧",
    "europe excluding germany": "德国以外的欧洲",
    "europe except germany": "德国以外的欧洲",
    "europe excluding united kingdom": "英国以外的欧洲",
    "europe and canada": "欧洲和加拿大",
    "europe africa": "欧洲和非洲",
    "other european countries": "其他欧洲国家",
    "other europe countries": "其他欧洲国家",
    "emea": "欧洲、中东和非洲",
    "emea region": "欧洲、中东和非洲",
    "emea excluding uk": "不含英国的欧洲、中东和非洲地区",
    "other emea": "其他欧洲、中东和非洲地区",
    "other europe middle east africa": "其他欧洲、中东和非洲地区",
    "rest of emea": "其他欧洲、中东和非洲地区",
    "restof emea": "其他欧洲、中东和非洲地区",
    "europe middle east and africa": "欧洲、中东和非洲",
    "europe middle east africa": "欧洲、中东和非洲",
    "europe the middle east and africa": "欧洲、中东和非洲",
    "europe middle east and africa excluding germany": "不含德国的欧洲、中东和非洲地区",
    "europe middle east india and africa": "欧洲、中东、印度和非洲",
    "middle east and africa": "中东和非洲",
    "middleeastandafrica": "中东和非洲",
    "middle east and north africa": "中东和北非",
    "middle east africa and oceania": "中东、非洲和大洋洲",
    "middle east": "中东",
    "africa": "非洲",
    "restofafrica": "其他非洲地区",
    "africa and asia": "非洲和亚洲",
    "africa eurasia": "非洲和欧亚地区",
    "asia": "亚洲",
    "other asia": "其他亚洲",
    "rest of asia": "其他亚洲",
    "restofasia": "其他亚洲",
    "other asian countries": "其他亚洲国家",
    "asia excluding china": "不含中国的亚洲地区",
    "asia excluding japan and china": "不含日本和中国的亚洲地区",
    "asia except japan and china": "不含日本和中国的亚洲地区",
    "asia and rest of world": "亚洲及世界其他地区",
    "asia and africa": "亚洲和非洲",
    "asia other": "其他亚洲",
    "asia and asia pacific": "亚洲及亚太地区",
    "asia pacific": "亚太地区",
    "apac": "亚太地区",
    "apj": "亚太及日本",
    "other asia pacific": "其他亚太地区",
    "asia pacific other": "其他亚太地区",
    "rest of apac": "其他亚太地区",
    "rest of asia pacific": "其他亚太地区",
    "asia pacific excluding japan": "不含日本的亚太地区",
    "asia pacific excluding china taiwan and japan": "不含中国、台湾和日本的亚太地区",
    "asia pacific excluding greater china": "不含大中华区的亚太地区",
    "apac excluding greater china": "不含大中华区的亚太地区",
    "south korea taiwan and other asia pacific": "韩国、台湾及其他亚太地区",
    "asia pacific and japan": "亚太及日本",
    "asia pacific and rest of world": "亚太地区及世界其他地区",
    "asia pacific and other": "亚太地区及其他地区",
    "asia pacific including australia and new zealand": "含澳新地区的亚太地区",
    "asia pacific and africa": "亚太和非洲",
    "asia pacific middle east africa": "亚太、中东和非洲",
    "asia pacific middle east and africa": "亚太、中东和非洲",
    "greater china": "大中华区",
    "china including hong kong": "中国（含香港）",
    "china includes hong kong": "中国（含香港）",
    "china and hong kong": "中国（含香港）",
    "international": "国际",
    "other international": "其他国际地区",
    "all other international": "所有其他国际地区",
    "total international": "国际合计",
    "international operations": "国际业务",
    "foreign": "海外",
    "foreign countries": "海外国家",
    "other foreign countries": "其他海外国家",
    "other foreign": "其他海外地区",
    "geographic distribution foreign": "海外地区",
    "geographic distribution domestic": "国内地区",
    "domestic country": "国内",
    "non us": "美国以外",
    "outside the united states": "美国以外",
    "outside north america": "北美以外",
    "non us and non canada": "非美国及加拿大地区",
    "non us or europe": "非美国及欧洲地区",
    "other countries": "其他国家",
    "all other countries": "所有其他国家",
    "other country": "其他国家",
    "other countries not separately reported": "其他未单独披露国家",
    "other country other than us and uk": "除美国和英国外的其他国家",
    "other countries outside of the united states and europe": "美国和欧洲以外的其他国家",
    "all other countries except china united states of america and switzerland": "除中国、美国和瑞士外的其他国家",
    "all other countries other than russia brazil and united states": "除俄罗斯、巴西和美国外的其他国家",
    "foreign country excluding specified country": "指定国家以外的其他国家",
    "country of domicile": "注册地国家",
    "rest of world": "世界其他地区",
    "rest of the world": "世界其他地区",
    "restofworld": "世界其他地区",
    "restoftheworld": "世界其他地区",
    "worldwide": "全球",
    "other region": "其他地区",
    "other regions": "其他地区",
    "other geographical areas": "其他地区",
    "other geographic areas": "其他地区",
    "other geographic area": "其他地区",
    "other geographical regions": "其他地区",
    "other geographic regions": "其他地区",
    "other geographical": "其他地区",
    "other geographic locations": "其他地区",
    "other geographical locations": "其他地区",
    "other foreign geographical locations": "其他海外地区",
    "other foreign locations": "其他海外地区",
    "otherforeignlocations": "其他海外地区",
    "all other geographical areas": "所有其他地区",
    "all other geographic region": "所有其他地区",
    "southeast asia": "东南亚",
    "south east asia": "东南亚",
    "emerging markets": "新兴市场",
    "emerging growth markets": "新兴增长市场",
    "established markets": "成熟市场",
    "australia and new zealand": "澳大利亚和新西兰",
    "australasia and other": "澳大拉西亚及其他地区",
    "u s canada and latin america": "美国、加拿大和拉丁美洲",
    "united states canada and latin america": "美国、加拿大和拉丁美洲",
    "canada and latin america": "加拿大和拉丁美洲",
    "latin america and canada": "拉丁美洲和加拿大",
    "mexico and peru": "墨西哥和秘鲁",
    "mexico and central america": "墨西哥和中美洲",
    "mexico central america and south america": "墨西哥、中美洲和南美洲",
    "canada caribbean and central america": "加拿大、加勒比和中美洲",
    "south and latin america": "南美和拉丁美洲",
    "south central america and the caribbean": "中南美洲和加勒比地区",
    "south and central amercia": "中南美洲",
    "americas excluding the united states": "不含美国的美洲地区",
    "americas excluding united states": "不含美国的美洲地区",
    "americas excluding us": "不含美国的美洲地区",
    "americas except united states and brazil": "除美国和巴西外的美洲地区",
    "americas excluding canada and united states": "不含加拿大和美国的美洲地区",
    "americas other": "其他美洲地区",
    "united states offshore": "美国海上地区",
    "united states onshore": "美国陆上地区",
    "canada offshore": "加拿大海上地区",
    "canada onshore": "加拿大陆上地区",
    "gulf coast": "墨西哥湾沿岸",
    "east": "东部地区",
    "west": "西部地区",
    "northeast region": "东北地区",
    "northwest region": "西北地区",
    "southwest region": "西南地区",
    "southern region": "南部地区",
    "great lakes region": "五大湖地区",
    "pacific region": "太平洋地区",
    "north america mainly united states": "以美国为主的北美地区",
    "eucan": "欧洲和加拿大",
    "all other": "其他全部",
    "other": "其他地区",
}

REGION_NAME_WRAPPER_PREFIXES = [
    "region of ",
    "segment ",
]

REGION_NAME_WRAPPER_SUFFIXES = [
    " country",
    " countries",
    " region",
    " regions",
]


def normalize_country_phrase(value: str) -> str:
    text = clean_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def build_country_name_map() -> dict[str, str]:
    mapping: dict[str, str] = {}
    for code, en_name in EN_LOCALE.territories.items():
        if not (len(code) == 2 and code.isalpha()):
            continue
        zh_name = ZH_LOCALE.territories.get(code)
        if not en_name or not zh_name:
            continue
        mapping[normalize_country_phrase(str(en_name))] = str(zh_name)
    for alias, code in COUNTRY_NAME_ALIASES.items():
        zh_name = ZH_LOCALE.territories.get(code)
        if zh_name:
            mapping[normalize_country_phrase(alias)] = str(zh_name)
    return mapping


def build_country_code_map() -> dict[str, str]:
    mapping: dict[str, str] = {}
    for code, zh_name in ZH_LOCALE.territories.items():
        if not (len(code) == 2 and code.isalpha()):
            continue
        if zh_name:
            mapping[code.upper()] = str(zh_name)
    return mapping

@dataclass
class ContextRow:
    adsh: str
    tag: str
    version: str | None
    ddate: str | None
    qtrs: int | None
    uom: str | None
    dimh: str
    value: float | None


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


COUNTRY_NAME_MAP = build_country_name_map()
COUNTRY_CODE_MAP = build_country_code_map()


def normalize_pool_ticker(code: Any) -> str:
    text = clean_text(code).upper().replace("_", "-")
    if "." in text:
        base, suffix = text.rsplit(".", 1)
        if len(suffix) <= 2 and suffix.isalpha():
            text = base
    return text


def normalize_report_date(value: Any) -> str | None:
    text = clean_text(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def camel_to_words(value: str) -> str:
    value = re.sub(r"(Member|Axis|Domain)$", "", value)
    value = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", value)
    value = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", value)
    value = value.replace("_", " ")
    return clean_text(value)


def dimension_category(dim_local: str) -> str | None:
    dim = clean_text(dim_local).lower()
    if any(token in dim for token in ["businesssegments", "operatingsegments", "reportablesegments"]):
        return "business_segment"
    if dim in {"segments", "segment"}:
        return "business_segment"
    if dim.endswith("segments") and not any(token in dim for token in ["portfolio", "loan", "receivable", "share", "contract"]):
        return "business_segment"
    if any(token in dim for token in ["consolidationitems", "reconcilingitems", "reconciliationitems"]):
        return "reconciling_item"
    if any(token in dim for token in ["timingoftransferofgoodorservice", "timingoftransferofgoodsorservices"]):
        return "recognition_timing"
    if any(token in dim for token in ["endmarket", "endmarkets", "marketsofcustomers", "marketofcustomers"]) or dim in {"market", "markets"}:
        return "end_market"
    if any(token in dim for token in ["productorservice", "productsandservices"]):
        return "product"
    if dim in {"product", "products", "service", "services"}:
        return "product"
    if any(token in dim for token in ["saleschannel", "channel"]):
        return "sales_channel"
    if any(token in dim for token in ["geograph", "region", "country", "area"]):
        return "region"
    if "industry" in dim:
        return "industry"
    if any(token in dim for token in ["counterparty", "relatedparty", "legalentity", "entities", "entity", "subsidiar", "associate", "jointventure", "jointventures"]):
        return "counterparty"
    if any(token in dim for token in ["arrangement", "contracttype", "typeofcontract", "typesofcontracts", "typeofcontracts", "contractduration", "pricing"]):
        return "contract_type"
    if any(token in dim for token in ["program", "project"]):
        return "program_or_project"
    if "customer" in dim:
        return "customer_type"
    return None


def translate_country_member_raw(member_raw: Any) -> str | None:
    raw = clean_text(member_raw)
    if not raw:
        return None
    if re.fullmatch(r"[A-Z]{1,4}", raw):
        return None
    words = camel_to_words(raw)
    candidates: list[str] = []
    normalized_words = normalize_country_phrase(words)
    normalized_raw = normalize_country_phrase(raw)
    for candidate in [normalized_words, normalized_raw]:
        if candidate and candidate not in candidates:
            candidates.append(candidate)
        for prefix in COUNTRY_NAME_WRAPPER_PREFIXES:
            if candidate.startswith(prefix):
                stripped = candidate[len(prefix) :].strip()
                if stripped and stripped not in candidates:
                    candidates.append(stripped)
        for suffix in COUNTRY_NAME_WRAPPER_SUFFIXES:
            if candidate.endswith(suffix):
                stripped = candidate[: -len(suffix)].strip()
                if stripped and stripped not in candidates:
                    candidates.append(stripped)
    for candidate in candidates:
        if not candidate or candidate in COUNTRY_SKIP_VALUES:
            continue
        zh_name = COUNTRY_NAME_MAP.get(candidate)
        if zh_name:
            return zh_name
    return None


def translate_region_member_raw(member_raw: Any) -> str | None:
    raw = clean_text(member_raw)
    if not raw:
        return None
    code_match = COUNTRY_CODE_MAP.get(raw.upper())
    if code_match:
        return code_match
    country_match = translate_country_member_raw(raw)
    if country_match:
        return country_match
    candidates: list[str] = []
    normalized_words = normalize_country_phrase(camel_to_words(raw))
    normalized_raw = normalize_country_phrase(raw)
    for candidate in [normalized_words, normalized_raw]:
        if candidate and candidate not in candidates:
            candidates.append(candidate)
        for prefix in REGION_NAME_WRAPPER_PREFIXES:
            if candidate.startswith(prefix):
                stripped = candidate[len(prefix) :].strip()
                if stripped and stripped not in candidates:
                    candidates.append(stripped)
        for suffix in REGION_NAME_WRAPPER_SUFFIXES:
            if candidate.endswith(suffix):
                stripped = candidate[: -len(suffix)].strip()
                if stripped and stripped not in candidates:
                    candidates.append(stripped)
    for candidate in candidates:
        if not candidate:
            continue
        region_match = REGION_NAME_MAP.get(candidate)
        if region_match:
            return region_match
        country_match = COUNTRY_NAME_MAP.get(candidate)
        if country_match:
            return country_match
    return None


def translate_segment_member_cn(member_raw: Any, segment_type: str | None) -> str | None:
    if clean_text(segment_type).lower() == "region":
        return translate_region_member_raw(member_raw)
    return translate_country_member_raw(member_raw)


def load_stock_pool() -> pd.DataFrame:
    pool = pd.read_excel(WORKBOOK_PATH, sheet_name=0).iloc[:557].copy()
    industries = pd.read_excel(WORKBOOK_PATH, sheet_name=1).copy()
    pool["ticker"] = pool["SECU_CODE"].map(normalize_pool_ticker)

    industry_cols = ["CCXDF_CN", "CCXDS_CN", "CCXDT_CN", "CCXDFourth_CN"]
    by_code = industries[["SECU_CODE", *industry_cols]].drop_duplicates("SECU_CODE")
    merged = pool.merge(by_code, on="SECU_CODE", how="left")

    by_cusip = industries[["CUSIP", *industry_cols]].drop_duplicates("CUSIP")
    fallback = pool[["CUSIP"]].merge(by_cusip, on="CUSIP", how="left")
    for col in industry_cols:
        merged[col] = merged[col].fillna(fallback[col])

    if EXTRA_STOCKBOOK_PATH.exists():
        extra = pd.read_excel(EXTRA_STOCKBOOK_PATH, sheet_name=0).copy()
        if len(extra.columns) >= 8:
            extra_pool = extra.iloc[:, [0, 1, 2, 3, 4, 5, 6, 7]].copy()
            extra_pool.columns = [
                "CUSIP",
                "SECU_CODE",
                "SECUNAME",
                "CCXDF_CN",
                "CCXDS_CN",
                "CCXDT_CN",
                "CCXDFourth_CN",
                "鎬诲競鍊硷紙浜跨編鍏冿級",
            ]
            extra_pool["ticker"] = extra_pool["SECU_CODE"].map(normalize_pool_ticker)
            merged = pd.concat([merged, extra_pool], ignore_index=True)
            merged = merged.drop_duplicates(subset=["SECU_CODE"], keep="first").reset_index(drop=True)

    return merged


def load_sec_ticker_map() -> pd.DataFrame:
    response = requests.get(SEC_TICKER_URL, headers={"User-Agent": USER_AGENT}, timeout=60)
    response.raise_for_status()
    data = response.json()
    df = pd.DataFrame.from_dict(data, orient="index")
    df["ticker"] = df["ticker"].astype(str).str.upper().str.strip()
    df["cik10"] = df["cik_str"].astype(str).astype(int).astype(str).str.zfill(10)
    return df[["ticker", "cik10", "title"]].drop_duplicates("ticker")


def load_latest_filings(stock_pool: pd.DataFrame) -> pd.DataFrame:
    sec_map = load_sec_ticker_map()
    pool = stock_pool.merge(sec_map, on="ticker", how="left")
    pool["ticker_match_method"] = pool["cik10"].map(lambda x: "sec_ticker_file" if clean_text(x) else "")

    frames: list[pd.DataFrame] = []
    for notes_dir in NOTES_DIRS:
        sub = pd.read_csv(
            notes_dir / "sub.tsv",
            sep="\t",
            dtype=str,
            usecols=[
                "adsh",
                "cik",
                "name",
                "sic",
                "countryba",
                "stprba",
                "cityba",
                "countryinc",
                "stprinc",
                "wksi",
                "fye",
                "form",
                "period",
                "fy",
                "fp",
                "filed",
                "accepted",
                "detail",
                "instance",
                "pubfloatusd",
                "floatdate",
            ],
        )
        sub = sub[sub["form"].isin(ANNUAL_FORMS) & sub["detail"].fillna("").isin(["1", "true", "TRUE"])].copy()
        sub["batch"] = notes_dir.name
        sub["notes_dir"] = str(notes_dir)
        sub["cik10"] = sub["cik"].astype(str).str.zfill(10)
        sub["filed_num"] = pd.to_numeric(sub["filed"], errors="coerce")
        sub["period_num"] = pd.to_numeric(sub["period"], errors="coerce")

        txt = pd.read_csv(notes_dir / "txt.tsv", sep="\t", dtype=str, usecols=["adsh", "tag", "value"])
        trading = txt[txt["tag"] == "TradingSymbol"][["adsh", "value"]].dropna().drop_duplicates()
        trading["ticker"] = trading["value"].map(normalize_pool_ticker)
        trading = trading[["adsh", "ticker"]].drop_duplicates()
        sub = sub.merge(trading, on="adsh", how="left")
        frames.append(sub)

    all_filings = pd.concat(frames, ignore_index=True)
    pool_tickers = set(pool["ticker"])

    by_ticker = all_filings[all_filings["ticker"].isin(pool_tickers)].copy()
    by_ticker = by_ticker.sort_values(["ticker", "filed_num", "period_num", "adsh"], ascending=[True, False, False, False])
    latest_by_ticker = by_ticker.drop_duplicates("ticker", keep="first")

    matched = pool.merge(latest_by_ticker, on="ticker", how="left", suffixes=("_pool", "_filing"))
    matched["matched"] = matched["adsh"].notna()
    matched["ticker_match_method"] = matched["ticker_match_method"].mask(matched["matched"], "notes_trading_symbol")
    return matched


def load_annual_filings_from_dir(notes_dir: Path) -> pd.DataFrame:
    sub = pd.read_csv(
        notes_dir / "sub.tsv",
        sep="\t",
        dtype=str,
        usecols=[
            "adsh",
            "cik",
            "name",
            "sic",
            "countryba",
            "stprba",
            "cityba",
            "countryinc",
            "stprinc",
            "wksi",
            "fye",
            "form",
            "period",
            "fy",
            "fp",
            "filed",
            "accepted",
            "detail",
            "instance",
            "pubfloatusd",
            "floatdate",
        ],
    )
    sub = sub[sub["form"].isin(ANNUAL_FORMS) & sub["detail"].fillna("").isin(["1", "true", "TRUE"])].copy()
    sub["batch"] = notes_dir.name
    sub["notes_dir"] = str(notes_dir)
    sub["filed_num"] = pd.to_numeric(sub["filed"], errors="coerce")
    sub["period_num"] = pd.to_numeric(sub["period"], errors="coerce")

    txt = pd.read_csv(notes_dir / "txt.tsv", sep="\t", dtype=str, usecols=["adsh", "tag", "value"])
    trading = txt[txt["tag"] == "TradingSymbol"][["adsh", "value"]].dropna().drop_duplicates()
    trading["ticker"] = trading["value"].map(normalize_pool_ticker)
    trading = trading[["adsh", "ticker"]].drop_duplicates()
    sub = sub.merge(trading, on="adsh", how="left")
    return sub


def find_single_company_stock_row(
    stock_pool: pd.DataFrame,
    notes_dir: Path,
    *,
    stock_code: str | None = None,
    company_name: str | None = None,
    adsh: str | None = None,
) -> pd.Series:
    filings = load_annual_filings_from_dir(notes_dir)
    pool_candidates = stock_pool.copy()

    stock_code_clean = clean_text(stock_code)
    company_name_clean = clean_text(company_name)
    adsh_clean = clean_text(adsh)

    if stock_code_clean:
        pool_candidates = pool_candidates[pool_candidates["SECU_CODE"].map(clean_text) == stock_code_clean]
    if company_name_clean:
        pool_candidates = pool_candidates[pool_candidates["SECUNAME"].map(clean_text) == company_name_clean]

    if pool_candidates.empty:
        raise ValueError("未在股票池中找到目标公司。")

    if adsh_clean:
        filing_candidates = filings[filings["adsh"].map(clean_text) == adsh_clean].copy()
    else:
        target_tickers = set(pool_candidates["ticker"].map(clean_text))
        filing_candidates = filings[filings["ticker"].map(clean_text).isin(target_tickers)].copy()

    if filing_candidates.empty and company_name_clean:
        filing_candidates = filings[filings["name"].fillna("").str.contains(company_name_clean, case=False, regex=False)].copy()

    if filing_candidates.empty:
        raise ValueError(f"未在 {notes_dir} 中找到目标公司的年报 filing。")

    filing_candidates = filing_candidates.sort_values(["filed_num", "period_num", "adsh"], ascending=[False, False, False])
    filing_row = filing_candidates.iloc[0]

    if clean_text(filing_row.get("ticker")):
        matching_pool = pool_candidates[pool_candidates["ticker"].map(clean_text) == clean_text(filing_row["ticker"])]
        if not matching_pool.empty:
            pool_candidates = matching_pool

    pool_row = pool_candidates.iloc[0]
    merged = {**pool_row.to_dict(), **filing_row.to_dict()}
    merged["ticker_match_method"] = "notes_trading_symbol" if clean_text(filing_row.get("ticker")) else "manual_company_lookup"
    merged["notes_dir"] = str(notes_dir)
    return pd.Series(merged)


def pick_text_value(group: pd.DataFrame, priorities: list[str]) -> tuple[str | None, str]:
    for tag in priorities:
        rows = group[group["tag"] == tag]
        if not rows.empty:
            values = rows["value"].map(clean_text)
            values = values[values != ""]
            if not values.empty:
                return tag, max(values, key=len)
    return None, ""


def choose_company_metric(
    num_df: pd.DataFrame,
    report_period: str,
    tag_priority: list[str],
) -> dict[str, Any]:
    candidates = num_df[
        num_df["tag"].isin(tag_priority)
        & (num_df["dimh"] == "0x00000000")
        & (num_df["value_num"].notna())
        & (num_df["ddate"] == report_period)
    ].copy()
    if candidates.empty:
        return {"value": None, "tag": None, "version": None, "uom": None}
    nonzero_candidates = candidates[candidates["value_num"] != 0].copy()
    if not nonzero_candidates.empty:
        candidates = nonzero_candidates

    def qtrs_rank(tag: str, qtrs: Any) -> tuple[int, int]:
        q = int(qtrs) if pd.notna(qtrs) else -1
        if tag in REVENUE_TAG_PRIORITY:
            return (0 if q == 4 else 1, -q)
        return (0 if q == 0 else 1, -q)

    candidates["priority_rank"] = candidates["tag"].map(lambda x: tag_priority.index(x) if x in tag_priority else len(tag_priority))
    candidates["qtrs_rank"] = candidates.apply(lambda r: qtrs_rank(r["tag"], r["qtrs_num"]), axis=1)
    candidates = candidates.sort_values(["priority_rank", "qtrs_rank", "value_num"], ascending=[True, True, False])
    row = candidates.iloc[0]
    return {
        "value": row["value_num"],
        "tag": row["tag"],
        "version": row["version"],
        "uom": row["uom"],
    }


def choose_segment_period_candidates(num_df: pd.DataFrame, report_period: str) -> pd.DataFrame:
    revenue_like = num_df[
        num_df["tag"].isin(REVENUE_TAG_PRIORITY)
        & (num_df["dimh"] != "0x00000000")
        & (num_df["value_num"].notna())
        & (num_df["value_num"] > 0)
        & (num_df["ddate"] == report_period)
    ].copy()
    if revenue_like.empty:
        return revenue_like

    revenue_like["tag_rank"] = revenue_like["tag"].map(lambda x: REVENUE_TAG_PRIORITY.index(x) if x in REVENUE_TAG_PRIORITY else 99)
    revenue_like["qtrs_rank"] = revenue_like["qtrs_num"].map(lambda q: 0 if pd.notna(q) and int(q) == 4 else 1)
    revenue_like = revenue_like.sort_values(
        ["qtrs_rank", "tag_rank", "value_num"],
        ascending=[True, True, False],
        kind="stable",
    )
    best_qtrs = revenue_like.iloc[0]["qtrs"]
    best_tag = revenue_like.iloc[0]["tag"]
    filtered = revenue_like[(revenue_like["qtrs"] == best_qtrs) & (revenue_like["tag"] == best_tag)].copy()
    return filtered


def choose_text_metadata(txt_df: pd.DataFrame, tag: str) -> str | None:
    rows = txt_df[txt_df["tag"] == tag]
    if rows.empty:
        return None
    values = rows["value"].map(clean_text)
    values = values[values != ""]
    if values.empty:
        return None
    return values.iloc[0]


def load_report_lookup(notes_dir: Path) -> pd.DataFrame:
    ren = pd.read_csv(notes_dir / "ren.tsv", sep="\t", dtype=str, usecols=["adsh", "report", "shortname", "longname"])
    ren["report"] = pd.to_numeric(ren["report"], errors="coerce")
    ren = ren.sort_values(["adsh", "report"]).drop_duplicates(["adsh", "report"])
    return ren


def load_notes_cache(notes_dir: Path) -> dict[str, pd.DataFrame]:
    key = notes_dir.name
    cached = NOTES_CACHE.get(key)
    if cached is not None:
        return cached

    txt = pd.read_csv(notes_dir / "txt.tsv", sep="\t", dtype=str, usecols=["adsh", "tag", "value"])
    num = pd.read_csv(
        notes_dir / "num.tsv",
        sep="\t",
        dtype=str,
        usecols=["adsh", "tag", "version", "ddate", "qtrs", "uom", "dimh", "value"],
    )
    num["value_num"] = pd.to_numeric(num["value"], errors="coerce")
    num["qtrs_num"] = pd.to_numeric(num["qtrs"], errors="coerce")
    num["dimh"] = num["dimh"].fillna("0x00000000")

    dim = pd.read_csv(notes_dir / "dim.tsv", sep="\t", dtype=str)
    pre = pd.read_csv(notes_dir / "pre.tsv", sep="\t", dtype=str, usecols=["adsh", "report", "line", "tag", "version", "plabel"])
    ren = load_report_lookup(notes_dir)

    cached = {"txt": txt, "num": num, "dim": dim, "pre": pre, "ren": ren}
    NOTES_CACHE[key] = cached
    return cached


def build_segment_rows(
    adsh: str,
    report_period: str,
    total_revenue: float | None,
    revenue_uom: str | None,
    num_df: pd.DataFrame,
    dim_df: pd.DataFrame,
    pre_df: pd.DataFrame,
    ren_df: pd.DataFrame,
    evidence_text: str,
) -> tuple[list[dict[str, Any]], list[str], str | None, bool]:
    revenue_candidates = choose_segment_period_candidates(num_df, report_period)
    if revenue_candidates.empty:
        return [], [], None, False

    dim_map = dim_df.set_index("dimhash").to_dict("index")
    pre_map = pre_df.sort_values(["adsh", "report", "line"]).drop_duplicates(["adsh", "tag", "version", "report"])
    pre_keyed = {(r["tag"], r["version"], r["report"]): r for r in pre_map.to_dict("records")}
    ren_keyed = {(r["adsh"], int(r["report"])): r for r in ren_df.dropna(subset=["report"]).to_dict("records")}

    gross_candidates = num_df[
        num_df["tag"].isin(["GrossProfit", "GrossProfitLoss", "GrossMargin", "GrossMarginPercentage"])
        & (num_df["ddate"] == report_period)
        & (num_df["dimh"] != "0x00000000")
        & (num_df["value_num"].notna())
    ].copy()

    rows: list[dict[str, Any]] = []
    dims_available: set[str] = set()
    has_margin = False
    seen_keys: set[tuple[str, str, str]] = set()

    revenue_candidates["tag_rank"] = revenue_candidates["tag"].map(lambda x: REVENUE_TAG_PRIORITY.index(x) if x in REVENUE_TAG_PRIORITY else 99)
    revenue_candidates = revenue_candidates.sort_values(["tag_rank", "value_num"], ascending=[True, False])

    for _, fact in revenue_candidates.iterrows():
        dim_info = dim_map.get(fact["dimh"])
        if not dim_info:
            continue
        segments_text = clean_text(dim_info.get("segments"))
        if not segments_text or "=" not in segments_text:
            continue

        parsed_dims: list[tuple[str, str, str]] = []
        for item in segments_text.split(";"):
            if "=" not in item:
                continue
            axis_name, member_name = item.split("=", 1)
            axis_name = clean_text(axis_name)
            member_name = clean_text(member_name)
            seg_type = dimension_category(axis_name) or "other_dimension"
            parsed_dims.append((seg_type, axis_name, member_name))
        if not parsed_dims:
            continue

        linked_pre = None
        source_report = None
        source_short = None
        source_long = None
        pre_candidates = pre_df[(pre_df["tag"] == fact["tag"]) & (pre_df["version"] == fact["version"])]
        if not pre_candidates.empty:
            linked_pre = pre_candidates.sort_values(["report", "line"]).iloc[0]
            report_no = pd.to_numeric(linked_pre["report"], errors="coerce")
            if pd.notna(report_no):
                source_report = int(report_no)
                ren_row = ren_keyed.get((adsh, int(report_no)))
                if ren_row:
                    source_short = clean_text(ren_row.get("shortname"))
                    source_long = clean_text(ren_row.get("longname"))

        gross_profit = None
        gross_margin = None
        gross_uom = None
        gross_match = gross_candidates[gross_candidates["dimh"] == fact["dimh"]]
        if not gross_match.empty:
            gp = gross_match[gross_match["tag"].isin(["GrossProfit", "GrossProfitLoss"])]
            gm = gross_match[gross_match["tag"].isin(["GrossMargin", "GrossMarginPercentage"])]
            if not gp.empty:
                gp = gp.sort_values("value_num", ascending=False).iloc[0]
                gross_profit = gp["value_num"]
                gross_uom = gp["uom"]
            if not gm.empty:
                gm = gm.sort_values("value_num", ascending=False).iloc[0]
                gross_margin = gm["value_num"] / 100 if gm["value_num"] and gm["value_num"] > 1 else gm["value_num"]
            if gross_margin is None and gross_profit is not None and fact["value_num"]:
                ratio = gross_profit / fact["value_num"]
                gross_margin = ratio if -1 <= ratio <= 1.5 else None
            if gross_profit is not None or gross_margin is not None:
                has_margin = True

        for seg_type, axis_name, member_name in parsed_dims:
            dedupe_key = (seg_type, axis_name, member_name)
            if dedupe_key in seen_keys:
                continue
            seen_keys.add(dedupe_key)
            dims_available.add(seg_type)

            member_label = camel_to_words(member_name)
            snippet = extract_snippet(evidence_text, member_label)
            if not snippet:
                snippet = build_segment_desc_fallback(member_label, source_short, source_long, segments_text)
            rows.append(
                {
                    "segment_name": member_label,
                    "segment_type": seg_type,
                    "segment_axis": axis_name,
                    "segment_member_raw": member_name,
                    "segment_member_country_cn": translate_segment_member_cn(member_name, seg_type),
                    "segment_revenue": fact["value_num"],
                    "segment_revenue_uom": fact["uom"] or revenue_uom,
                    "segment_revenue_ratio": fact["value_num"] / total_revenue if total_revenue else None,
                    "segment_gross_profit": gross_profit,
                    "segment_gross_profit_uom": gross_uom,
                    "segment_gross_margin": gross_margin,
                    "segment_desc": snippet,
                    "source_section": "notes",
                    "source_tag": fact["tag"],
                    "source_version": fact["version"],
                    "source_group_tag": fact["tag"],
                    "source_context": None,
                    "source_dim_hash": fact["dimh"],
                    "source_segments_text": segments_text,
                    "source_report": source_report,
                    "source_report_shortname": source_short,
                    "source_report_longname": source_long,
                }
            )

    primary_dim = None
    if rows:
        counts: dict[str, int] = {}
        for row in rows:
            counts[row["segment_type"]] = counts.get(row["segment_type"], 0) + 1
        primary_dim = sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0]

    rows = sorted(rows, key=lambda item: (item["segment_type"], -(item["segment_revenue"] or 0), item["segment_name"]))
    return rows, sorted(dims_available), primary_dim, has_margin


def build_raw_segment_dimension_rows(
    adsh: str,
    report_period: str,
    num_df: pd.DataFrame,
    dim_df: pd.DataFrame,
    pre_df: pd.DataFrame,
    ren_df: pd.DataFrame,
) -> pd.DataFrame:
    num = num_df.copy()
    num["dimh"] = num["dimh"].replace("", "0x00000000")
    revenue_candidates = choose_segment_period_candidates(num, report_period)
    if revenue_candidates.empty:
        return pd.DataFrame(
            columns=[
                "segment_type",
                "segment_axis",
                "segment_member_raw",
                "segment_name",
                "source_dim_hash",
                "source_segments_text",
                "source_tag",
                "source_version",
                "source_report",
                "source_report_shortname",
                "source_report_longname",
                "segment_revenue",
                "segment_revenue_uom",
            ]
        )

    revenue_candidates = revenue_candidates.drop_duplicates(["tag", "version", "ddate", "qtrs", "uom", "dimh", "value"], keep="first")

    dim_lookup = dim_df.set_index("dimhash").to_dict("index") if not dim_df.empty else {}
    ren_keyed = {(r["adsh"], int(r["report"])): r for r in ren_df.dropna(subset=["report"]).to_dict("records")}

    rows: list[dict[str, Any]] = []
    for _, fact in revenue_candidates.iterrows():
        dim_info = dim_lookup.get(clean_text(fact["dimh"]))
        if not dim_info:
            continue
        segments_text = clean_text(dim_info.get("segments"))
        if not segments_text or "=" not in segments_text:
            continue

        pre_candidates = pre_df[(pre_df["tag"] == fact["tag"]) & (pre_df["version"] == fact["version"])]
        source_report = None
        source_short = None
        source_long = None
        if not pre_candidates.empty:
            linked_pre = pre_candidates.sort_values(["report", "line"]).iloc[0]
            report_no = pd.to_numeric(linked_pre["report"], errors="coerce")
            if pd.notna(report_no):
                source_report = int(report_no)
                ren_row = ren_keyed.get((adsh, int(report_no)))
                if ren_row:
                    source_short = clean_text(ren_row.get("shortname"))
                    source_long = clean_text(ren_row.get("longname"))

        for item in segments_text.split(";"):
            if "=" not in item:
                continue
            axis_name, member_name = item.split("=", 1)
            axis_name = clean_text(axis_name)
            member_name = clean_text(member_name)
            seg_type = dimension_category(axis_name) or "other_dimension"
            rows.append(
                {
                    "segment_type": seg_type,
                    "segment_axis": axis_name,
                    "segment_member_raw": member_name,
                    "segment_member_country_cn": translate_segment_member_cn(member_name, seg_type),
                    "segment_name": camel_to_words(member_name),
                    "source_dim_hash": clean_text(fact["dimh"]),
                    "source_segments_text": segments_text,
                    "source_tag": fact["tag"],
                    "source_version": fact["version"],
                    "source_report": source_report,
                    "source_report_shortname": source_short,
                    "source_report_longname": source_long,
                    "segment_revenue": fact["value_num"],
                    "segment_revenue_uom": fact["uom"],
                }
            )

    return pd.DataFrame(rows)


def extract_snippet(text: str, term: str) -> str:
    text = clean_text(text)
    term = clean_text(term)
    if not text or not term:
        return ""
    lower_text = text.lower()
    lower_term = term.lower()
    pos = lower_text.find(lower_term)
    if pos < 0:
        compact_pos = lower_text.replace(" ", "").find(lower_term.replace(" ", ""))
        if compact_pos < 0:
            return text[:500]
        return text[:500]
    start = max(0, pos - 200)
    end = min(len(text), pos + len(term) + 240)
    return text[start:end]


def build_segment_desc_fallback(
    member_label: str,
    source_short: str | None,
    source_long: str | None,
    segments_text: str,
) -> str:
    report_label = clean_text(source_short) or clean_text(source_long)
    context = clean_text(segments_text)
    parts = [f"Derived from segment fact for {member_label}."]
    if report_label:
        parts.append(f"Source report: {report_label}.")
    if context:
        parts.append(f"Dimension context: {context}")
    return clean_text(" ".join(parts))


def build_company_and_segment_rows(stock_row: pd.Series) -> list[dict[str, Any]]:
    if pd.isna(stock_row.get("adsh")):
        return []

    batch = stock_row["batch"]
    notes_dir_value = stock_row.get("notes_dir")
    notes_dir = Path(notes_dir_value) if clean_text(notes_dir_value) else Path(batch)
    adsh = stock_row["adsh"]
    report_period_raw = clean_text(stock_row.get("period"))
    report_period = report_period_raw if report_period_raw else ""
    cache = load_notes_cache(notes_dir)
    txt = cache["txt"][cache["txt"]["adsh"] == adsh].copy()
    num = cache["num"][cache["num"]["adsh"] == adsh].copy()
    dim = cache["dim"]
    pre = cache["pre"][cache["pre"]["adsh"] == adsh].copy()
    ren = cache["ren"][cache["ren"]["adsh"] == adsh].copy()

    text_results: dict[str, tuple[str | None, str]] = {}
    for field_name, priorities in TEXT_TAG_PRIORITIES.items():
        text_results[field_name] = pick_text_value(txt, priorities)

    main_business_tag, main_business_desc = text_results["main_business_tag"]
    segment_text_tag, segment_text = text_results["segment_text_tag"]
    revenue_note_tag, revenue_note_text = text_results["revenue_note_tag"]

    if not main_business_desc:
        main_business_desc = segment_text or revenue_note_text
        if not main_business_tag:
            main_business_tag = segment_text_tag or revenue_note_tag

    total_revenue_metric = choose_company_metric(num, report_period, REVENUE_TAG_PRIORITY)
    gross_profit_metric = choose_company_metric(num, report_period, ["GrossProfit", "GrossProfitLoss"])
    operating_income_metric = choose_company_metric(num, report_period, ["OperatingIncomeLoss"])
    net_income_metric = choose_company_metric(num, report_period, ["NetIncomeLoss"])
    assets_metric = choose_company_metric(num, report_period, ["Assets"])
    liabilities_metric = choose_company_metric(num, report_period, ["Liabilities"])
    equity_metric = choose_company_metric(
        num,
        report_period,
        ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
    )
    cash_metric = choose_company_metric(num, report_period, ["CashAndCashEquivalentsAtCarryingValue"])
    ocf_metric = choose_company_metric(num, report_period, ["NetCashProvidedByUsedInOperatingActivities"])
    capex_metric = choose_company_metric(num, report_period, ["PaymentsToAcquirePropertyPlantAndEquipment"])
    rd_metric = choose_company_metric(num, report_period, ["ResearchAndDevelopmentExpense"])
    da_metric = choose_company_metric(num, report_period, ["DepreciationDepletionAndAmortization"])
    eps_basic_metric = choose_company_metric(num, report_period, ["EarningsPerShareBasic"])
    eps_diluted_metric = choose_company_metric(num, report_period, ["EarningsPerShareDiluted"])
    shares_outstanding_metric = choose_company_metric(num, report_period, ["CommonStockSharesOutstanding"])
    shares_issued_metric = choose_company_metric(num, report_period, ["CommonStockSharesIssued"])
    shares_authorized_metric = choose_company_metric(num, report_period, ["CommonStockSharesAuthorized"])
    weighted_avg_diluted_metric = choose_company_metric(
        num,
        report_period,
        ["WeightedAverageNumberOfDilutedSharesOutstanding", "WeightedAverageNumberOfShareOutstandingBasicAndDiluted"],
    )

    evidence_text = segment_text or revenue_note_text or main_business_desc
    segment_rows, dims_available, primary_dim, has_margin = build_segment_rows(
        adsh=adsh,
        report_period=report_period,
        total_revenue=total_revenue_metric["value"],
        revenue_uom=total_revenue_metric["uom"],
        num_df=num,
        dim_df=dim,
        pre_df=pre,
        ren_df=ren,
        evidence_text=evidence_text,
    )
    disclosure_dims_payload = list(dims_available)
    if not disclosure_dims_payload and (segment_text or revenue_note_text):
        disclosure_dims_payload = ["narrative_only"]

    company_row = {
        "row_type": "company",
        "stock_code": stock_row["SECU_CODE"],
        "company_name": stock_row["SECUNAME"],
        "CUSIP": stock_row["CUSIP"],
        "ticker": stock_row["ticker"],
        "ticker_match_method": stock_row["ticker_match_method"],
        "cik": clean_text(stock_row.get("cik")),
        "company_name_filing": clean_text(stock_row.get("name")),
        "report_type": clean_text(stock_row.get("form")),
        "report_period": normalize_report_date(report_period),
        "filed_date": normalize_report_date(stock_row.get("filed")),
        "accepted_datetime": clean_text(stock_row.get("accepted")),
        "filing_month_batch": batch,
        "adsh": adsh,
        "instance": clean_text(stock_row.get("instance")),
        "sic": clean_text(stock_row.get("sic")),
        "industry_l1": stock_row.get("CCXDF_CN"),
        "industry_l2": stock_row.get("CCXDS_CN"),
        "industry_l3": stock_row.get("CCXDT_CN"),
        "industry_l4": stock_row.get("CCXDFourth_CN"),
        "fiscal_year": clean_text(stock_row.get("fy")),
        "fiscal_period": clean_text(stock_row.get("fp")),
        "fiscal_year_end": clean_text(stock_row.get("fye")),
        "wksi": clean_text(stock_row.get("wksi")),
        "filer_category": choose_text_metadata(txt, "EntityFilerCategory"),
        "sec_file_number": choose_text_metadata(txt, "EntityFileNumber"),
        "tax_id": choose_text_metadata(txt, "EntityTaxIdentificationNumber"),
        "local_phone": choose_text_metadata(txt, "LocalPhoneNumber"),
        "public_float_usd": pd.to_numeric(stock_row.get("pubfloatusd"), errors="coerce"),
        "float_date": normalize_report_date(stock_row.get("floatdate")),
        "country_business_address": clean_text(stock_row.get("countryba")),
        "state_business_address": clean_text(stock_row.get("stprba")),
        "city_business_address": clean_text(stock_row.get("cityba")),
        "country_incorporation": clean_text(stock_row.get("countryinc")),
        "state_incorporation": clean_text(stock_row.get("stprinc")),
        "main_business_desc": main_business_desc[:30000],
        "main_business_tag": main_business_tag,
        "segment_text_tag": segment_text_tag,
        "revenue_note_tag": revenue_note_tag,
        "total_revenue": total_revenue_metric["value"],
        "total_revenue_tag": total_revenue_metric["tag"],
        "total_revenue_version": total_revenue_metric["version"],
        "total_revenue_uom": total_revenue_metric["uom"],
        "gross_profit": gross_profit_metric["value"],
        "gross_profit_tag": gross_profit_metric["tag"],
        "operating_income": operating_income_metric["value"],
        "operating_income_tag": operating_income_metric["tag"],
        "net_income": net_income_metric["value"],
        "net_income_tag": net_income_metric["tag"],
        "assets": assets_metric["value"],
        "assets_uom": assets_metric["uom"],
        "liabilities": liabilities_metric["value"],
        "liabilities_uom": liabilities_metric["uom"],
        "equity": equity_metric["value"],
        "equity_uom": equity_metric["uom"],
        "cash_and_equivalents": cash_metric["value"],
        "cash_uom": cash_metric["uom"],
        "operating_cash_flow": ocf_metric["value"],
        "operating_cash_flow_uom": ocf_metric["uom"],
        "capex": capex_metric["value"],
        "capex_uom": capex_metric["uom"],
        "rd_expense": rd_metric["value"],
        "rd_expense_uom": rd_metric["uom"],
        "depreciation_amortization": da_metric["value"],
        "depreciation_amortization_uom": da_metric["uom"],
        "eps_basic": eps_basic_metric["value"],
        "eps_diluted": eps_diluted_metric["value"],
        "shares_outstanding": shares_outstanding_metric["value"],
        "shares_outstanding_uom": shares_outstanding_metric["uom"],
        "shares_issued": shares_issued_metric["value"],
        "shares_issued_uom": shares_issued_metric["uom"],
        "shares_authorized": shares_authorized_metric["value"],
        "shares_authorized_uom": shares_authorized_metric["uom"],
        "weighted_avg_diluted_shares": weighted_avg_diluted_metric["value"],
        "weighted_avg_diluted_shares_uom": weighted_avg_diluted_metric["uom"],
        "segment_disclosure_flag": bool(segment_text or revenue_note_text or segment_rows),
        "segment_count": sum(1 for row in segment_rows if row["segment_type"] == primary_dim) if primary_dim else len(segment_rows),
        "primary_disclosure_dim": primary_dim or ("narrative_only" if segment_text or revenue_note_text else None),
        "disclosure_dims_available": json.dumps(disclosure_dims_payload, ensure_ascii=False),
        "has_revenue_breakdown": bool(segment_rows),
        "has_margin_breakdown": has_margin,
        "segment_id": None,
        "segment_name": None,
        "segment_type": None,
        "segment_axis": None,
        "segment_member_raw": None,
        "segment_member_country_cn": None,
        "segment_revenue": None,
        "segment_revenue_uom": None,
        "segment_revenue_ratio": None,
        "segment_gross_profit": None,
        "segment_gross_profit_uom": None,
        "segment_gross_margin": None,
        "segment_desc": None,
        "source_section": "company",
        "source_tag": main_business_tag or segment_text_tag or revenue_note_tag,
        "source_version": None,
        "source_group_tag": None,
        "source_context": None,
        "source_dim_hash": None,
        "source_segments_text": None,
        "source_report": None,
        "source_report_shortname": None,
        "source_report_longname": None,
        "evidence_text": (segment_text or revenue_note_text or main_business_desc)[:6000],
        "extraction_status": "ok" if total_revenue_metric["value"] is not None else "partial",
        "notes_file_batch": batch,
    }

    rows = [{col: company_row.get(col) for col in ROW_COLUMNS}]

    for idx, segment in enumerate(segment_rows, start=1):
        row = dict(company_row)
        row.update(segment)
        row["row_type"] = "segment"
        row["segment_id"] = f"{company_row['stock_code']}|{company_row['report_period'] or 'NA'}|{idx}"
        row["evidence_text"] = segment.get("segment_desc") or company_row["evidence_text"]
        rows.append({col: row.get(col) for col in ROW_COLUMNS})

    return rows


def validate_company_segments(stock_row: pd.Series, rows: list[dict[str, Any]]) -> dict[str, Any]:
    if pd.isna(stock_row.get("adsh")):
        return {
            "stock_code": stock_row.get("SECU_CODE"),
            "adsh": None,
            "expected_segment_types": [],
            "actual_segment_types": [],
            "missing_segment_types": [],
            "expected_rows": 0,
            "actual_rows": 0,
        }

    notes_dir_value = stock_row.get("notes_dir")
    notes_dir = Path(notes_dir_value) if clean_text(notes_dir_value) else Path(stock_row["batch"])
    adsh = stock_row["adsh"]
    report_period = clean_text(stock_row.get("period"))
    cache = load_notes_cache(notes_dir)
    num = cache["num"][cache["num"]["adsh"] == adsh].copy()
    dim = cache["dim"]
    pre = cache["pre"][cache["pre"]["adsh"] == adsh].copy()
    ren = cache["ren"][cache["ren"]["adsh"] == adsh].copy()

    raw_df = build_raw_segment_dimension_rows(
        adsh=adsh,
        report_period=report_period,
        num_df=num,
        dim_df=dim,
        pre_df=pre,
        ren_df=ren,
    )
    expected_types = sorted(set(raw_df["segment_type"].dropna().astype(str))) if not raw_df.empty else []
    actual_df = pd.DataFrame(rows)
    actual_seg = actual_df[actual_df.get("row_type").eq("segment")] if not actual_df.empty else pd.DataFrame()
    actual_types = sorted(set(actual_seg["segment_type"].dropna().astype(str))) if not actual_seg.empty else []
    missing_types = sorted(set(expected_types) - set(actual_types))
    return {
        "stock_code": clean_text(stock_row.get("SECU_CODE") or stock_row.get("stock_code")),
        "adsh": adsh,
        "expected_segment_types": expected_types,
        "actual_segment_types": actual_types,
        "missing_segment_types": missing_types,
        "expected_rows": int(len(raw_df)),
        "actual_rows": int(len(actual_seg)),
    }


def add_description_rows(workbook_path: Path) -> None:
    wb = load_workbook(workbook_path)
    ws = wb[wb.sheetnames[0]]
    ws.insert_rows(1, 2)
    ws.freeze_panes = "A4"

    note_fill = PatternFill(fill_type="solid", fgColor="EDEDED")
    desc_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    bold = Font(bold=True)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
    note_cell = ws.cell(1, 1)
    note_cell.value = TABLE_NOTE
    note_cell.fill = note_fill
    note_cell.font = bold
    note_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    for col_idx in range(1, ws.max_column + 1):
        header_cell = ws.cell(3, col_idx)
        desc_cell = ws.cell(2, col_idx)
        field = clean_text(header_cell.value)
        desc_cell.value = FIELD_DESCRIPTIONS.get(field, "")
        desc_cell.fill = desc_fill
        desc_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        header_cell.fill = header_fill
        header_cell.font = bold
        header_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 36
    ws.row_dimensions[2].height = 44
    ws.row_dimensions[3].height = 24
    wb.save(workbook_path)


def parse_inline_number(element: etree._Element) -> float | None:
    text = clean_text("".join(element.itertext()))
    if not text or text in {"-", "N/A", "n/a"}:
        return 0.0 if element.get("format") == "ixt:fixed-zero" else None
    text = text.replace(",", "").replace("$", "").replace("%", "")
    text = text.replace("(", "-").replace(")", "").replace("\u2212", "-")
    try:
        value = float(text)
    except ValueError:
        return None
    if element.get("sign") == "-":
        value = -abs(value)
    scale = element.get("scale")
    if scale:
        try:
            value *= 10 ** int(scale)
        except ValueError:
            pass
    return value


def qname_local(value: str | None) -> str:
    if not value:
        return ""
    return value.split(":", 1)[-1]


def extract_html_report(report_file: Path) -> dict[str, Any]:
    parser = etree.XMLParser(recover=True, huge_tree=True)
    root = etree.parse(str(report_file), parser).getroot()

    metadata_values: dict[str, list[str]] = {}
    text_blocks: dict[str, list[str]] = {}
    non_fraction_rows: list[dict[str, Any]] = []

    for fact in root.xpath("//ix:nonNumeric", namespaces={
        "ix": "http://www.xbrl.org/2013/inlineXBRL"
    }):
        local = qname_local(fact.get("name", ""))
        text = clean_text("".join(fact.itertext()))
        if not text:
            continue
        metadata_values.setdefault(local, []).append(text)
        if local.endswith("TextBlock") or local.endswith("Explanatory"):
            text_blocks.setdefault(local, []).append(text)

    for fact in root.xpath("//ix:nonFraction", namespaces={
        "ix": "http://www.xbrl.org/2013/inlineXBRL"
    }):
        non_fraction_rows.append(
            {
                "tag": qname_local(fact.get("name", "")),
                "context_ref": fact.get("contextRef"),
                "unit_ref": fact.get("unitRef"),
                "value": parse_inline_number(fact),
            }
        )

    report_period = normalize_report_date((metadata_values.get("DocumentPeriodEndDate") or [""])[0])
    report_type = clean_text((metadata_values.get("DocumentType") or [""])[0])
    company_name_filing = clean_text((metadata_values.get("EntityRegistrantName") or [""])[0])

    main_business_tag = None
    main_business_desc = ""
    segment_text_tag = None
    segment_text = ""
    revenue_note_tag = None
    revenue_note_text = ""

    for tag in TEXT_TAG_PRIORITIES["main_business_tag"]:
        values = text_blocks.get(tag)
        if values:
            main_business_tag = tag
            main_business_desc = max(values, key=len)
            break
    for tag in TEXT_TAG_PRIORITIES["segment_text_tag"]:
        values = text_blocks.get(tag)
        if values:
            segment_text_tag = tag
            segment_text = max(values, key=len)
            break
    for tag in TEXT_TAG_PRIORITIES["revenue_note_tag"]:
        values = text_blocks.get(tag)
        if values:
            revenue_note_tag = tag
            revenue_note_text = max(values, key=len)
            break
    if not main_business_desc:
        main_business_desc = segment_text or revenue_note_text
        main_business_tag = segment_text_tag or revenue_note_tag

    nf = pd.DataFrame(non_fraction_rows)
    if nf.empty:
        nf = pd.DataFrame(columns=["tag", "context_ref", "unit_ref", "value"])

    def html_metric(tag_priority: list[str], fallback_patterns: list[str] | None = None) -> dict[str, Any]:
        rows = nf[nf["tag"].isin(tag_priority) & nf["value"].notna() & (nf["value"] != 0)].copy()
        if rows.empty and fallback_patterns:
            pattern = "|".join(fallback_patterns)
            rows = nf[nf["tag"].str.contains(pattern, case=False, na=False) & nf["value"].notna() & (nf["value"] != 0)].copy()
        if rows.empty:
            return {"value": None, "tag": None, "uom": None}
        rows["priority"] = rows["tag"].map(lambda x: tag_priority.index(x) if x in tag_priority else 99)
        rows = rows.sort_values(["priority", "value"], ascending=[True, False])
        row = rows.iloc[0]
        return {"value": row["value"], "tag": row["tag"], "uom": row["unit_ref"]}

    total_revenue_metric = html_metric(
        REVENUE_TAG_PRIORITY,
        ["revenue", "sales", "turnover", "contractswithcustomers", "income"],
    )
    gross_profit_metric = html_metric(
        ["GrossProfit", "GrossProfitLoss"],
        ["grossprofit", "grossmargin"],
    )
    operating_income_metric = html_metric(
        ["OperatingIncomeLoss"],
        ["operatingincome", "operatingprofit", "incomefromoperations"],
    )
    net_income_metric = html_metric(
        ["NetIncomeLoss"],
        ["netincome", "profitloss", "profitattributable", "incomeloss"],
    )
    assets_metric = html_metric(["Assets"], ["^assets$", "totalassets"])
    liabilities_metric = html_metric(["Liabilities"], ["^liabilities$", "totalliabilities"])
    equity_metric = html_metric(
        ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
        ["equity", "totalequity", "stockholdersequity", "shareholdersequity"],
    )
    cash_metric = html_metric(
        ["CashAndCashEquivalentsAtCarryingValue"],
        ["cashandcashequivalents", "cash", "cashresources"],
    )
    ocf_metric = html_metric(
        ["NetCashProvidedByUsedInOperatingActivities"],
        ["cashflowsfromusedinoperating", "netcashgeneratedfromoperating", "operatingactivities"],
    )
    capex_metric = html_metric(
        ["PaymentsToAcquirePropertyPlantAndEquipment"],
        ["propertyplantandequipment", "purchaseofproperty", "capitalexpenditure", "additions"],
    )
    rd_metric = html_metric(
        ["ResearchAndDevelopmentExpense"],
        ["researchanddevelopment", "researchdevelopment"],
    )
    da_metric = html_metric(
        ["DepreciationDepletionAndAmortization"],
        ["depreciation", "amortization", "depletion"],
    )
    eps_basic_metric = html_metric(
        ["EarningsPerShareBasic"],
        ["basicearningspershare", "basiclosspershare"],
    )
    eps_diluted_metric = html_metric(
        ["EarningsPerShareDiluted"],
        ["dilutedearningspershare", "dilutedlosspershare"],
    )
    shares_outstanding_metric = html_metric(
        ["CommonStockSharesOutstanding"],
        ["sharesoutstanding", "entitycommonstocksharesoutstanding"],
    )
    shares_issued_metric = html_metric(
        ["CommonStockSharesIssued"],
        ["sharesissued", "stockissued"],
    )
    shares_authorized_metric = html_metric(
        ["CommonStockSharesAuthorized"],
        ["sharesauthorized"],
    )
    weighted_avg_diluted_metric = html_metric(
        ["WeightedAverageNumberOfDilutedSharesOutstanding", "WeightedAverageNumberOfShareOutstandingBasicAndDiluted"],
        ["weightedaverage.*diluted.*shares", "weightedaverage.*share"],
    )

    return {
        "company_name_filing": company_name_filing,
        "report_type": report_type,
        "report_period": report_period,
        "main_business_desc": main_business_desc[:30000] if main_business_desc else "",
        "main_business_tag": main_business_tag,
        "segment_text_tag": segment_text_tag,
        "revenue_note_tag": revenue_note_tag,
        "evidence_text": (segment_text or revenue_note_text or main_business_desc)[:6000],
        "total_revenue": total_revenue_metric["value"],
        "total_revenue_tag": total_revenue_metric["tag"],
        "total_revenue_uom": total_revenue_metric["uom"],
        "gross_profit": gross_profit_metric["value"],
        "gross_profit_tag": gross_profit_metric["tag"],
        "operating_income": operating_income_metric["value"],
        "operating_income_tag": operating_income_metric["tag"],
        "net_income": net_income_metric["value"],
        "net_income_tag": net_income_metric["tag"],
        "assets": assets_metric["value"],
        "assets_uom": assets_metric["uom"],
        "liabilities": liabilities_metric["value"],
        "liabilities_uom": liabilities_metric["uom"],
        "equity": equity_metric["value"],
        "equity_uom": equity_metric["uom"],
        "cash_and_equivalents": cash_metric["value"],
        "cash_uom": cash_metric["uom"],
        "operating_cash_flow": ocf_metric["value"],
        "operating_cash_flow_uom": ocf_metric["uom"],
        "capex": capex_metric["value"],
        "capex_uom": capex_metric["uom"],
        "rd_expense": rd_metric["value"],
        "rd_expense_uom": rd_metric["uom"],
        "depreciation_amortization": da_metric["value"],
        "depreciation_amortization_uom": da_metric["uom"],
        "eps_basic": eps_basic_metric["value"],
        "eps_diluted": eps_diluted_metric["value"],
        "shares_outstanding": shares_outstanding_metric["value"],
        "shares_outstanding_uom": shares_outstanding_metric["uom"],
        "shares_issued": shares_issued_metric["value"],
        "shares_issued_uom": shares_issued_metric["uom"],
        "shares_authorized": shares_authorized_metric["value"],
        "shares_authorized_uom": shares_authorized_metric["uom"],
        "weighted_avg_diluted_shares": weighted_avg_diluted_metric["value"],
        "weighted_avg_diluted_shares_uom": weighted_avg_diluted_metric["uom"],
        "filer_category": clean_text((metadata_values.get("EntityFilerCategory") or [""])[0]) or None,
        "sec_file_number": clean_text((metadata_values.get("EntityFileNumber") or [""])[0]) or None,
        "tax_id": clean_text((metadata_values.get("EntityTaxIdentificationNumber") or [""])[0]) or None,
        "local_phone": clean_text((metadata_values.get("LocalPhoneNumber") or [""])[0]) or None,
    }


def supplement_from_am_annual(result_df: pd.DataFrame, stock_pool: pd.DataFrame) -> pd.DataFrame:
    if not AM_ANNUAL_DIR.exists():
        return result_df

    existing_codes = set(result_df["stock_code"].dropna().astype(str))
    supplement_rows: list[dict[str, Any]] = []

    for stock_code, folder_name in AM_SUPPLEMENT_TARGETS.items():
        if stock_code in existing_codes:
            continue
        folder = AM_ANNUAL_DIR / folder_name
        if not folder.exists():
            continue

        annual_files = sorted(
            [p for p in folder.glob("*.htm") if ("Form 10-K" in p.name or "Form 20-F" in p.name or "Form 40-F" in p.name)],
            key=lambda p: p.name,
            reverse=True,
        )
        if not annual_files:
            continue

        report_file = annual_files[0]
        parsed = extract_html_report(report_file)
        pool_row = stock_pool[stock_pool["SECU_CODE"] == stock_code]
        if pool_row.empty:
            continue
        row = pool_row.iloc[0]

        company_row = {col: None for col in ROW_COLUMNS}
        company_row.update(
            {
                "row_type": "company",
                "stock_code": stock_code,
                "company_name": row["SECUNAME"],
                "CUSIP": row["CUSIP"],
                "ticker": normalize_pool_ticker(stock_code),
                "ticker_match_method": "am_annual_html",
                "cik": None,
                "company_name_filing": parsed["company_name_filing"],
                "report_type": parsed["report_type"],
                "report_period": parsed["report_period"],
                "filed_date": None,
                "accepted_datetime": None,
                "filing_month_batch": "Am_annual",
                "adsh": None,
                "instance": report_file.name,
                "sic": None,
                "industry_l1": row.get("CCXDF_CN"),
                "industry_l2": row.get("CCXDS_CN"),
                "industry_l3": row.get("CCXDT_CN"),
                "industry_l4": row.get("CCXDFourth_CN"),
                "fiscal_year": None,
                "fiscal_period": None,
                "fiscal_year_end": None,
                "wksi": None,
                "filer_category": parsed["filer_category"],
                "sec_file_number": parsed["sec_file_number"],
                "tax_id": parsed["tax_id"],
                "local_phone": parsed["local_phone"],
                "public_float_usd": None,
                "float_date": None,
                "country_business_address": None,
                "state_business_address": None,
                "city_business_address": None,
                "country_incorporation": None,
                "state_incorporation": None,
                "main_business_desc": parsed["main_business_desc"],
                "main_business_tag": parsed["main_business_tag"],
                "segment_text_tag": parsed["segment_text_tag"],
                "revenue_note_tag": parsed["revenue_note_tag"],
                "total_revenue": parsed["total_revenue"],
                "total_revenue_tag": parsed["total_revenue_tag"],
                "total_revenue_version": None,
                "total_revenue_uom": parsed["total_revenue_uom"],
                "gross_profit": parsed["gross_profit"],
                "gross_profit_tag": parsed["gross_profit_tag"],
                "operating_income": parsed["operating_income"],
                "operating_income_tag": parsed["operating_income_tag"],
                "net_income": parsed["net_income"],
                "net_income_tag": parsed["net_income_tag"],
                "assets": parsed["assets"],
                "assets_uom": parsed["assets_uom"],
                "liabilities": parsed["liabilities"],
                "liabilities_uom": parsed["liabilities_uom"],
                "equity": parsed["equity"],
                "equity_uom": parsed["equity_uom"],
                "cash_and_equivalents": parsed["cash_and_equivalents"],
                "cash_uom": parsed["cash_uom"],
                "operating_cash_flow": parsed["operating_cash_flow"],
                "operating_cash_flow_uom": parsed["operating_cash_flow_uom"],
                "capex": parsed["capex"],
                "capex_uom": parsed["capex_uom"],
                "rd_expense": parsed["rd_expense"],
                "rd_expense_uom": parsed["rd_expense_uom"],
                "depreciation_amortization": parsed["depreciation_amortization"],
                "depreciation_amortization_uom": parsed["depreciation_amortization_uom"],
                "eps_basic": parsed["eps_basic"],
                "eps_diluted": parsed["eps_diluted"],
                "shares_outstanding": parsed["shares_outstanding"],
                "shares_outstanding_uom": parsed["shares_outstanding_uom"],
                "shares_issued": parsed["shares_issued"],
                "shares_issued_uom": parsed["shares_issued_uom"],
                "shares_authorized": parsed["shares_authorized"],
                "shares_authorized_uom": parsed["shares_authorized_uom"],
                "weighted_avg_diluted_shares": parsed["weighted_avg_diluted_shares"],
                "weighted_avg_diluted_shares_uom": parsed["weighted_avg_diluted_shares_uom"],
                "segment_disclosure_flag": False,
                "segment_count": 0,
                "primary_disclosure_dim": "narrative_only" if parsed["main_business_desc"] else None,
                "disclosure_dims_available": json.dumps([], ensure_ascii=False),
                "has_revenue_breakdown": False,
                "has_margin_breakdown": False,
                "segment_id": None,
                "segment_name": None,
                "segment_type": None,
                "segment_axis": None,
                "segment_member_raw": None,
                "segment_member_country_cn": None,
                "segment_revenue": None,
                "segment_revenue_uom": None,
                "segment_revenue_ratio": None,
                "segment_gross_profit": None,
                "segment_gross_profit_uom": None,
                "segment_gross_margin": None,
                "segment_desc": None,
                "source_section": "company",
                "source_tag": parsed["main_business_tag"] or parsed["segment_text_tag"] or parsed["revenue_note_tag"],
                "source_version": None,
                "source_group_tag": None,
                "source_context": None,
                "source_dim_hash": None,
                "source_segments_text": None,
                "source_report": None,
                "source_report_shortname": None,
                "source_report_longname": None,
                "evidence_text": parsed["evidence_text"],
                "extraction_status": "ok" if parsed["total_revenue"] is not None else "partial",
                "notes_file_batch": "Am_annual",
            }
        )
        supplement_rows.append(company_row)

    if not supplement_rows:
        return result_df

    supplement_df = pd.DataFrame(supplement_rows, columns=ROW_COLUMNS)
    result_df = pd.concat([result_df, supplement_df], ignore_index=True)
    return result_df


def write_extract_outputs(result_df: pd.DataFrame, csv_path: Path, xlsx_path: Path) -> None:
    csv_path.parent.mkdir(exist_ok=True)
    xlsx_path.parent.mkdir(exist_ok=True)
    result_df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        result_df.to_excel(writer, index=False, sheet_name="总表")
    add_description_rows(xlsx_path)


def export_single_company(
    *,
    notes_dir: Path,
    stock_code: str | None = None,
    company_name: str | None = None,
    adsh: str | None = None,
    output_xlsx: Path | None = None,
    output_csv: Path | None = None,
) -> dict[str, Any]:
    stock_pool = load_stock_pool()
    stock_row = find_single_company_stock_row(
        stock_pool,
        notes_dir,
        stock_code=stock_code,
        company_name=company_name,
        adsh=adsh,
    )
    rows = build_company_and_segment_rows(stock_row)
    result_df = pd.DataFrame(rows, columns=ROW_COLUMNS)
    if not result_df.empty:
        result_df = result_df.sort_values(
            ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue"],
            ascending=[True, False, True, True, False],
            na_position="last",
        ).reset_index(drop=True)

    identifier = clean_text(stock_code) or clean_text(company_name) or clean_text(adsh) or "single_company"
    safe_identifier = re.sub(r"[^A-Za-z0-9._-]+", "_", identifier).strip("_") or "single_company"
    default_stem = f"{safe_identifier}_{notes_dir.name}_extract"
    final_xlsx = output_xlsx or (OUTPUT_DIR / f"{default_stem}.xlsx")
    final_csv = output_csv or (OUTPUT_DIR / f"{default_stem}.csv")

    write_extract_outputs(result_df, final_csv, final_xlsx)

    summary = {
        "mode": "single_company",
        "notes_dir": str(notes_dir),
        "stock_code": result_df.iloc[0]["stock_code"] if not result_df.empty else stock_code,
        "company_name": result_df.iloc[0]["company_name"] if not result_df.empty else company_name,
        "adsh": stock_row.get("adsh"),
        "company_rows": int((result_df["row_type"] == "company").sum()) if not result_df.empty else 0,
        "segment_rows": int((result_df["row_type"] == "segment").sum()) if not result_df.empty else 0,
        "output_rows": int(len(result_df)),
        "csv": str(final_csv),
        "excel": str(final_xlsx),
    }
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract SEC notes coverage tables.")
    parser.add_argument("--notes-dir", help="单公司导出时使用的 notes 批次目录，如 data/2026_01_notes")
    parser.add_argument("--stock-code", help="单公司导出时的股票代码，如 NOC.N")
    parser.add_argument("--company-name", help="单公司导出时的公司名称，如 诺斯罗普-格鲁曼")
    parser.add_argument("--adsh", help="单公司导出时指定 filing 的 ADSH")
    parser.add_argument("--output-xlsx", help="输出 Excel 路径")
    parser.add_argument("--output-csv", help="输出 CSV 路径")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.notes_dir or args.stock_code or args.company_name or args.adsh:
        if not args.notes_dir:
            raise SystemExit("单公司导出模式需要提供 --notes-dir。")
        summary = export_single_company(
            notes_dir=Path(args.notes_dir),
            stock_code=args.stock_code,
            company_name=args.company_name,
            adsh=args.adsh,
            output_xlsx=Path(args.output_xlsx) if args.output_xlsx else None,
            output_csv=Path(args.output_csv) if args.output_csv else None,
        )
        print(json.dumps(summary, ensure_ascii=False, indent=2))
        return

    OUTPUT_DIR.mkdir(exist_ok=True)
    stock_pool = load_stock_pool()
    latest_filings = load_latest_filings(stock_pool)
    covered = latest_filings[latest_filings["matched"]].copy()

    all_rows: list[dict[str, Any]] = []
    validation_rows: list[dict[str, Any]] = []
    for _, stock_row in covered.iterrows():
        company_rows = build_company_and_segment_rows(stock_row)
        all_rows.extend(company_rows)
        validation_rows.append(validate_company_segments(stock_row, company_rows))

    result_df = pd.DataFrame(all_rows, columns=ROW_COLUMNS)
    result_df = supplement_from_am_annual(result_df, stock_pool)
    if not result_df.empty:
        result_df = result_df.sort_values(
            ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue"],
            ascending=[True, False, True, True, False],
            na_position="last",
        ).reset_index(drop=True)

    summary = {
        "notes_batches": [p.name for p in NOTES_DIRS],
        "pool_rows": int(len(stock_pool)),
        "covered_stocks": int(result_df[result_df["row_type"] == "company"]["stock_code"].nunique()) if not result_df.empty else 0,
        "coverage_pct": round((result_df[result_df["row_type"] == "company"]["stock_code"].nunique() / len(stock_pool) * 100), 2) if len(stock_pool) and not result_df.empty else 0,
        "company_rows": int((result_df["row_type"] == "company").sum()) if not result_df.empty else 0,
        "segment_rows": int((result_df["row_type"] == "segment").sum()) if not result_df.empty else 0,
        "output_rows": int(len(result_df)),
        "validation_companies": int(len(validation_rows)),
        "validation_with_missing_segment_types": int(sum(1 for row in validation_rows if row["missing_segment_types"])),
    }

    write_extract_outputs(result_df, OUTPUT_CSV, OUTPUT_XLSX)
    OUTPUT_SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    pd.DataFrame(validation_rows).to_excel(OUTPUT_DIR / "notes_coverage_validation.xlsx", index=False)

    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print(f"csv={OUTPUT_CSV}")
    print(f"excel={OUTPUT_XLSX}")


if __name__ == "__main__":
    main()


