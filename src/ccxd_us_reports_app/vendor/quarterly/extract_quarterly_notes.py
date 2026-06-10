from __future__ import annotations

"""Low-level quarterly extractor.

Reads a stock pool workbook plus SEC notes data under inputs/data and produces
raw quarterly company/segment rows. This script is mainly used as the engine
behind run_folder_quarterly_pipeline.py and is only meant to be called
directly for debugging or small-scope extraction.
"""

import argparse
import json
import re
import urllib.request
from pathlib import Path
from typing import Any

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
INPUTS_DIR = PROJECT_ROOT / "inputs"
DATA_DIR = INPUTS_DIR / "data"
DEFAULT_WORKBOOK = INPUTS_DIR / "美股数据提取筛选.xlsx"
OUTPUT_DIR = SCRIPT_DIR / "outputs"
DEFAULT_RAW_CSV = OUTPUT_DIR / "quarterly_raw.csv"
DEFAULT_RAW_XLSX = OUTPUT_DIR / "quarterly_raw.xlsx"
DEFAULT_SUMMARY = OUTPUT_DIR / "quarterly_extract_summary.json"
SEC_TICKER_CACHE = OUTPUT_DIR / "sec_company_tickers_cache.json"


def first_existing(*candidates: Path) -> Path:
    for path in candidates:
        if path.exists():
            return path
    return candidates[0]


ANNUAL_PROJECT_DIR = first_existing(PROJECT_ROOT / "us_project", PROJECT_ROOT / "us_project_year")
INPUTS_DIR = first_existing(ANNUAL_PROJECT_DIR / "inputs", INPUTS_DIR)
DATA_DIR = first_existing(INPUTS_DIR / "data", INPUTS_DIR / "\u7f8e\u80a1data", DATA_DIR)
DEFAULT_WORKBOOK = INPUTS_DIR / "\u7f8e\u80a1\u6570\u636e\u63d0\u53d6\u7b5b\u9009.xlsx"

QUARTERLY_FORMS = {"10-Q", "10-Q/A", "6-K", "6-K/A"}
QUARTERLY_PERIODS = {"Q1", "Q2", "Q3"}
INFERABLE_6K_FORMS = {"6-K", "6-K/A"}
SEC_TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
USER_AGENT = "quarterly-notes-extract contact@example.com"

REVENUE_TAG_PRIORITY = [
    "RevenueFromExternalCustomers",
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "RevenueFromRenderingOfTelecommunicationServices",
    "RevenueFromContractsWithCustomers",
    "Revenues",
    "RevenuesAndOther",
    "RegulatedAndUnregulatedOperatingRevenue",
    "SalesRevenueNet",
    "Revenue",
    "NetSales",
    "RevenueFromContractWithCustomerIncludingAssessedTax",
    "TotalRevenue",
    "RevenueFromSaleOfGoods",
    "RevenueFromRenderingOfServices",
]

PERIOD_INFERENCE_TAGS = set(REVENUE_TAG_PRIORITY) | {
    "GrossProfit",
    "GrossProfitLoss",
    "OperatingIncomeLoss",
    "OperatingExpenses",
    "NetIncomeLoss",
    "ProfitLoss",
    "ResearchAndDevelopmentExpense",
}

TEXT_TAG_PRIORITIES = {
    "main_business_tag": [
        "BusinessDescriptionPolicyTextBlock",
        "BusinessDescriptionAndAccountingPoliciesTextBlock",
        "BusinessDescriptionAndBasisOfPresentationTextBlock",
        "BusinessOverviewPolicyTextBlock",
        "DescriptionOfBusinessTextBlock",
        "BusinessDescriptionTextBlock",
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
    "fiscal_year",
    "fiscal_period",
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
    "total_revenue_qtd",
    "total_revenue_ytd",
    "total_revenue_tag",
    "total_revenue_version",
    "total_revenue_uom",
    "gross_profit_qtd",
    "gross_profit_ytd",
    "operating_income_qtd",
    "operating_income_ytd",
    "net_income_qtd",
    "net_income_ytd",
    "rd_expense_qtd",
    "rd_expense_ytd",
    "operating_cash_flow_ytd",
    "capex_ytd",
    "eps_basic_qtd",
    "eps_basic_ytd",
    "eps_diluted_qtd",
    "eps_diluted_ytd",
    "assets",
    "assets_uom",
    "liabilities",
    "liabilities_uom",
    "equity",
    "equity_uom",
    "cash_and_equivalents",
    "cash_uom",
    "segment_disclosure_flag",
    "segment_count",
    "primary_disclosure_dim",
    "disclosure_dims_available",
    "has_revenue_breakdown",
    "segment_id",
    "segment_name",
    "segment_type",
    "segment_axis",
    "segment_member_raw",
    "segment_revenue_qtd",
    "segment_revenue_ytd",
    "segment_revenue_uom",
    "segment_revenue_ratio_qtd",
    "segment_revenue_ratio_ytd",
    "source_tag",
    "source_version",
    "source_dim_hash",
    "source_segments_text",
    "source_report",
    "source_report_shortname",
    "source_report_longname",
    "evidence_text",
    "extraction_status",
    "notes_file_batch",
]


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_pool_ticker(code: Any) -> str:
    text = clean_text(code).upper().replace("_", "-")
    if "." in text:
        base, suffix = text.rsplit(".", 1)
        if len(suffix) <= 2 and suffix.isalpha():
            return base
    return text


def normalize_cik(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        text = str(int(value))
    else:
        text = clean_text(value)
        if text.endswith(".0") and text[:-2].isdigit():
            text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits.zfill(10) if digits else ""


def normalize_report_date(value: Any) -> str | None:
    text = clean_text(value)
    if not text:
        return None
    parsed = pd.to_datetime(text, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.strftime("%Y-%m-%d")


def camel_to_words(value: Any) -> str:
    text = clean_text(value)
    text = re.sub(r"(Member|Axis|Domain)$", "", text)
    text = re.sub(r"([a-z0-9])([A-Z])", r"\1 \2", text)
    text = re.sub(r"([A-Z]+)([A-Z][a-z])", r"\1 \2", text)
    return clean_text(text.replace("_", " "))


def dimension_category(axis_name: Any) -> str | None:
    dim = clean_text(axis_name).lower()
    if any(token in dim for token in ["businesssegments", "operatingsegments", "reportablesegments"]):
        return "business_segment"
    if dim in {"segments", "segment"}:
        return "business_segment"
    if dim.endswith("segments") and not any(token in dim for token in ["portfolio", "loan", "receivable", "share", "contract"]):
        return "business_segment"
    if any(token in dim for token in ["productorservice", "productsandservices"]):
        return "product"
    if dim in {"product", "products", "service", "services"}:
        return "product"
    if any(token in dim for token in ["geograph", "region", "country", "area"]):
        return "region"
    if any(token in dim for token in ["endmarket", "marketsofcustomers"]):
        return "end_market"
    if any(token in dim for token in ["customer", "client"]):
        return "customer_type"
    if any(token in dim for token in ["saleschannel", "channel"]):
        return "sales_channel"
    if "industry" in dim:
        return "industry"
    if any(token in dim for token in ["contracttype", "typeofcontract", "arrangement", "program", "project"]):
        return "contract_or_program"
    if any(token in dim for token in ["reconciliation", "reconciling", "consolidationitems"]):
        return "reconciling_item"
    return None


def fiscal_period_to_ytd_qtrs(fp: Any) -> int | None:
    mapping = {"Q1": 1, "Q2": 2, "Q3": 3}
    return mapping.get(clean_text(fp).upper())


def load_stock_pool(workbook: Path, limit: int | None, stock_codes: list[str] | None) -> pd.DataFrame:
    pool = pd.read_excel(workbook, sheet_name=0).copy()
    industries = pd.read_excel(workbook, sheet_name=1).copy()

    required = {"CUSIP", "SECU_CODE", "SECUNAME"}
    missing = required - set(pool.columns)
    if missing:
        raise ValueError(f"股票池缺少字段: {', '.join(sorted(missing))}")

    pool["ticker"] = pool["SECU_CODE"].map(normalize_pool_ticker)
    if stock_codes:
        wanted = {clean_text(code) for code in stock_codes if clean_text(code)}
        pool = pool[pool["SECU_CODE"].map(clean_text).isin(wanted)].copy()

    industry_cols = ["CCXDF_CN", "CCXDS_CN", "CCXDT_CN", "CCXDFourth_CN"]
    if set(["SECU_CODE", *industry_cols]).issubset(industries.columns):
        by_code = industries[["SECU_CODE", *industry_cols]].drop_duplicates("SECU_CODE")
        pool = pool.merge(by_code, on="SECU_CODE", how="left")
    if set(["CUSIP", *industry_cols]).issubset(industries.columns):
        by_cusip = industries[["CUSIP", *industry_cols]].drop_duplicates("CUSIP")
        fallback = pool[["CUSIP"]].merge(by_cusip, on="CUSIP", how="left")
        for col in industry_cols:
            if col not in pool.columns:
                pool[col] = None
            pool[col] = pool[col].fillna(fallback[col])

    if limit:
        pool = pool.head(limit).copy()
    return pool.reset_index(drop=True)


def load_sec_ticker_map(cache_path: Path, refresh: bool = False) -> pd.DataFrame:
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    if cache_path.exists() and not refresh:
        data = json.loads(cache_path.read_text(encoding="utf-8"))
    else:
        req = urllib.request.Request(SEC_TICKER_URL, headers={"User-Agent": USER_AGENT})
        with urllib.request.urlopen(req, timeout=60) as response:
            data = json.loads(response.read().decode("utf-8"))
        cache_path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    df = pd.DataFrame.from_dict(data, orient="index")
    df["ticker"] = df["ticker"].astype(str).str.upper().str.strip()
    df["cik10"] = df["cik_str"].astype(int).astype(str).str.zfill(10)
    return df[["ticker", "cik10", "title"]].drop_duplicates("ticker")


def list_notes_dirs(data_dir: Path) -> list[Path]:
    return sorted([p for p in data_dir.iterdir() if p.is_dir() and p.name.endswith("_notes")])


def infer_6k_fiscal_periods_from_num(notes_dir: Path, candidates: pd.DataFrame) -> dict[str, str]:
    if candidates.empty:
        return {}
    num_path = notes_dir / "num.tsv"
    if not num_path.exists():
        return {}

    period_by_adsh = {
        clean_text(row.get("adsh")): clean_text(row.get("period"))
        for _, row in candidates[["adsh", "period"]].dropna(subset=["adsh", "period"]).iterrows()
        if clean_text(row.get("adsh")) and clean_text(row.get("period"))
    }
    if not period_by_adsh:
        return {}

    qtrs_seen: dict[str, set[int]] = {adsh: set() for adsh in period_by_adsh}
    for chunk in pd.read_csv(
        num_path,
        sep="\t",
        dtype=str,
        usecols=["adsh", "tag", "ddate", "qtrs", "value"],
        chunksize=300_000,
    ):
        keep = chunk[chunk["adsh"].isin(period_by_adsh)].copy()
        if keep.empty:
            continue
        keep = keep[keep["tag"].isin(PERIOD_INFERENCE_TAGS)]
        if keep.empty:
            continue
        keep["period"] = keep["adsh"].map(period_by_adsh)
        keep = keep[keep["ddate"].map(clean_text) == keep["period"]]
        if keep.empty:
            continue
        keep["qtrs_num"] = pd.to_numeric(keep["qtrs"], errors="coerce")
        keep["value_num"] = pd.to_numeric(keep["value"], errors="coerce")
        keep = keep[keep["qtrs_num"].notna() & keep["value_num"].notna()]
        for adsh, group in keep.groupby("adsh"):
            qtrs_seen.setdefault(adsh, set()).update(int(value) for value in group["qtrs_num"].dropna().tolist())

    inferred: dict[str, str] = {}
    for adsh, values in qtrs_seen.items():
        if any(value >= 4 for value in values):
            continue
        valid = {value for value in values if value in {1, 2, 3}}
        if 1 not in valid:
            continue
        if 3 in valid:
            inferred[adsh] = "Q3"
        elif 2 in valid:
            inferred[adsh] = "Q2"
        else:
            inferred[adsh] = "Q1"
    return inferred


def collect_latest_quarterly_filings(stock_pool: pd.DataFrame, data_dir: Path, refresh_sec_map: bool) -> pd.DataFrame:
    sec_map = load_sec_ticker_map(SEC_TICKER_CACHE, refresh=refresh_sec_map)
    stock_pool = stock_pool.copy()
    provided_cik = pd.Series([""] * len(stock_pool), index=stock_pool.index)
    for col in ["cik10", "cik", "CIK", "html_cik10"]:
        if col in stock_pool.columns:
            normalized = stock_pool[col].map(normalize_cik)
            provided_cik = provided_cik.mask(provided_cik.eq("") & normalized.ne(""), normalized)
    merge_pool = stock_pool.drop(columns=[col for col in ["cik10", "cik10_provided"] if col in stock_pool.columns])
    merge_pool["cik10_provided"] = provided_cik
    pool = merge_pool.merge(sec_map.rename(columns={"cik10": "cik10_sec"}), on="ticker", how="left")
    pool["cik10"] = pool["cik10_sec"].map(clean_text)
    pool["cik10"] = pool["cik10"].mask(pool["cik10"].eq(""), pool["cik10_provided"])
    pool["ticker_match_method"] = pool.apply(
        lambda row: "sec_ticker_file" if clean_text(row.get("cik10_sec")) else ("local_html_cik" if clean_text(row.get("cik10_provided")) else ""),
        axis=1,
    )
    target_ciks = {clean_text(value) for value in pool["cik10"].dropna().astype(str) if clean_text(value)}
    if not target_ciks:
        return pool.assign(matched=False)

    frames: list[pd.DataFrame] = []
    for notes_dir in list_notes_dirs(data_dir):
        sub_path = notes_dir / "sub.tsv"
        if not sub_path.exists():
            continue
        sub = pd.read_csv(
            sub_path,
            sep="\t",
            dtype=str,
            usecols=[
                "adsh",
                "cik",
                "name",
                "sic",
                "form",
                "period",
                "fy",
                "fp",
                "filed",
                "accepted",
                "detail",
                "instance",
                "countryba",
                "stprba",
                "cityba",
                "countryinc",
                "stprinc",
            ],
        )
        sub["cik10"] = sub["cik"].astype(str).str.zfill(10)
        detail_mask = sub["detail"].fillna("").isin(["1", "true", "TRUE"])
        target_mask = sub["cik10"].isin(target_ciks)
        fp_upper = sub["fp"].fillna("").str.upper()
        direct = sub[
            sub["form"].isin(QUARTERLY_FORMS)
            & fp_upper.isin(QUARTERLY_PERIODS)
            & detail_mask
            & target_mask
        ].copy()
        direct["fp_inferred_from"] = ""

        inferred_candidates = sub[
            sub["form"].isin(INFERABLE_6K_FORMS)
            & ~fp_upper.isin(QUARTERLY_PERIODS | {"FY"})
            & detail_mask
            & target_mask
        ].copy()
        inferred_map = infer_6k_fiscal_periods_from_num(notes_dir, inferred_candidates)
        inferred = inferred_candidates[inferred_candidates["adsh"].isin(inferred_map)].copy()
        if not inferred.empty:
            inferred["fp"] = inferred["adsh"].map(inferred_map)
            inferred["fp_inferred_from"] = "num_qtrs"

        selected = pd.concat([direct, inferred], ignore_index=True)
        if selected.empty:
            continue
        selected["batch"] = notes_dir.name
        selected["notes_dir"] = str(notes_dir)
        selected["filed_num"] = pd.to_numeric(selected["filed"], errors="coerce")
        selected["period_num"] = pd.to_numeric(selected["period"], errors="coerce")
        frames.append(selected)

    if not frames:
        return pool.assign(matched=False)

    filings = pd.concat(frames, ignore_index=True)
    filings = filings.sort_values(
        ["cik10", "filed_num", "period_num", "adsh"],
        ascending=[True, False, False, False],
        kind="stable",
    )
    latest = filings.drop_duplicates("cik10", keep="first")
    matched = pool.merge(latest, on="cik10", how="left", suffixes=("_pool", "_filing"))
    matched["matched"] = matched["adsh"].notna()
    return matched


def filter_txt_rows(txt_path: Path, adshs: set[str]) -> pd.DataFrame:
    chunks: list[pd.DataFrame] = []
    for chunk in pd.read_csv(txt_path, sep="\t", dtype=str, usecols=["adsh", "tag", "value"], chunksize=200_000):
        keep = chunk[chunk["adsh"].isin(adshs)]
        if not keep.empty:
            chunks.append(keep)
    if not chunks:
        return pd.DataFrame(columns=["adsh", "tag", "value"])
    return pd.concat(chunks, ignore_index=True)


def load_notes_subset(notes_dir: Path, adshs: set[str]) -> dict[str, pd.DataFrame]:
    num = pd.read_csv(
        notes_dir / "num.tsv",
        sep="\t",
        dtype=str,
        usecols=["adsh", "tag", "version", "ddate", "qtrs", "uom", "dimh", "value"],
    )
    num = num[num["adsh"].isin(adshs)].copy()
    num["value_num"] = pd.to_numeric(num["value"], errors="coerce")
    num["qtrs_num"] = pd.to_numeric(num["qtrs"], errors="coerce")
    num["dimh"] = num["dimh"].fillna("0x00000000")

    pre = pd.read_csv(
        notes_dir / "pre.tsv",
        sep="\t",
        dtype=str,
        usecols=["adsh", "report", "line", "stmt", "tag", "version", "plabel"],
    )
    pre = pre[pre["adsh"].isin(adshs)].copy()

    ren = pd.read_csv(notes_dir / "ren.tsv", sep="\t", dtype=str, usecols=["adsh", "report", "shortname", "longname"])
    ren = ren[ren["adsh"].isin(adshs)].copy()
    ren["report"] = pd.to_numeric(ren["report"], errors="coerce")

    dim = pd.read_csv(notes_dir / "dim.tsv", sep="\t", dtype=str)
    txt = filter_txt_rows(notes_dir / "txt.tsv", adshs)
    return {"num": num, "pre": pre, "ren": ren, "dim": dim, "txt": txt}


def metric_candidates(num_df: pd.DataFrame, report_period: str, tags: list[str], qtrs: int) -> pd.DataFrame:
    candidates = num_df[
        num_df["tag"].isin(tags)
        & (num_df["dimh"] == "0x00000000")
        & (num_df["ddate"] == report_period)
        & (num_df["qtrs_num"] == qtrs)
        & (num_df["value_num"].notna())
    ].copy()
    if candidates.empty:
        return candidates
    candidates["priority_rank"] = candidates["tag"].map(lambda tag: tags.index(tag) if tag in tags else len(tags))
    return candidates.sort_values(["priority_rank", "value_num"], ascending=[True, False], kind="stable")


def choose_metric(num_df: pd.DataFrame, report_period: str, tags: list[str], qtrs: int) -> dict[str, Any]:
    candidates = metric_candidates(num_df, report_period, tags, qtrs)
    if candidates.empty:
        return {"value": None, "tag": None, "version": None, "uom": None}
    row = candidates.iloc[0]
    return {"value": row["value_num"], "tag": row["tag"], "version": row["version"], "uom": row["uom"]}


def choose_instant_metric(num_df: pd.DataFrame, report_period: str, tags: list[str]) -> dict[str, Any]:
    return choose_metric(num_df, report_period, tags, 0)


def choose_text_value(txt_df: pd.DataFrame, priorities: list[str]) -> tuple[str | None, str]:
    for tag in priorities:
        rows = txt_df[txt_df["tag"] == tag]
        if rows.empty:
            continue
        values = rows["value"].map(clean_text)
        values = values[values != ""]
        if not values.empty:
            return tag, max(values, key=len)
    return None, ""


def segment_fact_candidates(num_df: pd.DataFrame, report_period: str, qtrs: int) -> pd.DataFrame:
    candidates = num_df[
        num_df["tag"].isin(REVENUE_TAG_PRIORITY)
        & (num_df["dimh"] != "0x00000000")
        & (num_df["ddate"] == report_period)
        & (num_df["qtrs_num"] == qtrs)
        & (num_df["value_num"].notna())
    ].copy()
    if candidates.empty:
        return candidates
    candidates["tag_rank"] = candidates["tag"].map(lambda tag: REVENUE_TAG_PRIORITY.index(tag) if tag in REVENUE_TAG_PRIORITY else 99)
    return candidates.sort_values(["tag_rank", "value_num"], ascending=[True, False], kind="stable")


def extract_snippet(text: str, term: str) -> str:
    text = clean_text(text)
    term = clean_text(term)
    if not text or not term:
        return ""
    lower = text.lower()
    pos = lower.find(term.lower())
    if pos < 0:
        return text[:500]
    start = max(0, pos - 180)
    end = min(len(text), pos + len(term) + 220)
    return text[start:end]


def build_segment_rows(
    *,
    adsh: str,
    report_period: str,
    fiscal_period: str,
    total_revenue_qtd: float | None,
    total_revenue_ytd: float | None,
    revenue_uom: str | None,
    num_df: pd.DataFrame,
    dim_df: pd.DataFrame,
    pre_df: pd.DataFrame,
    ren_df: pd.DataFrame,
    evidence_text: str,
) -> tuple[list[dict[str, Any]], list[str], str | None]:
    ytd_qtrs = fiscal_period_to_ytd_qtrs(fiscal_period)
    if ytd_qtrs is None:
        return [], [], None

    dim_map = dim_df.set_index("dimhash").to_dict("index") if "dimhash" in dim_df.columns else {}
    ren_keyed = {(r["adsh"], int(r["report"])): r for r in ren_df.dropna(subset=["report"]).to_dict("records")}

    rows_by_key: dict[tuple[str, str, str, str, str], dict[str, Any]] = {}
    dims_available: set[str] = set()

    for period_kind, qtrs in (("qtd", 1), ("ytd", ytd_qtrs)):
        facts = segment_fact_candidates(num_df, report_period, qtrs)
        for _, fact in facts.iterrows():
            dim_info = dim_map.get(clean_text(fact["dimh"]))
            if not dim_info:
                continue
            segments_text = clean_text(dim_info.get("segments"))
            if not segments_text or "=" not in segments_text:
                continue

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

            for item in segments_text.split(";"):
                if "=" not in item:
                    continue
                axis_name, member_name = item.split("=", 1)
                axis_name = clean_text(axis_name)
                member_name = clean_text(member_name)
                seg_type = dimension_category(axis_name) or "other_dimension"
                member_label = camel_to_words(member_name)
                key = (seg_type, axis_name, member_name, fact["tag"], clean_text(fact["dimh"]))
                row = rows_by_key.setdefault(
                    key,
                    {
                        "segment_name": member_label,
                        "segment_type": seg_type,
                        "segment_axis": axis_name,
                        "segment_member_raw": member_name,
                        "segment_revenue_qtd": None,
                        "segment_revenue_ytd": None,
                        "segment_revenue_uom": fact["uom"] or revenue_uom,
                        "segment_revenue_ratio_qtd": None,
                        "segment_revenue_ratio_ytd": None,
                        "source_tag": fact["tag"],
                        "source_version": fact["version"],
                        "source_dim_hash": clean_text(fact["dimh"]),
                        "source_segments_text": segments_text,
                        "source_report": source_report,
                        "source_report_shortname": source_short,
                        "source_report_longname": source_long,
                        "evidence_text": extract_snippet(evidence_text, member_label) or evidence_text[:500],
                    },
                )
                value = fact["value_num"]
                row[f"segment_revenue_{period_kind}"] = value
                if period_kind == "qtd" and total_revenue_qtd:
                    row["segment_revenue_ratio_qtd"] = value / total_revenue_qtd
                if period_kind == "ytd" and total_revenue_ytd:
                    row["segment_revenue_ratio_ytd"] = value / total_revenue_ytd
                dims_available.add(seg_type)

    rows = list(rows_by_key.values())
    if not rows:
        return [], [], None

    counts: dict[str, int] = {}
    for row in rows:
        counts[row["segment_type"]] = counts.get(row["segment_type"], 0) + 1
    primary_dim = sorted(counts.items(), key=lambda item: (-item[1], item[0]))[0][0]
    rows = sorted(rows, key=lambda row: (row["segment_type"], -(row["segment_revenue_qtd"] or row["segment_revenue_ytd"] or 0), row["segment_name"]))
    return rows, sorted(dims_available), primary_dim


def build_rows_for_filing(stock_row: pd.Series, cache: dict[str, pd.DataFrame]) -> list[dict[str, Any]]:
    adsh = clean_text(stock_row.get("adsh"))
    report_period_raw = clean_text(stock_row.get("period"))
    fiscal_period = clean_text(stock_row.get("fp")).upper()
    ytd_qtrs = fiscal_period_to_ytd_qtrs(fiscal_period)
    if not adsh or not report_period_raw or ytd_qtrs is None:
        return []

    txt = cache["txt"][cache["txt"]["adsh"] == adsh].copy()
    num = cache["num"][cache["num"]["adsh"] == adsh].copy()
    pre = cache["pre"][cache["pre"]["adsh"] == adsh].copy()
    ren = cache["ren"][cache["ren"]["adsh"] == adsh].copy()
    dim = cache["dim"]

    text_results = {name: choose_text_value(txt, priorities) for name, priorities in TEXT_TAG_PRIORITIES.items()}
    main_business_tag, main_business_desc = text_results["main_business_tag"]
    segment_text_tag, segment_text = text_results["segment_text_tag"]
    revenue_note_tag, revenue_note_text = text_results["revenue_note_tag"]
    if not main_business_desc:
        main_business_desc = segment_text or revenue_note_text
        main_business_tag = main_business_tag or segment_text_tag or revenue_note_tag

    revenue_qtd = choose_metric(num, report_period_raw, REVENUE_TAG_PRIORITY, 1)
    revenue_ytd = choose_metric(num, report_period_raw, REVENUE_TAG_PRIORITY, ytd_qtrs)
    if ytd_qtrs == 1 and revenue_ytd["value"] is None:
        revenue_ytd = revenue_qtd
    if revenue_qtd["value"] is None and ytd_qtrs == 1:
        revenue_qtd = revenue_ytd
    revenue_uom = revenue_qtd["uom"] or revenue_ytd["uom"]

    gross_qtd = choose_metric(num, report_period_raw, ["GrossProfit", "GrossProfitLoss"], 1)
    gross_ytd = choose_metric(num, report_period_raw, ["GrossProfit", "GrossProfitLoss"], ytd_qtrs)
    op_qtd = choose_metric(num, report_period_raw, ["OperatingIncomeLoss"], 1)
    op_ytd = choose_metric(num, report_period_raw, ["OperatingIncomeLoss"], ytd_qtrs)
    net_qtd = choose_metric(num, report_period_raw, ["NetIncomeLoss", "ProfitLoss"], 1)
    net_ytd = choose_metric(num, report_period_raw, ["NetIncomeLoss", "ProfitLoss"], ytd_qtrs)
    rd_qtd = choose_metric(num, report_period_raw, ["ResearchAndDevelopmentExpense"], 1)
    rd_ytd = choose_metric(num, report_period_raw, ["ResearchAndDevelopmentExpense"], ytd_qtrs)
    ocf_ytd = choose_metric(num, report_period_raw, ["NetCashProvidedByUsedInOperatingActivities"], ytd_qtrs)
    capex_ytd = choose_metric(num, report_period_raw, ["PaymentsToAcquirePropertyPlantAndEquipment"], ytd_qtrs)
    eps_basic_qtd = choose_metric(num, report_period_raw, ["EarningsPerShareBasic"], 1)
    eps_basic_ytd = choose_metric(num, report_period_raw, ["EarningsPerShareBasic"], ytd_qtrs)
    eps_diluted_qtd = choose_metric(num, report_period_raw, ["EarningsPerShareDiluted"], 1)
    eps_diluted_ytd = choose_metric(num, report_period_raw, ["EarningsPerShareDiluted"], ytd_qtrs)

    assets = choose_instant_metric(num, report_period_raw, ["Assets"])
    liabilities = choose_instant_metric(num, report_period_raw, ["Liabilities"])
    equity = choose_instant_metric(
        num,
        report_period_raw,
        ["StockholdersEquity", "StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest"],
    )
    cash = choose_instant_metric(num, report_period_raw, ["CashAndCashEquivalentsAtCarryingValue"])

    evidence_text = segment_text or revenue_note_text or main_business_desc
    segment_rows, dims_available, primary_dim = build_segment_rows(
        adsh=adsh,
        report_period=report_period_raw,
        fiscal_period=fiscal_period,
        total_revenue_qtd=revenue_qtd["value"],
        total_revenue_ytd=revenue_ytd["value"],
        revenue_uom=revenue_uom,
        num_df=num,
        dim_df=dim,
        pre_df=pre,
        ren_df=ren,
        evidence_text=evidence_text,
    )

    disclosure_dims = dims_available or (["narrative_only"] if segment_text or revenue_note_text else [])
    company_row = {
        "row_type": "company",
        "stock_code": stock_row.get("SECU_CODE"),
        "company_name": stock_row.get("SECUNAME"),
        "CUSIP": stock_row.get("CUSIP"),
        "ticker": stock_row.get("ticker"),
        "ticker_match_method": stock_row.get("ticker_match_method"),
        "cik": clean_text(stock_row.get("cik")),
        "company_name_filing": clean_text(stock_row.get("name")),
        "report_type": clean_text(stock_row.get("form")),
        "report_period": normalize_report_date(report_period_raw),
        "fiscal_year": clean_text(stock_row.get("fy")),
        "fiscal_period": fiscal_period,
        "filed_date": normalize_report_date(stock_row.get("filed")),
        "accepted_datetime": clean_text(stock_row.get("accepted")),
        "filing_month_batch": clean_text(stock_row.get("batch")),
        "adsh": adsh,
        "instance": clean_text(stock_row.get("instance")),
        "sic": clean_text(stock_row.get("sic")),
        "industry_l1": stock_row.get("CCXDF_CN"),
        "industry_l2": stock_row.get("CCXDS_CN"),
        "industry_l3": stock_row.get("CCXDT_CN"),
        "industry_l4": stock_row.get("CCXDFourth_CN"),
        "total_revenue_qtd": revenue_qtd["value"],
        "total_revenue_ytd": revenue_ytd["value"],
        "total_revenue_tag": revenue_qtd["tag"] or revenue_ytd["tag"],
        "total_revenue_version": revenue_qtd["version"] or revenue_ytd["version"],
        "total_revenue_uom": revenue_uom,
        "gross_profit_qtd": gross_qtd["value"],
        "gross_profit_ytd": gross_ytd["value"],
        "operating_income_qtd": op_qtd["value"],
        "operating_income_ytd": op_ytd["value"],
        "net_income_qtd": net_qtd["value"],
        "net_income_ytd": net_ytd["value"],
        "rd_expense_qtd": rd_qtd["value"],
        "rd_expense_ytd": rd_ytd["value"],
        "operating_cash_flow_ytd": ocf_ytd["value"],
        "capex_ytd": capex_ytd["value"],
        "eps_basic_qtd": eps_basic_qtd["value"],
        "eps_basic_ytd": eps_basic_ytd["value"],
        "eps_diluted_qtd": eps_diluted_qtd["value"],
        "eps_diluted_ytd": eps_diluted_ytd["value"],
        "assets": assets["value"],
        "assets_uom": assets["uom"],
        "liabilities": liabilities["value"],
        "liabilities_uom": liabilities["uom"],
        "equity": equity["value"],
        "equity_uom": equity["uom"],
        "cash_and_equivalents": cash["value"],
        "cash_uom": cash["uom"],
        "segment_disclosure_flag": bool(segment_text or revenue_note_text or segment_rows),
        "segment_count": sum(1 for row in segment_rows if row["segment_type"] == primary_dim) if primary_dim else len(segment_rows),
        "primary_disclosure_dim": primary_dim or ("narrative_only" if segment_text or revenue_note_text else None),
        "disclosure_dims_available": json.dumps(disclosure_dims, ensure_ascii=False),
        "has_revenue_breakdown": bool(segment_rows),
        "segment_id": None,
        "segment_name": None,
        "segment_type": None,
        "segment_axis": None,
        "segment_member_raw": None,
        "segment_revenue_qtd": None,
        "segment_revenue_ytd": None,
        "segment_revenue_uom": None,
        "segment_revenue_ratio_qtd": None,
        "segment_revenue_ratio_ytd": None,
        "source_tag": main_business_tag or segment_text_tag or revenue_note_tag,
        "source_version": None,
        "source_dim_hash": None,
        "source_segments_text": None,
        "source_report": None,
        "source_report_shortname": None,
        "source_report_longname": None,
        "evidence_text": evidence_text[:6000],
        "extraction_status": "ok" if revenue_qtd["value"] is not None or revenue_ytd["value"] is not None else "partial",
        "notes_file_batch": clean_text(stock_row.get("batch")),
    }

    rows = [{col: company_row.get(col) for col in ROW_COLUMNS}]
    for idx, segment in enumerate(segment_rows, start=1):
        row = dict(company_row)
        row.update(segment)
        row["row_type"] = "segment"
        row["segment_id"] = f"{company_row['stock_code']}|{company_row['report_period'] or 'NA'}|{company_row['fiscal_period']}|{idx}"
        rows.append({col: row.get(col) for col in ROW_COLUMNS})
    return rows


def extract_quarterly(args: argparse.Namespace) -> pd.DataFrame:
    stock_codes = [item.strip() for item in args.stock_codes.split(",")] if args.stock_codes else None
    stock_pool = load_stock_pool(Path(args.workbook), args.limit, stock_codes)
    matched = collect_latest_quarterly_filings(stock_pool, Path(args.data_dir), args.refresh_sec_map)
    covered = matched[matched["matched"]].copy()
    if covered.empty:
        return pd.DataFrame(columns=ROW_COLUMNS)

    all_rows: list[dict[str, Any]] = []
    for notes_dir_text, group in covered.groupby("notes_dir", sort=True):
        notes_dir = Path(notes_dir_text)
        adshs = set(group["adsh"].map(clean_text))
        cache = load_notes_subset(notes_dir, adshs)
        for _, row in group.iterrows():
            all_rows.extend(build_rows_for_filing(row, cache))

    result = pd.DataFrame(all_rows, columns=ROW_COLUMNS)
    if not result.empty:
        result = result.sort_values(
            ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue_qtd", "segment_revenue_ytd"],
            ascending=[True, False, True, True, False, False],
            na_position="last",
            kind="stable",
        ).reset_index(drop=True)
    return result


def write_outputs(result: pd.DataFrame, csv_path: Path, xlsx_path: Path, summary_path: Path, stock_pool_size: int) -> None:
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    result.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="raw_quarterly")

    company = result[result["row_type"] == "company"] if not result.empty else pd.DataFrame(columns=result.columns)
    summary = {
        "stock_pool_requested": stock_pool_size,
        "company_rows": int(len(company)),
        "segment_rows": int((result["row_type"] == "segment").sum()) if not result.empty else 0,
        "output_rows": int(len(result)),
        "forms": sorted(company["report_type"].dropna().unique().tolist()) if not company.empty else [],
        "fiscal_periods": sorted(company["fiscal_period"].dropna().unique().tolist()) if not company.empty else [],
        "batches": sorted(company["notes_file_batch"].dropna().unique().tolist()) if not company.empty else [],
        "csv": str(csv_path),
        "excel": str(xlsx_path),
    }
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract latest 10-Q/10-Q/A and quarterly 6-K SEC notes data for US stocks.")
    parser.add_argument("--workbook", default=str(DEFAULT_WORKBOOK), help="股票池 Excel，默认 inputs/美股数据提取筛选.xlsx")
    parser.add_argument("--data-dir", default=str(DATA_DIR), help="SEC notes data 目录")
    parser.add_argument("--limit", type=int, default=None, help="只处理股票池前 N 只股票，测试时建议设置")
    parser.add_argument("--stock-codes", default="", help="逗号分隔的股票代码，如 NVDA.O,AAPL.O,MSFT.O")
    parser.add_argument("--output-csv", default=str(DEFAULT_RAW_CSV), help="原始结果 CSV 路径")
    parser.add_argument("--output-xlsx", default=str(DEFAULT_RAW_XLSX), help="原始结果 Excel 路径")
    parser.add_argument("--summary-json", default=str(DEFAULT_SUMMARY), help="提取摘要 JSON 路径")
    parser.add_argument("--refresh-sec-map", action="store_true", help="重新下载 SEC ticker-CIK 映射")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    stock_codes = [item.strip() for item in args.stock_codes.split(",") if item.strip()] if args.stock_codes else None
    stock_pool_size = len(load_stock_pool(Path(args.workbook), args.limit, stock_codes))
    result = extract_quarterly(args)
    write_outputs(result, Path(args.output_csv), Path(args.output_xlsx), Path(args.summary_json), stock_pool_size)


if __name__ == "__main__":
    main()
