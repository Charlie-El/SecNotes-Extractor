from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from ccxd_us_reports_app.runtime_paths import bundled_source_template
from ccxd_us_reports_app.vendor.annual import extract_notes_coverage as annual_notes
from ccxd_us_reports_app.vendor.quarterly import extract_quarterly_notes as quarterly_notes
from ccxd_us_reports_app.quarterly_output_utils import to_annual_compatible
from ccxd_us_reports_app.quarterly_output_utils import write_source_workbook
from ccxd_us_reports_app.vendor.quarterly import remap_quarterly_output as remapper


SOURCE_FILENAME = "美股季报提取_源表.xlsx"
MAPPED_FILENAME = "美股季报提取_修改映射.xlsx"
UNMATCHED_FILENAME = "美股季报提取_未匹配公司.xlsx"
SUMMARY_FILENAME = "季报提取摘要.json"


@dataclass
class QuarterlyRunConfig:
    notes_data_dir: Path
    main_pool_workbook: Path
    output_dir: Path
    extra_pool_workbook: Path | None = None
    mapping_workbook: Path | None = None
    enable_test_features: bool = False
    company_dir: Path | None = None


def _configure_shared_state(config: QuarterlyRunConfig) -> None:
    quarterly_notes.DATA_DIR = config.notes_data_dir
    annual_notes.DATA_DIR = config.notes_data_dir
    annual_notes.NOTES_DIRS = (
        sorted(path for path in config.notes_data_dir.iterdir() if path.is_dir() and path.name.endswith("_notes"))
        if config.notes_data_dir.exists()
        else []
    )
    annual_notes.WORKBOOK_PATH = config.main_pool_workbook
    annual_notes.EXTRA_STOCKBOOK_PATH = config.extra_pool_workbook or Path("__missing_extra_pool__.xlsx")
    annual_notes.AM_ANNUAL_DIR = (
        config.company_dir if config.enable_test_features and config.company_dir else Path("__missing_am_annual__")
    )

def _load_stock_pool(config: QuarterlyRunConfig) -> pd.DataFrame:
    return annual_notes.load_stock_pool()


def _extract_quarterly_note_source(config: QuarterlyRunConfig, stock_pool: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    matched = quarterly_notes.collect_latest_quarterly_filings(stock_pool, config.notes_data_dir, refresh_sec_map=False)
    covered = matched[matched["matched"]].copy()
    if covered.empty:
        return pd.DataFrame(columns=quarterly_notes.ROW_COLUMNS), matched

    all_rows: list[dict[str, Any]] = []
    for notes_dir_text, group in covered.groupby("notes_dir", sort=True):
        notes_dir = Path(notes_dir_text)
        adshs = set(group["adsh"].map(quarterly_notes.clean_text))
        cache = quarterly_notes.load_notes_subset(notes_dir, adshs)
        for _, row in group.iterrows():
            all_rows.extend(quarterly_notes.build_rows_for_filing(row, cache))

    raw = pd.DataFrame(all_rows, columns=quarterly_notes.ROW_COLUMNS)
    if raw.empty:
        return raw, matched
    raw = raw.sort_values(
        ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue_qtd", "segment_revenue_ytd"],
        ascending=[True, False, True, True, False, False],
        na_position="last",
        kind="stable",
    ).reset_index(drop=True)
    source = to_annual_compatible(raw, bundled_source_template())
    return source, matched


def _load_annual_matches(config: QuarterlyRunConfig, stock_pool: pd.DataFrame) -> pd.DataFrame:
    latest_filings = annual_notes.load_latest_filings(stock_pool)
    return latest_filings


def _extract_selected_annual_note_rows(
    config: QuarterlyRunConfig,
    stock_pool: pd.DataFrame,
    annual_matches: pd.DataFrame,
    selected_annual_codes: list[str],
) -> pd.DataFrame:
    if not selected_annual_codes:
        return pd.DataFrame(columns=annual_notes.ROW_COLUMNS)

    covered = annual_matches[annual_matches["matched"]].copy()
    covered = covered[covered["SECU_CODE"].astype(str).isin(selected_annual_codes)].copy()
    if covered.empty:
        return pd.DataFrame(columns=annual_notes.ROW_COLUMNS)

    all_rows: list[dict[str, Any]] = []
    for _, stock_row in covered.iterrows():
        all_rows.extend(annual_notes.build_company_and_segment_rows(stock_row))

    source = pd.DataFrame(all_rows, columns=annual_notes.ROW_COLUMNS)
    if config.enable_test_features:
        source = annual_notes.supplement_from_am_annual(source, stock_pool)
    if not source.empty:
        source = source.sort_values(
            ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue"],
            ascending=[True, False, True, True, False],
            na_position="last",
        ).reset_index(drop=True)
    return source


def _latest_company_rows(source: pd.DataFrame) -> pd.DataFrame:
    if source.empty:
        return source.copy()
    companies = source[source["row_type"].map(annual_notes.clean_text).eq("company")].copy()
    if companies.empty:
        return companies
    companies["_filed_ts"] = pd.to_datetime(companies.get("filed_date"), errors="coerce")
    companies["_period_ts"] = pd.to_datetime(companies.get("report_period"), errors="coerce")
    companies = companies.sort_values(
        ["stock_code", "_filed_ts", "_period_ts"],
        ascending=[True, False, False],
        na_position="last",
        kind="stable",
    )
    return companies.drop_duplicates("stock_code", keep="first").drop(columns=["_filed_ts", "_period_ts"])


def _build_latest_quarterly_company_map(source: pd.DataFrame) -> dict[str, dict[str, Any]]:
    companies = _latest_company_rows(source)
    return {
        str(row.get("stock_code")): row.to_dict()
        for _, row in companies.iterrows()
        if str(row.get("stock_code", "")).strip()
    }


def _build_latest_annual_company_map(annual_matches: pd.DataFrame) -> dict[str, dict[str, Any]]:
    covered = annual_matches[annual_matches["matched"]].copy()
    if covered.empty:
        return {}

    covered["_filed_ts"] = pd.to_datetime(covered.get("filed"), format="%Y%m%d", errors="coerce")
    covered["_period_ts"] = pd.to_datetime(covered.get("period"), format="%Y%m%d", errors="coerce")
    covered = covered.sort_values(
        ["SECU_CODE", "_filed_ts", "_period_ts", "adsh"],
        ascending=[True, False, False, False],
        na_position="last",
        kind="stable",
    )
    covered = covered.drop_duplicates("SECU_CODE", keep="first")

    annual_map: dict[str, dict[str, Any]] = {}
    for _, row in covered.iterrows():
        stock_code = str(row.get("SECU_CODE", "")).strip()
        if not stock_code:
            continue
        annual_map[stock_code] = {
            "stock_code": stock_code,
            "report_type": row.get("form"),
            "report_period": annual_notes.normalize_report_date(row.get("period")),
            "filed_date": annual_notes.normalize_report_date(row.get("filed")),
            "fiscal_period": row.get("fp"),
            "notes_file_batch": row.get("batch"),
            "adsh": row.get("adsh"),
        }
    return annual_map


def _combine_note_sources(
    quarterly_source: pd.DataFrame,
    annual_matches: pd.DataFrame,
) -> tuple[list[str], list[str]]:
    quarterly_by_code = _build_latest_quarterly_company_map(quarterly_source)
    annual_by_code = _build_latest_annual_company_map(annual_matches)
    selected_quarterly_codes: list[str] = []
    selected_annual_codes: list[str] = []
    all_codes = sorted(set(quarterly_by_code) | set(annual_by_code))
    for code in all_codes:
        quarterly_row = quarterly_by_code.get(code)
        annual_row = annual_by_code.get(code)
        best = quarterly_row
        annual_filed = pd.to_datetime(annual_row.get("filed_date") if annual_row else None, errors="coerce")
        quarterly_filed = pd.to_datetime(best.get("filed_date") if best else None, errors="coerce")
        if annual_row is not None and (pd.isna(quarterly_filed) or (pd.notna(annual_filed) and annual_filed > quarterly_filed)):
            best = annual_row
        elif annual_row is not None and best is None:
            best = annual_row

        if best is annual_row:
            selected_annual_codes.append(code)
        elif best is quarterly_row:
            selected_quarterly_codes.append(code)
    return selected_quarterly_codes, selected_annual_codes


def _assemble_final_source(
    quarterly_source: pd.DataFrame,
    annual_source: pd.DataFrame,
    selected_quarterly_codes: list[str],
    selected_annual_codes: list[str],
) -> pd.DataFrame:
    quarterly_rows = quarterly_source[quarterly_source["stock_code"].isin(selected_quarterly_codes)].copy()
    annual_rows = annual_source[annual_source["stock_code"].isin(selected_annual_codes)].copy()
    frames = [frame for frame in (quarterly_rows, annual_rows) if not frame.empty]
    if not frames:
        return pd.DataFrame(columns=quarterly_source.columns if not quarterly_source.empty else annual_source.columns)
    combined = pd.concat(frames, ignore_index=True)
    if combined.empty:
        return combined

    combined["_row_order"] = combined["row_type"].map({"company": 0, "segment": 1}).fillna(9)
    combined["_period"] = combined["report_period"].fillna("")
    combined = combined.sort_values(
        ["stock_code", "_period", "_row_order", "segment_type", "segment_name"],
        ascending=[True, False, True, True, True],
        na_position="last",
        kind="stable",
    ).drop(columns=["_row_order", "_period"])
    return combined.reset_index(drop=True)


def _write_unmatched_workbook(
    stock_pool: pd.DataFrame,
    quarterly_matches: pd.DataFrame,
    annual_matches: pd.DataFrame,
    output_path: Path,
) -> pd.DataFrame:
    quarter_match_map = {}
    if not quarterly_matches.empty:
        quarter_match_map = dict(
            zip(
                quarterly_matches["SECU_CODE"].astype(str),
                quarterly_matches["matched"].fillna(False),
            )
        )
    annual_match_map = {}
    if not annual_matches.empty:
        annual_match_map = dict(
            zip(
                annual_matches["SECU_CODE"].astype(str),
                annual_matches["matched"].fillna(False),
            )
        )

    rows: list[dict[str, Any]] = []
    for _, row in stock_pool.iterrows():
        stock_code = str(row.get("SECU_CODE", ""))
        quarter_hit = bool(quarter_match_map.get(stock_code, False))
        annual_hit = bool(annual_match_map.get(stock_code, False))
        if quarter_hit or annual_hit:
            continue
        rows.append(
            {
                "stock_code": row.get("SECU_CODE"),
                "company_name": row.get("SECUNAME"),
                "CUSIP": row.get("CUSIP"),
                "ticker": row.get("ticker"),
                "unmatched_reason": "未在季度 notes 或年报 notes 中匹配到 filing",
            }
        )
    unmatched = pd.DataFrame(rows)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        unmatched.to_excel(writer, index=False, sheet_name="未匹配公司")
    return unmatched


def run_quarterly_extract(config: QuarterlyRunConfig) -> dict[str, Any]:
    config.output_dir.mkdir(parents=True, exist_ok=True)
    _configure_shared_state(config)

    stock_pool = _load_stock_pool(config)
    quarterly_source, quarterly_matches = _extract_quarterly_note_source(config, stock_pool)
    annual_matches = _load_annual_matches(config, stock_pool)
    selected_quarterly_codes, selected_annual_codes = _combine_note_sources(quarterly_source, annual_matches)
    annual_source = _extract_selected_annual_note_rows(config, stock_pool, annual_matches, selected_annual_codes)
    combined_source = _assemble_final_source(quarterly_source, annual_source, selected_quarterly_codes, selected_annual_codes)

    source_path = config.output_dir / SOURCE_FILENAME
    write_source_workbook(combined_source, bundled_source_template(), source_path)

    unmatched_path = config.output_dir / UNMATCHED_FILENAME
    unmatched_df = _write_unmatched_workbook(stock_pool, quarterly_matches, annual_matches, unmatched_path)

    mapped_written = False
    mapped_path = config.output_dir / MAPPED_FILENAME
    if config.mapping_workbook:
        remapper.build_mapped_output(source_path, config.mapping_workbook, mapped_path)
        mapped_written = True
    elif mapped_path.exists():
        mapped_path.unlink()

    company_rows = combined_source[combined_source["row_type"] == "company"] if not combined_source.empty else pd.DataFrame()
    summary = {
        "company_rows": int((combined_source["row_type"] == "company").sum()) if not combined_source.empty else 0,
        "segment_rows": int((combined_source["row_type"] == "segment").sum()) if not combined_source.empty else 0,
        "output_rows": int(len(combined_source)),
        "quarterly_note_selected_companies": int(len(selected_quarterly_codes)),
        "annual_note_selected_companies": int(len(selected_annual_codes)),
        "unmatched_stocks": int(len(unmatched_df)),
        "covered_stocks": int(company_rows["stock_code"].nunique()) if not company_rows.empty else 0,
        "source_workbook": str(source_path),
        "mapped_workbook": str(mapped_path) if mapped_written else "",
        "unmatched_workbook": str(unmatched_path),
    }
    (config.output_dir / SUMMARY_FILENAME).write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        **summary,
        "mapped_written": mapped_written,
    }
