from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd

from ccxd_us_reports_app.vendor.annual import extract_notes_coverage as annual_notes
from ccxd_us_reports_app.vendor.quarterly import remap_quarterly_output as remapper


SOURCE_FILENAME = "美股年报提取_源表.xlsx"
MAPPED_FILENAME = "美股年报提取_修改映射.xlsx"
UNMATCHED_FILENAME = "美股年报提取_未匹配公司.xlsx"
SUMMARY_FILENAME = "年报提取摘要.json"


@dataclass
class AnnualRunConfig:
    notes_data_dir: Path
    main_pool_workbook: Path
    output_dir: Path
    extra_pool_workbook: Path | None = None
    mapping_workbook: Path | None = None
    enable_test_features: bool = False
    am_annual_dir: Path | None = None


def _configure_vendor(config: AnnualRunConfig) -> None:
    annual_notes.DATA_DIR = config.notes_data_dir
    annual_notes.NOTES_DIRS = (
        sorted(path for path in config.notes_data_dir.iterdir() if path.is_dir() and path.name.endswith("_notes"))
        if config.notes_data_dir.exists()
        else []
    )
    annual_notes.WORKBOOK_PATH = config.main_pool_workbook
    annual_notes.EXTRA_STOCKBOOK_PATH = config.extra_pool_workbook or Path("__missing_extra_pool__.xlsx")
    annual_notes.AM_ANNUAL_DIR = (
        config.am_annual_dir if config.enable_test_features and config.am_annual_dir else Path("__missing_am_annual__")
    )
    annual_notes.OUTPUT_DIR = config.output_dir
    annual_notes.OUTPUT_XLSX = config.output_dir / SOURCE_FILENAME
    annual_notes.OUTPUT_CSV = config.output_dir / "美股年报提取_源表.csv"
    annual_notes.OUTPUT_SUMMARY = config.output_dir / SUMMARY_FILENAME

def _write_unmatched_workbook(stock_pool: pd.DataFrame, latest_filings: pd.DataFrame, output_path: Path) -> pd.DataFrame:
    if latest_filings.empty:
        unmatched = stock_pool.copy()
        unmatched["unmatched_reason"] = "未在年报 notes 中匹配到 filing"
    else:
        unmatched = latest_filings[~latest_filings["matched"].fillna(False)].copy()
        unmatched["unmatched_reason"] = "未在年报 notes 中匹配到 filing"

    keep_cols = [
        "SECU_CODE",
        "SECUNAME",
        "CUSIP",
        "ticker",
        "ticker_match_method",
        "unmatched_reason",
    ]
    unmatched = unmatched[[col for col in keep_cols if col in unmatched.columns]].copy()
    unmatched = unmatched.rename(columns={"SECU_CODE": "stock_code", "SECUNAME": "company_name"})
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        unmatched.to_excel(writer, index=False, sheet_name="未匹配公司")
    return unmatched


def _build_summary(
    result_df: pd.DataFrame,
    stock_pool: pd.DataFrame,
    unmatched_df: pd.DataFrame,
    mapped_written: bool,
    source_path: Path,
    mapped_path: Path,
    unmatched_path: Path,
) -> dict[str, Any]:
    company_rows = result_df[result_df["row_type"] == "company"] if not result_df.empty else pd.DataFrame()
    covered_count = int(company_rows["stock_code"].nunique()) if not company_rows.empty else 0
    coverage_pct = round((covered_count / len(stock_pool) * 100), 2) if len(stock_pool) else 0
    return {
        "notes_batches": [path.name for path in annual_notes.NOTES_DIRS],
        "pool_rows": int(len(stock_pool)),
        "covered_stocks": covered_count,
        "unmatched_stocks": int(len(unmatched_df)),
        "coverage_pct": coverage_pct,
        "company_rows": int((result_df["row_type"] == "company").sum()) if not result_df.empty else 0,
        "segment_rows": int((result_df["row_type"] == "segment").sum()) if not result_df.empty else 0,
        "output_rows": int(len(result_df)),
        "source_workbook": str(source_path),
        "mapped_workbook": str(mapped_path) if mapped_written else "",
        "unmatched_workbook": str(unmatched_path),
    }


def run_annual_extract(config: AnnualRunConfig) -> dict[str, Any]:
    config.output_dir.mkdir(parents=True, exist_ok=True)
    _configure_vendor(config)

    stock_pool = annual_notes.load_stock_pool()
    latest_filings = annual_notes.load_latest_filings(stock_pool)
    covered = latest_filings[latest_filings["matched"]].copy()

    all_rows: list[dict[str, Any]] = []
    for _, stock_row in covered.iterrows():
        all_rows.extend(annual_notes.build_company_and_segment_rows(stock_row))

    result_df = pd.DataFrame(all_rows, columns=annual_notes.ROW_COLUMNS)
    if config.enable_test_features:
        result_df = annual_notes.supplement_from_am_annual(result_df, stock_pool)
    if not result_df.empty:
        result_df = result_df.sort_values(
            ["stock_code", "report_period", "row_type", "segment_type", "segment_revenue"],
            ascending=[True, False, True, True, False],
            na_position="last",
        ).reset_index(drop=True)

    source_path = config.output_dir / SOURCE_FILENAME
    annual_notes.write_extract_outputs(result_df, annual_notes.OUTPUT_CSV, source_path)

    unmatched_path = config.output_dir / UNMATCHED_FILENAME
    unmatched_df = _write_unmatched_workbook(stock_pool, latest_filings, unmatched_path)

    mapped_written = False
    mapped_path = config.output_dir / MAPPED_FILENAME
    if config.mapping_workbook:
        remapper.build_mapped_output(source_path, config.mapping_workbook, mapped_path)
        mapped_written = True
    elif mapped_path.exists():
        mapped_path.unlink()

    summary = _build_summary(result_df, stock_pool, unmatched_df, mapped_written, source_path, mapped_path, unmatched_path)
    annual_notes.OUTPUT_SUMMARY.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        **summary,
        "mapped_written": mapped_written,
    }
