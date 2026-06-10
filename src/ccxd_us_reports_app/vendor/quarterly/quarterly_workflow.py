from __future__ import annotations

"""Unified entry for the quarterly delivery workflow.

Modes:
- full: base extraction -> annual fallback fill -> remap -> tracking workbook
- remap-only: rebuild mapped workbook from source and refresh tracking workbook
- track-only: refresh tracking workbook from existing outputs
"""

import argparse
import json
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook

from ccxd_us_reports_app.vendor.quarterly import fill_full_coverage_from_annual as filler
from ccxd_us_reports_app.vendor.quarterly import remap_quarterly_output as remapper
from ccxd_us_reports_app.vendor.quarterly import run_folder_quarterly_pipeline as qpipe


SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "outputs" / "folder_latest_quarterly"
TRACKING_WORKBOOK = OUTPUT_DIR / "\u7f8e\u80a1\u5b63\u62a5\u6700\u65b0\u516c\u544a\u8ddf\u8e2a.xlsx"
RULE_SHEET = "\u5224\u5b9a\u89c4\u5219"
SNAPSHOT_SHEET = "\u5f53\u524d\u5feb\u7167"
NEW_ONLY_SHEET = "\u65b0\u589e\u516c\u544a"
HISTORY_SHEET = "\u5386\u53f2\u8fd0\u884c"
SOURCE_SHEET = "\u603b\u8868"
QUARTERLY_FORMS = {"10-Q", "10-Q/A", "6-K", "6-K/A"}


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def text_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column in df.columns:
        return df[column].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index, dtype=object)


def make_signature(df: pd.DataFrame) -> pd.Series:
    parts = [
        text_series(df, "stock_code"),
        text_series(df, "report_type"),
        text_series(df, "report_period"),
        text_series(df, "filed_date"),
        text_series(df, "adsh"),
        text_series(df, "coverage_source"),
    ]
    return parts[0].str.cat(parts[1:], sep="|")


def build_run_args(args: argparse.Namespace) -> argparse.Namespace:
    return argparse.Namespace(
        company_dir=args.company_dir,
        main_pool=args.main_pool,
        extra_pool=args.extra_pool,
        dim_mapping=args.base_dim_mapping,
        annual_template=args.annual_template,
    )


def build_fill_args(args: argparse.Namespace) -> argparse.Namespace:
    return argparse.Namespace(
        company_dir=args.company_dir,
        annual_source=args.annual_source,
        quarterly_source=args.source_workbook,
        company_mapping=args.company_mapping_workbook,
        coverage_output=args.coverage_workbook,
        detail_output=args.detail_workbook,
        formatted_output=args.formatted_workbook,
    )


def load_current_snapshot(source_path: Path, coverage_path: Path) -> pd.DataFrame:
    source = pd.read_excel(source_path, sheet_name=SOURCE_SHEET, header=2, dtype=object)
    companies = source[source["row_type"].map(clean_text).eq("company")].copy()
    companies = companies.drop_duplicates("stock_code", keep="first")

    if coverage_path.exists():
        coverage = pd.read_excel(coverage_path, sheet_name="coverage", dtype=object)
        if "stock_code" not in coverage.columns and "SECU_CODE" in coverage.columns:
            coverage = coverage.rename(columns={"SECU_CODE": "stock_code"})
        elif "stock_code" in coverage.columns and "SECU_CODE" in coverage.columns:
            coverage = coverage.drop(columns=["SECU_CODE"])
        keep_cols = [
            "stock_code",
            "quarterly_covered",
            "final_covered",
            "coverage_source",
            "coverage_note",
            "quarterly_status",
            "folder_name",
            "folder_path",
        ]
        companies = companies.merge(coverage[[c for c in keep_cols if c in coverage.columns]], on="stock_code", how="left")

    companies["is_quarterly_result"] = text_series(companies, "report_type").map(clean_text).isin(QUARTERLY_FORMS)
    companies["selection_signature"] = make_signature(companies)
    companies["filed_date_ts"] = pd.to_datetime(companies.get("filed_date"), errors="coerce")
    companies["accepted_datetime_ts"] = pd.to_datetime(companies.get("accepted_datetime"), errors="coerce")
    companies["report_period_ts"] = pd.to_datetime(companies.get("report_period"), errors="coerce")

    ordered_cols = [
        "stock_code",
        "company_name",
        "report_type",
        "fiscal_period",
        "report_period",
        "filed_date",
        "accepted_datetime",
        "adsh",
        "notes_file_batch",
        "extraction_status",
        "coverage_source",
        "quarterly_status",
        "primary_disclosure_dim",
        "disclosure_dims_available",
        "segment_count",
        "total_revenue",
        "total_revenue_uom",
        "selection_signature",
        "folder_name",
        "folder_path",
        "coverage_note",
    ]
    ordered_cols = [c for c in ordered_cols if c in companies.columns]
    companies = companies[ordered_cols + [c for c in companies.columns if c not in ordered_cols]]
    companies = companies.sort_values(["filed_date_ts", "report_period_ts", "stock_code"], ascending=[False, False, True], kind="stable")
    return companies.reset_index(drop=True)


def load_previous_snapshot(tracking_path: Path) -> pd.DataFrame:
    if not tracking_path.exists():
        return pd.DataFrame()
    try:
        wb = load_workbook(tracking_path, read_only=True)
        if SNAPSHOT_SHEET not in wb.sheetnames:
            wb.close()
            return pd.DataFrame()
        wb.close()
        return pd.read_excel(tracking_path, sheet_name=SNAPSHOT_SHEET, dtype=object)
    except Exception:
        return pd.DataFrame()


def compare_snapshots(previous: pd.DataFrame, current: pd.DataFrame) -> pd.DataFrame:
    current = current.copy()
    if previous.empty:
        out = current.copy()
        out["change_type"] = "initial_baseline"
        out["previous_report_type"] = ""
        out["previous_report_period"] = ""
        out["previous_filed_date"] = ""
        out["previous_adsh"] = ""
        out["previous_coverage_source"] = ""
        return out

    prev = previous.copy()
    prev = prev.drop_duplicates("stock_code", keep="first")
    current = current.drop_duplicates("stock_code", keep="first")
    prev = prev.rename(
        columns={
            "report_type": "previous_report_type",
            "report_period": "previous_report_period",
            "filed_date": "previous_filed_date",
            "adsh": "previous_adsh",
            "coverage_source": "previous_coverage_source",
            "selection_signature": "previous_selection_signature",
        }
    )
    merged = current.merge(
        prev[
            [
                "stock_code",
                "previous_report_type",
                "previous_report_period",
                "previous_filed_date",
                "previous_adsh",
                "previous_coverage_source",
                "previous_selection_signature",
            ]
        ],
        on="stock_code",
        how="left",
    )

    def classify(row: pd.Series) -> str:
        prev_sig = clean_text(row.get("previous_selection_signature"))
        curr_sig = clean_text(row.get("selection_signature"))
        if not prev_sig:
            return "new_company"
        if prev_sig != curr_sig:
            return "new_filing"
        return "unchanged"

    merged["change_type"] = merged.apply(classify, axis=1)
    changed = merged[merged["change_type"] != "unchanged"].copy()
    return changed.reset_index(drop=True)


def build_rule_sheet(current: pd.DataFrame, new_only: pd.DataFrame, run_mode: str, previous: pd.DataFrame) -> pd.DataFrame:
    run_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    quarterly = current[text_series(current, "coverage_source").eq("quarterly_filing")].copy()
    cutoff_pool = quarterly if not quarterly.empty else current

    current_cutoff_filed = ""
    current_cutoff_report = ""
    if not cutoff_pool.empty:
        cutoff_row = cutoff_pool.sort_values(["filed_date_ts", "report_period_ts", "stock_code"], ascending=[False, False, True], kind="stable").iloc[0]
        current_cutoff_filed = clean_text(cutoff_row.get("filed_date"))
        current_cutoff_report = clean_text(cutoff_row.get("report_period"))

    prev_cutoff = ""
    if not previous.empty and "filed_date" in previous.columns:
        prev_dates = pd.to_datetime(previous["filed_date"], errors="coerce").dropna()
        if not prev_dates.empty:
            prev_cutoff = prev_dates.max().strftime("%Y-%m-%d")

    annual_fallback_count = int(text_series(current, "coverage_source").isin(["annual_source_workbook", "local_annual_html_fallback"]).sum())
    rows = [
        ["run_mode", run_mode, "\u672c\u6b21\u8fd0\u884c\u6a21\u5f0f"],
        ["run_timestamp", run_ts, "\u811a\u672c\u6267\u884c\u65f6\u95f4"],
        ["as_of_date", datetime.now().strftime("%Y-%m-%d"), "\u672c\u6b21\u7ed3\u679c\u89c6\u4f5c\u622a\u6b62\u5230\u5f53\u5929"],
        [
            "latest_quarterly_selection_rule",
            "per cik10: sort by filed_date desc -> report_period desc -> adsh desc, then keep first quarterly filing",
            "\u5e95\u5c42 latest \u5b63\u62a5\u9009\u53d6\u89c4\u5219",
        ],
        [
            "full_coverage_rule",
            "prefer latest quarterly filing from notes; if latest annual source/local annual html is newer by filed_date, replace the quarterly result; if quarterly is missing, also fall back",
            "\u5168\u8986\u76d6\u8865\u9f50\u89c4\u5219",
        ],
        [
            "mapped_output_rule",
            "segment_type is remapped from segment_axis by the mapping workbook; table1 primary_disclosure_dim and disclosure_dims_available are recomputed from remapped table2",
            "\u4fee\u6539\u6620\u5c04\u8868\u751f\u6210\u89c4\u5219",
        ],
        ["current_cutoff_filed_date", current_cutoff_filed, "\u672c\u6b21\u5df2\u8986\u76d6\u5230\u7684\u6700\u65b0 filed_date"],
        ["current_max_report_period", current_cutoff_report, "\u4e0e\u6700\u65b0 filed_date \u5bf9\u5e94\u7684 report_period"],
        ["previous_cutoff_filed_date", prev_cutoff, "\u4e0a\u6b21\u57fa\u7ebf\u7684 filed_date \u622a\u6b62\u65e5"],
        [
            "next_incremental_rule",
            f"prefer filings with filed_date > {current_cutoff_filed or 'current_cutoff'}; if the same company gets a new adsh or report_period, treat it as new announcement too",
            "\u4e0b\u4e00\u5b63\u5ea6\u53ef\u76f4\u63a5\u6309\u8fd9\u6761\u89c4\u5219\u505a\u589e\u91cf",
        ],
        ["current_company_count", int(len(current)), "\u672c\u6b21\u8f93\u51fa\u516c\u53f8\u603b\u6570"],
        ["current_quarterly_company_count", int(text_series(current, "coverage_source").eq("quarterly_filing").sum()), "\u672c\u6b21\u4fdd\u7559\u5b63\u62a5\u7684\u516c\u53f8\u6570"],
        ["current_annual_fallback_company_count", annual_fallback_count, "\u672c\u6b21\u7528\u5e74\u62a5\u5145\u586b\u7684\u516c\u53f8\u6570"],
        ["new_announcement_count_vs_previous_run", int(len(new_only)), "\u76f8\u5bf9\u4e0a\u6b21\u8fd0\u884c\u7684\u65b0\u516c\u544a\u6570"],
    ]
    return pd.DataFrame(rows, columns=["item", "value", "description"])


def load_history(tracking_path: Path) -> pd.DataFrame:
    if not tracking_path.exists():
        return pd.DataFrame()
    try:
        wb = load_workbook(tracking_path, read_only=True)
        if HISTORY_SHEET not in wb.sheetnames:
            wb.close()
            return pd.DataFrame()
        wb.close()
        return pd.read_excel(tracking_path, sheet_name=HISTORY_SHEET, dtype=object)
    except Exception:
        return pd.DataFrame()


def append_history(history: pd.DataFrame, current: pd.DataFrame, new_only: pd.DataFrame, run_mode: str) -> pd.DataFrame:
    current_dates = pd.to_datetime(text_series(current, "filed_date"), errors="coerce")
    current_cutoff = current_dates.max().strftime("%Y-%m-%d") if not current_dates.dropna().empty else ""
    period_dates = pd.to_datetime(text_series(current, "report_period"), errors="coerce")
    current_max_period = period_dates.max().strftime("%Y-%m-%d") if not period_dates.dropna().empty else ""
    row = pd.DataFrame(
        [
            {
                "run_timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "run_mode": run_mode,
                "current_cutoff_filed_date": current_cutoff,
                "current_max_report_period": current_max_period,
                "company_count": int(len(current)),
                "quarterly_company_count": int(text_series(current, "coverage_source").eq("quarterly_filing").sum()),
                "annual_fallback_company_count": int(text_series(current, "coverage_source").isin(["annual_source_workbook", "local_annual_html_fallback"]).sum()),
                "new_announcement_count": int(len(new_only)),
            }
        ]
    )
    out = pd.concat([history, row], ignore_index=True) if not history.empty else row
    return out


def write_tracking_workbook(path: Path, rules: pd.DataFrame, current: pd.DataFrame, new_only: pd.DataFrame, history: pd.DataFrame) -> None:
    export_current = current.drop(columns=[c for c in ["filed_date_ts", "accepted_datetime_ts", "report_period_ts"] if c in current.columns], errors="ignore")
    export_new = new_only.drop(columns=[c for c in ["filed_date_ts", "accepted_datetime_ts", "report_period_ts", "previous_selection_signature"] if c in new_only.columns], errors="ignore")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        rules.to_excel(writer, sheet_name=RULE_SHEET, index=False)
        export_current.to_excel(writer, sheet_name=SNAPSHOT_SHEET, index=False)
        export_new.to_excel(writer, sheet_name=NEW_ONLY_SHEET, index=False)
        history.to_excel(writer, sheet_name=HISTORY_SHEET, index=False)
    filler.format_simple_workbook(path)


def refresh_tracking(args: argparse.Namespace, run_mode: str) -> dict[str, Any]:
    source_path = Path(args.source_workbook)
    coverage_path = Path(args.coverage_workbook)
    tracking_path = Path(args.tracking_workbook)
    current = load_current_snapshot(source_path, coverage_path)
    previous = load_previous_snapshot(tracking_path)
    new_only = compare_snapshots(previous, current)
    rules = build_rule_sheet(current, new_only, run_mode, previous)
    history = append_history(load_history(tracking_path), current, new_only, run_mode)
    write_tracking_workbook(tracking_path, rules, current, new_only, history)
    return {
        "tracking_workbook": str(tracking_path),
        "current_company_count": int(len(current)),
        "new_announcement_count": int(len(new_only)),
        "run_mode": run_mode,
    }


def run_full_workflow(args: argparse.Namespace) -> dict[str, Any]:
    pipeline_summary = qpipe.run_pipeline(build_run_args(args))
    fill_summary = filler.run_fill_full_coverage(build_fill_args(args))
    remap_summary = remapper.build_mapped_output(
        Path(args.source_workbook),
        Path(args.final_axis_mapping),
        Path(args.formatted_workbook),
    )
    tracking_summary = refresh_tracking(args, "full")
    return {
        "pipeline": pipeline_summary,
        "fill_coverage": fill_summary,
        "remap": remap_summary,
        "tracking": tracking_summary,
    }


def run_remap_only(args: argparse.Namespace) -> dict[str, Any]:
    remap_summary = remapper.build_mapped_output(
        Path(args.source_workbook),
        Path(args.final_axis_mapping),
        Path(args.formatted_workbook),
    )
    tracking_summary = refresh_tracking(args, "remap-only")
    return {"remap": remap_summary, "tracking": tracking_summary}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Unified quarterly workflow runner.")
    parser.add_argument("--mode", choices=["full", "remap-only", "track-only"], default="full")
    parser.add_argument("--company-dir", default=str(qpipe.DEFAULT_COMPANY_DIR))
    parser.add_argument("--main-pool", default=str(qpipe.DEFAULT_MAIN_POOL))
    parser.add_argument("--extra-pool", default=str(qpipe.DEFAULT_EXTRA_POOL))
    parser.add_argument("--base-dim-mapping", default=str(qpipe.DEFAULT_DIM_MAPPING))
    parser.add_argument("--annual-template", default=str(qpipe.DEFAULT_ANNUAL_TEMPLATE))
    parser.add_argument("--annual-source", default=str(filler.DEFAULT_ANNUAL_SOURCE))
    parser.add_argument("--source-workbook", default=str(filler.DEFAULT_QUARTERLY_SOURCE))
    parser.add_argument("--formatted-workbook", default=str(filler.DEFAULT_FORMATTED))
    parser.add_argument("--company-mapping-workbook", default=str(filler.DEFAULT_COMPANY_MAPPING))
    parser.add_argument("--coverage-workbook", default=str(filler.DEFAULT_COVERAGE))
    parser.add_argument("--detail-workbook", default=str(filler.DEFAULT_DETAIL))
    parser.add_argument("--final-axis-mapping", default=str(remapper.DEFAULT_MAPPING))
    parser.add_argument("--tracking-workbook", default=str(TRACKING_WORKBOOK))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.mode == "full":
        summary = run_full_workflow(args)
    elif args.mode == "remap-only":
        summary = run_remap_only(args)
    else:
        summary = refresh_tracking(args, "track-only")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
