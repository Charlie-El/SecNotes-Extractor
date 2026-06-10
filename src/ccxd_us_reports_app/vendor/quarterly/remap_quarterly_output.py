from __future__ import annotations

"""Rebuild the mapped quarterly workbook from the source workbook only.

This script never edits the source workbook. It reads 缇庤偂瀛ｆ姤鎻愬彇_婧愯〃.xlsx plus
鍒嗙被缁村害鏄犲皠.xlsx and regenerates 缇庤偂瀛ｆ姤鎻愬彇_淇敼鏄犲皠.xlsx.
"""

import argparse
import json
from pathlib import Path
from typing import Any

import pandas as pd

from ccxd_us_reports_app.quarterly_output_utils import add_country_translation
from ccxd_us_reports_app.quarterly_output_utils import load_dimension_revisions
from ccxd_us_reports_app.quarterly_output_utils import write_formatted_workbook


SCRIPT_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = SCRIPT_DIR / "outputs" / "folder_latest_quarterly"
DEFAULT_SOURCE = OUTPUT_DIR / "\u7f8e\u80a1\u5b63\u62a5\u63d0\u53d6_\u6e90\u8868.xlsx"
DEFAULT_MAPPING = SCRIPT_DIR / "\u5206\u7c7b\u7ef4\u5ea6\u6620\u5c04.xlsx"
DEFAULT_FORMATTED = OUTPUT_DIR / "\u7f8e\u80a1\u5b63\u62a5\u63d0\u53d6_\u4fee\u6539\u6620\u5c04.xlsx"
SOURCE_SHEET = "\u603b\u8868"


def clean_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def load_source(path: Path) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=SOURCE_SHEET, header=2, dtype=object)


def load_axis_mapping(path: Path) -> dict[str, str]:
    return load_dimension_revisions(path)


def row_key(row: pd.Series) -> tuple[str, str, str]:
    return (
        clean_text(row.get("stock_code")),
        clean_text(row.get("report_period")),
        clean_text(row.get("adsh")),
    )


def apply_axis_mapping(source_df: pd.DataFrame, axis_mapping: dict[str, str]) -> pd.DataFrame:
    out = source_df.copy()
    segment_mask = out["row_type"].map(clean_text).eq("segment")
    mapped_axes = out.loc[segment_mask, "segment_axis"].map(lambda value: axis_mapping.get(clean_text(value), ""))
    hit_index = mapped_axes[mapped_axes != ""].index
    out.loc[hit_index, "segment_type"] = mapped_axes.loc[hit_index]
    out = add_country_translation(out)
    out.attrs["axis_mapping_changes"] = int(len(hit_index))
    return out


def recompute_company_dimensions(source_df: pd.DataFrame) -> pd.DataFrame:
    out = source_df.copy()
    out["_company_key"] = out.apply(row_key, axis=1)

    company_mask = out["row_type"].map(clean_text).eq("company")
    segment_mask = out["row_type"].map(clean_text).eq("segment")
    segment_groups = {
        key: group.copy()
        for key, group in out.loc[segment_mask].groupby("_company_key", sort=False, dropna=False)
    }

    changed = 0
    for idx, row in out.loc[company_mask].iterrows():
        segments = segment_groups.get(row["_company_key"])
        if segments is None or segments.empty:
            continue

        dim_counts: dict[str, int] = {}
        for value in segments["segment_type"]:
            segment_type = clean_text(value)
            if segment_type:
                dim_counts[segment_type] = dim_counts.get(segment_type, 0) + 1

        if not dim_counts:
            continue

        dims_available = sorted(dim_counts)
        primary_dim = sorted(dim_counts.items(), key=lambda item: (-item[1], item[0]))[0][0]

        out.at[idx, "primary_disclosure_dim"] = primary_dim
        out.at[idx, "disclosure_dims_available"] = json.dumps(dims_available, ensure_ascii=False)
        out.at[idx, "segment_count"] = dim_counts.get(primary_dim, 0)
        out.at[idx, "segment_disclosure_flag"] = True
        out.at[idx, "has_revenue_breakdown"] = True
        changed += 1

    out = out.drop(columns=["_company_key"])
    out.attrs["company_dimension_updates"] = changed
    return out


def build_mapped_output(source_path: Path, mapping_path: Path, output_path: Path) -> dict[str, Any]:
    source_df = load_source(source_path)
    axis_mapping = load_axis_mapping(mapping_path)
    remapped = apply_axis_mapping(source_df, axis_mapping)
    remapped = recompute_company_dimensions(remapped)
    write_formatted_workbook(remapped, output_path)

    summary = {
        "source_rows": int(len(source_df)),
        "company_rows": int((remapped["row_type"].map(clean_text) == "company").sum()),
        "segment_rows": int((remapped["row_type"].map(clean_text) == "segment").sum()),
        "axis_mapping_rules": int(len(axis_mapping)),
        "axis_mapping_changes": int(remapped.attrs.get("axis_mapping_changes", 0)),
        "company_dimension_updates": int(remapped.attrs.get("company_dimension_updates", 0)),
        "source_workbook": str(source_path),
        "formatted_workbook": str(output_path),
    }
    return summary


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Rebuild quarterly mapped workbook from source workbook.")
    parser.add_argument("--source", default=str(DEFAULT_SOURCE))
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING))
    parser.add_argument("--output", default=str(DEFAULT_FORMATTED))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    summary = build_mapped_output(Path(args.source), Path(args.mapping), Path(args.output))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

