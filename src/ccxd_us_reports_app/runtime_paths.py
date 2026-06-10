from __future__ import annotations

import sys
from pathlib import Path


def app_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def templates_dir() -> Path:
    return app_root() / "templates"


def bundled_source_template() -> Path:
    return templates_dir() / "内置原表模板.xlsx"


def bundled_mapped_template() -> Path:
    return templates_dir() / "内置修改映射模板.xlsx"


def quarterly_mapping_template_path() -> Path:
    return templates_dir() / "季报映射模板.xlsx"


def annual_mapping_template_path() -> Path:
    return templates_dir() / "年报映射模板.xlsx"
