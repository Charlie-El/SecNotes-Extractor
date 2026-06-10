from __future__ import annotations

import shutil
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent
DIST_DIR = ROOT / "dist" / "ccxd_us_reports_tool"
RELEASE_ROOT = ROOT / "release"
RELEASE_NAME = "美股年报季报提取工具"
RELEASE_DIR = RELEASE_ROOT / RELEASE_NAME
README_FILE = ROOT / "README.md"
README_EN_FILE = ROOT / "README_EN.md"
TEMPLATES_DIR = ROOT / "templates"


def _ensure_templates() -> None:
    sys.path.insert(0, str(ROOT / "src"))
    from ccxd_us_reports_app.mapping_template import (
        create_annual_mapping_template,
        create_quarterly_mapping_template,
    )
    from ccxd_us_reports_app.release_templates import create_mapped_template, create_source_template

    TEMPLATES_DIR.mkdir(parents=True, exist_ok=True)
    quarterly_path = TEMPLATES_DIR / "季报映射模板.xlsx"
    annual_path = TEMPLATES_DIR / "年报映射模板.xlsx"
    source_template_path = TEMPLATES_DIR / "内置原表模板.xlsx"
    mapped_template_path = TEMPLATES_DIR / "内置修改映射模板.xlsx"
    create_quarterly_mapping_template(quarterly_path)
    create_annual_mapping_template(annual_path)
    create_source_template(source_template_path)
    create_mapped_template(mapped_template_path)


def main() -> None:
    if not DIST_DIR.exists():
        raise SystemExit(f"Missing build output: {DIST_DIR}")

    _ensure_templates()

    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    RELEASE_ROOT.mkdir(parents=True, exist_ok=True)

    shutil.copytree(DIST_DIR, RELEASE_DIR)

    src_exe = RELEASE_DIR / "ccxd_us_reports_tool.exe"
    dst_exe = RELEASE_DIR / f"{RELEASE_NAME}.exe"
    if src_exe.exists():
        src_exe.rename(dst_exe)

    shutil.copy2(README_FILE, RELEASE_DIR / "README.md")
    if README_EN_FILE.exists():
        shutil.copy2(README_EN_FILE, RELEASE_DIR / "README_EN.md")
    if TEMPLATES_DIR.exists():
        shutil.copytree(TEMPLATES_DIR, RELEASE_DIR / "templates", dirs_exist_ok=True)

    archive_base = RELEASE_ROOT / RELEASE_NAME
    archive_file = archive_base.with_suffix(".zip")
    if archive_file.exists():
        archive_file.unlink()
    shutil.make_archive(str(archive_base), "zip", RELEASE_ROOT, RELEASE_NAME)


if __name__ == "__main__":
    main()
