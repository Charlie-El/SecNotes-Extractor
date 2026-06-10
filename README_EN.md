# US SEC Note Annual and Quarterly Report Extractor

[中文说明](README.md)

US SEC Note Annual and Quarterly Report Extractor is a Windows desktop tool designed for business users. It extracts US annual and quarterly report data from the SEC Financial Statement and Notes Data Sets and exports standardized Excel workbooks. In normal use, the tool runs on local `note` datasets and does not require code changes.

The application provides two separate workflows:

- `Annual Extraction`
- `Quarterly Extraction`

Both workflows always generate a source workbook and an unmatched company workbook. If a mapping workbook is provided, a remapped workbook is also generated for downstream review and delivery.

## Key Features

- Extract annual report data from SEC `note` datasets.
- Extract quarterly report data from SEC `note` datasets.
- In quarterly mode, prefer the latest quarterly filing; if the latest annual filing in `note` data has a later `filed_date`, use the annual filing instead.
- Support separate annual and quarterly mapping workbooks.
- Export unmatched company workbooks for review.
- Keep local raw filing / HTML fallback hidden by default as a testing-only feature.

## Data Source

This tool is designed for the SEC Financial Statement and Notes Data Sets.

Official download page:

- https://www.sec.gov/data-research/sec-markets-data/financial-statement-notes-data-sets

According to the SEC, these datasets provide structured text and numeric information from financial statements and their footnotes. The data is extracted from XBRL exhibits filed with the SEC. Compared with the Financial Statement Data Sets, this dataset contains significantly more disclosure detail from notes and is better suited for structured extraction, cross-company comparison, and batch analysis. The datasets are currently updated monthly.

The application does not download SEC data automatically. Prepare the required local `note` folders before use.

## Quick Start

If you are using the packaged Windows application, simply run the executable from the release package.

If you want to run from source, use:

```powershell
python -m ccxd_us_reports_app.app
```

To rebuild the desktop application locally, run:

```powershell
.\build.ps1
```

The packaged files will be generated in the local `release\` folder.

## GUI Workflow

1. Open the application.
2. Choose `Annual Extraction` or `Quarterly Extraction`.
3. Select the `note` data folder.
4. Select the main stock pool workbook.
5. Optionally select the extra stock pool workbook.
6. Optionally select the mapping workbook.
7. Select the output folder.
8. Click `开始运行`.
9. Review the generated files in the output folder.

The GUI shows only the fields required for normal use by default. Local filing / HTML fallback fields are shown only when testing mode is enabled.

## Input Guide

### Quarterly Extraction

- `季报 note 数据目录`
  Main quarterly data source. The application first extracts the latest quarterly filing for a company, then compares it with the latest annual filing in `note` data and keeps the newer one by `filed_date`.

- `主股票池工作簿`
  Defines which companies will be processed.

- `补充股票池工作簿`
  Used to add companies beyond the main stock pool. Can be left empty.

- `季报映射表（可选）`
  Maps `segment_axis` to `segment_type`. The quarterly remapped workbook is generated only when this file is provided.

- `输出目录`
  Destination folder for output files.

- `本地年报原件目录（测试）`
  Hidden by default and shown only in testing mode. Used for local fallback when `note` data is insufficient.

### Annual Extraction

- `年报 note 数据目录`
  Main annual data source. The application extracts the latest annual filing from this folder.

- `主股票池工作簿`
  Defines which companies will be processed.

- `补充股票池工作簿`
  Used to add companies beyond the main stock pool. Can be left empty.

- `年报映射表（可选）`
  Maps `segment_axis` to `segment_type`. The annual remapped workbook is generated only when this file is provided.

- `输出目录`
  Destination folder for output files.

- `本地年报原件目录（测试）`
  Hidden by default and shown only in testing mode. Used for local fallback testing.

## Output Files

### Annual Mode

Always generated:

- `美股年报提取_源表.xlsx`
- `美股年报提取_未匹配公司.xlsx`

Optional:

- `美股年报提取_修改映射.xlsx`

### Quarterly Mode

Always generated:

- `美股季报提取_源表.xlsx`
- `美股季报提取_未匹配公司.xlsx`

Optional:

- `美股季报提取_修改映射.xlsx`

The remapped workbook is generated only when a mapping workbook is provided.

## Mapping Workbooks

Annual and quarterly workflows use separate mapping templates.

Template files:

- [templates/季报映射模板.xlsx](templates/季报映射模板.xlsx)
- [templates/年报映射模板.xlsx](templates/年报映射模板.xlsx)

The GUI can also generate templates directly:

- `生成季报映射模板`
- `生成年报映射模板`

Rules:

- Do not modify the first two header rows.
- The actual mapping rule is `segment_axis -> segment_type`.

The required header rows are:

Row 1:

- `原始XBRL维度名称`
- `primary_disclosure_dim/disclosure_dims_available/segment_axis`

Row 2:

- `segment_axis`
- `segment_type`

## Project Structure

```text
secnotesextractor/
├─ src/
│  └─ ccxd_us_reports_app/
├─ templates/
│  ├─ 季报映射模板.xlsx
│  └─ 年报映射模板.xlsx
├─ README.md
├─ README_EN.md
├─ build.ps1
├─ package_release.py
└─ .gitignore
```

## Notes

- The application does not download SEC data automatically.
- Prepare local `note` datasets and stock pool workbooks before use.
- Annual and quarterly workflows are independent.
- Remapped workbooks are generated only when a mapping file is provided.
- Local raw filing fallback is hidden by default and is intended mainly for testing.
- Output files should still be reviewed against the original filings when needed.

## Acknowledgements

Thanks to all users who provided feedback during data preparation, template validation, and tool testing. Thanks also to open-source projects such as `pandas`, `openpyxl`, and `PyInstaller` for the foundational capabilities behind this tool.

## Copyright

Copyright (c) 2026 Liu Juncheng. All rights reserved.

This project is intended for SEC note data extraction, structured processing, and Excel output support. Users should still review generated results independently. The output of this tool does not constitute investment, financial, legal, or audit advice.
