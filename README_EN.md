# US SEC Note Annual and Quarterly Report Extractor

[中文说明](README.md)

This is a Windows desktop application for extracting US annual and quarterly report data from the SEC Financial Statement and Notes Data Sets and exporting standardized Excel workbooks.

The tool is designed for end users. In normal use, no code changes are required. Users only need to select the note data folder, stock pool workbook, optional mapping workbook, and output folder.

## Overview

The application provides two separate workflows:

- `Annual Extraction`
- `Quarterly Extraction`

Both workflows always produce a source workbook. If a mapping workbook is provided, a remapped workbook is also generated. An unmatched company workbook is also exported for review.

## Data Source

Official SEC download page:

- https://www.sec.gov/data-research/sec-markets-data/financial-statement-notes-data-sets

According to the SEC official page:

- The dataset provides text and detailed numeric information from financial statements and their notes.
- The data is extracted from exhibits to corporate financial reports filed with the SEC using XBRL.
- Compared with the Financial Statement Data Sets, this dataset contains significantly more disclosure detail.
- The information is provided as filed and presented in a flattened format to make cross-company and cross-period analysis easier.
- The datasets also include additional fields such as Standard Industrial Classification.
- The datasets are currently updated monthly. Before November 2020, they were updated quarterly.
- The SEC notes that these datasets are intended to assist analysis and are not a substitute for reviewing full filings.

If you are new to this workflow, review the SEC page and its documentation before preparing your local note data folders.

## Features

- Extract annual report data from SEC note datasets
- Extract quarterly report data from SEC note datasets
- In quarterly mode, prefer the latest quarterly filing; if the latest annual filing in notes has a later `filed_date`, use the annual filing instead
- Support optional mapping workbooks for both annual and quarterly workflows
- Export unmatched company workbooks in both workflows
- Keep local raw filing / HTML fallback hidden by default as a testing-only feature

## Output Files

### Annual Extraction

Always generated:

- `美股年报提取_源表.xlsx`
- `美股年报提取_未匹配公司.xlsx`

Optional:

- `美股年报提取_修改映射.xlsx`

This workbook is generated only when an annual mapping workbook is provided.

### Quarterly Extraction

Always generated:

- `美股季报提取_源表.xlsx`
- `美股季报提取_未匹配公司.xlsx`

Optional:

- `美股季报提取_修改映射.xlsx`

This workbook is generated only when a quarterly mapping workbook is provided.

## GUI Field Guide

### Quarterly Extraction

- `季报 note 数据目录`
  Main quarterly note data source. The application first selects the latest quarterly filing, then compares it with the latest annual filing in notes and keeps the newer one by `filed_date`.

- `主股票池工作簿`
  Main stock pool workbook that defines which companies will be processed.

- `补充股票池工作簿`
  Optional extra stock pool workbook.

- `季报映射表（可选）`
  Optional mapping workbook used to map `segment_axis` to `segment_type`.

- `输出目录`
  Folder where result files are written.

- `本地年报原件目录（测试）`
  Hidden by default. Only shown when testing mode is enabled. Used only for local fallback testing.

### Annual Extraction

- `年报 note 数据目录`
  Main annual note data source.

- `主股票池工作簿`
  Main stock pool workbook that defines which companies will be processed.

- `补充股票池工作簿`
  Optional extra stock pool workbook.

- `年报映射表（可选）`
  Optional mapping workbook used to map `segment_axis` to `segment_type`.

- `输出目录`
  Folder where result files are written.

- `本地年报原件目录（测试）`
  Hidden by default. Only shown when testing mode is enabled. Used only for local fallback testing.

## Mapping Workbooks

Annual and quarterly workflows use separate mapping templates.

Rules:

- Do not modify the first two header rows
- The actual mapping rule is `segment_axis -> segment_type`

Required header rows:

Row 1:

- `原始XBRL维度名称`
- `primary_disclosure_dim/disclosure_dims_available/segment_axis`

Row 2:

- `segment_axis`
- `segment_type`

Templates:

- [templates/季报映射模板.xlsx](templates/季报映射模板.xlsx)
- [templates/年报映射模板.xlsx](templates/年报映射模板.xlsx)

The application can also generate templates directly from the GUI:

- `生成季报映射模板`
- `生成年报映射模板`

## How to Use

### Run the packaged application

The packaged executable is not stored directly in the repository.

For end users, distribute the packaged application through GitHub Releases.

### Quarterly Workflow

1. Open the application and choose `季报提取`
2. Select the quarterly note data folder
3. Select the main stock pool workbook
4. Optionally select the extra stock pool workbook
5. Optionally select the quarterly mapping workbook
6. Select the output folder
7. Click `开始运行`

### Annual Workflow

1. Open the application and choose `年报提取`
2. Select the annual note data folder
3. Select the main stock pool workbook
4. Optionally select the extra stock pool workbook
5. Optionally select the annual mapping workbook
6. Select the output folder
7. Click `开始运行`

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

## GitHub Publishing Guidance

Recommended to upload:

- `src/`
- `templates/`
- `README.md`
- `README_EN.md`
- `build.ps1`
- `package_release.py`
- `.gitignore`

Do not upload:

- private note datasets
- real business outputs
- local test folders
- temporary caches
- intermediate build folders
- packaged EXE folders or zip release artifacts

## Rebuild

To rebuild the packaged release:

```powershell
.\build.ps1
```

The script will:

- remove old `build/`
- remove old `dist/`
- rebuild the EXE
- recreate the local `release/` folder
- recreate the local zip package

## Notes

- The application does not download SEC data automatically
- Users must prepare the local note datasets and stock pool workbooks
- Annual and quarterly workflows are independent
- Remapped workbooks are generated only when a mapping file is provided
- Local raw filing fallback is hidden by default and should be treated as a testing feature

## Public Repository Note

This repository has been cleaned for public publishing:

- no private business datasets
- no real business output files
- no local test outputs
- no cached or intermediate build artifacts
- no large packaged EXE artifacts tracked in Git

For external distribution, upload the packaged artifacts to GitHub Releases instead of committing them to the repository.
