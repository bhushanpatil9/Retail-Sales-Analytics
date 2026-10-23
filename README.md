# Retail Sales Analytics Dashboard

An interactive Tableau portfolio project for exploring retail sales, customer behavior, order fulfillment, and quarterly costs. The repository includes the Tableau packaged workbook and its companion Excel source data, plus a lightweight validation script.

## Project contents

| Path | Description |
| --- | --- |
| [`tableau/retail_sales_dashboard.twbx`](tableau/retail_sales_dashboard.twbx) | Packaged Tableau workbook with the `Overview`, `Sales`, and `Orders` dashboards and bundled Hyper extracts. |
| [`data/retail_sales_source.xlsx`](data/retail_sales_source.xlsx) | Excel source workbook with four sheets: order-sales records, segment metrics, quarterly costs, and waffle-chart coordinates. |
| [`scripts/validate_project.py`](scripts/validate_project.py) | Checks the source workbook structure and the packaged Tableau workbook contents. |
| [`requirements.txt`](requirements.txt) | Python dependency for the validation script. |

## Dashboards

- **Overview** — headline sales, order, customer, and return views.
- **Sales** — sales breakdowns by time, geography, segment, category, and subcategory.
- **Orders** — order fulfillment, return, and repeat-customer views.

The workbook also contains supporting worksheets for cost trends, KPI metrics, maps, waffle charts, and small-multiple time trends. Open the `.twbx` file in Tableau to interact with filters, tooltips, and dashboard actions.

## Dataset

The source Excel workbook contains:

| Sheet | Grain / contents |
| --- | --- |
| `Order-Sales Data` | 21,900 transaction records with 14 named fields, including order and ship timestamps, sales, customer, geography, segment, fulfillment status, returns, repeat-customer flag, and on-time flag. |
| `Metrics` | 12 segment-level values across GPM, OPM, ROA, and ROE. |
| `Cost Data` | 60 quarterly cost values: five cost types over 12 quarters from 2022 Q1 through 2024 Q4. |
| `Waffle Chart` | 100 coordinate pairs for a 10 × 10 waffle chart. |

The transaction dates run from **2022-01-01 through 2024-12-30**. The data includes 10 countries, 3 segments, and 3 product categories (Furniture, Office Supplies, and Technology). The order ID appears on multiple records, so the 21,900-row count is a record count, not a count of distinct orders. Customer IDs identify 2,726 distinct customers in the source workbook.

### Data notes

- An extra empty trailing column is present in the order sheet; it is not a named data field and is ignored by the validator.
- The workbook includes derived KPI and cost values. Their business definitions and underlying accounting assumptions are not documented in the supplied files, so treat them as provided metrics rather than independently verified financial statements.
- The repository does not include data-source, privacy, or redistribution terms. Confirm those rights before republishing the source workbook.

## Key findings from the supplied data

These are descriptive summaries computed from the workbook, not claims about an identified real-world company:

- Total sales across the 21,900 records are **$3,203,440.00** (sum of the `Sales` field).
- The Consumer segment has the largest sales total (**$1,461,542.86**); Furniture has the largest category total (**$1,345,629.19**).
- **19,265 records (88.0%)** have `Complete` status; **1,759** are `Pending` and **876** are `Cancelled`.
- The supplied segment metrics show Consumer with the highest GPM (38%), while Corporate has the highest ROA (10%) and ROE (25%).

## Tools

- **Tableau Desktop or Tableau Public** — open and interact with the packaged `.twbx` dashboard.
- **Python 3.9+** and **openpyxl 3.1.5** — run the repository validation script.

No Tableau installation is needed to inspect the Excel data or run the validator. Tableau is required to open the interactive dashboards.

## Quick start

1. Clone or download this repository.
2. Open `tableau/retail_sales_dashboard.twbx` in Tableau Desktop or Tableau Public.
3. Select the `Overview`, `Sales`, or `Orders` dashboard tab. Use the controls and marks in the workbook to explore the data.

### Validate the repository files

From the repository root, create an environment:

```bash
python -m venv .venv
```

On Windows PowerShell, activate it and install the pinned dependency:

```powershell
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python scripts/validate_project.py
```

On macOS or Linux, use:

```bash
source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/validate_project.py
```

The validator checks the Excel sheets, required columns, completeness and basic value domains, then verifies that the Tableau package contains its workbook definition, expected dashboards, and referenced extracts. It prints a compact dataset summary and exits with an error if a structural check fails.

## Methodology

1. Use the Excel workbook as the documented source for the four data tables.
2. Use the bundled extracts in the Tableau packaged workbook for interactive exploration; the `.twbx` travels as a self-contained workbook package.
3. Summarize sales by summing `Sales`; summarize flags as record counts or rates over the relevant records. The key findings above state their denominators explicitly where applicable.
4. Read GPM, OPM, ROA, and ROE from the supplied segment metrics sheet. These are source-provided values, not recomputed from transaction-level profit or asset data.
5. Use quarter and cost type as supplied in the cost sheet. The repository does not provide a cost calculation method.

## Limitations

- No profit or asset fields are provided at transaction level, so segment financial KPIs cannot be independently recalculated from the order records.
- Duplicate order IDs mean order-level counts require a documented deduplication or line-item rule; this project reports row counts unless explicitly stated otherwise.
- The source data's origin, update cadence, and license are not stated in the repository.
- The Tableau workbook requires Tableau Desktop or Public for its interactive views.

## Repository layout

```text
.
├── data/
│   └── retail_sales_source.xlsx
├── scripts/
│   └── validate_project.py
├── tableau/
│   └── retail_sales_dashboard.twbx
├── .gitignore
├── requirements.txt
└── README.md
```
