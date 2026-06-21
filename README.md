# CSI Daily Monitor

An automated R reporting and forecasting workflow for daily customer
satisfaction monitoring at branch and service-advisor levels.

> This repository contains a sanitized portfolio version. Customer records,
> employee names, company locations, credentials, internal database names, and
> generated management reports are not included.

## Key Features

- Loads survey responses and outreach records from MariaDB
- Applies configurable monthly reporting periods
- Compares current cumulative KPIs with the previous day's results
- Monitors service execution, high satisfaction, and low satisfaction
- Calculates response rates and follow-up opportunity volumes
- Generates card-based HTML management reports
- Produces branch and service-advisor performance summaries
- Forecasts cumulative CSI indicators with ARIMA and Prophet
- Compares actual and forecast trends through interactive Plotly charts
- Identifies questionnaire items associated with core CSI outcomes
- Exports follow-up lists and detailed records to multi-sheet Excel files

## Monitored KPIs

- Survey sample size
- Service execution rate
- High-satisfaction rate
- Low-satisfaction rate
- Survey response rate
- Follow-up opportunity count and rate
- Daily and cumulative performance changes

## Reporting Period Logic

The CSI reporting month uses a custom survey window:

- Survey responses: previous month day 21 through current month day 20
- Outreach response monitoring: current calendar month
- Data is capped at the day before the script runs

The target reporting month is configured with `REPORT_YM`.

## Forecasting and Analysis

- ARIMA cumulative KPI forecasts
- Prophet cumulative KPI forecasts
- Forecast error and threshold evaluation
- Point-biserial correlation for questionnaire drivers
- Strongest and weakest question analysis
- Branch-level and advisor-level interactive trend views

## Technology

- R
- MariaDB
- `DBI` and `RMariaDB`
- `dplyr`, `tidyr`, `lubridate`, and `stringr`
- `gt`, `htmltools`, `webshot2`, and `pagedown`
- `ggplot2` and `plotly`
- `forecast` and `prophet`
- `psych`, `ltm`, and `fmsb`
- `writexl`

## Project Structure

```text
.
├── scripts/
│   └── csi_daily_monitor.R
├── .env.example
├── .gitignore
└── README.md
```

## Configuration

Set the following environment variables before running the monitor:

```text
MARIADB_HOST
MARIADB_PORT
MARIADB_USER
MARIADB_PASSWORD
MARIADB_DATABASE
MARIADB_TABLE
REPORT_YM
REPORT_OUTPUT_DIR
```

See `.env.example` for non-secret example values.

## Installation

```r
install.packages(c(
  "readxl", "dplyr", "lubridate", "ggplot2", "gt", "scales",
  "htmltools", "webshot2", "pdftools", "pagedown", "forecast",
  "tidyr", "forcats", "psych", "ltm", "fmsb", "plotly",
  "prophet", "writexl", "stringr", "DBI", "RMariaDB"
))
```

## Running the Monitor

After configuring the environment:

```r
source("scripts/csi_daily_monitor.R", encoding = "UTF-8")
```

The workflow expects the configured MariaDB table to provide the survey,
outreach, branch, advisor, and questionnaire fields referenced by the script.

## Main Outputs

- Daily CSI HTML management report
- Branch-level KPI summaries
- Advisor-level cumulative trend visualizations
- ARIMA and Prophet forecasts
- Questionnaire-driver charts
- Follow-up opportunity Excel report
- Survey and response-detail Excel workbook

## Data Privacy

Generated reports can contain customer comments, employee names, operational
performance, and contact-level records. Keep the `outputs` directory private
and never commit real reports or source data.
