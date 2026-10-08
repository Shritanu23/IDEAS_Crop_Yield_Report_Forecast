# Crop Yield Forecast Reporting & Data Pipeline
Data Science Internship Institution: IDEAS - Institute of Data Engineering, Analytics, and Science Technology Innovation Hub: Indian Statistical Institute, Kolkata.

> **Data Analyst / Data Engineering Project**  
> A PostgreSQL-driven reporting pipeline that extracts crop-yield forecast data by crop, state, year, season, and prediction method, then generates formatted Word reports automatically.

## 📌 Project Overview

This project automates the generation of crop-wise yield forecast reports from a PostgreSQL database.

The pipeline:

1. Accepts a **prediction year** and **season** as inputs.
2. Identifies the prediction methods available for that period.
3. Retrieves forecast yield and RMSE values from PostgreSQL.
4. Joins crop, state, and yield tables.
5. Compares the forecast year with the previous two years of reported data.
6. Organizes results by **crop → state → year → method**.
7. Generates a structured `.docx` report with dynamic tables.
8. Adds report formatting, headers, logo, page breaks, and footer information.

The project demonstrates practical skills in **SQL, PostgreSQL, Python, data extraction, data transformation, reporting automation, and business-oriented data presentation**.

---

## 🎯 Business Problem

Agricultural forecasting produces large amounts of crop- and state-level data. Manually converting these results into standardized reports can be time-consuming and error-prone.

This project addresses that problem by creating a reusable reporting workflow that can:

- reduce manual reporting effort;
- standardize report formatting;
- compare current forecasts with historical values;
- display forecast accuracy using RMSE;
- support multiple crops, states, seasons, and prediction methods;
- generate reports on demand using command-line parameters.

---

## 🏗️ Project Architecture

```text
PostgreSQL Database
        │
        ▼
  SQL Data Extraction
        │
        ├── crop_yields
        ├── crops
        └── states
        │
        ▼
 Python Data Processing
        │
        ├── Year validation
        ├── Historical-year calculation
        ├── Crop/state organization
        ├── Prediction-method detection
        └── Yield + RMSE mapping
        │
        ▼
 Dynamic Report Generator
        │
        ▼
 Microsoft Word (.docx)
```

---

## 🧰 Tech Stack

| Category | Tools / Technologies |
|---|---|
| Programming | Python |
| Database | PostgreSQL |
| SQL Connectivity | psycopg2 |
| Reporting | python-docx |
| Data Processing | Python dictionaries / structured transformations |
| Version Control | Git / GitHub |
| CI/CD | GitHub Actions |
| Output | Microsoft Word (.docx) |

### Key Python Libraries

- `psycopg2`
- `python-docx`
- `argparse`
- `datetime`
- `collections`

---

## 🗃️ Database Structure

The reporting script expects a PostgreSQL database containing the `NPCYF` schema.

### Main tables

#### `NPCYF.crop_yields`

Contains crop-level forecast/reporting data.

Important fields used by the pipeline:

- `crop_id`
- `state_id`
- `year`
- `season`
- `method`
- `yield_value`
- `rmse_value`

#### `NPCYF.crops`

Maps crop IDs to crop names.

- `corp_id`
- `corp_name`

#### `NPCYF.states`

Maps state IDs to state names.

- `state_id`
- `state_name`

### Relationship

text
crops
  │
  │ crop_id
  ▼
crop_yields
  ▲
  │ state_id
  │
states


---

## 🔎 SQL Logic

The project uses SQL to:

- identify available prediction methods;
- filter data by prediction year and season;
- retrieve historical comparison years;
- join crop and state reference tables;
- return yield and RMSE values.

Example logic:

```sql
SELECT
    c.corp_name,
    s.state_name,
    cy.year,
    cy.method,
    cy.yield_value,
    cy.rmse_value
FROM NPCYF.crop_yields cy
JOIN NPCYF.crops c
    ON c.corp_id = cy.crop_id
JOIN NPCYF.states s
    ON s.state_id = cy.state_id
WHERE cy.year = ANY(%s)
  AND cy.season = %s
ORDER BY
    c.corp_name,
    s.state_name,
    cy.year DESC,
    cy.method;
```

This demonstrates practical use of:

- `JOIN`
- filtering with `WHERE`
- parameterized SQL
- `ANY()`
- `ORDER BY`
- relational database design

---

## ⚙️ How the Pipeline Works

### 1. Input parameters

The script accepts:

- template path
- output path
- page orientation
- logo path
- prediction year
- season

Example:

```bash
python gen_report.py ^
  --template fasal_templete.docx ^
  --output crop_yield_report.docx ^
  --format LANDSCAPE ^
  --logo ISI_Logo.jpg ^
  --year 2025-26 ^
  --season Kharif
```

### 2. Historical comparison

For a prediction year such as:

```text
2025-26
```

the pipeline automatically identifies:

```text
2024-25
2023-24
```

This allows the generated report to provide historical context alongside the latest forecast.

### 3. Dynamic prediction methods

The script queries PostgreSQL for the prediction methods available for the selected year and season.

This means the report does not need a hard-coded list of forecasting methods.

### 4. Data organization

The retrieved records are organized into a nested structure:

```text
Crop
 └── State
      └── Year
           └── Prediction Method
                ├── Yield
                └── RMSE
```

### 5. Automated report generation

The pipeline generates crop-specific tables containing:

- State
- Forecast year
- Prediction method
- Yield
- RMSE
- Historical yield values

Each crop receives a separate section/page.

---

## 📊 Example Business Output

A generated report can answer questions such as:

- What is the forecasted yield for a particular crop and state?
- Which prediction method was used?
- How accurate is the prediction based on RMSE?
- How does the forecast compare with the previous two years?
- Which crops/states require closer monitoring?

This makes the project relevant to **data reporting, business intelligence, operational analytics, and decision support**.

---


### . Generate a report

```bash
python gen_report.py \
  --template fasal_templete.docx \
  --output crop_yield_report.docx \
  --format LANDSCAPE \
  --logo ISI_Logo.jpg \
  --year 2025-26 \
  --season Kharif
```

---

## 📈 Data Analyst Skills Demonstrated

This project demonstrates several skills that are valuable for Data Analyst positions:

### SQL & Database Analysis

- PostgreSQL
- Multi-table joins
- Filtering
- Sorting
- Parameterized queries
- Relational data modeling
- Historical comparisons

### Python

- Database connectivity
- Data transformation
- Dictionaries and nested structures
- Functions
- Exception handling
- Command-line arguments
- Automated reporting

### Data Quality

- Input validation
- Error handling
- Missing-data handling in report generation
- Historical-period validation

### Reporting & Business Communication

- Automated report generation
- Dynamic tables
- Forecast vs historical comparison
- RMSE-based model evaluation
- Standardized reporting

---

## 🧪 Testing & Code Quality

The repository includes GitHub Actions configuration for:

- Python environment setup
- dependency installation
- `flake8` linting
- `pytest`

For a stronger production-quality repository, add unit tests for:

- financial-year parsing;
- previous-year calculation;
- SQL result transformation;
- missing-data scenarios;
- invalid seasons;
- invalid prediction-year formats.

Example:

```python
def test_get_previous_years():
    assert get_previous_years("2025-26") == [
        "2024-25",
        "2023-24"
    ]
```

---

## ⚠️ Important Project Scope

This repository primarily implements the **data extraction and automated reporting layer**.

The Python script does **not itself train a forecasting model**. Forecast values and RMSE values are retrieved from the PostgreSQL `crop_yields` table.

Therefore, the project should be described accurately as:

> **A PostgreSQL-driven crop yield forecast reporting and automation pipeline**

rather than claiming that the repository itself performs machine-learning forecasting.

---


