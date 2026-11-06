# Retail Product Data: Cleaning & Transformation in Microsoft Excel

## 📌 Repository Description (GitHub About Section)
> End-to-end Microsoft Excel data preparation workflow showcasing missing value imputation, text standardization, feature splitting/merging, and dynamic conditional formatting on a retail product dataset.

---

## 📖 Overview

Raw operational data often suffers from inconsistencies, missing attributes, duplicate records, and unstructured composite identifiers. This project documents a complete data preparation and transformation pipeline performed in **Microsoft Excel**.

Using a sample dataset containing product attributes, pricing, categories, and inventory metrics, this project implements industry-standard data cleaning techniques to ensure data integrity, readiness for downstream business intelligence (BI), and statistical reliability.

---

## 🗂 Project Structure

```text
├── data/
│   ├── raw/
│   │   └── raw_product_data.xlsx         # Unprocessed input dataset
│   └── cleaned/
│       └── cleaned_product_data.xlsx     # Cleaned, standardized dataset
├── Excel Assignment 2 - Data Cleaning and Transformation.xlsx
└── README.md                             # Project documentation & formula reference
```

---

## 🛠 Tasks & Methodology

### 1. Handling Missing Values
* **Price Imputation (`Price ($)`):**
  * *Issue:* Blank price entries skew total valuation and basic summaries.
  * *Solution:* Imputed blanks using the mean average of valid price entries with an absolute reference range.
  * *Formula:*
    ```excel
    =IF(ISBLANK(H2), AVERAGE($H$2:$H$32), H2)
    ```
* **Category Backfilling (`Category`):**
  * *Issue:* Unassigned product classifications.
  * *Solution:* Identified null cells and populated them with `"Unknown"` (or inferred from context).
  * *Formula:*
    ```excel
    =IF(ISBLANK(K2), "Unknown", K2)
    ```

---

### 2. Correcting Inconsistent Data
* **Standardizing Text & Whitespace:**
  * *Issue:* Inconsistent casing and trailing/leading accidental spaces in `Product Name`.
  * *Solution:* Applied `TRIM` to strip unwanted whitespace and `PROPER` for uniform title casing.
  * *Formula:*
    ```excel
    =PROPER(TRIM(D2))
    ```
* **Correcting Category Typos:**
  * *Issue:* Misspelled category names (e.g., `"Electrnics"` instead of `"Electronics"`).
  * *Solution:* Utilized Excel's global Find & Replace dialog (`Ctrl + H`) to ensure clean categorical groupings.

---

### 3. Deduplication
* **Eliminating Duplicate Records:**
  * *Issue:* Identical repeated rows inflating inventory and product counts.
  * *Solution:* Ran the native **Remove Duplicates** utility (`Data` > `Data Tools` > `Remove Duplicates`) across all columns (`A:K`) to retain only unique product transactions.

---

### 4. Feature Extraction & Engineering
* **Splitting Composite `Product ID` (e.g., `28-JAN-US`):**
  * **Manufacturing Date:** Parsed the date string component (day and month) and converted it into a serial date value.
    ```excel
    =DATEVALUE(LEFT(A2, 6) & "-2026")
    ```
  * **Country Code:** Extracted the terminal 2-character country identifier.
    ```excel
    =RIGHT(A2, 2)
    ```
* **Merging Brand and Product:**
  * Combined separate `Brand` and `Product Name` fields into a unified `Product Brand` column for clearer labeling.
  * *Formula using `TEXTJOIN`:*
    ```excel
    =TEXTJOIN(" ", TRUE, E2, D2)
    ```
  * *Formula using concatenation:*
    ```excel
    =E2 & " " & D2
    ```

---

### 5. Data Formatting & Normalization
* **Currency Formatting:**
  * Standardized the `Price ($)` column to standard currency formatting with two decimal places (`[$$-409]#,##0.00`).
* **Date Normalization:**
  * Converted varied date inputs to a standard `DD-MM-YYYY` format via Custom Cell Formatting (`Ctrl + 1`).

---

### 6. Visual Analysis & Conditional Formatting
* **Dynamic Data Bars:**
  * Configured in-cell gradient data bars on the price column (`Home` > `Conditional Formatting` > `Data Bars`) to provide immediate relative price indicators without needing a separate chart.
* **Category Highlighting:**
  * Configured an automated cell highlighting rule (`Equal To... "Electronics"`) to highlight high-priority tech products instantly.

---

## 📊 Summary of Functions & Tools Used

| Category | Function / Tool | Purpose |
| :--- | :--- | :--- |
| **Logic & Imputation** | `IF`, `ISBLANK`, `AVERAGE` | Detecting empty cells and replacing with statistical measures |
| **Text Transformation** | `TRIM`, `PROPER` | Cleaning erratic casing and removing extraneous whitespace |
| **Text Extraction** | `LEFT`, `RIGHT` | Parsing composite codes (`Product ID`) |
| **Date Transformation** | `DATEVALUE` | Converting parsed string segments into valid Excel serial dates |
| **String Aggregation** | `TEXTJOIN`, `&` | Merging brand and product names cleanly |
| **Data Hygiene** | Remove Duplicates, Find & Replace | Eliminating redundant records and typographical errors |
| **Visualization** | Conditional Formatting Rules | Applying data bars and category-specific cell styles |

---

## 🚀 How to Use

1. Clone or download this repository:
   ```bash
   git clone https://github.com/<your-username>/<repo-name>.git
   ```
2. Open `Excel Assignment 2 - Data Cleaning and Transformation.xlsx` in **Microsoft Excel 2016** or newer (or Google Sheets / Excel Online).
3. Review the formulas in the calculated columns to see the data transformation pipeline in action.
