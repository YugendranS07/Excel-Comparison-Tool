# 🧮 Excel Comparison Tool

A lightweight yet powerful **Python automation utility** that compares data between two Excel files — typically a *Source* and an *Org Data* file — using a configurable *Mapping file*.
It’s designed to help QA engineers, analysts, and developers quickly validate data migrations or integration consistency.

---

## 🚀 Overview

The **Excel Comparison Tool** identifies field-level mismatches between two Excel datasets and generates a comprehensive Excel output containing:

- ✅ **Summary Sheet** — Pass/Fail status with total match & mismatch counts
- 📊 **Comparison Sheet** — record-by-record comparison between Source and Org Data
- ❌ **Failed Record Sheets** — rows and fields that didn’t match
- 🎨 **Color-coded output** — green for Pass, red for Fail for easy visual review

> Perfect for **data validation**, **migration QA**, and **cross-system reconciliation**.

---

## 🧩 How It Works

1. Reads all sheets from both Excel files (Source and Org Data)
2. Maps fields based on a provided mapping file
3. Compares each record using `Source ID` as the key
4. Treats blanks, boolean equivalents (`Y/1`, `N/0`), and matching values as *True Matches*
5. Generates an **Excel report** with detailed insights, summaries, and color formatting

---

## 📂 Input Files

| File | Purpose | Required Columns |
|------|----------|------------------|
| `Source_File_Account_details.xlsx` | The *Source* dataset | `Source ID` + mapped columns |
| `OrgData_Account.xlsx` | The *Org Data* dataset | `Source ID` + mapped columns |
| `Account_Mapping.xlsx` | Defines field relationships between Source and Org Data | `Source_Column`, `OrgData_Column` |

> ⚠️ Ensure both Excel files include the column **`Source ID`** as the unique key.

---

## 🧠 Key Features

- Reads **all sheets automatically** from both workbooks
- Handles **boolean equivalence** (`Y ↔ 1`, `N ↔ 0`, etc.)
- Treats **blank vs blank** as a match
- Supports **partial text containment** for fuzzy matching
- Produces a **multi-sheet output Excel report**
- Auto-applies **filters & color formatting** in the Summary sheet
- Entity/Cycle name used: **Object - Validation**

---

## 🧾 Example Output

### 🟢 Summary Sheet

| Source Field | OrgData Field | Total | Matches | Mismatches | Outcome |
|---------------|---------------|--------|----------|-------------|----------|
| Account Name | Account Name | 3 | 2 | 1 | ❌ Fail |
| Status | Status | 3 | 2 | 1 | ❌ Fail |
| Active | Active | 3 | 3 | 0 | ✅ Pass |
| **TOTAL** | | **9** | **7** | **2** |  |

> 🟢 Green = Pass 🔴 Red = Fail

---

### 📊 Comparison Sheet

| Source ID | Account Name (Source) | Account Name (OrgData) | Match |
|------------|----------------------|------------------------|--------|
| 101 | ABC Corp | ABC Corporation | False |
| 102 | XYZ Inc | XYZ Inc | True |
| 103 | LMN Pvt Ltd | LMN Pvt Ltd | True |

---

### ❌ Failed Record Sheet

| Source ID | Account Name (Source) | Account Name (OrgData) | OrgData Contains Source |
|------------|----------------------|------------------------|--------------------------|
| 101 | ABC Corp | ABC Corporation | True |

---

## ⚙️ Installation & Usage

### Step 1 — Clone the repository

```bash
git clone https://github.com/YugendranS07/Excel-Comparison-Tool.git
cd Excel-Comparison-Tool
```

### Step 2 — Install dependencies

```bash
pip install -r requirements.txt
```

### Step 3 — Run the tool

```bash
python excel_comparator.py
```

This generates an Excel output file named `comparison_output.xlsx`.

---

## 🧩 Project Structure

```
Excel-Comparison-Tool/
├── excel_comparator.py
├── requirements.txt
├── README.md
└── sample_files/
    ├── Source_File_Account_details.xlsx
    ├── OrgData_Account.xlsx
    ├── Account_Mapping.xlsx
```

---

## 🧪 Requirements

- Python ≥ 3.8
- pandas
- openpyxl

Install manually with:
```bash
pip install pandas openpyxl
```

---

## 💡 Example Use Cases

- Validating post-migration data between systems
- Comparing Source vs Target after ETL transformations
- Checking mapping consistency in crosswalk files
- QA testing for data warehouse or integration projects

---

## 🏷️ License

This project is licensed under the **MIT License** — free to use, share, and modify.

---

## ✨ Author

**Yugendran S**
_Data QA | Automation Engineer_  
📧 [LinkedIn Profile](https://www.linkedin.com/in/yugendran07/)

---

## ❤️ Acknowledgments

Special thanks to data QA professionals who inspired the idea behind automating field-level Excel comparisons — saving hours of manual validation work every day.
