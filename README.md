# Excel Comparison Tool (Source vs Org Data)

**Author:** Yugendran S
**License:** MIT

A Python-based automation tool that validates large-scale data migrations by comparing a **Source Excel file** against a **Target/Org Data Excel file**, using a configurable **Mapping file** to determine which fields correspond to each other. It produces a single, structured Excel report covering summary results, full field-by-field comparison, and isolated mismatch details.

---

## 📌 The Problem This Tool Solves

### The Real-World Business Pain Point

In most data migration projects — moving data from a legacy system to a new system, or from a source system to an org/target system — the business needs assurance that the migrated data is **accurate, complete, and unchanged in meaning**, even if it has changed in structure or format.

This sounds simple in theory. In practice, it becomes one of the most time-consuming and error-prone parts of any migration project, for the following reasons:

#### 1. Volume of data makes manual testing impossible
Source files in real migration projects often contain **hundreds of thousands to millions of records**, spread across multiple columns and sheets. Manually opening two Excel files side-by-side and comparing values row by row simply does not scale.

- Validating a few hundred rows manually? Feasible.
- Validating a few thousand rows? Painful, but doable with effort.
- Validating **millions** of rows, across dozens of mapped columns? **Not humanly possible** within any reasonable project timeline.

#### 2. Testers are forced into sampling — and sampling hides risk
Because full manual validation isn't feasible, testers/QA teams fall back on **random sampling**: pick a small subset of records (say, 50–100 out of millions) and manually verify those.

This creates a serious, often unspoken risk:

> **The records that are *not* tested might be exactly the ones that contain errors.**

A sampled test can pass with flying colors while thousands of untested records downstream are corrupted, mismatched, or incorrectly transformed — and nobody finds out until it surfaces in production, sometimes weeks or months later, when it's far more expensive to fix.

#### 3. Timelines suffer
Even partial manual validation efforts can stretch to **weeks or months** depending on data volume, especially when testers have to re-check after every new data drop, mapping change, or system update. This directly delays go-live decisions and increases pressure on QA teams to "just sign off" without full confidence.

#### 4. Inconsistent data representations cause false mismatches
Even when data *is* correct, differences in formatting between source and target systems create noise:
- `Y` vs `1` vs `TRUE` — logically the same, but textually different
- Blank vs `NULL` vs empty string — often meant to represent the same "no value" state
- Case sensitivity, whitespace, and partial/substring matches (e.g., a code embedded inside a longer descriptive string)

Manual testers have to mentally account for all of this every time they compare two cells — which is both slow and inconsistent from tester to tester.

#### 5. Lack of a structured, auditable report
Even when manual testing is completed, the output is often just tester notes, screenshots, or a scattered set of spreadsheet comments — not a clean, structured artifact that stakeholders can review, audit, or use as documented sign-off evidence.

---

## 💡 The Solution

Instead of relying on manual, partial, sample-based testing, this tool performs **100% automated validation** of every record and every mapped field — completing in **minutes** what would otherwise take **weeks or months** manually.

### Core Idea
Give the tool three inputs:
1. **Source File** — the original/legacy data
2. **Org Data File** — the migrated/target data
3. **Mapping File** — a simple two-column file defining which Source column maps to which Org Data column

The tool then automatically:
- Matches every record between both files using a unique **Source ID**, regardless of row order or which sheet the data sits in
- Compares every mapped field for every matched record — not a sample, **all of it**
- Intelligently normalizes values so that formatting differences (Y/N, 1/0, TRUE/FALSE, blank vs blank) don't produce false mismatches
- Flags genuine mismatches only, with an additional "contains" check to catch partial/substring matches that a strict equality check would incorrectly flag as failed
- Outputs everything into one clean, structured, color-coded Excel workbook

---

## ⚙️ How It Works (Technical Overview)

### Inputs
| File | Purpose |
|---|---|
| `Source_File_*.xlsx` | The original data before migration |
| `OrgData_*.xlsx` | The migrated/target data after migration |
| `Mapping_File.xlsx` | Two columns: `Source_Column`, `OrgData_Column` — defines field-to-field correspondence |

### Processing Steps
1. **Read all sheets** from both the Source and Org Data files and combine them into unified DataFrames using `pandas` + `openpyxl`.
2. **Deduplicate and index** both datasets by a unique `Source ID` column, so records can be reliably matched regardless of row order.
3. **Validate the mapping file** — only mapped column pairs that actually exist in both files are used for comparison; invalid/missing mappings are automatically skipped.
4. **Find common Source IDs** between both files (the intersection) — these are the records eligible for comparison.
5. **Field-by-field comparison** for every mapped column across every common record:
   - Normalizes boolean-like values (`Y`/`N`, `1`/`0`, `TRUE`/`FALSE`, `NULL`, blank) into a consistent form before comparing
   - Treats blank-vs-blank as a match (no false failures from missing values on both sides)
   - Falls back to a **substring "contains" check** when exact values don't match but one value is logically contained within the other
6. **Aggregate results** into:
   - Per-field match/mismatch counts
   - Overall Pass/Fail outcome per field
   - A grand total row summarizing the entire comparison run

### Output — Single Excel Workbook
| Sheet | Contents |
|---|---|
| **Summary** | One row per mapped field pair, showing Total Records, Matched Count, Mismatched Count, and Outcome (Pass/Fail) — with **color-coded highlighting** (green = pass, red = fail) for instant visual triage |
| **Comparison** | The complete side-by-side comparison of every field for every record — the full audit trail |
| **`<Field>_vs_<Field>`** (one per mismatched field pair) | Isolated sheets showing *only* the genuinely mismatched records for that specific field pair, so testers can jump directly to what needs investigation instead of scrolling through millions of matched rows |

All sheets have **auto-filters** applied automatically for fast, on-the-fly filtering and sorting.

---

## 🎯 Key Design Decisions & Accepted Behaviors

These behaviors were deliberately built in based on real testing edge cases encountered during use:

- ✅ Records are matched by **Source ID**, independent of row order in either file
- ✅ All fields for a given Source ID are consolidated into a **single row** in the Comparison output — no fragmented per-field rows
- ✅ **Blank vs blank** is treated as a match, not a mismatch
- ✅ **`Y`/`N`** are treated as equivalent to **`1`/`0`** and **`TRUE`/`FALSE`**
- ✅ Original Excel formats (especially **dates**) are preserved rather than being reformatted or corrupted during processing
- ✅ Literal **`'null'`** string values are preserved as-is, not silently converted to blanks — since in some systems, `'null'` is a meaningful data value, not an absence of data
- ✅ Sheet names are automatically sanitized to remove invalid Excel characters and truncated to Excel's 31-character sheet name limit

---

## 📈 Real-World Impact

| Traditional Manual Testing | This Tool |
|---|---|
| Samples a small subset of records | Validates **100% of records** |
| Takes **weeks to months** for large datasets | Runs in **minutes** |
| Risk of undetected errors in untested records | **No blind spots** — every record, every field, checked |
| Inconsistent handling of formatting differences (Y/N, blanks, nulls) across testers | **Consistent, automated normalization** logic every time |
| Manual notes/screenshots as "evidence" | Structured, color-coded, auditable **Excel report** |
| Re-testing after every data refresh is a fresh manual effort | Re-run the script in minutes after any new data drop |

This directly addresses one of the most common and costly gaps in data migration QA: **the false confidence created by sample-based testing**, where the untested majority of records could be silently carrying errors into production.

---

## 📁 Sample Files (For Reference)

To make it easy to understand exactly how this tool behaves on a realistic dataset, I've uploaded a **sample file package containing 10,000 records** alongside this project. It includes:

| File | Description |
|---|---|
| `Source_File_Account_Details.xlsx` | 10,000 sample source account records (names, account numbers, emails, phone numbers, status, active flag, created date, balance, region, notes) |
| `OrgData_Account.xlsx` | The corresponding "migrated" version of the same 10,000 records — with row order intentionally shuffled to prove that matching happens by **Source ID**, not by row position |
| `Account_Mapping.xlsx` | The field mapping file used to tell the tool which Source column corresponds to which OrgData column |
| `Sample_Comparison_Output.xlsx` | The **actual output** produced by running this tool against the above three files — included purely as a reference so you can see the Summary, Comparison, and mismatch sheets without having to run anything yourself |

The sample data intentionally includes realistic migration issues — flipped status flags, dropped emails, reformatted phone numbers, rounding differences in balances, shifted dates, and formatting differences like `Y`/`N` vs `TRUE`/`FALSE` — so the output report tells a genuine, believable validation story rather than showing a dataset that matches perfectly by default.

You're welcome to open the sample output file directly to see exactly what the tool produces before running it on your own data.

---

## 🚀 Usage

### Requirements
```bash
pip install pandas openpyxl
```

### Configuration
Edit the config section at the top of the script:

```python
SOURCE_FILE = "sample_files/Source_File_Account_details.xlsx"
ORG_DATA_FILE = "sample_files/OrgData_Account.xlsx"
MAPPING_FILE = "sample_files/Account_Mapping.xlsx"
OUTPUT_FILE = "comparison_output.xlsx"
ENTITY_NAME = "Object - Validation"
SOURCE_ID_COL = "Source ID"
```

### Run
```bash
python excel_comparator.py
```

### Output
A single file, `comparison_output.xlsx`, containing the Summary, Comparison, and per-field mismatch sheets described above.

---

## 🔮 Potential Future Enhancements

- Configurable matching logic beyond substring "contains" checks (e.g., fuzzy matching, tolerance thresholds for numeric fields)
- Command-line arguments / config file support instead of hardcoded paths
- Support for very large datasets via chunked processing to reduce memory usage
- HTML/PDF report generation alongside the Excel output
- Email or Slack notification integration for automated CI-based validation runs

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 🔗 Repository

[https://github.com/YugendranS07/Excel-Comparison-Tool](https://github.com/YugendranS07/Excel-Comparison-Tool/tree/main)
