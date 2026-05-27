# Project Description

## Project Title
**Executive Records Collation and Categorization System**
*For CACYOF OYSCATECH ALUMNI*

---

## Background

CACYOF OYSCATECH ALUMNI maintains a record of elected executives for each session year. Over the years, these records have been stored as separate Excel files — one per session — making it difficult to answer questions like:

- *Who has served as President across all sessions?*
- *Who has held the Bible Coordinator role over the years?*
- *How many people have served in the Prayer Coordinator position?*

Manually compiling these records across six or more files is time-consuming and error-prone, especially as the number of sessions grows.

---

## Problem Statement

1. Executive records are spread across multiple yearly Excel files with no unified view.
2. The same position is written differently across files (e.g. `"Ass. Bible Coordinator"` vs `"Bible Coordinator II"` vs `"Bible Cord"`), making grouping by office unreliable.
3. Some executives hold more than one office in the same session (e.g. `"Vice President / Prayer Coordinator"`), and must appear under both positions in any organized record.
4. There is no chronological ordering of who held each office over the years.

---

## Solution

A Python-based automation script that:

- Reads all yearly Excel files from a folder
- Extracts executive names, positions, phone numbers, and session years
- Splits compound positions so each person appears under every office they held
- Normalizes inconsistent position titles into a single standard name per office
- Sorts all records chronologically within each position
- Generates one formatted Excel report per position, plus a master summary file

---

## Scope

### In Scope
- Reading `.xlsx` executive files (one per session year)
- Extracting: Full Name, Position(s), Phone Number, Year/Session
- Position normalization (30+ title variants mapped)
- Compound position splitting (e.g. `A / B` → two separate entries)
- Chronological sorting of records per position
- Formatted Excel output per position
- Master summary Excel file
- Reusable: re-running with new files updates all outputs automatically

### Out of Scope
- Web or mobile interface
- Database storage
- Online/cloud sync
- Email or print distribution

---

## Inputs

| Item | Description |
|------|-------------|
| Yearly Excel files | One `.xlsx` file per session, named by year |
| File location | All files placed in the `input_files/` folder |
| File structure | Row 1: session year · Row 2: headers · Row 3+: records |

---

## Outputs

| File | Description |
|------|-------------|
| `President.xlsx` | All presidents across all sessions |
| `Vice_President.xlsx` | All vice presidents |
| `Bible_Coordinator.xlsx` | All bible coordinators |
| *(one file per position)* | … |
| `_SUMMARY_ALL_EXECUTIVES.xlsx` | Every executive, every position, every year |

Each output file contains:
- Organization heading: **CACYOF OYSCATECH ALUMNI**
- Position title as subheading
- Table: S/N · Name · Year Served · Phone Number
- Records sorted from earliest to latest session

---

## Key Design Decisions

### 1. Position Normalization Dictionary
A centralized `NORMALIZATION` dictionary maps raw position fragments to canonical names. Sorted by key length (longest first) to prevent short keys from incorrectly overriding longer, more specific ones.

### 2. Compound Position Splitting
Before splitting on `/`, `&`, or `and`, the script runs a set of regex pre-mappings to correctly resolve known compound titles (e.g. `"Drama/Tech Cord"` → `"Drama Coordinator / Technical Coordinator"`). This prevents ambiguous splits.

### 3. Chronological Sorting
A `year_sort_key()` function extracts the first 4-digit year from a session string (e.g. `"2019/2020"` → `2019`) and uses it to sort all records within each position group before writing to Excel.

### 4. Reusability
The script is stateless — it reads all files fresh every run. Adding a new session requires only dropping the new file into `input_files/` and re-running. No configuration changes are needed unless a new position title variant is introduced.

---

## Technologies

| Technology | Role |
|------------|------|
| Python 3.9+ | Core language |
| pandas | Reading and parsing Excel input files |
| openpyxl | Writing formatted Excel output files |
| os | Folder creation and file listing |
| re | Regex-based position splitting and normalization |

---

## How to Run

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Place yearly Excel files in input_files/

# 3. Run
python collate_executives.py

# 4. Collect results from Output/
```

---

## Extending the System

### Add a new position variant
In `collate_executives.py`, find the `NORMALIZATION` dictionary and add:
```python
"your new title": "Standardized Position Name",
```

### Add a new compound position pattern
In the `normalize_position()` function, find the `pre_map` dictionary and add:
```python
r"pattern one \/ pattern two": "Resolved Title 1 / Resolved Title 2",
```

### Change input or output folder
At the top of `collate_executives.py`:
```python
INPUT_FOLDER = "input_files"   # change this
OUTPUT_FOLDER = "Output"       # change this
```

---

## Project Info

| Field | Detail |
|-------|--------|
| Organization | CACYOF OYSCATECH ALUMNI |
| Developer | Olakunle Sunday Olalekan |
| Company | Crevagroup / Clevatec |
| Language | Python 3 |
| Status | Stable — Production Ready |
