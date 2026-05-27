# 📋 CACYOF OYSCATECH ALUMNI — Executive Records Collation System

A Python automation tool that reads multiple yearly executive list Excel files and automatically reorganizes records by executive office/position, generating clean categorized reports for **CACYOF OYSCATECH ALUMNI**.

---

## 📌 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Project Structure](#project-structure)
- [Requirements](#requirements)
- [Installation](#installation)
- [How to Use](#how-to-use)
- [Input File Format](#input-file-format)
- [Output Structure](#output-structure)
- [Adding New Sessions](#adding-new-sessions)
- [Position Normalization](#position-normalization)
- [Technologies Used](#technologies-used)
- [Author](#author)

---

## Overview

This tool was built to solve a real administrative challenge: executive records spanning multiple years were stored in separate yearly Excel files, making it difficult to see who held a particular office across all sessions.

The system reads all yearly files, standardizes position titles, groups executives by office, and generates one clean Excel file per position — sorted chronologically from the earliest session to the latest.

---

## ✨ Features

- ✅ Reads multiple yearly Excel files automatically from a folder
- ✅ Extracts Name, Position, Phone Number, and Year Served
- ✅ Splits compound roles (e.g. `Vice President / Prayer Coordinator`) — person appears in **both** output files
- ✅ Normalizes position variants into a single standard title (e.g. `Ass. Bible Coordinator`, `Bible Coordinator II` → `Bible Coordinator`)
- ✅ Sorts records chronologically (2019 → 2020 → 2021 → ...)
- ✅ Generates one formatted Excel file per position
- ✅ Generates a master summary file with all records
- ✅ Professional Excel formatting with headers, alternating row colours, and borders
- ✅ Auto-creates output folder
- ✅ Easily expandable — just add new yearly files and re-run

---

## 📁 Project Structure

```
cacyof-executives/
│
├── collate_executives.py       # Main script
├── requirements.txt            # Python dependencies
├── README.md                   # This file
│
├── input_files/                # Put all yearly Excel files here
│   ├── EXECUTIVES__2019-2020_SESSION_.xlsx
│   ├── EXECUTIVES__2020-2021_SESSION_.xlsx
│   ├── EXECUTIVES__2021-2022_SESSION_.xlsx
│   ├── EXECUTIVES__2022-2023_SESSION_.xlsx
│   ├── EXECUTIVES__2023-2024_SESSION_.xlsx
│   └── EXECUTIVES__2024-2025_SESSION_.xlsx
│
└── Output/                     # Auto-generated output folder
    ├── President.xlsx
    ├── Vice_President.xlsx
    ├── General_Secretary.xlsx
    ├── Bible_Coordinator.xlsx
    ├── ... (one file per position)
    └── _SUMMARY_ALL_EXECUTIVES.xlsx
```

---

## 🛠 Requirements

- Python 3.9 or higher
- pip (Python package installer)

See `requirements.txt` for full dependency list.

---

## ⚙️ Installation

**1. Clone or download this repository**

```bash
git clone https://github.com/your-username/cacyof-executives.git
cd cacyof-executives
```

Or simply download and extract the ZIP file.

**2. Install dependencies**

```bash
pip install -r requirements.txt
```

---

## ▶️ How to Use

**1. Place your yearly Excel files in the `input_files/` folder**

Each file should represent one session (e.g. `EXECUTIVES__2024-2025_SESSION_.xlsx`).

**2. Run the script**

```bash
python collate_executives.py
```

**3. Find your results in the `Output/` folder**

One `.xlsx` file will be created for every unique position found across all sessions, plus a master summary file `_SUMMARY_ALL_EXECUTIVES.xlsx`.

---

## 📄 Input File Format

Each yearly Excel file should follow this general structure:

| Row | Content |
|-----|---------|
| Row 1 | Session year (e.g. `2024/2025 SESSION`) |
| Row 2 | Column headers (Name, Position, Phone, etc.) |
| Row 3+ | Executive records |

The script auto-detects the Name, Position, and Phone Number columns by scanning the header row — so minor variations in column order between files are handled automatically.

### Supported Position Column Formats

- Single position: `President`
- Compound position: `Vice President / Prayer Coordinator`
- Abbreviated: `Ass. Gen. Sec. / PRO`
- Shorthand: `Drama/Tech Cord`
- Numbered: `Bible Coordinator 2`, `Choir Coordinator 1`

---

## 📂 Output Structure

### Per-Position Files

Each file contains:

```
CACYOF OYSCATECH ALUMNI
[POSITION NAME]

S/N | Name | Year Served | Phone Number
 1  | ...  |  2019/2020  | ...
 2  | ...  |  2020/2021  | ...
```

Records are sorted from the earliest session to the most recent.

### Summary File (`_SUMMARY_ALL_EXECUTIVES.xlsx`)

Contains all executives from all positions and all years in one sheet, with an additional **Position** column.

---

## ➕ Adding New Sessions

When a new session is available:

1. Copy the new Excel file into the `input_files/` folder
2. Run the script again:
   ```bash
   python collate_executives.py
   ```

The `Output/` folder will be fully regenerated with the new session included and sorted in the correct chronological position.

> **Note:** The `Output/` folder is overwritten on every run. Always keep your source files in `input_files/`.

---

## 🔄 Position Normalization

The script uses a built-in normalization dictionary (`NORMALIZATION`) to merge similar position titles into one standard name.

### Examples

| Raw Position in File | Standardized Output |
|----------------------|---------------------|
| `Ass. Bible Coordinator` | `Bible Coordinator` |
| `Bible Coordinator II` | `Bible Coordinator` |
| `Bible Cord` | `Bible Coordinator` |
| `Asst. Gen. Sec. / PRO` | `Assistant General Secretary` + `PRO (Public Relations Officer)` |
| `Drama/Tech Cord` | `Drama Coordinator` + `Technical Coordinator` |
| `Vice President / Prayer Coordinator` | `Vice President` + `Prayer Coordinator` |

### Adding a New Position

Open `collate_executives.py` and find the `NORMALIZATION` dictionary. Add your entry:

```python
"new title variant": "Standardized Position Name",
```

Example:
```python
"youth coordinator": "Youth Coordinator",
"youth coord": "Youth Coordinator",
```

Save and re-run the script.

---

## 🧰 Technologies Used

| Tool | Purpose |
|------|---------|
| Python 3 | Core programming language |
| pandas | Reading and parsing Excel files |
| openpyxl | Writing formatted Excel output files |
| os / re | File system operations and text processing |

---

## 👤 Author

Built for **CACYOF OYSCATECH ALUMNI**

Developed by **Olakunle Sunday Olalekan**
Founder, [Crevagroup](https://crevagroup.com) | Clevatec

---

## 📃 License

This project is for internal administrative use by CACYOF OYSCATECH ALUMNI.
