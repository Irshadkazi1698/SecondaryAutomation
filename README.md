# SecondaryAutomation
This is a Secondary Automation Tool which can be use to Validate Counts, Bases and Sanity Checks

# Secondary QC Automation V8 - Project Documentation

## 1. Project Overview

This project is a desktop automation tool for secondary QC of banner/count outputs used in tabulation workflows.  
It validates alignment between Banner and Counts files, generates comparison outputs, performs sanity checks, and can optionally generate Grid Tables.

Primary application title:
- `Ipsos Banner Validation Tool` (PyQt5 desktop GUI)

---

## 2. Core Functional Scope

The tool performs the following major tasks:

1. Build/refresh matched variable mapping (`Matched_Variables.xlsx`)
2. Validate Banner vs Counts data at table/row level
3. Generate final comparison workbook with formula-based checks and color coding
4. Create unmatched title summary text output
5. Run sanity check module and append `Sanity Check` sheet
6. Optionally generate `Grid Tables` sheet

---

## 3. UI Specification

UI implementation file:
- `GUI.py`

### 3.1 Window & Layout

- Framework: `PyQt5`
- Main window title: `Ipsos Banner Validation Tool`
- Icon: `icon.png`
- Window behavior:
  - Responsive to screen size (`~85%` of available width/height)
  - Minimum size protection for usability
  - Scrollable content area (`QScrollArea`)
  - Maximized on launch

### 3.2 Main Sections

#### A. Files (GroupBox)
Input controls with `QLineEdit + Browse`:
- Input directory
- Output directory
- Banner file (name without extension)
- Counts file (name without extension)
- Numeric variable file (`.inc`) (name without extension)
- TabPlan file (name without extension)

#### B. TabPlan Category (GroupBox)
- Dropdown (`QComboBox`) options:
  - `1 - AutoPlan`
  - `2 - Tabplan`
  - `3 - Custom`
- `Custom Details` button is enabled only when `3 - Custom` is selected.

#### C. Custom TabPlan Dialog
Shown via `Custom Details`:
- Question column index
- Label column index
- Base text column index
- Sheet name

#### D. Grid Tables (Optional) (GroupBox)
- Checkbox: `Enable Grid Tables Generation`
- Grid Counts file picker (enabled only when checkbox is checked)

#### E. Execution & Feedback
- Progress bar (`0-100%`)
- Read-only log/output text area (`QTextEdit`)
- Run button (`Run`)

### 3.3 Validation & Interaction Rules

- Input and Output directory are mandatory
- If Grid option is enabled, Grid Counts file is mandatory
- If TabPlan category is Custom:
  - Custom details must be provided
  - Question/Label/Base indices must be integers
  - Sheet name is required
- Run button is disabled while background processing is active
- UI remains responsive via `QThread` worker execution

### 3.4 Progress Milestones (Approximate)

- 5%: Start validation
- 15-35%: Match file preparation and Banner QC automation
- 70%: Validation complete
- 80-90%: Sanity check
- 93-98%: Optional grid generation
- 100%: Completed

---

## 4. Functional Workflow (Execution Pipeline)

1. User provides file/folder names in GUI.
2. Worker thread creates `main.BannerValidation(...)`.
3. If `Matched_Variables.xlsx` is missing, run matching creation step.
4. Run banner validation automation.
5. Run sanity check:
   - TabPlan option `2`: `SanityCheckingTabPlan2`
   - Other options: `SanityChecking`
6. If enabled, run grid table generation.
7. Final logs and progress update in GUI.

---

## 5. Input Specifications

Input folder is expected to contain files referenced by base names entered in UI.

Typical required inputs:
- `<BannerFile>.xlsx`
- `<CountFile>.xlsx`
- `<NumericVarFile>.inc`
- `<TabPlanFile>.xlsm`
- `Var_name.txt` (used in matching file creation)

Optional input:
- `<GridCountsFile>.xlsx` (if Grid Tables enabled)

Output folder must be writable.

---

## 6. Output Specifications

Main generated artifacts:
- `Matched_df.xlsx`
- `Matched_Variables.xlsx`
- `Unmatched_Summary.txt`
- `Final Comparison.xlsx`
  - `Tables` sheet populated and formatted
  - comparison columns with formula checks
  - conditional formatting for matches/mismatches
  - `Sanity Check` sheet appended
  - optional `Grid Tables` sheet appended

---

## 7. Module Architecture

### 7.1 Entry/Orchestration
- `GUI.py`
  - UI layer
  - background thread orchestration
  - progress/log updates

- `main.py`
  - `BannerValidation` class
  - orchestrates matching and validation modules

### 7.2 Matching & Banner QC
- `DSCValidationAutomation/MatchingFileCreation.py`
  - validates dependencies/files/sheets
  - extracts variables from `.inc`
  - creates and enriches matched-variable outputs

- `DSCValidationAutomation/BannerQCAutomation.py`
  - preflight checks
  - table alignment and comparison
  - output workbook population
  - summary creation and formatting

### 7.3 Sanity Checks
- `SanityCheckModule/SanityChecking.py`
- `SanityCheckModule/SanityCheckingTabPlan2.py`
  - sigma checks
  - title/base checks
  - missing table checks
  - base size duplication checks
  - junk character checks

### 7.4 Grid Table Generation
- `GridTable/CreateGridTables.py`
  - derives grid variable datasets
  - creates `Grid Tables` sheet in final workbook

### 7.5 Counts Cleaning Utility
- `CountsCleaning.py`
- `CountsFileCleaning/CountsCleaning.py`
  - CSV counts cleanup and transformation
  - copies cleaned input files into `Input` folder

---

## 8. Technology Stack

### 8.1 Language & Runtime
- Python `3.x` (project currently running with CPython 3.10 bytecode in cache)

### 8.2 UI Framework
- `PyQt5`

### 8.3 Data Processing
- `pandas`
- `numpy`
- `re` (regex, stdlib)
- `collections` / `copy` (stdlib)

### 8.4 Matching Logic
- `fuzzywuzzy` (`fuzz`, `process`)

### 8.5 Excel Processing
- `openpyxl`

### 8.6 Encoding & CSV Handling
- `chardet`
- `csv` (stdlib)

### 8.7 Packaging / Distribution
- `PyInstaller`
  - Build command (from `Readme.txt`):
  - `python -m PyInstaller --onefile --icon=icon.png --noconsole GUI.py`
  - Produces executable artifacts in `dist/` (for example `GUI.exe` / `SV Tool.exe`)

---

## 9. Dependency List

From `requirements.txt`:
- pandas
- numpy
- fuzzywuzzy
- openpyxl
- chardet

Note:
- `PyQt5` is used by the GUI and should also be available in the runtime environment.

---

## 10. Run Instructions

### 10.1 Run from source
1. Install dependencies:
   - `pip install -r requirements.txt`
   - `pip install PyQt5`
2. Launch:
   - `python GUI.py`

### 10.2 Run packaged executable
- Use generated exe from `dist/` (for example `GUI.exe` or `SV Tool.exe`)

---

## 11. Error Handling & Validation Design

Validation checks are implemented in processing modules for:
- dependency availability
- input/output folder existence
- output folder write permission
- required files existence
- required Excel sheet presence
- required columns/index bounds (where applicable)
- critical content markers (for example variable prefix checks)

Errors are surfaced to logs and/or stop execution when critical.

---

## 12. Current Constraints / Assumptions

- GUI file selectors store base filename only; files are resolved inside selected input folder.
- Several modules assume fixed sheet names such as:
  - `Tables`, `Titles`, `Stub Specs`, `STUB SPECS`
- Matching and alignment quality depends on naming consistency and fuzzy matching thresholds.
- Some output file names are fixed by design (`Final Comparison`, `Matched_Variables`, etc.).

---

## 13. Repository Artifacts (High-level)

- Source scripts: root + module folders
- Input samples: `Input/`, `CountsInputFiles/`
- Output samples: `Output/`
- Build artifacts: `build/`, `dist/`
- Executables: `SV Tool.exe`, `dist/GUI.exe`

---

## 14. Recommended Documentation Next Additions

1. Data dictionary for each expected input file schema
2. Example run with screenshots of each UI section
3. Troubleshooting matrix (error -> likely cause -> resolution)
4. Versioned changelog and release notes

