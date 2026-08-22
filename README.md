# EasyClaim — AI-Assisted Insurance Claim Filing Platform

> A Robotic Process Automation (RPA) toolkit — with a broader AI-assisted vision — for automating medical insurance claim submission on government provider portals (NHA / ECHS / CGHS).

Repository: [`abhijeetrogye/easyclaim`](https://github.com/abhijeetrogye/easyclaim)

---

## Table of Contents

1. [About This Project](#about-this-project)
2. [Academic Context](#academic-context)
3. [What's Actually in This Repository](#whats-actually-in-this-repository)
4. [Repository Structure](#repository-structure)
5. [How the Automation Works](#how-the-automation-works)
6. [Script-by-Script Breakdown](#script-by-script-breakdown)
7. [Data Files](#data-files)
8. [Getting Started](#getting-started)
9. [Usage Workflow](#usage-workflow)
10. [Tech Stack](#tech-stack)
11. [The Larger Vision (Full Platform Design)](#the-larger-vision-full-platform-design)
12. [Functional & Non-Functional Requirements](#functional--non-functional-requirements)
13. [Reported Results (Academic Study)](#reported-results-academic-study)
14. [Known Limitations & Security Notes](#known-limitations--security-notes)
15. [Documentation Included in the Repo](#documentation-included-in-the-repo)
16. [Team / Authors](#team--authors)
17. [Roadmap](#roadmap)
18. [License](#license)

---

## About This Project

**EasyClaim** targets one of the most tedious parts of Indian government/PSU health insurance schemes (like **ECHS** — Ex-Servicemen Contributory Health Scheme, and **CGHS** — Central Government Health Scheme): manually re-typing the same patient, doctor, diagnosis, and procedure details into slow, legacy web portals for every single claim.

The code in this repository automates that data-entry step using **browser automation (Playwright)** driven by an **Excel sheet of claim entries**, so a bulk of claims can be filed against the [NHA Provider Portal](https://provider.nha.gov.in/) with minimal manual clicking.

Alongside the automation code, the repo also carries the **academic project report** ("blackbook") for a Bachelor of Engineering final-year project that frames EasyClaim as a much larger AI-assisted platform — OCR-based document extraction, NLP field-mapping, ML fraud detection, and a full web dashboard. That larger system is the long-term vision; the code currently in this repo is the **RPA layer** that proves out the portal-automation piece of it.

## Academic Context

This repository doubles as the codebase for a B.E. (Information Technology) final year project, submitted to the **University of Mumbai** via **St. John College of Engineering and Management, Palghar** (academic year 2025–2026), under the guidance of **Mr. Gautam Jha** (Assistant Professor).

The project report is titled *"Easy Claim – An AI Assisted Insurance Claim Filing Platform"* and documents motivation, literature review, requirements, UML/DFD design, and evaluation results for the full envisioned system (see [The Larger Vision](#the-larger-vision-full-platform-design) below).

## What's Actually in This Repository

Right now, this repo is primarily a working area for the **RPA/automation prototype**, not the full production platform. It contains:

- Playwright-based Python scripts that log into the NHA provider portal and fill in pre-authorization / claim forms field-by-field.
- A separate script for responding to **claim queries** (uploading supporting documents against claims that insurers have queried).
- Excel/CSV files used as the input "database" of claims to process, and as output logs of what succeeded or failed.
- Early exploratory/legacy scripts (under `ROUGH/`) from an earlier iteration of the automation, including a Playwright **Codegen recorder** transcript.
- The B.E. project report (`easyclaim - blackbook.pdf`) and a supporting Word document (`EasyClaim_ID_1133.docx`).

There is currently **no React frontend, FastAPI/Flask backend, OCR/NLP pipeline, database, or ML fraud-detection module in the repo** — those are part of the documented vision/design, not yet implemented code.

## Repository Structure

```
easyclaim/
├── EasyClaim_ID_1133.docx              # Supporting Word document (claim/reference doc)
├── easyclaim - blackbook.pdf           # Full B.E. project report ("blackbook")
│
├── QUERY UPLOAD/                       # Automation for responding to claim queries
│   ├── Query_upload.py                 # Main script: uploads docs for queried claims
│   ├── query_upload.xlsx               # Input sheet: registration IDs + doc paths + remarks
│   └── query_upload_backup.xlsx        # Auto-generated backup of the input sheet
│
└── ROUGH/
    └── tpa-auto-scripts-main/          # Earlier/experimental automation scripts
        ├── .gitignore
        ├── browser_launcher.py         # Opens a persistent, debuggable Chrome session
        ├── form_fill.py                # Main claim pre-authorization filler
        ├── script.py                   # Older/rougher end-to-end filler prototype
        ├── recorder_steps.py           # Playwright Codegen output (recorded login/nav steps)
        ├── entries.csv                 # Sample of a single claim entry, in CSV form
        ├── entries.xlsx                # Input sheet: claims to file (main dataset)
        ├── entries_with_paths.xlsx     # Variant of entries.xlsx with file paths resolved
        ├── output_results_*.xlsx       # Auto-generated run results (success/failure log)
        ├── requirements.txt            # Python dependencies
        └── readme.md                   # Original short setup note
```

## How the Automation Works

The core idea is a **two-process, human-in-the-loop RPA pattern**:

1. **`browser_launcher.py`** opens a real, persistent Chrome browser profile with remote debugging enabled (`--remote-debugging-port=9222`) and navigates to the NHA provider portal. A human logs in manually here (captcha + OTP make this a required manual step).
2. **`form_fill.py`** (or `Query_upload.py` / `script.py`) then **attaches to that already-logged-in browser** via `connect_over_cdp`, reads a spreadsheet of claims, and drives the UI: searching for a registration ID, expanding accordion sections, filling diagnosis/procedure/doctor fields, uploading supporting documents, and submitting for pre-authorization.
3. At a few points the script **pauses and waits for `Enter`**, handing control back to a human for steps that need manual judgement (e.g., completing "Admission Info", or verifying uploaded files) before continuing automatically.
4. Every row's outcome (`SUCCESS` / `FAILED` + error message) is written back out to a results spreadsheet, so a batch run can be audited or resumed.

This "launch browser separately, then attach and automate" pattern is deliberate — it avoids re-solving CAPTCHA/OTP challenges programmatically and lets a person stay in the loop for anything risky, while the repetitive form-filling is automated.

## Script-by-Script Breakdown

### `ROUGH/tpa-auto-scripts-main/browser_launcher.py`
Launches a persistent Chromium context (profile stored in `chrome-profile/`) with remote debugging exposed on port `9222`, and opens the NHA provider portal. Meant to be left running in its own terminal while you log in manually. Includes a reminder to delete `chrome-profile/SingletonLock` if the profile fails to reopen after a crash.

### `ROUGH/tpa-auto-scripts-main/form_fill.py`
The main, more mature automation script. It:
- Connects to the already-open browser from `browser_launcher.py`.
- Loads `entries.xlsx` (columns like `Registration ID*`, `Diagnose*`, `Doctor Name*`, `Doctor Id *`, `Procedures*`, `Amounts`, `Card File*`, `Patient Photo*`).
- For each row: searches the claim by registration ID, opens **Medical Information** (Personal History — allergies/habits), pauses for manual **Admission Info** entry, then automates **Treatment** — filling diagnosis, adding one or more procedures (falling back to a rotating list of "Unspecified Code" entries when a procedure isn't found in the portal's dropdown), uploading the CGHS card copy and patient photo, filling **Care Team Details**, and finally clicking through **Preview & Validate → Initiate Pre-Authorization → Confirm**.
- Writes a timestamped `output_results_<timestamp>.xlsx` summarizing success/failure per entry.

### `QUERY UPLOAD/Query_upload.py`
A separate, hardened automation for a different workflow: responding to claims that an insurer has **queried** (asking for more documents/clarification). It:
- Logs in manually (captcha/OTP), then navigates to the **Claims Queried** list.
- For each row in `query_upload.xlsx` (columns include a registration ID, a supporting-document path, and remarks): searches the claim, opens the **Query Response** panel, uploads the requested document, fills in remarks, and saves.
- Uses a `retry_on_failure` decorator (exponential backoff) around key UI actions, multiple fallback CSS/XPath selector strategies for finding claim cards and file inputs, periodic session-validity checks, and safe Excel saving (writes to a temp file, keeps a `_backup.xlsx`, and warns you to close Excel if the file is locked).
- The final "SUBMIT CLAIM" / confirmation step is deliberately commented out (`# TODO: Uncomment below after testing`) as a safety guard until the flow is verified.

### `ROUGH/tpa-auto-scripts-main/script.py`
An earlier, rougher version of the end-to-end filler (targets a different Excel file, `NEW PORTAL.xlsx`, and a `DataPDF` download directory). Useful as a reference for how the automation evolved, but superseded by `form_fill.py`.

### `ROUGH/tpa-auto-scripts-main/recorder_steps.py`
Raw output from **Playwright's Codegen recorder** — a literal transcript of a manual login + navigation session (captcha entry, OTP, searching a claim, opening Medical Information/Edit). It's not meant to be run as-is in production; it's a reference for the exact selectors the portal uses. *(Note: as recorded, this file contains example/test login and OTP values from a real session — treat it as a template to rebuild selectors from, not a script to execute unmodified.)*

## Data Files

| File | Purpose |
|---|---|
| `entries.xlsx` | Primary input: one row per claim to file, with registration ID, diagnosis, doctor info, procedures/amounts (pipe-`\|`-separated for multi-procedure claims), and document paths. |
| `entries_with_paths.xlsx` | Same data, with local file paths for supporting documents already resolved. |
| `entries.csv` | A single sample row in CSV form, useful for understanding the expected column layout. |
| `query_upload.xlsx` / `query_upload_backup.xlsx` | Input/backup for the "respond to queried claims" workflow — registration ID, supporting document path, remarks, and a `status` column the script updates as it runs. |
| `output_results_*.xlsx` | Auto-generated after each `form_fill.py` run — one row per claim with `Status` (SUCCESS/FAILED) and `Error` detail. |

## Getting Started

> These steps apply to the scripts under `ROUGH/tpa-auto-scripts-main/` (the most complete automation). Adjust paths for `QUERY UPLOAD/` similarly.

1. **Clone the repo**
   ```bash
   git clone https://github.com/abhijeetrogye/easyclaim.git
   cd easyclaim/ROUGH/tpa-auto-scripts-main
   ```

2. **Create and activate a virtual environment**
   ```bash
   python -m venv .venv
   # Windows
   .venv\Scripts\activate
   # macOS / Linux
   source .venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

4. **Prepare your data**
   Fill in `entries.xlsx` with your claim rows, matching the column headers the script expects (see [Data Files](#data-files)). Make sure any referenced document paths (card copy, patient photo) actually exist on disk.

## Usage Workflow

Run the two scripts in **separate terminals**, in this order:

```bash
# Terminal 1 — launch and stay logged into the portal
python browser_launcher.py
```
Complete the manual login (captcha + OTP) in the browser window that opens, then leave it running.

```bash
# Terminal 2 — run the automated filler against your open session
python form_fill.py
```
The script will process each row of `entries.xlsx`, pausing where indicated for manual verification (Admission Info, file-upload confirmation) before continuing. When it finishes, check the generated `output_results_*.xlsx` for a per-claim success/failure report.

For the query-response workflow, run `python "QUERY UPLOAD/Query_upload.py"` instead — it manages its own browser session end-to-end (login → process every unresolved row in `query_upload.xlsx` → summary).

## Tech Stack

**Currently implemented (this repo):**

| Layer | Technology |
|---|---|
| Browser automation | [Playwright](https://playwright.dev/python/) (sync API), Chromium |
| Data handling | `pandas`, `openpyxl` |
| Language | Python 3 |
| Input/output format | Excel (`.xlsx`) and CSV |
| Target system | [NHA Provider Portal](https://provider.nha.gov.in/) (ECHS/CGHS claim workflows) |

**Full stated project stack (from the academic report — not yet in this repo):**

| Layer | Technology |
|---|---|
| Frontend | React.js |
| Backend | Python — FastAPI or Flask |
| Database | PostgreSQL or MongoDB |
| Document OCR | AWS Textract |
| ML (fraud/anomaly detection) | Scikit-learn (Decision Trees, Logistic Regression) |
| Cloud storage | AWS S3 |
| Automation (production target) | Playwright or UiPath |
| Deployment | Docker / cloud VMs |

## The Larger Vision (Full Platform Design)

The accompanying B.E. project report frames EasyClaim as an end-to-end **AI-assisted insurance claim management system**, designed to address fragmented, manual, and non-transparent claim filing across insurers, hospitals, and third-party administrators. Per the report, the intended system would:

- Use **OCR (AWS Textract)** and **NLP** to extract data directly from unstructured medical bills and prescriptions, and map it into claim form fields — removing manual entry.
- Use **Robotic Process Automation (RPA)** — the piece this repo currently implements — to simulate a human filling out external portals like ECHS/CGHS.
- Apply **supervised ML models** (Logistic Regression, Decision Trees) to flag anomalous or potentially fraudulent claims.
- Offer a centralized web dashboard with real-time status tracking, notifications, and an audit trail, targeting policyholders, hospitals/insurance agents, and administrators as distinct user roles.
- Be designed for extensibility toward blockchain-based immutable claim records and predictive analytics on approval likelihood, per the report's stated future scope.

The design documentation (use-case, class, activity, sequence, and data-flow diagrams) for this full system lives in the project report rather than as code in this repository at this time.

## Functional & Non-Functional Requirements

As specified in the project report, the target system's requirements include:

**Functional:** user registration & authentication; claim initiation across categories (Health, Vehicle, Travel); AI-assisted form filling via OCR/NLP; secure multi-document upload tied to a unique Claim ID; automated portal submission via RPA; an AI guidance assistant; historical claim logs/audit trail; and automated document categorization.

**Non-functional:** ≥90% OCR extraction accuracy and ~90.7% validation accuracy targets; ~70% reduction in manual filing time; encrypted, access-controlled handling of sensitive medical/personal data; modular scalability for new insurers; robust error handling with backup/recovery; a simple, responsive UI; interoperability via standard APIs; and maintainable, update-friendly architecture.

**Suggested hardware baseline** (per the report): Intel i5+ CPU, 8GB+ RAM (16GB recommended), 256GB SSD, stable broadband, and a modern browser (Chrome/Firefox/Edge).

## Reported Results (Academic Study)

The project report documents evaluation against a simulated dataset of 500 claims across four categories (Cashless, Cashless Anywhere, Reimbursement With Financing, Cashless With Financing — 186 valid / 314 intentionally incomplete). Headline findings reported:

- **~84.8%** reduction in overall claim-processing time versus manual filing.
- **>90%** OCR extraction accuracy (AWS Textract) on unstructured medical documents.
- **~90.7%** validation accuracy from the ML-based anomaly detection models.
- **93%** of claims correctly completed on the first submission attempt.
- Ablation testing indicated the guided workflow and automated validation steps materially drove these accuracy/completion gains.

*(These figures describe the evaluated academic prototype/dataset as documented in the report, not necessarily production metrics of the code currently in this repo.)*

## Known Limitations & Security Notes

- **Manual login required.** Captcha and OTP steps are intentionally left manual; no credential automation or OTP-bypassing is implemented (nor should it be, for a government portal).
- **Brittle selectors.** Automation relies on the portal's current CSS classes, XPath structure, and hashed class names (e.g., `zKnrJTgkPSZ9htEXwleT`), which will break if the portal's frontend changes.
- **Hardcoded values.** Some scripts (`script.py`, parts of `form_fill.py`) contain hardcoded dates, contact numbers, and a single registration ID as leftovers from development/testing — check these before reusing the scripts for new data.
- **Sensitive data in plain files.** Excel inputs contain real-looking registration IDs, diagnoses, and doctor details, and `recorder_steps.py` is a literal recording of a login session. Treat all of these as sensitive and avoid committing real patient data to version control going forward (consider adding `*.xlsx` claim-data files to `.gitignore` alongside the existing `chrome-profile` exclusion).
- **No automated tests.** There is no test suite for the automation scripts; correctness currently relies on manual verification pauses built into the scripts themselves.
- **Duplicate/legacy code.** `script.py` and `form_fill.py` overlap significantly — `script.py` is the older, less robust version and could be removed or archived once `form_fill.py` is confirmed to fully replace it.

## Documentation Included in the Repo

- **`easyclaim - blackbook.pdf`** — the full 48-page B.E. project report: introduction, literature review, requirements analysis, UML/DFD design, proposed architecture, implementation notes, results & discussion, and conclusion.
- **`EasyClaim_ID_1133.docx`** — a supporting Word document (claim/reference-ID related).
- **`ROUGH/tpa-auto-scripts-main/readme.md`** — the original short setup note this comprehensive README supersedes; it noted creating a venv first and running `browser_launcher.py` and `form_fill.py` in separate terminals.

## Team / Authors

Final-year B.E. (Information Technology) project, University of Mumbai, 2025–2026 — St. John College of Engineering and Management, Palghar:

- **Arya Uday Kurup**
- **Abhijeet Pradip Rogye** ([@abhijeetrogye](https://github.com/abhijeetrogye))
- **Viraj Atul Tamhanekar**

Project guide: **Mr. Gautam Jha**, Assistant Professor, Department of Information Technology.

## Roadmap

Based on the gap between this repo's current scripts and the report's stated scope, natural next steps include:

- [ ] Extract the OCR/NLP form-mapping step so procedures/diagnoses can be auto-populated from an uploaded bill instead of a manually prepared Excel sheet.
- [ ] Replace hardcoded selectors with a more resilient locator strategy (data-testid attributes, or a small selector-config file) to reduce breakage when the portal updates.
- [ ] Build the FastAPI/Flask backend and React dashboard described in the report, with the current Playwright scripts becoming a background "submission worker."
- [ ] Add a proper test suite (even smoke tests against a staging/mock portal) before relying on this for real submissions.
- [ ] Formalize secrets/config handling (the `NHA_USERNAME` environment variable pattern in `Query_upload.py` is a good starting point — extend it to remove all hardcoded values from scripts).

## License

No license file is currently present in the repository. Until one is added, all rights are reserved by the authors — check with the repository owner before reusing or redistributing this code.
