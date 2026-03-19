# Project Context — Digital Lending MIS

## Project Overview

**Digital Lending MIS** — a Management Information System for digital lending operations at Rupitol (AllCloud platform).

The goal is to fully automate the pipeline:
1. Download reports from the Rupitol web app
2. Process/transform them via Excel
3. Produce a dashboard from the final Excel outputs

---

## Platform

- **App:** Rupitol — `https://prod-ui-rupitol.allcloud.app/`
- **Login:** username `mis.rupitol` (password stored encrypted in Power Automate)
- **Report URLs** are all under `/Report/` on the same domain

---

## Folder Structure

```
MIS/
├── CLAUDE.md
├── PROJECT_CONTEXT.md                 ← this file
├── currentWorkflow/
│   └── Power Automade Flow/           ← 12 Power Automate Desktop flow files (.txt)
└── sourceOfTruth/
    └── Reports/                       ← final Excel outputs (source of truth for dashboard)
        ├── Insta Collection-Main.xlsm
        ├── Insta Collection-Mar.xlsx
        ├── LAP&USL Collection Due Report Mar26-Hozefa.xlsm
        └── MUL Collection Due Report Mar26-Hozefa.xlsm
```

---

## Control File

**`D:\Automate.xlsx`** — the central config file. Every flow reads a specific row range from this file to get:
- `URL` — the Rupitol report URL to open
- `On Date` / `From Date` / `To Date` — date filters for the report
- `Download Loaction` [sic] — the local file path to save the downloaded report

---

## Power Automate Flows (in execution order)

Each flow is a `.txt` file in `currentWorkflow/Power Automade Flow/`. They run sequentially, each calling the next via `External.RunFlow`.

| # | File | Internal Name | What It Does | Automate.xlsx Rows |
|---|------|---------------|--------------|-------------------|
| 1 | `1. Login.txt` | `I_Login` | Launches Chrome, logs into Rupitol, calls next flow | — |
| 2 | `2. Collection Status Report.txt` | `II_Download_CSR` | Downloads **Daily Collection Status Report** (`PLDailyCollectionStatusReport`) — needs "On Date" filter | A84:C94 |
| 3 | `3. LCC.txt` | `III_Download_LCC` | Downloads **LCC Report** (`MFILCCReport`) | A15:B19 |
| 4 | `4. DDR.txt` | `IV_Download_DDR` | Downloads **Detailed Due Report** (`DetailedDueReportForMFI`) | A8:B12 |
| 5 | `5. EMI Due Report.txt` | `V_Download_EDR` | Downloads **Loan EMI Due Report** (`MFIEMIDueReport`) | A59:B63 |
| 6 | `6. Loan Collection Report.txt` | `VI_Download_LCR` | Downloads **Loan Collection Report** — needs From Date + To Date; closes browser after; has time-of-day logic (before 9AM / after 6PM) | A23:D27 |
| 7 | `7. Collection Data Move.txt` | — | Moves/renames the downloaded raw files to working locations | — |
| 8 | `8. Prepare Collection Report.txt` | — | Opens Excel files and processes/transforms downloaded data | — |
| 9 | `9. Data Move for Collection Merge.txt` | — | Moves processed files into position for the merge step | — |
| 10 | `10. Update Merge Collection.txt` | — | Updates the merged collection Excel workbook | — |
| 11 | `11. INSTA Collection Update.txt` | — | Updates the Insta Collection Excel files | — |
| 12 | `12. Collection Division.txt` | — | Splits/divides the collection data (likely by product or branch) | — |

---

## Report Types Downloaded

| Report | Rupitol Endpoint | Filter |
|--------|-----------------|--------|
| Daily Collection Status Report | `/Report/PLDailyCollectionStatusReport` | On Date |
| LCC Report | `/Report/MFILCCReport` | (default) |
| Detailed Due Report | `/Report/DetailedDueReportForMFI` | (default) |
| Loan EMI Due Report | `/Report/MFIEMIDueReport` | (default) |
| Loan Collection Report | `/Report/` (LAP/USL, MUL variants) | From Date + To Date |

---

## Final Output Files (Source of Truth)

These are macro-enabled Excel files with formulas, located in `sourceOfTruth/Reports/`. They are the basis for the dashboard.

| File | Contents |
|------|----------|
| `Insta Collection-Main.xlsm` | Master Insta collection workbook |
| `Insta Collection-Mar.xlsx` | Monthly Insta collection (March) |
| `LAP&USL Collection Due Report Mar26-Hozefa.xlsm` | LAP & USL product collection dues |
| `MUL Collection Due Report Mar26-Hozefa.xlsm` | MUL product collection dues |

---

## Key Notes

- The `currentWorkflow/` folder is the **reference for the current process** — always compare any new automation against it.
- Download paths and URLs are not hardcoded in flows; they are driven by `D:\Automate.xlsx`.
- Flows 2–6 are the download phase; flows 7–12 are the processing/transformation phase.
- The project is in active development — the dashboard has not been built yet.
