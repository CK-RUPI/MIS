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
| 7 | `7. Collection Data Move.txt` | — | Reads FROM/TO table from `D:\Automate.xlsx` rows 173–195; copies raw downloaded files to working locations (overwrites) | A173:B195 |
| 8 | `8. Prepare Collection Report.txt` | — | Reads list of Excel files from rows 156–162 (two batches); opens each → refreshes Power Query (Ctrl+Alt+F5) → recalculates (F9) → runs VBA macro → saves & closes | A156:B162 |
| 9 | `9. Data Move for Collection Merge.txt` | — | Reads 2 file paths from rows 157–158; copies both to `D:\My work\Collection Count` | A157:A158 |
| 10 | `10. Update Merge Collection.txt` | — | Opens `MUL_USL_LAP-collection.xlsx` on OneDrive; refreshes Power Query (waits 30s); saves & closes. No macros. | — |
| 11 | `11. INSTA Collection Update.txt` | — | Copies files per FROM/TO rows 207–213; opens main merge file → refreshes → runs `SyncTableSize` macro; opens source files → refreshes each. **Time-gated: exits early if run after 11:30 AM.** | A207:B213 |
| 12 | `12. Collection Division.txt` | — | Opens `Insta-collection-Query.xlsm` → runs `SplitDataBySubBranch`; opens `MUL_USL_LAP-collection-UNPAID Query.xlsm` → runs `SplitDataBySubBranch`, `SplitDataBySubState`, `SplitDataBySubcluster`; moves output `.xlsx` files to OneDrive; runs `SendEmailUsingOutlook` macro | — |

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

| File | Size | Contents |
|------|------|----------|
| `Insta Collection-Main.xlsm` | 8.7 MB | Master Insta collection workbook; contains `SyncTableSize` macro and a "Master" sheet with table `Master` + `Table3` |
| `Insta Collection-Mar.xlsx` | 2.8 MB | Monthly Insta collection (March) |
| `LAP&USL Collection Due Report Mar26-Hozefa.xlsm` | 22 MB | LAP & USL collection due report; contains `AutomateLAPUSLReport` macro |
| `MUL Collection Due Report Mar26-Hozefa.xlsm` | 31 MB | MUL collection due report; contains `AutomateMULReport` macro |

### Sheet Structure Inside LAP&USL and MUL Workbooks
These sheets are the **actual dashboard views** that the macros produce:

| Sheet | What It Shows |
|-------|--------------|
| `Summary` | High-level collection KPIs |
| `Master` | Full raw collection data |
| `1-6 Unpaid` | Loans with 1–6 EMIs overdue |
| `Bucket` | DPD (Days Past Due) bucket-wise breakdown |
| `Emp wise` | Collection performance by employee |
| `Commitment vs Ach` | Target vs actual collection achievement |
| `Vintage Wise` | Collection performance by loan vintage |

---

---

## Flow Analysis

For detailed step-by-step breakdown of each flow, see **[FLOW_ANALYSIS.md](FLOW_ANALYSIS.md)**.

### Architecture Insight
The Power Automate flows are **orchestrators only** — they open Excel files, press Ctrl+Alt+F5 to refresh Power Query, and call VBA macros. The real transformation logic lives inside those macros.

### Pipeline Summary (Flows 7–12)
```
Raw downloaded files
       ↓
Flow 7:  Copy files to working paths (config-driven FROM/TO)
       ↓
Flow 8:  Refresh Power Query (×2) + run transformation macros
       ↓
Flow 9:  Copy 2 files to Collection Count folder
       ↓
Flow 10: Refresh master merged collection file (MUL_USL_LAP-collection.xlsx)
       ↓
Flow 11: Sync INSTA collection (morning-only, exits after 11:30 AM)
       ↓
Flow 12: Split by Branch / State / Cluster → distribute to OneDrive → send emails
```

### Key VBA Macros

| Macro | File | Reviewed? | Purpose |
|-------|------|-----------|---------|
| `SyncTableSize` | `Insta Collection-Main.xlsm` | Yes | Resizes `Table3` to match `Master` table row count after refresh |
| `AutomateLAPUSLReport` | `LAP&USL Collection Due Report Mar26-Hozefa.xlsm` | Yes | Creates delivery copy: copies 7 sheets, converts formulas to values, saves new xlsm |
| `AutomateMULReport` | `MUL Collection Due Report Mar26-Hozefa.xlsm` | Yes | Same as above for MUL data |
| `SplitDataBySubBranch` | `Insta-collection-Query.xlsm` & `MUL_USL_LAP-collection-UNPAID Query.xlsm` | No | Splits data by branch — files on D:\ not yet accessible |
| `SplitDataBySubState` | `MUL_USL_LAP-collection-UNPAID Query.xlsm` | No | Splits data by state |
| `SplitDataBySubcluster` | `MUL_USL_LAP-collection-UNPAID Query.xlsm` | No | Splits data by cluster |
| `SendEmailUsingOutlook` | `Send_Multiple_Email_Ver_2.0...xlsm` | No | Sends emails to branch/state/cluster managers |

---

## Status & Direction (as of 2026-03-19)

- **Flows 1–6** are being replaced by an AllCloud API (in progress, not yet delivered).
- **Flows 7–12** are the active focus — these need to be understood, automated, and used to feed an in-house dashboard.
- **Dashboard** has not been built yet. Tech stack and KPIs are not yet decided.
- **Macro code reviewed:** `SyncTableSize`, `AutomateLAPUSLReport`, `AutomateMULReport` — logic understood.
- **Macro code NOT yet reviewed:** `SplitDataBySubBranch`, `SplitDataBySubState`, `SplitDataBySubcluster`, `SendEmailUsingOutlook` — these are in files on `D:\My work\` which is not in the project folder. Need to copy those files into the project or paste macro code directly.
- **Dashboard views identified:** 7 sheet types per product — Summary, Master, 1-6 Unpaid, Bucket, Emp wise, Commitment vs Ach, Vintage Wise.

---

## Key Notes

- The `currentWorkflow/` folder is the **reference for the current process** — always compare any new automation against it.
- Download paths and URLs are not hardcoded in flows; they are driven by `D:\Automate.xlsx`.
- Flows 1–6 are the download phase (being replaced by API); flows 7–12 are the processing/transformation phase.
- The project is in active development — the dashboard has not been built yet.
