# Flow Analysis — Digital Lending MIS

Detailed breakdown of each Power Automate Desktop flow (flows 7–12).
Flows 1–6 are being replaced by the AllCloud API and are not analysed here.

---

## Flow 7 — Collection Data Move

**Purpose:** Copy raw downloaded files to their working locations.

- Opens `D:\Automate.xlsx`, reads rows A173:B195 (columns: `FROM`, `TO`)
- Loops through each row and copies file from `FROM` path to `TO` path (overwrites if exists)
- Purely a file move/rename step — no data transformation

---

## Flow 8 — Prepare Collection Report

**Purpose:** Refresh Power Query connections and run transformation macros on report workbooks.

### Mechanism
Flow 8 processes files in **two batches**, both configured via `D:\Automate.xlsx`.

#### Batch 1 — Rows 156–158 (up to 2 files)
For each file:
1. Open the Excel file
2. Wait 10s (let file load)
3. **Ctrl+Alt+F5** — refresh all Power Query connections
4. Wait 70s
5. **F9** — recalculate all formulas
6. Wait 10s
7. **Ctrl+Alt+F5 again** — second refresh (picks up derived queries dependent on first)
8. Wait 70s
9. **F9 again** — recalculate
10. Wait 10s
11. Save
12. Wait 20s
13. **Run VBA macro** (name from `MACRO` column)
14. Wait 12s
15. Save again
16. Wait 20s
17. Close

**Why double refresh?** The first refresh pulls raw data from source. Some Power Query steps depend on the output of earlier steps — the second refresh ensures those derived queries also update.

#### Batch 2 — Rows 159–162 (up to 4 files)
For each file:
1. Open the Excel file
2. Wait 5s
3. **Ctrl+Alt+F5** — single refresh only
4. Wait 15s
5. **F9** — recalculate
6. Wait 10s
7. **Run VBA macro**
8. Wait 10s
9. Save
10. Wait 10s
11. Close

#### Batch Comparison

| | Batch 1 (rows 156–158) | Batch 2 (rows 159–162) |
|---|---|---|
| Files | Up to 2 | Up to 4 |
| Refreshes | 2× | 1× |
| Wait after refresh | 70s each | 15s |
| Saves | 2× (before + after macro) | 1× |
| Likely files | Heavy/complex workbooks | Lighter workbooks |

### What the Macros Do (reviewed)
The macros triggered by Flow 8 (from `sourceOfTruth/Reports/`) are:

- **`AutomateLAPUSLReport`** (in `LAP&USL Collection Due Report Mar26-Hozefa.xlsm`):
  Creates a new workbook, copies sheets (Summary, 1-6 Unpaid, Commitment vs Ach, Bucket, Emp wise, Master, Vintage Wise), converts all formulas to values, saves as `LAP&USL Collection Due Report.xlsm`.

- **`AutomateMULReport`** (in `MUL Collection Due Report Mar26-Hozefa.xlsm`):
  Identical logic for MUL data. Saves as `MUL Collection Due Report.xlsm`.

### Known Gap
The actual file paths and macro names for rows 156–162 are stored in `D:\Automate.xlsx` — not yet read. We know the mechanism fully but not which specific files are in each batch.

---

## Flow 9 — Data Move for Collection Merge

**Purpose:** Stage 2 freshly-processed files into the merge working folder so Flow 10 can pick them up.

### Steps (6 lines total — simplest flow in the pipeline)
1. Open `D:\Automate.xlsx`
2. Read **cell A157** (single cell, not a range) → variable `File`
3. Read **cell A158** (single cell) → variable `File1`
4. Close `D:\Automate.xlsx`
5. Copy `File` → `D:\My work\Collection Count` (overwrite if exists)
6. Copy `File1` → `D:\My work\Collection Count` (overwrite if exists)

No loops, no macros, no Excel manipulation — purely a file copy.

### Key Observations
- **Rows 157–158 overlap with Flow 8 Batch 1.** Flow 8 reads A156:B158 to process files; Flow 9 reads A157 and A158 to copy those same files. The files Flow 8 just refreshed and ran macros on are immediately staged here.
- **Copies, not moves.** Original files remain at their source; only a copy goes to `Collection Count`.
- **`D:\My work\Collection Count` is an intermediate staging folder.** The master merge file in Flow 10 (`MUL_USL_LAP-collection.xlsx`) almost certainly has a Power Query connection pointing here.

### Pipeline Role
```
Flow 8:  Refreshed + ran macros on files at rows A157 and A158
              ↓
Flow 9:  Copies those 2 files → D:\My work\Collection Count
              ↓
Flow 10: Opens MUL_USL_LAP-collection.xlsx, refreshes Power Query
         (pulls from Collection Count via its connection)
```

---

## Flow 10 — Update Merge Collection

**Purpose:** Refresh the master merged workbook that consolidates LAP, MUL, and USL data into one file.

### Steps (7 lines — second simplest flow)
1. Open `C:\Users\Administrator\OneDrive - Rupitol Finance Private Limited\1. Collection\Collection Merged-LAP,MUL,USL\MUL_USL_LAP-collection.xlsx`
2. Wait 5s (let file load)
3. **Ctrl+Alt+F5** — refresh all Power Query connections
4. Wait 30s
5. Save
6. Wait 10s
7. Close

No macros. No F9 recalculation. No loops.

### Key Observations
- **Master merge file.** `MUL_USL_LAP-collection.xlsx` consolidates all three products (MUL, USL, LAP) into one workbook. Its Power Query connections point to the files Flow 9 just staged in `D:\My work\Collection Count`.
- **Single refresh only.** Unlike Flow 8 Batch 1 (double refresh), one pass is enough here — it's pulling already-processed data, not raw downloads.
- **No F9 recalculation.** Suggests the file has automatic calculation on, or formulas recalculate as part of the Power Query refresh.
- **File lives on OneDrive.** Once saved, it auto-syncs to the cloud — the team can access the merged data without manual sharing.
- **Shorter wait (30s vs 70s in Flow 8).** Reflects that this file pulls from clean local files, not complex raw queries.

### Pipeline Role
```
Flow 9:  Staged 2 processed files → D:\My work\Collection Count
              ↓
Flow 10: MUL_USL_LAP-collection.xlsx refreshes Power Query
         (pulls from Collection Count, merges LAP + MUL + USL)
         Saves → auto-syncs to OneDrive
              ↓
Flow 12: Splits the merged data by branch / state / cluster
```

Flow 10 is the **merge point** for LAP, MUL, and USL. After this, all three product lines are combined in one OneDrive file, ready for Flow 12 to split and distribute.

---

## Flow 11 — INSTA Collection Update

**Purpose:** Sync INSTA collection data. Morning-only (time-gated).

### Config Reads from Automate.xlsx (all in one open)

| Variable | Row(s) | What it is |
|----------|--------|-----------|
| `Movedata` | A207:B213 (with header) | FROM/TO file copy table — up to 6 raw INSTA download files to move to working paths |
| `File` | A219 (single cell) | Path to `Insta Collection-Main.xlsm` (INSTA master merge file) |
| `File1` | B218:B220 (column `EXCEL1`, with header) | Up to 3 source/feeder files that feed into the master |

### Phase 1 — File Copy Loop (lines 6–8)

```
FOR EACH row in Movedata:
    Copy FROM → TO (overwrite if exists)
```

Same pattern as Flow 7 — moves fresh INSTA download files to their working locations before processing.

### Phase 2 — Refresh Master Merge File (lines 9–17)

```
Open Insta Collection-Main.xlsm (path = File)
Ctrl+Alt+F5            → single Power Query refresh
Wait 60s               → heavy file (8.7MB), longer wait than source files
Run macro: SyncTableSize
F9                     → recalculate all formulas
Wait 10s
Save → Close
```

**Order matters:** `SyncTableSize` resizes `Table3` to match the `Master` table row count first, then F9 recalculates. If you ran F9 before resizing, formulas referencing rows not yet added to `Table3` would error.

**Single refresh only** — unlike Flow 8 Batch 1 (double refresh). The master is pulling from already-clean source files, so one pass is enough.

### Phase 3 — Refresh Source Files Loop (lines 18–29)

```
FOR EACH file in File1 (up to 3):
    Open file
    Wait 10s
    Ctrl+Alt+F5    → first refresh
    Wait 15s
    F9             → recalculate formulas  ← between the two refreshes (unusual)
    Ctrl+Alt+F5    → second refresh
    Wait 15s
    Save → Close
```

**F9 between two refreshes** is unique to this loop. The source files likely have calculated columns or formula-derived values that Power Query reads as data. The F9 forces those cells to update before the second refresh so Power Query picks up the freshly calculated values.

**Column header is `EXCEL1`** (not `EXCEL`) — minor naming difference from Flow 8's column headers.

### Phase 4 — Time Gate (lines 30–36)

```
Get current system time
Convert "11:30 AM" → ThresholdTime
Format both as HH:mm (24-hour)
IF current time > "11:30" THEN EXIT code 0
```

- Uses 24-hour string comparison: `"12:00" > "11:30"` = true (exits); `"09:00" > "11:30"` = false (continues to Flow 12)
- **Exit code 0** = clean success exit, not an error
- Morning-only by design: INSTA data comes in fresh every morning. Running after 11:30 AM would reprocess stale data pointlessly.
- **Flow 12 still runs regardless** — this gate only stops Flow 11 from processing; the pipeline continues downstream.

### Timing Estimate

| Step | Wait |
|------|------|
| 6 file copies | ~instant |
| Master refresh (60s) + macro + F9 + save | ~90s |
| 3 source files × (10 + 15 + 15 + 15)s | ~165s |
| **Total** | **~5–6 minutes** |

### Refresh Pattern Comparison

| | Flow 8 Batch 1 | Flow 11 Master | Flow 11 Sources |
|--|---|---|---|
| Refreshes | 2× | 1× | 2× |
| F9 between refreshes | No | N/A | **Yes** |
| Macro | Yes (from config) | `SyncTableSize` | No |
| Wait after refresh | 70s | 60s | 15s |

### INSTA Pipeline Architecture

Flow 11 makes clear that INSTA has a **fully separate pipeline** from LAP/MUL/USL:
- Own raw download files (copied in Phase 1)
- Own master merge file (`Insta Collection-Main.xlsm`)
- Own source/feeder files (up to 3)
- Own split macro (`SplitDataBySubBranch` on `Master_Insta` table) in Flow 12

The two pipelines (INSTA and LAP/MUL/USL) run independently through Flows 8–11 and converge only in **Flow 12** at the file distribution and email blast steps.

### `SyncTableSize` Macro (reviewed)
Located in `Insta Collection-Main.xlsm`, sheet "Master":
- Reads row count of table `Master`
- Resizes `Table3` to match that row count exactly
- Ensures table dimensions stay in sync after Power Query refresh so dependent formulas don't break

---

## Flow 12 — Collection Division

**Purpose:** Split collection data by branch/state/cluster, distribute files to OneDrive, send automated emails via Outlook. Final flow in the pipeline.

### Step 1 — INSTA Split (lines 1–7)

```
Open D:\My work\Collection Division - Product Wise\Insta-collection-Query.xlsm
Ctrl+Alt+F5          → single refresh
Wait 15s
Run macro: SplitDataBySubBranch
Clear filters on table: Master_Insta   ← clears after macro (macro uses filters to loop through branches)
Wait 10s
Close and Save (CloseAndSave — saves on close)
```

- **INSTA gets branch-split only** — no state or cluster split. Simpler dimension than LAP/MUL/USL.
- **`Master_Insta`** is the source table.
- Filter clear after macro resets the workbook view — the macro loops through branches by applying/clearing filters repeatedly; this final clear leaves it clean.

### Step 2 — MUL/USL/LAP Unpaid Split (lines 8–18)

```
Open D:\My work\Collection Division - Product Wise\MUL_USL_LAP-collection-UNPAID Query.xlsm
Ctrl+Alt+F5     → single refresh, Wait 30s

Run: SplitDataBySubBranch
Clear filters on ENTIRE WORKSHEET (not just a table)   ← different from INSTA and the other two macros

Run: SplitDataBySubState
Clear filters on table: Collection_Count

Run: SplitDataBySubcluster
Clear filters on table: Collection_Count

Wait 10s
Close and Save
```

**Filter clear differences are meaningful:**
- After `SplitDataBySubBranch` → worksheet-level clear (macro touches cells outside any named table, or affects multiple tables)
- After `SplitDataBySubState` and `SplitDataBySubcluster` → table-level clear on `Collection_Count` only

**Why filter clear between each macro?** Each macro applies a filter to `Collection_Count` to loop through unique values (branches/states/clusters). If you skip the clear, the next macro runs on already-filtered data → wrong or incomplete outputs.

**`Collection_Count`** is the source table for LAP/MUL/USL splits — same table name that Flow 10's merged file feeds into.

### Inferred Split Macro Behavior

We haven't read the macro code, but the calling pattern tells us the output structure:

| Macro | Source table | Split dimension | Output |
|-------|-------------|----------------|--------|
| `SplitDataBySubBranch` | `Master_Insta` / `Collection_Count` | Branch | One `.xlsx` per branch |
| `SplitDataBySubState` | `Collection_Count` | State | One `.xlsx` per state |
| `SplitDataBySubcluster` | `Collection_Count` | Cluster | One `.xlsx` per cluster |

Pattern: loop unique values in dimension column → filter → copy filtered rows to new workbook → save as `[DimensionValue].xlsx` → clear filter → next.

### Step 3 — File Distribution (lines 19–20)

```
Get all *.xlsx files from D:\My work\Collection Division - Product Wise
MOVE → C:\Users\Administrator\OneDrive - Rupitol Finance Private Limited\
        1. Collection\Collection Merged-LAP,MUL,USL\Collection Division - Product Wise
(overwrite existing)
```

- **`*.xlsx` not `*.xlsm`** — only the generated output files are moved. The source query files (`Insta-collection-Query.xlsm`, `MUL_USL_LAP-collection-UNPAID Query.xlsm`) are `.xlsm` and stay on `D:\My work\`.
- **Move, not copy** — files are removed from `D:\My work\` after transfer (unlike Flows 7 and 9 which copy).
- **OneDrive path → auto-syncs** — team members can access branch/state/cluster split files immediately after this step without manual sharing.

### Step 4 — Email Blast (lines 21–29)

```
Open C:\...\OneDrive\...\Send_Multiple_Email_Ver_2.0 - For Retailer Scheme -Updated.xlsm
Ctrl+Alt+F5     → first refresh, Wait 15s
Ctrl+Alt+F5     → second refresh (double refresh), Wait 15s
Activate sheet: "Formet"   ← typo in the original file name (not "Format")
Wait 5s
Run macro: SendEmailUsingOutlook
Close and Save
```

- **File is on OneDrive** — email config/template file is cloud-synced.
- **Double refresh** — the email workbook pulls from the just-moved split files (or from `MUL_USL_LAP-collection.xlsx`) to build recipient lists and email body data. Two passes ensure everything is current.
- **Sheet `Formet` activated before macro** — the macro reads from this sheet for email addresses, subject, body template. The activation is deliberate; the macro depends on being on this sheet.
- **`SendEmailUsingOutlook`** — uses Outlook COM automation. The machine must have Outlook installed and the sending account configured.
- **Recipients unknown** — likely branch managers, state managers, cluster managers (matching the three split dimensions). Need macro code to confirm.

### Macros Not Yet Reviewed
`SplitDataBySubBranch`, `SplitDataBySubState`, `SplitDataBySubcluster`, `SendEmailUsingOutlook` — all in files on `D:\My work\` which is outside the project folder.

### Complete Pipeline (Flows 7–12)

```
AllCloud API → raw downloaded files
        ↓
Flow 7:  Move raw files to working paths (config-driven FROM/TO)
        ↓
Flow 8:  Refresh PQ (×2) + run macros (LAP/MUL/USL transformation)
        ↓
Flow 9:  Copy 2 processed files → D:\My work\Collection Count
        ↓
Flow 10: Refresh MUL_USL_LAP-collection.xlsx → saves to OneDrive
        ↓
Flow 11: Move INSTA raw files → refresh Insta Collection-Main.xlsm
         + SyncTableSize macro (morning only; exits after 11:30 AM)
        ↓
Flow 12: [1] Insta-collection-Query.xlsm → SplitDataBySubBranch
              → INSTA branch files (*.xlsx)
         [2] MUL_USL_LAP-collection-UNPAID Query.xlsm
              → SplitDataBySubBranch (branch files)
              → SplitDataBySubState (state files)
              → SplitDataBySubcluster (cluster files)
         [3] Move all *.xlsx → OneDrive (auto-sync to team)
         [4] Refresh email workbook → SendEmailUsingOutlook
```

---

## Open Questions

1. What are the exact file paths in `D:\Automate.xlsx` rows 156–162? (Flow 8 batch files)
2. What does `SplitDataBySubBranch/State/Cluster` produce — one file per branch? One sheet per branch?
3. Who receives the emails from `SendEmailUsingOutlook`? What does the email contain?
4. What is in the `D:\My work\Collection Count` folder after Flow 9 copies files there?
