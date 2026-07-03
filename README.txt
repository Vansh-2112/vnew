TRS MASTER MONITORING — AGGREGATION AUTOMATION (VBA)
====================================================

WHAT IT DOES
  Reads your "TRS Master Monitoring.xlsm", aggregates every client tab down to
  ONE ROW PER (CLIENT x ASSET), runs all business calculations, and writes a
  formatted output workbook with a results tab and a warnings Log tab.

FILES (import all 8 into the VBA editor)
  clsAssetBucket.cls     class module — the per (client,asset) bucket
  modConfig.bas          all editable names/headers/settings  <-- EDIT THIS ONE
  modMain.bas            entry point: RunAutomation
  modUtilities.bas       safe conversions, array/dict/log helpers
  modMapping.bas         Trade Code -> Asset, Asset -> Price lookups
  modClientProcessor.bas reads client tabs, feeds trades to aggregation
  modAggregation.bas     the buckets and how trades pour in
  modCalculation.bas     ALL business formulas (one function each)
  modOutput.bas          builds/formats/saves the output workbook

HOW TO IMPORT
  1. Open your TRS workbook in Excel.
  2. Press Alt+F11 to open the VBA editor.
  3. File > Import File... and import ALL 8 files above (import the .cls too).
     (Or drag them onto the Project window.)
  4. Press Alt+F11 to go back to Excel.

HOW TO RUN
  1. Make sure your TRS workbook is the active workbook.
  2. Press Alt+F8, choose RunAutomation, click Run.
  3. If prompted, pick where to save the output workbook.
  4. Read the success message; open the Log tab if there were warnings.

WHAT TO EDIT BEFORE FIRST RUN  (all in modConfig.bas)
  * SHEET_MAPPING        -> your mapping tab name (default "Mapping")
  * SHEET_EXCLUDE_LIST   -> pipe-separated list of non-client tabs
  * HDR_* constants      -> your client-sheet column headers
  * MAP_HDR_* constants  -> the header names inside the Mapping tab
  Columns are found BY HEADER NAME, so column order on the sheet does not matter.

REPLACING THE SAMPLE FORMULAS
  Open modCalculation.bas and search for  SAMPLE FORMULA - REPLACE
  Each such function is a single self-contained calculation you can swap out.
  The PROVIDED FORMULAs (weighted price/strike/IA, max cash out) are already
  implemented exactly as specified.

EXTENDING IT LATER (three small edits)
  1. clsAssetBucket.cls     : add a Public field   ("ADD NEW TOTALS HERE")
  2. modAggregation.bas     : accumulate it        ("ADD NEW FIELD HERE")
  3. modCalculation + modConfig.OutputHeaders + modOutput.BuildOutputRow
     if you want it as a new report column.

NOTES
  * No Select/Activate/Copy/Paste on data; everything runs in memory via arrays
    and Scripting.Dictionary for speed.
  * Late binding is used for dictionaries, so you do NOT need to add any VBA
    references by hand — it just works.
  * The source workbook is never modified; a brand-new .xlsx is produced.
