# Vinzor Trade Lookup — Implementation Guide

Button-driven Excel tool that looks up multiple Trade IDs across a source
workbook made of **plain formatted ranges** (any sheets, any layout, nothing
hardcoded) and writes each trade's row in next to your input. Includes
counterparty code→name translation and shows the selected source path on-sheet.

---

## 1. Project structure

```
Vinzor_TradeLookup/
├── modTradeLookup.bas    ALL code (config, engine, 3 button macros, mapping)
└── ThisWorkbook_code.txt  Paste-only safety code for the ThisWorkbook object
```

One standard module + one ThisWorkbook snippet. No class modules, no UserForm
(the native file picker is used).

## 2. Required worksheets

- **Working sheet** (any name): where you paste Trade IDs and read results. The
  tool always acts on the **active sheet**.
- **`Mapping`** sheet (optional, for counterparty translation): two columns with
  headers `Counterparty Code` and `Counterparty Name`. If this sheet is absent,
  extraction still runs and simply skips translation.

## 3. Working-sheet layout

Place these label/header cells anywhere; the tool finds them dynamically.

- A cell containing exactly **`Trade ID`** — Trade IDs go beneath it, outputs
  auto-fill to its right.
- (Optional) a cell containing exactly **`Source File`** — the selected source
  path is written into the cell immediately to its right.

```
        E             F                         G          H        ...
1    Source File   C:\Data\Trades.xlsx
3                  Trade ID                    <output headers auto-fill>
4                  TR-1001
5                  TR-1002
6                  TR-1001   (duplicate: searched once, filled twice)
```

## 4. Mapping sheet layout

```
        A                   B
1   Counterparty Code   Counterparty Name
2   CP001               Acme Bank
3   CP002               Globex Capital
```

Columns need not be A/B and need not be adjacent — headers are located by name.
After extraction, any value in the **Counterparty** output column that matches a
code is replaced by its name; unmatched codes are left as-is.

## 5. Buttons and macros

| Button label      | Assigned macro          |
|-------------------|-------------------------|
| Browse Workbook   | `BrowseSourceWorkbook`  |
| Extract Trades    | `ExtractTrades`         |
| Reset Sheet       | `ResetSheet`            |

## 6. ThisWorkbook code

Paste `ThisWorkbook_code.txt` into the ThisWorkbook object (a Workbook_Open
safety net that restores Excel settings if a prior run crashed).

## 7. References / libraries

**None required.** Dictionaries use late binding; `FileDialog` and
`CustomDocumentProperties` are always available. Optional: Tools ▸ References ▸
*Microsoft Scripting Runtime* for IntelliSense only.

## 8. One-time setup

1. New workbook → **Alt+F11** → **File ▸ Import File…** → import
   `modTradeLookup.bas`.
2. Double-click **ThisWorkbook**, paste `ThisWorkbook_code.txt`.
3. On the working sheet add a `Trade ID` header (and optionally a `Source File`
   label).
4. (Optional) add a `Mapping` sheet with the two headers.
5. **Developer ▸ Insert ▸ Button (Form Control)** ×3, assign the macros in §5,
   label them.
6. **File ▸ Save As ▸ Excel Macro-Enabled Workbook (\*.xlsm)**.

## 9. How extraction works (flow)

Find `Trade ID` header → read IDs into dictionaries (duplicates → one search,
many fills) → open/reuse source → establish/detect output headers → scan every
`Barclays Trade ID` block in memory via arrays, with early exit once all IDs are
found → translate counterparty codes → one bulk write → report any not-found IDs.

## 10. Testing

1. Build a source workbook with plain ranges (no Ctrl+T) on two sheets, each a
   block with a `Barclays Trade ID` column in a non-first position, separated by
   blank rows/columns. Use unique IDs. Save & close.
2. **Browse Workbook** → select it (path appears next to `Source File`).
3. Enter IDs including one duplicate and one bogus ID → **Extract Trades**.
   First run creates output headers; duplicate row fills too; bogus ID reported.
4. Add a `Mapping` row for a counterparty code present in results → re-run:
   that column now shows the name.
5. **Reset Sheet** → IDs and outputs clear; headers, widths, buttons, formulas
   preserved.

## 11. Troubleshooting

| Symptom | Cause / fix |
|---|---|
| "Header 'Trade ID' not found" | Header text isn't exactly `Trade ID`. |
| "No source workbook selected" | Click **Browse Workbook** first. |
| "No range containing 'Barclays Trade ID'" | That exact header must exist in the source; check for stray/non-breaking spaces. |
| Some IDs not found | ID absent, or formatting mismatch (spaces / leading apostrophe). |
| Counterparty still shows a code | No `Mapping` sheet, missing headers, or no matching code row. |
| Path not shown | Add a `Source File` label cell; path saves regardless. |
| A block is cut short | Plain-range blocks must have no fully blank row inside and be separated from other blocks by a blank row and column (CurrentRegion rule). |
| Excel seems frozen after an error | Re-run any macro (restores settings) or reopen the workbook. |

## 12. Performance

Arrays not cells (one bulk read per block), O(1) dictionary lookups, duplicates
searched once, early exit on completion, single bulk write, no
Select/Activate/Copy-Paste, and ScreenUpdating/Events/Calc disabled then
restored.

## 13. Future enhancements

- FOUND/NOT-FOUND status column.
- Map multiple columns (e.g. Broker) via a list of target headers.
- Multi-source folders. Timestamped run log. Status-bar progress.
- Header-anchored scan to allow blank rows inside a block.
