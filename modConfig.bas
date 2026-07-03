Attribute VB_Name = "modConfig"
'==============================================================================
' MODULE 1 : modConfig
'------------------------------------------------------------------------------
' PURPOSE
'   The ONE place you edit when the workbook structure changes.
'   Every worksheet name, every header name, every output column and every
'   default setting lives here. Nothing structural is hardcoded anywhere else.
'
' WHY A CONFIG MODULE?
'   Real workbooks change: someone renames a sheet, adds a column, or changes a
'   header. If those strings were scattered across the code you would have to
'   hunt through hundreds of lines. Keeping them in one module means one edit
'   fixes everything.
'
' HOW TO USE
'   Scroll down, change the text inside the quotes to match YOUR workbook,
'   and you are done. You should almost never edit any other module.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' SECTION A : WORKSHEET NAMES
'   These are the tabs in your TRS Master Monitoring workbook.
'   Change the text in quotes to match your real tab names EXACTLY
'   (spelling, spaces and capitalisation are matched case-insensitively,
'    but it is safest to match exactly).
'------------------------------------------------------------------------------

' The sheet that maps a Trade Code to its Underlying Asset,
' and (optionally) an Asset to its Current Price. See SECTION D for the
' header names inside this sheet.
Public Const SHEET_MAPPING As String = "Mapping"

' The Work-In-Progress input sheet. It is NOT a client, so we skip it.
Public Const SHEET_INPUT As String = "Inputs WIP"

' Any sheet whose name appears in this list is treated as a "system" sheet
' and is NEVER treated as a client. Separate names with a vertical bar "|".
' Add "Config", "Log", "Summary", etc. here if you have them.
' (Comparison is case-insensitive.)
Public Const SHEET_EXCLUDE_LIST As String = "Mapping|Inputs WIP|Inputs|Output|Config|Log|Summary|Dashboard|Notes"

' Should hidden sheets be treated as clients?
' False (recommended) = hidden sheets are skipped.
Public Const INCLUDE_HIDDEN_SHEETS As Boolean = False

'------------------------------------------------------------------------------
' SECTION B : CLIENT-SHEET HEADER NAMES
'   These are the COLUMN HEADERS the code looks for on each client tab.
'   The code finds columns by these names, so the column ORDER on the sheet
'   does not matter and can change without breaking anything.
'
'   If a header below does not exist on a client sheet, that field is simply
'   treated as blank/zero and a warning is logged (processing continues).
'------------------------------------------------------------------------------

Public Const HDR_TRADE_CODE As String = "Trade Code"        ' unique code per trade
Public Const HDR_NOTIONAL As String = "Notional"            ' trade notional value
Public Const HDR_QUANTITY As String = "Quantity"            ' number of units
Public Const HDR_STRIKE As String = "Strike"               ' strike level (e.g. 0.95)
Public Const HDR_DAY1_CASHOUT As String = "Day 1 Cash Out"  ' day-1 cash out amount
Public Const HDR_INDEP_AMOUNT As String = "Independent Amount" ' independent amount (IA)
Public Const HDR_CURRENT_PRICE As String = "Current Price"   ' price on the client sheet
Public Const HDR_MARKET_PRICE As String = "Market Price"     ' per-trade market price
' To read a NEW field from client sheets, add a Const here (one line) and then
' read it inside modClientProcessor (search for "ADD NEW FIELD HERE").

'------------------------------------------------------------------------------
' SECTION C : HEADER ROW POSITION
'   Which row on each sheet contains the headers. Almost always 1.
'------------------------------------------------------------------------------
Public Const HEADER_ROW As Long = 1

'------------------------------------------------------------------------------
' SECTION D : MAPPING SHEET HEADER NAMES
'   The column headers used INSIDE the Mapping sheet.
'   Mapping 1 : Trade Code  ->  Underlying Asset
'   Mapping 2 : Underlying Asset -> Current Price
'------------------------------------------------------------------------------
Public Const MAP_HDR_TRADE_CODE As String = "Trade Code"       ' key of mapping 1
Public Const MAP_HDR_ASSET As String = "Underlying Asset"      ' value of mapping 1
Public Const MAP_HDR_PRICE_ASSET As String = "Underlying Asset" ' key of mapping 2
Public Const MAP_HDR_PRICE As String = "Current Price"          ' value of mapping 2
' To add a NEW mapping (e.g. Asset -> Currency), add the two header consts here
' and load it in modMapping (search for "ADD NEW MAPPING HERE").

'------------------------------------------------------------------------------
' SECTION E : OUTPUT WORKBOOK SETTINGS
'------------------------------------------------------------------------------

' Base name for the generated output workbook (date/time is appended at save).
Public Const OUTPUT_WORKBOOK_BASENAME As String = "TRS Aggregated Output"

' Name of the results tab inside the output workbook.
Public Const OUTPUT_SHEET_NAME As String = "Aggregated Result"

' Name of the warnings/log tab inside the output workbook.
Public Const OUTPUT_LOG_SHEET_NAME As String = "Log"

' If True, the code prompts you for a save location.
' If False, it saves next to the source workbook automatically.
Public Const PROMPT_FOR_SAVE_LOCATION As Boolean = True

'------------------------------------------------------------------------------
' SECTION F : OUTPUT COLUMN HEADERS
'   The headers (and therefore the column order) of the output report.
'   IMPORTANT: This order MUST match the order that modOutput writes values in
'   (see BuildOutputRow in modOutput). If you add a column here, add the matching
'   value there. They are kept side-by-side and commented to make that easy.
'------------------------------------------------------------------------------
Public Function OutputHeaders() As Variant
    ' Returns the ordered list of output column headers as a 1-D array.
    OutputHeaders = Array( _
        "Client", _
        "Asset", _
        "Trade Count", _
        "Total Quantity", _
        "Total Notional", _
        "Total Independent Amount", _
        "Weighted Average Price", _
        "Weighted Average Strike", _
        "Weighted Average IA", _
        "Average Market Price", _
        "Current Price", _
        "Day 1 Cash Out Total", _
        "Max Cash Out", _
        "Exposure", _
        "Margin", _
        "Variation Margin", _
        "Current Exposure", _
        "Future Exposure")
End Function

'------------------------------------------------------------------------------
' SECTION G : CALCULATION SETTINGS (editable assumptions)
'   These feed the SAMPLE formulas in modCalculation. Change freely.
'------------------------------------------------------------------------------

' Weighting basis for weighted averages: "Notional" or "Quantity".
Public Const WEIGHT_BASIS As String = "Notional"

' Sample margin rate used by CalculateMargin (e.g. 0.1 = 10%). SAMPLE ONLY.
Public Const SAMPLE_MARGIN_RATE As Double = 0.1

' Sample future-exposure add-on factor used by CalculateFutureExposure. SAMPLE.
Public Const SAMPLE_FUTURE_EXPOSURE_FACTOR As Double = 0.15

'------------------------------------------------------------------------------
' SECTION H : BEHAVIOUR SETTINGS
'------------------------------------------------------------------------------

' If True, warnings are also printed to the VBA Immediate window (Ctrl+G).
Public Const ECHO_WARNINGS_TO_IMMEDIATE As Boolean = True
