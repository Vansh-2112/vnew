Attribute VB_Name = "modMain"
'==============================================================================
' MODULE 2 : modMain
'------------------------------------------------------------------------------
' PURPOSE
'   The entry point and conductor. RunAutomation() is the button you press. It
'   validates the workbook, loads mappings, processes every client, then hands
'   the buckets to the Output Engine - all wrapped in one error handler so any
'   failure produces a friendly message and always restores Excel's settings.
'
' HOW TO RUN
'   1. Import all modules (see the README).
'   2. Open your "TRS Master Monitoring.xlsm".
'   3. Press Alt+F8, choose RunAutomation, click Run.
'==============================================================================
Option Explicit

Public Sub RunAutomation()
    ' Purpose : top-level workflow. This is the macro the user runs.
    Dim sourceWb As Workbook
    Dim clientSheets As Collection
    Dim ws As Worksheet
    Dim i As Long
    Dim outWb As Workbook

    ' --- Enter a fast, quiet state and set up the error trap -------------------
    On Error GoTo Fatal
    BeginFastMode
    LogReset                 ' start with an empty warning log
    AggregationReset         ' start with empty buckets

    ' --- STEP 1 : validate the source workbook --------------------------------
    Application.StatusBar = "TRS: validating workbook..."
    Set sourceWb = GetSourceWorkbook()
    If sourceWb Is Nothing Then
        EndFastMode
        MsgBox "No workbook is open to process. Please open your TRS Master " & _
               "Monitoring workbook and run again.", vbExclamation, "TRS Automation"
        Exit Sub
    End If

    ' --- STEP 2 : validate required sheets exist ------------------------------
    Application.StatusBar = "TRS: checking required sheets..."
    If Not ValidateRequiredSheets(sourceWb) Then
        EndFastMode
        MsgBox "Required sheet '" & SHEET_MAPPING & "' was not found in '" & _
               sourceWb.name & "'." & vbCrLf & vbCrLf & _
               "Open modConfig and set SHEET_MAPPING to your real tab name.", _
               vbCritical, "TRS Automation"
        Exit Sub
    End If

    ' --- STEP 3 : load mapping tables into memory -----------------------------
    Application.StatusBar = "TRS: loading mapping tables..."
    LoadAllMappings sourceWb

    ' --- STEP 4 : find and process every client sheet -------------------------
    Set clientSheets = GetClientSheets(sourceWb)
    If clientSheets.Count = 0 Then
        EndFastMode
        MsgBox "No client sheets were found. Every tab matched the exclude list." & _
               vbCrLf & "Check SHEET_EXCLUDE_LIST in modConfig.", _
               vbExclamation, "TRS Automation"
        Exit Sub
    End If

    For i = 1 To clientSheets.Count
        Set ws = clientSheets.Item(i)
        Application.StatusBar = "TRS: processing client " & i & " of " & _
                                clientSheets.Count & "  (" & ws.name & ")..."
        ProcessClient ws              ' reads sheet, feeds trades into aggregation
    Next i

    ' --- STEP 5 : build, format and save the output workbook ------------------
    Application.StatusBar = "TRS: generating output workbook..."
    Set outWb = GenerateOutputWorkbook(sourceWb)

    ' --- Restore Excel, then report the outcome -------------------------------
    EndFastMode

    If outWb Is Nothing Then
        MsgBox "Processing finished but the output workbook was not saved." & vbCrLf & _
               "See the Log for details (or it may have been cancelled).", _
               vbExclamation, "TRS Automation"
        Exit Sub
    End If

    MsgBox BuildSuccessMessage(clientSheets.Count), vbInformation, "TRS Automation - Done"
    Exit Sub

' --- Central error handler ----------------------------------------------------
Fatal:
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description
    EndFastMode              ' ALWAYS restore Excel, even on a crash
    MsgBox "The automation stopped due to an unexpected error." & vbCrLf & vbCrLf & _
           "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
           "No changes were made to your source workbook.", _
           vbCritical, "TRS Automation - Error"
End Sub

'------------------------------------------------------------------------------
' VALIDATION & SOURCE SELECTION
'------------------------------------------------------------------------------

Private Function GetSourceWorkbook() As Workbook
    ' Purpose : decide which open workbook to treat as the source.
    ' Rule    : use the ActiveWorkbook, but never the Personal Macro Workbook or
    '           an add-in. If nothing suitable is open, return Nothing.
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    If wb.name = "PERSONAL.XLSB" Then Exit Function
    Set GetSourceWorkbook = wb
End Function

Private Function ValidateRequiredSheets(ByVal wb As Workbook) As Boolean
    ' Purpose : confirm the sheets the engine cannot run without are present.
    ' Currently only the Mapping sheet is strictly required.
    ValidateRequiredSheets = WorksheetExists(wb, SHEET_MAPPING)
End Function

'------------------------------------------------------------------------------
' EXCEL STATE MANAGEMENT (speed + always restore)
'------------------------------------------------------------------------------

Private Sub BeginFastMode()
    ' Purpose : switch Excel into a fast, silent state for the run.
    Application.ScreenUpdating = False       ' don't redraw the screen
    Application.EnableEvents = False         ' don't fire Worksheet_Change etc.
    Application.Calculation = xlCalculationManual   ' don't recalc on every write
    Application.DisplayStatusBar = True      ' but DO show our progress text
End Sub

Private Sub EndFastMode()
    ' Purpose : restore Excel to normal. Safe to call more than once.
    Application.StatusBar = False            ' hand the status bar back to Excel
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'------------------------------------------------------------------------------
' SUCCESS MESSAGE
'------------------------------------------------------------------------------

Private Function BuildSuccessMessage(ByVal clientCount As Long) As String
    ' Purpose : a friendly summary of what happened.
    Dim msg As String
    msg = "TRS aggregation complete." & vbCrLf & vbCrLf
    msg = msg & "Clients processed : " & clientCount & vbCrLf
    msg = msg & "Output rows (client x asset) : " & BucketCount() & vbCrLf
    msg = msg & "Warnings logged : " & LogCount() & vbCrLf & vbCrLf
    If LogCount() > 0 Then
        msg = msg & "See the '" & OUTPUT_LOG_SHEET_NAME & "' tab in the new workbook " & _
                    "for warning details." & vbCrLf & vbCrLf
    End If
    msg = msg & "The results are in the '" & OUTPUT_SHEET_NAME & "' tab of the new workbook."
    BuildSuccessMessage = msg
End Function
