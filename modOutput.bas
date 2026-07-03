Attribute VB_Name = "modOutput"
'==============================================================================
' MODULE 8 : modOutput  (Output Engine)
'------------------------------------------------------------------------------
' PURPOSE
'   Turn the finished buckets into a formatted report workbook:
'     1. build a single 2-D output array (headers + one row per bucket)
'     2. write it to a NEW workbook in ONE operation (fast)
'     3. format it (bold header, freeze row, number formats, autofit)
'     4. write a Log sheet of any warnings
'     5. save it (prompting for a location if configured)
'
' WHY BUILD AN ARRAY AND WRITE ONCE?
'   Writing cell-by-cell is extremely slow. We assemble everything in memory and
'   drop it onto the sheet in a single assignment: range.Value = array.
'==============================================================================
Option Explicit

Public Function GenerateOutputWorkbook(ByVal sourceWb As Workbook) As Workbook
    ' Purpose : create, fill, format and save the output workbook.
    ' Input   : sourceWb - the source workbook (used to derive a save folder).
    ' Returns : the new workbook (already saved), or Nothing on failure.
    Dim outWb As Workbook
    Dim wsOut As Worksheet
    Dim outArray As Variant

    ' 1) Build the full output array from the buckets.
    outArray = BuildOutputArray()

    ' 2) Create a brand-new one-sheet workbook.
    Set outWb = Application.Workbooks.Add(xlWBATWorksheet)   ' single blank sheet
    Set wsOut = outWb.Worksheets(1)
    wsOut.name = OUTPUT_SHEET_NAME

    ' 3) Write the whole array in a single shot.
    Dim nRows As Long, nCols As Long
    nRows = UBound(outArray, 1)
    nCols = UBound(outArray, 2)
    wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(nRows, nCols)).value = outArray

    ' 4) Format the result sheet.
    FormatResultSheet wsOut, nRows, nCols

    ' 5) Add the Log sheet (warnings collected during the run).
    WriteLogSheet outWb

    ' 6) Save, prompting for a location if configured.
    If Not SaveOutputWorkbook(outWb, sourceWb) Then
        Set GenerateOutputWorkbook = Nothing
        Exit Function
    End If

    Set GenerateOutputWorkbook = outWb
End Function

Private Function BuildOutputArray() As Variant
    ' Purpose : assemble headers + one data row per bucket into a 2-D array.
    ' Layout  : row 1 = headers; rows 2..n = data. Columns follow OutputHeaders().
    Dim headers As Variant
    Dim dict As Object
    Dim keys As Variant
    Dim nCols As Long, nRows As Long
    Dim result As Variant
    Dim i As Long, c As Long

    headers = OutputHeaders()                 ' 0-based 1-D array of header texts
    nCols = UBound(headers) - LBound(headers) + 1

    Set dict = Buckets()
    Dim bucketTotal As Long
    If dict Is Nothing Then bucketTotal = 0 Else bucketTotal = dict.Count

    ' +1 row for the header line. If there are zero buckets we still emit headers.
    nRows = bucketTotal + 1
    ReDim result(1 To nRows, 1 To nCols)

    ' Header row.
    For c = 1 To nCols
        result(1, c) = headers(LBound(headers) + c - 1)
    Next c

    ' Data rows.
    If bucketTotal > 0 Then
        keys = dict.keys
        Dim rowValues As Variant
        For i = 0 To UBound(keys)
            Dim b As clsAssetBucket
            Set b = dict.Item(keys(i))
            rowValues = BuildOutputRow(b)     ' a 1-based 1-D array, length nCols
            For c = 1 To nCols
                result(i + 2, c) = rowValues(c)
            Next c
        Next i
    End If

    BuildOutputArray = result
End Function

Private Function BuildOutputRow(ByVal b As clsAssetBucket) As Variant
    ' Purpose : produce ONE output row for a bucket.
    ' CRITICAL: the ORDER here MUST match modConfig.OutputHeaders() exactly.
    '           Each line is numbered to match its header for easy maintenance.
    Dim v() As Variant
    ReDim v(1 To 18)

    v(1) = b.ClientName                                   ' 1  Client
    v(2) = b.Asset                                        ' 2  Asset
    v(3) = b.TradeCount                                   ' 3  Trade Count
    v(4) = b.TotalQuantity                                ' 4  Total Quantity
    v(5) = b.TotalNotional                                ' 5  Total Notional
    v(6) = b.TotalIndependentAmount                       ' 6  Total Independent Amount
    v(7) = CalculateWeightedAveragePrice(b)               ' 7  Weighted Average Price
    v(8) = CalculateWeightedAverageStrike(b)              ' 8  Weighted Average Strike
    v(9) = CalculateWeightedAverageIA(b)                  ' 9  Weighted Average IA
    v(10) = CalculateAverageMarketPrice(b)                ' 10 Average Market Price
    v(11) = b.CurrentPrice                                ' 11 Current Price
    v(12) = b.Day1CashOutTotal                            ' 12 Day 1 Cash Out Total
    v(13) = CalculateMaxCashOut(b)                        ' 13 Max Cash Out
    v(14) = CalculateExposure(b)                          ' 14 Exposure
    v(15) = CalculateMargin(b)                            ' 15 Margin
    v(16) = CalculateVariationMargin(b)                   ' 16 Variation Margin
    v(17) = CalculateCurrentExposure(b)                   ' 17 Current Exposure
    v(18) = CalculateFutureExposure(b)                    ' 18 Future Exposure
    ' ADD NEW OUTPUT VALUE HERE (and a matching header in modConfig, and grow the
    ' ReDim above from 18 to the new count).

    BuildOutputRow = v
End Function

Private Sub FormatResultSheet(ByVal ws As Worksheet, ByVal nRows As Long, ByVal nCols As Long)
    ' Purpose : make the result sheet readable: bold header, freeze it, number
    '           formats on numeric columns, and autofit widths.
    Dim headerRng As Range
    Set headerRng = ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))

    ' Bold, coloured header row.
    With headerRng
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 120)     ' dark blue
        .Font.Color = RGB(255, 255, 255)       ' white text
        .HorizontalAlignment = xlCenter
    End With

    ' Freeze the header row so it stays visible while scrolling.
    ' Done without Select; wrapped so a freeze hiccup never aborts the run.
    On Error Resume Next
    ws.Activate
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    On Error GoTo 0

    ' Apply number formats to the numeric columns (columns 3..18 in this layout).
    If nRows >= 2 Then
        Dim dataRng As Range
        ' Integer-style column: Trade Count (col 3).
        Set dataRng = ws.Range(ws.Cells(2, 3), ws.Cells(nRows, 3))
        dataRng.NumberFormat = "#,##0"

        ' Amount/price columns: 4..18 -> two decimals with thousands separators.
        Set dataRng = ws.Range(ws.Cells(2, 4), ws.Cells(nRows, nCols))
        dataRng.NumberFormat = "#,##0.00"
    End If

    ' Autofit all used columns for a tidy result.
    ws.Columns.AutoFit
End Sub

Private Sub WriteLogSheet(ByVal wb As Workbook)
    ' Purpose : add a Log tab listing every warning captured during the run.
    Dim wsLog As Worksheet
    Dim n As Long, i As Long

    Set wsLog = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsLog.name = OUTPUT_LOG_SHEET_NAME

    wsLog.Cells(1, 1).value = "#"
    wsLog.Cells(1, 2).value = "Warning"
    wsLog.Range("A1:B1").Font.Bold = True

    n = LogCount()
    If n = 0 Then
        wsLog.Cells(2, 1).value = 1
        wsLog.Cells(2, 2).value = "No warnings. All data processed cleanly."
    Else
        Dim logArray() As Variant
        ReDim logArray(1 To n, 1 To 2)
        For i = 1 To n
            logArray(i, 1) = i
            logArray(i, 2) = LogItem(i)
        Next i
        wsLog.Range(wsLog.Cells(2, 1), wsLog.Cells(n + 1, 2)).value = logArray
    End If

    wsLog.Columns("A").ColumnWidth = 6
    wsLog.Columns("B").ColumnWidth = 110
End Sub

Private Function SaveOutputWorkbook(ByVal outWb As Workbook, ByVal sourceWb As Workbook) As Boolean
    ' Purpose : save the output workbook, optionally asking the user where.
    ' Returns : True on success, False if the user cancelled or a save error occurred.
    Dim baseName As String
    Dim defaultFolder As String
    Dim fullPath As String
    Dim chosen As Variant

    baseName = OUTPUT_WORKBOOK_BASENAME & " " & TimeStampSuffix() & ".xlsx"

    ' Default folder = wherever the source workbook lives (fallback to Documents).
    On Error Resume Next
    defaultFolder = sourceWb.Path
    On Error GoTo 0
    If defaultFolder = "" Then defaultFolder = Environ$("USERPROFILE") & "\Documents"

    If PROMPT_FOR_SAVE_LOCATION Then
        ' Ask the user for a location, pre-filling the default file name.
        chosen = Application.GetSaveAsFilename( _
                    InitialFileName:=defaultFolder & "\" & baseName, _
                    FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                    Title:="Save TRS Aggregated Output")
        If chosen = False Then
            ' User cancelled the dialog.
            LogWarning "Save cancelled by user; output workbook left open, unsaved."
            SaveOutputWorkbook = False
            Exit Function
        End If
        fullPath = CStr(chosen)
    Else
        fullPath = defaultFolder & "\" & baseName
    End If

    On Error GoTo SaveError
    Application.DisplayAlerts = False          ' suppress overwrite prompt
    outWb.SaveAs Filename:=fullPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    SaveOutputWorkbook = True
    Exit Function

SaveError:
    Application.DisplayAlerts = True
    LogWarning "Could not save output workbook to '" & fullPath & "': " & Err.Description
    SaveOutputWorkbook = False
End Function
