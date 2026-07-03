Attribute VB_Name = "modUtilities"
'==============================================================================
' MODULE 3 : modUtilities
'------------------------------------------------------------------------------
' PURPOSE
'   Small, reusable, dependency-free helper functions used by every other
'   module: safe type conversion, finding sheets/columns, reading ranges into
'   arrays, and logging warnings.
'
' DESIGN NOTE
'   None of these functions read cells one-at-a-time in a loop over the sheet.
'   We read a whole range into a Variant array ONCE and work in memory, because
'   touching the worksheet thousands of times is the #1 cause of slow VBA.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' LOGGING
'   We collect warnings in a module-level Collection instead of popping a
'   MsgBox for every problem (which would be unusable on a big workbook).
'   At the end we write them all to a Log sheet and summarise them once.
'------------------------------------------------------------------------------
Private mLog As Collection   ' holds warning strings for the whole run

Public Sub LogReset()
    ' Purpose : start a fresh, empty log at the beginning of a run.
    Set mLog = New Collection
End Sub

Public Sub LogWarning(ByVal message As String)
    ' Purpose : record a non-fatal problem and keep going.
    ' Input   : message - human-readable description of the issue.
    If mLog Is Nothing Then Set mLog = New Collection
    mLog.Add message
    ' Optionally echo to the Immediate window (Ctrl+G) for live debugging.
    If ECHO_WARNINGS_TO_IMMEDIATE Then Debug.Print "WARNING: " & message
End Sub

Public Function LogCount() As Long
    ' Purpose : how many warnings were recorded.
    If mLog Is Nothing Then
        LogCount = 0
    Else
        LogCount = mLog.Count
    End If
End Function

Public Function LogItem(ByVal index As Long) As String
    ' Purpose : read one warning by its 1-based position.
    If mLog Is Nothing Then Exit Function
    If index >= 1 And index <= mLog.Count Then LogItem = mLog.Item(index)
End Function

'------------------------------------------------------------------------------
' SAFE TYPE CONVERSION
'   Real sheets contain blanks, text where you expect numbers, and error cells.
'   These "Safe*" functions never crash: they return a sensible default instead.
'------------------------------------------------------------------------------

Public Function SafeDouble(ByVal value As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    ' Purpose : turn any cell value into a Double without ever raising an error.
    ' Returns : the number, or defaultValue if the cell is blank/text/error.
    On Error GoTo Fallback
    If IsError(value) Then GoTo Fallback           ' cell holds #N/A, #VALUE! etc.
    If IsEmpty(value) Then GoTo Fallback           ' truly empty cell
    If Trim$(CStr(value)) = "" Then GoTo Fallback  ' blank / spaces only
    If Not IsNumeric(value) Then GoTo Fallback     ' text that isn't a number
    SafeDouble = CDbl(value)
    Exit Function
Fallback:
    SafeDouble = defaultValue
End Function

Public Function SafeString(ByVal value As Variant, Optional ByVal defaultValue As String = "") As String
    ' Purpose : turn any cell value into a trimmed String safely.
    On Error GoTo Fallback
    If IsError(value) Then GoTo Fallback
    If IsEmpty(value) Then GoTo Fallback
    SafeString = Trim$(CStr(value))
    Exit Function
Fallback:
    SafeString = defaultValue
End Function

Public Function SafeDate(ByVal value As Variant, Optional ByVal defaultValue As Date = 0) As Date
    ' Purpose : turn any cell value into a Date safely.
    On Error GoTo Fallback
    If IsError(value) Then GoTo Fallback
    If IsEmpty(value) Then GoTo Fallback
    If Trim$(CStr(value)) = "" Then GoTo Fallback
    If Not IsDate(value) Then GoTo Fallback
    SafeDate = CDate(value)
    Exit Function
Fallback:
    SafeDate = defaultValue
End Function

Public Function SafeBoolean(ByVal value As Variant, Optional ByVal defaultValue As Boolean = False) As Boolean
    ' Purpose : interpret common truthy/falsey cell values as a Boolean.
    Dim s As String
    On Error GoTo Fallback
    If IsError(value) Or IsEmpty(value) Then GoTo Fallback
    s = LCase$(Trim$(CStr(value)))
    Select Case s
        Case "true", "yes", "y", "1", "-1", "on":  SafeBoolean = True
        Case "false", "no", "n", "0", "off", "":    SafeBoolean = False
        Case Else:                                   SafeBoolean = defaultValue
    End Select
    Exit Function
Fallback:
    SafeBoolean = defaultValue
End Function

'------------------------------------------------------------------------------
' WORKBOOK / WORKSHEET HELPERS
'------------------------------------------------------------------------------

Public Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    ' Purpose : True if a sheet with this name exists in the workbook.
    ' Method  : try to grab it; if that errors, it does not exist.
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    WorksheetExists = Not (ws Is Nothing)
End Function

Public Function IsExcludedSheet(ByVal sheetName As String) As Boolean
    ' Purpose : True if this sheet is a "system" sheet and must NOT be a client.
    ' Method  : case-insensitive match against the "|"-separated exclude list.
    Dim parts() As String
    Dim i As Long
    parts = Split(SHEET_EXCLUDE_LIST, "|")
    For i = LBound(parts) To UBound(parts)
        If StrComp(Trim$(parts(i)), Trim$(sheetName), vbTextCompare) = 0 Then
            IsExcludedSheet = True
            Exit Function
        End If
    Next i
    IsExcludedSheet = False
End Function

'------------------------------------------------------------------------------
' ARRAY / RANGE HELPERS
'------------------------------------------------------------------------------

Public Function ReadSheetToArray(ByVal ws As Worksheet) As Variant
    ' Purpose : read the sheet's UsedRange into a 2-D, 1-based Variant array
    '           in a SINGLE operation (fast), so we never touch cells in a loop.
    ' Returns : a 2-D array. If the sheet is empty, returns an empty (Empty) value
    '           which the caller must check with IsEmpty / IsArray.
    ' Note    : Range.Value on a single cell returns a scalar, not an array, so we
    '           normalise that edge-case into a 1x1 array for consistent handling.
    Dim rng As Range
    Dim result As Variant

    Set rng = ws.UsedRange
    If rng Is Nothing Then
        ReadSheetToArray = Empty
        Exit Function
    End If

    If rng.Cells.Count = 1 Then
        ' Single populated cell: build a 1x1 array so callers can always use (r, c).
        ReDim result(1 To 1, 1 To 1)
        result(1, 1) = rng.value
        ReadSheetToArray = result
    Else
        ReadSheetToArray = rng.value   ' bulk read -> 2-D array, 1-based indices
    End If
End Function

Public Function FindColumn(ByVal dataArray As Variant, ByVal headerName As String, _
                           Optional ByVal headerRow As Long = 1) As Long
    ' Purpose : find which COLUMN in the array holds a given header text.
    ' Inputs  : dataArray - a 2-D array read by ReadSheetToArray
    '           headerName - the header to look for (e.g. "Trade Code")
    '           headerRow  - which row in the array holds headers (default 1)
    ' Returns : the 1-based column index, or 0 if the header was not found.
    ' Matching: case-insensitive and trims surrounding spaces, so "Trade Code "
    '           still matches "Trade Code".
    Dim c As Long
    Dim cellText As String

    If Not IsArray(dataArray) Then
        FindColumn = 0
        Exit Function
    End If
    If headerRow < LBound(dataArray, 1) Or headerRow > UBound(dataArray, 1) Then
        FindColumn = 0
        Exit Function
    End If

    For c = LBound(dataArray, 2) To UBound(dataArray, 2)
        cellText = SafeString(dataArray(headerRow, c))
        If StrComp(cellText, Trim$(headerName), vbTextCompare) = 0 Then
            FindColumn = c
            Exit Function
        End If
    Next c

    FindColumn = 0   ' not found
End Function

Public Function LastRowOfArray(ByVal dataArray As Variant) As Long
    ' Purpose : the last row index of a 2-D array (0 if not an array).
    If IsArray(dataArray) Then LastRowOfArray = UBound(dataArray, 1) Else LastRowOfArray = 0
End Function

Public Function LastColOfArray(ByVal dataArray As Variant) As Long
    ' Purpose : the last column index of a 2-D array (0 if not an array).
    If IsArray(dataArray) Then LastColOfArray = UBound(dataArray, 2) Else LastColOfArray = 0
End Function

Public Function IsRowBlank(ByVal dataArray As Variant, ByVal r As Long) As Boolean
    ' Purpose : True if EVERY cell in row r is empty/blank. Such rows are skipped.
    Dim c As Long
    IsRowBlank = True
    For c = LBound(dataArray, 2) To UBound(dataArray, 2)
        If SafeString(dataArray(r, c)) <> "" Then
            IsRowBlank = False
            Exit Function
        End If
    Next c
End Function

'------------------------------------------------------------------------------
' DICTIONARY HELPERS
'   We use late binding (CreateObject) so you do NOT have to add the
'   "Microsoft Scripting Runtime" reference by hand - the project just works.
'------------------------------------------------------------------------------

Public Function NewDictionary() As Object
    ' Purpose : create a fresh, case-insensitive Scripting.Dictionary.
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1   ' 1 = vbTextCompare = case-insensitive keys
    Set NewDictionary = d
End Function

Public Sub SafeDictSet(ByVal dict As Object, ByVal key As String, ByVal value As Variant)
    ' Purpose : add or overwrite a key without ever raising a duplicate error.
    If dict.Exists(key) Then
        dict.Item(key) = value
    Else
        dict.Add key, value
    End If
End Sub

'------------------------------------------------------------------------------
' OUTPUT / FORMAT HELPERS
'------------------------------------------------------------------------------

Public Function TimeStampSuffix() As String
    ' Purpose : a filename-safe date/time stamp like "2026-07-03 14-05-32".
    TimeStampSuffix = Format$(Now, "yyyy-mm-dd hh-nn-ss")
End Function
