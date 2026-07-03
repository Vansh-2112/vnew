Attribute VB_Name = "modMapping"
'==============================================================================
' MODULE 4 : modMapping  (Mapping Engine)
'------------------------------------------------------------------------------
' PURPOSE
'   Turn the Mapping sheet into fast in-memory lookup tables (dictionaries):
'       Mapping 1 : Trade Code       -> Underlying Asset
'       Mapping 2 : Underlying Asset -> Current Price
'
' WHY DICTIONARIES?
'   A Scripting.Dictionary is a hash table: looking up a key is effectively
'   instant no matter how many entries it holds. If we instead searched the
'   sheet for every trade we'd re-scan thousands of rows over and over. We build
'   each lookup ONCE, then reuse it for every trade.
'
' EXTENSIBILITY
'   The core builder BuildLookup() is generic: give it a sheet, a key-header and
'   a value-header and it returns a dictionary. Adding "Asset -> Currency",
'   "Asset -> Counterparty", etc. later is one extra call (search for
'   "ADD NEW MAPPING HERE").
'==============================================================================
Option Explicit

' Module-level lookups, populated by LoadAllMappings and read by the processor.
Private mTradeToAsset As Object   ' key = Trade Code (String) -> value = Asset (String)
Private mAssetToPrice As Object   ' key = Asset (String)      -> value = Price (Double)

Public Sub LoadAllMappings(ByVal wb As Workbook)
    ' Purpose : build every lookup dictionary from the Mapping sheet.
    ' Input   : wb - the source workbook (already validated to contain the sheet).
    ' Output  : populates the module-level dictionaries above.
    Dim wsMap As Worksheet
    Dim mapArray As Variant

    Set wsMap = wb.Worksheets(SHEET_MAPPING)
    mapArray = ReadSheetToArray(wsMap)   ' read the mapping sheet once, into memory

    If Not IsArray(mapArray) Then
        LogWarning "Mapping sheet '" & SHEET_MAPPING & "' appears to be empty."
        Set mTradeToAsset = NewDictionary
        Set mAssetToPrice = NewDictionary
        Exit Sub
    End If

    ' Mapping 1 : Trade Code -> Underlying Asset (values are text).
    Set mTradeToAsset = BuildLookup(mapArray, MAP_HDR_TRADE_CODE, MAP_HDR_ASSET, False)

    ' Mapping 2 : Underlying Asset -> Current Price (values are numbers).
    Set mAssetToPrice = BuildLookup(mapArray, MAP_HDR_PRICE_ASSET, MAP_HDR_PRICE, True)

    ' ADD NEW MAPPING HERE:
    '   Set mAssetToCurrency = BuildLookup(mapArray, "Underlying Asset", "Currency", False)
    '   (declare a matching Private module variable and a public getter below).
End Sub

Private Function BuildLookup(ByVal dataArray As Variant, ByVal keyHeader As String, _
                             ByVal valueHeader As String, ByVal valueIsNumeric As Boolean) As Object
    ' Purpose : generic mapping builder. Reads two columns (found by header name)
    '           and returns a dictionary of key -> value.
    ' Inputs  : dataArray      - the mapping sheet as a 2-D array
    '           keyHeader      - header text of the key column
    '           valueHeader    - header text of the value column
    '           valueIsNumeric - True to store values as Double, False for String
    ' Behaviour on bad data (never fatal):
    '           * missing header  -> warn, return empty dictionary
    '           * blank key row   -> skipped
    '           * duplicate key   -> keep the FIRST value, warn about the clash
    Dim dict As Object
    Dim keyCol As Long, valCol As Long
    Dim r As Long
    Dim keyText As String
    Dim rawVal As Variant

    Set dict = NewDictionary

    keyCol = FindColumn(dataArray, keyHeader, HEADER_ROW)
    valCol = FindColumn(dataArray, valueHeader, HEADER_ROW)

    If keyCol = 0 Then
        LogWarning "Mapping header '" & keyHeader & "' not found on '" & SHEET_MAPPING & "'."
        Set BuildLookup = dict
        Exit Function
    End If
    If valCol = 0 Then
        LogWarning "Mapping header '" & valueHeader & "' not found on '" & SHEET_MAPPING & "'."
        Set BuildLookup = dict
        Exit Function
    End If

    ' Walk every data row (rows below the header row).
    For r = HEADER_ROW + 1 To UBound(dataArray, 1)
        keyText = SafeString(dataArray(r, keyCol))
        If keyText <> "" Then                       ' ignore blank-key rows
            rawVal = dataArray(r, valCol)
            If dict.Exists(keyText) Then
                ' Duplicate key: real workbooks sometimes repeat rows. We keep the
                ' first value (deterministic) and warn so you can clean the source.
                LogWarning "Duplicate mapping key '" & keyText & "' for '" & _
                           valueHeader & "' on '" & SHEET_MAPPING & "'. First value kept."
            Else
                If valueIsNumeric Then
                    dict.Add keyText, SafeDouble(rawVal)
                Else
                    dict.Add keyText, SafeString(rawVal)
                End If
            End If
        End If
    Next r

    Set BuildLookup = dict
End Function

'------------------------------------------------------------------------------
' PUBLIC LOOKUP GETTERS
'   The rest of the project asks questions through these, never by touching the
'   dictionaries directly. Each one degrades gracefully on a missing entry.
'------------------------------------------------------------------------------

Public Function LookupAsset(ByVal tradeCode As String, ByRef found As Boolean) As String
    ' Purpose : Trade Code -> Underlying Asset.
    ' Output  : found = True if the trade code was in the mapping.
    found = False
    If mTradeToAsset Is Nothing Then Exit Function
    If mTradeToAsset.Exists(tradeCode) Then
        LookupAsset = CStr(mTradeToAsset.Item(tradeCode))
        found = True
    End If
End Function

Public Function LookupPrice(ByVal asset As String, ByRef found As Boolean) As Double
    ' Purpose : Underlying Asset -> Current Price.
    ' Output  : found = True if a price existed for this asset.
    found = False
    If mAssetToPrice Is Nothing Then Exit Function
    If mAssetToPrice.Exists(asset) Then
        LookupPrice = SafeDouble(mAssetToPrice.Item(asset))
        found = True
    End If
End Function
