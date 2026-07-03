Attribute VB_Name = "modClientProcessor"
'==============================================================================
' MODULE 5 : modClientProcessor  (Client Processor)
'------------------------------------------------------------------------------
' PURPOSE
'   Decide which tabs are "clients", then read each client tab into memory and
'   turn every valid trade row into a call to the Aggregation Engine.
'
' KEY IDEAS
'   * A client sheet is any visible tab that is NOT on the exclude list.
'   * Each sheet is read into a Variant array in ONE go (fast); we never read
'     cells individually inside the row loop.
'   * Columns are located by HEADER NAME (via FindColumn), so column order on the
'     sheet can change freely without breaking anything.
'==============================================================================
Option Explicit

Public Function GetClientSheets(ByVal wb As Workbook) As Collection
    ' Purpose : return a Collection of the worksheets that count as clients.
    ' Rule    : keep every sheet that is not excluded and (unless configured
    '           otherwise) not hidden.
    Dim result As New Collection
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If Not IsExcludedSheet(ws.name) Then
            If ws.Visible = xlSheetVisible Or INCLUDE_HIDDEN_SHEETS Then
                result.Add ws
            End If
        End If
    Next ws

    Set GetClientSheets = result
End Function

Public Sub ProcessClient(ByVal ws As Worksheet)
    ' Purpose : read one client sheet and feed all its trades into aggregation.
    ' Input   : ws - a worksheet already confirmed to be a client.
    Dim data As Variant
    Dim clientName As String

    clientName = ws.name
    data = ReadSheetToArray(ws)     ' single bulk read into memory

    If Not IsArray(data) Then
        LogWarning "Client sheet '" & clientName & "' is empty - skipped."
        Exit Sub
    End If

    ' --- Locate all needed columns ONCE, by header name -----------------------
    Dim colTrade As Long, colNotional As Long, colQty As Long, colStrike As Long
    Dim colDay1 As Long, colIA As Long, colCurPrice As Long, colMktPrice As Long

    colTrade = FindColumn(data, HDR_TRADE_CODE, HEADER_ROW)
    colNotional = FindColumn(data, HDR_NOTIONAL, HEADER_ROW)
    colQty = FindColumn(data, HDR_QUANTITY, HEADER_ROW)
    colStrike = FindColumn(data, HDR_STRIKE, HEADER_ROW)
    colDay1 = FindColumn(data, HDR_DAY1_CASHOUT, HEADER_ROW)
    colIA = FindColumn(data, HDR_INDEP_AMOUNT, HEADER_ROW)
    colCurPrice = FindColumn(data, HDR_CURRENT_PRICE, HEADER_ROW)
    colMktPrice = FindColumn(data, HDR_MARKET_PRICE, HEADER_ROW)

    ' Trade Code is the one column we truly cannot work without.
    If colTrade = 0 Then
        LogWarning "Client sheet '" & clientName & "': header '" & HDR_TRADE_CODE & _
                   "' not found - sheet skipped."
        Exit Sub
    End If

    ' Warn (but continue) for helpful-but-optional columns that are missing.
    If colNotional = 0 Then LogWarning "'" & clientName & "': '" & HDR_NOTIONAL & "' missing - treated as 0."
    If colQty = 0 Then LogWarning "'" & clientName & "': '" & HDR_QUANTITY & "' missing - treated as 0."
    If colStrike = 0 Then LogWarning "'" & clientName & "': '" & HDR_STRIKE & "' missing - treated as 0."

    ' --- Walk every data row --------------------------------------------------
    Dim r As Long
    Dim tradeCode As String, asset As String
    Dim foundAsset As Boolean, foundPrice As Boolean
    Dim qty As Double, notional As Double, strike As Double
    Dim ia As Double, day1 As Double
    Dim curPrice As Double, mktPrice As Double, hasMkt As Boolean

    For r = HEADER_ROW + 1 To UBound(data, 1)

        ' Skip completely blank rows silently (very common at the bottom of sheets).
        If Not IsRowBlank(data, r) Then

            tradeCode = SafeString(data(r, colTrade))

            If tradeCode = "" Then
                ' A row with data but no trade code is unusable for mapping.
                LogWarning "'" & clientName & "' row " & r & ": blank Trade Code - row skipped."
            Else
                ' 1) Trade Code -> Asset (via the mapping engine).
                asset = LookupAsset(tradeCode, foundAsset)
                If Not foundAsset Then
                    LogWarning "'" & clientName & "' row " & r & ": no asset mapping for Trade Code '" & _
                               tradeCode & "' - row skipped."
                Else
                    ' 2) Read the trade's numeric fields safely (missing col -> 0).
                    notional = ColValueDouble(data, r, colNotional)
                    qty = ColValueDouble(data, r, colQty)
                    strike = ColValueDouble(data, r, colStrike)
                    ia = ColValueDouble(data, r, colIA)
                    day1 = ColValueDouble(data, r, colDay1)

                    ' 3) Market price: only "present" if the column exists and has a value.
                    hasMkt = False
                    mktPrice = 0#
                    If colMktPrice > 0 Then
                        If SafeString(data(r, colMktPrice)) <> "" Then
                            mktPrice = SafeDouble(data(r, colMktPrice))
                            hasMkt = True
                        End If
                    End If

                    ' 4) Current price: prefer the value on the client sheet; if the
                    '    sheet has none, fall back to the Asset->Price mapping table.
                    curPrice = 0#
                    If colCurPrice > 0 Then
                        If SafeString(data(r, colCurPrice)) <> "" Then curPrice = SafeDouble(data(r, colCurPrice))
                    End If
                    If curPrice = 0# Then
                        curPrice = LookupPrice(asset, foundPrice)
                        If Not foundPrice Then
                            LogWarning "'" & clientName & "' row " & r & ": no Current Price for asset '" & _
                                       asset & "' (sheet or mapping). Using 0."
                        End If
                    End If

                    ' ADD NEW FIELD HERE:
                    '   Dim fees As Double
                    '   fees = ColValueDouble(data, r, FindColumn(data, "Fees", HEADER_ROW))
                    '   ...and pass it to AddTradeToBucket.

                    ' 5) Hand the trade to the Aggregation Engine.
                    AddTradeToBucket clientName, asset, qty, notional, strike, _
                                     ia, day1, mktPrice, hasMkt, curPrice
                End If
            End If

        End If
    Next r
End Sub

Private Function ColValueDouble(ByVal data As Variant, ByVal r As Long, ByVal col As Long) As Double
    ' Purpose : read a numeric cell from the array, returning 0 if the column
    '           does not exist (col = 0) or the cell is blank/non-numeric.
    If col = 0 Then
        ColValueDouble = 0#
    Else
        ColValueDouble = SafeDouble(data(r, col))
    End If
End Function
