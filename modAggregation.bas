Attribute VB_Name = "modAggregation"
'==============================================================================
' MODULE 6 : modAggregation  (Aggregation Engine)
'------------------------------------------------------------------------------
' PURPOSE
'   Owns the collection of buckets. There is ONE bucket per unique
'   (Client, Asset) pair, keyed as  ClientName & "|" & Asset.
'   As trades stream in, we find (or create) the right bucket and add the
'   trade's raw numbers into its running totals.
'
' WHAT THIS MODULE DOES NOT DO
'   It performs NO business calculations - only additions. Weighted averages,
'   exposure, margin etc. are computed later by modCalculation from the sums we
'   store here. Keeping "adding up" and "working out" apart is deliberate: it
'   means a formula change never risks corrupting the aggregation, and vice versa.
'==============================================================================
Option Explicit

' The master store of buckets. key = "Client|Asset", value = clsAssetBucket.
Private mBuckets As Object

Public Sub AggregationReset()
    ' Purpose : empty the store before a new run.
    Set mBuckets = NewDictionary
End Sub

Public Function BucketCount() As Long
    ' Purpose : how many unique (client, asset) rows we have built.
    If mBuckets Is Nothing Then BucketCount = 0 Else BucketCount = mBuckets.Count
End Function

Public Function Buckets() As Object
    ' Purpose : expose the store (read-only intent) so modOutput can iterate it.
    Set Buckets = mBuckets
End Function

Private Function BucketKey(ByVal clientName As String, ByVal asset As String) As String
    ' Purpose : the single, consistent key format used everywhere.
    BucketKey = clientName & "|" & asset
End Function

Private Function GetOrCreateBucket(ByVal clientName As String, ByVal asset As String) As clsAssetBucket
    ' Purpose : return the bucket for this (client, asset), creating it if new.
    Dim k As String
    Dim b As clsAssetBucket

    If mBuckets Is Nothing Then Set mBuckets = NewDictionary
    k = BucketKey(clientName, asset)

    If mBuckets.Exists(k) Then
        Set b = mBuckets.Item(k)            ' reuse the existing bucket
    Else
        Set b = New clsAssetBucket          ' first trade for this pair
        b.ClientName = clientName
        b.Asset = asset
        mBuckets.Add k, b
    End If

    Set GetOrCreateBucket = b
End Function

Public Sub AddTradeToBucket(ByVal clientName As String, ByVal asset As String, _
                            ByVal quantity As Double, ByVal notional As Double, _
                            ByVal strike As Double, ByVal independentAmount As Double, _
                            ByVal day1CashOut As Double, ByVal marketPrice As Double, _
                            ByVal hasMarketPrice As Boolean, ByVal currentPrice As Double)
    ' Purpose : pour ONE trade's numbers into the correct (client, asset) bucket.
    ' Inputs  : the trade's fields, already safely converted to Double by the caller.
    '           hasMarketPrice - True if this trade actually supplied a market price
    '                            (so the simple average only counts real values).
    '           currentPrice   - asset-level current price (same for the whole bucket).
    Dim b As clsAssetBucket
    Dim weight As Double

    Set b = GetOrCreateBucket(clientName, asset)

    ' --- Simple running totals ------------------------------------------------
    b.TradeCount = b.TradeCount + 1
    b.TotalQuantity = b.TotalQuantity + quantity
    b.TotalNotional = b.TotalNotional + notional
    b.TotalIndependentAmount = b.TotalIndependentAmount + independentAmount
    b.Day1CashOutTotal = b.Day1CashOutTotal + day1CashOut

    ' --- Weighted-average building blocks -------------------------------------
    ' The weight is the "importance" of each trade in the average. We support
    ' weighting by Notional or by Quantity (set in modConfig, WEIGHT_BASIS).
    If StrComp(WEIGHT_BASIS, "Quantity", vbTextCompare) = 0 Then
        weight = quantity
    Else
        weight = notional           ' default and recommended for TRS
    End If

    ' Weighted average price = Sum(price*weight) / Sum(weight).
    ' We use the trade's market price as the "price" being averaged; if the trade
    ' had no market price we fall back to the asset current price so the weight is
    ' still represented.
    Dim priceForAvg As Double
    If hasMarketPrice Then priceForAvg = marketPrice Else priceForAvg = currentPrice
    b.WeightedPriceNumerator = b.WeightedPriceNumerator + priceForAvg * weight
    b.WeightedPriceDenominator = b.WeightedPriceDenominator + weight

    ' Weighted average strike = Sum(strike*weight) / Sum(weight).
    b.WeightedStrikeNumerator = b.WeightedStrikeNumerator + strike * weight
    b.WeightedStrikeDenominator = b.WeightedStrikeDenominator + weight

    ' --- Simple average market price ------------------------------------------
    If hasMarketPrice Then
        b.SumMarketPrice = b.SumMarketPrice + marketPrice
        b.MarketPriceCount = b.MarketPriceCount + 1
    End If

    ' --- Asset-level current price --------------------------------------------
    ' Same for every trade of an asset; we just store the latest non-zero value.
    If currentPrice <> 0 Then b.CurrentPrice = currentPrice

    ' --- ADD NEW FIELD HERE ---------------------------------------------------
    '   Example: b.TotalFees = b.TotalFees + fees
    '   (Declare the field in clsAssetBucket, add "fees" as a parameter above and
    '    read it in modClientProcessor.)
End Sub
