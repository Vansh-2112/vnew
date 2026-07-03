VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAssetBucket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' CLASS MODULE : clsAssetBucket
'------------------------------------------------------------------------------
' PURPOSE
'   Holds all the RUNNING TOTALS for one unique (Client, Asset) combination.
'   There is exactly one bucket per output row. As we loop through trades we
'   pour each trade's numbers into the matching bucket.
'
' WHY A CLASS (instead of a plain dictionary of values)?
'   A class gives every field a NAME and a TYPE. That makes the aggregation and
'   calculation code readable (bucket.TotalNotional instead of dict("TN")) and
'   makes it trivial to add a new field: add one Public line below and one line
'   in modAggregation. IntelliSense will then autocomplete it everywhere.
'
' IMPORTANT (accumulators only!)
'   This class stores RAW SUMS ONLY. It does NOT compute business results.
'   All business maths (weighted averages, exposure, margin, ...) lives in
'   modCalculation. Keeping accumulation and calculation separate is the single
'   most important design rule of this project.
'==============================================================================
Option Explicit

' --- Identity -----------------------------------------------------------------
Public ClientName As String          ' which client this bucket belongs to
Public Asset As String               ' which underlying asset this bucket is for

' --- Simple running totals ----------------------------------------------------
Public TradeCount As Long            ' how many trades landed in this bucket
Public TotalQuantity As Double       ' sum of Quantity across trades
Public TotalNotional As Double       ' sum of Notional across trades
Public TotalIndependentAmount As Double ' sum of Independent Amount (IA)
Public Day1CashOutTotal As Double    ' sum of Day 1 Cash Out

' --- Weighted-average numerators / denominators -------------------------------
'   A weighted average = numerator / denominator. We accumulate both parts here
'   and let modCalculation do the final division (so it can guard divide-by-zero).
Public WeightedPriceNumerator As Double    ' Sum of (price  * weight)
Public WeightedPriceDenominator As Double  ' Sum of (weight)
Public WeightedStrikeNumerator As Double   ' Sum of (strike * weight)
Public WeightedStrikeDenominator As Double ' Sum of (weight)

' --- For a simple (unweighted) average market price ---------------------------
Public SumMarketPrice As Double      ' sum of per-trade market price
Public MarketPriceCount As Long      ' how many trades actually had a market price

' --- Asset-level current price (comes from the Mapping / price table) ---------
Public CurrentPrice As Double        ' current price for this asset

' --- ADD NEW TOTALS HERE ------------------------------------------------------
'   To track a new number, add a Public line above this comment, then add the
'   accumulation line in modAggregation (search "ADD NEW FIELD HERE") and, if you
'   want it in the report, a calculation + output column. That is the whole loop.
