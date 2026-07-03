Attribute VB_Name = "modCalculation"
'==============================================================================
' MODULE 7 : modCalculation  (Calculation Engine)
'------------------------------------------------------------------------------
' PURPOSE
'   The ONLY place business maths happens. Each function takes a finished bucket
'   (a clsAssetBucket full of raw sums) and returns ONE result. Nothing here
'   changes the bucket - these are pure "read the sums, return a number" helpers.
'
' WHY ISOLATED?
'   When a formula needs to change, you edit exactly one small function here and
'   nothing else in the project can break. Aggregation stays untouched.
'
' FORMULAS
'   * Where you supplied a formula it is implemented exactly and marked
'     ' PROVIDED FORMULA.
'   * Where no formula was given, a reasonable, executable SAMPLE is implemented
'     and marked ' SAMPLE FORMULA - REPLACE. Search for that tag to find every
'     place you should drop in production maths later.
'
' DIVIDE-BY-ZERO
'   Every division goes through SafeDivide() so an empty bucket can never crash
'   the run; it returns 0 instead.
'==============================================================================
Option Explicit

Private Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double) As Double
    ' Purpose : divide without ever raising a "Division by zero" error.
    If denominator = 0# Then
        SafeDivide = 0#
    Else
        SafeDivide = numerator / denominator
    End If
End Function

'------------------------------------------------------------------------------
' PROVIDED FORMULAS
'------------------------------------------------------------------------------

Public Function CalculateWeightedAveragePrice(ByVal b As clsAssetBucket) As Double
    ' PROVIDED FORMULA
    '   WeightedAveragePrice = WeightedPriceNumerator / WeightedPriceDenominator
    CalculateWeightedAveragePrice = SafeDivide(b.WeightedPriceNumerator, b.WeightedPriceDenominator)
End Function

Public Function CalculateWeightedAverageStrike(ByVal b As clsAssetBucket) As Double
    ' PROVIDED FORMULA
    '   WeightedAverageStrike = WeightedStrikeNumerator / WeightedStrikeDenominator
    CalculateWeightedAverageStrike = SafeDivide(b.WeightedStrikeNumerator, b.WeightedStrikeDenominator)
End Function

Public Function CalculateWeightedAverageIA(ByVal b As clsAssetBucket) As Double
    ' PROVIDED FORMULA
    '   WeightedAverageIA = TotalIndependentAmount / TradeCount
    '   (This is the simple average of IA per trade, exactly as specified.)
    CalculateWeightedAverageIA = SafeDivide(b.TotalIndependentAmount, CDbl(b.TradeCount))
End Function

Public Function CalculateMaxCashOut(ByVal b As clsAssetBucket) As Double
    ' PROVIDED FORMULA (aggregated form)
    '   Per trade you gave:  MaxCashOut = Day1CashOut + ((1 - Strike) * Notional)
    '   At the aggregated (client, asset) level we use the bucket totals and the
    '   weighted-average strike:
    '       MaxCashOut = Day1CashOutTotal + ((1 - WeightedAvgStrike) * TotalNotional)
    Dim was As Double
    was = CalculateWeightedAverageStrike(b)
    CalculateMaxCashOut = b.Day1CashOutTotal + ((1# - was) * b.TotalNotional)
End Function

'------------------------------------------------------------------------------
' SAMPLE FORMULAS  (executable now, replace with production maths later)
'------------------------------------------------------------------------------

Public Function CalculateAverageMarketPrice(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   Simple mean of the per-trade market prices that were actually present.
    '   AverageMarketPrice = SumMarketPrice / MarketPriceCount
    CalculateAverageMarketPrice = SafeDivide(b.SumMarketPrice, CDbl(b.MarketPriceCount))
End Function

Public Function CalculateExposure(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   A basic mark-to-market style exposure:
    '   Exposure = (CurrentPrice - WeightedAveragePrice) * TotalQuantity
    Dim wap As Double
    wap = CalculateWeightedAveragePrice(b)
    CalculateExposure = (b.CurrentPrice - wap) * b.TotalQuantity
End Function

Public Function CalculateCurrentExposure(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   Current (positive-only) exposure: never counts a gain as exposure.
    '   CurrentExposure = Max(0, Exposure)
    Dim e As Double
    e = CalculateExposure(b)
    If e < 0# Then e = 0#
    CalculateCurrentExposure = e
End Function

Public Function CalculateFutureExposure(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   Potential future exposure as a flat add-on to notional:
    '   FutureExposure = TotalNotional * SAMPLE_FUTURE_EXPOSURE_FACTOR
    CalculateFutureExposure = b.TotalNotional * SAMPLE_FUTURE_EXPOSURE_FACTOR
End Function

Public Function CalculateMargin(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   Total margin requirement = Independent Amount + a % of current exposure.
    '   Margin = TotalIndependentAmount + (CurrentExposure * SAMPLE_MARGIN_RATE)
    CalculateMargin = b.TotalIndependentAmount + (CalculateCurrentExposure(b) * SAMPLE_MARGIN_RATE)
End Function

Public Function CalculateVariationMargin(ByVal b As clsAssetBucket) As Double
    ' SAMPLE FORMULA - REPLACE
    '   Variation margin tracks mark-to-market move; here we equate it to exposure.
    '   VariationMargin = Exposure
    CalculateVariationMargin = CalculateExposure(b)
End Function

' ADD NEW CALCULATION HERE:
'   Public Function CalculateXyz(ByVal b As clsAssetBucket) As Double
'       CalculateXyz = ...   ' read only from b.* fields
'   End Function
'   Then add a matching output column in modConfig.OutputHeaders and a matching
'   value in modOutput.BuildOutputRow.
