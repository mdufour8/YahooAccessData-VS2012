Namespace OptionValuation
	Public Interface IStockPriceVolatilityEstimateData
		ReadOnly Property NumberTradingDays As Double
		ReadOnly Property StockPrice As IPriceVol
		Property StockPriceNext As IPriceVol
		ReadOnly Property StockPriceStartValue As Double
		ReadOnly Property GainLog As Double
		ReadOnly Property Gain As Double
		ReadOnly Property GainDerivative As Double
		ReadOnly Property Volatility As Double
		ReadOnly Property VolatilityTotal As Double
		ReadOnly Property VolatilityTotalHigh As Double
		ReadOnly Property VolatilityTotalLow As Double
		ReadOnly Property ProbabilityOfInterval As Double
		ReadOnly Property IsBandExceeded As Boolean
		ReadOnly Property IsBandExceededHigh As Boolean
		ReadOnly Property IsBandExceededLow As Boolean
		ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		Sub Refresh()
		Sub Refresh(VolatilityDelta As Double)
		Sub Refresh(VolatilityDeltaLow As Double, VolatilityDeltaHigh As Double)
		Sub Reset()
	End Interface
End Namespace