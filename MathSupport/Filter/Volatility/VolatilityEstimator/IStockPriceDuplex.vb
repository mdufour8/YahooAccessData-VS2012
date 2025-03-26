Namespace OptionValuation
	Interface IStockPriceDuplex
		ReadOnly Property StockPrice As IPriceVol
		Property StockPriceNext As IPriceVol
		ReadOnly Property TimePeriodInDay As Double

		ReadOnly Property StockPriceStart As Double
		ReadOnly Property StockPriceHigh As Double

		ReadOnly Property StockPriceLow As Double

		ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
	End Interface
End Namespace
