Namespace OptionValuation
	''' <summary>
	''' This class encapsulate teh logic needed to test the stock price volatility prediction band via predicted price high and low.
	''' The class is used to combine together the stock price and the next stock price and calculate the High and Low in function
	''' of the Volatility calculation requirement. Note that the StockPriceNext is not mandatory to be set for the class creation since 
	''' usually it is still an unknown. Later on it can be addded to the same class via the StockPriceNext property.
	''' </summary>
	Public Class StockPriceDuplex
		Implements IStockPriceDuplex

		Dim MyStockPrice As IPriceVol
		Dim MyStockPriceNext As IPriceVol

		Dim MyTimePeriodInDay As Double
		Dim MyTimePeriodInDayLocal As Double
		Dim MyStockPriceStart As Double
		Dim MyStockPriceHigh As Double
		Dim MyStockPriceLow As Double

		Dim MyVolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		'


		''' <summary>
		''' This class is uses to combine together the stock price and the next stock price and calculate the High and Low in function
		''' of the Volatility calculation requirement. Note that the StockPriceNext is not mandatory to be set for the class cration since 
		''' usually it is still an unknown. Later on it can be addded to the same class via the StockPriceNext property.
		''' </summary>
		''' <param name="StockPrice">The stock price</param>
		''' <param name="TimePeriodInDay">The default time period in day. This parameter may be internally changed if the StockPriceNext 
		''' is not yet availaible. In that case the time period calculation is changed to the standard 8.00 hour daily trading period for an 
		''' intraday measurement range calculation.</param>
		''' <param name="volatilityPredictionBandType">The volatility prediction band type</param>
		''' <remarks></remarks>
		Public Sub New(
			ByVal StockPrice As IPriceVol,
			ByVal TimePeriodInDay As Double,
			ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)

			Me.New(StockPrice, Nothing, TimePeriodInDay, VolatilityPredictionBandType)
		End Sub

		Public Sub New(
			ByVal StockPrice As IPriceVol,
			ByVal StockPriceNext As IPriceVol,
			ByVal TimePeriodInDay As Double,
			ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)

			MyStockPrice = StockPrice
			MyStockPriceNext = StockPriceNext
			MyTimePeriodInDay = TimePeriodInDay
			MyTimePeriodInDayLocal = MyTimePeriodInDay 'by default the local time period for calculation is the same as the standard user sample time period
			MyVolatilityPredictionBandType = VolatilityPredictionBandType

			If MyStockPriceNext Is Nothing Then
				Select Case MyVolatilityPredictionBandType
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose
						'since we do not have the next sample calculate the on the open price and teh high and low intraday price	
						'in thatcase the time period need to be adjusted to the intraday trading period	
						MyTimePeriodInDayLocal = ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT
						MyStockPriceStart = MyStockPrice.Open
						MyStockPriceHigh = MyStockPrice.High
						MyStockPriceLow = MyStockPrice.Low
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen
						''calculate the threshold excess base on the last close since we do not have the next sample
						'MyTimePeriodInDayLocal =
						'MyStockPriceStart = MyStockPrice.LastPrevious
						'MyStockPriceHigh = MyStockPrice.Open
						'MyStockPriceLow = MyStockPrice.Open
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromOpenToClose
						'calculate the threshold excess base on the intraday stock price range
						'also need to reduice the daily period of trading to a inside daily period of trading
						MyTimePeriodInDay = ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT
						MyStockPriceStart = MyStockPrice.Open
						MyStockPriceHigh = MyStockPrice.High
						MyStockPriceLow = MyStockPrice.Low
				End Select
			Else
				Select Case MyVolatilityPredictionBandType
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose
						'calculate the threshold excess base on the last close since we do not have the next sample
						MyStockPriceStart = MyStockPrice.Last
						MyStockPriceHigh = MyStockPriceNext.High
						MyStockPriceLow = MyStockPriceNext.Low
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen
						'calculate the threshold excess base on the last close since we do not have the next sample
						MyStockPriceStart = MyStockPrice.Last
						MyStockPriceHigh = MyStockPrice.Open
						MyStockPriceLow = MyStockPrice.Open
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromOpenToClose
						'calculate the threshold excess base on the intraday stock price range
						'also need to reduice the daily period of trading to a inside daily period of trading
						MyTimePeriodInDay = ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT
						MyStockPriceStart = MyStockPrice.Open
						MyStockPriceHigh = MyStockPrice.High
						MyStockPriceLow = MyStockPrice.Low
				End Select
			End If
		End Sub


		Public ReadOnly Property StockPrice As IPriceVol Implements IStockPriceDuplex.StockPrice
			Get
				Return MyStockPrice
			End Get
		End Property

		Public Property StockPriceNext As IPriceVol Implements IStockPriceDuplex.StockPriceNext
			Get
				Return MyStockPriceNext
			End Get
			Set(value As IPriceVol)
				MyStockPriceNext = value
				Select Case MyVolatilityPredictionBandType
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose
						'calculate the threshold excess base on the last close since we do not have the next sample
						MyStockPriceStart = MyStockPrice.Last
						MyStockPriceHigh = MyStockPriceNext.High
						MyStockPriceLow = MyStockPriceNext.Low
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen
						'calculate the threshold excess base on the last close since we do not have the next sample
						MyStockPriceStart = MyStockPrice.Last
						MyStockPriceHigh = MyStockPrice.Open
						MyStockPriceLow = MyStockPrice.Open
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromOpenToClose
						'calculate the threshold exc ess base on the intraday stock price range
						'also need to reduice the daily period of trading to a inside daily period of trading
						MyTimePeriodInDay = ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT
						MyStockPriceStart = MyStockPrice.Open
						MyStockPriceHigh = MyStockPrice.High
						MyStockPriceLow = MyStockPrice.Low
				End Select
			End Set
		End Property

		Public ReadOnly Property TimePeriodInDay As Double Implements IStockPriceDuplex.TimePeriodInDay
			Get
				Return MyTimePeriodInDay
			End Get
		End Property

		Public ReadOnly Property StockPriceStart As Double Implements IStockPriceDuplex.StockPriceStart
			Get
				Return MyStockPriceStart
			End Get
		End Property

		Public ReadOnly Property StockPriceHigh As Double Implements IStockPriceDuplex.StockPriceHigh
			Get
				Return MyStockPriceHigh
			End Get
		End Property

		Public ReadOnly Property StockPriceLow As Double Implements IStockPriceDuplex.StockPriceLow
			Get
				Return MyStockPriceLow
			End Get
		End Property

		Public ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType Implements IStockPriceDuplex.VolatilityPredictionBandType
			Get
				Return MyVolatilityPredictionBandType
			End Get
		End Property
	End Class
End Namespace

