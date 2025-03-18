
Namespace OptionValuation
	Public Class StockPriceVolatilityEstimateData
		Implements IStockPriceVolatilityEstimateData

		Private Const VOLATILITY_TOTAL_MINIMUM As Double = 0.01

		Private MyNumberTradingDays As Double
		Private MyStockPrice As IPriceVol
		Private MyStockPriceNext As IPriceVol
		Private MyStockPriceStartValue As Double
		Private MyStockPriceHighValue As Double
		Private MyStockPriceLowValue As Double
		Private MyGain As Double
		Private MyGainDerivative As Double
		Private MyVolatility As Double
		Private MyVolatilityTotal As Double
		Private MyProbabilityHigh As Double
		Private MyProbabilityLow As Double
		Private MyProbabilityOfInterval As Double
		Private MyVolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		Private _IsBandExceeded As Boolean
		Private _IsBandExceededHigh As Boolean
		Private _IsBandExceededLow As Boolean
		Private MyGainLog As Double


		''' <summary>
		''' This class hold the data for the stock price prediction run. 
		''' It also calculate the stock price threshold prediction data and calculate if the data fall inside a given probabilty value for the next period.	
		''' </summary>
		''' <param name="NumberTradingDays">Usually one days</param>
		''' <param name="StockPrice"></param>
		''' <param name="Gain"></param>
		''' <param name="GainDerivative"></param>
		''' <param name="Volatility"></param>
		''' <param name="ProbabilityOfInterval"></param>
		''' <param name="VolatilityPredictionBandType"></param>
		Public Sub New(
								 ByVal NumberTradingDays As Double,
								 ByVal StockPrice As IPriceVol,
								 ByVal Gain As Double,
								 ByVal GainDerivative As Double,
								 ByVal Volatility As Double,
								 ByVal ProbabilityOfInterval As Double,
								 ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)


			MyNumberTradingDays = NumberTradingDays
			MyStockPrice = StockPrice ' This is a reference assignment, not a copy
			MyStockPriceNext = StockPriceNext
			MyStockPriceStartValue = StockPriceStartValue
			MyGain = Gain
			MyGainDerivative = GainDerivative
			MyVolatility = Volatility
			MyVolatilityTotal = MyVolatility
			MyProbabilityOfInterval = ProbabilityOfInterval
			MyVolatilityPredictionBandType = VolatilityPredictionBandType
			MyProbabilityHigh = 0.5 + MyProbabilityOfInterval / 2
			MyProbabilityLow = 0.5 - MyProbabilityOfInterval / 2
			MyGainLog = MathPlus.Measure.Measure.GainLog(MyStockPrice.Last, MyStockPrice.LastPrevious)
		End Sub

		Public ReadOnly Property NumberTradingDays As Double Implements IStockPriceVolatilityEstimateData.NumberTradingDays
			Get
				Return MyNumberTradingDays
			End Get
		End Property

		Public ReadOnly Property StockPrice As IPriceVol Implements IStockPriceVolatilityEstimateData.StockPrice
			Get
				Return MyStockPrice
			End Get
		End Property

		Public Property StockPriceNext As IPriceVol Implements IStockPriceVolatilityEstimateData.StockPriceNext
			Get
				Return MyStockPriceNext
			End Get
			Set(value As IPriceVol)
				MyStockPriceNext = value
			End Set
		End Property

		''' <summary>
		''' Refreshes the stock price prediction data by recalculating the thresholds based on the current volatility.
		''' If the stock price exceeds the calculated thresholds, it records this as a band exceeded event, indicating that
		''' the calculated volatility may not provide a good estimate of the current market risk.
		''' </summary>
		Public Sub Refresh() Implements IStockPriceVolatilityEstimateData.Refresh
			Dim ThisNumberTradingDays As Double = MyNumberTradingDays

			If MyVolatilityTotal < VOLATILITY_TOTAL_MINIMUM Then
				MyVolatilityTotal = VOLATILITY_TOTAL_MINIMUM
			End If
			_IsBandExceeded = False
			_IsBandExceededHigh = False
			_IsBandExceededLow = False

			If MyStockPriceNext Is Nothing Then
				'calculate the threshold excess base on the intraday stock price range
				'also need to consider the daily period of trading
				ThisNumberTradingDays = ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT
				MyStockPriceStartValue = MyStockPrice.Open
			Else
				MyStockPriceStartValue = MyStockPrice.Last
			End If
			' Update the threshold based on the total volatility, which includes any corrections provided by the user via the Refresh method.
			' If not updated via the Refresh method, the threshold calculation will be based on the original volatility.
			MyStockPriceHighValue = StockOption.StockPricePrediction(
					NumberTradingDays:=MyNumberTradingDays,
					StockPrice:=MyStockPriceStartValue,
					Gain:=MyGain,
					GainDerivative:=MyGainDerivative,
					Volatility:=MyVolatilityTotal,
					Probability:=MyProbabilityHigh)

			MyStockPriceLowValue = StockOption.StockPricePrediction(
								NumberTradingDays:=MyNumberTradingDays,
								StockPrice:=MyStockPriceStartValue,
								Gain:=MyGain,
								GainDerivative:=MyGainDerivative,
								Volatility:=MyVolatilityTotal,
								Probability:=MyProbabilityLow)

			If MyStockPriceNext IsNot Nothing Then
				If MyStockPriceNext.High > MyStockPriceHighValue Then
					_IsBandExceeded = True
					_IsBandExceededHigh = True
				End If
				If MyStockPriceNext.Low < MyStockPriceLowValue Then
					_IsBandExceeded = True
					_IsBandExceededLow = True
				End If
			Else
				'compare with the current value since the future is not yet available
				If MyStockPrice.High > MyStockPriceHighValue Then
					_IsBandExceeded = True
					_IsBandExceededHigh = True
				End If
				If MyStockPrice.Low < MyStockPriceLowValue Then
					_IsBandExceeded = True
					_IsBandExceededLow = True
				End If
			End If
		End Sub

		Public Sub Refresh(VolatilityDelta As Double) Implements IStockPriceVolatilityEstimateData.Refresh
			MyVolatilityTotal += VolatilityDelta
			Me.Refresh()
		End Sub

		Public Sub Reset() Implements IStockPriceVolatilityEstimateData.Reset
			MyVolatilityTotal = MyVolatility
			Me.Refresh()
		End Sub

		Public ReadOnly Property VolatilityTotal As Double Implements IStockPriceVolatilityEstimateData.VolatilityTotal
			Get
				Return MyVolatilityTotal
			End Get
		End Property

		Public ReadOnly Property StockPriceStartValue As Double Implements IStockPriceVolatilityEstimateData.StockPriceStartValue
			Get
				Return MyStockPriceStartValue
			End Get
		End Property

		Public ReadOnly Property Gain As Double Implements IStockPriceVolatilityEstimateData.Gain
			Get
				Return MyGain
			End Get
		End Property

		Public ReadOnly Property GainDerivative As Double Implements IStockPriceVolatilityEstimateData.GainDerivative
			Get
				Return MyGainDerivative
			End Get
		End Property

		Public ReadOnly Property Volatility As Double Implements IStockPriceVolatilityEstimateData.Volatility
			Get
				Return MyVolatility
			End Get
		End Property

		Public ReadOnly Property ProbabilityOfInterval As Double Implements IStockPriceVolatilityEstimateData.ProbabilityOfInterval
			Get
				Return MyProbabilityOfInterval
			End Get
		End Property

		Public ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType Implements IStockPriceVolatilityEstimateData.VolatilityPredictionBandType
			Get
				Return MyVolatilityPredictionBandType
			End Get
		End Property

		Public ReadOnly Property IsBandExceeded As Boolean Implements IStockPriceVolatilityEstimateData.IsBandExceeded
			Get
				Return _IsBandExceeded
			End Get
		End Property

		Public ReadOnly Property IsBandExceededHigh As Boolean Implements IStockPriceVolatilityEstimateData.IsBandExceededHigh
			Get
				Return _IsBandExceededHigh
			End Get
		End Property

		Public ReadOnly Property IsBandExceededLow As Boolean Implements IStockPriceVolatilityEstimateData.IsBandExceededLow
			Get
				Return _IsBandExceededLow
			End Get
		End Property

		Public ReadOnly Property GainLog As Double Implements IStockPriceVolatilityEstimateData.GainLog
			Get
				Return MyGainLog
			End Get
		End Property
	End Class
End Namespace

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
		ReadOnly Property ProbabilityOfInterval As Double
		ReadOnly Property IsBandExceeded As Boolean
		ReadOnly Property IsBandExceededHigh As Boolean
		ReadOnly Property IsBandExceededLow As Boolean
		ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		Sub Refresh()
		Sub Refresh(VolatilityDelta As Double)
		Sub Reset()
	End Interface
End Namespace

