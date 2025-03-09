Namespace OptionValuation
	Public Class StockPricePredictionData
		Implements IStockPricePredictionData

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

		Public Sub New(
								 ByVal NumberTradingDays As Double,
								 ByRef StockPrice As IPriceVol,
								 ByRef StockPriceNext As IPriceVol,
								 ByVal StockPriceStartValue As Double,
								 ByVal Gain As Double,
								 ByVal GainDerivative As Double,
								 ByVal Volatility As Double,
								 ByVal ProbabilityOfInterval As Double,
								 ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)

			MyNumberTradingDays = NumberTradingDays
			MyStockPrice = StockPrice
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
			Me.Refresh()
		End Sub

		Public ReadOnly Property NumberTradingDays As Double Implements IStockPricePredictionData.NumberTradingDays
			Get
				Return MyNumberTradingDays
			End Get
		End Property

		Public Function GetStockPrice() As IPriceVol Implements IStockPricePredictionData.GetStockPrice
			Return MyStockPrice
		End Function

		Public Function GetStockPriceNext() As IPriceVol Implements IStockPricePredictionData.GetStockPriceNext
			Return MyStockPriceNext
		End Function

		Public Sub Refresh() Implements IStockPricePredictionData.Refresh
			If MyVolatilityTotal < VOLATILITY_TOTAL_MINIMUM Then
				MyVolatilityTotal = VOLATILITY_TOTAL_MINIMUM
			End If
			_IsBandExceeded = False
			_IsBandExceededHigh = False
			_IsBandExceededLow = False

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
				If MyStockPriceNext.High >= MyStockPriceHighValue Then
					_IsBandExceeded = True
					_IsBandExceededHigh = True
				End If
				If MyStockPriceNext.Low <= MyStockPriceLowValue Then
					_IsBandExceeded = True
					_IsBandExceededLow = True
				End If
			Else
				'compare with the current value since the future is not yet availaible
				If MyStockPrice.High >= MyStockPriceHighValue Then
					_IsBandExceeded = True
					_IsBandExceededHigh = True
				End If
				If MyStockPrice.Low <= MyStockPriceLowValue Then
					_IsBandExceeded = True
					_IsBandExceededLow = True
				End If
			End If
		End Sub

		Public Sub Refresh(VolatilityDelta As Double) Implements IStockPricePredictionData.Refresh
			MyVolatilityTotal += VolatilityDelta
			Me.Refresh()
		End Sub

		Public Sub Reset() Implements IStockPricePredictionData.Reset
			MyVolatilityTotal = MyVolatility
			Me.Refresh()
		End Sub

		Public ReadOnly Property VolatilityTotal As Double Implements IStockPricePredictionData.VolatilityTotal
			Get
				Return MyVolatilityTotal
			End Get
		End Property


		Public ReadOnly Property StockPriceStartValue As Double Implements IStockPricePredictionData.StockPriceStartValue
			Get
				Return MyStockPriceStartValue
			End Get
		End Property

		Public ReadOnly Property Gain As Double Implements IStockPricePredictionData.Gain
			Get
				Return MyGain
			End Get
		End Property

		Public ReadOnly Property GainDerivative As Double Implements IStockPricePredictionData.GainDerivative
			Get
				Return MyGainDerivative
			End Get
		End Property

		Public ReadOnly Property Volatility As Double Implements IStockPricePredictionData.Volatility
			Get
				Return MyVolatility
			End Get
		End Property

		Public ReadOnly Property ProbabilityOfInterval As Double Implements IStockPricePredictionData.ProbabilityOfInterval
			Get
				Return MyProbabilityOfInterval
			End Get
		End Property

		Public ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType Implements IStockPricePredictionData.VolatilityPredictionBandType
			Get
				Return MyVolatilityPredictionBandType
			End Get
		End Property

		Public ReadOnly Property IsBandExceeded As Boolean Implements IStockPricePredictionData.IsBandExceeded
			Get
				Return _IsBandExceeded
			End Get
		End Property

		Public ReadOnly Property IsBandExceededHigh As Boolean Implements IStockPricePredictionData.IsBandExceededHigh
			Get
				Return _IsBandExceededHigh
			End Get
		End Property

		Public ReadOnly Property IsBandExceededLow As Boolean Implements IStockPricePredictionData.IsBandExceededLow
			Get
				Return _IsBandExceededLow
			End Get
		End Property
	End Class
End Namespace

Namespace OptionValuation
	Public Interface IStockPricePredictionData
		ReadOnly Property NumberTradingDays As Double
		Function GetStockPrice() As IPriceVol
		Function GetStockPriceNext() As IPriceVol
		ReadOnly Property StockPriceStartValue As Double
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
