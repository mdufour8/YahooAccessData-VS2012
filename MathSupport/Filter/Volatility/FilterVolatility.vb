#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region

Namespace MathPlus.Filter
#Region "FilterVolatility"
	''' <summary>
	''' This class implements the close-to-close standard volatility measurement.
	''' By default, it returns the annualized volatility assuming a daily sampling rate.
	''' The user can modify the annualization factor by specifying a custom scale factor during object creation.
	''' The volatility is calculated using the standard deviation of the logarithmic returns.
	''' Consideration Details:
	'''
	'''  Volatility is a standardized measure of the variability of returns, not raw prices.
	''' 
	''' - Stock prices are non-stationary: they trend over time and do not revert to a fixed mean.
	''' - Using raw prices in volatility calculation leads to overestimated and inconsistent results.
	''' - Returns (especially log returns) are stationary and mean-reverting, making statistical analysis valid.
	''' - Log returns are scale-invariant, time-additive, and comparable across different stocks.
	''' 
	''' Preferred formula:
	'''   LogReturn = Log(P_t / P_{t-1})
	''' 
	''' This approach ensures volatility measures reflect relative changes, not price levels or long-term trends.
	''' 
	''' see more details definition:
	''' http://en.wikipedia.org/wiki/Rate_of_return#Logarithmic_or_continuously_compounded_return
	''' http://en.wikipedia.org/wiki/Volatility_%28finance%29
	''' https://www.youtube.com/watch?v=eiTCTibH010
	''' http://en.wikipedia.org/wiki/Volatility_(finance)
	''' https://en.wikipedia.org/wiki/Stochastic_volatility
	''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
	''' </summary>
	''' <remarks>
	''' The simplest and most common type of calculation that
	''' benefits from only using reliable prices from closing auctions. We note that the
	''' volatility should be the standard deviation multiplied by √N/(N-1) to take into
	''' account the fact we are sampling a smaller subset of the population.
	''' This class also use the normalized logarithm compounded return to measure the volatility and
	''' is only valid for positive value
	''' </remarks>
	<Serializable()>
  Public Class FilterVolatility
		Implements IFilterRun
		Implements IFilter
		Implements IRegisterKey(Of String)

		Public Enum enuVolatilityStatisticType
			Standard
			Exponential
		End Enum


		Public Const VOLATILITY_FILTER_RATE_DEFAULT As Integer = MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12

		Private MyRate As Integer
		Private MyFilterValueLastK1 As Double
		Private MyFilterValueLast As Double
		Private MyValueLast As Double
		Private ValueLastK1 As Double
		Private MyFilterVolatilityYearlyCorrection As Double
		Private MyListOfFilterValue As ListScaled
		Private MyStatistical As IFilter(Of IStatistical)
		Private MyStatisticType As enuVolatilityStatisticType
		'Private MyPriceNextDailyHighPreviousCloseToOpenSigma2 As Double
		Private IsSpecialDividendPayoutLocal As Boolean

		Public Sub New()
			Me.New(VOLATILITY_FILTER_RATE_DEFAULT, Math.Sqrt(NUMBER_TRADINGDAY_PER_YEAR))
		End Sub

		Public Sub New(ByVal FilterRate As Integer, Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard, Optional SamplingRatePerDay As Double = 1.0)
			Me.New(FilterRate, Math.Sqrt(SamplingRatePerDay * NUMBER_TRADINGDAY_PER_YEAR), StatisticType)
		End Sub

		''' <summary>
		''' By observation the standard deviation of the measured Volatility using an exponential filter
		''' is slighlty larger than using a square window method
		''' The square windows might be preferable for a better estimate with less deviation of the volatility:
		''' i.e. 10% deviation standard for he exponential filter versus 6% for the square windows for 30% volatility
		''' and a window size or filter exponential equivalent of 120 sample. 
		''' </summary>
		''' <param name="FilterRate"></param>
		''' <param name="ScaleCorrection"></param>
		''' <param name="StatisticType"></param>
		Public Sub New(
			ByVal FilterRate As Integer,
			ByVal ScaleCorrection As Double,
			Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard)

			MyFilterVolatilityYearlyCorrection = ScaleCorrection
			MyListOfFilterValue = New ListScaled
			MyStatisticType = StatisticType
			If FilterRate < 1 Then FilterRate = 1
			MyRate = CInt(FilterRate)
			MyStatistical = New FilterStatistical(FilterRate, StatisticType:=MyStatisticType)
			MyFilterValueLast = 0
			MyFilterValueLastK1 = 0
			MyValueLast = 0
			ValueLastK1 = 0
			IsSpecialDividendPayoutLocal = False
		End Sub



		''' <summary>
		''' True for ignoring the volatility jump due to an open price exceeding the expected 2x sigma price peak value.
		''' This can be usuful for reducing the impact of unexpected news on the stock volatility. 
		''' </summary>
		''' <returns>The current state</returns>
		Public Property IsFilterVolatilityJump As Boolean

		''' <summary>
		''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
		''' This method expect positive value of asset price.
		''' </summary>
		''' <param name="Value">The current positive value of the asset</param>
		''' <param name="ValueRef">
		''' The reference value for the asset return calculation normally the last value of the sample. This function may likely require to adjust
		''' the scale factor to unity for raw volatility measurement
		''' </param>
		''' <remarks>
		''' The function assume by default a daily data input. The scale factor may need to be adjusted if the data
		''' is not at the daily sample rate and the yearly volatility is needed.
		''' </remarks>
		Public Function Filter(ByVal Value As Double, ByVal ValueRef As Double) As Double
			Dim ThisReturnLog As Double

			If MyListOfFilterValue.Count = 0 Then
				'assume volatility of zero at start
				MyFilterValueLast = 0
			Else
				If IsFilterVolatilityJump Then
					If MyFilterValueLast > 0 Then
						'nothing to calculate if the volatility is zero
						'calculate the 2 sigma range for the current stock and volatility
						Dim ThisPriceNextDailyHighPreviousCloseToClose = StockOption.StockPricePrediction(
							NumberTradingDays:=1,
							Me.MyValueLast,
							Gain:=0.0,
							GainDerivative:=0.0,
							Me.MyFilterValueLast,
							GAUSSIAN_PROBABILITY_SIGMA3)

						If Value > ThisPriceNextDailyHighPreviousCloseToClose Then
							ThisPriceNextDailyHighPreviousCloseToClose = ThisPriceNextDailyHighPreviousCloseToClose
						End If
					End If
				End If
			End If
			If IsSpecialDividendPayoutLocal Then
				'ignore the current data and use the previous calculation
				IsSpecialDividendPayoutLocal = False
				ThisReturnLog = MyStatistical.Last
			Else
				'same thing should be fixed
				'ThisReturnLog = GainLog(Value, ValueRef)
				ThisReturnLog = LogPriceReturn(Value, ValueRef)
			End If
			If MyStatistical.Count = 0 Then
				'start filtering only if the first valid data i.e. not zero
				'this help to eliminate the first day of trading when there is no data from the previous day
				If ThisReturnLog <> 0 Then
					'first valid measurement of volatility
					MyStatistical.Filter(ThisReturnLog)
				End If
			Else
				MyStatistical.Filter(ThisReturnLog)
			End If
			MyFilterValueLastK1 = MyFilterValueLast
			'correct the value for the yearly variation
			MyFilterValueLast = MyFilterVolatilityYearlyCorrection * MyStatistical.FilterLast.StandardDeviation
			MyListOfFilterValue.Add(MyFilterValueLast)
			'calculate the next sample 2 sigma normal last price range
			If MyListOfFilterValue.Count > 0 Then
				If IsFilterVolatilityJump Then
					'MyPriceNextDailyHighPreviousCloseToOpenSigma2 = OptionValuation.StockOption.StockPricePrediction(
					'  NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
					'  StockPrice:=Value,
					'  Gain:=0.0,
					'  GainDerivative:=0.0,
					'  Volatility:=MyFilterValueLast,
					'  Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
				End If
			End If
			ValueLastK1 = MyValueLast
			MyValueLast = Value
			Return MyFilterValueLast
		End Function

		''' <summary>
		''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
		''' This method expect positive value of asset price.
		''' </summary>
		''' <param name="Value">The current positive value of the asset</param>
		''' <returns>The current volatility corrected by default to a yearly period assuming a daily sample rate.</returns>
		''' <remarks>
		''' The function assume by default a daily data input. The scale factor may need to be adjusted if the data
		''' is not at the daily sample rate and the yearly volatility is needed.
		''' </remarks>
		Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
			Return Me.Filter(Value, MyValueLast)
		End Function


		Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
			IsSpecialDividendPayoutLocal = Value.IsSpecialDividendPayout
			If Value.Vol > 0 Then
				Return Me.Filter(CDbl(Value.Last), MyValueLast)
			Else
				'no volume mean no volatility
				'and the point should not be included in the calculation
				MyListOfFilterValue.Add(Me.MyFilterValueLast)
				ValueLastK1 = MyValueLast
				MyValueLast = Value.Last
				Return Me.MyFilterValueLast
			End If
		End Function

		' ''' <summary>
		' ''' Compute the yearly Volatility based on the Garman and Klauss estimator taking in account the high and low
		' ''' </summary>
		' ''' <param name="Value"></param>
		' ''' <param name="ValueHigh"></param>
		' ''' <param name="ValueLow"></param>
		' ''' <returns></returns>
		' ''' <remarks></remarks>
		'Public Function Filter(ByVal Value As YahooAccessData.IPriceVol) As Double
		'  Dim ThisReturnLog As Double
		'  Dim ThisReturnLogHighLow As Double


		'  Throw New NotImplementedException
		'  'If MyListOfFilterValue.Count = 0 Then
		'  '  MyFilterValueLast = Value.Last
		'  'End If
		'  'If MyValueLast <= 0 Then
		'  '  ThisReturnLog = 0
		'  '  ThisReturnLogHighLow = 0
		'  'Else
		'  '  If Value <= 0 Then
		'  '    ThisReturnLog = 0
		'  '    ThisReturnLogHighLow = 0
		'  '  Else
		'  '    ThisReturnLog = Math.Log(Value / MyValueLast)
		'  '  End If
		'  'End If
		'  'MyStatistical.Filter(ThisReturnLog)
		'  'MyFilterValueLastK1 = MyFilterValueLast
		'  ''correct the value for the yearly variation
		'  'MyFilterValueLast = MyFilterVolatilityYearlyCorrection * MyStatistical.FilterLast.StandardDeviation
		'  'MyListOfFilterValue.Add(MyFilterValueLast)
		'  'ValueLastK1 = MyValueLast
		'  'MyValueLast = Value
		'  'Return MyFilterValueLast
		'End Function

		Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
			Dim ThisValue As Double
			For Each ThisValue In Value
				Me.Filter(ThisValue)
			Next
			Return Me.ToArray
		End Function

		Public Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
			Throw New NotSupportedException
		End Function

		Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
			Throw New NotSupportedException
		End Function

		Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
			Throw New NotSupportedException
		End Function

		Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
			Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
			With ThisPriceVol
				.LastPrevious = CSng(MyFilterValueLastK1)
				If Me.FilterLast > .Last Then
					.High = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.Last < .Last Then
					.Low = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
			Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
			With ThisPriceVol
				.LastPrevious = CSng(ValueLastK1)
				If Me.FilterLast > .Last Then
					.High = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.FilterLast < .Last Then
					.Low = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
			Return CSng(Me.Filter(CDbl(Value)))
		End Function

		Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
			Throw New NotSupportedException
		End Function

		Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
			Return Me.FilterPredictionNext(CDbl(Value))
		End Function

		Public Function FilterLast() As Double Implements IFilter.FilterLast
			Return MyFilterValueLast
		End Function

		Public Function Last() As Double Implements IFilter.Last
			Return MyValueLast
		End Function

		Public ReadOnly Property Rate As Integer Implements IFilter.Rate
			Get
				Return MyRate
			End Get
		End Property

		Public ReadOnly Property Count As Integer Implements IFilter.Count
			Get
				Return MyListOfFilterValue.Count
			End Get
		End Property

		Public ReadOnly Property Max As Double Implements IFilter.Max
			Get
				Return MyListOfFilterValue.Max
			End Get
		End Property

		Public ReadOnly Property Min As Double Implements IFilter.Min
			Get
				Return MyListOfFilterValue.Min
			End Get
		End Property

		Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
			Get
				Return MyListOfFilterValue
			End Get
		End Property

		Private ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
			Get
				Throw New NotSupportedException
			End Get
		End Property

		Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
			Get
				Return MyListOfFilterValue
			End Get
		End Property

		Public Function ToArray() As Double() Implements IFilter.ToArray
			Return MyListOfFilterValue.ToArray
		End Function

		Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Return MyListOfFilterValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
		End Function

		Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Return MyListOfFilterValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
		End Function

		Public Property Tag As String Implements IFilter.Tag

		Public Overrides Function ToString() As String Implements IFilter.ToString
			Return Me.FilterLast.ToString
		End Function

#Region "IFilterRun"
		Private ReadOnly Property IFilterRun_InputLast As Double Implements IFilterRun.InputLast
			Get
				Return MyValueLast
			End Get
		End Property

		Private ReadOnly Property IFilterRun_FilterLast As Double Implements IFilterRun.FilterLast
			Get
				Return MyFilterValueLast
			End Get
		End Property


		''' <summary>
		''' Compare to a standard list to get the data at a specific index, the current 
		''' data access is reversed.
		''' The index is in the range [0, MyListOfFilterValue.Count - 1].
		''' The index 0 is the most recent value and MyListOfFilterValue.Count -1 is the oldest value.	
		''' </summary>
		''' <param name="Index"></param>
		''' <returns></returns>
		Private ReadOnly Property IFilterRun_FilterLast(Index As Integer) As Double Implements IFilterRun.FilterLast
			Get
				'we can use the current List to get data at a specific index
				'note 0 is the oldest value MyCircularBuffer.Count -1 is the most recent value.
				'The index reversed in the range [0, MyListOfFilterValue.Count - 1].
				'change the access range to be in the range [0, MyListOfFilterValue.Count - 1]	
				Dim ThisBufferIndex As Integer = MyListOfFilterValue.Count - 1 - Index
				Select Case ThisBufferIndex
					Case < 0
						'return the oldest value
						Return MyListOfFilterValue.First
					Case >= MyListOfFilterValue.Count
						'return the last value (most recent value)
						Return MyListOfFilterValue.Last
					Case Else
						'return at a specific location in the buffer	
						Return MyListOfFilterValue.Item(index:=ThisBufferIndex)
				End Select
			End Get
		End Property

		Private ReadOnly Property IFilterRun_FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
			Get
				Throw New NotImplementedException(message:=Me.ToString & " does not support FilterTrendLast")
			End Get
		End Property

		Private ReadOnly Property IFilterRun_FilterRate As Double Implements IFilterRun.FilterRate
			Get
				Return MyRate
			End Get
		End Property

		''' <summary>
		''' is not supported here, the filter is not a circular buffer but a list of values.
		''' The data can be accessed using the IFilter ToList or ToArray method.
		''' </summary>
		''' <returns></returns>
		Private ReadOnly Property IFilterRun_ToBufferList As IList(Of Double) Implements IFilterRun.ToBufferList
			Get
				Throw New NotImplementedException(message:=Me.ToString & " does not support IFilterRun_ToBufferList")
			End Get
		End Property

		Private ReadOnly Property IFilterRun_FilterDetails As String Implements IFilterRun.FilterDetails
			Get
				Return $"{Me.GetType().Name}({MyRate},{MyStatisticType})"
			End Get
		End Property

		Public ReadOnly Property IsReset As Boolean Implements IFilterRun.IsReset
			Get
				Throw New NotImplementedException()
			End Get
		End Property
		Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
			Throw New NotImplementedException()
		End Function

		Public Sub Reset() Implements IFilterRun.Reset
			Throw New NotImplementedException()
		End Sub

		Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun.Reset
			Throw New NotImplementedException()
		End Sub
#End Region
#Region "IRegisterKey"
		Public Function AsIRegisterKey() As IRegisterKey(Of String)
			Return Me
		End Function

		Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
		Dim MyKeyValue As String
		Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
			Get
				Return MyKeyValue
			End Get
			Set(value As String)
				MyKeyValue = value
			End Set
		End Property
#End Region
	End Class
#End Region
End Namespace
