Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter



''' <summary>
'''	Brown 's double exponential smoothing is a powerful and versatile method for time series forecasting. It captures both the level and 
'''	trend of the series, smooths out short-term fluctuations, adapts to changes, and is simple to implement and computationally efficient. 
'''	These advantages make it a popular choice for practitioners in various fields who need to produce reliable forecasts with limited data 
'''	and computational resources.
''' </summary>
Public Class FilterExpPredict
	Implements IFilterRun

	Implements IFilter
	Implements IFilterState

	Private MyRate As Integer
	Private MyFilterRate As Double
	Private MyFilterALast As Double
	Private MyFilterBLast As Double
	Private MyFilterDeltaBLast As Double
	Private MyFilterDeltaALast As Double
	Private ABRatio As Double
	Private FilterValueLastK1 As Double
	Private FilterValueLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MyFilter As IFilter
	Private MyFilterY As IFilter
	Private MyNumberToPredict As Double
	Private MyGainYearlyEstimate As Double
	Private IsReset As Boolean
	Private MyStatisticalForGain As FilterStatisticalQueue
	Private MyFilterRateYearlyScaling As Double
	Private MyFilterRateYearlyGainVolatilityScaling As Double


	''' <summary>
	'''	Construct a new instance of the FilterExpPredict class with the specified filter rate.
	''' </summary>
	''' <param name="FilterRate">The filter rate.</param>
	''' <remarks></remarks>

	Public Sub New(ByVal FilterRate As Double)
		Me.New(NumberToPredict:=0, FilterHead:=New FilterExp(FilterRate), FilterBase:=New FilterExp(FilterRate))
	End Sub

	Public Sub New(ByVal FilterRate As Double, ByVal NumberToPredict As Double)
		Me.New(NumberToPredict:=NumberToPredict, FilterHead:=New FilterExp(FilterRate), FilterBase:=New FilterExp(FilterRate))
	End Sub

	Public Sub New(ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
		Me.New(NumberToPredict:=0, FilterHead:=FilterHead, FilterBase:=FilterBase)
	End Sub

	Public Sub New(ByVal NumberToPredict As Double, ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
		Dim A As Double
		Dim B As Double

		'check parameter validity before proceeding
		If (FilterHead Is Nothing) Then
			'filter cannot be nothing a reference is needed for the filter rate
			Throw New ArgumentException("Invalid Filter type FilterHead in FilterExpPredict!")
		End If
		'FilterHead determine the filter rate
		If TypeOf FilterHead Is IFilterControl Then
			MyFilterRate = DirectCast(FilterHead, IFilterControl).FilterRate
		Else
			MyFilterRate = FilterHead.Rate
		End If
		If MyFilterRate < 2 Then
			Throw New ArgumentException("Invalid filter rate in FilterExpPredict!")
		End If
		If (FilterBase Is Nothing) Then
			'filter cannot be nothing a reference is needed for the filter rate
			Throw New ArgumentException("Invalid Filter type FilterHead in FilterExpPredict!")
		End If
		MyFilterRateYearlyScaling = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR / MyFilterRate
		MyFilterRateYearlyGainVolatilityScaling = Math.Sqrt(MyFilterRateYearlyScaling)


		MyFilter = FilterHead
		MyFilterY = FilterBase
		MyStatisticalForGain = New FilterStatisticalQueue(FilterRate:=CInt(MyFilterRate))
		MyNumberToPredict = NumberToPredict
		Dim ThisFilterRateForStatistical As Double = 5 * MyFilterRate
		If ThisFilterRateForStatistical < YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_MONTH Then
			ThisFilterRateForStatistical = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_MONTH
		End If
		FilterValueLast = 0
		FilterValueLastK1 = 0
		ValueLast = 0
		ValueLastK1 = 0
		A = CDbl((2 / (MyFilterRate + 1)))
		B = 1 - A
		ABRatio = A / B

		IsReset = True
	End Sub

	Public Overridable Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		Dim Ap As Double
		Dim Bp As Double
		Dim Result As Double
		Dim ResultY As Double
		If IsReset Then
			'initialization
			If TypeOf MyFilter Is IFilterRun Then
				DirectCast(MyFilter, IFilterRun).Reset()
			End If
			If TypeOf MyFilterY Is IFilterRun Then
				DirectCast(MyFilterY, IFilterRun).Reset()
			End If
			FilterValueLast = Value
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		Result = MyFilter.Filter(Value)
		ResultY = MyFilterY.Filter(Result)
		Ap = Result + (Result - ResultY)
		Bp = ABRatio * (Result - ResultY)
		MyFilterDeltaALast = Ap - MyFilterALast
		MyFilterDeltaBLast = Bp - MyFilterBLast
		MyFilterALast = Ap
		MyFilterBLast = Bp
		'note that Bp is the average trend
		FilterValueLast = Ap + Bp * MyNumberToPredict
		'do not use the gainLog here 
		'it is important for a low level filter not to assume that the signal is only positive
		'this filter want to be generic for any range of signal input. Instead return teh B value and leave the 
		'usee to deal with any other type of processing
		'MyStatisticalForGain.Filter(Measure.Measure.GainLog(Value:=FilterValuePredictH1, ValueRef:=Ap))
		'ThisFilterPredictionGainYearly = MyStatisticalForGain.FilterLast.ToGaussianScale(ScaleToSignedUnit:=True)
		'While the above is a good idea, we can still provide an estimate of the gain anualized by scaling the filter rate as indicated below:
		'See description 'Brown_LES_Annualized_Trend.pdf' for more information	

		'Scaling b_t to an Annualized Value
		'To ensure comparability across different filter rates, we need to scale the trend to a one-year period.
		'Given that b_t Is averaged over MyFilterRate samples, the appropriate scaling factor Is
		'f_scale = f_s / MyFilterRate
		'where:
		'- f_s = total samples per year (e.g., 252 for daily trading, 52 for weekly
		'Data).
		'- MyFilterRate = number of samples in the smoothing window.
		'Thus, the annualized trend estimate Is computed as
		'b_annual = b_t * f_scale = b_t * (f_s / MyFilterRate)
		'

		Dim ThisGainVolatilityCorrected As Double = MyFilterRateYearlyGainVolatilityScaling * MyStatisticalForGain.Filter(Value:=Bp).StandardDeviation
		MyGainYearlyEstimate = Bp * (YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR) / (1 + ThisGainVolatilityCorrected)
		'Console.WriteLine(ThisGainVolatilityCorrected)

		'note that the gain is not the same as the gainLog here. This etimate is valid for any range of signal positive or negative
		ValueLastK1 = ValueLast
		ValueLast = Value
		Return FilterValueLast
	End Function

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return FilterValueLast
		End Get
	End Property

	''' <summary>
	''' The smooted signal level before the trend is added
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterLevelLast As Double
		Get
			Return MyFilterALast
		End Get
	End Property

	''' <summary>
	''' The filtered trend or slope of the signal over the filtered sample.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterTrendLast As Double
		Get
			Return MyFilterBLast
		End Get
	End Property

	''' <summary>
	''' The logarithmic gain of the signal.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property GainLog As Double
		Get
			Return Measure.Measure.GainLog(MyFilterALast + MyFilterBLast, MyFilterALast)
		End Get
	End Property

	''' <summary>
	''' The logarithmic gain derivative of the signal taken immediatly between two consecutive samples.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property GainLogDerivative As Double
		Get
			Return Measure.Measure.GainLog(MyFilterALast + MyFilterDeltaBLast, MyFilterALast)
		End Get
	End Property

	''' <summary>
	''' Return a yearly estimate of the gain of the signal.
	''' The gain is scaled to the filter rate and the volatility of the gain.
	''' The gain is not the same as the gainLog here and the estimate is valid for any range of signal positive or negative	
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property GainYearlyEstimate As Double
		Get
			Return MyGainYearlyEstimate
		End Get
	End Property

	Public ReadOnly Property FilterRate() As Double Implements IFilterRun.FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		IsReset = True
	End Sub

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyFilterRate},{MyFilter.GetType().Name}({MyFilter.Rate}),{MyFilterY.GetType().Name}({MyFilterY.Rate}))"
		End Get
	End Property


	Public Overrides Function ToString() As String
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
	End Function

#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		Throw New NotImplementedException()
		'Try
		'	If MyQueueForState.Count = 0 Then Return
		'	ValueLast = MyQueueForState.Dequeue
		'	ValueLastK1 = MyQueueForState.Dequeue
		'	FilterValueLast = MyQueueForState.Dequeue
		'	FilterValueLastK1 = MyQueueForState.Dequeue
		'Catch ex As InvalidOperationException
		'	' Handle error, perhaps log it or rethrow with additional info
		'	Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
		'End Try
	End Sub

	Private Sub IFilterState_Save() Implements IFilterState.Save
		Throw New NotImplementedException()
		'MyQueueForState.Enqueue(ValueLast)
		'MyQueueForState.Enqueue(ValueLastK1)
		'MyQueueForState.Enqueue(FilterValueLast)
		'MyQueueForState.Enqueue(FilterValueLastK1)
	End Sub
#End Region

#Region "IFilter"
	Private ReadOnly Property IFilter_Rate As Integer Implements IFilter.Rate
		Get
			Return CInt(MyFilterRate)
		End Get
	End Property

	Private ReadOnly Property IFilter_Count As Integer Implements IFilter.Count
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_Max As Double Implements IFilter.Max
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_Min As Double Implements IFilter.Min
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_ToList As IList(Of Double) Implements IFilter.ToList
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_ToListOfError As IList(Of Double) Implements IFilter.ToListOfError
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_ToListScaled As ListScaled Implements IFilter.ToListScaled
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public Property Tag As String Implements IFilter.Tag

	Private Function IFilter_Filter(Value As Double) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Private Function IFilter_Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_Filter(Value As Single) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Private Function IFilter_Filter(Value As IPriceVol) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value.Last)
	End Function

	Private Function IFilter_FilterErrorLast() As Double Implements IFilter.FilterErrorLast
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterLast() As Double Implements IFilter.FilterLast
		Return Me.FilterLast
	End Function

	Private Function IFilter_Last() As Double Implements IFilter.Last
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray() As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToString() As String Implements IFilter.ToString
		Return Me.ToString
	End Function
#End Region
End Class
