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
	Private ABRatio As Double
	Private FilterValueLastK1 As Double
	Private FilterValueLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MyFilter As IFilter
	Private MyFilterY As IFilter
	Private MyNumberToPredict As Double
	Private IsReset As Boolean

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
		MyFilter = FilterHead
		MyFilterY = FilterBase

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

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		Dim Ap As Double
		Dim Bp As Double
		Dim Result As Double
		Dim ResultY As Double
		If IsReset Then
			'initialization
			FilterValueLast = Value
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		Result = MyFilter.Filter(Value)
		ResultY = MyFilterY.Filter(Result)
		Ap = (2 * Result) - ResultY
		Bp = ABRatio * (Result - ResultY)
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
		ValueLastK1 = ValueLast
		ValueLast = Value
		Return FilterValueLast
	End Function

	Private Function IFilterRun_FilterRun(Value As Double, FilterPLLDetector As IFilterPLLDetector) As Double Implements IFilterRun.FilterRun
		Throw New NotImplementedException()
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
	''' The filtered trend of the signal or the average trend.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterTrendLast As Double
		Get
			Return MyFilterBLast
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

	Public Overrides Function ToString() As String
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate}"
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

	Private Property IFilter_Tag As String Implements IFilter.Tag

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
