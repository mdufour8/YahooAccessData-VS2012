Imports YahooAccessData.MathPlus.Filter

Public Class FilterExpPredict
	Implements IFilterRun
	Implements IFilter
	Implements IFilterState

	Private MyRate As Integer
	Private MyFilterRate As Double
	Private AFilterLast As Double
	Private BFilterLast As Double
	Private ABRatio As Double
	Private FilterValueLastK1 As Double

	Private FilterValuePredictH1 As Double     'future 1 point
	Private FilterValuePredictH1Last As Double
	Private MyFilterPredictionGainYearlyLast As Double
	Private MyGainStandardDeviationLast As Double
	Private FilterValueLast As Double
	Private FilterValueLastY As Double
	Private FilterValueSlopeLastK1 As Double
	Private FilterValueSlopeLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MyListOfValue As ListScaled
	Private MyListOfPredictionGainPerYear As ListScaled
	Private MyListOfStatisticalVarianceError As ListScaled
	Private MyListOfAFilter As List(Of Double)
	Private MyListOfBFilter As List(Of Double)
	Private MyStatisticalForPredictionError As FilterStatistical
	Private MyStatisticalForGain As FilterStatistical
	Private MyFilter As IFilter
	Private MyFilterY As IFilter
	Private MyNumberToPredict As Double
	Private MyInputValue() As Double
	Private IsReset As Boolean

	Protected Sub New(ByVal NumberToPredict As Double, ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
		Dim A As Double
		Dim B As Double

		'check parameter validity before proceeding
		If (FilterHead Is Nothing) Then
			'filter cannot be nothing a reference is needed for the filter rate
			Throw New ArgumentException("Invalid Filter type FilterHead in FilterExpPredict!")
		End If
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

		MyStatisticalForPredictionError = New FilterStatistical(CInt(ThisFilterRateForStatistical))
		MyStatisticalForGain = New FilterStatistical(CInt(ThisFilterRateForStatistical))
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
		If IsReset Then
			'initialization
			FilterValueLast = Value
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		'FilterValueLast = A * Value + B * FilterValueLast
		ValueLastK1 = ValueLast
		ValueLast = Value
		Return FilterValueLast
	End Function

	Public Function FilterRun(Value As Double, FilterPLLDetector As IFilterPLLDetector) As Double Implements IFilterRun.FilterRun
		Return Me.FilterRun(Value)
	End Function

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return FilterValueLast
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

#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		Try
			If MyQueueForState.Count = 0 Then Return
			ValueLast = MyQueueForState.Dequeue
			ValueLastK1 = MyQueueForState.Dequeue
			FilterValueLast = MyQueueForState.Dequeue
			FilterValueLastK1 = MyQueueForState.Dequeue
		Catch ex As InvalidOperationException
			' Handle error, perhaps log it or rethrow with additional info
			Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
		End Try
	End Sub

	Private Sub IFilterState_Save() Implements IFilterState.Save
		MyQueueForState.Enqueue(ValueLast)
		MyQueueForState.Enqueue(ValueLastK1)
		MyQueueForState.Enqueue(FilterValueLast)
		MyQueueForState.Enqueue(FilterValueLastK1)
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
		Return $"FilterExp: {Me.FilterLast}"
	End Function
#End Region

End Class
