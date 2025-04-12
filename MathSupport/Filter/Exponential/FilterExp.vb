Imports YahooAccessData.MathPlus.Filter

Public Class FilterExp
	Implements IFilterRun

	Implements IFilter
	Implements IFilterState

	Private MyRate As Integer
	Private MyFilterRate As Double
	Private A As Double
	Private B As Double
	Private FilterValueLastK1 As Double
	Private FilterValueLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private IsReset As Boolean

	Public Sub New(ByVal FilterRate As Double)

		If FilterRate < 1 Then FilterRate = 1
		MyFilterRate = FilterRate
		MyRate = CInt(MyFilterRate)

		'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
		'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
		'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
		'Note this is the noise equivalent bandwidth (NEB), not the 3 dB bandwidth. The NEB is the bandwidth of a brick-wall filter
		'that would let the same amount of noise through as the filter in question.
		'see the explication in the file included in this directory 'IIR_Filter_NEB_vs_3dB_Bandwidth_FINAL.pdf'


		'THe formula that follow originate from this technical aspect:
		'Comparison with standard moving average
		'Exponential smoothing And moving average have similar defects Of introducing a lag relative To the input data.
		'While this can be corrected by shifting the result by half the window length For a symmetrical kernel, such As a
		'moving average Or gaussian, it Is unclear how appropriate this would be for exponential smoothing.
		'They (moving average with symmetrical kernels) also both have roughly the same distribution of forecast error
		'when α = 2/(k + 1) where k Is the number of past data points in consideration of moving average. They differ in that
		'exponential smoothing takes into account all past data, whereas moving average only takes into account k past data points.
		'Computationally speaking, they also differ in that moving average requires that the past k data points,
		'Or the data point at lag k + 1 plus the most recent forecast value, to be kept, whereas exponential
		'smoothing only needs the most recent forecast value to be kept.[11]


		'So the following term is not related to 3 dB bandwidth nor to the noise equivalent bandwidth (NEB)
		'It is related to the equivalent number of point than give a similar result that a flat moving average with a fixed number of point.
		A = CDbl((2 / (MyFilterRate + 1)))   'for MyFilterRate=10 this give A=0.1818

		'however given the Alpha the 3 dB badwidth is approximativly given by:
		'Ferquency 3dB = (Sampling Frequency*ln(1-A)/2PI
		'where A is the factor for the previous value	

		'The equivalent noise Bandwidth is given by PI*(A/(2-A))

		'Another aspect to consider is time or number of sample to reach the 63% of the final value	for a unit step input.
		'This is given by the formula: Time Constant = Sampling Period/ln(1-A). This time constant value can be use to estimate the number of sample 
		'requires to reach 95% of the final value of the unit step. This is given by the formula: 3 * Time Constant.	


		'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
		' B = 1 - A is the factor for the previous value
		B = 1 - A
		FilterValueLast = 0
		FilterValueLastK1 = 0
		ValueLast = 0
		ValueLastK1 = 0
		IsReset = True
	End Sub

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		If IsReset Then
			'initialization
			FilterValueLast = Value
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		FilterValueLast = A * Value + B * FilterValueLast
		ValueLastK1 = ValueLast
		ValueLast = Value
		Return FilterValueLast
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

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyFilterRate})"
		End Get
	End Property

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

	Public Overrides Function ToString() As String Implements IFilter.ToString
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
	End Function
#End Region
End Class

