Imports YahooAccessData.MathPlus.Filter
Imports MathNet.Filtering
Imports MathNet.Filtering.IIR

Public Class FilterLPButterworth
	Implements IFilterRun
	Implements IFilter
	'Implements IFilterState

	Private MyRate As Integer
	Private MyFilterRate As Double
	Private A As Double
	Private B As Double
	Private FilterValueLastK1 As Double
	Private FilterValueLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MyFilter As OnlineFilter
	Private IsReset As Boolean

	''' <summary>
	''' Implements a Butterworth Low Pass filter with e 3 dB bandwith of 1/FilterRate
	''' The filter is implemented using the MathNet library.
	''' The filter is a second-order low-pass Butterworth filter.
	''' The filter is initialized with a cutoff frequency of 1/FilterRate Hz and a sample rate of 1 sample.
	''' The filter can be reset to its initial state.
	''' </summary>
	Public Sub New(ByVal FilterRate As Double)

		If FilterRate < 1 Then FilterRate = 2
		MyFilterRate = FilterRate
		MyRate = CInt(MyFilterRate)

		' Define the sample rate and the desired cutoff frequency
		Dim sampleRate As Double = 1.0 ' Sample rate in Hz  
		Dim cutoffFrequency As Double = 1 / MyFilterRate ' Cutoff frequency in Hz

		' Create a second-order low-pass Butterworth filter
		MyFilter = OnlineIirFilter.CreateLowpass(
				mode:=ImpulseResponse.Infinite,
				sampleRate:=sampleRate,
				cutoffRate:=cutoffFrequency,
				order:=2) ' Order of the filter


		FilterValueLast = 0
		FilterValueLastK1 = 0
		ValueLast = 0
		ValueLastK1 = 0
		IsReset = True
		MyFilter.Reset()
		MyFilter.ProcessSample(0.0) ' Initialize the filter with a sample of 0	
	End Sub

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		If IsReset Then
			'initialization
			FilterValueLast = Value
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		FilterValueLast = MyFilter.ProcessSample(Value)
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

	'#Region "IFilterState"
	'	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
	'		Return Me
	'	End Function

	'	Private MyQueueForState As New Queue(Of Double)
	'	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
	'		Try
	'			If MyQueueForState.Count = 0 Then Return
	'			ValueLast = MyQueueForState.Dequeue
	'			ValueLastK1 = MyQueueForState.Dequeue
	'			FilterValueLast = MyQueueForState.Dequeue
	'			FilterValueLastK1 = MyQueueForState.Dequeue
	'		Catch ex As InvalidOperationException
	'			' Handle error, perhaps log it or rethrow with additional info
	'			Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
	'		End Try
	'	End Sub

	'	Private Sub IFilterState_Save() Implements IFilterState.Save
	'		MyQueueForState.Enqueue(ValueLast)
	'		MyQueueForState.Enqueue(ValueLastK1)
	'		MyQueueForState.Enqueue(FilterValueLast)
	'		MyQueueForState.Enqueue(FilterValueLastK1)
	'	End Sub
	'#End Region

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

