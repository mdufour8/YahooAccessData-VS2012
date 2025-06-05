Public Class StatisticExponential
	Implements IFilterRun(Of IStatistical)

	Private MyFilterForMean As FilterExp
	Private MyFilterForMeanSquare As FilterExp
	Private MyFilterRate As Integer
	Private MyVarianceCorrectionForPopulationSize As Double
	Private MyCircularBuffer As CircularBuffer(Of IStatistical)

	Public Sub New(ByVal FilterRate As Integer, Optional BufferCapacity As Integer = 0)
		'note the IIR filtering is done using this:
		' For additional context, see: https://en.wikipedia.org/wiki/Low-pass_filter.
		' Note that Beta is the factor applied to the previous value and ALpha to the new value
		'Alpha = 2 / (FilterRate + 1)
		'Beta = 1 - Alpha
		MyFilterForMean = New FilterExp(FilterRate)
		MyFilterForMeanSquare = New FilterExp(FilterRate)
		'just for simplification and speed
		MyVarianceCorrectionForPopulationSize = FilterRate / (FilterRate - 1)
		MyCircularBuffer = New CircularBuffer(Of IStatistical)(capacity:=BufferCapacity, Nothing)
	End Sub

#Region "IFilterRun"
	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		Return Me.Filter(Value)
	End Function

	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return Me.Last()
		End Get
	End Property

	Private ReadOnly Property IFilterRun_FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyCircularBuffer.PeekLast
		End Get
	End Property

	Private ReadOnly Property IFilterRun_FilterLast(Index As Integer) As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			'For the CircularBuffer note 0 is the oldest value MyCircularBuffer.Count -1 is
			'the most recent value.
			'The index is in the range [0, FilterRate-1].
			Dim ThisBufferIndex As Integer = MyCircularBuffer.Count - 1 - Index
			Select Case ThisBufferIndex
				Case < 0
					'return the oldest value
					Return MyCircularBuffer.PeekFirst
				Case >= MyCircularBuffer.Count
					'return the last value (most recent value)
					Return MyCircularBuffer.PeekLast
				Case Else
					'return at a sppecific location in the buffer	
					Return MyCircularBuffer.Item(BufferIndex:=ThisBufferIndex)
			End Select
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As IStatistical Implements IFilterRun(Of IStatistical).FilterTrendLast
		Get
			Return New StatisticalData(0, 0, 0)
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun(Of IStatistical).FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Private ReadOnly Property IFilterRun_Count As Integer Implements IFilterRun(Of IStatistical).Count
		Get
			Return MyCircularBuffer.Count
		End Get
	End Property

	Private ReadOnly Property IFilterRun_ToList As IList(Of IStatistical) Implements IFilterRun(Of IStatistical).ToList
		Get
			Return MyCircularBuffer.ToList()
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun(Of IStatistical).FilterDetails
		Get
			Return $"{Me.GetType().Name}({Me.FilterRate})"
		End Get
	End Property

	Public ReadOnly Property IsReset As Boolean Implements IFilterRun(Of IStatistical).IsReset
		Get
			Return _IsReset
		End Get
	End Property

	Private _IsReset As Boolean
	Public Sub Reset() Implements IFilterRun(Of IStatistical).Reset
		_IsReset = True
	End Sub

	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun(Of IStatistical).Reset
		Me.Reset()
		MyCircularBuffer = New CircularBuffer(Of IStatistical)(capacity:=BufferCapacity, Nothing)
	End Sub
#End Region
End Class
