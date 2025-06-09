Public Class StatisticWindows
	Implements IFilterRun(Of IStatistical)

	Private MyQueueOfValue As Queue(Of Double)
	Private MyQueueOfValueSquare As Queue(Of Double)
	Private MySumOfValue As Double
	Private MySumOfValueSquare As Double
	Private MyFilterRate As Integer
	Private MyCircularBuffer As CircularBuffer(Of IStatistical)

	Public Sub New(ByVal FilterRate As Integer, Optional BufferCapacity As Integer = 0)
		MyQueueOfValue = New Queue(Of Double)
		MyQueueOfValueSquare = New Queue(Of Double)
		MyFilterRate = FilterRate
		'just for simplification and speed
		Me.Reset(BufferCapacity:=BufferCapacity)
	End Sub

#Region "IFilterRun"
	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		If _IsReset Then
			_IsReset = False
			MyCircularBuffer.Clear()
			MyQueueOfValue.Clear()
			MyQueueOfValueSquare.Clear()
			MySumOfValue = 0.0
			MySumOfValueSquare = 0.0
		End If
		Dim ThisValueToRemove As Double
		Dim ThisValueSquareToRemove As Double
		Dim ThisM2 As Double
		Dim ThisMean As Double
		Dim ThisVariance As Double

		If MyQueueOfValue.Count >= MyFilterRate Then
			'start removing the oldest value from the queue
			ThisValueToRemove = MyQueueOfValue.Dequeue()
			MyQueueOfValue.Enqueue(Value)
			MySumOfValue = MySumOfValue + Value - ThisValueToRemove
			ThisMean = MySumOfValue / MyFilterRate
			ThisM2 = (Value - ThisMean) ^ 2
			ThisValueSquareToRemove = MyQueueOfValueSquare.Dequeue()
			MyQueueOfValueSquare.Enqueue(ThisM2)
			MySumOfValueSquare = MySumOfValueSquare + ThisM2 - ThisValueSquareToRemove
			'the -1 is for the finite number of samples correction from an infinite number of samples
			ThisVariance = MySumOfValueSquare / (MyFilterRate - 1)
		Else
			MyQueueOfValue.Enqueue(Value)
			MySumOfValue = MySumOfValue + Value
			ThisMean = MySumOfValue / MyQueueOfValue.Count
			ThisM2 = (Value - ThisMean) ^ 2
			MyQueueOfValueSquare.Enqueue(ThisM2)
			MySumOfValueSquare = MySumOfValueSquare + ThisM2
			If MyQueueOfValue.Count > 1 Then
				'-1 take care of the finite population effect
				ThisVariance = MySumOfValueSquare / (MyQueueOfValue.Count - 1)
			Else
				ThisVariance = MySumOfValueSquare
			End If
		End If
		Dim ThisStatisticalData = New StatisticalData(Mean:=ThisMean, Variance:=ThisVariance, NumberPoint:=MyQueueOfValue.Count, ValueLast:=Value)
		MyCircularBuffer.AddLast(ThisStatisticalData)
		Return ThisStatisticalData
	End Function

	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return MyCircularBuffer.PeekLast.ValueLast
		End Get
	End Property

	Public ReadOnly Property FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyCircularBuffer.PeekLast
		End Get
	End Property

	''' <summary>
	''' 'The index is in the range [0, FilterRate-1]. Zero is the most recent data
	''' </summary>
	''' <param name="Index"></param>
	''' <returns></returns>
	Public ReadOnly Property FilterLast(Index As Integer) As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
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

	''' <summary>
	''' does not exist in the current statistical concept
	''' </summary>
	''' <returns></returns>
	Private ReadOnly Property FilterTrendLast As IStatistical Implements IFilterRun(Of IStatistical).FilterTrendLast
		Get
			Return New StatisticalData(0, 0, 0)
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun(Of IStatistical).FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Public ReadOnly Property ToBufferList As IList(Of IStatistical) Implements IFilterRun(Of IStatistical).ToBufferList
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
