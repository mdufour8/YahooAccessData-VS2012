Public Class FilterRunStatisticQueue
	Implements IFilterRun(Of IStatistical)

	Private MyRate As Integer
	Private FilterValueLast As IStatistical
	Private IsRunReady As Boolean
	Private _IsReset As Boolean
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private IsValueInitial As Boolean
	Private MyQueueOfValue As Queue(Of Double)
	Private MyQueueOfValueSquare As Queue(Of Double)
	Private MySumOfValue As Double
	Private MySumOfValueSquare As Double
	Private MyCircularBuffer As CircularBuffer(Of IStatistical)

#Region "New"
	''' <summary>
	''' Calculate the statistical information based on a windows size given by FilterRate. 
	''' This is a basic implementation for the statistical measurement using the square windows only.
	''' It does not use any exponential filtering and follow the usual statistic standard method.
	''' </summary>
	''' <param name="FilterRate">
	''' Should be set to be greater than two
	''' </param>
	''' <remarks></remarks>
	Public Sub New(ByVal FilterRate As Integer, Optional BufferCapacity As Integer = 0)
		MyQueueOfValue = New Queue(Of Double)
		MyQueueOfValueSquare = New Queue(Of Double)
		If FilterRate < 2 Then FilterRate = 2
		MyRate = CInt(FilterRate)
		FilterValueLast = New StatisticalData(0, 0, 0)
		ValueLast = 0
		ValueLastK1 = 0
		IsRunReady = False
		MyCircularBuffer = New CircularBuffer(Of IStatistical)(capacity:=BufferCapacity, Nothing)
	End Sub


#End Region
	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return ValueLast
		End Get
	End Property

	''' <summary>
	''' Returns the last value of the filter run. Index 0 is the most recent value added to the filter.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterLast(Index As Integer) As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			'For the CircularBuffer note 0 is the oldest value MyCircularBuffer.Count -1 is
			'the most recent value.
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
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilterRun(Of IStatistical).Count
		Get
			Return MyCircularBuffer.Count
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of IStatistical) Implements IFilterRun(Of IStatistical).ToList
		Get
			Return MyCircularBuffer.ToList()
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun(Of IStatistical).FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyRate})"
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun(Of IStatistical).Reset
		_IsReset = True
	End Sub

	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun(Of IStatistical).Reset
		Me.Reset()
		MyCircularBuffer = New CircularBuffer(Of IStatistical)(capacity:=BufferCapacity, Nothing)
	End Sub
	Public ReadOnly Property IsReset As Boolean Implements IFilterRun(Of IStatistical).IsReset
		Get
			Return _IsReset
		End Get
	End Property

	Public ReadOnly Property FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return FilterValueLast ' Return the most recent value added to the filter	
		End Get
	End Property

	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		Dim ThisValueToRemove As Double
		Dim ThisValueSquareToRemove As Double
		Dim ThisM2 As Double
		Dim ThisMean As Double
		Dim ThisVariance As Double

		If _IsReset Then
			MyCircularBuffer.Clear()
			MyQueueOfValue.Clear()
			MyQueueOfValueSquare.Clear()
			MySumOfValue = 0
			MySumOfValueSquare = 0
			ValueLast = Value
			'initialization
			FilterValueLast = New StatisticalData(Value, 0, 1)
			IsRunReady = False
		Else
			If IsRunReady = False Then
				'wait until the value change to start the volatility measurement filter
				'this is to avoid initial value of zero to be used in the calculation
				If Value <> ValueLast Then
					IsRunReady = True
				End If
			End If
		End If
		If IsRunReady Then
			If MyQueueOfValue.Count >= MyRate Then
				'start removing the oldest value from the queue
				ThisValueToRemove = MyQueueOfValue.Dequeue()
				MyQueueOfValue.Enqueue(Value)
				MySumOfValue = MySumOfValue + Value - ThisValueToRemove
				ThisMean = MySumOfValue / MyRate
				ThisM2 = (Value - ThisMean) ^ 2
				ThisValueSquareToRemove = MyQueueOfValueSquare.Dequeue()
				MyQueueOfValueSquare.Enqueue(ThisM2)
				MySumOfValueSquare = MySumOfValueSquare + ThisM2 - ThisValueSquareToRemove
				'the -1 is for the finite number of samples correction from an infinite number of samples
				ThisVariance = MySumOfValueSquare / (MyRate - 1)
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
		Else
			ThisMean = Value
			ThisVariance = 0
		End If
		FilterValueLast = New StatisticalData(ThisMean, ThisVariance, MyQueueOfValue.Count) With {.ValueLast = Value}
		ValueLastK1 = ValueLast
		ValueLast = Value
		MyCircularBuffer.AddLast(FilterValueLast)
		Return FilterValueLast
	End Function
End Class
