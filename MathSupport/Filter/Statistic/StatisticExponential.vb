Public Class StatisticExponential
	Implements IFilterRun(Of IStatistical)

	Private MyFilterForMean As FilterExp
	Private MyFilterForMeanSquare As FilterExp
	Private MyFilterRate As Integer
	Private StatisticCount As Integer
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
		Me.Reset(BufferCapacity:=BufferCapacity)
	End Sub

#Region "IFilterRun"
	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		If _IsReset Then
			MyCircularBuffer.Clear()
			MyFilterForMean.Reset()
			MyFilterForMeanSquare.Reset()
			StatisticCount = 0
		End If
		Dim ThisMean = MyFilterForMean.FilterRun(Value)
		Dim ThisM2 = (Value - ThisMean) ^ 2
		Dim ThisVariance = MyVarianceCorrectionForPopulationSize * MyFilterForMeanSquare.FilterRun(ThisM2)
		StatisticCount += 1
		If StatisticCount > FilterRate Then StatisticCount = CInt(MyFilterRate)
		MyCircularBuffer.AddLast(New StatisticalData(Mean:=ThisMean, Variance:=ThisVariance, NumberPoint:=StatisticCount, ValueLast:=Value))
		Return MyCircularBuffer.PeekLast
	End Function

	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return MyCircularBuffer.PeekLast.ValueLast
		End Get
	End Property

	Private ReadOnly Property IFilterRun_FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
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
