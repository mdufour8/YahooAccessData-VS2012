
Public Class FilterLinearRegressionSlope

	Private ReadOnly _windowSize As Integer
	Private ReadOnly _buffer As Queue(Of Double)

	Private ReadOnly _sumX As Double
	Private ReadOnly _sumX2 As Double
	Private ReadOnly _denominator As Double

	Private _sumY As Double
	Private _sumXY As Double

	Public Sub New(windowSize As Integer)

		If windowSize < 2 Then
			Throw New ArgumentException("Window size must be at least 2.")
		End If

		_windowSize = windowSize
		_buffer = New Queue(Of Double)(windowSize)

		_sumX = windowSize * (windowSize - 1) / 2.0
		_sumX2 = (windowSize - 1) * windowSize * (2 * windowSize - 1) / 6.0

		_denominator = windowSize * _sumX2 - _sumX * _sumX

	End Sub

	Public Function Add(value As Double) As Double

		If _buffer.Count < _windowSize Then

			Dim index As Integer = _buffer.Count

			_buffer.Enqueue(value)

			_sumY += value
			_sumXY += index * value

		Else

			Dim oldSumY As Double = _sumY
			Dim oldest As Double = _buffer.Dequeue()

			_buffer.Enqueue(value)

			_sumXY = _sumXY - (oldSumY - oldest) + (_windowSize - 1) * value
			_sumY = _sumY - oldest + value

		End If
		Return Slope
	End Function

	Public ReadOnly Property WindowSize As Integer
		Get
			Return _windowSize
		End Get
	End Property

	Public ReadOnly Property IsReady As Boolean
		Get
			Return _buffer.Count = _windowSize
		End Get
	End Property

	Public Function SlopeNormalized() As Double
		Return If(_sumY = 0, 0, Slope() / (_sumY / _buffer.Count))
	End Function
	Public Function Slope() As Double
		If _buffer.Count < 2 Then Return 0

		Dim n As Integer = _buffer.Count

		If n <> _windowSize Then
			' During warm-up, use direct calculation.
			Return CalculateSlopeDirect()
		End If

		'do not Normalize by average X to get slope per unit time leave it to the outside if it is needed
		'Return ((_windowSize * _sumXY - _sumX * _sumY) / _denominator) / _buffer.Last
		Return ((_windowSize * _sumXY - _sumX * _sumY) / _denominator)
	End Function

	Private Function CalculateSlopeDirect() As Double

		Dim values = _buffer.ToArray()
		Dim n As Integer = values.Length

		Dim sumX As Double = n * (n - 1) / 2.0
		Dim sumX2 As Double = (n - 1) * n * (2 * n - 1) / 6.0

		Dim sumY As Double = 0
		Dim sumXY As Double = 0

		For i As Integer = 0 To n - 1
			sumY += values(i)
			sumXY += i * values(i)
		Next

		Dim denominator As Double = n * sumX2 - sumX * sumX

		If denominator = 0 Then Return 0

		Return ((n * sumXY - sumX * sumY) / denominator)
	End Function
End Class