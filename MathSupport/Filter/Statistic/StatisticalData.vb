Imports YahooAccessData.MathPlus

Public Class StatisticalData
	Implements IStatistical

	Private MyMean As Double
	Private MyVariance As Double
	Private MyStandardDeviation As Double
	Private MyHigh As Double
	Private MyLow As Double
	Private MyNumberPoint As Integer

	Public Sub New(ByVal Mean As Double, ByVal Variance As Double, ByVal NumberPoint As Integer, Optional ValueLast As Double = 0.0)
		MyMean = Mean
		MyVariance = Variance
		MyStandardDeviation = Math.Sqrt(MyVariance)
		MyHigh = MyMean + MyStandardDeviation
		MyLow = MyMean - MyStandardDeviation
		MyNumberPoint = NumberPoint
		Me.ValueLast = ValueLast
	End Sub

	Public Sub New(ByVal StatisticalData As IStatistical)
		Me.New(StatisticalData.Mean, StatisticalData.Variance, StatisticalData.NumberPoint, ValueLast:=StatisticalData.ValueLast)
	End Sub

	Public ReadOnly Property Mean As Double Implements IStatistical.Mean
		Get
			Return MyMean
		End Get
	End Property

	Public ReadOnly Property StandardDeviation As Double Implements IStatistical.StandardDeviation
		Get
			Return MyStandardDeviation
		End Get
	End Property

	Public ReadOnly Property Variance As Double Implements IStatistical.Variance
		Get
			Return MyVariance
		End Get
	End Property

	Public Function Copy() As IStatistical Implements IStatistical.Copy
		Return New StatisticalData(MyMean, MyVariance, MyNumberPoint, Me.ValueLast)
	End Function

	Public ReadOnly Property High As Double Implements IStatistical.High
		Get
			Return MyHigh
		End Get
	End Property

	Public ReadOnly Property Low As Double Implements IStatistical.Low
		Get
			Return MyLow
		End Get
	End Property

	Public Overrides Function ToString() As String
		Return String.Format("Mean={0:n3}, SD={1:n3}, High={2:n3}, Low={3:n3}", Me.Mean, Me.StandardDeviation, Me.High, Me.Low)
	End Function

	Public ReadOnly Property NumberPoint As Integer Implements IStatistical.NumberPoint
		Get
			Return MyNumberPoint
		End Get
	End Property

	Public Property ValueLast As Double Implements IStatistical.ValueLast

	Public Sub Add(Value As IStatistical) Implements IStatistical.Add
		Dim ThisWeight1 As Double
		Dim ThisWeight2 As Double
		Dim ThisMeanTotal As Double
		Dim ThisMeanSquare1 As Double
		Dim ThisMeanSquare2 As Double
		Dim ThisMeanSquareTotal As Double
		Dim ThisNumberPoint As Integer

		ThisNumberPoint = Me.NumberPoint + Value.NumberPoint
		ThisWeight1 = Me.NumberPoint / ThisNumberPoint
		ThisWeight2 = 1 - ThisWeight1
		ThisMeanTotal = ThisWeight1 * Me.Mean + ThisWeight2 * Value.Mean
		ThisMeanSquare1 = Me.Variance + Me.Mean ^ 2
		ThisMeanSquare2 = Value.Variance + Value.Mean ^ 2
		ThisMeanSquareTotal = ThisWeight1 * ThisMeanSquare1 + ThisWeight2 * ThisMeanSquare2
		ThisMeanSquareTotal = ThisMeanSquareTotal - ThisMeanTotal ^ 2
		Me.CopyTo(New StatisticalData(ThisMeanTotal, ThisMeanSquareTotal, ThisNumberPoint, Value.ValueLast))
	End Sub

	Public Sub CopyTo(Value As IStatistical) Implements IStatistical.CopyTo
		MyMean = Value.Mean
		MyVariance = Value.Variance
		MyStandardDeviation = Value.StandardDeviation
		MyHigh = Value.High
		MyLow = Value.Low
		MyNumberPoint = Value.NumberPoint
		Me.ValueLast = Value.ValueLast
	End Sub

	Public Function ToGaussianScale(Optional ScaleToSignedUnit As Boolean = False) As Double Implements IStatistical.ToGaussianScale
		Dim ThisGaussianRatio As Double
		If Me.StandardDeviation > 0 Then
			'transform the gain in PerCent per Year as a probability mainly for display
			ThisGaussianRatio = Measure.Measure.CDFGaussian(Me.ValueLast / Me.StandardDeviation)
			If ScaleToSignedUnit Then
				ThisGaussianRatio = 2 * (ThisGaussianRatio - 0.5)
			End If
		Else
			ThisGaussianRatio = 0.0
		End If
		Return ThisGaussianRatio
	End Function

	'''' <summary>
	'''' Calculate the LogNormal Mu parameter. This function is valid if teh function follow a logmal distribution
	'''' see: https://en.wikipedia.org/wiki/Log-normal_distribution
	'''' </summary>
	'Private _LogNormalMu As Double?
	'Public Function LogNormalMu() As Double Implements IStatistical.LogNormalMu
	'	If _LogNormalMu.HasValue Then
	'		Return _LogNormalMu.Value
	'	End If
	'	_LogNormalMu = Math.Log(Me.Mean / Math.Sqrt((Me.Variance / Me.Mean ^ 2)) + 1)
	'	Return _LogNormalMu.Value
	'End Function

	'''' <summary>
	'''' Calculate the LogNormal Sigma parameter. This function is valid if the function follow a logmal distribution
	'''' </summary>
	'Private _LogNormalSigma As Double?
	'Public Function LogNormalSigma() As Double Implements IStatistical.LogNormalSigma
	'	If _LogNormalSigma.HasValue Then
	'		Return _LogNormalSigma.Value
	'	End If
	'	_LogNormalSigma = Math.Sqrt(Math.Log(Math.Sqrt((Me.Variance / Me.Mean ^ 2)) + 1))
	'	Return _LogNormalSigma.Value
	'End Function
End Class
