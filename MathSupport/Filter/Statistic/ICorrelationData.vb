Public Interface ICorrelationData
	ReadOnly Property Correlation As Double
	ReadOnly Property JoinProbability As Double
	ReadOnly Property RegressionCoefficient As Double
	ReadOnly Property Mean As Double
	ReadOnly Property StandardDeviation As Double
End Interface

Public Class CorrelationData
	Implements ICorrelationData
	Public Sub New(
		Correlation As Double,
		JoinProbability As Double,
		RegressionCoefficient As Double,
		Mean As Double,
		StandardDeviation As Double)

		Me.Correlation = Correlation
		Me.JoinProbability = JoinProbability
		Me.RegressionCoefficient = RegressionCoefficient
		Me.Mean = Mean
		Me.StandardDeviation = StandardDeviation
	End Sub

	Public ReadOnly Property Correlation As Double Implements ICorrelationData.Correlation
	Public ReadOnly Property JoinProbability As Double Implements ICorrelationData.JoinProbability
	Public ReadOnly Property RegressionCoefficient As Double Implements ICorrelationData.RegressionCoefficient
	Public ReadOnly Property Mean As Double Implements ICorrelationData.Mean
	Public ReadOnly Property StandardDeviation As Double Implements ICorrelationData.StandardDeviation
End Class
