''' <summary>
''' Hold the Probability state of a stock being classified as "Up" or "Down" 
''' based on the probability of the stock price going up. 
''' </summary>
Public Class ProbabilityState
	Implements IProbabilityState
	Public Sub New(ProbabilityUp As Double, StateFlipProbability As Double)
		Me.ProbabilityUp = ProbabilityUp
		Me.StateFlipProbability = StateFlipProbability
	End Sub

	Public ReadOnly Property ProbabilityUp As Double Implements IProbabilityState.ProbabilityUp

	''' <summary>
	''' Probability that the current state classification
	''' (Up or Down) could reverse.
	''' Range: 0.0 to 0.5.
	''' A value near 0 indicates high confidence.
	''' A value near 0.5 indicates low confidence.
	''' </summary>
	Public ReadOnly Property StateFlipProbability As Double Implements IProbabilityState.StateFlipProbability
End Class

Public Interface IProbabilityState
	ReadOnly Property ProbabilityUp As Double

	''' <summary>
	''' Probability that the current state classification
	''' (Up or Down) could reverse.
	''' Range: 0.0 to 0.5.
	''' A value near 0 indicates high confidence.
	''' A value near 0.5 indicates low confidence.
	''' </summary>
	ReadOnly Property StateFlipProbability As Double
End Interface

