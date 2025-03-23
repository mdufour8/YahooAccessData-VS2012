'Namespace MathPlus
Public Interface IStatistical
	ReadOnly Property Mean As Double
	ReadOnly Property StandardDeviation As Double
	ReadOnly Property Variance As Double
	ReadOnly Property High As Double
	ReadOnly Property Low As Double
	ReadOnly Property NumberPoint As Integer
	Property ValueLast As Double

	'Function LogNormalMu() As Double

	'Function LogNormalSigma() As Double

	''' <summary>
	''' Output the ratio between the ValueLast and the Standard Deviation on a Gaussian probability scaled between [0,1] or optionally between [-1,1] 
	''' </summary>
	''' <returns></returns>
	Function ToGaussianScale(Optional ScaleToSignedUnit As Boolean = False) As Double
	Function Copy() As IStatistical
	Sub CopyTo(ByVal Value As IStatistical)
	Sub Add(ByVal Value As IStatistical)
End Interface
'End Namespace