Imports System.Runtime.CompilerServices
Imports MathNet.Numerics.Distributions

Namespace MathPlus.Probability
	''' <summary>
	''' Utilities for mapping between probability space [0,1] and real-value axes
	''' using Normal (Gaussian) and Log-Normal distributions.
	''' 
	''' Recommended API:
	'''  • ProbabilityToGaussianX(p, rangeX)
	'''  • ProbabilityToLogNormalX(p, [mu], [sigma], [maxX])
	'''  • ProbabilityToGaussianScale(p, scaleOfX)          ' tail-expansion in [0,1]
	'''  • GaussianScaleToProbability(pPrime, scaleOfX)      ' inverse of the above
	'''  • GaussianCdf(x, [mean], [standardDeviation])
	'''  • LogNormalCdf(x, [mu], [sigma])
	'''  • GaussianPdf(x, [mean], [standardDeviation])
	'''  • LogNormalPdf(x, [mu], [sigma])
	''' 
	''' Extension helpers for IEnumerable(Of Double):
	'''  • probs.ToGaussianX(rangeX)
	'''  • probs.ToGaussianScale(scaleOfX)
	'''  • scaled.FromGaussianScale(scaleOfX)
	''' 
	''' Legacy-compatible wrappers are provided below so you don’t need to refactor callers immediately.
	''' </summary>
	Public Module ProbabilityMapping

		Private Const EPS As Double = 0.000000000001 ' 1e-12

#Region "Recommended API (use these going forward)"
		''' <summary>
		''' Maps a probability in [0,1] to a symmetric Gaussian-style X axis using the standard Normal quantile,
		''' then clamps to [-rangeX, +rangeX] for numerical stability and charting.
		''' 
		''' Typical use cases:
		'''  • Place probabilities on a fixed, symmetric X axis (for example, -3 to +3).
		'''  • Create a smooth nonlinear axis where the center is 0 and extremes compress at the edges.
		''' 
		''' Mapping behavior:
		'''  p = 0.5 returns 0
		'''  p near 0 returns a large negative X (clamped to -rangeX)
		'''  p near 1 returns a large positive X (clamped to +rangeX)
		''' </summary>
		''' <param name="p">Probability in [0,1]. Values slightly outside are clamped to [1E-12, 1-1E-12].</param>
		''' <param name="rangeX">Maximum absolute X. The result is clamped to [-rangeX, +rangeX].</param>
		Public Function ProbabilityToGaussianX(p As Double, rangeX As Double) As Double
			Dim pp = Math.Max(EPS, Math.Min(1.0 - EPS, p))
			Dim x = Normal.InvCDF(0.0, 1.0, pp)
			If x > rangeX Then
				x = rangeX
			ElseIf x < -rangeX Then
				x = -rangeX
			End If
			Return x
		End Function

		''' <summary>
		''' Maps a probability in [0,1] to a positive, skewed X axis using the Log-Normal quantile.
		''' Optional clamping prevents extremely large right-tail values.
		''' 
		''' Typical use cases:
		'''  • Map probabilities to a positive axis (growth, multiplicative factors, log-scale visuals).
		'''  • Generate log-normal values by inverse transform sampling.
		''' 
		''' Mapping behavior:
		'''  p = 0.5 returns approximately exp(mu)
		'''  p near 0 returns values near 0
		'''  p near 1 returns very large values (optionally clamped by maxX)
		''' </summary>
		''' <param name="p">Probability in [0,1]. Values slightly outside are clamped to [1E-12, 1-1E-12].</param>
		''' <param name="mu">Mean of the underlying normal (default 0).</param>
		''' <param name="sigma">Standard deviation of the underlying normal (default 1).</param>
		''' <param name="maxX">Optional maximum X; if provided, the result is capped at this value.</param>
		Public Function ProbabilityToLogNormalX(p As Double,
																						Optional mu As Double = 0.0,
																						Optional sigma As Double = 1.0,
																						Optional maxX As Double? = Nothing) As Double
			Dim pp = Math.Max(EPS, Math.Min(1.0 - EPS, p))
			Dim x = LogNormal.InvCDF(mu, sigma, pp)
			If maxX.HasValue AndAlso x > maxX.Value Then x = maxX.Value
			Return x
		End Function

		''' <summary>
		''' Nonlinearly rescales a probability in [0,1] so that small changes near 0 or 1
		''' become more pronounced while mid-range values are relatively compressed.
		''' 
		''' Implementation:
		'''  1) Convert p to X on a standard Normal axis, clamped to ±ScaleOfX.
		'''  2) Linearly reproject X back to [0,1]: p' = (X / (2*ScaleOfX)) + 0.5.
		''' 
		''' Typical use cases:
		'''  • Emphasize subtle changes at high certainty (e.g., 0.999 → 0.9999) for visualization or thresholds.
		'''  • Make tail regions more “human-visible” without leaving the [0,1] range.
		''' 
		''' Mapping behavior:
		'''  p = 0.5 returns 0.5
		'''  Larger ScaleOfX ⇒ stronger tail expansion (greater sensitivity near 0 and 1)
		''' </summary>
		Public Function ProbabilityToGaussianScale(p As Double, ScaleOfX As Double) As Double
			Dim x = ProbabilityToGaussianX(p, ScaleOfX)     ' X ∈ [-ScaleOfX, +ScaleOfX]
			Return (x / (2 * ScaleOfX)) + 0.5               ' back to [0,1]
		End Function

		''' <summary>
		''' Inverse of ProbabilityToGaussianScale. Recovers the original linear probability from the
		''' Gaussian-scaled value p' in [0,1].
		''' 
		''' Implementation:
		'''  1) Undo the linear reproject: X = (p' - 0.5) * (2*ScaleOfX), clamped to ±ScaleOfX.
		'''  2) Map X back to probability p using the standard Normal CDF.
		''' 
		''' Note: If the forward transform clamped at ±ScaleOfX, this inverse mirrors that behavior.
		''' </summary>
		Public Function GaussianScaleToProbability(pPrime As Double, ScaleOfX As Double) As Double
			Dim x = (pPrime - 0.5) * (2 * ScaleOfX)
			If x > ScaleOfX Then
				x = ScaleOfX
			ElseIf x < -ScaleOfX Then
				x = -ScaleOfX
			End If

			Return Normal.CDF(0.0, 1.0, x)
		End Function

		''' <summary>
		''' Gaussian (Normal) cumulative probability P(X ≤ x).
		''' 
		''' Typical use cases:
		'''  • Convert a real value to probability space.
		'''  • Evaluate tail probabilities and thresholds for a Normal model.
		''' </summary>
		Public Function GaussianCDF(x As Double, Optional mean As Double = 0.0, Optional standardDeviation As Double = 1.0) As Double
			Return New Normal(mean, standardDeviation).CumulativeDistribution(x)
		End Function

		''' <summary>
		''' Log-Normal cumulative probability P(X ≤ x).
		''' 
		''' Typical use cases:
		'''  • Convert a positive real value to probability space on a log-normal model.
		'''  • Evaluate tail probabilities for multiplicative processes.
		''' </summary>
		Public Function LogNormalCDF(x As Double, Optional mu As Double = 0.0, Optional sigma As Double = 1.0) As Double
			Return LogNormal.CDF(mu, sigma, x)
		End Function

		''' <summary>
		''' Gaussian (Normal) probability density at x.
		''' Useful for weighting/likelihoods and visualizing bell-curve shapes.
		''' </summary>
		Public Function GaussianPDF(x As Double,
																Optional mean As Double = 0.0,
																Optional standardDeviation As Double = 1.0) As Double
			Return New Normal(mean, standardDeviation).Density(x)
		End Function

		''' <summary>
		''' Log-Normal probability density at x.
		''' Useful for positive, skewed processes and log-scale charts.
		''' </summary>
		Public Function LogNormalPDF(x As Double,
																 Optional mu As Double = 0.0,
																 Optional sigma As Double = 1.0) As Double
			Return LogNormal.PDF(mu, sigma, x)
		End Function
#End Region

#Region "Enumerable extensions (inline helpers for lists/sequences)"
		''' <summary>
		''' Converts each probability to a symmetric Gaussian X and clamps to ±rangeX.
		''' Example: probs.ToGaussianX(rangeX:=3)
		''' </summary>
		<Extension>
		Public Function ToGaussianX(probs As IEnumerable(Of Double), rangeX As Double) As List(Of Double)
			Dim out As New List(Of Double)
			For Each p In probs
				out.Add(ProbabilityToGaussianX(p, rangeX))
			Next
			Return out
		End Function

		''' <summary>
		''' Applies the Gaussian tail-expansion transform to each probability:
		''' p → p' in [0,1] with more sensitivity near 0 and 1.
		''' Example: probs.ToGaussianScale(scaleOfX:=3)
		''' </summary>
		<Extension>
		Public Function ToGaussianScale(probs As IEnumerable(Of Double), scaleOfX As Double) As List(Of Double)
			Dim out As New List(Of Double)
			For Each p In probs
				out.Add(ProbabilityToGaussianScale(p, scaleOfX))
			Next
			Return out
		End Function

		''' <summary>
		''' Inverse of ToGaussianScale. Recovers the original linear probabilities from Gaussian-scaled values.
		''' Example: scaled.FromGaussianScale(scaleOfX:=3)
		''' </summary>
		<Extension>
		Public Function FromGaussianScale(scaled As IEnumerable(Of Double), scaleOfX As Double) As List(Of Double)
			Dim out As New List(Of Double)
			For Each pPrime In scaled
				out.Add(GaussianScaleToProbability(pPrime, scaleOfX))
			Next
			Return out
		End Function
#End Region

#Region "Legacy-compatible wrappers (keep callers working; switch to the API above when convenient)"
		''' <summary>
		''' Converts a probability to a standard Normal X and clamps to [-RangeOfX, +RangeOfX].
		''' Prefer <c>ProbabilityToGaussianX</c> for new code.
		''' </summary>
		Public Function InverseNormal(ProbabilityValue As Double, RangeOfX As Double) As Double
			Return ProbabilityToGaussianX(ProbabilityValue, RangeOfX)
		End Function

		''' <summary>
		''' Converts a probability to a Log-Normal X (μ=0, σ=1).
		''' Prefer <c>ProbabilityToLogNormalX</c> (with parameters if needed) for new code.
		''' </summary>
		Public Function InverseLogNormal(ProbabilityValue As Double) As Double
			Return ProbabilityToLogNormalX(ProbabilityValue, 0.0, 1.0, Nothing)
		End Function

		''' <summary>
		''' Standard Gaussian CDF approximation P(X ≤ x) for μ=0, σ=1.
		''' Prefer <c>GaussianCdf(x)</c> for new code.
		''' </summary>
		Public Function CDFGaussian(x As Double) As Double
			' Use MathNet for accuracy and simplicity
			Return GaussianCdf(x, 0.0, 1.0)
		End Function

		''' <summary>
		''' Gaussian CDF P(X ≤ x) for specified mean and standard deviation.
		''' Prefer <c>GaussianCdf(x, mean, standardDeviation)</c> for new code.
		''' </summary>
		Public Function CDFGaussian(Mean As Double, StandardDeviation As Double, X As Double) As Double
			Return GaussianCdf(X, Mean, StandardDeviation)
		End Function

		''' <summary>
		''' Gaussian inverse CDF (quantile). Returns X such that P(X' ≤ X) = Probability.
		''' Prefer <c>ProbabilityToGaussianX</c> for mapping p to X on a symmetric axis.
		''' </summary>
		Public Function InverseCDFGaussian(Mean As Double, StandardDeviation As Double, Probability As Double) As Double
			Dim pp = Math.Max(EPS, Math.Min(1.0 - EPS, Probability))
			Return New Normal(Mean, StandardDeviation).InverseCumulativeDistribution(pp)
		End Function

		''' <summary>
		''' Legacy list wrapper forwarding to the extension ToGaussianScale.
		''' </summary>
		Public Function ToGaussianScale(ByVal DataProbability As IList(Of Double),
																		ByVal GaussianPropabilityRangeOfX As Double) As List(Of Double)
			Return CType(DataProbability, IEnumerable(Of Double)).ToGaussianScale(GaussianPropabilityRangeOfX)
		End Function
#End Region
	End Module
End Namespace
