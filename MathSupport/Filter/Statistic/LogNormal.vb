Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports MathNet.Numerics.Distributions
Imports MathNet.Numerics.Random
Imports MathNet.Numerics.Statistics
Imports MathNet.Numerics.Threading

Namespace MathNet.Numerics.Distributions
	''' <summary>
	''' Continuous Univariate Log-Normal distribution.
	''' For details about this distribution, see
	''' <a href="http://en.wikipedia.org/wiki/Log-normal_distribution">Wikipedia - Log-Normal distribution</a>.
	''' </summary>
	Public Class LogNormal
		'Implements IContinuousDistribution

		Private _random As System.Random
		Private ReadOnly _mu As Double
		Private ReadOnly _sigma As Double

		''' <summary>
		''' Initializes a new instance of the <see cref="LogNormal"/> class.
		''' The distribution will be initialized with the default <seealso cref="System.Random"/>
		''' random number generator.
		''' </summary>
		''' <param name="mu">The log-scale (μ) of the logarithm of the distribution.</param>
		''' <param name="sigma">The shape (σ) of the logarithm of the distribution. Range: σ ≥ 0.</param>
		Public Sub New(mu As Double, sigma As Double)
			If Not IsValidParameterSet(mu, sigma) Then
				Throw New ArgumentException("Invalid parametrization for the distribution.")
			End If

			_random = SystemRandomSource.Default
			_mu = mu
			_sigma = sigma
		End Sub

		''' <summary>
		''' Initializes a new instance of the <see cref="LogNormal"/> class.
		''' The distribution will be initialized with the default <seealso cref="System.Random"/>
		''' random number generator.
		''' </summary>
		''' <param name="mu">The log-scale (μ) of the distribution.</param>
		''' <param name="sigma">The shape (σ) of the distribution. Range: σ ≥ 0.</param>
		''' <param name="randomSource">The random number generator which is used to draw random samples.</param>
		Public Sub New(mu As Double, sigma As Double, randomSource As System.Random)
			If Not IsValidParameterSet(mu, sigma) Then
				Throw New ArgumentException("Invalid parametrization for the distribution.")
			End If

			_random = If(randomSource, SystemRandomSource.Default)
			_mu = mu
			_sigma = sigma
		End Sub

		''' <summary>
		''' Constructs a log-normal distribution with the desired mu and sigma parameters.
		''' </summary>
		Public Shared Function WithMuSigma(mu As Double, sigma As Double, Optional randomSource As System.Random = Nothing) As LogNormal
			Return New LogNormal(mu, sigma, randomSource)
		End Function

		''' <summary>
		''' Constructs a log-normal distribution with the desired mean and variance.
		''' </summary>
		Public Shared Function WithMeanVariance(mean As Double, var As Double, Optional randomSource As System.Random = Nothing) As LogNormal
			Dim sigma2 = Math.Log(var / (mean * mean) + 1.0)
			Return New LogNormal(Math.Log(mean) - sigma2 / 2.0, Math.Sqrt(sigma2), randomSource)
		End Function

		''' <summary>
		''' Estimates the log-normal distribution parameters from sample data with maximum-likelihood.
		''' </summary>
		Public Shared Function Estimate(samples As IEnumerable(Of Double), Optional randomSource As System.Random = Nothing) As LogNormal
			Dim muSigma = samples.Select(Function(s) Math.Log(s)).MeanStandardDeviation()
			Return New LogNormal(muSigma.Mean, muSigma.StandardDeviation, randomSource)
		End Function

		''' <summary>
		''' Tests whether the provided values are valid parameters for this distribution.
		''' </summary>
		Public Shared Function IsValidParameterSet(mu As Double, sigma As Double) As Boolean
			Return sigma >= 0.0 AndAlso Not Double.IsNaN(mu)
		End Function

		''' <summary>
		''' Gets the log-scale (μ) (mean of the logarithm) of the distribution.
		''' </summary>
		Public ReadOnly Property Mu As Double
			Get
				Return _mu
			End Get
		End Property

		''' <summary>
		''' Gets the shape (σ) (standard deviation of the logarithm) of the distribution. Range: σ ≥ 0.
		''' </summary>
		Public ReadOnly Property Sigma As Double
			Get
				Return _sigma
			End Get
		End Property

		''' <summary>
		''' Gets or sets the random number generator which is used to draw random samples.
		''' </summary>
		Public Property RandomSource As System.Random
			Get
				Return _random
			End Get
			Set(value As System.Random)
				_random = If(value, SystemRandomSource.Default)
			End Set
		End Property

		''' <summary>
		''' Gets the mean of the log-normal distribution.
		''' </summary>
		Public ReadOnly Property Mean As Double
			Get
				Return Math.Exp(_mu + (_sigma * _sigma / 2.0))
			End Get
		End Property

		''' <summary>
		''' Gets the variance of the log-normal distribution.
		''' </summary>
		Public ReadOnly Property Variance As Double
			Get
				Dim sigma2 = _sigma * _sigma
				Return (Math.Exp(sigma2) - 1.0) * Math.Exp(_mu + _mu + sigma2)
			End Get
		End Property

		''' <summary>
		''' Gets the standard deviation of the log-normal distribution.
		''' </summary>
		Public ReadOnly Property StdDev As Double
			Get
				Dim sigma2 = _sigma * _sigma
				Return Math.Sqrt((Math.Exp(sigma2) - 1.0) * Math.Exp(_mu + _mu + sigma2))
			End Get
		End Property

		''' <summary>
		''' Generates a sample from the log-normal distribution.
		''' </summary>
		Public Function Sample() As Double
			Return Math.Exp(Normal.Sample(_random, _mu, _sigma))
		End Function

		''' <summary>
		''' Generates a sequence of samples from the log-normal distribution.
		''' </summary>
		Public Function Samples() As IEnumerable(Of Double)
			Return Normal.Samples(_random, _mu, _sigma).Select(Function(x) Math.Exp(x))
		End Function
	End Class
End Namespace
