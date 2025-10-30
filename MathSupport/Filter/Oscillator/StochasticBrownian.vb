Option Strict On
Option Infer On
Option Explicit On

Imports System.Collections.Generic
Imports YahooAccessData.MathPlus.Filter

Namespace MathPlus.Filter
	''' <summary>
	''' Facade over <see cref="FilterStochasticBrownian"/> exposing each stochastic series as a read-only property.
	''' </summary>
	''' <remarks>
	''' Properties are lazily fetched and cached per instance. Values reflect the current state of the underlying
	''' <see cref="FilterStochasticBrownian"/> source at the time of first access.
	''' </remarks>
	Public Class StochasticBrownian
		Implements IStochasticBrownianData

		Private ReadOnly _src As FilterStochasticBrownian
		Private ReadOnly _cache As New Dictionary(Of IStochastic.enuStochasticType, IList(Of Double))()

		Public Sub New(
			ByVal FilterRate As Integer,
			ByVal FilterOutputRate As Double,
			Optional ByVal IsFilterPeakEnabled As Boolean = False,
			Optional ByVal FilterVolatilityRate As Integer = FilterStochasticBrownian.FILTER_RATE_FOR_VOLATILITY,
			Optional ByVal FilterPeakRate As Integer = -1)

			_src = New FilterStochasticBrownian(
				FilterRate,
				FilterOutputRate,
				IsFilterPeakEnabled,
				FilterVolatilityRate,
				FilterPeakRate)

			_ToList = _src.ToList
		End Sub
		''' <summary>
		''' Create a facade over an existing <see cref="FilterStochasticBrownian"/> instance.
		''' </summary>
		''' <param name="stochastic">Source stochastic/brownian filter instance.</param>
		''' <exception cref="ArgumentNullException">If <paramref name="stochastic"/> is null.</exception>

		Public Sub New(stochastic As FilterStochasticBrownian)
			If stochastic Is Nothing Then Throw New ArgumentNullException(NameOf(stochastic))
			_src = stochastic
		End Sub

		' Centralized lazy cache getter.
		Private Function GetList(kind As IStochastic.enuStochasticType) As IList(Of Double)
			Dim list As IList(Of Double) = Nothing
			If _cache.TryGetValue(kind, list) = True Then Return list
			list = _src.ToList(kind)
			_cache(kind) = list
			Return list
		End Function

		''' <summary>Stochastic Fast/Slow combined series.</summary>
		Public ReadOnly Property FastSlow As IList(Of Double) Implements IStochasticBrownianData.FastSlow
			Get
				Return GetList(IStochastic.enuStochasticType.FastSlow)
			End Get
		End Property

		''' <summary>Stochastic Fast line.</summary>
		Public ReadOnly Property Fast As IList(Of Double) Implements IStochasticBrownianData.Fast
			Get
				Return GetList(IStochastic.enuStochasticType.Fast)
			End Get
		End Property

		''' <summary>Stochastic Slow line (often smoothed).</summary>
		Public ReadOnly Property Slow As IList(Of Double) Implements IStochasticBrownianData.Slow
			Get
				Return GetList(IStochastic.enuStochasticType.Slow)
			End Get
		End Property

		''' <summary>Upper price band.</summary>
		Public ReadOnly Property PriceBandHigh As IList(Of Double) Implements IStochasticBrownianData.PriceBandHigh
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandHigh)
			End Get
		End Property

		''' <summary>Lower price band.</summary>
		Public ReadOnly Property PriceBandLow As IList(Of Double) Implements IStochasticBrownianData.PriceBandLow
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandLow)
			End Get
		End Property

		''' <summary>Predicted upper price band.</summary>
		Public ReadOnly Property PriceBandHighPrediction As IList(Of Double) Implements IStochasticBrownianData.PriceBandHighPrediction
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandHighPrediction)
			End Get
		End Property

		''' <summary>Predicted lower price band.</summary>
		Public ReadOnly Property PriceBandLowPrediction As IList(Of Double) Implements IStochasticBrownianData.PriceBandLowPrediction
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandLowPrediction)
			End Get
		End Property

		''' <summary>General range volatility metric.</summary>
		Public ReadOnly Property RangeVolatility As IList(Of Double) Implements IStochasticBrownianData.RangeVolatility
			Get
				Return GetList(IStochastic.enuStochasticType.RangeVolatility)
			End Get
		End Property

		''' <summary>Probability of moving up (if supported in source).</summary>
		Public ReadOnly Property ProbabilityHigh As IList(Of Double) Implements IStochasticBrownianData.ProbabilityHigh
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityHigh)
			End Get
		End Property

		''' <summary>Probability of moving down (if supported in source).</summary>
		Public ReadOnly Property ProbabilityLow As IList(Of Double) Implements IStochasticBrownianData.ProbabilityLow
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityLow)
			End Get
		End Property

		''' <summary>Volatility derived upper band.</summary>
		Public ReadOnly Property PriceBandVolatilityHigh As IList(Of Double) Implements IStochasticBrownianData.PriceBandVolatilityHigh
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandVolatilityHigh)
			End Get
		End Property

		''' <summary>Volatility derived lower band.</summary>
		Public ReadOnly Property PriceBandVolatilityLow As IList(Of Double) Implements IStochasticBrownianData.PriceBandVolatilityLow
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandVolatilityLow)
			End Get
		End Property

		''' <summary>Volatility band gain (e.g., high−low or related gain metric).</summary>
		Public ReadOnly Property PriceBandVolatilityGain As IList(Of Double) Implements IStochasticBrownianData.PriceBandVolatilityGain
			Get
				Return GetList(IStochastic.enuStochasticType.PriceBandVolatilityGain)
			End Get
		End Property

		''' <summary>Standard (baseline) price volatility.</summary>
		Public ReadOnly Property PriceStandardVolatility As IList(Of Double) Implements IStochasticBrownianData.PriceStandardVolatility
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStandardVolatility)
			End Get
		End Property

		''' <summary>Price volatility PDF.</summary>
		Public ReadOnly Property PriceVolatilityPDF As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityPDF
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityPDF)
			End Get
		End Property

		''' <summary>Price volatility LCR (Level Crossing Rate e.g., CDF or tail metric).</summary>
		Public ReadOnly Property PriceVolatilityLCR As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityLCR
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityLCR)
			End Get
		End Property

		''' <summary>Simulated price volatility PDF.</summary>
		Public ReadOnly Property PriceVolatilityPDFSimulated As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityPDFSimulated
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityPDFSimulated)
			End Get
		End Property

		''' <summary>Simulated price volatility LCR.</summary>
		Public ReadOnly Property PriceVolatilityLCRSimulated As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityLCRSimulated
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityLCRSimulated)
			End Get
		End Property

		''' <summary>Median price probability curve.</summary>
		Public ReadOnly Property PriceProbabilityMedian As IList(Of Double) Implements IStochasticBrownianData.PriceProbabilityMedian
			Get
				Return GetList(IStochastic.enuStochasticType.PriceProbabilityMedian)
			End Get
		End Property

		''' <summary>Probability derived from band volatility.</summary>
		Public ReadOnly Property ProbabilityFromBandVolatility As IList(Of Double) Implements IStochasticBrownianData.ProbabilityFromBandVolatility
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityFromBandVolatility)
			End Get
		End Property

		''' <summary>Time probability of price volatility (hazard-like).</summary>
		Public ReadOnly Property TimeProbabilityOfPriceVolatility As IList(Of Double) Implements IStochasticBrownianData.TimeProbabilityOfPriceVolatility
			Get
				Return GetList(IStochastic.enuStochasticType.TimeProbabilityOfPriceVolatility)
			End Get
		End Property

		''' <summary>Slow stochastic computed from price band volatility (Low band).</summary>
		Public ReadOnly Property StochasticSlowFromPriceBandVolatilityLow As IList(Of Double) Implements IStochasticBrownianData.StochasticSlowFromPriceBandVolatilityLow
			Get
				Return GetList(IStochastic.enuStochasticType.StochasticSlowFromPriceBandVolatilityLow)
			End Get
		End Property

		''' <summary>Slow stochastic computed from price band volatility (High band).</summary>
		Public ReadOnly Property StochasticSlowFromPriceBandVolatilityHigh As IList(Of Double) Implements IStochasticBrownianData.StochasticSlowFromPriceBandVolatilityHigh
			Get
				Return GetList(IStochastic.enuStochasticType.StochasticSlowFromPriceBandVolatilityHigh)
			End Get
		End Property

		''' <summary>Slow stochastic from peak median (not supported by source will throw if called).</summary>
		Public ReadOnly Property StochasticSlowFromPricePeakMedian As IList(Of Double) Implements IStochasticBrownianData.StochasticSlowFromPricePeakMedian
			Get
				Return GetList(IStochastic.enuStochasticType.StochasticSlowFromPricePeakMedian)
			End Get
		End Property

		''' <summary>Median of the “price stochastic”.</summary>
		Public ReadOnly Property PriceStochacticMedian As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedian
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedian)
			End Get
		End Property

		''' <summary>Median with gain overlay.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGain As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGain
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGain)
			End Get
		End Property

		''' <summary>Next-day range (up) driven by median model.</summary>
		Public ReadOnly Property PriceStochacticMedianRangeDailyUp As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianRangeDailyUp
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyUp)
			End Get
		End Property

		''' <summary>Next-day range (up) with gain overlay.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUp As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUp
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUp)
			End Get
		End Property

		''' <summary>Next-day range (down) driven by median model.</summary>
		Public ReadOnly Property PriceStochacticMedianRangeDailyDown As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianRangeDailyDown
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyDown)
			End Get
		End Property

		''' <summary>Next-day range (down) with gain overlay.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyDown As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyDown
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDown)
			End Get
		End Property

		''' <summary>Probability that daily sigma is exceeded.</summary>
		Public ReadOnly Property ProbabilityPriceDailySigmaExceeded As IList(Of Double) Implements IStochasticBrownianData.ProbabilityPriceDailySigmaExceeded
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityPriceDailySigmaExceeded)
			End Get
		End Property

		''' <summary>Regulated price volatility series.</summary>
		Public ReadOnly Property PriceVolatilityRegulated As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityRegulated
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityRegulated)
			End Get
		End Property

		''' <summary>“Last point trail” of volatility (tracking tail).</summary>
		Public ReadOnly Property PriceVolatilityLastPointTrail As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityLastPointTrail
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityLastPointTrail)
			End Get
		End Property

		''' <summary>Detector balance for volatility prediction (e.g., error/convergence metric).</summary>
		Public ReadOnly Property PriceVolatilityDetectorBalance As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityDetectorBalance
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityDetectorBalance)
			End Get
		End Property

		''' <summary>Median+gain prediction series.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainPrediction As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainPrediction
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainPrediction)
			End Get
		End Property

		''' <summary>Median+gain prediction with hysteresis (HIGH) — not supported.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainHisteresisHigh As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainHisteresisHigh
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainHisteresisHigh)
			End Get
		End Property

		''' <summary>Median+gain prediction with hysteresis (LOW) — not supported.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainHisteresisLow As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainHisteresisLow
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainHisteresisLow)
			End Get
		End Property

		''' <summary>Prediction series for regulated volatility.</summary>
		Public ReadOnly Property PriceVolatilityRegulatedPrediction As IList(Of Double) Implements IStochasticBrownianData.PriceVolatilityRegulatedPrediction
			Get
				Return GetList(IStochastic.enuStochasticType.PriceVolatilityRegulatedPrediction)
			End Get
		End Property

		''' <summary>Peak value gain prediction series.</summary>
		Public ReadOnly Property PricePeakValueGainPrediction As IList(Of Double) Implements IStochasticBrownianData.PricePeakValueGainPrediction
			Get
				Return GetList(IStochastic.enuStochasticType.PricePeakValueGainPrediction)
			End Get
		End Property

		''' <summary>Volatility derived from median stochastic.</summary>
		Public ReadOnly Property PriceStochasticMedianVolatility As IList(Of Double) Implements IStochasticBrownianData.PriceStochasticMedianVolatility
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochasticMedianVolatility)
			End Get
		End Property

		''' <summary>Next day low based on median stochastic.</summary>
		Public ReadOnly Property PriceStochasticMedianNextDayLow As IList(Of Double) Implements IStochasticBrownianData.PriceStochasticMedianNextDayLow
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochasticMedianNextDayLow)
			End Get
		End Property

		''' <summary>Next day high based on median stochastic.</summary>
		Public ReadOnly Property PriceStochasticMedianNextDayHigh As IList(Of Double) Implements IStochasticBrownianData.PriceStochasticMedianNextDayHigh
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochasticMedianNextDayHigh)
			End Get
		End Property

		''' <summary>Median+gain prediction (LOW).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainPredictionLow As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainPredictionLow
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainPredictionLow)
			End Get
		End Property

		''' <summary>Median+gain prediction (HIGH).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainPredictionHigh As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainPredictionHigh
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainPredictionHigh)
			End Get
		End Property

		''' <summary>Probability that 2× daily sigma is exceeded.</summary>
		Public ReadOnly Property ProbabilityPriceDailySigmaDoubleExceeded As IList(Of Double) Implements IStochasticBrownianData.ProbabilityPriceDailySigmaDoubleExceeded
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityPriceDailySigmaDoubleExceeded)
			End Get
		End Property

		''' <summary>Day+2: median+gain range (up).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpDay2 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUpDay2
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpDay2)
			End Get
		End Property

		''' <summary>Day+2: median+gain range (down).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownDay2 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyDownDay2
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownDay2)
			End Get
		End Property

		''' <summary>Range to next open (up scenario).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpToOpen As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUpToOpen
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpToOpen)
			End Get
		End Property

		''' <summary>Range to next open (down scenario).</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownToOpen As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyDownToOpen
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownToOpen)
			End Get
		End Property

		''' <summary>Volatility from previous close → open.</summary>
		Public ReadOnly Property RangeVolatilityFromPreviousCloseToOpen As IList(Of Double) Implements IStochasticBrownianData.RangeVolatilityFromPreviousCloseToOpen
			Get
				Return GetList(IStochastic.enuStochasticType.RangeVolatilityFromPreviousCloseToOpen)
			End Get
		End Property

		''' <summary>Ratio form of previous close → open volatility.</summary>
		Public ReadOnly Property RangeVolatilityFromPreviousCloseToOpenRatio As IList(Of Double) Implements IStochasticBrownianData.RangeVolatilityFromPreviousCloseToOpenRatio
			Get
				Return GetList(IStochastic.enuStochasticType.RangeVolatilityFromPreviousCloseToOpenRatio)
			End Get
		End Property

		''' <summary>Intraday open → close volatility.</summary>
		Public ReadOnly Property RangeVolatilityFromOpenToClose As IList(Of Double) Implements IStochasticBrownianData.RangeVolatilityFromOpenToClose
			Get
				Return GetList(IStochastic.enuStochasticType.RangeVolatilityFromOpenToClose)
			End Get
		End Property

		''' <summary>Range (up) from open to close with gain overlay.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose)
			End Get
		End Property

		''' <summary>Range (down) from open to close with gain overlay.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose)
			End Get
		End Property

		''' <summary>Probability tied to median+gain model.</summary>
		Public ReadOnly Property ProbabilityOfPriceStochacticMedianWithGain As IList(Of Double) Implements IStochasticBrownianData.ProbabilityOfPriceStochacticMedianWithGain
			Get
				Return GetList(IStochastic.enuStochasticType.ProbabilityOfPriceStochacticMedianWithGain)
			End Get
		End Property

		''' <summary>Annualized gain estimate.</summary>
		Public ReadOnly Property PriceGainPerYear As IList(Of Double) Implements IStochasticBrownianData.PriceGainPerYear
			Get
				Return GetList(IStochastic.enuStochasticType.PriceGainPerYear)
			End Get
		End Property

		''' <summary>Derivative of annualized gain.</summary>
		Public ReadOnly Property PriceGainPerYearDerivative As IList(Of Double) Implements IStochasticBrownianData.PriceGainPerYearDerivative
			Get
				Return GetList(IStochastic.enuStochasticType.PriceGainPerYearDerivative)
			End Get
		End Property

		''' <summary>Day+1 “up” boundary at 2σ.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpAtSigma2 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUpAtSigma2
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpAtSigma2)
			End Get
		End Property

		''' <summary>Day+1 “down” boundary at 2σ.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyLowAtSigma2 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyLowAtSigma2
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyLowAtSigma2)
			End Get
		End Property

		''' <summary>Day+1 “up” boundary at 3σ.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpAtSigma3 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyUpAtSigma3
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpAtSigma3)
			End Get
		End Property

		''' <summary>Day+1 “down” boundary at 3σ.</summary>
		Public ReadOnly Property PriceStochacticMedianWithGainRangeDailyLowAtSigma3 As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticMedianWithGainRangeDailyLowAtSigma3
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyLowAtSigma3)
			End Get
		End Property

		''' <summary>Regulated range from previous close → open.</summary>
		Public ReadOnly Property RangeVolatilityRegulatedFromPreviousCloseToOpen As IList(Of Double) Implements IStochasticBrownianData.RangeVolatilityRegulatedFromPreviousCloseToOpen
			Get
				Return GetList(IStochastic.enuStochasticType.RangeVolatilityRegulatedFromPreviousCloseToOpen)
			End Get
		End Property

		''' <summary>“Positive” component of stochastic volatility.</summary>
		Public ReadOnly Property PriceStochacticVolatilityPositive As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticVolatilityPositive
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticVolatilityPositive)
			End Get
		End Property

		''' <summary>“Negative” component of stochastic volatility.</summary>
		Public ReadOnly Property PriceStochacticVolatilityNegative As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticVolatilityNegative
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticVolatilityNegative)
			End Get
		End Property

		''' <summary>Ratio of positive/negative stochastic volatility.</summary>
		Public ReadOnly Property PriceStochacticVolatilityPositiveToNegativeRatio As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticVolatilityPositiveToNegativeRatio
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticVolatilityPositiveToNegativeRatio)
			End Get
		End Property

		''' <summary>Filtered ratio of positive/negative stochastic volatility.</summary>
		Public ReadOnly Property PriceStochacticVolatilityPositiveToNegativeRatioFiltered As IList(Of Double) Implements IStochasticBrownianData.PriceStochacticVolatilityPositiveToNegativeRatioFiltered
			Get
				Return GetList(IStochastic.enuStochasticType.PriceStochacticVolatilityPositiveToNegativeRatioFiltered)
			End Get
		End Property

		Public Sub SetListOfPriceOBV(Value As IList(Of Double))
			_GetListOfPriceOBV = Value
		End Sub


		Private _GetListOfPriceOBV As IList(Of Double)

		''' <summary>
		''' Return the On-Balance Log scaled Volume (OBV) related data series.
		''' </summary>
		Public ReadOnly Property GetListOfPriceOBV As IList(Of Double) Implements IStochasticBrownianData.GetListOfPriceOBV
			Get
				Return _GetListOfPriceOBV
			End Get
		End Property

		Public Sub SetListOfRSIOBV(Value As IList(Of Double))
			_GetListOfRSIOBV = Value
		End Sub


		Private _GetListOfRSIOBV As IList(Of Double)

		Public ReadOnly Property GetListOfRSIOBV As IList(Of Double) Implements IStochasticBrownianData.GetListOfRSIOBV
			Get
				Return _GetListOfRSIOBV
			End Get
		End Property

		Public Sub SetToList(Value As IList(Of Double))
			_ToList = Value
		End Sub

		Private _ToList As IList(Of Double)
		Public ReadOnly Property ToList As IList(Of Double) Implements IStochasticBrownianData.ToList
			Get
				Return _ToList
			End Get
		End Property


		Public ReadOnly Property Count As Integer Implements IStochasticBrownianData.Count
			Get
				Return _ToList.Count
			End Get
		End Property

		Public ReadOnly Property Rate As Integer Implements IStochasticBrownianData.Rate
			Get
				Return _src.Rate
			End Get
		End Property
	End Class
End Namespace