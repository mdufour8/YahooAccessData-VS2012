Option Strict On
Option Infer On
Option Explicit On

Imports System.Collections.Generic

''' <summary>
''' Read-only stochastic/brownian data surface. 
''' Replace enum-based access with discoverable properties and XML-docs.
''' </summary>
Public Interface IStochasticBrownianData

	' --- Core stochastic lines ---
	ReadOnly Property FastSlow As IList(Of Double)
	ReadOnly Property Fast As IList(Of Double)
	ReadOnly Property Slow As IList(Of Double)

	' --- Price bands & predictions ---
	ReadOnly Property PriceBandHigh As IList(Of Double)
	ReadOnly Property PriceBandLow As IList(Of Double)
	ReadOnly Property PriceBandHighPrediction As IList(Of Double)
	ReadOnly Property PriceBandLowPrediction As IList(Of Double)

	' --- Volatility & ranges ---
	ReadOnly Property RangeVolatility As IList(Of Double)
	ReadOnly Property RangeVolatilityFromPreviousCloseToOpen As IList(Of Double)
	ReadOnly Property RangeVolatilityFromPreviousCloseToOpenRatio As IList(Of Double)
	ReadOnly Property RangeVolatilityFromOpenToClose As IList(Of Double)
	ReadOnly Property RangeVolatilityRegulatedFromPreviousCloseToOpen As IList(Of Double)
	ReadOnly Property PriceStandardVolatility As IList(Of Double)

	' --- Probability & time-probability ---
	ReadOnly Property ProbabilityHigh As IList(Of Double)
	ReadOnly Property ProbabilityLow As IList(Of Double)
	ReadOnly Property ProbabilityFromBandVolatility As IList(Of Double)
	ReadOnly Property TimeProbabilityOfPriceVolatility As IList(Of Double)
	ReadOnly Property ProbabilityPriceDailySigmaExceeded As IList(Of Double)
	ReadOnly Property ProbabilityPriceDailySigmaDoubleExceeded As IList(Of Double)
	ReadOnly Property ProbabilityOfPriceStochacticMedianWithGain As IList(Of Double)

	' --- Volatility PDF/LCR (empirical & simulated) ---
	ReadOnly Property PriceVolatilityPDF As IList(Of Double)
	ReadOnly Property PriceVolatilityLCR As IList(Of Double)
	ReadOnly Property PriceVolatilityPDFSimulated As IList(Of Double)
	ReadOnly Property PriceVolatilityLCRSimulated As IList(Of Double)

	' --- Volatility bands & derivatives ---
	ReadOnly Property PriceBandVolatilityHigh As IList(Of Double)
	ReadOnly Property PriceBandVolatilityLow As IList(Of Double)
	ReadOnly Property PriceBandVolatilityGain As IList(Of Double)

	' --- Median/peak models & predictions ---
	ReadOnly Property PriceProbabilityMedian As IList(Of Double)
	ReadOnly Property StochasticSlowFromPriceBandVolatilityLow As IList(Of Double)
	ReadOnly Property StochasticSlowFromPriceBandVolatilityHigh As IList(Of Double)
	ReadOnly Property StochasticSlowFromPricePeakMedian As IList(Of Double)             ' may throw NotSupportedException upstream
	ReadOnly Property PriceStochacticMedian As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGain As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainPrediction As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainHisteresisHigh As IList(Of Double)    ' not supported upstream
	ReadOnly Property PriceStochacticMedianWithGainHisteresisLow As IList(Of Double)     ' not supported upstream
	ReadOnly Property PriceStochacticMedianWithGainPredictionHigh As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainPredictionLow As IList(Of Double)

	' --- Median-based daily ranges (D+1) ---
	ReadOnly Property PriceStochacticMedianRangeDailyUp As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUp As IList(Of Double)
	ReadOnly Property PriceStochacticMedianRangeDailyDown As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyDown As IList(Of Double)

	' --- Median-based daily ranges (open-to-close + variants) ---
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpToOpen As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownToOpen As IList(Of Double)

	' --- Median-based daily ranges (D+2) ---
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpDay2 As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyDownDay2 As IList(Of Double)

	' --- Median-based daily ranges at sigma thresholds ---
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpAtSigma2 As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyLowAtSigma2 As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyUpAtSigma3 As IList(Of Double)
	ReadOnly Property PriceStochacticMedianWithGainRangeDailyLowAtSigma3 As IList(Of Double)

	' --- Regulated volatility & detectors ---
	ReadOnly Property PriceVolatilityRegulated As IList(Of Double)
	ReadOnly Property PriceVolatilityRegulatedPrediction As IList(Of Double)
	ReadOnly Property PriceVolatilityLastPointTrail As IList(Of Double)
	ReadOnly Property PriceVolatilityDetectorBalance As IList(Of Double)

	' --- Median-linked volatility ---
	ReadOnly Property PriceStochasticMedianVolatility As IList(Of Double)
	ReadOnly Property PriceStochasticMedianNextDayLow As IList(Of Double)
	ReadOnly Property PriceStochasticMedianNextDayHigh As IList(Of Double)

	' --- Gain metrics ---
	ReadOnly Property PriceGainPerYear As IList(Of Double)
	ReadOnly Property PriceGainPerYearDerivative As IList(Of Double)

	' --- Positive/negative volatility components ---
	ReadOnly Property PriceStochacticVolatilityPositive As IList(Of Double)
	ReadOnly Property PriceStochacticVolatilityNegative As IList(Of Double)
	ReadOnly Property PriceStochacticVolatilityPositiveToNegativeRatio As IList(Of Double)
	ReadOnly Property PriceStochacticVolatilityPositiveToNegativeRatioFiltered As IList(Of Double)
	ReadOnly Property PricePeakValueGainPrediction As IList(Of Double)

	''' <summary>
	''' Return the On-Balance Log scaled Volume (OBV) related data series.
	''' </summary>
	ReadOnly Property GetListOfPriceOBV As IList(Of Double)

	''' <summary>
	''' Return the RSI of the On-Balance Log Scaled Volume (OBVLS) related data series.
	''' </summary>
	ReadOnly Property GetListOfRSIOBV As IList(Of Double)

	ReadOnly Property ToList As IList(Of Double)

	ReadOnly Property Count As Integer

	ReadOnly Property Rate As Integer


End Interface

