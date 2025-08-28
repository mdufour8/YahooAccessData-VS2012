Option Explicit On
Option Strict On
Option Infer On

#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
Imports System.Runtime.CompilerServices

#End Region


Namespace MathPlus.Filter
	Public Class FilterStochasticBrownian
		Implements IStochastic
		Implements IStochastic1
		Implements IStochastic2

		Implements IFilterRunAsync
		Implements IStochasticPriceGain
		Public Enum enuStatisticDistribution
			Volatility
			VolatilitySimulated
			VolatilityPositive
			VolatilityNegative
			VolatilityDistributionPositive
			VolatilityDistributionNegative
		End Enum

		Private Const STATISTIC_VOLATILITY_WINDOWS_SIZE As Integer = 1 * NUMBER_WORKDAY_PER_YEAR
		'Private Const STATISTIC_VOLATILITY_WINDOWS_SIZE As Integer = 10
		Private Const STATISTIC_DB_MINIMUM As Integer = -20
		Private Const STATISTIC_DB_MAXIMUM As Integer = 0
		Private Const STATISTIC_BUCKET_RANGE As Integer = STATISTIC_DB_MAXIMUM - STATISTIC_DB_MINIMUM
		Private Const STATISTIC_BUCKET_NUMBER As Integer = STATISTIC_BUCKET_RANGE + 1
		Private Const STATISTIC_BUCKET_TO_ARRAY_MAP As Integer = 1 - STATISTIC_DB_MINIMUM
		Private Const THRESHOLD_LEVEL As Double = 0.5

		'https://en.wikipedia.org/wiki/68%E2%80%9395%E2%80%9399.7_rule

		Private Const FILTER_PLL_DETECTOR_COUNT_LIMIT As Integer = 20
		Private Const FILTER_PLL_DETECTOR_ERROR_LIMIT As Double = 0.0001
		Private Const FILTER_RATE_FOR_VOLATILITY As Integer = 30

		Private MyRate As Integer
		Private MyRateOutput As Double
		Private MyRatePreFilter As Integer
		Private MyLogNormalForVolatilityStatistic As MathNet.Numerics.Distributions.LogNormal
		Private MyFilterVolatilityYangZhangForStatistic As FilterVolatilityYangZhang
		Private MyFilterVolatilityYangZhangForStatisticLastPointTrail As FilterVolatilityYangZhang
		Private MyListOfValue As IList(Of IPriceVol)
		Private MyListOfPriceProbabilityMedian As List(Of Double)
		Private MyListOfProbabilityDailySigmaExcess As List(Of Double)
		Private MyListOfProbabilityDailySigmaDoubleExcess As List(Of Double)
		Private MyListOfProbabilityBandHigh As List(Of Double)
		Private MyListOfProbabilityBandLow As List(Of Double)
		Private MyListOfPriceVolatilityGain As List(Of Double)
		Private MyListOfPriceVolatilityHigh As List(Of Double)
		Private MyListOfPriceVolatilityLow As List(Of Double)
		Private MyListOfPriceVolatilityTimeProbability As List(Of Double)
		Private MyListOfProbabilityPDF() As Integer
		Private MyListOfProbabilityLCR() As Integer
		'Private MyFilterLPForPrice As IFilter
		Private MyFilterLPForPrice As FilterLowPassPLL
		Private MyFilterLPForProbabilityFromBandVolatility As FilterLowPassExp
		Private MyFilterLPForStochasticFromPriceVolatilityHigh As IFilter
		Private MyFilterLPForStochasticFromPriceVolatilityLow As IFilter
		Private MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain As IList(Of Double)
		Private MyListForVolatilityRegulatedPreviousCloseToOpenWithGain As IList(Of Double)
		Private MyListForVolatilityRegulatedFromOpenToCloseWithGain As IList(Of Double)
		Private MyListForVolatilityRegulatedNoGainFromOpenToClose As IList(Of Double)
		Private MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain As FilterLowPassExpPredict
		Private MyListForVolatilityDetectorBalance As IList(Of Double)
		Private MyValueLast As IPriceVol
		Private MyProcessorCount As Integer

		'Private MyFilterStochastic As FilterStochastic
		Private MyMeasurePeakValueRange As MeasurePeakValueRange
		Private MyMeasurePeakValueRangeUsingNoPeakFilter As MeasurePeakValueRange
		Private MyListOfPeakValueGainPrediction As IList(Of Double)

		Private MyFilterPLLForGain As FilterLowPassPLL
		Private MyFilterPLLForGainPrediction As FilterLowPassPLL

		Private MyListOfPriceNextDailyHigh As IList(Of Double)
		Private MyListOfPriceNextDailyLow As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGain As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGain As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGainAtSigma2 As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGainAtSigma3 As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGainAtSigma2 As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGainAtSigma3 As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGainK2 As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGainK2 As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen As IList(Of Double)
		Private MyListOfPriceNextDailyHighWithGainOpenToClose As IList(Of Double)
		Private MyListOfPriceNextDailyLowWithGainOpenToClose As IList(Of Double)
		Private MyPLLErrorDetectorForPriceStochacticMedian As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGain As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh As FilterPLLDetectorForCDFToZero
		Private MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow As FilterPLLDetectorForCDFToZero

		Private MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain As FilterPLLDetectorForVolatilitySigma
		Private MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToOpenWithGain As FilterPLLDetectorForVolatilitySigma
		Private MyFilterForVolatilityStatistic As FilterStatistical
		Private MyStatisticalDistributionForVolatility As StatisticalDistribution
		Private MyStatisticalDistributionForVolatilityPositive As StatisticalDistribution
		Private MyStatisticalDistributionForVolatilityNegative As StatisticalDistribution
		Private MyStatisticalDistributionForVolatilitySimulation As StatisticalDistribution
		Private MyStatisticalDistributionForPrice As StatisticalDistribution
		Private MyQueueOfDailyRangeSigmaExcess As Queue(Of Tuple(Of Integer, Integer, Integer))
		Private MySumOfDailyRangeSigmaExcess As Integer
		Private MySumOfDailyRangeSigmaExcessHigh As Integer
		Private MySumOfDailyRangeSigmaExcessLow As Integer
		Private MyRateForVolatility As Integer
		Private MyStockPriceVolatilityPredictionBand() As StockPriceVolatilityPredictionBand
		Private MyStockPriceVolatilityPredictionBandWithGain() As StockPriceVolatilityPredictionBand
		Private MyStockPriceVolatilityPredictionBandWithGainCloseToOpen() As StockPriceVolatilityPredictionBand
		Private MyStatisticRangeOfExcess As StatisticRangeExcess

		Private MyStocFastSlowLast As Double
		Private MyListOfStochasticFast As ListScaled
		Private MyListOfStochasticFastSlow As ListScaled
		Private MyListOfPriceBandHigh As List(Of Double)
		Private MyListOfPriceBandLow As List(Of Double)
		Private MyListOfPriceBandHighPrediction As List(Of Double)
		Private MyListOfPriceBandLowPrediction As List(Of Double)

		Private MyListOfProbabilityOfStockMedian As List(Of Double)

		Private MyFilterVolatilityForPositifNegatif As Filter.FilterVolatilityYangZhang

		Private MyListOfPriceRangeVolatility As IList(Of Double)
		Private MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio As IList(Of Double)
		Private MyFilterLPOfStochasticSlow As IFilter
		Private MyStocFastLast As Double
		Private MyStocRangeVolatility As Double
		Private MyStochasticPriceGain As StochasticPriceGain


#Region "New"
		''' <summary>
		''' 
		''' </summary>
		''' <param name="FilterRate"></param>
		''' <param name="FilterOutputRate"></param>
		''' <param name="IsFilterPeakEnabled"></param>
		''' <param name="FilterVolatilityRate"></param>
		''' <param name="FilterPeakRate">
		''' A filter peak rate of -1 lock it's value to the FilterRate value
		''' </param>
		Public Sub New(
									ByVal FilterRate As Integer,
									ByVal FilterOutputRate As Double,
									Optional ByVal IsFilterPeakEnabled As Boolean = False,
									Optional ByVal FilterVolatilityRate As Integer = FILTER_RATE_FOR_VOLATILITY,
									Optional ByVal FilterPeakRate As Integer = -1)

			MyProcessorCount = Environment.ProcessorCount
			If MyProcessorCount < 2 Then MyProcessorCount = 2


			MyRate = FilterRate
			MyRateOutput = FilterOutputRate
			MyRatePreFilter = CInt(MyRateOutput)

			MyRateForVolatility = FilterVolatilityRate
			If FilterPeakRate > 0 Then
				MyMeasurePeakValueRange = New MeasurePeakValueRange(FilterPeakRate, Me.IsFilterPeak)
				MyMeasurePeakValueRangeUsingNoPeakFilter = New MeasurePeakValueRange(FilterRate:=FilterPeakRate, IsFilterPeakEnabled:=False)
			Else
				'this is the normal default operation 
				MyMeasurePeakValueRange = New MeasurePeakValueRange(FilterRate, Me.IsFilterPeak)
				MyMeasurePeakValueRangeUsingNoPeakFilter = New MeasurePeakValueRange(FilterRate:=FilterRate, IsFilterPeakEnabled:=False)
			End If

			MyListOfPeakValueGainPrediction = New List(Of Double)

			MyLogNormalForVolatilityStatistic = New MathNet.Numerics.Distributions.LogNormal(0, 1)

			Me.IsFilterPeak = IsFilterPeakEnabled
			If FilterOutputRate < 1 Then FilterOutputRate = 1
			MyQueueOfDailyRangeSigmaExcess = New Queue(Of Tuple(Of Integer, Integer, Integer))(capacity:=MyRateForVolatility)
			MyFilterVolatilityYangZhangForStatistic = New FilterVolatilityYangZhang(MyRateForVolatility, FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=False)
			MyFilterVolatilityForPositifNegatif = New FilterVolatilityYangZhang(FilterRate, FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=False)
			MyFilterVolatilityYangZhangForStatisticLastPointTrail = New FilterVolatilityYangZhang(MyRateForVolatility, FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=True)
			MyFilterLPForPrice = New FilterLowPassPLL(FilterRate, NumberOfPredictionOutput:=0)

			MyFilterPLLForGain = New FilterLowPassPLL(FilterRate, IsPredictionEnabled:=True)
			MyFilterPLLForGainPrediction = New FilterLowPassPLL(FilterRate:=FilterRate, NumberOfPredictionOutput:=1, IsPredictionEnabled:=True)

			MyListOfProbabilityBandHigh = New List(Of Double)
			MyListOfProbabilityDailySigmaExcess = New List(Of Double)
			MyListOfProbabilityDailySigmaDoubleExcess = New List(Of Double)
			MyListOfValue = New List(Of IPriceVol)
			MyListOfPriceProbabilityMedian = New List(Of Double)
			MyListOfProbabilityBandLow = New List(Of Double)
			MyListOfPriceVolatilityGain = New List(Of Double)
			MyListOfPriceVolatilityHigh = New List(Of Double)
			MyListOfPriceVolatilityLow = New List(Of Double)
			MyListOfPriceVolatilityTimeProbability = New List(Of Double)
			MyStatisticalDistributionForVolatility = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
			MyStatisticalDistributionForVolatilityPositive = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE \ 2, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
			MyStatisticalDistributionForVolatilityNegative = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE \ 2, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))

			MyStatisticalDistributionForPrice = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
			MyFilterForVolatilityStatistic = New FilterStatistical(STATISTIC_VOLATILITY_WINDOWS_SIZE)
			MyFilterLPForProbabilityFromBandVolatility = New FilterLowPassExp(FilterOutputRate)
			MyFilterLPForStochasticFromPriceVolatilityHigh = New FilterLowPassExp(FilterOutputRate)
			MyFilterLPForStochasticFromPriceVolatilityLow = New FilterLowPassExp(FilterOutputRate)
			MyPLLErrorDetectorForPriceStochacticMedian = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGain = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)

			MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain = New FilterPLLDetectorForVolatilitySigma(
				FilterRate,
				ToCountLimit:=FILTER_PLL_DETECTOR_COUNT_LIMIT,
				ToErrorLimit:=FILTER_PLL_DETECTOR_ERROR_LIMIT)
			MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToOpenWithGain = New FilterPLLDetectorForVolatilitySigma(
				FilterRate,
				ToCountLimit:=FILTER_PLL_DETECTOR_COUNT_LIMIT,
				ToErrorLimit:=FILTER_PLL_DETECTOR_ERROR_LIMIT)


			MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain = New List(Of Double)
			MyListForVolatilityRegulatedPreviousCloseToOpenWithGain = New List(Of Double)
			MyListForVolatilityRegulatedFromOpenToCloseWithGain = New List(Of Double)
			MyListForVolatilityRegulatedNoGainFromOpenToClose = New List(Of Double)
			'predict the next sample using the last 5 samples
			'MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain = New FilterLowPassPLL(FilterRate:=5, NumberOfPredictionOutput:=1)
			MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain = New FilterLowPassExpPredict(FilterRate:=5, NumberToPredict:=1)
			'MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain = New FilterLowPassPLL(FilterRate:=Rate)
			MyListForVolatilityDetectorBalance = New List(Of Double)
			'MyFilterLPOfStochasticSlow = New FilterLowPassExpHull(FilterOutputRate)
			MyListOfStochasticFast = New ListScaled
			MyFilterLPOfStochasticSlow = New FilterLowPassExp(FilterOutputRate)
			MyListOfPriceBandHigh = New List(Of Double)
			MyListOfPriceBandLow = New List(Of Double)
			MyListOfPriceBandHighPrediction = New List(Of Double)
			MyListOfPriceBandLowPrediction = New List(Of Double)
			MyListOfProbabilityOfStockMedian = New List(Of Double)
			MyListOfPriceNextDailyHigh = New List(Of Double)
			MyListOfPriceNextDailyLow = New List(Of Double)
			MyListOfPriceNextDailyHighWithGain = New List(Of Double)
			MyListOfStochasticFastSlow = New ListScaled
			MyListOfPriceRangeVolatility = New List(Of Double)
			MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio = New List(Of Double)
			MyListOfPriceNextDailyLowWithGain = New List(Of Double)
			MyListOfPriceNextDailyLowWithGainAtSigma2 = New List(Of Double)
			MyListOfPriceNextDailyLowWithGainAtSigma3 = New List(Of Double)
			MyListOfPriceNextDailyHighWithGainAtSigma2 = New List(Of Double)
			MyListOfPriceNextDailyHighWithGainAtSigma3 = New List(Of Double)
			MyListOfPriceNextDailyLowWithGainK2 = New List(Of Double)
			MyListOfPriceNextDailyHighWithGainK2 = New List(Of Double)
			MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen = New List(Of Double)
			MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen = New List(Of Double)
			MyListOfPriceNextDailyHighWithGainOpenToClose = New List(Of Double)
			MyListOfPriceNextDailyLowWithGainOpenToClose = New List(Of Double)
			MyStatisticRangeOfExcess = New StatisticRangeExcess(MyRateForVolatility)
			'by design nothing unless initialize at the interface level
			MyStochasticPriceGain = Nothing
		End Sub

		''' <summary>
		''' create a new  stochastic object with the same basic parameters than the reference object
		''' </summary>
		''' <param name="StochasticParameterReference"></param>
		Public Sub New(ByVal StochasticParameterReference As FilterStochasticBrownian)
			Me.New(FilterRate:=StochasticParameterReference.Rate,
						 FilterOutputRate:=StochasticParameterReference.Rate(IStochastic.enuStochasticType.Slow),
						 IsFilterPeakEnabled:=StochasticParameterReference.IsFilterPeak,
						 FilterVolatilityRate:=StochasticParameterReference.Rate(IStochastic.enuStochasticType.PriceVolatilityRegulated))

			With Me
				.Tag = StochasticParameterReference.Tag
				.IsFilterRange = StochasticParameterReference.IsFilterRange
				.IsUseFeedbackRegulatedVolatility = StochasticParameterReference.IsUseFeedbackRegulatedVolatility
				.IsUseFeedbackRegulatedVolatilityFastAttackEvent = StochasticParameterReference.IsUseFeedbackRegulatedVolatilityFastAttackEvent
			End With
		End Sub

		''' <summary>
		''' create a new  stochastic object with the same basic parameters than the reference object 
		''' except for the main FilterRate parameters that can be changed
		''' </summary>
		''' <param name="FilterRate"></param>
		''' <param name="StochasticParameterReference"></param>
		Public Sub New(
			ByVal FilterRate As Integer,
			ByVal StochasticParameterReference As FilterStochasticBrownian)

			Me.New(FilterRate:=FilterRate,
						 FilterOutputRate:=StochasticParameterReference.Rate(IStochastic.enuStochasticType.Slow),
						 IsFilterPeakEnabled:=StochasticParameterReference.IsFilterPeak,
						 FilterVolatilityRate:=StochasticParameterReference.Rate(IStochastic.enuStochasticType.PriceVolatilityRegulated))

			With Me
				.Tag = StochasticParameterReference.Tag
				.IsFilterRange = StochasticParameterReference.IsFilterRange
				.IsUseFeedbackRegulatedVolatility = StochasticParameterReference.IsUseFeedbackRegulatedVolatility
				.IsUseFeedbackRegulatedVolatilityFastAttackEvent = StochasticParameterReference.IsUseFeedbackRegulatedVolatilityFastAttackEvent
			End With
		End Sub
#End Region
#Region "Private Functions"
		Private Function FilterLocal(ByRef Value As IPriceVol) As Double
			Dim I As Integer
			Dim ThisStockPriceHighValueFromOpenToClose As Double
			Dim ThisStockPriceLowValueFromOpenToClose As Double
			Dim ThisValueRemoved As IPriceVolLarge = Nothing
			Dim ThisValueHigh As Double
			Dim ThisValueLow As Double
			Dim ThisValueLowPrediction As Double
			Dim ThisValueHighPrediction As Double
			Dim ThisFilterBasedVolatilityTotal As Double
			Dim ThisFilterBasedVolatilityFromPreviousCloseToOpen As Double
			Dim ThisFilterBasedVolatilityFromOpenToClose As Double

			Dim ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen As Double
			Dim ThisFilterBasedVolatilityRatioFromOpenToClose As Double
			Dim ThisFilterBasedVolatilityFromLastPointTrailing As Double
			Dim ThisProbHigh As Double
			Dim ThisProbLow As Double
			Dim ThisProbOfStockMedian As Double
			Dim ThisProbHighFromVolatilityBand As Double
			Dim ThisProbLowFromVolatilityBand As Double
			Dim ThisProbHighHalfRate As Double
			Dim ThisProbLowHalfRate As Double
			Dim ThisStochasticResult As Double

			'Dim ThisGainPerYearFast As Double
			'Dim ThisGainPerYearDerivativeFast As Double
			'Dim ThisGainPerYearSlow As Double
			'Dim ThisGainPerYearDerivativeSlow As Double
			Dim ThisGainPerYear As Double
			Dim ThisGainPerYearDerivative As Double
			Dim ThisGainPerYearPrediction As Double
			Dim ThisGainPerYearDerivativePrediction As Double

			Dim ThisPriceVolatilityHigh As Double
			Dim ThisPriceVolatilityLow As Double
			Dim ThisRate As Integer = Me.Rate
			Dim ThisRateHalf As Integer = Me.Rate \ 2
			Dim ThisGainFromStep As Double
			Dim ThisPriceMedian As Double
			Dim ThisMu As Double
			Dim ThisSigmaSquare As Double
			Dim ThisVolatilityStatistic As IStatistical
			Dim ThisPriceNextDailyHigh As Double
			Dim ThisPriceNextDailyLow As Double
			Dim ThisPriceNextDailyHighNoGainAtSigma2 As Double
			Dim ThisPriceNextDailyHighNoGainAtSigma3 As Double
			Dim ThisPriceNextDailyLowNoGainAtSigma2 As Double
			Dim ThisPriceNextDailyLowNoGainAtSigma3 As Double

			Dim ThisPriceNextDailyHighWithGain As Double
			Dim ThisPriceNextDailyHighWithGainAtSigma2 As Double
			Dim ThisPriceNextDailyHighWithGainAtSigma3 As Double
			Dim ThisPriceNextDailyLowWithGain As Double
			Dim ThisPriceNextDailyLowWithGainAtSigma2 As Double
			Dim ThisPriceNextDailyLowWithGainAtSigma3 As Double
			Dim ThisPriceNextDailyHighWithGainK2 As Double
			Dim ThisPriceNextDailyLowWithGainK2 As Double
			Dim ThisPriceNextDailyHighPreviousCloseToOpenWithGain As Double
			Dim ThisPriceNextDailyLowPreviousCloseToOpenWithGain As Double
			Dim ThisVolatilityRegulatedFromOpenToCloseWithGain As Double

			If Me.Count = 0 Then
				MyValueLast = Value
			End If
			MyListOfValue.Add(Value)
			MyFilterLPForPrice.Filter(Value.Last)
			Dim MyPriceNextDailyHighPreviousCloseToOpenSigma3 As Double
			Dim MyPriceNextDailyLowPreviousCloseToOpenSigma3 As Double
			Dim ThisVolatilityLast As Double = 0.0
			Dim IsVolatilityJump As Boolean = False
			If MyListOfPriceRangeVolatility.Count > 0 Then
				ThisVolatilityLast = MyListOfPriceRangeVolatility.Last
				MyPriceNextDailyHighPreviousCloseToOpenSigma3 = OptionValuation.StockOption.StockPricePrediction(
						NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
						StockPrice:=Value.LastPrevious,
						Gain:=0.0,
						GainDerivative:=0.0,
						Volatility:=ThisVolatilityLast,
						Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
				MyPriceNextDailyLowPreviousCloseToOpenSigma3 = OptionValuation.StockOption.StockPricePrediction(
						NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
						StockPrice:=Value.LastPrevious,
						Gain:=0.0,
						GainDerivative:=0.0,
						Volatility:=ThisVolatilityLast,
						Probability:=GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
				If Value.Open > MyPriceNextDailyHighPreviousCloseToOpenSigma3 Then
					IsVolatilityJump = True
				ElseIf Value.Open < MyPriceNextDailyLowPreviousCloseToOpenSigma3 Then
					IsVolatilityJump = True
				End If
			End If
			'#If DEBUG Then
			'      'use to debug for a specific tag and point
			'      If Me.Tag = "FB" Then
			'        If Me.Count = 1843 Then
			'          ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.Filter(Value, IsVolatityHoldToLast:=IsVolatilityJump)
			'        Else
			'          ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.Filter(Value, IsVolatityHoldToLast:=IsVolatilityJump)
			'        End If
			'      Else
			'        ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.Filter(Value, IsVolatityHoldToLast:=IsVolatilityJump)
			'      End If
			'#Else
			ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.Filter(Value, IsVolatityHoldToLast:=IsVolatilityJump)
			'#End If
			ThisFilterBasedVolatilityFromPreviousCloseToOpen = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.PreviousCloseToOpen).Last
			ThisFilterBasedVolatilityFromOpenToClose = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToClose).Last
			ThisFilterBasedVolatilityFromLastPointTrailing = MyFilterVolatilityYangZhangForStatisticLastPointTrail.Filter(Value)
			If ThisFilterBasedVolatilityTotal > 0 Then
				ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen = ThisFilterBasedVolatilityFromPreviousCloseToOpen / ThisFilterBasedVolatilityTotal
				ThisFilterBasedVolatilityRatioFromOpenToClose = ThisFilterBasedVolatilityFromOpenToClose / ThisFilterBasedVolatilityTotal
				'Used later to study the variation of the Volatility compared to a Gaussian distribution
				MyStatisticalDistributionForVolatility.BucketFill(ThisFilterBasedVolatilityTotal)
				Select Case Value.Last - Me.Last
					Case Is = 0
						'no decision
						MyStatisticalDistributionForVolatilityPositive.BucketFill(ThisFilterBasedVolatilityTotal)
						MyStatisticalDistributionForVolatilityNegative.BucketFill(ThisFilterBasedVolatilityTotal)
					Case Is > 0
						MyStatisticalDistributionForVolatilityPositive.BucketFill(ThisFilterBasedVolatilityTotal)
					Case Is < 0
						MyStatisticalDistributionForVolatilityNegative.BucketFill(ThisFilterBasedVolatilityTotal)
				End Select
				ThisVolatilityStatistic = MyFilterForVolatilityStatistic.Filter(ThisFilterBasedVolatilityTotal)
				If ThisVolatilityStatistic.Variance > 0 Then
					'see wikipedia on lognormal distribution
					'https://en.wikipedia.org/wiki/Log-normal_distribution
					'http://www.mathworks.com/help/stats/lognstat.html
					ThisMu = Math.Log((ThisVolatilityStatistic.Mean ^ 2) / (Math.Sqrt(ThisVolatilityStatistic.Variance + ThisVolatilityStatistic.Mean ^ 2)))
					ThisSigmaSquare = Math.Log(1 + (ThisVolatilityStatistic.Variance / ThisVolatilityStatistic.Mean ^ 2))
					MyLogNormalForVolatilityStatistic = New MathNet.Numerics.Distributions.LogNormal(ThisMu, Math.Sqrt(ThisSigmaSquare))
					MyListOfPriceVolatilityTimeProbability.Add(MyLogNormalForVolatilityStatistic.CumulativeDistribution(ThisFilterBasedVolatilityTotal))
				Else
					MyListOfPriceVolatilityTimeProbability.Add(0.5)
				End If
			Else
				ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen = 1.0
				ThisFilterBasedVolatilityRatioFromOpenToClose = 0
				MyListOfPriceVolatilityTimeProbability.Add(0.5)
			End If
			MyFilterPLLForGain.Filter(Value.Last)
			MyFilterPLLForGainPrediction.Filter(Value.Last)
			ThisGainPerYear = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear.Last
			If Double.IsNaN(ThisGainPerYear) Then
				ThisGainPerYear = ThisGainPerYear
			End If
			ThisGainPerYearDerivative = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last
			ThisGainPerYearPrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYear.Last
			ThisGainPerYearDerivativePrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last

			'create the object that adjust the volatility based the the latest 
			'variation of probability of excess for a given probability threshold
			'this is the object for 1 day prediction not taking into account the gain variation of the signal
			Dim ThisVolatilityPredictionFromPreviousCloseToCloseNoGain = New StockPriceVolatilityPredictionBand(
				NumberTradingDays:=1,
				StockPrice:=Value,
				StockPriceStartValue:=Value.Last,
				Gain:=0.0,
				GainDerivative:=0.0,
				Volatility:=ThisFilterBasedVolatilityTotal,
				ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1) With {
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing,
					.IsVolatilityMaximumEnabled = True}


			If Me.Count = 500 Then
				I = I
			End If
			'same but including the gain 
			Dim ThisVolatilityPredictionFromPreviousCloseToCloseWithGain = New StockPriceVolatilityPredictionBand(
				NumberTradingDays:=1,
				StockPrice:=Value,
				StockPriceStartValue:=Value.Last,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisFilterBasedVolatilityTotal,
				ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1) With {
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing,
					.IsVolatilityMaximumEnabled = True}

			ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.Refresh(0.0)

			'this one is based on the previous close to open price volatility with no gain
			Dim ThisVolatilityPredictionPreviousCloseToOpenNoGain = New StockPriceVolatilityPredictionBand(
				NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
				StockPrice:=Value,
				StockPriceStartValue:=Value.Last,
				Gain:=0,
				GainDerivative:=0,
				Volatility:=ThisFilterBasedVolatilityFromPreviousCloseToOpen,
				ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1,
				VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen) With {
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing,
					.IsVolatilityMaximumEnabled = True}

			ThisVolatilityPredictionPreviousCloseToOpenNoGain.Refresh(0.0)

			'this one is based on the previous close to open price volatility but with gain
			Dim ThisVolatilityPredictionPreviousCloseToOpenWithGain = New StockPriceVolatilityPredictionBand(
				NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
				StockPrice:=Value,
				StockPriceStartValue:=Value.Last,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisFilterBasedVolatilityFromPreviousCloseToOpen,
				ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1,
				VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen) With {
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing,
					.IsVolatilityMaximumEnabled = True}

			ThisVolatilityPredictionPreviousCloseToOpenWithGain.Refresh(0)

			'~~~~~~~~~~~~~~Check this code before put back~~~~~~~~~~~~~~~~~~ 
			'run the PLL error volatility correction method on all the StockPriceVolatilityPredictionBand object created above
			'note: 
			'MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain use the 
			'StockPriceVolatilityPredictionBand input object and modify it with a new volatility
			'to make that clear .Update on the object should return the input object when called (to do later)

			''Previous Close to close volatility with gain
			'MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain.Update(ThisVolatilityPredictionFromPreviousCloseToCloseWithGain)
			''ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.Refresh(0.0)
			''save the result in a list
			MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain.Add(ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal)
			''filter the result for prediction purpose
			MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.Filter(ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal)

			''previous close to open
			'MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToOpenWithGain.Update(ThisVolatilityPredictionPreviousCloseToOpenWithGain)
			'ThisVolatilityPredictionPreviousCloseToOpenWithGain.Refresh(0.0)
			'save the result in a list
			MyListForVolatilityRegulatedPreviousCloseToOpenWithGain.Add(ThisVolatilityPredictionPreviousCloseToOpenWithGain.VolatilityTotal)
			'Note:this should be change for the exact calculation 

			'calculate the volatility from open to close assuming an independant energy relation between previous 
			'close to close total volatility and the previous close to open and open to close volatility:
			' (Volatility Total)^2 = (Volatility Previous Close To Open)^2 + (Volatility From Open To Close)^2
			'this approach is much faster than using the PLL error volatility correction method above
			'If ThisVolatilityPredictionPreviousCloseToOpenWithGain.VolatilityTotal > ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal Then
			'  'this is not possible assume no volatility from open to close
			'  Debugger.Break()
			'  MyListForVolatilityRegulatedFromOpenToCloseWithGain.Add(0.0)
			'Else
			'  MyListForVolatilityRegulatedFromOpenToCloseWithGain.Add(Math.Sqrt((ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal ^ 2) - (ThisVolatilityPredictionPreviousCloseToOpenWithGain.VolatilityTotal ^ 2)))
			'End If
			'again to save processing time assume tha volatility correction for the no gain is the same than with gain
			ThisVolatilityRegulatedFromOpenToCloseWithGain = ThisFilterBasedVolatilityRatioFromOpenToClose * ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal
			MyListForVolatilityRegulatedFromOpenToCloseWithGain.Add(ThisVolatilityRegulatedFromOpenToCloseWithGain)
			'If Me.Count = 900 Then
			'  I = I
			'End If
			ThisVolatilityPredictionFromPreviousCloseToCloseNoGain.Refresh(ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityDelta)

			'now we have all the different volatility and can start estimating the different corresponding price range
			Dim ThisVolatilityRegulated As Double
			Dim ThisVolatilityRegulatedForPreviousCloseToOpen As Double
			Dim ThisVolatilityForStochasticPrediction As Double
			If Me.IsUseFeedbackRegulatedVolatility Then
				'calculate the probability to reach the peak over the specified band
				'the main stochactic brownian is calculated in this section
				'many of the other calculation are test evaluation that may be removed in the future
				ThisVolatilityRegulated = ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.VolatilityTotal
				ThisVolatilityForStochasticPrediction = MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList.Last
			Else
				ThisVolatilityRegulated = ThisFilterBasedVolatilityTotal
				ThisVolatilityForStochasticPrediction = ThisFilterBasedVolatilityTotal
			End If
			ThisVolatilityRegulatedForPreviousCloseToOpen = ThisVolatilityPredictionPreviousCloseToOpenWithGain.VolatilityTotal
			'ThisVolatilityRegulatedForPreviousCloseToOpen = ThisFilterBasedVolatilityFromPreviousCloseToOpen

			ThisValueHigh = Value.High
			ThisValueLow = Value.Low
			Dim ThisMeasurementPeakForGainPeakPrediction = MyMeasurePeakValueRangeUsingNoPeakFilter.Filter(ValueLow:=ThisValueLow, ValueHigh:=ThisValueHigh)
			With MyMeasurePeakValueRange.Filter(ValueLow:=ThisValueLow, ValueHigh:=ThisValueHigh)
				ThisValueLow = .Low
				ThisValueHigh = .High
			End With
			Dim ThisMeasurementPeakForGainPeakPredictionEstimate As IPeakValueRange
			Dim ThisMeasurementPeakForGainPeakPredictionEstimateHigh As IPeakValueRange
			Dim ThisMeasurementPeakForGainPeakPredictionEstimateLow As IPeakValueRange
			With MyMeasurePeakValueRange
				ThisMeasurementPeakForGainPeakPredictionEstimate = .FilterPredictionEstimate(
					ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue,
					ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue)
				ThisMeasurementPeakForGainPeakPredictionEstimateHigh = .FilterPredictionEstimate(
					ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue,
					ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue)
				ThisMeasurementPeakForGainPeakPredictionEstimateLow = .FilterPredictionEstimate(
					ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue,
					ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue)
			End With

			ThisValueLowPrediction = ThisMeasurementPeakForGainPeakPredictionEstimate.Low
			ThisValueHighPrediction = ThisMeasurementPeakForGainPeakPredictionEstimate.High
			MyListOfPriceBandHighPrediction.Add(ThisValueHighPrediction)
			MyListOfPriceBandLowPrediction.Add(ThisValueLowPrediction)


			'calculate the price Volatility High and low over the specified rate
			'do not use the gain for the Volatility High and Low calculation
			'and use the one sigma band for the calculation range
			'the gain is maintained to zero since the current price always reflect the sum view that the investor
			'are egually positive and negative on the stock.
			'of course you view on the stock is likely different than the average but here the measurement
			'is a market average based price range.
			ThisPriceVolatilityHigh = StockOption.StockPricePrediction(
				ThisRate,
				MyFilterLPForPrice.FilterLast,
				0,
				0,
				ThisVolatilityRegulated,
				GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)
			ThisPriceVolatilityLow = StockOption.StockPricePrediction(
				ThisRate,
				MyFilterLPForPrice.FilterLast,
				0,
				0,
				ThisVolatilityRegulated,
				GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)
			ThisPriceMedian = StockOption.StockPricePredictionMedian(ThisRate, MyFilterLPForPrice.FilterLast, 0, 0, ThisVolatilityRegulated)
			If Me.Count = 0 Then
				'initialization
				'these are predictive value for the next ThisRate sample
				'one value is added to shift the samples by 1 day toward it prediction
				'but to be exact it should be shifted by the calculating rate period
				MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
				MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
				MyListOfPriceProbabilityMedian.Add(ThisPriceMedian)
				MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList.Add(
					MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList.Last)
			End If
			MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
			MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
			MyListOfPriceProbabilityMedian.Add(ThisPriceMedian)
			'given the current price calculate the probability High and low to reach the volatility 1 sigma band over the rate period
			'note that this calculation does not seem to be as predictive of price movement than using the peak high and low value for a given period
			ThisProbHighFromVolatilityBand = 1 - StockOption.StockPricePredictionInverse(ThisRate, Value.Last, 0, 0, ThisVolatilityRegulated, ThisPriceVolatilityHigh)
			ThisProbLowFromVolatilityBand = StockOption.StockPricePredictionInverse(ThisRate, Value.Last, 0, 0, ThisVolatilityRegulated, ThisPriceVolatilityLow)
			'save the probability result
			'note: does not appear to be very useful but need to be checked again
			MyFilterLPForProbabilityFromBandVolatility.Filter(ThisProbHighFromVolatilityBand / (ThisProbHighFromVolatilityBand + ThisProbLowFromVolatilityBand))
			If Me.IsUseFeedbackRegulatedVolatility Then
				With ThisVolatilityPredictionFromPreviousCloseToCloseNoGain
					ThisPriceNextDailyHigh = .StockPriceHighValue
					ThisPriceNextDailyLow = .StockPriceLowValue
					ThisPriceNextDailyHighNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
					ThisPriceNextDailyHighNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
					ThisPriceNextDailyLowNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
					ThisPriceNextDailyLowNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
				End With
				With ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
					ThisPriceNextDailyHighWithGain = .StockPriceHighValue
					ThisPriceNextDailyLowWithGain = .StockPriceLowValue
					ThisPriceNextDailyHighWithGainK2 = .StockPriceHighPrediction(Index:=2.0)
					ThisPriceNextDailyLowWithGainK2 = .StockPriceLowPrediction(Index:=2.0)
					ThisPriceNextDailyHighWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
					ThisPriceNextDailyHighWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
					ThisPriceNextDailyLowWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
					ThisPriceNextDailyLowWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
				End With
			Else
				With ThisVolatilityPredictionFromPreviousCloseToCloseNoGain
					ThisPriceNextDailyHigh = .StockPriceHighValueStandard
					ThisPriceNextDailyLow = .StockPriceLowValueStandard
					ThisPriceNextDailyHighNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
					ThisPriceNextDailyHighNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
					ThisPriceNextDailyLowNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
					ThisPriceNextDailyLowNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
				End With
				With ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
					ThisPriceNextDailyHighWithGain = .StockPriceHighValueStandard
					ThisPriceNextDailyLowWithGain = .StockPriceLowValueStandard
					ThisPriceNextDailyHighWithGainK2 = .StockPriceHighPrediction(Index:=2.0)
					ThisPriceNextDailyLowWithGainK2 = .StockPriceLowPrediction(Index:=2.0)
					ThisPriceNextDailyHighWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
					ThisPriceNextDailyHighWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
					ThisPriceNextDailyLowWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
					ThisPriceNextDailyLowWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
				End With
			End If

			'If Me.Count = 1000 Then
			'just a test
			'these two call give the same the same
			'ThisPriceNextDailyHighPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
			'  1.0,
			'  Value.Last,
			'  ThisGainPerYear,
			'  ThisGainPerYearDerivative,
			'  ThisVolatilityPredictionFromPreviousCloseToCloseNoGain.VolatilityTotal,
			'  GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)

			'ThisPriceNextDailyHighPreviousCloseToOpenWithGain = ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPricePrediction(1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
			'End If
			ThisPriceNextDailyHighPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
				TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
				Value.Last,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisVolatilityRegulatedForPreviousCloseToOpen,
				GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)
			ThisPriceNextDailyLowPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
				TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
				Value.Last,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisVolatilityRegulatedForPreviousCloseToOpen,
				GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)

			ThisStockPriceHighValueFromOpenToClose = StockOption.StockPricePrediction(
				NumberTradingDays:=TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY,
				StockPrice:=Value.Open,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisFilterBasedVolatilityFromOpenToClose,
				Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)

			ThisStockPriceLowValueFromOpenToClose = StockOption.StockPricePrediction(
				NumberTradingDays:=TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY,
				StockPrice:=Value.Open,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisFilterBasedVolatilityFromOpenToClose,
				Probability:=GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)



			'ThisVolatilityRegulatedFromOpenToCloseWithGain

			Dim ThisRange As Double = (ThisPriceNextDailyHighWithGain - ThisPriceNextDailyLowWithGain) / (ThisPriceNextDailyHighWithGain + ThisPriceNextDailyLowWithGain)
			If ThisRange < 0.002 Then
				ThisRange = ThisRange
			End If

			'~~~~~~~~~~~~~~~~~~~~~~~~
			'Calculate the price median that is needed to bring the Stochastic to 50%
			'this value is found using a second order PLL filter tracking the price that bring the Stochastic to 50%
			'for this we provide the filter with a special phase detector that contain the equation that return zero at the 50% stochactic
			'this method converge iterativly rapidly toward the price value for the 50% stochastic
			'update the phase detector with the latest parameters
			'we need two one with the gain and the other one without
			'run the filter with the special phase detector
			'ThisPricePeakMedian is the starting search point since it is not generally to far from the searched solution
			'the filter contain the price value that bring the 50% stochastic
			'Note: this is a very important price threshold value that can be used to help start or terminate a trade
			'the value that include the gain seem to be the most usuful to enter or leave a trade.
			MyPLLErrorDetectorForPriceStochacticMedian.Update(
				ThisVolatilityRegulated,
				0.0,
				0.0,
				ThisValueHigh,
				ThisValueLow)
			MyPLLErrorDetectorForPriceStochacticMedianWithGain.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisValueHigh,
				ThisValueLow)
			MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisValueHighPrediction,
				ThisValueLowPrediction)

			MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisMeasurementPeakForGainPeakPrediction.High,
				ThisMeasurementPeakForGainPeakPrediction.Low)

			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisMeasurementPeakForGainPeakPredictionEstimate.High,
				ThisMeasurementPeakForGainPeakPredictionEstimate.Low)

			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisMeasurementPeakForGainPeakPredictionEstimateHigh.High,
				ThisMeasurementPeakForGainPeakPredictionEstimateHigh.Low)

			MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow.Update(
				ThisVolatilityRegulated,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisMeasurementPeakForGainPeakPredictionEstimateLow.High,
				ThisMeasurementPeakForGainPeakPredictionEstimateLow.Low)


			Dim ThisFilterPredictionGainYearly As Double = MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(
																																																	MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter.ToList.Last,
																																																	MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter.ToList.Last)

			ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
			MyListOfPeakValueGainPrediction.Add(ThisFilterPredictionGainYearly)

			'calculate the Peak Gain prediction
			If Me.Count = 0 Then
				'initialization
				'these are predictive value for the next sample 
				'one value is added initialy to shift the samples by 1 day toward it next day prediction
				'MyListOfPeakValueGainPrediction.Add(ThisFilterPredictionGainYearly)
				With MyPLLErrorDetectorForPriceStochacticMedian
					.ToList.Add(.ToList.Last)
					.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
					.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
				End With
				With MyPLLErrorDetectorForPriceStochacticMedianWithGain
					.ToList.Add(.ToList.Last)
					.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
					.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
				End With
				With MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction
					.ToList.Add(.ToList.Last)
					.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
					.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
				End With
				With MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh
					.ToList.Add(.ToList.Last)
					.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
					.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
				End With
				With MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow
					.ToList.Add(.ToList.Last)
					.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
					.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
				End With
			End If
			If Me.Count = 1000 Then
				ThisProbHigh = ThisProbHigh
			End If
			If Me.Count > 0 Then
				MyStatisticRangeOfExcess.Run(Value, MyListOfPriceNextDailyLowWithGain.Last, MyListOfPriceNextDailyHighWithGain.Last)
			Else
				'initialization
				'these are predictive value for the next sample 
				'one extra value is added just at the beginning to shift the samples by 1 day toward it prediction
				MyListOfPriceNextDailyHigh.Add(ThisPriceNextDailyHigh)
				MyListOfPriceNextDailyLow.Add(ThisPriceNextDailyLow)
				MyListOfPriceNextDailyHighWithGain.Add(ThisPriceNextDailyHighWithGain)
				MyListOfPriceNextDailyLowWithGain.Add(ThisPriceNextDailyLowWithGain)
				MyListOfPriceNextDailyLowWithGainAtSigma2.Add(ThisPriceNextDailyLowWithGainAtSigma2)
				MyListOfPriceNextDailyLowWithGainAtSigma3.Add(ThisPriceNextDailyLowWithGainAtSigma3)
				MyListOfPriceNextDailyHighWithGainAtSigma2.Add(ThisPriceNextDailyHighWithGainAtSigma2)
				MyListOfPriceNextDailyHighWithGainAtSigma3.Add(ThisPriceNextDailyHighWithGainAtSigma3)
				MyListOfPriceNextDailyHighWithGainK2.Add(ThisPriceNextDailyHighWithGainK2)
				MyListOfPriceNextDailyLowWithGainK2.Add(ThisPriceNextDailyLowWithGainK2)
				MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyHighPreviousCloseToOpenWithGain)
				MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyLowPreviousCloseToOpenWithGain)
				'and initilize the queue use to measure the excess rate
				'always no excess at the beginning
				'MyQueueOfDailyRangeSigmaExcess.Enqueue(New Tuple(Of Integer, Integer, Integer)(0, 0, 0))
				'MySumOfDailyRangeSigmaExcess = 0
				'ThisProbabilityDailySigmaExcess = 0.0
				MyStatisticRangeOfExcess.Run(Value, Value.Last, Value.Last)
			End If
			MyListOfProbabilityDailySigmaExcess.Add(MyStatisticRangeOfExcess.BandLimitExcess)
			MyListOfProbabilityDailySigmaDoubleExcess.Add(MyStatisticRangeOfExcess.BandLimitDoubleExcess)
			MyListForVolatilityDetectorBalance.Add(MyStatisticRangeOfExcess.BandLimitHighLowBalance)

			'leave this here after the measurement of the probability of excess
			MyListOfPriceNextDailyHigh.Add(ThisPriceNextDailyHigh)
			MyListOfPriceNextDailyLow.Add(ThisPriceNextDailyLow)
			MyListOfPriceNextDailyHighWithGain.Add(ThisPriceNextDailyHighWithGain)
			MyListOfPriceNextDailyLowWithGain.Add(ThisPriceNextDailyLowWithGain)
			MyListOfPriceNextDailyLowWithGainAtSigma2.Add(ThisPriceNextDailyLowWithGainAtSigma2)
			MyListOfPriceNextDailyLowWithGainAtSigma3.Add(ThisPriceNextDailyLowWithGainAtSigma3)
			MyListOfPriceNextDailyHighWithGainAtSigma2.Add(ThisPriceNextDailyHighWithGainAtSigma2)
			MyListOfPriceNextDailyHighWithGainAtSigma3.Add(ThisPriceNextDailyHighWithGainAtSigma3)
			MyListOfPriceNextDailyHighWithGainK2.Add(ThisPriceNextDailyHighWithGainK2)
			MyListOfPriceNextDailyLowWithGainK2.Add(ThisPriceNextDailyLowWithGainK2)
			MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyHighPreviousCloseToOpenWithGain)
			MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyLowPreviousCloseToOpenWithGain)
			MyListOfPriceNextDailyLowWithGainOpenToClose.Add(ThisStockPriceLowValueFromOpenToClose)
			MyListOfPriceNextDailyHighWithGainOpenToClose.Add(ThisStockPriceHighValueFromOpenToClose)

			'If Me.Count = 1000 Then
			'  ThisProbHigh = ThisProbHigh
			'End If
			ThisProbHigh = 1 - StockOption.StockPricePredictionInverse(
				ThisRate,
				Value.Last,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisVolatilityRegulated,
				ThisValueHigh)

			ThisProbLow = StockOption.StockPricePredictionInverse(
				ThisRate,
				Value.Last,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				ThisVolatilityRegulated,
				ThisValueLow)

			'ThisProbHigh = 1 - StockOption.StockPricePredictionInverse(
			'  10,
			'  Value.Last,
			'  ThisGainPerYear,
			'  ThisGainPerYearDerivative,
			'  ThisVolatilityRegulated,
			'  ThisValueHigh)



			'ThisProbLow = StockOption.StockPricePredictionInverse(
			'  10,
			'  Value.Last,
			'  ThisGainPerYear,
			'  ThisGainPerYearDerivative,
			'  ThisVolatilityRegulated,
			'  ThisValueLow)

			ThisProbHighHalfRate = 1 - StockOption.StockPricePredictionInverse(ThisRateHalf, Value.Last, ThisGainPerYear, ThisGainPerYearDerivative, ThisVolatilityRegulated, ThisValueHigh)
			ThisProbLowHalfRate = StockOption.StockPricePredictionInverse(ThisRateHalf, Value.Last, ThisGainPerYear, ThisGainPerYearDerivative, ThisVolatilityRegulated, ThisValueLow)
			ThisProbHighHalfRate = ThisProbHigh
			ThisProbLowHalfRate = ThisProbLow

			ThisStochasticResult = ThisProbHigh / (ThisProbHigh + ThisProbLow)

			'calculate the probability to reach or exceed the Stock Price median 
			'over 20% of the FilterRate period. 
			'for PriceStochacticMedianWithGain
			ThisProbOfStockMedian = 1 - StockOption.StockPricePredictionInverse(
				NumberTradingDays:=ThisRate / 5,
				StockPriceStart:=Value.Last,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisVolatilityRegulated,
				StockPriceEnd:=MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToList.Last)

			'Dim ThisProbOfStockMedian1 = 1 - StockOption.StockPricePredictionInverse(
			'	NumberTradingDays:=5,
			'	StockPriceStart:=Value.Last,
			'	Gain:=ThisGainPerYear,
			'	GainDerivative:=ThisGainPerYearDerivative,
			'	Volatility:=ThisVolatilityRegulated,
			'	StockPriceEnd:=MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToList.Last)



			'main stochactic result
			MyListOfProbabilityOfStockMedian.Add(ThisProbOfStockMedian)
			MyFilterVolatilityForPositifNegatif.Filter(Value, IsVolatityHoldToLast:=False)

			'ThisStochasticResultHalfRate = ThisProbHighHalfRate / (ThisProbHighHalfRate + ThisProbLowHalfRate)
			'average the result for the both rate 
			'ThisStochasticResult = (ThisStochasticResult + ThisStochasticResultHalfRate) / 2

			'MyListOfProbabilityBandHigh.Add((ThisProbHigh + ThisProbHighHalfRate) / 2)
			'MyListOfProbabilityBandLow.Add((ThisProbLow + ThisProbLowHalfRate) / 2)
			MyListOfProbabilityBandHigh.Add(ThisProbHigh)
			MyListOfProbabilityBandLow.Add(ThisProbLow)
			ThisGainFromStep = WaveForm.SignalLimit(((NUMBER_WORKDAY_PER_YEAR * Measure.Measure.GainLog(ThisPriceVolatilityHigh + ThisPriceVolatilityLow, ThisValueHigh + ThisValueLow) / ThisRate) / 2 + 0.5), 0.5, 0.5)
			'ThisGainFromStep = WaveForm.SignalLimit(((NUMBER_WORKDAY_PER_YEAR * Measure.Measure.GainLog(ThisPriceVolatilityHigh, ThisValueHigh) / ThisRate) / 2 + 0.5), 0.5, 0.5)
			MyListOfPriceVolatilityGain.Add(ThisGainFromStep)
			MyValueLast = Value
			'note for now always send the range volatility here
			'Return MyFilterStochastic.ListDataUpdate(ThisValueHigh, ThisValueLow, ThisFilterBasedVolatilityTotal, ThisStochasticResult)
			MyListOfPriceBandHigh.Add(ThisValueHigh)
			MyListOfPriceBandLow.Add(ThisValueLow)

			MyListOfPriceRangeVolatility.Add(ThisFilterBasedVolatilityTotal)
			MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio.Add(ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen)
			MyFilterLPOfStochasticSlow.Filter(ThisStochasticResult)
			MyListOfStochasticFast.Add(ThisStochasticResult)
			MyStocFastSlowLast = ThisStochasticResult - MyFilterLPOfStochasticSlow.FilterLast
			MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
			MyStocRangeVolatility = ThisFilterBasedVolatilityTotal
			MyStocFastLast = ThisStochasticResult
			Return MyFilterLPOfStochasticSlow.FilterLast
		End Function

		Private Sub FilterLocalPart1(ByVal ReportPrices As YahooAccessData.RecordPrices)
			Dim ThisValueRemoved As IPriceVolLarge = Nothing
			Dim ThisFilterBasedVolatilityTotal As Double
			Dim ThisFilterBasedVolatilityFromPreviousCloseToOpen As Double
			Dim ThisFilterBasedVolatilityFromOpenToClose As Double
			Dim ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen As Double
			Dim ThisFilterBasedVolatilityFromLastPointTrailing As Double
			Dim ThisGainPerYear As Double
			Dim ThisGainPerYearDerivative As Double
			Dim ThisGainPerYearPrediction As Double
			Dim ThisGainPerYearDerivativePrediction As Double
			Dim ThisMu As Double
			Dim ThisSigmaSquare As Double
			Dim ThisVolatilityStatistic As IStatistical
			Dim I, J As Integer
			Dim Value As IPriceVol

			ReDim MyStockPriceVolatilityPredictionBand(0 To ReportPrices.NumberPoint - 1)
			ReDim MyStockPriceVolatilityPredictionBandWithGain(0 To ReportPrices.NumberPoint - 1)
			ReDim MyStockPriceVolatilityPredictionBandWithGainCloseToOpen(0 To ReportPrices.NumberPoint - 1)

			If ReportPrices.NumberPoint = 0 Then Exit Sub
			MyValueLast = ReportPrices.GetPriceVolInterface(0)
			For I = 0 To ReportPrices.NumberPoint - 1
				Value = ReportPrices.GetPriceVolInterface(I)
				MyListOfValue.Add(Value)
				MyFilterLPForPrice.Filter(Value.Last)
				ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.Filter(Value)
				ThisFilterBasedVolatilityFromPreviousCloseToOpen = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.PreviousCloseToOpen).Last
				ThisFilterBasedVolatilityFromOpenToClose = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToClose).Last
				ThisFilterBasedVolatilityFromLastPointTrailing = MyFilterVolatilityYangZhangForStatisticLastPointTrail.Filter(Value)
				If ThisFilterBasedVolatilityTotal > 0 Then
					ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen = ThisFilterBasedVolatilityFromPreviousCloseToOpen / ThisFilterBasedVolatilityTotal
					'Used later to study the variation of the Volatility compared to a Gaussian distribution
					MyStatisticalDistributionForVolatility.BucketFill(ThisFilterBasedVolatilityTotal)
					Select Case Value.Last - Me.Last
						Case Is = 0
							'no decision
							MyStatisticalDistributionForVolatilityPositive.BucketFill(ThisFilterBasedVolatilityTotal)
							MyStatisticalDistributionForVolatilityNegative.BucketFill(ThisFilterBasedVolatilityTotal)
						Case Is > 0
							MyStatisticalDistributionForVolatilityPositive.BucketFill(ThisFilterBasedVolatilityTotal)
						Case Is < 0
							MyStatisticalDistributionForVolatilityNegative.BucketFill(ThisFilterBasedVolatilityTotal)
					End Select
					ThisVolatilityStatistic = MyFilterForVolatilityStatistic.Filter(ThisFilterBasedVolatilityTotal)
					If ThisVolatilityStatistic.Variance > 0 Then
						'see wikipedia on lognormal distribution
						'https://en.wikipedia.org/wiki/Log-normal_distribution
						'http://www.mathworks.com/help/stats/lognstat.html
						ThisMu = Math.Log((ThisVolatilityStatistic.Mean ^ 2) / (Math.Sqrt(ThisVolatilityStatistic.Variance + ThisVolatilityStatistic.Mean ^ 2)))
						ThisSigmaSquare = Math.Log(1 + (ThisVolatilityStatistic.Variance / ThisVolatilityStatistic.Mean ^ 2))
						MyLogNormalForVolatilityStatistic = New MathNet.Numerics.Distributions.LogNormal(ThisMu, Math.Sqrt(ThisSigmaSquare))
						MyListOfPriceVolatilityTimeProbability.Add(MyLogNormalForVolatilityStatistic.CumulativeDistribution(ThisFilterBasedVolatilityTotal))
					Else
						MyListOfPriceVolatilityTimeProbability.Add(0.5)
					End If
				Else
					ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen = 1.0
					MyListOfPriceVolatilityTimeProbability.Add(0.5)
				End If
				MyFilterPLLForGain.Filter(Value.Last)
				MyFilterPLLForGainPrediction.Filter(Value.Last)
				ThisGainPerYear = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear.Last
				'If Double.IsNaN(ThisGainPerYear) Then
				'  Debugger.Break()
				'End If
				ThisGainPerYearDerivative = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last
				ThisGainPerYearPrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYear.Last
				ThisGainPerYearDerivativePrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last

				'correction for the volatility prediction on next day
				Dim ThisVolatilityPredictionFromPreviousCloseToCloseNoGain = New StockPriceVolatilityPredictionBand(
																																														NumberTradingDays:=1,
																																														StockPrice:=Value,
																																														StockPriceStartValue:=Value.Last,
																																														Gain:=0.0,
																																														GainDerivative:=0.0,
																																														Volatility:=ThisFilterBasedVolatilityTotal,
																																														ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1)
				Dim ThisVolatilityPredictionFromPreviousCloseToCloseWithGain = New StockPriceVolatilityPredictionBand(
																																														 NumberTradingDays:=1,
																																														 StockPrice:=Value,
																																														 StockPriceStartValue:=Value.Last,
																																														 Gain:=ThisGainPerYear,
																																														 GainDerivative:=ThisGainPerYearDerivative,
																																														 Volatility:=ThisFilterBasedVolatilityTotal,
																																														 ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1)

				Dim ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen = New StockPriceVolatilityPredictionBand(
																																													 NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
																																													 StockPrice:=Value,
																																													 StockPriceStartValue:=Value.Last,
																																													 Gain:=ThisGainPerYear,
																																													 GainDerivative:=ThisGainPerYearDerivative,
																																													 Volatility:=ThisFilterBasedVolatilityFromPreviousCloseToOpen,
																																													 ProbabilityOfInterval:=GAUSSIAN_PROBABILITY_SIGMA1,
																																													 VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen)

				'update the price high and low that is use for price volatility variation excess measurement
				J = J + 1
				If J < ReportPrices.NumberPoint Then
					With ReportPrices
						'If ReportPrices.Symbol = "NFLX" Then
						'  If .DateLastTrade = DateSerial(2014, 11, 19) Then
						'    J = J
						'  End If
						'End If
						ThisVolatilityPredictionFromPreviousCloseToCloseNoGain.Refresh(.GetPriceVolInterface(J))
						ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.Refresh(.GetPriceVolInterface(J))
						ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen.Refresh(.GetPriceVolInterface(J))
					End With
				End If
				MyListOfPriceRangeVolatility.Add(ThisFilterBasedVolatilityTotal)
				MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio.Add(ThisFilterBasedVolatilityRatioFromPreviousCloseToOpen)
				MyStockPriceVolatilityPredictionBand(I) = ThisVolatilityPredictionFromPreviousCloseToCloseNoGain
				MyStockPriceVolatilityPredictionBandWithGain(I) = ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
				MyStockPriceVolatilityPredictionBandWithGainCloseToOpen(I) = ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen
				MyValueLast = Value
			Next
		End Sub

		Private Sub FilterLocalPart2(ByVal ReportPrices As YahooAccessData.RecordPrices)
			Dim ThisFilterBasedVolatilityFromPreviousCloseToOpen As Double
			Dim ThisFilterBasedVolatilityFromOpenToClose As Double
			Dim ThisStockPriceHighValueFromOpenToClose As Double
			Dim ThisStockPriceLowValueFromOpenToClose As Double
			Dim ThisValueRemoved As IPriceVolLarge = Nothing
			Dim ThisValueHigh As Double
			Dim ThisValueLow As Double
			Dim ThisValueLowPrediction As Double
			Dim ThisValueHighPrediction As Double
			Dim ThisProbOfStockMedian As Double
			Dim ThisFilterBasedVolatilityTotal As Double
			Dim ThisFilterBasedVolatilityFromLastPointTrailing As Double
			Dim ThisProbHigh As Double
			Dim ThisProbLow As Double
			Dim ThisProbHighFromVolatilityBand As Double
			Dim ThisProbLowFromVolatilityBand As Double
			Dim ThisProbHighHalfRate As Double
			Dim ThisProbLowHalfRate As Double
			Dim ThisStochasticResult As Double
			Dim ThisStochasticResultHalfRate As Double
			Dim ThisGainPerYear As Double
			Dim ThisGainPerYearDerivative As Double
			Dim ThisGainPerYearPrediction As Double
			Dim ThisGainPerYearDerivativePrediction As Double

			Dim ThisPriceVolatilityHigh As Double
			Dim ThisPriceVolatilityLow As Double
			Dim ThisRate As Integer = Me.Rate
			Dim ThisRateHalf As Integer = Me.Rate \ 2
			Dim ThisGainFromStep As Double
			Dim ThisPriceMedian As Double
			Dim ThisPriceNextDailyHigh As Double
			Dim ThisPriceNextDailyLow As Double
			Dim ThisPriceNextDailyHighNoGainAtSigma2 As Double
			Dim ThisPriceNextDailyHighNoGainAtSigma3 As Double
			Dim ThisPriceNextDailyLowNoGainAtSigma2 As Double
			Dim ThisPriceNextDailyLowNoGainAtSigma3 As Double
			Dim ThisPriceNextDailyHighWithGain As Double
			Dim ThisPriceNextDailyHighWithGainAtSigma2 As Double
			Dim ThisPriceNextDailyHighWithGainAtSigma3 As Double
			Dim ThisPriceNextDailyLowWithGain As Double
			Dim ThisPriceNextDailyLowWithGainAtSigma2 As Double
			Dim ThisPriceNextDailyLowWithGainAtSigma3 As Double


			Dim ThisVolatilityRegulated As Double
			Dim ThisVolatilityForStochasticPrediction As Double
			Dim I As Integer
			Dim Value As IPriceVol
			Dim ThisVolatilityPredictionFromPreviousCloseToCloseNoGain As StockPriceVolatilityPredictionBand
			Dim ThisVolatilityPredictionFromPreviousCloseToCloseWithGain As StockPriceVolatilityPredictionBand
			Dim ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen As StockPriceVolatilityPredictionBand
			Dim ThisPriceNextDailyHighWithGainK2 As Double
			Dim ThisPriceNextDailyLowWithGainK2 As Double
			Dim ThisPriceNextDailyHighPreviousCloseToOpenWithGain As Double
			Dim ThisPriceNextDailyLowPreviousCloseToOpenWithGain As Double

			For I = 0 To ReportPrices.NumberPoint - 1
				Value = ReportPrices.GetPriceVolInterface(I)
				ThisVolatilityPredictionFromPreviousCloseToCloseNoGain = MyStockPriceVolatilityPredictionBand(I)
				ThisVolatilityPredictionFromPreviousCloseToCloseWithGain = MyStockPriceVolatilityPredictionBandWithGain(I)
				ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen = MyStockPriceVolatilityPredictionBandWithGainCloseToOpen(I)

				'update the object with the latest volatility obtained on the threaded calculation
				ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.Refresh(MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain(I) - ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.Volatility)
				ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen.Refresh(MyListForVolatilityRegulatedPreviousCloseToOpenWithGain(I) - ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen.Volatility)
				ThisFilterBasedVolatilityTotal = MyFilterVolatilityYangZhangForStatistic.ToList.Item(I)
				ThisFilterBasedVolatilityFromLastPointTrailing = MyFilterVolatilityYangZhangForStatisticLastPointTrail.ToList.Item(I)

				ThisGainPerYear = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear(I)
				ThisGainPerYearDerivative = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative(I)
				ThisGainPerYearPrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYear(I)
				ThisGainPerYearDerivativePrediction = MyFilterPLLForGainPrediction.AsIFilterPrediction.ToListOfGainPerYearDerivative(I)

				With ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing
					.IsVolatilityMaximumEnabled = True
				End With
				With ThisStockPriceVolatilityPredictionBandWithGainCloseToOpen
					.VolatilityMaximum = ThisFilterBasedVolatilityFromLastPointTrailing
					.IsVolatilityMaximumEnabled = True
				End With
				ThisVolatilityPredictionFromPreviousCloseToCloseNoGain.Refresh(MyStockPriceVolatilityPredictionBandWithGain(I).VolatilityDelta)


				If Me.IsUseFeedbackRegulatedVolatility Then
					'calculate the probability to reach the peak ove the specified band
					'the main stochactic brownian is calculated in this section
					'many of the pther calculation are test evaluation that may be removed in the future
					ThisVolatilityRegulated = MyStockPriceVolatilityPredictionBandWithGain(I).VolatilityTotal
					ThisVolatilityForStochasticPrediction = MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList(I)
				Else
					ThisVolatilityRegulated = ThisFilterBasedVolatilityTotal
					ThisVolatilityForStochasticPrediction = ThisVolatilityRegulated
				End If

				ThisValueHigh = Value.High
				ThisValueLow = Value.Low

				Dim ThisMeasurementPeakForGainPeakPrediction = MyMeasurePeakValueRangeUsingNoPeakFilter.Filter(ValueLow:=ThisValueLow, ValueHigh:=ThisValueHigh)
				With MyMeasurePeakValueRange.Filter(ValueLow:=ThisValueLow, ValueHigh:=ThisValueHigh)
					ThisValueLow = .Low
					ThisValueHigh = .High
				End With
				Dim ThisMeasurementPeakForGainPeakPredictionEstimate As IPeakValueRange
				Dim ThisMeasurementPeakForGainPeakPredictionEstimateHigh As IPeakValueRange
				Dim ThisMeasurementPeakForGainPeakPredictionEstimateLow As IPeakValueRange
				With MyMeasurePeakValueRange
					ThisMeasurementPeakForGainPeakPredictionEstimate = .FilterPredictionEstimate(
						ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue,
						ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue)
					ThisMeasurementPeakForGainPeakPredictionEstimateHigh = .FilterPredictionEstimate(
						ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue,
						ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceHighValue)
					ThisMeasurementPeakForGainPeakPredictionEstimateLow = .FilterPredictionEstimate(
						ValueLow:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue,
						ValueHigh:=ThisVolatilityPredictionFromPreviousCloseToCloseWithGain.StockPriceLowValue)
				End With

				ThisValueLowPrediction = ThisMeasurementPeakForGainPeakPredictionEstimate.Low
				ThisValueHighPrediction = ThisMeasurementPeakForGainPeakPredictionEstimate.High
				MyListOfPriceBandHighPrediction.Add(ThisValueHighPrediction)
				MyListOfPriceBandLowPrediction.Add(ThisValueLowPrediction)



				'calculate the price Volatility High and low over the specified rate
				'do not use the gain for the Volatility High and Low calculation
				'and use the one sigma band for the calculation range
				'the gain is maintained to zero since the current price always reflect the sum view that the investor
				'are egually positive and negative on the stock.
				'of course you view on the stock is likely different than the average but here the measurement
				'is a market average based price range.
				Dim ThisFilterLast As Double = MyFilterLPForPrice.ToList(I)
				ThisPriceVolatilityHigh = StockOption.StockPricePrediction(ThisRate, ThisFilterLast, 0, 0, ThisVolatilityRegulated, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)
				ThisPriceVolatilityLow = StockOption.StockPricePrediction(ThisRate, ThisFilterLast, 0, 0, ThisVolatilityRegulated, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)
				ThisPriceMedian = StockOption.StockPricePredictionMedian(ThisRate, ThisFilterLast, 0, 0, ThisVolatilityRegulated)
				If Me.Count = 0 Then
					'initialization
					'these are predictive value for the next ThisRate sample
					'one value is added to shift the samples by 1 day toward it prediction
					'but to be exact it should be shifted by the calculating rate period
					MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
					MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
					MyListOfPriceProbabilityMedian.Add(ThisPriceMedian)
				End If
				MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
				MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
				MyListOfPriceProbabilityMedian.Add(ThisPriceMedian)
				'given the current price calculate the probability High and low to reach the volatility 1 sigma band over the rate period
				'note that this calculation does not seem to be as predictive of price movement than using the peak high and low value for a given period
				ThisProbHighFromVolatilityBand = 1 - StockOption.StockPricePredictionInverse(ThisRate, Value.Last, 0, 0, ThisVolatilityRegulated, ThisPriceVolatilityHigh)
				ThisProbLowFromVolatilityBand = StockOption.StockPricePredictionInverse(ThisRate, Value.Last, 0, 0, ThisVolatilityRegulated, ThisPriceVolatilityLow)
				'save the probability result
				'note: does not appear to be very useful but need to be checked again
				MyFilterLPForProbabilityFromBandVolatility.Filter(ThisProbHighFromVolatilityBand / (ThisProbHighFromVolatilityBand + ThisProbLowFromVolatilityBand))

				If Me.IsUseFeedbackRegulatedVolatility Then
					With ThisVolatilityPredictionFromPreviousCloseToCloseNoGain
						ThisPriceNextDailyHigh = .StockPriceHighValue
						ThisPriceNextDailyLow = .StockPriceLowValue
						ThisPriceNextDailyHighNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
						ThisPriceNextDailyHighNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
						ThisPriceNextDailyLowNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
						ThisPriceNextDailyLowNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
					End With
					With ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
						ThisPriceNextDailyHighWithGain = .StockPriceHighValue
						ThisPriceNextDailyLowWithGain = .StockPriceLowValue
						ThisPriceNextDailyHighWithGainK2 = .StockPriceHighPrediction(Index:=2.0)
						ThisPriceNextDailyLowWithGainK2 = .StockPriceLowPrediction(Index:=2.0)
						ThisPriceNextDailyHighWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
						ThisPriceNextDailyHighWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
						ThisPriceNextDailyLowWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
						ThisPriceNextDailyLowWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
					End With
				Else
					With ThisVolatilityPredictionFromPreviousCloseToCloseNoGain
						ThisPriceNextDailyHigh = .StockPriceHighValueStandard
						ThisPriceNextDailyLow = .StockPriceLowValueStandard
						ThisPriceNextDailyHighNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
						ThisPriceNextDailyHighNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
						ThisPriceNextDailyLowNoGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
						ThisPriceNextDailyLowNoGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
					End With
					With ThisVolatilityPredictionFromPreviousCloseToCloseWithGain
						ThisPriceNextDailyHighWithGain = .StockPriceHighValueStandard
						ThisPriceNextDailyLowWithGain = .StockPriceLowValueStandard
						ThisPriceNextDailyHighWithGainK2 = .StockPriceHighPrediction(Index:=2.0)
						ThisPriceNextDailyLowWithGainK2 = .StockPriceLowPrediction(Index:=2.0)
						ThisPriceNextDailyHighWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
						ThisPriceNextDailyHighWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3)
						ThisPriceNextDailyLowWithGainAtSigma2 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2)
						ThisPriceNextDailyLowWithGainAtSigma3 = .StockPricePrediction(Index:=1.0, GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3)
					End With
				End If
				'ThisFilterBasedVolatilityFromPreviousCloseToOpen = MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio(I) * MyStockPriceVolatilityPredictionBandWithGain(I).VolatilityTotal
				ThisFilterBasedVolatilityFromPreviousCloseToOpen = MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio(I) * MyStockPriceVolatilityPredictionBandWithGain(I).Volatility

				ThisPriceNextDailyHighPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
					TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
					Value.Last,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisFilterBasedVolatilityFromPreviousCloseToOpen,
					GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)

				ThisPriceNextDailyLowPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
					TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
					Value.Last,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisFilterBasedVolatilityFromPreviousCloseToOpen,
					GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)

				ThisStockPriceHighValueFromOpenToClose = StockOption.StockPricePrediction(
					NumberTradingDays:=TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY,
					StockPrice:=Value.Open,
					Gain:=ThisGainPerYear,
					GainDerivative:=ThisGainPerYearDerivative,
					Volatility:=ThisFilterBasedVolatilityFromOpenToClose,
					Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)

				ThisStockPriceLowValueFromOpenToClose = StockOption.StockPricePrediction(
				NumberTradingDays:=TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY,
				StockPrice:=Value.Open,
				Gain:=ThisGainPerYear,
				GainDerivative:=ThisGainPerYearDerivative,
				Volatility:=ThisFilterBasedVolatilityFromOpenToClose,
				Probability:=GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)


				'~~~~~~~~~~~~~~~~~~~~~~~~
				'Calculate the price median that is needed to bring the Stochastic to 50%
				'this value is found using a second order PLL filter tracking the price that bring the Stochastic to 50%
				'for this we provide the filter with a special phase detector that contain the equation that return zero at the 50% stochactic
				'this method converge iterativly rapidly toward the price value for the 50% stochastic
				'update the phase detector with the latest parameters
				'we need two one with the gain and the other one without
				'run the filter with the special phase detector
				'ThisPricePeakMedian is the starting search point since it is not generally to far from the searched solution
				'the filter contain the price value that bring the 50% stochastic
				'Note: this is a very important price threshold value that can be used to help start or terminate a trade
				'the value that include the gain seem to be the most usuful to enter or leave a trade.
				MyPLLErrorDetectorForPriceStochacticMedian.Update(
					ThisVolatilityRegulated,
					0.0,
					0.0,
					ThisValueHigh,
					ThisValueLow)
				MyPLLErrorDetectorForPriceStochacticMedianWithGain.Update(
					ThisVolatilityRegulated,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisValueHigh,
					ThisValueLow)
				'MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction.Update(
				'  ThisVolatilityForStochasticPrediction,
				'  ThisGainPerYearPrediction,
				'  ThisGainPerYearDerivativePrediction,
				'  ThisValueHighPrediction,
				'  ThisValueLowPrediction)
				MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction.Update(
					ThisVolatilityForStochasticPrediction,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisValueHighPrediction,
					ThisValueLowPrediction)



				MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter.Update(
					ThisVolatilityRegulated,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisMeasurementPeakForGainPeakPrediction.High,
					ThisMeasurementPeakForGainPeakPrediction.Low)

				MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter.Update(
					ThisVolatilityRegulated,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisMeasurementPeakForGainPeakPredictionEstimate.High,
					ThisMeasurementPeakForGainPeakPredictionEstimate.Low)

				MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh.Update(
					ThisVolatilityRegulated,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisMeasurementPeakForGainPeakPredictionEstimateHigh.High,
					ThisMeasurementPeakForGainPeakPredictionEstimateHigh.Low)
				MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow.Update(
					ThisVolatilityRegulated,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisMeasurementPeakForGainPeakPredictionEstimateLow.High,
					ThisMeasurementPeakForGainPeakPredictionEstimateLow.Low)


				Dim ThisFilterPredictionGainYearly As Double = MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(
																												MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter.ToList.Last,
																												MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter.ToList.Last)

				ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
				MyListOfPeakValueGainPrediction.Add(ThisFilterPredictionGainYearly)

				If Me.Count = 0 Then
					'initialization
					'these are predictive value for the next sample 
					'one value is added initialy to shift the samples by 1 day toward it next day prediction
					'MyListOfPeakValueGainPrediction.Add(ThisFilterPredictionGainYearly)
					With MyPLLErrorDetectorForPriceStochacticMedian
						.ToList.Add(.ToList.Last)
						.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
						.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
					End With
					With MyPLLErrorDetectorForPriceStochacticMedianWithGain
						.ToList.Add(.ToList.Last)
						.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
						.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
					End With
					With MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction
						.ToList.Add(.ToList.Last)
						.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
						.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
					End With
					With MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh
						.ToList.Add(.ToList.Last)
						.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
						.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
					End With
					With MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow
						.ToList.Add(.ToList.Last)
						.ToListOfPriceMedianNextDayLow.Add(.ToListOfPriceMedianNextDayLow.Last)
						.ToListOfPriceMedianNextDayHigh.Add(.ToListOfPriceMedianNextDayHigh.Last)
					End With
				End If
				If Me.Count = 1000 Then
					'Debugger.Break()
					ThisProbHigh = ThisProbHigh
				End If

				If Me.Count > 0 Then
					MyStatisticRangeOfExcess.Run(Value, MyListOfPriceNextDailyLowWithGain.Last, MyListOfPriceNextDailyHighWithGain.Last)
				Else
					'initialization
					'these are predictive value for the next sample 
					'one extra value is added just at the beginning to shift the samples by 1 day toward it prediction
					MyListOfPriceNextDailyHigh.Add(ThisPriceNextDailyHigh)
					MyListOfPriceNextDailyLow.Add(ThisPriceNextDailyLow)
					MyListOfPriceNextDailyHighWithGain.Add(ThisPriceNextDailyHighWithGain)
					MyListOfPriceNextDailyLowWithGain.Add(ThisPriceNextDailyLowWithGain)
					MyListOfPriceNextDailyLowWithGainAtSigma2.Add(ThisPriceNextDailyLowWithGainAtSigma2)
					MyListOfPriceNextDailyLowWithGainAtSigma3.Add(ThisPriceNextDailyLowWithGainAtSigma3)
					MyListOfPriceNextDailyHighWithGainAtSigma2.Add(ThisPriceNextDailyHighWithGainAtSigma2)
					MyListOfPriceNextDailyHighWithGainAtSigma3.Add(ThisPriceNextDailyHighWithGainAtSigma3)
					MyListOfPriceNextDailyHighWithGainK2.Add(ThisPriceNextDailyHighWithGainK2)
					MyListOfPriceNextDailyLowWithGainK2.Add(ThisPriceNextDailyLowWithGainK2)
					MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyHighPreviousCloseToOpenWithGain)
					MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyLowPreviousCloseToOpenWithGain)

					'and initilize the queue use to measure the excess rate
					'always no excess at the beginning
					'MyQueueOfDailyRangeSigmaExcess.Enqueue(New Tuple(Of Integer, Integer, Integer)(0, 0, 0))
					'MySumOfDailyRangeSigmaExcess = 0
					'ThisProbabilityDailySigmaExcess = 0.0
					MyStatisticRangeOfExcess.Run(Value, Value.Last, Value.Last)
				End If
				MyListOfProbabilityDailySigmaExcess.Add(MyStatisticRangeOfExcess.BandLimitExcess)
				MyListOfProbabilityDailySigmaDoubleExcess.Add(MyStatisticRangeOfExcess.BandLimitDoubleExcess)
				MyListForVolatilityDetectorBalance.Add(MyStatisticRangeOfExcess.BandLimitHighLowBalance)
				'MyListOfProbabilityDailySigmaExcess.Add(ThisProbabilityDailySigmaExcess)
				'MyListForVolatilityDetectorBalance.Add(ThisProbabilityDailySigmaExcessHigh - ThisProbabilityDailySigmaExcessLow)

				'leave this here after the measurement of the probability of excess
				MyListOfPriceNextDailyHigh.Add(ThisPriceNextDailyHigh)
				MyListOfPriceNextDailyLow.Add(ThisPriceNextDailyLow)
				MyListOfPriceNextDailyHighWithGain.Add(ThisPriceNextDailyHighWithGain)
				MyListOfPriceNextDailyLowWithGain.Add(ThisPriceNextDailyLowWithGain)
				MyListOfPriceNextDailyLowWithGainAtSigma2.Add(ThisPriceNextDailyLowWithGainAtSigma2)
				MyListOfPriceNextDailyLowWithGainAtSigma3.Add(ThisPriceNextDailyLowWithGainAtSigma3)
				MyListOfPriceNextDailyHighWithGainAtSigma2.Add(ThisPriceNextDailyHighWithGainAtSigma2)
				MyListOfPriceNextDailyHighWithGainAtSigma3.Add(ThisPriceNextDailyHighWithGainAtSigma3)
				MyListOfPriceNextDailyHighWithGainK2.Add(ThisPriceNextDailyHighWithGainK2)
				MyListOfPriceNextDailyLowWithGainK2.Add(ThisPriceNextDailyLowWithGainK2)
				MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyHighPreviousCloseToOpenWithGain)
				MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen.Add(ThisPriceNextDailyLowPreviousCloseToOpenWithGain)
				MyListOfPriceNextDailyLowWithGainOpenToClose.Add(ThisStockPriceLowValueFromOpenToClose)
				MyListOfPriceNextDailyHighWithGainOpenToClose.Add(ThisStockPriceHighValueFromOpenToClose)


				ThisProbHigh = 1 - StockOption.StockPricePredictionInverse(
														 ThisRate,
														 Value.Last,
														 ThisGainPerYear,
														 ThisGainPerYearDerivative,
														 ThisVolatilityRegulated,
														 ThisValueHigh)

				ThisProbLow = StockOption.StockPricePredictionInverse(
												ThisRate,
												Value.Last,
												ThisGainPerYear,
												ThisGainPerYearDerivative,
												ThisVolatilityRegulated,
												ThisValueLow)
				ThisProbHighHalfRate = 1 - StockOption.StockPricePredictionInverse(ThisRateHalf, Value.Last, ThisGainPerYear, ThisGainPerYearDerivative, ThisVolatilityRegulated, ThisValueHigh)
				ThisProbLowHalfRate = StockOption.StockPricePredictionInverse(ThisRateHalf, Value.Last, ThisGainPerYear, ThisGainPerYearDerivative, ThisVolatilityRegulated, ThisValueLow)
				ThisProbHighHalfRate = ThisProbHigh
				'ThisProbLowHalfRate = ThisProbLow

				ThisStochasticResult = ThisProbHigh / (ThisProbHigh + ThisProbLow)
				ThisStochasticResultHalfRate = ThisProbHighHalfRate / (ThisProbHighHalfRate + ThisProbLowHalfRate)
				'average the result for the both rate 
				'ThisStochasticResult = (ThisStochasticResult + ThisStochasticResultHalfRate) / 2

				'MyListOfProbabilityBandHigh.Add((ThisProbHigh + ThisProbHighHalfRate) / 2)
				'MyListOfProbabilityBandLow.Add((ThisProbLow + ThisProbLowHalfRate) / 2)
				MyListOfProbabilityBandHigh.Add(ThisProbHigh)
				MyListOfProbabilityBandLow.Add(ThisProbLow)

				'calculate the probability to reach or exceed the Stock Price median at the end of the filter rate
				'for PriceStochacticMedianWithGain
				ThisProbOfStockMedian = 1 - StockOption.StockPricePredictionInverse(
					ThisRate,
					Value.Last,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisVolatilityRegulated,
					MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToList.Last)

				MyListOfProbabilityOfStockMedian.Add(ThisProbOfStockMedian)
				MyFilterVolatilityForPositifNegatif.Filter(Value, IsVolatityHoldToLast:=False)

				ThisGainFromStep = WaveForm.SignalLimit(((NUMBER_WORKDAY_PER_YEAR * Measure.Measure.GainLog(ThisPriceVolatilityHigh + ThisPriceVolatilityLow, ThisValueHigh + ThisValueLow) / ThisRate) / 2 + 0.5), 0.5, 0.5)
				'ThisGainFromStep = WaveForm.SignalLimit(((NUMBER_WORKDAY_PER_YEAR * Measure.Measure.GainLog(ThisPriceVolatilityHigh, ThisValueHigh) / ThisRate) / 2 + 0.5), 0.5, 0.5)
				MyListOfPriceVolatilityGain.Add(ThisGainFromStep)
				MyValueLast = Value
				'note for now always send the range volatility here
				'MyFilterStochastic.ListDataUpdate(ThisValueHigh, ThisValueLow, ThisFilterBasedVolatilityTotal, ThisStochasticResult)
				MyListOfPriceBandHigh.Add(ThisValueHigh)
				MyListOfPriceBandLow.Add(ThisValueLow)
				MyFilterLPOfStochasticSlow.Filter(ThisStochasticResult)
				MyListOfStochasticFast.Add(ThisStochasticResult)
				MyStocFastSlowLast = ThisStochasticResult - MyFilterLPOfStochasticSlow.FilterLast
				MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
				MyStocRangeVolatility = ThisFilterBasedVolatilityTotal
				MyStocFastLast = ThisStochasticResult
			Next
		End Sub

		''' <summary>
		''' see:https://www.mathworks.com/help/stats/lognstat.html
		'''
		''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
		''' for lab experiment see this excellent demo: 
		''' http://www.math.uah.edu/stat/apps/GeometricBrownianMotion.html
		''' See also for more general interest on statistic analysis and definition:
		''' http://www.math.uah.edu/stat/brown/index.html
		''' https://en.wikipedia.org/wiki/Log-normal_distribution
		''' 
		''' </summary>
		''' <param name="VolatilityStatistic"></param>
		''' <remarks></remarks>
		Private Function CalculateVolatilityStatisticVolatilitySimulated(ByVal VolatilityStatistic As IStatistical) As StatisticalDistribution
			Dim ThisVolatilityYangZhang As FilterVolatilityYangZhang
			Dim ThisStochasticProcess As MathProcess.StochasticProcess
			Dim ThisStochasticProcessData() As Double
			Dim ThisStatisticalDistributionForVolatilitySimulation = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
			Dim I As Integer

			'see wikipedia on lognormal distribution
			' https://en.wikipedia.org/wiki/Log-normal_distribution

			'ThisMu = Math.Log((VolatilityStatistic.Mean ^ 2) / (Math.Sqrt(VolatilityStatistic.Variance + VolatilityStatistic.Mean ^ 2)))
			'ThisSigma = Math.Sqrt(Math.Log(1 + (VolatilityStatistic.Variance / VolatilityStatistic.Mean ^ 2)))

			ThisStatisticalDistributionForVolatilitySimulation.Tag = Me.Tag
			ThisStochasticProcess = New MathProcess.StochasticProcess(ValueStart:=100, Gain:=0.0, Volatility:=VolatilityStatistic.Mean)
			'ThisFilterVolatility = New FilterVolatility(Me.Rate, FilterVolatility.enuVolatilityStatisticType.Standard)
			ThisVolatilityYangZhang = New FilterVolatilityYangZhang(Me.Rate, FilterVolatility.enuVolatilityStatisticType.Standard)

			ThisStochasticProcessData = ThisStochasticProcess.Samples(VolatilityStatistic.NumberPoint + Me.Rate)
			ThisVolatilityYangZhang.Filter(ThisStochasticProcessData)
			For I = Me.Rate To ThisStochasticProcessData.Length - 1
				ThisStatisticalDistributionForVolatilitySimulation.BucketFill(ThisVolatilityYangZhang.ToList.Item(I))
			Next
			Return ThisStatisticalDistributionForVolatilitySimulation
		End Function
#End Region
		Private Function CalculateStochasticPriceStop(
			ByVal IsFilterGainPriceStopOneSigmaEnabled As Boolean,
			ByVal IsSochasticPriceMedianIncludingGain As Boolean,
			ByVal IsPriceStopBoundToDailyOneSigmaEnabled As Boolean) As IList(Of Double)

			Dim ThisPriceStop As Double
			Dim I, J As Integer

			Dim ThisListOfPriceStopStochasticResult As List(Of Double) = New List(Of Double)
			'extract the data to a local reference list
			Dim ThisListOfPriceStochastic = Me.ToList
			Dim ThisListOfPriceStopFromStochasticMedianWithGain = Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedianWithGain)
			Dim ThisListOfPriceStopFromStochasticMedianNoGain = Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedian)
			Dim ThisListOfPriceBandVolatilityLow = Me.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityLow)
			Dim ThisListOfPriceBandVolatilityHigh = Me.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityHigh)
			Dim ThisListOfPriceStochasticMedianWithGainDailyBandLow = Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDown)
			Dim ThisListOfPriceStochasticMedianWithGainDailyBandHigh = Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUp)

			'set the price stochastic stop
			If IsFilterGainPriceStopOneSigmaEnabled Then
				'this is a closer tight stop that use the one sigma limit 
				ThisListOfPriceStopStochasticResult = New List(Of Double)
				If IsSochasticPriceMedianIncludingGain Then
					For I = 0 To Me.Count - 1
						J = I + 1
						If ThisListOfPriceStochastic(I) >= 0.5 Then
							If ThisListOfPriceBandVolatilityLow(J) > ThisListOfPriceStopFromStochasticMedianWithGain(J) Then
								ThisPriceStop = ThisListOfPriceBandVolatilityLow(J)
							Else
								ThisPriceStop = ThisListOfPriceStopFromStochasticMedianWithGain(J)
							End If
						Else
							If ThisListOfPriceBandVolatilityHigh(J) < ThisListOfPriceStopFromStochasticMedianWithGain(J) Then
								ThisPriceStop = ThisListOfPriceBandVolatilityHigh(J)
							Else
								ThisPriceStop = ThisListOfPriceStopFromStochasticMedianWithGain(J)
							End If
						End If
						ThisListOfPriceStopStochasticResult.Add(ThisPriceStop)
						If I = 0 Then
							ThisListOfPriceStopStochasticResult.Add(ThisPriceStop)
						End If
					Next
				Else
					ThisListOfPriceStopStochasticResult = New List(Of Double)(ThisListOfPriceStopFromStochasticMedianNoGain)
				End If
			Else
				If IsSochasticPriceMedianIncludingGain Then
					ThisListOfPriceStopStochasticResult = New List(Of Double)(ThisListOfPriceStopFromStochasticMedianWithGain)
				Else
					ThisListOfPriceStopStochasticResult = New List(Of Double)(ThisListOfPriceStopFromStochasticMedianNoGain)
				End If
			End If
			If IsPriceStopBoundToDailyOneSigmaEnabled Then
				Dim ThisPriceMedianDaily As Double
				Dim ThisPriceDailyBandHigh As Double
				Dim ThisPriceDailyBandLow As Double
				For I = 0 To ThisListOfPriceStopStochasticResult.Count - 1
					ThisPriceStop = ThisListOfPriceStopStochasticResult(I)
					ThisPriceDailyBandLow = ThisListOfPriceStochasticMedianWithGainDailyBandLow(I)
					ThisPriceDailyBandHigh = ThisListOfPriceStochasticMedianWithGainDailyBandHigh(I)
					If (ThisPriceStop > ThisPriceDailyBandLow) And (ThisPriceStop < ThisPriceDailyBandHigh) Then
						ThisPriceMedianDaily = (ThisPriceDailyBandLow + ThisPriceDailyBandHigh) / 2
						'if necessary replace the value
						If ThisPriceStop >= ThisPriceMedianDaily Then
							ThisListOfPriceStopStochasticResult(I) = ThisPriceDailyBandHigh
						Else
							ThisListOfPriceStopStochasticResult(I) = ThisPriceDailyBandLow
						End If
					End If
				Next
			End If
			Return ThisListOfPriceStopStochasticResult
		End Function

		Public ReadOnly Property ToStatisticalDistribution(ByVal Type As enuStatisticDistribution) As IStatisticalDistribution
			Get
				Select Case Type
					Case enuStatisticDistribution.Volatility
						Return MyStatisticalDistributionForVolatility
					Case enuStatisticDistribution.VolatilitySimulated
						If MyStatisticalDistributionForVolatilitySimulation Is Nothing Then
							'generate the simulation on the first time
							MyStatisticalDistributionForVolatilitySimulation = CalculateVolatilityStatisticVolatilitySimulated(MyFilterForVolatilityStatistic.FilterLast)
						End If
						Return MyStatisticalDistributionForVolatilitySimulation
					Case enuStatisticDistribution.VolatilityDistributionPositive
						Return MyStatisticalDistributionForVolatilityPositive
					Case enuStatisticDistribution.VolatilityDistributionNegative
						Return MyStatisticalDistributionForVolatilityNegative
					Case Else
						Throw New NotImplementedException("ToStatisticalDistribution")
				End Select
			End Get
		End Property
#Region "Friend function and properties"
		Friend ReadOnly Property ToFilterPrice() As IFilter
			Get
				Return MyFilterLPForPrice
			End Get
		End Property

		Friend ReadOnly Property ToListOfGain() As IList(Of Double)
			Get
				Return MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear
			End Get
		End Property

		Friend ReadOnly Property ToListOfGainDerivative() As IList(Of Double)
			Get
				Return MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative
			End Get
		End Property
#End Region
#Region "IStochastic"
		Public ReadOnly Property Count As Integer Implements IStochastic.Count
			Get
				Return MyFilterLPOfStochasticSlow.Count
			End Get
		End Property

		Private Function Filter(ByRef Value As Double) As Double Implements IStochastic.Filter
			Throw New NotImplementedException()
		End Function

		Private Function Filter(ByRef Value() As Double) As Double() Implements IStochastic.Filter
			Throw New NotImplementedException()
		End Function

		Private Function Filter(Value As Single) As Double Implements IStochastic.Filter
			Throw New NotImplementedException()
		End Function

		Public Function Filter(ByRef Value As IPriceVol) As Double Implements IStochastic.Filter
			Return Me.FilterLocal(Value)
		End Function

		'Public Async Function FilterAsync(ReportPrices As RecordPrices, IsUseParallelBlock As Boolean) As Task(Of Boolean) Implements IFilterRunAsync.FilterAsync
		'	Dim ThisTaskRun As Task(Of Boolean)

		'	If IsUseParallelBlock Then
		'		ThisTaskRun = New Task(Of Boolean)(
		'			Function()
		'				Return Me.FilterAsync1(ReportPrices).Result
		'			End Function)

		'		ThisTaskRun.Start()
		'		Await ThisTaskRun
		'		Return ThisTaskRun.Result
		'	Else
		'		ThisTaskRun = FilterAsync(ReportPrices)
		'		Await ThisTaskRun
		'		Return ThisTaskRun.Result
		'	End If
		'End Function

		''' <summary>
		''' simpler version re-written in a more efficient manner for async operation
		''' </summary>
		''' <param name="ReportPrices"></param>
		''' <param name="IsUseParallelBlock"></param>
		''' <returns></returns>
		Public Async Function FilterAsync(ReportPrices As RecordPrices, IsUseParallelBlock As Boolean) As Task(Of Boolean) Implements IFilterRunAsync.FilterAsync

			If IsUseParallelBlock Then
				' Offload to background thread
				Return Await Task.Run(Function() FilterAsync1(ReportPrices))
			Else
				' Run directly without blocking or parallelization
				Return FilterSync(ReportPrices)
			End If
		End Function

		Public Function FilterSync(ReportPrices As RecordPrices) As Boolean
			Dim I As Integer

			For I = 0 To ReportPrices.NumberPoint - 1
				If I = 500 Then
					I = I
				End If
				Me.FilterLocal(ReportPrices.GetPriceVolInterface(I))
			Next

			Return True
		End Function

		'Public Async Function FilterAsync(ReportPrices As RecordPrices) As Task(Of Boolean) Implements IFilterRunAsync.FilterAsync
		'	Dim ThisTaskRun As Task(Of Boolean)

		'	ThisTaskRun = New Task(Of Boolean)(
		'		Function()
		'			Dim I As Integer

		'			For I = 0 To ReportPrices.NumberPoint - 1
		'				Me.FilterLocal(ReportPrices.GetPriceVolInterface(I))
		'			Next
		'			Return True
		'		End Function)

		'	ThisTaskRun.Start()
		'	Await ThisTaskRun
		'	Return ThisTaskRun.Result
		'End Function

		Public Async Function FilterAsync(ReportPrices As RecordPrices) As Task(Of Boolean) Implements IFilterRunAsync.FilterAsync
			Return Await Task.Run(Function() FilterSync(ReportPrices))
		End Function




		Private Async Function FilterAsync1(ByVal ReportPrices As YahooAccessData.RecordPrices) As Task(Of Boolean)
			Dim ThisListOfPLLErrorDetectorWithGain As List(Of FilterPLLDetectorForVolatilitySigmaAsync)
			Dim ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen As List(Of FilterPLLDetectorForVolatilitySigmaAsync)
			Dim ThisNumberOfPointPerThread As Integer
			Dim ThisNumberOfPointPerThreadToAdd As Integer
			Dim MyTasksForWaitUpdateAll As List(Of Task(Of Boolean))
			Dim ThisStartPoint, ThisStopPoint As Integer
			Dim ThisNumberPoint As Integer = ReportPrices.NumberPoint
			Dim ThisFilterPLLDetectorForVolatilitySigmaAsync As FilterPLLDetectorForVolatilitySigmaAsync
			Dim ThisFilterPLLDetectorForVolatilitySigmaCloseToOpenAsync As FilterPLLDetectorForVolatilitySigmaAsync
			Dim I As Integer


			Throw New NotImplementedException
			'adjust the the number of threading based on the number of points
			If ThisNumberPoint > MyProcessorCount Then
				ThisNumberOfPointPerThread = ThisNumberPoint \ MyProcessorCount
				ThisNumberOfPointPerThreadToAdd = ThisNumberPoint Mod MyProcessorCount
			Else
				'do not thread with small number of point
				ThisNumberOfPointPerThread = ThisNumberPoint
				ThisNumberOfPointPerThreadToAdd = 0
			End If
			'this first section is fast preparation calculation and is not threaded
			Me.FilterLocalPart1(ReportPrices)
			'ThisListOfPLLErrorDetector = New List(Of FilterPLLDetectorForVolatilitySigmaAsync)
			ThisListOfPLLErrorDetectorWithGain = New List(Of FilterPLLDetectorForVolatilitySigmaAsync)
			ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen = New List(Of FilterPLLDetectorForVolatilitySigmaAsync)
			ThisStartPoint = 0

			Do
				ThisStopPoint = ThisStartPoint + (ThisNumberOfPointPerThread - 1)
				If ThisNumberOfPointPerThreadToAdd > 0 Then
					ThisStopPoint = ThisStopPoint + 1
					ThisNumberOfPointPerThreadToAdd = ThisNumberOfPointPerThreadToAdd - 1
				End If

				ThisFilterPLLDetectorForVolatilitySigmaAsync = New FilterPLLDetectorForVolatilitySigmaAsync(
						MyStockPriceVolatilityPredictionBandWithGain,
						StartPoint:=ThisStartPoint,
						StopPoint:=ThisStopPoint,
						Rate:=Me.Rate,
						ToCountLimit:=FILTER_PLL_DETECTOR_COUNT_LIMIT,
						ToErrorLimit:=FILTER_PLL_DETECTOR_ERROR_LIMIT) With {.Tag = Me.Tag, .IsUseFeedbackRegulatedVolatilityFastAttackEvent = Me.IsUseFeedbackRegulatedVolatilityFastAttackEventLocal}

				ThisFilterPLLDetectorForVolatilitySigmaCloseToOpenAsync = New FilterPLLDetectorForVolatilitySigmaAsync(
						MyStockPriceVolatilityPredictionBandWithGainCloseToOpen,
						StartPoint:=ThisStartPoint,
						StopPoint:=ThisStopPoint,
						Rate:=Me.Rate,
						ToCountLimit:=FILTER_PLL_DETECTOR_COUNT_LIMIT,
						ToErrorLimit:=FILTER_PLL_DETECTOR_ERROR_LIMIT) With {.Tag = Me.Tag, .IsUseFeedbackRegulatedVolatilityFastAttackEvent = Me.IsUseFeedbackRegulatedVolatilityFastAttackEventLocal}


				ThisListOfPLLErrorDetectorWithGain.Add(ThisFilterPLLDetectorForVolatilitySigmaAsync)
				ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen.Add(ThisFilterPLLDetectorForVolatilitySigmaCloseToOpenAsync)

				'update the startpoint
				ThisStartPoint = ThisStartPoint + (ThisStopPoint - ThisStartPoint) + 1
				'Exit Do
			Loop Until ThisStartPoint = ThisNumberPoint
			MyTasksForWaitUpdateAll = New List(Of Task(Of Boolean))
			For Each ThisPLLErrorDetector In ThisListOfPLLErrorDetectorWithGain
				'run all in parallel
				MyTasksForWaitUpdateAll.Add(ThisPLLErrorDetector.UpdateAsync)
				'run them one after the other for debugging
				'Await ThisPLLErrorDetector.UpdateAsync
			Next
			For Each ThisPLLErrorDetector In ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen
				'run all in parallel
				MyTasksForWaitUpdateAll.Add(ThisPLLErrorDetector.UpdateAsync)
				'run them one after the other for debugging
				'Await ThisPLLErrorDetector.UpdateAsync
			Next
			Dim ThisTaskWaitAll = Task.WhenAll(MyTasksForWaitUpdateAll)
			Await ThisTaskWaitAll

			'pack all the result data together
			'add the first element in the prediction list before we start
			MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList.Add(ThisListOfPLLErrorDetectorWithGain.First.ToList.First)
			MyListForVolatilityRegulatedPreviousCloseToOpenWithGain.Add(ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen.First.ToList.First)
			For Each ThisPLLErrorDetector In ThisListOfPLLErrorDetectorWithGain
				For Each ThisPLLVolatilityResult In ThisPLLErrorDetector.ToList
					MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain.Add(ThisPLLVolatilityResult)
					MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.Filter(ThisPLLVolatilityResult)
				Next
				'For Each ThisPLLVolatilityResult In ThisPLLErrorDetector.ToListOfProbabilityOfExcess
				'  MyListOfProbabilityDailySigmaExcess.Add(ThisPLLVolatilityResult)
				'Next
				'For Each ThisPLLVolatilityResult In ThisPLLErrorDetector.ToListOfProbabilityOfExcess
				'  MyListForVolatilityDetectorBalance.Add(ThisPLLVolatilityResult)
				'Next
			Next
			For Each ThisPLLErrorDetector In ThisListOfPLLErrorDetectorWithGainPreviousCloseToOpen
				For Each ThisPLLVolatilityResult In ThisPLLErrorDetector.ToList
					MyListForVolatilityRegulatedPreviousCloseToOpenWithGain.Add(ThisPLLVolatilityResult)
				Next
			Next
			Dim ThisVolatilityTotal As Double
			Dim ThisVolatilityPreviousCloseToOpen As Double
			For I = 0 To MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain.Count - 1
				ThisVolatilityTotal = MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain(I)
				ThisVolatilityPreviousCloseToOpen = MyListForVolatilityRegulatedPreviousCloseToOpenWithGain(I)
				If ThisVolatilityPreviousCloseToOpen > ThisVolatilityTotal Then
					MyListForVolatilityRegulatedFromOpenToCloseWithGain.Add(0.0)
				Else
					MyListForVolatilityRegulatedFromOpenToCloseWithGain.Add(Math.Sqrt((ThisVolatilityTotal ^ 2) - (ThisVolatilityPreviousCloseToOpen ^ 2)))
				End If
			Next
			'the last section is fast and not threaded
			Me.FilterLocalPart2(ReportPrices)
			Return True
		End Function

		Private Function Filter(ByRef Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double Implements IStochastic.Filter
			Throw New NotImplementedException()
		End Function

		Private Function Filter(ByRef Value As IPriceVol, FilterRate As Integer) As Double Implements IStochastic.Filter
			Throw New NotImplementedException()
		End Function

		Private Function FilterBackTo(ByRef Value As Double, Optional IsPreFilter As Boolean = True) As Double Implements IStochastic.FilterBackTo
			Throw New NotImplementedException()
		End Function

		Public Function FilterLast() As Double Implements IStochastic.FilterLast
			Return MyFilterLPOfStochasticSlow.FilterLast
		End Function

		Public Function FilterLast(Type As IStochastic.enuStochasticType) As Double Implements IStochastic.FilterLast
			Select Case Type
				Case IStochastic.enuStochasticType.FastSlow
					Return MyStocFastSlowLast
				Case IStochastic.enuStochasticType.Fast
					Return MyStocFastLast
				Case IStochastic.enuStochasticType.Slow
					Return MyFilterLPOfStochasticSlow.FilterLast
				Case IStochastic.enuStochasticType.RangeVolatility
					Return MyListOfPriceRangeVolatility.Last
				Case IStochastic.enuStochasticType.PriceBandHigh
					Return MyListOfPriceBandHigh.Last
				Case IStochastic.enuStochasticType.PriceBandLow
					Return MyListOfPriceBandLow.Last
				Case IStochastic.enuStochasticType.PriceBandHighPrediction
					Return MyListOfPriceBandHighPrediction.Last
				Case IStochastic.enuStochasticType.PriceBandLowPrediction
					Return MyListOfPriceBandLowPrediction.Last
				Case Else
					Return MyStocFastSlowLast
			End Select
		End Function

		Public Function PredictionOpenToClose(
			ByVal PriceValue As Double,
			ByVal MarketTimeToCloseInDay As Double,
			Optional ByVal SigmaRangeProbability As Double = GAUSSIAN_PROBABILITY_SIGMA1) As IPriceVol

			Dim ThisFilterBasedVolatilityFromOpenToClose = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.FullDay).Last

			Return Me.PredictionOpenToClose(
				PriceValue,
				MarketTimeToCloseInDay,
				ThisFilterBasedVolatilityFromOpenToClose,
				ThisFilterBasedVolatilityFromOpenToClose,
				SigmaRangeProbability)

		End Function

		Public Function PredictionOpenToClose(
			ByVal PriceValue As Double,
			ByVal MarketTimeToCloseInDay As Double,
			ByVal Volatility As Double,
			Optional ByVal SigmaRangeProbability As Double = GAUSSIAN_PROBABILITY_SIGMA1) As IPriceVol

			Return Me.PredictionOpenToClose(
				PriceValue,
				MarketTimeToCloseInDay,
				Volatility,
				Volatility,
				SigmaRangeProbability)

		End Function

		Public Function PredictionOpenToClose(
			ByVal PriceValue As Double,
			ByVal MarketTimeToCloseInDay As Double,
			ByVal VolatilityPositif As Double,
			ByVal VolatilityNegatif As Double,
			Optional ByVal SigmaRangeProbability As Double = GAUSSIAN_PROBABILITY_SIGMA1) As IPriceVol

			Dim ThisFilterBasedVolatilityFromOpenToClose As Double
			Dim ThisGainPerYear As Double
			Dim ThisGainPerYearDerivative As Double
			Dim ThisProbabilityHigh As Double
			Dim ThisProbabilityLow As Double
			Dim ThisPriceMedianLow As Double
			Dim ThisPriceMedianHigh As Double


			If SigmaRangeProbability < 0.5 Then
				SigmaRangeProbability = 1 - SigmaRangeProbability
			End If
			ThisProbabilityHigh = 0.5 + SigmaRangeProbability / 2
			ThisProbabilityLow = 1 - ThisProbabilityHigh
			ThisFilterBasedVolatilityFromOpenToClose = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToClose).Last
			With MyFilterPLLForGain.AsIFilterPrediction
				ThisGainPerYear = .ToListOfGainPerYear.Last
				ThisGainPerYearDerivative = .ToListOfGainPerYearDerivative.Last
			End With
			Dim ThisPriceVolStart = PriceValue
			'Dim ThisDateStartTime As Date = ThisPriceVolStart.DateDay.Date.AddSeconds(YahooAccessData.ReportDate.MARKET_OPEN_TIME_SEC_DEFAULT)
			'Dim ThisDateStopTime As Date = ThisPriceVolStart.DateDay.Date.AddSeconds(YahooAccessData.ReportDate.MARKET_CLOSE_TIME_SEC_DEFAULT)
			'Dim ThisStepSizeInDay = ThisTimeMarketMaximumOpenTimePeriodInDay / NumberStepPerDay
			Dim ThisPriceVol As IPriceVol = New PriceVol
			ThisPriceVol.Open = CSng(PriceValue)
			ThisPriceVol.Low = CSng(StockOption.StockPricePrediction(
				MarketTimeToCloseInDay,
				PriceValue,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				VolatilityNegatif,
				ThisProbabilityLow))
			ThisPriceVol.High = CSng(StockOption.StockPricePrediction(
				MarketTimeToCloseInDay,
				PriceValue,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				VolatilityPositif,
				ThisProbabilityHigh))

			ThisPriceMedianLow = StockOption.StockPricePrediction(
				MarketTimeToCloseInDay,
				PriceValue,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				VolatilityNegatif,
				0.5)

			ThisPriceMedianHigh = StockOption.StockPricePrediction(
				MarketTimeToCloseInDay,
				PriceValue,
				ThisGainPerYear,
				ThisGainPerYearDerivative,
				VolatilityPositif,
				0.5)

			ThisPriceVol.Last = CSng((ThisPriceMedianLow + ThisPriceMedianHigh) / 2)
			Return ThisPriceVol
		End Function


		Public Function FilterPredictionToClose(ByRef Index As Integer, ByVal NumberStepPerDay As Integer) As IList(Of IPriceVol)
			Dim ThisFilterBasedVolatilityFromOpenToClose As Double
			Dim ThisGainPerYear As Double
			Dim ThisGainPerYearDerivative As Double
			Dim ThisTimeToMarketInDay As Double
			Dim ThisTimeMarketMaximumOpenTimePeriodInDay As Double
			Dim ThisPriceNextDailyHighPreviousCloseToOpenWithGain As Double
			Dim ThisPriceNextDailyLowPreviousCloseToOpenWithGain As Double
			Dim ThisPriceNextDailyHighWithGainOpenStart As Double
			Dim ThisPriceNextDailyLowWithGainOpenStart As Double


			ThisFilterBasedVolatilityFromOpenToClose = MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToClose)(Index)
			ThisTimeMarketMaximumOpenTimePeriodInDay = YahooAccessData.ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_HOUR_DEFAULT / 24

			'get the predictive value by adding 1 to the index
			ThisPriceNextDailyHighWithGainOpenStart = MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen(Index + 1)
			ThisPriceNextDailyLowWithGainOpenStart = MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen(Index + 1)
			Dim ThisResult As New List(Of IPriceVol)


			With MyFilterPLLForGain.AsIFilterPrediction
				ThisGainPerYear = .ToListOfGainPerYear(Index)
				ThisGainPerYearDerivative = .ToListOfGainPerYearDerivative(Index)
			End With
			Dim ThisPriceVolStart = MyListOfValue(Index)

			Dim ThisDateStartTime As Date = ThisPriceVolStart.DateDay.Date.AddSeconds(YahooAccessData.ReportDate.MARKET_OPEN_TIME_SEC_DEFAULT)
			Dim ThisDateStopTime As Date = ThisPriceVolStart.DateDay.Date.AddSeconds(YahooAccessData.ReportDate.MARKET_CLOSE_TIME_SEC_DEFAULT)
			Dim ThisStepSizeInDay = ThisTimeMarketMaximumOpenTimePeriodInDay / NumberStepPerDay
			Dim ThisPriceVol As PriceVol
			ThisTimeToMarketInDay = 0
			Do
				ThisPriceNextDailyHighPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
					ThisTimeToMarketInDay,
					ThisPriceNextDailyHighWithGainOpenStart,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisFilterBasedVolatilityFromOpenToClose,
					GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)

				ThisPriceNextDailyLowPreviousCloseToOpenWithGain = StockOption.StockPricePrediction(
					ThisTimeToMarketInDay,
					ThisPriceNextDailyLowWithGainOpenStart,
					ThisGainPerYear,
					ThisGainPerYearDerivative,
					ThisFilterBasedVolatilityFromOpenToClose,
					GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)

				ThisPriceVol = New PriceVol(ThisPriceVolStart)
				ThisPriceVol.DateLastTrade = ThisDateStartTime.AddDays(ThisTimeToMarketInDay)
				ThisPriceVol.High = CSng(ThisPriceNextDailyHighPreviousCloseToOpenWithGain)
				ThisPriceVol.Low = CSng(ThisPriceNextDailyLowPreviousCloseToOpenWithGain)
				ThisPriceVol.Last = (ThisPriceVol.High + ThisPriceVol.Low) / 2
				ThisTimeToMarketInDay = ThisTimeToMarketInDay + ThisStepSizeInDay
				ThisResult.Add(ThisPriceVol)
			Loop Until ThisTimeToMarketInDay > ThisTimeMarketMaximumOpenTimePeriodInDay
			Return ThisResult
		End Function

		Private Function FilterPredictionNext(ByRef Value As Double) As Double Implements IStochastic.FilterPredictionNext
			Throw New NotImplementedException
		End Function

		Public Function FilterPriceBandHigh() As Double Implements IStochastic.FilterPriceBandHigh
			Return MyListOfPriceBandHigh.Last
		End Function

		Public Function FilterPriceBandLow() As Double Implements IStochastic.FilterPriceBandLow
			Return MyListOfPriceBandLow.Last
		End Function

		Public Property IsUseFeedbackRegulatedVolatility As Boolean

		Private IsUseFeedbackRegulatedVolatilityFastAttackEventLocal As Boolean
		Public Property IsUseFeedbackRegulatedVolatilityFastAttackEvent As Boolean
			Get
				Return IsUseFeedbackRegulatedVolatilityFastAttackEventLocal
			End Get
			Set(value As Boolean)
				IsUseFeedbackRegulatedVolatilityFastAttackEventLocal = value
				MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain.IsUseFeedbackRegulatedVolatilityFastAttackEvent = IsUseFeedbackRegulatedVolatilityFastAttackEventLocal
			End Set
		End Property

		Private IsFilterPeakLocal As Boolean
		Public Property IsFilterPeak As Boolean Implements IStochastic.IsFilterPeak
			Get
				Return IsFilterPeakLocal
			End Get
			Set(value As Boolean)
				IsFilterPeakLocal = value
				MyMeasurePeakValueRange.IsFilterEnabled = value
			End Set
		End Property

		Private IsFilterRangeLocal As Boolean
		Public Property IsFilterRange As Boolean Implements IStochastic.IsFilterRange
			Get
				Return IsFilterRangeLocal
			End Get
			Set(value As Boolean)
				IsFilterRangeLocal = value
			End Set
		End Property

		Public Function Last() As Double Implements IStochastic.Last
			Return MyValueLast.Last
		End Function

		Public ReadOnly Property Max(Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Max
			Get
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.Max
					Case IStochastic.enuStochasticType.Fast
						Return 1.0
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.Max
					Case Else
						Return MyListOfStochasticFastSlow.Max
				End Select
			End Get
		End Property

		Public ReadOnly Property Min(Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Min
			Get
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.Min
					Case IStochastic.enuStochasticType.Fast
						Return 0.0
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.Min
					Case Else
						Return MyListOfStochasticFastSlow.Min
				End Select
			End Get
		End Property

		Public Property Rate(Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Integer Implements IStochastic.Rate
			Get
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyRate
					Case IStochastic.enuStochasticType.Fast
						Return MyRatePreFilter
					Case IStochastic.enuStochasticType.Slow
						Return CInt(MyRateOutput)
					Case IStochastic.enuStochasticType.PriceVolatilityRegulated,
							 IStochastic.enuStochasticType.PriceStandardVolatility
						Return MyRateForVolatility
					Case Else
						Return 0
				End Select
			End Get
			Set(value As Integer)
				'do not set the rate here
			End Set
		End Property

		Private _Tag As String
		Public Property Tag As String Implements IStochastic.Tag
			Get
				Return _Tag
			End Get
			Set(value As String)
				_Tag = value
			End Set
		End Property

		Private _Symbol As String
		Public Property Symbol As String
			Get
				Return _Symbol
			End Get
			Set(value As String)
				_Tag = _Symbol
			End Set
		End Property

		Private Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double, Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
			Throw New NotImplementedException
		End Function

		Private Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double, Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
			Throw New NotImplementedException
		End Function

		Private Function ToArray(Optional Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
			Throw New NotImplementedException
		End Function

		Public ReadOnly Property ToList As System.Collections.Generic.IList(Of Double) Implements IStochastic.ToList
			Get
				'Return MyFilterLPOfStochasticSlow.ToList
				Return MyListOfProbabilityOfStockMedian
			End Get
		End Property

		Private MyDictionaryOfStochasticType As Dictionary(Of IStochastic.enuStochasticType, IList(Of Double))
		Public ReadOnly Property ToList(Type As IStochastic.enuStochasticType) As System.Collections.Generic.IList(Of Double) Implements IStochastic.ToList
			Get
				If MyDictionaryOfStochasticType Is Nothing Then
					MyDictionaryOfStochasticType = New Dictionary(Of IStochastic.enuStochasticType, IList(Of Double))
				End If
				If MyDictionaryOfStochasticType.ContainsKey(Type) Then
					Return MyDictionaryOfStochasticType(Type)
				Else
					MyDictionaryOfStochasticType.Add(Type, Me.ToListData(Type))
					Return MyDictionaryOfStochasticType(Type)
				End If
			End Get
		End Property

		Private Function ToListData(ByVal Type As IStochastic.enuStochasticType) As IList(Of Double)
			Select Case Type
				Case IStochastic.enuStochasticType.PriceStochacticVolatilityPositive
					Return MyFilterVolatilityForPositifNegatif.ToList(FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToHighClose)
				Case IStochastic.enuStochasticType.PriceStochacticVolatilityNegative
					Return MyFilterVolatilityForPositifNegatif.ToList(FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToLowClose)
				Case IStochastic.enuStochasticType.PriceStochacticVolatilityPositiveToNegativeRatio
					Return MyFilterVolatilityForPositifNegatif.ToList(FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio)
				Case IStochastic.enuStochasticType.PriceStochacticVolatilityPositiveToNegativeRatioFiltered
					Return MyFilterVolatilityForPositifNegatif.ToList(FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatioFiltered)
				Case IStochastic.enuStochasticType.ProbabilityHigh
					Return MyListOfProbabilityBandHigh
				Case IStochastic.enuStochasticType.ProbabilityLow
					Return MyListOfProbabilityBandLow
				Case IStochastic.enuStochasticType.PriceBandVolatilityHigh
					Return MyListOfPriceVolatilityHigh
				Case IStochastic.enuStochasticType.PriceBandVolatilityLow
					Return MyListOfPriceVolatilityLow
				Case IStochastic.enuStochasticType.PriceVolatilityPDF
					MyStatisticalDistributionForVolatility.Refresh(IStatisticalDistribution.enuRefreshType.ArrayStandard)
					Return MyStatisticalDistributionForVolatility.ToListOfPDF
				Case IStochastic.enuStochasticType.PriceVolatilityLCR
					MyStatisticalDistributionForVolatility.Refresh(IStatisticalDistribution.enuRefreshType.ArrayStandard)
					Return MyStatisticalDistributionForVolatility.ToListOfLCR
				Case IStochastic.enuStochasticType.PriceVolatilityPDFSimulated
					If MyStatisticalDistributionForVolatilitySimulation Is Nothing Then
						'generate the simulation on the first time
						MyStatisticalDistributionForVolatilitySimulation = CalculateVolatilityStatisticVolatilitySimulated(MyFilterForVolatilityStatistic.FilterLast)
					End If
					MyStatisticalDistributionForVolatilitySimulation.Refresh(IStatisticalDistribution.enuRefreshType.ArrayStandard)
					Return MyStatisticalDistributionForVolatilitySimulation.ToListOfPDF
				Case IStochastic.enuStochasticType.PriceVolatilityLCRSimulated
					If MyStatisticalDistributionForVolatilitySimulation Is Nothing Then
						'generate the simulation on the first time
						MyStatisticalDistributionForVolatilitySimulation = CalculateVolatilityStatisticVolatilitySimulated(MyFilterForVolatilityStatistic.FilterLast)
					End If
					MyStatisticalDistributionForVolatilitySimulation.Refresh(IStatisticalDistribution.enuRefreshType.ArrayStandard)
					Return MyStatisticalDistributionForVolatilitySimulation.ToListOfLCR
				Case IStochastic.enuStochasticType.PriceBandVolatilityGain
					Return MyListOfPriceVolatilityGain
				Case IStochastic.enuStochasticType.PriceProbabilityMedian
					Return MyListOfPriceProbabilityMedian
				Case IStochastic.enuStochasticType.PriceStandardVolatility, IStochastic.enuStochasticType.RangeVolatility
					Return MyListOfPriceRangeVolatility
				Case IStochastic.enuStochasticType.RangeVolatilityFromPreviousCloseToOpen
					Return MyFilterVolatilityYangZhangForStatistic.ToList(Type:=FilterVolatilityYangZhang.enuVolatilityDailyPeriodType.PreviousCloseToOpen)
				Case IStochastic.enuStochasticType.RangeVolatilityFromPreviousCloseToOpenRatio
					Return MyListOfPriceRangeVolatilityFromPreviousCloseToOpenRatio
				Case IStochastic.enuStochasticType.RangeVolatilityRegulatedFromPreviousCloseToOpen
					Return MyListForVolatilityRegulatedPreviousCloseToOpenWithGain
				Case IStochastic.enuStochasticType.RangeVolatilityFromOpenToClose
					Return MyListForVolatilityRegulatedFromOpenToCloseWithGain
				Case IStochastic.enuStochasticType.ProbabilityFromBandVolatility
					Return MyFilterLPForProbabilityFromBandVolatility.ToList
				Case IStochastic.enuStochasticType.TimeProbabilityOfPriceVolatility
					Return MyListOfPriceVolatilityTimeProbability
				Case IStochastic.enuStochasticType.StochasticSlowFromPriceBandVolatilityLow
					Return MyFilterLPForStochasticFromPriceVolatilityLow.ToList
				Case IStochastic.enuStochasticType.StochasticSlowFromPriceBandVolatilityHigh
					Return MyFilterLPForStochasticFromPriceVolatilityHigh.ToList
				Case IStochastic.enuStochasticType.StochasticSlowFromPricePeakMedian
					Throw New NotSupportedException
				Case IStochastic.enuStochasticType.PriceStochacticMedian
					Return MyPLLErrorDetectorForPriceStochacticMedian.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGain
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToList
				Case IStochastic.enuStochasticType.PriceStochasticMedianNextDayLow
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToListOfPriceMedianNextDayLow
				Case IStochastic.enuStochasticType.PriceStochasticMedianNextDayHigh
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToListOfPriceMedianNextDayHigh
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainHisteresisHigh
					Throw New NotSupportedException
						'Return MyPLLErrorDetectorForPriceStochacticMedianWithGainFast.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainHisteresisLow
					Throw New NotSupportedException
						'Return MyPLLErrorDetectorForPriceStochacticMedianWithGainSlow.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainPrediction
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainPredictionHigh
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainPredictionLow
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow.ToList
				Case IStochastic.enuStochasticType.PricePeakValueGainPrediction
					Return MyListOfPeakValueGainPrediction.ToList
				Case IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyDown
					Return MyListOfPriceNextDailyLow
				Case IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyUp
					Return MyListOfPriceNextDailyHigh
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDown
					Return MyListOfPriceNextDailyLowWithGain
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose
					Return MyListOfPriceNextDailyHighWithGain
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose
					Return MyListOfPriceNextDailyHighWithGain
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUp
					Return MyListOfPriceNextDailyHighWithGain
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownDay2
					Return MyListOfPriceNextDailyLowWithGainK2
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpDay2
					Return MyListOfPriceNextDailyHighWithGainK2
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpAtSigma2
					Return MyListOfPriceNextDailyHighWithGainAtSigma2
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpAtSigma3
					Return MyListOfPriceNextDailyHighWithGainAtSigma3
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyLowAtSigma2
					Return MyListOfPriceNextDailyLowWithGainAtSigma2
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyLowAtSigma3
					Return MyListOfPriceNextDailyLowWithGainAtSigma3
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyDownToOpen
					Return MyListOfPriceNextDailyLowWithGainPreviousCloseToOpen
				Case IStochastic.enuStochasticType.PriceStochacticMedianWithGainRangeDailyUpToOpen
					Return MyListOfPriceNextDailyHighWithGainPreviousCloseToOpen
				Case IStochastic.enuStochasticType.ProbabilityPriceDailySigmaExceeded
					Return MyListOfProbabilityDailySigmaExcess
				Case IStochastic.enuStochasticType.ProbabilityPriceDailySigmaDoubleExceeded
					Return MyListOfProbabilityDailySigmaDoubleExcess
						'Return MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain.ToListOfProbabilityOfExcess
				Case IStochastic.enuStochasticType.PriceVolatilityRegulated
					Return MyListOfVolatilityRegulatedFromPreviousCloseToCloseWithGain
				Case IStochastic.enuStochasticType.PriceVolatilityRegulatedPrediction
					Return MyFilterPLLForVolatilityRegulatedFromPreviousCloseToCloseWithGain.ToList
				Case IStochastic.enuStochasticType.PriceVolatilityLastPointTrail
					Return MyFilterVolatilityYangZhangForStatisticLastPointTrail.ToList
				Case IStochastic.enuStochasticType.PriceVolatilityDetectorBalance
					'Return MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain.ToListOfProbabilityOfExcessBalance
					Return MyListForVolatilityDetectorBalance
						'Return MyPLLErrorDetectorForVolatilityPredictionFromPreviousCloseToCloseWithGain.ToListOfConvergence
				Case IStochastic.enuStochasticType.PriceStochasticMedianVolatility
					Return MyPLLErrorDetectorForPriceStochacticMedianWithGain.ToListOfVolatility
				Case IStochastic.enuStochasticType.FastSlow
					Return MyListOfStochasticFastSlow
				Case IStochastic.enuStochasticType.Fast
					Return MyListOfStochasticFast
				Case IStochastic.enuStochasticType.Slow
					Return MyFilterLPOfStochasticSlow.ToList
				Case IStochastic.enuStochasticType.PriceBandHigh
					Return MyListOfPriceBandHigh
				Case IStochastic.enuStochasticType.PriceBandLow
					Return MyListOfPriceBandLow
				Case IStochastic.enuStochasticType.PriceBandHighPrediction
					Return MyListOfPriceBandHighPrediction
				Case IStochastic.enuStochasticType.PriceBandLowPrediction
					Return MyListOfPriceBandLowPrediction
				Case IStochastic.enuStochasticType.RangeVolatility
					Return MyListOfPriceRangeVolatility
				Case IStochastic.enuStochasticType.PriceGainPerYear
					Return MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear
				Case IStochastic.enuStochasticType.PriceGainPerYearDerivative
					Return MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative
				Case IStochastic.enuStochasticType.ProbabilityHigh, IStochastic.enuStochasticType.ProbabilityLow
					Throw New NotSupportedException
				Case Else
					Return MyListOfStochasticFastSlow
			End Select
		End Function


		Public Overrides Function ToString() As String Implements IStochastic.ToString
			Return Me.FilterLast.ToString
		End Function
#End Region
#Region "IStochasticPriceGain"
		Private Sub Init(
			FilterRate As Integer,
			FilterRateForGainMeasurement As Integer,
			IsGainFunctionWeightedMethod As Boolean,
			IsPriceStopEnabled As Boolean,
			IsInversePositionOnPriceStopEnabled As Boolean,
			TransactionCostPerCent As Double,
			GainLimiting As Double,
			ListOfPriceVol As IList(Of IPriceVol),
			ListOfPriceStopFromStochastic As IList(Of Double),
			ListOfPriceStochasticMedianDailyBandHigh As IList(Of Double),
			ListOfPriceStochasticMedianDailyBandLow As IList(Of Double)) Implements IStochasticPriceGain.Init

			Throw New NotImplementedException()
		End Sub

		Private IsFilterGainPriceStopOneSigmaEnabledLocal As Boolean
		Private IsSochasticPriceMedianIncludingGainLocal As Boolean
		Private IsPriceStopBoundToDailyOneSigmaEnabledLocal As Boolean

		Private Sub Init(
			FilterRateForGainMeasurement As Integer,
			IsGainFunctionWeightedMethod As Boolean,
			IsPriceStopEnabled As Boolean,
			IsInversePositionOnPriceStopEnabled As Boolean,
			TransactionCostPerCent As Double,
			GainLimiting As Double,
			IsFilterGainPriceStopOneSigmaEnabled As Boolean,
			IsSochasticPriceMedianIncludingGain As Boolean,
			IsPriceStopBoundToDailyOneSigmaEnabled As Boolean) Implements IStochasticPriceGain.Init

			IsFilterGainPriceStopOneSigmaEnabledLocal = IsFilterGainPriceStopOneSigmaEnabledLocal
			IsSochasticPriceMedianIncludingGainLocal = IsSochasticPriceMedianIncludingGain
			IsPriceStopBoundToDailyOneSigmaEnabledLocal = IsPriceStopBoundToDailyOneSigmaEnabled

			Dim ThisListOfPriceStopFromStochastic = Me.CalculateStochasticPriceStop(
				IsFilterGainPriceStopOneSigmaEnabled,
				IsSochasticPriceMedianIncludingGain,
				IsPriceStopBoundToDailyOneSigmaEnabled)

			'MyStochasticPriceGain = New StochasticPriceGain(
			'  FilterRate:=Me.FilterRate,
			'  FilterRateForGainMeasurement:=FilterRateForGainMeasurement,
			'  IsGainFunctionWeightedMethod:=IsGainFunctionWeightedMethod,
			'  IsPriceStopEnabled:=IsPriceStopEnabled,
			'  IsInversePositionOnPriceStopEnabled:=IsInversePositionOnPriceStopEnabled,
			'  TransactionCostPerCent:=TransactionCostPerCent,
			'  GainLimiting:=GainLimiting,
			'  ListOfPriceVol:=MyListOfValue,
			'  ListOfPriceStopFromStochastic:=ThisListOfPriceStopFromStochastic,
			'  ListOfPriceStochasticMedianDailyBandHigh:=Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyUp),
			'  ListOfPriceStochasticMedianDailyBandLow:=Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticMedianRangeDailyDown))


			MyStochasticPriceGain = New StochasticPriceGain(
				FilterRate:=Me.FilterRate,
				FilterRateForGainMeasurement:=FilterRateForGainMeasurement,
				IsGainFunctionWeightedMethod:=IsGainFunctionWeightedMethod,
				IsPriceStopEnabled:=IsPriceStopEnabled,
				IsInversePositionOnPriceStopEnabled:=IsInversePositionOnPriceStopEnabled,
				TransactionCostPerCent:=TransactionCostPerCent,
				GainLimiting:=GainLimiting,
				ListOfPriceVol:=MyListOfValue,
				ListOfStochasticProbability:=Me.ToList,
				ThresholdLevel:=THRESHOLD_LEVEL)
		End Sub

		Public Sub Init(
			FilterRate As Integer,
			FilterGainMeasurementPeriod As Integer,
			IsGainFunctionWeightedMethod As Boolean,
			IsPriceStopEnabled As Boolean,
			IsInversePositionOnPriceStopEnabled As Boolean,
			TransactionCostPerCent As Double,
			GainLimiting As Double,
			ListOfPriceVol As IList(Of IPriceVol),
			ListOfStochasticProbability As IList(Of Double),
			ThresholdLevel As Double) Implements IStochasticPriceGain.Init

			IsFilterGainPriceStopOneSigmaEnabledLocal = IsFilterGainPriceStopOneSigmaEnabledLocal
			IsSochasticPriceMedianIncludingGainLocal = IsSochasticPriceMedianIncludingGain
			IsPriceStopBoundToDailyOneSigmaEnabledLocal = IsPriceStopBoundToDailyOneSigmaEnabled

			'Dim ThisListOfPriceStopFromStochastic = Me.CalculateStochasticPriceStop(
			'  IsFilter GainPriceStopOneSigmaEnabled,
			'  IsSochasticPriceMedianIncludingGain,
			'  IsPriceStopBoundToDailyOneSigmaEnabled)

			MyStochasticPriceGain = New StochasticPriceGain(
				FilterRate:=Me.FilterRate,
				FilterRateForGainMeasurement:=FilterGainMeasurementPeriod,
				IsGainFunctionWeightedMethod:=IsGainFunctionWeightedMethod,
				IsPriceStopEnabled:=IsPriceStopEnabled,
				IsInversePositionOnPriceStopEnabled:=IsInversePositionOnPriceStopEnabled,
				TransactionCostPerCent:=TransactionCostPerCent,
				GainLimiting:=GainLimiting,
				ListOfPriceVol:=MyListOfValue,
				ListOfStochasticProbability:=ListOfStochasticProbability,
				ThresholdLevel:=ThresholdLevel)
		End Sub

		Public Sub Init(
			FilterGainMeasurementPeriod As Integer,
			IsGainFunctionWeightedMethod As Boolean,
			IsPriceStopEnabled As Boolean,
			IsInversePositionOnPriceStopEnabled As Boolean,
			TransactionCostPerCent As Double,
			GainLimiting As Double,
			IsFilterGainPriceStopOneSigmaEnabled As Boolean,
			IsStochasticPriceMedianIncludingGain As Boolean,
			IsPriceStopBoundToDailyOneSigmaEnabled As Boolean,
			ThresholdLevel As Double) Implements IStochasticPriceGain.Init


			MyStochasticPriceGain = New StochasticPriceGain(
				FilterRate:=Me.FilterRate,
				FilterRateForGainMeasurement:=FilterGainMeasurementPeriod,
				IsGainFunctionWeightedMethod:=IsGainFunctionWeightedMethod,
				IsPriceStopEnabled:=IsPriceStopEnabled,
				IsInversePositionOnPriceStopEnabled:=IsInversePositionOnPriceStopEnabled,
				TransactionCostPerCent:=TransactionCostPerCent,
				GainLimiting:=GainLimiting,
				ListOfPriceVol:=MyListOfValue,
				ListOfStochasticProbability:=Me.ToList,
				ThresholdLevel:=ThresholdLevel)

			'MyStochasticPriceGain = New StochasticPriceGain(
			'  FilterRate:=Me.FilterRate,
			'  FilterRateForGainMeasurement:=FilterGainMeasurementPeriod,
			'  IsGainFunctionWeightedMethod:=IsGainFunctionWeightedMethod,
			'  IsPriceStopEnabled:=IsPriceStopEnabled,
			'  IsInversePositionOnPriceStopEnabled:=IsInversePositionOnPriceStopEnabled,
			'  TransactionCostPerCent:=TransactionCostPerCent,
			'  GainLimiting:=GainLimiting,
			'  ListOfPriceVol:=MyListOfValue,
			'  ListOfStochasticProbability:=Me.ToList(Type:=IStochastic.enuStochasticType.PriceStochacticVolatilityPositiveToNegativeRatio),
			'  ThresholdLevel:=ThresholdLevel)

		End Sub

		Public ReadOnly Property AsIStochasticPriceGain As IStochasticPriceGain Implements IStochasticPriceGain.AsIStochasticPriceGain
			Get
				Return Me
			End Get
		End Property

		Private ReadOnly Property FilterRate As Integer Implements IStochasticPriceGain.FilterRate
			Get
				Return Me.Rate
			End Get
		End Property
		Private ReadOnly Property FilterRateForGain As Integer Implements IStochasticPriceGain.FilterRateForGain
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.FilterRateForGain
				Else
					Return Me.Rate
				End If
			End Get
		End Property

		Private ReadOnly Property IsGainFunctionWeightedMethod As Boolean Implements IStochasticPriceGain.IsGainFunctionWeightedMethod
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.IsGainFunctionWeightedMethod
				Else
					Return False
				End If
			End Get
		End Property

		Private ReadOnly Property IsPriceStopEnabled As Boolean Implements IStochasticPriceGain.IsPriceStopEnabled
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.IsPriceStopEnabled
				Else
					Return False
				End If
			End Get
		End Property

		Private ReadOnly Property IsInversePositionOnPriceStopEnabled As Boolean Implements IStochasticPriceGain.IsInversePositionOnPriceStopEnabled
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.IsInversePositionOnPriceStopEnabled
				Else
					Return False
				End If
			End Get
		End Property

		Private ReadOnly Property TransactionCostPerCent As Double Implements IStochasticPriceGain.TransactionCostPerCent
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.TransactionCostPerCent
				Else
					Return 0.0
				End If
			End Get
		End Property

		Private ReadOnly Property GainLimiting As Double Implements IStochasticPriceGain.GainLimiting
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.GainLimiting
				Else
					Return 1.0
				End If
			End Get
		End Property

		Private ReadOnly Property FilterTransactionGainLog As FilterTransactionGainLog Implements IStochasticPriceGain.FilterTransactionGainLog
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.FilterTransactionGainLog
				Else
					Return Nothing
				End If
			End Get
		End Property

		Private ReadOnly Property FilterTransactionGainLogFast As FilterTransactionGainLog Implements IStochasticPriceGain.FilterTransactionGainLogFast
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return MyStochasticPriceGain.FilterTransactionGainLogFast
				Else
					Return Nothing
				End If
			End Get
		End Property

		Private ReadOnly Property IsFilterGainPriceStopOneSigmaEnabled As Boolean Implements IStochasticPriceGain.IsFilterGainPriceStopOneSigmaEnabled
			Get
				Return IsFilterGainPriceStopOneSigmaEnabledLocal
			End Get
		End Property

		Private ReadOnly Property IsSochasticPriceMedianIncludingGain As Boolean Implements IStochasticPriceGain.IsStochasticPriceMedianIncludingGain
			Get
				Return IsSochasticPriceMedianIncludingGainLocal
			End Get
		End Property

		Private ReadOnly Property IsPriceStopBoundToDailyOneSigmaEnabled As Boolean Implements IStochasticPriceGain.IsPriceStopBoundToDailyOneSigmaEnabled
			Get
				Return IsPriceStopBoundToDailyOneSigmaEnabledLocal
			End Get
		End Property

		Private ReadOnly Property IStochasticPriceGain_AsIStochastic1 As IStochastic1 Implements IStochasticPriceGain.AsIStochastic1
			Get
				Return Me
			End Get
		End Property

		Public ReadOnly Property AsIStochastic2 As IStochastic2 Implements IStochastic1.AsIStochastic2
			Get
				Return Me
			End Get
		End Property
		Private ReadOnly Property IStochasticPriceGain_AsIStochastic As IStochastic Implements IStochasticPriceGain.AsIStochastic
			Get
				Return Me
			End Get
		End Property

		Private ReadOnly Property IStochastic1_AsIStochasticPriceGain As IStochasticPriceGain Implements IStochastic1.AsIStochasticPriceGain
			Get
				Return Me
			End Get
		End Property

		Public ReadOnly Property AsIStochastic1 As IStochastic1 Implements IStochastic1.AsIStochastic1
			Get
				Return Me
			End Get
		End Property

		Public ReadOnly Property AsIStochastic As IStochastic Implements IStochastic1.AsIStochastic
			Get
				Return Me
			End Get
		End Property

		Private Property IStochastic2_Symbol As String Implements IStochastic2.Symbol
			Get
				Return Me.Symbol
			End Get
			Set(value As String)
				Me.Symbol = value
			End Set
		End Property

		Private ReadOnly Property IsInit As Boolean Implements IStochasticPriceGain.IsInit
			Get
				If MyStochasticPriceGain IsNot Nothing Then
					Return True
				Else
					Return False
				End If
			End Get
		End Property

		Private ReadOnly Property IStochasticPriceGain_ToList(GainType As IStochasticPriceGain.EnuGainType) As IList(Of Double) Implements IStochasticPriceGain.ToList
			Get
				Return MyStochasticPriceGain.ToList(GainType:=GainType)
			End Get
		End Property

		Private ReadOnly Property IStochasticPriceGain_ToList(GainType As IStochasticPriceGain.EnuGainType, IsFromFastFilter As Boolean) As IList(Of Double) Implements IStochasticPriceGain.ToList
			Get
				Return MyStochasticPriceGain.ToList(GainType:=GainType, IsFromFastFilter:=IsFromFastFilter)
			End Get
		End Property
#End Region
	End Class
End Namespace