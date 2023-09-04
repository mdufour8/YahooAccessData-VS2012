Namespace MathPlus.Filter
  Public Interface IStochastic
    Enum enuStochasticType
      FastSlow
      Fast
      Slow
      PriceBandHigh
      PriceBandLow
      PriceBandHighPrediction
      PriceBandLowPrediction
      RangeVolatility
      ProbabilityHigh
      ProbabilityLow
      PriceBandVolatilityHigh
      PriceBandVolatilityLow
      PriceBandVolatilityGain
      PriceStandardVolatility
      PriceVolatilityPDF
      PriceVolatilityLCR
      PriceVolatilityPDFSimulated
      PriceVolatilityLCRSimulated
      PriceProbabilityMedian
      ProbabilityFromBandVolatility
      TimeProbabilityOfPriceVolatility
      StochasticSlowFromPriceBandVolatilityLow
      StochasticSlowFromPriceBandVolatilityHigh
      StochasticSlowFromPricePeakMedian
      PriceStochacticMedian
      PriceStochacticMedianWithGain
      PriceStochacticMedianRangeDailyUp
      PriceStochacticMedianWithGainRangeDailyUp
      PriceStochacticMedianRangeDailyDown
      PriceStochacticMedianWithGainRangeDailyDown
      ProbabilityPriceDailySigmaExceeded
      PriceVolatilityRegulated
      PriceVolatilityLastPointTrail
      PriceVolatilityDetectorBalance
      PriceStochacticMedianWithGainPrediction
      PriceStochacticMedianWithGainHisteresisHigh  'not supported
      PriceStochacticMedianWithGainHisteresisLow   'not supported
      PriceVolatilityRegulatedPrediction
      PricePeakValueGainPrediction
      PriceStochasticMedianVolatility
      PriceStochasticMedianNextDayLow
      PriceStochasticMedianNextDayHigh
      PriceStochacticMedianWithGainPredictionLow
      PriceStochacticMedianWithGainPredictionHigh
      ProbabilityPriceDailySigmaDoubleExceeded
      PriceStochacticMedianWithGainRangeDailyUpDay2
      PriceStochacticMedianWithGainRangeDailyDownDay2
      PriceStochacticMedianWithGainRangeDailyUpToOpen
      PriceStochacticMedianWithGainRangeDailyDownToOpen
      RangeVolatilityFromPreviousCloseToOpen
      RangeVolatilityFromPreviousCloseToOpenRatio
      RangeVolatilityFromOpenToClose
      PriceStochacticMedianWithGainRangeDailyUpFromOpenToClose
      PriceStochacticMedianWithGainRangeDailyDownFromOpenToClose
      ProbabilityOfPriceStochacticMedianWithGain
      PriceGainPerYear
      PriceGainPerYearDerivative
      PriceStochacticMedianWithGainRangeDailyUpAtSigma2
      PriceStochacticMedianWithGainRangeDailyLowAtSigma2
      PriceStochacticMedianWithGainRangeDailyUpAtSigma3
      PriceStochacticMedianWithGainRangeDailyLowAtSigma3
      RangeVolatilityRegulatedFromPreviousCloseToOpen
      PriceStochacticVolatilityPositive
      PriceStochacticVolatilityNegative
      PriceStochacticVolatilityPositiveToNegativeRatio
      PriceStochacticVolatilityPositiveToNegativeRatioFiltered
    End Enum

    Function Filter(ByVal Value As Single) As Double
    Function Filter(ByRef Value As Double) As Double
    Function Filter(ByRef Value As IPriceVol) As Double
    Function Filter(ByRef Value As IPriceVol, ByVal ValueExpectedMin As Double, ByVal ValueExpectedMax As Double) As Double
    Function Filter(ByRef Value As IPriceVol, ByVal FilterRate As Integer) As Double
    Function Filter(ByRef Value() As Double) As Double()
    Function FilterPredictionNext(ByRef Value As Double) As Double
    Function FilterBackTo(ByRef Value As Double, Optional ByVal IsPreFilter As Boolean = True) As Double
    Function FilterPriceBandHigh() As Double
    Function FilterPriceBandLow() As Double
    Function FilterLast() As Double
    Function FilterLast(ByVal Type As enuStochasticType) As Double
    Function Last() As Double
    Property Rate(Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Integer
    Property IsFilterPeak As Boolean
    Property IsFilterRange As Boolean
    ReadOnly Property Count As Integer
    ReadOnly Property Max(Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Double
    ReadOnly Property Min(Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Double
    ReadOnly Property ToList() As IList(Of Double)
    ReadOnly Property ToList(ByVal Type As enuStochasticType) As IList(Of Double)
    Function ToArray(Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Double()
    Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Double()
    Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As enuStochasticType = enuStochasticType.FastSlow) As Double()
    Property Tag As String
    Function ToString() As String
  End Interface


  Public Interface IStochastic1
    Inherits IStochastic
    ReadOnly Property AsIStochasticPriceGain As IStochasticPriceGain
    ReadOnly Property AsIStochastic1 As IStochastic1

    ReadOnly Property AsIStochastic As IStochastic
    ReadOnly Property AsIStochastic2 As IStochastic2
  End Interface

  Public Interface IStochastic2
    Inherits IStochastic1

    Property Symbol As String

  End Interface
End Namespace