Imports YahooAccessData.MathPlus.Filter
Public Interface IStochasticPriceGain
  Enum EnuGainType
    Total
    TotalScaled
    GainMonthlyPerYear
    GainMonthlyPerYearSquared
    GainAverage
    GainTransactionPerformace
  End Enum

  ReadOnly Property AsIStochasticPriceGain As IStochasticPriceGain
  ReadOnly Property AsIStochastic1 As IStochastic1

  ReadOnly Property AsIStochastic As IStochastic

  Sub Init(
    ByVal FilterRate As Integer,
    ByVal FilterGainMeasurementPeriod As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal ListOfPriceVol As IList(Of IPriceVol),
    ByVal ListOfPriceStopFromStochastic As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandHigh As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandLow As IList(Of Double))

  Sub Init(
    ByVal FilterGainMeasurementPeriod As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal IsFilterGainPriceStopOneSigmaEnabled As Boolean,
    ByVal IsStochasticPriceMedianIncludingGain As Boolean,
    ByVal IsPriceStopBoundToDailyOneSigmaEnabled As Boolean)

  Sub Init(
    ByVal FilterGainMeasurementPeriod As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal IsFilterGainPriceStopOneSigmaEnabled As Boolean,
    ByVal IsStochasticPriceMedianIncludingGain As Boolean,
    ByVal IsPriceStopBoundToDailyOneSigmaEnabled As Boolean,
    ByVal ThresholdLevel As Double)

  Sub Init(
    ByVal FilterRate As Integer,
    ByVal FilterGainMeasurementPeriod As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal ListOfPriceVol As IList(Of IPriceVol),
    ByVal ListOfStochasticProbability As IList(Of Double),
    ByVal ThresholdLevel As Double)

  ReadOnly Property IsInit As Boolean
  ReadOnly Property FilterRate As Integer
  ReadOnly Property FilterRateForGain As Integer

  ReadOnly Property IsGainFunctionWeightedMethod As Boolean

  ReadOnly Property IsPriceStopEnabled As Boolean

  ReadOnly Property IsInversePositionOnPriceStopEnabled As Boolean

  ReadOnly Property TransactionCostPerCent As Double

  ReadOnly Property GainLimiting As Double

  ReadOnly Property IsFilterGainPriceStopOneSigmaEnabled As Boolean
  ReadOnly Property IsStochasticPriceMedianIncludingGain As Boolean
  ReadOnly Property IsPriceStopBoundToDailyOneSigmaEnabled As Boolean

  ReadOnly Property FilterTransactionGainLog As FilterTransactionGainLog
  ReadOnly Property FilterTransactionGainLogFast As FilterTransactionGainLog

  ReadOnly Property ToList(ByVal GainType As EnuGainType) As IList(Of Double)

  ReadOnly Property ToList(ByVal GainType As EnuGainType, ByVal IsFromFastFilter As Boolean) As IList(Of Double)
End Interface
