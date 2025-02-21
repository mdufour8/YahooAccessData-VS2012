Imports YahooAccessData.MathPlus.Filter

Public Class FilterStochasticBrownianModulated
  Implements IStochastic

#Region "New"
  Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer, Optional ByVal IsFilterPeakEnabled As Boolean = False)
    Me.New(FilterRate, CDbl(FilterOutputRate))
  End Sub

  Public Sub New(ByVal IsPreFilter As Boolean, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double, Optional ByVal IsFilterPeakEnabled As Boolean = False)
    Me.New(FilterRate, FilterRate, FilterOutputRate)
  End Sub

  Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer, Optional ByVal IsFilterPeakEnabled As Boolean = False)
    Me.New(PreFilterRate, FilterRate, CDbl(FilterOutputRate))
  End Sub

  Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double, Optional ByVal IsFilterPeakEnabled As Boolean = False)
    'MyProcessorCount = Environment.ProcessorCount
    'If MyProcessorCount < 2 Then MyProcessorCount = 2

    'MyRateForVolatility = CInt(FilterPLLDetectorForVolatilitySigma.VolatilityRate)
    'MyFilterStochastic = New FilterStochastic(PreFilterRate, FilterRate, FilterOutputRate)
    'MyMeasurePeakValueRange = New MeasurePeakValueRange(FilterRate, Me.IsFilterPeak)
    'MyMeasurePeakValueRangeUsingNoPeakFilter = New MeasurePeakValueRange(FilterRate:=FilterRate, IsFilterPeakEnabled:=False)
    'MyListOfPeakValueGainPrediction = New List(Of Double)

    'MyLogNormalForVolatilityStatistic = New MathNet.Numerics.Distributions.LogNormal(0, 1)

    'Me.IsFilterPeak = IsFilterPeakEnabled
    'If FilterOutputRate < 1 Then FilterOutputRate = 1
    'MyQueueOfDailyRangeSigmaExcess = New Queue(Of Tuple(Of Integer, Integer, Integer))(capacity:=MyRateForVolatility)
    'MyFilterVolatilityYangZhangForStatistic = New FilterVolatilityYangZhang(MyRateForVolatility, FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=False)
    'MyFilterVolatilityYangZhangForStatisticLastPointTrail = New FilterVolatilityYangZhang(MyRateForVolatility, FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=True)
    'MyFilterLPForPrice = New FilterLowPassPLL(FilterRate, NumberOfPredictionOutput:=0)

    'MyFilterPLLForGain = New FilterLowPassPLL(FilterRate, IsPredictionEnabled:=True)
    'MyFilterPLLForGainPrediction = New FilterLowPassPLL(FilterRate:=FilterRate, NumberOfPredictionOutput:=1, IsPredictionEnabled:=True)

    'MyListOfProbabilityBandHigh = New List(Of Double)
    'MyListOfProbabilityDailySigmaExcess = New List(Of Double)
    'MyListOfPriceProbabilityMedian = New List(Of Double)
    'MyListOfProbabilityBandLow = New List(Of Double)
    'MyListOfPriceVolatilityGain = New List(Of Double)
    'MyListOfPriceVolatilityHigh = New List(Of Double)
    'MyListOfPriceVolatilityLow = New List(Of Double)
    'MyListOfPriceVolatilityTimeProbability = New List(Of Double)
    'MyStatisticalDistributionForVolatility = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
    'MyStatisticalDistributionForVolatilityPositive = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE \ 2, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
    'MyStatisticalDistributionForVolatilityNegative = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE \ 2, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))

    'MyStatisticalDistributionForPrice = New StatisticalDistribution(STATISTIC_VOLATILITY_WINDOWS_SIZE, New StatisticalDistributionFunctionLog(STATISTIC_DB_MINIMUM, STATISTIC_DB_MAXIMUM))
    'MyFilterForVolatilityStatistic = New Filter.FilterStatistical(STATISTIC_VOLATILITY_WINDOWS_SIZE)
    'MyFilterLPForProbabilityFromBandVolatility = New FilterLowPassExp(FilterOutputRate)
    'MyFilterLPForStochasticFromPriceVolatilityHigh = New FilterLowPassExp(FilterOutputRate)
    'MyFilterLPForStochasticFromPriceVolatilityLow = New FilterLowPassExp(FilterOutputRate)
    'MyPLLErrorDetectorForPriceStochacticMedian = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGain = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGainPrediction = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionHigh = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionLow = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGainNoFilter = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForPriceStochacticMedianWithGainPredictionNoFilter = New FilterPLLDetectorForCDFToZero(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)

    ''MyPLLErrorDetectorForForVolatilitySigma = New FilterPLLDetectorForVolatilitySigma(FilterRate, ToCountLimit:=20, ToErrorLimit:=0.0001)
    'MyPLLErrorDetectorForForVolatilitySigma = New FilterPLLDetectorForVolatilitySigma(
    '  FilterRate,
    '  ToCountLimit:=FILTER_PLL_DETECTOR_COUNT_LIMIT,
    '  ToErrorLimit:=FILTER_PLL_DETECTOR_ERROR_LIMIT)

    'MyListForVolatilityRegulated = New List(Of Double)
    ''predict the next sample using the last 5 samples
    'MyFilterPLLForVolatilityRegulatedPrediction = New FilterLowPassPLL(FilterRate:=5, NumberOfPredictionOutput:=1)
    ''MyFilterPLLForVolatilityRegulatedPrediction = New FilterLowPassPLL(FilterRate:=Rate)
    'MyListForVolatilityDetectorBalance = New List(Of Double)

    'MyListOfPriceNextDailyHigh = New List(Of Double)
    'MyListOfPriceNextDailyLow = New List(Of Double)
    'MyListOfPriceNextDailyHighWithGain = New List(Of Double)
    'MyListOfPriceNextDailyLowWithGain = New List(Of Double)
  End Sub

  Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Double, Optional ByVal IsFilterPeakEnabled As Boolean = False)
    Me.New(FilterRate, FilterRate, FilterOutputRate)
  End Sub
#End Region


  Public ReadOnly Property Count As Integer Implements IStochastic.Count
    Get
      Throw New NotImplementedException
    End Get
  End Property

  Public Function Filter(ByRef Value As Double) As Double Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function Filter(ByRef Value() As Double) As Double() Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function Filter(Value As Single) As Double Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function Filter(ByRef Value As IPriceVol) As Double Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function Filter(ByRef Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function Filter(ByRef Value As IPriceVol, FilterRate As Integer) As Double Implements IStochastic.Filter
    Throw New NotImplementedException
  End Function

  Public Function FilterBackTo(ByRef Value As Double, Optional IsPreFilter As Boolean = True) As Double Implements IStochastic.FilterBackTo
    Throw New NotImplementedException
  End Function

  Public Function FilterLast() As Double Implements IStochastic.FilterLast
    Throw New NotImplementedException
  End Function

  Public Function FilterLast(Type As IStochastic.enuStochasticType) As Double Implements IStochastic.FilterLast
    Throw New NotImplementedException
  End Function

  Public Function FilterPredictionNext(ByRef Value As Double) As Double Implements IStochastic.FilterPredictionNext
    Throw New NotImplementedException
  End Function

  Public Function FilterPriceBandHigh() As Double Implements IStochastic.FilterPriceBandHigh
    Throw New NotImplementedException
  End Function

  Public Function FilterPriceBandLow() As Double Implements IStochastic.FilterPriceBandLow
    Throw New NotImplementedException
  End Function

  Public Property IsFilterPeak As Boolean Implements IStochastic.IsFilterPeak

  Public Property IsFilterRange As Boolean Implements IStochastic.IsFilterRange

  Public Function Last() As Double Implements IStochastic.Last
    Throw New NotImplementedException
  End Function

  Public ReadOnly Property Max(Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Max
    Get
      Throw New NotImplementedException
    End Get
  End Property

  Public ReadOnly Property Min(Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Min
    Get
      Throw New NotImplementedException
    End Get
  End Property

  Public Property Rate(Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Integer Implements IStochastic.Rate
    Get
      Throw New NotImplementedException
    End Get
    Set(value As Integer)

    End Set
  End Property

  Public Property Tag As String Implements IStochastic.Tag

  Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double, Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
    Throw New NotImplementedException
  End Function

  Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double, Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
    Throw New NotImplementedException
  End Function

  Public Function ToArray(Optional Type As IStochastic.enuStochasticType = MathPlus.Filter.IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
    Throw New NotImplementedException
  End Function

  Public ReadOnly Property ToList As IList(Of Double) Implements IStochastic.ToList
    Get
      Throw New NotImplementedException
    End Get
  End Property

  Public ReadOnly Property ToList(Type As IStochastic.enuStochasticType) As IList(Of Double) Implements IStochastic.ToList
    Get
      Throw New NotImplementedException
    End Get
  End Property

  Public Function ToString1() As String Implements IStochastic.ToString
    Throw New NotImplementedException
  End Function
End Class
