Imports System.Threading.Tasks
Imports YahooAccessData.MathPlus.Filter

Public Class StochasticPriceGain
  Implements IStochasticPriceGain

  Private Const GAIN_MEASUREMENT_PERIOD_RATIO_TO_FILTER_RATE_FAST As Double = 2 / 3

  Private MyFilterRate As Integer
  Private MyFilterRateForGainMeasurement As Integer
  Private IsGainFunctionWeightedMethodLocal As Boolean
  Private IsPriceStopEnabledLocal As Boolean
  Private IsInversePositionOnPriceStopEnabledLocal As Boolean
  Private MyTransactionCost As Double
  Private MyGainLimiting As Double
  Private MyListOfPriceVol As IList(Of IPriceVol)
  Private MyListOfPriceStopFromStochastic As IList(Of Double)
  Private MyListOfPriceStochasticMedianDailyBandHigh As IList(Of Double)
  Private MyListOfPriceStochasticMedianDailyBandLow As IList(Of Double)

  Private MyMeasureOfPriceGainLog As FilterTransactionGainLog
  Private MyMeasureOfPriceGainLogFast As FilterTransactionGainLog

  Public Sub New(
    ByVal FilterRate As Integer,
    ByVal FilterRateForGainMeasurement As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal ReportPrices As YahooAccessData.RecordPrices,
    ByVal ListOfPriceStopFromStochastic As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandHigh As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandLow As IList(Of Double))

    MyListOfPriceVol = New List(Of IPriceVol)
    For Each ThisPriceVol As IPriceVol In ReportPrices.PriceVolsData
      MyListOfPriceVol.Add(ThisPriceVol)
    Next
    Me.Init(
      FilterRate,
      FilterRateForGainMeasurement,
      IsGainFunctionWeightedMethod,
      IsPriceStopEnabled,
      IsInversePositionOnPriceStopEnabled,
      TransactionCostPerCent,
      GainLimiting,
      MyListOfPriceVol,
      ListOfPriceStopFromStochastic,
      ListOfPriceStochasticMedianDailyBandHigh,
      ListOfPriceStochasticMedianDailyBandLow)
  End Sub

  Public Sub New(
    ByVal FilterRate As Integer,
    ByVal FilterRateForGainMeasurement As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal ListOfPriceVol As IList(Of IPriceVol),
    ByVal ListOfPriceStopFromStochastic As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandHigh As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandLow As IList(Of Double))

    Me.Init(
      FilterRate,
      FilterRateForGainMeasurement,
      IsGainFunctionWeightedMethod,
      IsPriceStopEnabled,
      IsInversePositionOnPriceStopEnabled,
      TransactionCostPerCent,
      GainLimiting,
      ListOfPriceVol,
      ListOfPriceStopFromStochastic,
      ListOfPriceStochasticMedianDailyBandHigh,
      ListOfPriceStochasticMedianDailyBandLow)
  End Sub

  Public Sub New(
    FilterRate As Integer,
    FilterRateForGainMeasurement As Integer,
    IsGainFunctionWeightedMethod As Boolean,
    IsPriceStopEnabled As Boolean,
    IsInversePositionOnPriceStopEnabled As Boolean,
    TransactionCostPerCent As Double,
    GainLimiting As Double,
    ListOfPriceVol As IList(Of IPriceVol),
    ListOfStochasticProbability As IList(Of Double),
    ThresholdLevel As Double)

    Me.Init(
      FilterRate,
      FilterRateForGainMeasurement,
      IsGainFunctionWeightedMethod,
      IsPriceStopEnabled,
      IsInversePositionOnPriceStopEnabled,
      TransactionCostPerCent,
      GainLimiting,
      ListOfPriceVol,
      ListOfStochasticProbability,
      ThresholdLevel)

  End Sub

  Private Sub Init(
    ByVal FilterRate As Integer,
    ByVal FilterRateForGainMeasurement As Integer,
    ByVal IsGainFunctionWeightedMethod As Boolean,
    ByVal IsPriceStopEnabled As Boolean,
    ByVal IsInversePositionOnPriceStopEnabled As Boolean,
    ByVal TransactionCostPerCent As Double,
    ByVal GainLimiting As Double,
    ByVal ListOfPriceVol As IList(Of IPriceVol),
    ByVal ListOfPriceStopFromStochastic As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandHigh As IList(Of Double),
    ByVal ListOfPriceStochasticMedianDailyBandLow As IList(Of Double)) Implements IStochasticPriceGain.Init

    Dim I As Integer
    Dim ThisGainWeight As Double = 1.0

    MyListOfPriceVol = ListOfPriceVol
    MyFilterRate = FilterRate
    MyFilterRateForGainMeasurement = FilterRateForGainMeasurement
    IsGainFunctionWeightedMethodLocal = IsGainFunctionWeightedMethod
    IsPriceStopEnabledLocal = IsPriceStopEnabled
    IsInversePositionOnPriceStopEnabledLocal = IsInversePositionOnPriceStopEnabled
    MyTransactionCost = TransactionCostPerCent
    MyGainLimiting = GainLimiting
    MyListOfPriceStopFromStochastic = ListOfPriceStopFromStochastic
    MyListOfPriceStochasticMedianDailyBandHigh = ListOfPriceStochasticMedianDailyBandHigh
    MyListOfPriceStochasticMedianDailyBandLow = ListOfPriceStochasticMedianDailyBandLow

    MyMeasureOfPriceGainLog = New FilterTransactionGainLog(
      FilterRate:=MyFilterRate,
      GainMeasurementPeriod:=MyFilterRateForGainMeasurement,
      TransactionCostPerCent:=MyTransactionCost) With {
          .GainSaturationForLimiting = MyGainLimiting,
          .IsPriceStopEnabled = IsPriceStopEnabledLocal,
          .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

    MyMeasureOfPriceGainLogFast = New FilterTransactionGainLog(
      FilterRate:=MyFilterRate,
      GainMeasurementPeriod:=CInt(GAIN_MEASUREMENT_PERIOD_RATIO_TO_FILTER_RATE_FAST * MyFilterRateForGainMeasurement),
      TransactionCostPerCent:=MyTransactionCost) With {
          .GainSaturationForLimiting = MyGainLimiting,
          .IsPriceStopEnabled = IsPriceStopEnabledLocal,
          .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

    If IsGainFunctionWeightedMethodLocal Then
      Dim ThisMeasureOfPriceGainLogPreCalcul As FilterTransactionGainLog
      Dim ThisMeasureOfPriceGainLogFastPreCalcul As FilterTransactionGainLog

      ThisMeasureOfPriceGainLogPreCalcul = New FilterTransactionGainLog(FilterRate:=MyFilterRate, GainMeasurementPeriod:=MyFilterRateForGainMeasurement, TransactionCostPerCent:=MyTransactionCost) With {
        .GainSaturationForLimiting = MyGainLimiting,
        .IsPriceStopEnabled = IsPriceStopEnabledLocal,
        .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

      ThisMeasureOfPriceGainLogFastPreCalcul = New FilterTransactionGainLog(FilterRate:=MyFilterRate, GainMeasurementPeriod:=MyFilterRateForGainMeasurement \ 2, TransactionCostPerCent:=MyTransactionCost) With {
        .Tag = "",
        .GainSaturationForLimiting = MyGainLimiting,
        .IsPriceStopEnabled = IsPriceStopEnabledLocal,
        .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

      For I = 0 To MyListOfPriceVol.Count - 1
        ThisMeasureOfPriceGainLogPreCalcul.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=1.0,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
        ThisMeasureOfPriceGainLogFastPreCalcul.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=1.0,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))

        With ThisMeasureOfPriceGainLogPreCalcul.AsIFilterPrediction
          ThisGainWeight = .ToListOfGainPerYear.Last - .ToListOfGainPerYearDerivative.Last + 1
        End With
        If ThisGainWeight < 0.3 Then
          ThisGainWeight = 0
        Else
          ThisGainWeight = 1.0
        End If
        MyMeasureOfPriceGainLog.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
        MyMeasureOfPriceGainLogFast.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
      Next
    Else
      For I = 0 To MyListOfPriceVol.Count - 1
        MyMeasureOfPriceGainLog.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
        MyMeasureOfPriceGainLogFast.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
      Next
    End If
  End Sub


  Public Sub Init(
    FilterRate As Integer,
    FilterRateForGainMeasurement As Integer,
    IsGainFunctionWeightedMethod As Boolean,
    IsPriceStopEnabled As Boolean,
    IsInversePositionOnPriceStopEnabled As Boolean,
    TransactionCostPerCent As Double,
    GainLimiting As Double,
    ListOfPriceVol As IList(Of IPriceVol),
    ListOfStochasticProbability As IList(Of Double),
    ThresholdLevel As Double) Implements IStochasticPriceGain.Init

    Dim I As Integer
    Dim ThisGainWeight As Double = 1.0
    Dim ThisThresholdLevelPlus As Double = ThresholdLevel
    Dim ThisThresholdLevelMinus As Double = 1 - ThisThresholdLevelPlus


    MyListOfPriceVol = ListOfPriceVol
    MyFilterRate = FilterRate
    MyFilterRateForGainMeasurement = FilterRateForGainMeasurement
    IsGainFunctionWeightedMethodLocal = IsGainFunctionWeightedMethod
    IsPriceStopEnabledLocal = IsPriceStopEnabled
    IsInversePositionOnPriceStopEnabledLocal = IsInversePositionOnPriceStopEnabled
    MyTransactionCost = TransactionCostPerCent
    MyGainLimiting = GainLimiting

    MyMeasureOfPriceGainLog = New FilterTransactionGainLog(
      FilterRate:=MyFilterRate,
      GainMeasurementPeriod:=MyFilterRateForGainMeasurement,
      TransactionCostPerCent:=MyTransactionCost) With {
          .GainSaturationForLimiting = MyGainLimiting,
          .IsPriceStopEnabled = IsPriceStopEnabledLocal,
          .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

    MyMeasureOfPriceGainLogFast = New FilterTransactionGainLog(
      FilterRate:=MyFilterRate,
      GainMeasurementPeriod:=CInt(GAIN_MEASUREMENT_PERIOD_RATIO_TO_FILTER_RATE_FAST * MyFilterRateForGainMeasurement),
      TransactionCostPerCent:=MyTransactionCost) With {
          .GainSaturationForLimiting = MyGainLimiting,
          .IsPriceStopEnabled = IsPriceStopEnabledLocal,
          .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

    If IsGainFunctionWeightedMethodLocal Then
      Dim ThisMeasureOfPriceGainLogPreCalcul As FilterTransactionGainLog
      Dim ThisMeasureOfPriceGainLogFastPreCalcul As FilterTransactionGainLog

      ThisMeasureOfPriceGainLogPreCalcul = New FilterTransactionGainLog(FilterRate:=MyFilterRate, GainMeasurementPeriod:=MyFilterRateForGainMeasurement, TransactionCostPerCent:=MyTransactionCost) With {
        .GainSaturationForLimiting = MyGainLimiting,
        .IsPriceStopEnabled = IsPriceStopEnabledLocal,
        .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

      ThisMeasureOfPriceGainLogFastPreCalcul = New FilterTransactionGainLog(FilterRate:=MyFilterRate, GainMeasurementPeriod:=MyFilterRateForGainMeasurement \ 2, TransactionCostPerCent:=MyTransactionCost) With {
        .Tag = "",
        .GainSaturationForLimiting = MyGainLimiting,
        .IsPriceStopEnabled = IsPriceStopEnabledLocal,
        .IsInversePositionOnPriceStopEnabled = IsInversePositionOnPriceStopEnabledLocal}

      For I = 0 To MyListOfPriceVol.Count - 1
        ThisMeasureOfPriceGainLogPreCalcul.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=1.0,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
        ThisMeasureOfPriceGainLogFastPreCalcul.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=1.0,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))

        With ThisMeasureOfPriceGainLogPreCalcul.AsIFilterPrediction
          ThisGainWeight = .ToListOfGainPerYear.Last - .ToListOfGainPerYearDerivative.Last + 1
        End With
        If ThisGainWeight < 0.3 Then
          ThisGainWeight = 0
        Else
          ThisGainWeight = 1.0
        End If
        MyMeasureOfPriceGainLog.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
        MyMeasureOfPriceGainLogFast.Filter1(
            MyListOfPriceVol(I),
            WeightControl:=ThisGainWeight,
            ValueTransactionStop:=MyListOfPriceStopFromStochastic(I + 1),
            PriceRangeHigh:=MyListOfPriceStochasticMedianDailyBandHigh(I + 1),
            PriceRangeLow:=MyListOfPriceStochasticMedianDailyBandLow(I + 1))
      Next
    Else
      For I = 0 To MyListOfPriceVol.Count - 1
        Select Case ListOfStochasticProbability(I)
          Case >= ThisThresholdLevelPlus
            MyMeasureOfPriceGainLog.Filter(MyListOfPriceVol(I), 1)
            MyMeasureOfPriceGainLogFast.Filter(MyListOfPriceVol(I), 1)
          Case < ThisThresholdLevelMinus
            MyMeasureOfPriceGainLog.Filter(MyListOfPriceVol(I), -1)
            MyMeasureOfPriceGainLogFast.Filter(MyListOfPriceVol(I), -1)
          Case Else
            MyMeasureOfPriceGainLog.Filter(MyListOfPriceVol(I), 0)
            MyMeasureOfPriceGainLogFast.Filter(MyListOfPriceVol(I), 0)
        End Select
      Next
    End If
  End Sub

  Private Sub Init(
    FilterGainMeasurementPeriod As Integer,
    IsGainFunctionWeightedMethod As Boolean,
    IsPriceStopEnabled As Boolean,
    IsInversePositionOnPriceStopEnabled As Boolean,
    TransactionCostPerCent As Double,
    GainLimiting As Double,
    IsFilterGainPriceStopOneSigmaEnabled As Boolean,
    IsStochasticPriceMedianIncludingGain As Boolean,
    IsPriceStopBoundToDailyOneSigmaEnabled As Boolean) Implements IStochasticPriceGain.Init

    Throw New NotImplementedException()
  End Sub

  Public ReadOnly Property FilterRate As Integer Implements IStochasticPriceGain.FilterRate
    Get
      Return MyFilterRate
    End Get
  End Property

  Public ReadOnly Property FilterRateForGain As Integer Implements IStochasticPriceGain.FilterRateForGain
    Get
      Return MyFilterRateForGainMeasurement
    End Get
  End Property

  Public ReadOnly Property IsGainFunctionWeightedMethod As Boolean Implements IStochasticPriceGain.IsGainFunctionWeightedMethod
    Get
      Return IsGainFunctionWeightedMethodLocal
    End Get
  End Property

  Public ReadOnly Property IsPriceStopEnabled As Boolean Implements IStochasticPriceGain.IsPriceStopEnabled
    Get
      Return IsPriceStopEnabledLocal
    End Get
  End Property

  Public ReadOnly Property IsInversePositionOnPriceStopEnabled As Boolean Implements IStochasticPriceGain.IsInversePositionOnPriceStopEnabled
    Get
      Return IsInversePositionOnPriceStopEnabledLocal
    End Get
  End Property

  Public ReadOnly Property TransactionCostPerCent As Double Implements IStochasticPriceGain.TransactionCostPerCent
    Get
      Return MyTransactionCost
    End Get
  End Property

  Public ReadOnly Property GainLimiting As Double Implements IStochasticPriceGain.GainLimiting
    Get
      Return MyGainLimiting
    End Get
  End Property

  Public ReadOnly Property FilterTransactionGainLog As FilterTransactionGainLog Implements IStochasticPriceGain.FilterTransactionGainLog
    Get
      Return MyMeasureOfPriceGainLog
    End Get
  End Property

  Public ReadOnly Property FilterTransactionGainLogFast As FilterTransactionGainLog Implements IStochasticPriceGain.FilterTransactionGainLogFast
    Get
      Return MyMeasureOfPriceGainLogFast
    End Get
  End Property

#Region "Private function"
  Private ReadOnly Property IsFilterGainPriceStopOneSigmaEnabled As Boolean Implements IStochasticPriceGain.IsFilterGainPriceStopOneSigmaEnabled
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Private ReadOnly Property IsStochasticPriceMedianIncludingGain As Boolean Implements IStochasticPriceGain.IsStochasticPriceMedianIncludingGain
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Private ReadOnly Property IsPriceStopBoundToDailyOneSigmaEnabled As Boolean Implements IStochasticPriceGain.IsPriceStopBoundToDailyOneSigmaEnabled
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Public ReadOnly Property AsIStochasticPriceGain As IStochasticPriceGain Implements IStochasticPriceGain.AsIStochasticPriceGain
    Get
      Return Me
    End Get
  End Property

  Private ReadOnly Property AsIStochastic1 As IStochastic1 Implements IStochasticPriceGain.AsIStochastic1
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Private ReadOnly Property AsIStochastic As IStochastic Implements IStochasticPriceGain.AsIStochastic
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Public ReadOnly Property IsInit As Boolean Implements IStochasticPriceGain.IsInit
    Get
      Return True
    End Get
  End Property

  Public ReadOnly Property ToList(GainType As IStochasticPriceGain.EnuGainType) As IList(Of Double) Implements IStochasticPriceGain.ToList
    Get
      Return GetList(MyMeasureOfPriceGainLog, GainType:=GainType)
    End Get
  End Property

  Public ReadOnly Property ToList(GainType As IStochasticPriceGain.EnuGainType, IsFromFastFilter As Boolean) As IList(Of Double) Implements IStochasticPriceGain.ToList
    Get
      Return GetList(MyMeasureOfPriceGainLogFast, GainType:=GainType)
    End Get
  End Property

  Private Function GetList(MeasureOfPriceGainLog As FilterTransactionGainLog, GainType As IStochasticPriceGain.EnuGainType) As IList(Of Double)
    Select Case GainType
      Case IStochasticPriceGain.EnuGainType.Total
        Return MeasureOfPriceGainLog.ToList
      Case IStochasticPriceGain.EnuGainType.TotalScaled
        Return MeasureOfPriceGainLog.ToListScaled
      Case IStochasticPriceGain.EnuGainType.GainMonthlyPerYear
        Return MeasureOfPriceGainLog.AsIFilterPrediction.ToListOfGainPerYear
      Case IStochasticPriceGain.EnuGainType.GainMonthlyPerYearSquared
        Return MeasureOfPriceGainLog.ToListOfGainPerYearRMS
      Case IStochasticPriceGain.EnuGainType.GainAverage
        Return MeasureOfPriceGainLog.ToListOfGainPerYearAverage
      Case IStochasticPriceGain.EnuGainType.GainTransactionPerformace
        Return MeasureOfPriceGainLog.ToListOfGainTransactionPerformance
      Case Else
        Return MeasureOfPriceGainLog.ToList
    End Select
  End Function

  Public Sub Init(FilterGainMeasurementPeriod As Integer, IsGainFunctionWeightedMethod As Boolean, IsPriceStopEnabled As Boolean, IsInversePositionOnPriceStopEnabled As Boolean, TransactionCostPerCent As Double, GainLimiting As Double, IsFilterGainPriceStopOneSigmaEnabled As Boolean, IsStochasticPriceMedianIncludingGain As Boolean, IsPriceStopBoundToDailyOneSigmaEnabled As Boolean, ThresholdLevel As Double) Implements IStochasticPriceGain.Init
    Throw New NotImplementedException()
  End Sub
#End Region
End Class


