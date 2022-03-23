Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions

Namespace MathPlus.Filter
  Friend Class FilterPLLDetectorForVolatilitySigma
    Implements IFilterPLLDetector

    Public Const RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE As Double = 1
    Public Const BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED As Integer = 20

    Private MyRate As Integer
    Private MyRateForSigmaStatisticDaily As Integer
    Private MyCount As Integer
    Private MyErrorLast As Double
    Private MyToCountLimit As Integer
    Private MyToCountLimitSelected As Integer
    Private MyToErrorLimit As Double
    Private MyValueInit As Double
    Private MyValueOutput As Double
    Private MyQueueForSigmaStatisticDaily As Queue(Of IStockPriceVolatilityPredictionBand)
    Private MySumForSigmaStatisticDaily As Double
    Private MyCountOfPLLRun As Integer
    Private MyStatus As Boolean
    Private MyVolatilityAverage As Double
    Private MyFilterPLL As FilterLowPassPLL
    Private MyDetectorBandExcessBalanceSum As Double
    Private MyDetectorBalanceLast As Double
    Private MyMaximum As Double
    Private MyListOfConvergence As IList(Of Double)
    Private MyProbabilityOfExcessMeasuredLast As Double
    Private MyListOfProbabilityOfExcess As IList(Of Double)
    Private MyListOfProbabilityOfExcessBalance As IList(Of Double)
    Private MyFastAttackCountForBandExceededLow As Double
    Private MyFastAttackCountForBandExceededHigh As Double

    Public Sub New(ByVal Rate As Integer, Optional ByVal ToCountLimit As Integer = 1, Optional ToErrorLimit As Double = 0.001)
      MyRate = Rate
      'MyRateForSigmaStatisticDaily = CInt(FilterPLLDetectorForVolatilitySigma.VolatilityRate)
      Me.IsUseFeedbackRegulatedVolatilityFastAttackEvent = False    'by default
      MyRateForSigmaStatisticDaily = BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
      MyToCountLimit = ToCountLimit
      MyToCountLimitSelected = ToCountLimit
      MyToErrorLimit = ToErrorLimit
      MyCount = 0
      MyValueInit = 0
      MyCountOfPLLRun = 0
      MyStatus = False
      MyQueueForSigmaStatisticDaily = New Queue(Of IStockPriceVolatilityPredictionBand)(capacity:=CInt(MyRateForSigmaStatisticDaily))
      MyFilterPLL = New FilterLowPassPLL(FilterRate:=7, DampingFactor:=1.0, NumberOfPredictionOutput:=0)
      MyListOfConvergence = New List(Of Double)
      MyListOfProbabilityOfExcess = New List(Of Double)
      MyListOfProbabilityOfExcessBalance = New List(Of Double)
      Me.Tag = TypeName(Me)
    End Sub

    Public Function RunErrorDetector(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.RunErrorDetector
      Dim ThisStockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand
      Dim IsBandExceededLast As Boolean
      Dim ThisProbOfBandExceedExpected As Double
      Dim ThisGradientSum As Double
      Dim ThisGradientMean As Double
      Dim ThisValueOuput As Double = Me.ValueOutput(Input, InputFeedback)
      Dim ThisVolatilityChangePerCent As Double
      Dim ThisCount As Integer
      Dim ThisWeight As Double
      Dim ThisWeightStep As Double
      Dim I As Integer


      If Me.IsUseFeedbackRegulatedVolatilityFastAttackEvent Then

      Else

      End If


      MyStatus = False
      MyErrorLast = 0   'by default
      'If MyCount Mod MyRateForSigmaStatisticDaily = 0 Then
      If MyCount Mod 1 = 0 Then
        'process for volatility correction
        MyCountOfPLLRun = MyCountOfPLLRun + 1
        Do
          MySumForSigmaStatisticDaily = 0
          MyDetectorBandExcessBalanceSum = 0
          ThisGradientSum = 0
          ThisCount = 0
          If Input <= 0 Then Exit Do
          If Me.Count = 500 Then
            I = I
          End If
          ThisVolatilityChangePerCent = InputFeedback / Input

          ThisProbOfBandExceedExpected = 1 - MyQueueForSigmaStatisticDaily(0).ProbabilityOfInterval
          'ThisWeight = 0.5
          'ThisWeightStep = 1 / MyQueueForSigmaStatisticDaily.Count
          'Dim A As Double = CDbl((2 / (5 + 1)))
          'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
          'Dim B As Double = 1 - A

          Dim ThisCountThresholdForFastAttack As Integer = CInt(0.8 * MyRateForSigmaStatisticDaily)
          For I = 0 To MyQueueForSigmaStatisticDaily.Count - 1
            ThisStockPriceVolatilityPredictionBand = MyQueueForSigmaStatisticDaily(I)
            'For Each ThisStockPriceVolatilityPredictionBand In MyQueueForSigmaStatisticDaily
            ThisWeight = ThisWeight + ThisWeightStep
            With ThisStockPriceVolatilityPredictionBand
              If .Volatility = 0 Then
                Exit Do
              Else
                'note that with the last data .IsStockPriceValueRealEnabled is always false because it does not have received the real data yet for comparaison
                'it only hold the expected high and low value
                .Refresh(VolatilityDelta:=ThisVolatilityChangePerCent * .Volatility)
                ThisCount = ThisCount + 1
                ThisGradientSum = ThisGradientSum + .RatioOfΔProbabilityToΔVolatility
                'If .IsStockPriceValueRealEnabled Then
                If .IsBandExceeded Then
                  If IsBandExceededLast Then
                    ''increase the weight for consecutive hit
                    ''this is to speed up the reaction time to fast change that require
                    ''a faster volatility increase
                    'If .IsBandExceededLow Then
                    '  'accelerate quickly on the way down
                    '  'MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + ThisWeight * MyFastAttackCountForBandExceededLow
                    '  MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + .VolatilityExcessRatio
                    'Else
                    '  'not so fast on the up side
                    '  'MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + ThisWeight * MyFastAttackCountForBandExceededHigh
                    '  MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + .VolatilityExcessRatio
                    'End If
                  Else
                    IsBandExceededLast = True
                    'MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + ThisWeight * 1
                  End If
                  If IsUseFeedbackRegulatedVolatilityFastAttackEventLocal Then
                    If I >= ThisCountThresholdForFastAttack Then
                      MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + .VolatilityExcessRatio
                      'MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + 1
                    Else
                      MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + 1
                    End If
                  Else
                    MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + 1
                  End If
                  If .IsBandExceededHigh Then
                    MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum + 1
                  End If
                  If .IsBandExceededLow Then
                    MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum - 1
                  End If
                Else
                  IsBandExceededLast = False
                End If
                'Else
                'IsBandExceededLast = False
                'End If
              End If
            End With
          Next
          'If Me.Count = 500 Then
          '  MyCount = MyCount
          'End If
          If ThisCount > 1 Then
            MyDetectorBalanceLast = MyDetectorBandExcessBalanceSum / ThisCount
            ThisGradientMean = ThisGradientSum / ThisCount
            If MySumForSigmaStatisticDaily > ThisCount Then
              MySumForSigmaStatisticDaily = ThisCount
            End If
            MyProbabilityOfExcessMeasuredLast = MySumForSigmaStatisticDaily / ThisCount
            MyErrorLast = (MyProbabilityOfExcessMeasuredLast - ThisProbOfBandExceedExpected)
            MyErrorLast = -MyErrorLast / ThisGradientMean
            'this error is equivalent to 0.5 sample of the buffer size
            'MyToErrorLimit = (1 / (MyQueueForSigmaStatisticDaily.Count - 1)) * ThisGradientMean / 4
            MyStatus = True
            Exit Do
          Else
            MyDetectorBalanceLast = 0
            ThisGradientMean = 0
            MyProbabilityOfExcessMeasuredLast = 0
            MyErrorLast = 0
            MyStatus = False
            Exit Do
          End If
        Loop
      End If
      Return MyErrorLast
    End Function

    Public Sub Update(ByVal StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand)
      Dim ThisQueueDataLast As IStockPriceVolatilityPredictionBand = Nothing
      Dim ThisQueueDataRemoved As IStockPriceVolatilityPredictionBand = Nothing
      Dim ThisVolatilityDelta As Double
      Dim ThisQueueDataLastDate As Date
      Dim ThisQueueDataRemovedDate As Date

      MyCountOfPLLRun = 0
      MyCount = MyCount + 1
      If MyQueueForSigmaStatisticDaily.Count > 0 Then
        'update the previous sample with the actual price
        'it can only be updated now because future data was not yet availaible for the previous sample
        ThisQueueDataLast = MyQueueForSigmaStatisticDaily.Last
        ThisQueueDataLastDate = ThisQueueDataLast.StockPrice.DateDay
        With ThisQueueDataLast
          .Refresh(ThisQueueDataLast.VolatilityDelta, StockPriceVolatilityPredictionBand.StockPrice)
          ''and refresh the current item with the previous sample estimate of the high and low
          ''this way the current sample can still estimate if it current value exceed the estimated threshold
          'Dim ThisPriceVol As IPriceVol = New PriceVol(CSng((.StockPriceHighValue + .StockPriceLowValue) / 2))
          'ThisPriceVol.High = CSng(.StockPriceHighValue)
          'ThisPriceVol.Low = CSng(.StockPriceLowValue)
          'StockPriceVolatilityPredictionBand.Refresh(ThisPriceVol)
        End With
      End If
      If MyQueueForSigmaStatisticDaily.Count = MyRateForSigmaStatisticDaily Then
        'remove from the old data and adjust the sum and the statisitc
        ThisQueueDataRemoved = MyQueueForSigmaStatisticDaily.Dequeue
        ThisQueueDataRemovedDate = ThisQueueDataRemoved.StockPrice.DateDay
      End If
      ''this is to try to limit the PLL convergence processing when there is no band exceeded
      ''not use anymore because the threading has speed up the process by a factor 10
      'If (ThisQueueDataLast IsNot Nothing) And (ThisQueueDataRemoved IsNot Nothing) Then
      '  If (ThisQueueDataLast.IsBandExceeded) Or (ThisQueueDataRemoved.IsBandExceeded) Then
      '    MyToCountLimitSelected = MyToCountLimit
      '  Else
      '    'this help a lot to limit the calculation time
      '    'MyToCountLimitSelected = MyFilterPLL.Rate
      '    'do not used for now
      '    MyToCountLimitSelected = MyToCountLimit
      '  End If
      'Else
      '  MyToCountLimitSelected = MyToCountLimit
      'End If
      MyToCountLimitSelected = MyToCountLimit
      MyQueueForSigmaStatisticDaily.Enqueue(StockPriceVolatilityPredictionBand)
      'If MyCount >= 1576 Then
      '  MyCount = MyCount
      'End If

      'If ThisQueueDataLast Then
      'note the PLL return the error VolatilityDelta, however it finish the calculation with the Volatility Total
      'obtained by calling the object ValueOutput function. The filter PLL list contain the volatility total as the final result
      'and also the FilterLast value
      ThisVolatilityDelta = MyFilterPLL.Filter(StockPriceVolatilityPredictionBand.Volatility, Me)
      'call RunErrorDetector for the last statistic update
      Me.RunErrorDetector(StockPriceVolatilityPredictionBand.Volatility, ThisVolatilityDelta)
      StockPriceVolatilityPredictionBand.Refresh(VolatilityDelta:=ThisVolatilityDelta)


      MyListOfProbabilityOfExcess.Add(MyProbabilityOfExcessMeasuredLast)
      MyListOfProbabilityOfExcessBalance.Add(MyDetectorBalanceLast)
      MyToCountLimitSelected = MyToCountLimit
    End Sub

    Private IsUseFeedbackRegulatedVolatilityFastAttackEventLocal As Boolean
    Public Property IsUseFeedbackRegulatedVolatilityFastAttackEvent As Boolean
      Get
        Return IsUseFeedbackRegulatedVolatilityFastAttackEventLocal
      End Get
      Set(value As Boolean)
        IsUseFeedbackRegulatedVolatilityFastAttackEventLocal = value
        If IsUseFeedbackRegulatedVolatilityFastAttackEventLocal Then
          MyFastAttackCountForBandExceededHigh = 1.5
          MyFastAttackCountForBandExceededLow = 3
        Else
          MyFastAttackCountForBandExceededLow = 1
          MyFastAttackCountForBandExceededHigh = 1
        End If
      End Set
    End Property

    Public ReadOnly Property ToErrorLimit As Double Implements IFilterPLLDetector.ToErrorLimit
      Get
        Return MyToErrorLimit
      End Get
    End Property

    Public ReadOnly Property ToCount As Integer Implements IFilterPLLDetector.ToCount
      Get
        Return MyToCountLimitSelected
      End Get
    End Property

    Public ReadOnly Property Count As Double Implements IFilterPLLDetector.Count
      Get
        Return MyCount
      End Get
    End Property

    Public ReadOnly Property ErrorLast As Double Implements IFilterPLLDetector.ErrorLast
      Get
        Return MyErrorLast
      End Get
    End Property

    Public ReadOnly Property ValueInit As Double Implements IFilterPLLDetector.ValueInit
      Get
        Return MyValueInit
      End Get
    End Property

    Public Function ValueOutput(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.ValueOutput
      Return Input + InputFeedback
    End Function

    Public ReadOnly Property Status As Boolean Implements IFilterPLLDetector.Status
      Get
        Return MyStatus
      End Get
    End Property

    Public ReadOnly Property IsMaximum As Boolean Implements IFilterPLLDetector.IsMaximum
      Get
        Return False
      End Get
    End Property

    Public ReadOnly Property IsMinimum As Boolean Implements IFilterPLLDetector.IsMinimum
      Get
        Return False
      End Get
    End Property

    Public ReadOnly Property Maximum As Double Implements IFilterPLLDetector.Maximum
      Get
        Return MyMaximum
      End Get
    End Property

    Public ReadOnly Property Minimum As Double Implements IFilterPLLDetector.Minimum
      Get
        Return 0
      End Get
    End Property

    Public ReadOnly Property DetectorBalance As Double Implements IFilterPLLDetector.DetectorBalance
      Get
        Return MyDetectorBalanceLast
      End Get
    End Property

    Public ReadOnly Property ToListOfProbabilityOfExcess As IList(Of Double)
      Get
        Return MyListOfProbabilityOfExcess
      End Get
    End Property

    Public ReadOnly Property ToListOfProbabilityOfExcessBalance As IList(Of Double)
      Get
        Return MyListOfProbabilityOfExcessBalance
      End Get
    End Property

    Public ReadOnly Property ToList As IList(Of Double) Implements IFilterPLLDetector.ToList
      Get
        Return MyFilterPLL.ToList
      End Get
    End Property

    Public Shared Function VolatilityRate() As Double
      Return RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
    End Function

    Public Shared Function VolatilityRate(ByVal FilterRate As Double) As Double
      Return RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
    End Function

    Public Property Tag As String Implements IFilterPLLDetector.Tag

    Public ReadOnly Property ToListOfConvergence As IList(Of Double) Implements IFilterPLLDetector.ToListOfConvergence
      Get
        Return MyListOfConvergence
      End Get
    End Property

    Public Sub RunConvergence(NumberOfIteration As Integer, ValueBegin As Double, ValueEnd As Double) Implements IFilterPLLDetector.RunConvergence
      If ValueBegin <> 0 Then
        MyListOfConvergence.Add(Math.Abs((ValueEnd - ValueBegin)) / ValueBegin)
      Else
        MyListOfConvergence.Add(0)
      End If
    End Sub

    Private ReadOnly Property ToListOfVolatility As IList(Of Double) Implements IFilterPLLDetector.ToListOfVolatility
      Get

      End Get
    End Property

    Private ReadOnly Property ToListOfPriceMedianNextDayHigh As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayHigh
      Get

      End Get
    End Property

    Private ReadOnly Property ToListOfPriceMedianNextDayLow As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayLow
      Get

      End Get
    End Property
  End Class
End Namespace