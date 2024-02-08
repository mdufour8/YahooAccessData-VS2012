Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks

Namespace MathPlus.Filter
  Friend Class FilterPLLDetectorForVolatilitySigmaAsync
    Implements IFilterPLLDetector

    Private MyRate As Integer
    Private MyRateForSigmaStatisticDaily As Integer
    Private MyCount As Integer
    Private MyErrorLast As Double
    Private MyToCountLimit As Integer
    Private MyToCountLimitSelected As Integer
    Private MyToErrorLimit As Double
    Private MyValueInit As Double
    Private MyValueOutput As Double
    Private MySumForSigmaStatisticDaily As Double
    Private MyCountOfPLLRun As Integer
    Private MyStatus As Boolean
    Private MyVolatilityAverage As Double
    Private MyFilterPLL As FilterLowPassPLL
    Private MyDetectorBandExcessBalanceSum As Double
    Private MyDetectorBalanceLast As Double
    Private MyMaximum As Double
    Private MyListOfConvergence As IList(Of Double)
    Private MyStartPoint As Integer
    Private MyStopPoint As Integer
    Private MyMapToLocalRunPoint As Integer
    Private MyMapToLocalStartPoint As Integer
    Private MyMapToLocalStopPoint As Integer
    Private MyProbOfBandExceedExpected As Double
    Private MyProbabilityOfExcessMeasuredLast As Double
    Private MyListOfProbabilityOfExcess As IList(Of Double)
    Private MyListOfProbabilityOfExcessBalance As IList(Of Double)
    Private MyFastAttackCountForBandExceededLow As Double
    Private MyFastAttackCountForBandExceededHigh As Double
    Private MyVolatilityPredictionBandLocal() As StockPriceVolatilityPredictionBand

    Public Sub New(
                  ByRef VolatilityPredictionBandArray() As StockPriceVolatilityPredictionBand,
                  ByVal StartPoint As Integer,
                  ByVal StopPoint As Integer,
                  ByVal Rate As Integer,
                  Optional ByVal ToCountLimit As Integer = 1,
                  Optional ToErrorLimit As Double = 0.001)

      Dim ThisStartPointInInputArray As Integer
      Dim ThisStopPointInInputArray As Integer
      Dim ThisPointInInputArray As Integer
      Dim ThisPointInWorkArray As Integer

      Me.IsUseFeedbackRegulatedVolatilityFastAttackEvent = False   'by default
      'bound to the input array size
      If ThisStartPointInInputArray < 0 Then
        ThisStartPointInInputArray = 0
      End If
      If ThisStopPointInInputArray >= VolatilityPredictionBandArray.Length Then
        ThisStopPointInInputArray = VolatilityPredictionBandArray.Length - 1
      End If
      MyStartPoint = StartPoint
      MyStopPoint = StopPoint
      MyRate = Rate
      MyRateForSigmaStatisticDaily = FilterPLLDetectorForVolatilitySigma.BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
      MyToCountLimit = ToCountLimit
      MyToCountLimitSelected = ToCountLimit
      MyToErrorLimit = ToErrorLimit
      MyProbOfBandExceedExpected = 1 - VolatilityPredictionBandArray(MyStartPoint).ProbabilityOfInterval

      'ThisStartPointInInputArray is adjusted to include additional points to take into account the include the buffer size needed for tha statistc calculation
      'note ThisStartPointInInputArray could be less than zero at the beginning of the array since the data is not provided below zero
      ThisStartPointInInputArray = MyStartPoint - MyRateForSigmaStatisticDaily + 1
      ThisStopPointInInputArray = MyStopPoint
      'next copy the data to a local array so that the data become independant from other threading object process and become 
      'completely local to this processing object
      ReDim MyVolatilityPredictionBandLocal(0 To (ThisStopPointInInputArray - ThisStartPointInInputArray))
      ThisPointInWorkArray = 0
      For ThisPointInInputArray = ThisStartPointInInputArray To ThisStopPointInInputArray
        'note that ThisPointInInputArray could be less than zero at the beginning
        'and in that case the object in the array is assigned to nothing
        If ThisPointInInputArray >= 0 Then
          'copy the object in the local working array
          MyVolatilityPredictionBandLocal(ThisPointInWorkArray) = New StockPriceVolatilityPredictionBand(VolatilityPredictionBandArray(ThisPointInInputArray))
        Else
          'no data availaible at that position
          'assign it to nothing
          MyVolatilityPredictionBandLocal(ThisPointInWorkArray) = Nothing
        End If
        ThisPointInWorkArray = ThisPointInWorkArray + 1
      Next
      'indicate the exact position of the start and stop point provided by the user in the local array
      'the extra point in the array are there for statictic measurement
      MyMapToLocalStartPoint = MyRateForSigmaStatisticDaily - 1
      MyMapToLocalStopPoint = MyMapToLocalStartPoint + MyStopPoint - MyStartPoint
      MyCount = 0
      MyValueInit = 0
      MyCountOfPLLRun = 0
      MyStatus = False
      MyFilterPLL = New FilterLowPassPLL(FilterRate:=7, DampingFactor:=1.0, NumberOfPredictionOutput:=0)
      MyListOfConvergence = New List(Of Double)
      MyListOfProbabilityOfExcess = New List(Of Double)
      MyListOfProbabilityOfExcessBalance = New List(Of Double)
      Me.Tag = TypeName(Me)
    End Sub

    Public Function RunErrorDetector(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.RunErrorDetector
      Dim ThisValueStart As Double = Me.ValueOutput(Input, InputFeedback)
      Dim ThisStockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand
      Dim IsBandExceededLast As Boolean

      Dim ThisGradientSum As Double
      Dim ThisGradientMean As Double
      Dim ThisValueOuput As Double = Me.ValueOutput(Input, InputFeedback)
      Dim ThisCount As Integer
      Dim I As Integer
      Dim ThisVolatilityChangePerCent As Double

      MyStatus = False
      MyErrorLast = 0   'by default
      'If MyCount Mod MyRateForSigmaStatisticDaily = 0 Then

      'process for volatility correction
      MyCountOfPLLRun = MyCountOfPLLRun + 1
      Do
        MySumForSigmaStatisticDaily = 0
        MyDetectorBandExcessBalanceSum = 0
        ThisGradientSum = 0
        ThisCount = 0
        If Input <= 0 Then Exit Do
        ThisVolatilityChangePerCent = InputFeedback / Input
        Dim ThisCountThresholdForFastAttack As Integer = CInt(0.8 * MyRateForSigmaStatisticDaily)
        ThisCountThresholdForFastAttack = (MyMapToLocalRunPoint - MyRateForSigmaStatisticDaily + 1) + ThisCountThresholdForFastAttack
        'MathPlus.WaveForm.SignalLimit(ThisVolatilityChangePerCent, MinScale:=-0.1, MaxScale:=1, Offset:=0)
        'do not calculate the statistic with the last point since it it would correspond to a point in the future
        For I = (MyMapToLocalRunPoint - MyRateForSigmaStatisticDaily + 1) To MyMapToLocalRunPoint - 1
          ThisStockPriceVolatilityPredictionBand = MyVolatilityPredictionBandLocal(I)
          If ThisStockPriceVolatilityPredictionBand IsNot Nothing Then
            With ThisStockPriceVolatilityPredictionBand
              If .Volatility = 0 Then
                Exit Do
              Else
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
          End If
        Next
        If ThisCount > 1 Then
          MyDetectorBalanceLast = MyDetectorBandExcessBalanceSum / ThisCount
          ThisGradientMean = ThisGradientSum / ThisCount
          If MySumForSigmaStatisticDaily > ThisCount Then
            MySumForSigmaStatisticDaily = ThisCount
          End If
          MyProbabilityOfExcessMeasuredLast = MySumForSigmaStatisticDaily / ThisCount
          MyErrorLast = (MyProbabilityOfExcessMeasuredLast - MyProbOfBandExceedExpected)
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
      Return MyErrorLast
    End Function

    Public Async Function UpdateAsync() As Task(Of Boolean)
      Dim ThisVolatilityDelta As Double
      Dim I As Integer
      Dim ThisVolatilityPredictionBandLocal As StockPriceVolatilityPredictionBand



      Dim ThisTaskRun = New Task(Of Boolean)(
        Function()
          For I = MyMapToLocalStartPoint To MyMapToLocalStopPoint
            MyMapToLocalRunPoint = I
            ThisVolatilityPredictionBandLocal = MyVolatilityPredictionBandLocal(I)
            'If Me.StopPoint = 870 Then
            '  If Me.Tag = "NFLX" Then
            '    If ThisVolatilityPredictionBandLocal.StockPrice.DateDay = New DateTime(2014, 11, 18) Then
            '      I = I
            '    End If
            '  End If
            'End If

            'note the PLL return the error VolatilityDelta, however it finish the calculation with the Volatility Total
            'obtained by calling the object ValueOutput function of this class. 
            'The filter PLL list and FilterLast contain the volatility total for the  final result
            ThisVolatilityDelta = MyFilterPLL.Filter(ThisVolatilityPredictionBandLocal.Volatility, Me)
            Me.RunErrorDetector(ThisVolatilityPredictionBandLocal.Volatility, ThisVolatilityDelta)
            ThisVolatilityPredictionBandLocal.Refresh(VolatilityDelta:=ThisVolatilityDelta)
            MyListOfProbabilityOfExcess.Add(MyProbabilityOfExcessMeasuredLast)
            MyListOfProbabilityOfExcessBalance.Add(MyDetectorBalanceLast)
            MyToCountLimitSelected = MyToCountLimit
          Next
          Return True
        End Function)

      ThisTaskRun.Start()
      'ThisTaskRun.RunSynchronously()
      Await ThisTaskRun
      Return ThisTaskRun.Result
    End Function

    Public ReadOnly Property StartPoint As Integer
      Get
        Return MyStartPoint
      End Get
    End Property

    Public ReadOnly Property StopPoint As Integer
      Get
        Return MyStopPoint
      End Get
    End Property

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
      Return FilterPLLDetectorForVolatilitySigma.RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * FilterPLLDetectorForVolatilitySigma.BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
    End Function

    Public Shared Function VolatilityRate(ByVal FilterRate As Double) As Double
      Return FilterPLLDetectorForVolatilitySigma.RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * FilterPLLDetectorForVolatilitySigma.BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
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
        Throw New NotImplementedException
      End Get
    End Property

    Private ReadOnly Property ToListOfPriceMedianNextDayHigh As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayHigh
      Get
        Throw New NotImplementedException
      End Get
    End Property

    Private ReadOnly Property ToListOfPriceMedianNextDayLow As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayLow
      Get
        Throw New NotImplementedException
      End Get
    End Property
  End Class
End Namespace
