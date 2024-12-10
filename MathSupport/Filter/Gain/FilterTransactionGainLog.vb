#Region "Imports"
'Imports MathNet.Numerics
'Imports MathNet.Numerics.RootFinding
'Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region
Namespace MathPlus.Filter
  <Serializable()>
  Public Class FilterTransactionGainLog
    Implements IFilterPrediction

    Private Enum enuControlStart
      Start
      WaitForRate
      WaitForTransaction
      Run
    End Enum

    Private MyRate As Integer
    Private MyFilterTotalGainValueLast As Double
    Private ValueLast As Double
    Private ValueTransactionLast As Double
    Private MySwitchRef As Double
    Private MyListOfFilterTotalGainValueLast As ListScaled
    Private MySwitchControlLast As FilterTransactionGain.enumControlBuySell
    Private MyNumberOfTransaction As Integer
    Private MyTransactionStart As Integer
    Private IsValueChanging As Boolean
    Private MyNumberOfValueWithChange As Integer
    Private MyValueOfSwitchControl As Double
    Private MyValueOfSwitchControlLast As Double
    Private MyTransaction As ITransaction
    Private MyControlStart As enuControlStart
    Private MyTransactionCostPerCent As Double
    Private MyStartPoint As Integer
    'Private MyFilterPrediction As Filter.FilterLowPassExpPredict
    Private MyFilterPrediction As IFilter
    Private MyFilterPredictionForGainMeasurement As IFilter
    Private MyListOfGainRMS As IList(Of Double)
    Private MyListOfGainTransactionPerformance As IList(Of Double)

    Private MyValueTransactionStopLast As Double
    Private MyGainMeasurementPeriod As Integer
    Private MyFilterGainAverage As FilterLowPassExp
    Private MyFilterGainRMSAverage As FilterLowPassExp


    Public Sub New()
      Me.New(1, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer)
      Me.New(FilterRate, 0.0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal GainMeasurementPeriod As Integer)
      Me.New(FilterRate, GainMeasurementPeriod, 0.0, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal GainMeasurementPeriod As Integer, ByVal TransactionCostPerCent As Double)
      Me.New(FilterRate, GainMeasurementPeriod, TransactionCostPerCent, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal TransactionCostPerCent As Double)
      Me.New(FilterRate, 20, TransactionCostPerCent, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal GainMeasurementPeriod As Integer, ByVal TransactionCostPerCent As Double, ByVal StartPoint As Integer)
      MyListOfFilterTotalGainValueLast = New ListScaled
      MyRate = FilterRate
      MyFilterTotalGainValueLast = 0
      ValueLast = 0.0
      ValueTransactionLast = 0.0
      MySwitchRef = 0
      MyValueOfSwitchControl = 0
      MyTransactionCostPerCent = TransactionCostPerCent
      MyStartPoint = StartPoint
      MyGainMeasurementPeriod = GainMeasurementPeriod  'by default
      Me.GainSaturationForLimiting = 1.0

      'MyFilterPrediction = New Filter.FilterLowPassExpPredict(NUMBER_TRADINGDAY_PER_YEAR, 0)
      'If FilterRate < 5 Then
      '  MyFilterPrediction = New Filter.FilterLowPassExpPredict(5, 0)
      'Else
      '  MyFilterPrediction = New Filter.FilterLowPassExpPredict(FilterRate, 0)
      'End If
      'this is similar to the measure of the gain for the average period of 4 transaction
      'MyFilterPrediction = New Filter.FilterLowPassExpPredict(4 * FilterRate, 0)
      MyFilterPrediction = New Filter.FilterLowPassExpPredict(MyRate, 0)
      'MyFilterPrediction = New Filter.FilterLowPassPLL(MyGainMeasurementPeriod, IsPredictionEnabled:=True)
      'MyFilterPredictionForGainMeasurement = New Filter.FilterLowPassPLL(MyGainMeasurementPeriod, IsPredictionEnabled:=True)
      MyFilterPredictionForGainMeasurement = New Filter.FilterLowPassExpPredict(MyRate, 0)
      MyListOfGainRMS = New List(Of Double)
      MyFilterGainAverage = New FilterLowPassExp(FilterRate:=MyGainMeasurementPeriod)
      MyFilterGainRMSAverage = New FilterLowPassExp(FilterRate:=MyGainMeasurementPeriod)
      MyListOfGainTransactionPerformance = New List(Of Double)
    End Sub

    ''' <summary>
    '''   Calculate the gain base on the following linear weighted value:
    '''              n
    '''   Gain(n)= Sum[(Weight(n)*(Price(n) - Price(n-1))]
    '''              0
    ''' </summary>
    ''' <param name="Value">
    '''   The price
    ''' </param>
    ''' <param name="WeightControl">
    '''   The weight generally between -1 to +1 
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Filter(
      ByVal Value As Double,
      ByVal WeightControl As Double,
      Optional ByVal ValueTransactionStop As Double = -1) As Double

      Return Me.Filter(New PriceVol(CSng(Value)), WeightControl:=WeightControl, ValueTransactionStop:=ValueTransactionStop)
    End Function

		Public Function Filter(
			ByVal Value As IPriceVol,
			ByVal WeightControl As Double) As Double

			Return Me.Filter(Value, WeightControl, -1)
		End Function

		Public Shared Function ScaleToUnit(Data As Double, ThresholdDown As Double, ThresholdUp As Double) As Double
			Select Case Data
				Case < ThresholdDown
					Return -1.0
				Case > ThresholdUp
					Return 1.0
				Case Else
					Return 0.0
			End Select
		End Function

		Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal WeightControl As Double,
      ByVal ValueTransactionStop As Double) As Double

			Return Me.Filter(Value, WeightControl, ValueTransactionStop, PriceRangeHigh:=-1, PriceRangeLow:=-1)
		End Function

    Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal WeightControl As Double,
      ByVal ValueTransactionStop As Double,
      ByVal PriceRangeHigh As Double,
      ByVal PriceRangeLow As Double) As Double


      Dim ThisSampleDeltaGain As Double
      Dim ThisSwitchControl As FilterTransactionGain.enumControlBuySell
      Dim ThisWeightControlLimit As Double
      Dim ThisValueTransactionStop As Double
      Dim IsLocalValueStopEnabled As Boolean = False
      Dim ThisNumberOfTransaction As Integer = 0
      Dim ThisGainRMS As Double

      If Me.GainSaturationForLimiting > 0 Then
        ThisWeightControlLimit = MathPlus.WaveForm.SignalLimit(Me.GainSaturationForLimiting * WeightControl, 1)
      Else
        'no limiting here
        ThisWeightControlLimit = WeightControl
      End If
      If Me.IsPriceStopEnabled Then
        If ValueTransactionStop > 0 Then
          IsLocalValueStopEnabled = True
        End If
      End If
			ThisSwitchControl = FilterTransactionGain.ToSwitchScale(ThisWeightControlLimit, ThresholdLevelLow:=0, ThresholdLevelHigh:=0)

			'If Value.DateDay = #1/26/2015# Then
			'  Value = Value
			'End If
			MyValueOfSwitchControl = ThisWeightControlLimit
      If MyListOfFilterTotalGainValueLast.Count = 0 Then
        MyControlStart = enuControlStart.Start
        MyFilterTotalGainValueLast = 0
        ValueLast = Value.Last
        ValueTransactionLast = ValueLast
        IsValueChanging = False
        MyNumberOfValueWithChange = 0
        MyValueOfSwitchControl = 0
        MyValueTransactionStopLast = ValueTransactionStop
      End If
      Select Case MyControlStart
        Case enuControlStart.Start
          If Value.Last <> ValueLast Then
            'this is when the data started to change
            IsValueChanging = True
            MyNumberOfValueWithChange = 1
            MyControlStart = enuControlStart.WaitForRate
          End If
        Case enuControlStart.WaitForRate
          If MyListOfFilterTotalGainValueLast.Count >= MyStartPoint Then
            If MyNumberOfValueWithChange >= MyRate Then
              MyControlStart = enuControlStart.WaitForTransaction
            Else
              MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
            End If
          Else
            MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
          End If
        Case enuControlStart.WaitForTransaction
          If MyRate = 0 Then
            'do not wait for a transaction
            MyControlStart = enuControlStart.Run
          Else
            If MySwitchControlLast <> ThisSwitchControl Then
              MyControlStart = enuControlStart.Run
            End If
          End If
        Case enuControlStart.Run
          If IsLocalValueStopEnabled Then
            If MyValueOfSwitchControlLast > 0 Then
              'this is a buy situation
              'check if price low value reached the stop loss limit
              'if it did we assume the sale at the value of stop loss if it is contained inside the high and low for the day
              ThisValueTransactionStop = MyValueTransactionStopLast
              If Value.Low < ThisValueTransactionStop Then
                'yes the stop loss has been it 
                'check the stock opening price
                If Value.Open < ThisValueTransactionStop Then
                  'the sale likeley happen at the open price
                  'set the last switch control to zero since the sale was completed today
                  'at this point we assume the transaction was not reversed on the other side
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Open, ValueLast))
                Else
                  'if it did not happen at the open then if did happen at the stop value sometime in the day
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(ThisValueTransactionStop, ValueLast))
                End If
                'set the last switch control to zero since the sale was completed
                ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Hold
                MyValueOfSwitchControlLast = 0
                'selling on a stop is one transaction
                ThisNumberOfTransaction = 1
                If IsInversePositionOnPriceStopEnabled Then
                  'the transaction was completed at the price Stop
                  'check if the last price is still below the stop price
                  'in that case assume that the trading desk is now negatif and not with zero position
                  If Value.Last < ThisValueTransactionStop Then
                    'one additional transaction to reverse position
                    ThisNumberOfTransaction = 2
                    ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Sell
                    MyValueOfSwitchControlLast = -1
                  End If
                End If
              Else
                'stop was not reach
                ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
              End If
            ElseIf MyValueOfSwitchControlLast < 0 Then
              'this is a sale situation
              'check if price high value reached the stop buy limit
              'if it did we assume the sale at the value of stop loss if it is contained inside the high and low for the day
              ThisValueTransactionStop = MyValueTransactionStopLast
              If Value.High > ThisValueTransactionStop Then
                'yes the stop buy limit has been it 
                If Value.Open > ThisValueTransactionStop Then
                  'the buy likeley happened at the open price
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Open, ValueLast))
                Else
                  'if it did not happen at the open then if did happen at the stop value sometime in the day
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(ThisValueTransactionStop, ValueLast))
                End If
                'set the last switch control to zero since the sale was completed
                ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Hold
                MyValueOfSwitchControlLast = 0
                'selling on a stop is one transaction
                ThisNumberOfTransaction = 1
                If IsInversePositionOnPriceStopEnabled Then
                  'the transaction was completed at the price Stop
                  'check if the last price is still higher than the stop price
                  'in that case assume that the trading desk is now positif and not with zero position
                  If Value.Last > ThisValueTransactionStop Then
                    'one additional transaction to reverse position
                    ThisNumberOfTransaction = 2
                    MyValueOfSwitchControlLast = 1
                    ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Buy
                  End If
                End If
              Else
                'stop was not reach
                ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
              End If
            Else
              'no open position on the last trade when MyValueOfSwitchControlLast=0
              ThisSampleDeltaGain = 0
            End If
          Else
            ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
          End If
          If ThisSwitchControl <> MySwitchControlLast Then
            If ThisNumberOfTransaction = 0 Then
              ThisNumberOfTransaction = 1
            End If
            MyNumberOfTransaction = MyNumberOfTransaction + ThisNumberOfTransaction
            MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain - (ThisNumberOfTransaction * (MyTransactionCostPerCent / 100))
            'MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain
          Else
            MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain
          End If
      End Select
      MyListOfFilterTotalGainValueLast.Add(MyFilterTotalGainValueLast)
      'simulate the price gain on a stock with a price starting at 100
      MyFilterPrediction.Filter(GainLogInverse(MyFilterTotalGainValueLast, ValueRef:=100.0))
      MyFilterPredictionForGainMeasurement.Filter(Value.Last)
      ThisGainRMS = Math.Sqrt(DirectCast(MyFilterPredictionForGainMeasurement, IFilterPrediction).ToListOfGainPerYear.Last ^ 2)
      If MyControlStart = enuControlStart.Run Then
        MyListOfGainRMS.Add(ThisGainRMS)
      Else
        MyListOfGainRMS.Add(0.0)
      End If
      MyFilterGainAverage.Filter(Me.ToListOfGainPerYear.Last)
      MyFilterGainRMSAverage.Filter(ThisGainRMS)
      If MyFilterGainRMSAverage.FilterLast > 0 Then
        MyListOfGainTransactionPerformance.Add(MyFilterGainAverage.FilterLast / MyFilterGainRMSAverage.FilterLast)
      Else
        MyListOfGainTransactionPerformance.Add(0.0)
      End If
      MyValueOfSwitchControlLast = MyValueOfSwitchControl
      MySwitchControlLast = ThisSwitchControl
      ValueLast = Value.Last
      ValueTransactionLast = ValueLast
      MyValueTransactionStopLast = ValueTransactionStop
      Return MyFilterTotalGainValueLast
    End Function

    Public Async Function FilterAsync(
      ByVal ReportPrices As YahooAccessData.RecordPrices,
      ByVal WeightControl As IList(Of Double),
      ByVal ValueTransactionStop As IList(Of Double),
      ByVal PriceRangeHigh As IList(Of Double),
      ByVal PriceRangeLow As IList(Of Double)) As Task(Of Boolean)

      Dim ThisTaskRun As Task(Of Boolean)

      ThisTaskRun = New Task(Of Boolean)(
        Function()
          Dim J As Integer
          'note that a positive value indicate thst this list is a predictive list 
          Dim IShifted As Integer = ValueTransactionStop.Count - ReportPrices.NumberPoint
          For I = 0 To ReportPrices.NumberPoint - 1
            J = I + IShifted
            Me.Filter(ReportPrices.PriceVols(I), WeightControl(I), ValueTransactionStop(J), PriceRangeHigh(J), PriceRangeLow(J))
          Next
          Return True
        End Function)

      ThisTaskRun.Start()
      Await ThisTaskRun
      Return ThisTaskRun.Result
    End Function

    Public Async Function Filter1Async(
      ByVal ReportPrices As YahooAccessData.RecordPrices,
      ByVal WeightControl As IList(Of Double),
      ByVal ValueTransactionStop As IList(Of Double),
      ByVal PriceRangeHigh As IList(Of Double),
      ByVal PriceRangeLow As IList(Of Double)) As Task(Of Boolean)

      Dim ThisTaskRun As Task(Of Boolean)

      ThisTaskRun = New Task(Of Boolean)(
        Function()
          Dim I As Integer
          Dim J As Integer
          'note that a positive value indicate thst this list is a predictive list 
          Dim IShifted As Integer = ValueTransactionStop.Count - ReportPrices.NumberPoint
          For I = 0 To ReportPrices.NumberPoint - 1
            J = I + IShifted
            Me.Filter1(ReportPrices.PriceVols(I), WeightControl(I), ValueTransactionStop(J), PriceRangeHigh(J), PriceRangeLow(J))
          Next
          Return True
        End Function)

      ThisTaskRun.Start()
      Await ThisTaskRun
      Return ThisTaskRun.Result
    End Function

    Public Function Filter1(
      ByVal Value As IPriceVol,
      ByVal WeightControl As Double,
      ByVal ValueTransactionStop As Double,
      ByVal PriceRangeHigh As Double,
      ByVal PriceRangeLow As Double) As Double


      Dim ThisSampleDeltaGain As Double
      Dim ThisSwitchControl As FilterTransactionGain.enumControlBuySell
      Dim ThisWeightControlLimit As Double
      Dim ThisValueTransactionStop As Double
      Dim IsLocalValueStopEnabled As Boolean = False
      Dim ThisNumberOfTransaction As Integer = 0
      Dim ThisWeightControl As Double
      Dim ThisGainRMS As Double

      Dim ThisPriceMedian As Double = (PriceRangeHigh + PriceRangeLow) / 2
      If ValueTransactionStop < ThisPriceMedian Then
        ThisWeightControl = 1 * WeightControl
      Else
        ThisWeightControl = -1 * WeightControl
      End If

      If Me.GainSaturationForLimiting > 0 Then
        'ThisWeightControlLimit = MathPlus.WaveForm.SignalLimit(Me.GainSaturationForLimiting * ThisWeightControl, 1)
        ThisWeightControlLimit = ThisWeightControl
      Else
        'no limiting here
        ThisWeightControlLimit = ThisWeightControl
      End If
      If Me.IsPriceStopEnabled Then
        If ValueTransactionStop > 0 Then
          IsLocalValueStopEnabled = True
        End If
      End If
      ThisSwitchControl = FilterTransactionGain.ToSwitchScale(ThisWeightControlLimit, 0, 0)

      'If Value.DateDay = #1/26/2015# Then
      '  Value = Value
      'End If
      MyValueOfSwitchControl = ThisWeightControlLimit
      If MyListOfFilterTotalGainValueLast.Count = 0 Then
        MyControlStart = enuControlStart.Start
        MyFilterTotalGainValueLast = 0
        ValueLast = Value.Last
        ValueTransactionLast = ValueLast
        IsValueChanging = False
        MyNumberOfValueWithChange = 0
        MyValueOfSwitchControl = 0
        MyValueTransactionStopLast = ValueTransactionStop
      End If
      Select Case MyControlStart
        Case enuControlStart.Start
          If Value.Last <> ValueLast Then
            'this is when the data started to change
            IsValueChanging = True
            MyNumberOfValueWithChange = 1
            MyControlStart = enuControlStart.WaitForRate
          End If
        Case enuControlStart.WaitForRate
          If MyListOfFilterTotalGainValueLast.Count >= MyStartPoint Then
            If MyNumberOfValueWithChange >= MyRate Then
              MyControlStart = enuControlStart.WaitForTransaction
            Else
              MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
            End If
          Else
            MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
          End If
        Case enuControlStart.WaitForTransaction
          If MyRate = 0 Then
            'do not wait for a transaction
            MyControlStart = enuControlStart.Run
          Else
            If MySwitchControlLast <> ThisSwitchControl Then
              MyControlStart = enuControlStart.Run
            End If
          End If
        Case enuControlStart.Run
          If IsLocalValueStopEnabled Then
            If MyValueOfSwitchControlLast > 0 Then
              'this is a buy situation
              'check if price low value reached the stop loss limit
              'if it did we assume the sale at the value of stop loss if it is contained inside the high and low for the day
              ThisValueTransactionStop = MyValueTransactionStopLast
              If Value.Low < ThisValueTransactionStop Then
                'yes the stop loss has been it 
                'check the stock opening price
                If Value.Open < ThisValueTransactionStop Then
                  'the sale likeley happen at the open price
                  'set the last switch control to zero since the sale was completed today
                  'at this point we assume the transaction was not reversed on the other side
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Open, ValueLast))
                Else
                  'if it did not happen at the open then if did happen at the stop value sometime in the day
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(ThisValueTransactionStop, ValueLast))
                End If
                'set the last switch control to zero since the sale was completed
                ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Hold
                MyValueOfSwitchControlLast = 0
                'selling on a stop is one transaction
                ThisNumberOfTransaction = 1
                If IsInversePositionOnPriceStopEnabled Then
                  'the transaction was completed at the price Stop
                  'check if the last price is still below the stop price
                  'in that case assume that the trading desk is now negatif and not with zero position
                  If Value.Last < ThisValueTransactionStop Then
                    'one additional transaction to reverse position
                    ThisNumberOfTransaction = 2
                    ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Sell
                    MyValueOfSwitchControlLast = -1
                  End If
                End If
              Else
                'stop was not reach
                ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
              End If
            ElseIf MyValueOfSwitchControlLast < 0 Then
              'this is a sale situation
              'check if price high value reached the stop buy limit
              'if it did we assume the sale at the value of stop loss if it is contained inside the high and low for the day
              ThisValueTransactionStop = MyValueTransactionStopLast
              If Value.High > ThisValueTransactionStop Then
                'yes the stop buy limit has been it 
                If Value.Open > ThisValueTransactionStop Then
                  'the buy likeley happened at the open price
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Open, ValueLast))
                Else
                  'if it did not happen at the open then if did happen at the stop value sometime in the day
                  ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(ThisValueTransactionStop, ValueLast))
                End If
                'set the last switch control to zero since the sale was completed
                ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Hold
                MyValueOfSwitchControlLast = 0
                'selling on a stop is one transaction
                ThisNumberOfTransaction = 1
                If IsInversePositionOnPriceStopEnabled Then
                  'the transaction was completed at the price Stop
                  'check if the last price is still higher than the stop price
                  'in that case assume that the trading desk is now positif and not with zero position
                  If Value.Last > ThisValueTransactionStop Then
                    'one additional transaction to reverse position
                    ThisNumberOfTransaction = 2
                    MyValueOfSwitchControlLast = 1
                    ThisSwitchControl = FilterTransactionGain.enumControlBuySell.Buy
                  End If
                End If
              Else
                'stop was not reach
                ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
              End If
            Else
              'no open position on the last trade when MyValueOfSwitchControlLast=0
              ThisSampleDeltaGain = 0
            End If
          Else
            ThisSampleDeltaGain = MyValueOfSwitchControlLast * (GainLog(Value.Last, ValueLast))
          End If
          If ThisSwitchControl <> MySwitchControlLast Then
            If ThisNumberOfTransaction = 0 Then
              ThisNumberOfTransaction = 1
            End If
            MyNumberOfTransaction = MyNumberOfTransaction + ThisNumberOfTransaction
            MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain - (ThisNumberOfTransaction * (MyTransactionCostPerCent / 100))
            'MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain
          Else
            MyFilterTotalGainValueLast = MyFilterTotalGainValueLast + ThisSampleDeltaGain
          End If
      End Select
      MyListOfFilterTotalGainValueLast.Add(MyFilterTotalGainValueLast)
      'simulate the price gain on a stock with a price starting at 100
      MyFilterPrediction.Filter(GainLogInverse(MyFilterTotalGainValueLast, 100.0))
      MyFilterPredictionForGainMeasurement.Filter(Value.Last)
      ThisGainRMS = Math.Sqrt(DirectCast(MyFilterPredictionForGainMeasurement, IFilterPrediction).ToListOfGainPerYear.Last ^ 2)
      If MyControlStart = enuControlStart.Run Then
        MyListOfGainRMS.Add(ThisGainRMS)
      Else
        MyListOfGainRMS.Add(0.0)
      End If
      MyFilterGainAverage.Filter(Me.ToListOfGainPerYear.Last)
      If Double.IsNaN(MyFilterGainAverage.FilterLast) Then
        Debugger.Break()
      End If
      MyFilterGainRMSAverage.Filter(ThisGainRMS)
      If MyFilterGainRMSAverage.FilterLast > 0 Then
        MyListOfGainTransactionPerformance.Add(MyFilterGainAverage.FilterLast / MyFilterGainRMSAverage.FilterLast)
      Else
        MyListOfGainTransactionPerformance.Add(0.0)
      End If
      MyValueOfSwitchControlLast = MyValueOfSwitchControl
      MySwitchControlLast = ThisSwitchControl
      ValueLast = Value.Last
      ValueTransactionLast = ValueLast
      MyValueTransactionStopLast = ValueTransactionStop
      Return MyFilterTotalGainValueLast
    End Function


    Public Property IsInversePositionOnPriceStopEnabled As Boolean
    Public Property IsPriceStopEnabled As Boolean
    Public Property GainSaturationForLimiting As Double

    ''' <summary>
    ''' The price cost per transaction in percent
    ''' </summary>
    ''' <value></value>
    ''' <returns>The price cost per transaction in %</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TransactionCost As Double
      Get
        Return MyTransactionCostPerCent
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return MyNumberOfTransaction
      End Get
    End Property

    Public ReadOnly Property TransactionStart() As Integer
      Get
        Return MyTransactionStart
      End Get
    End Property

    Public Function FilterLast() As Double
      Return MyFilterTotalGainValueLast
    End Function

    Public Function Last() As Double
      Return ValueLast
    End Function

    Public ReadOnly Property Rate As Integer
      Get
        Return MyRate
      End Get
    End Property

    Public ReadOnly Property Count As Integer
      Get
        Return MyListOfFilterTotalGainValueLast.Count
      End Get
    End Property

    Public ReadOnly Property ScaleRange As Double
      Get
        If Me.Max > Math.Abs(Me.Min) Then
          Return Me.Max
        Else
          Return Math.Abs(Me.Min)
        End If
      End Get
    End Property

    Public Property Max As Double
      Set(value As Double)
        MyListOfFilterTotalGainValueLast.Max = value
      End Set
      Get
        Return MyListOfFilterTotalGainValueLast.Max
      End Get
    End Property

    Public Property Min As Double
      Set(value As Double)
        MyListOfFilterTotalGainValueLast.Min = value
      End Set
      Get
        Return MyListOfFilterTotalGainValueLast.Min
      End Get
    End Property

    Public ReadOnly Property ToList() As IList(Of Double)
      Get
        Return MyListOfFilterTotalGainValueLast
      End Get
    End Property

    Public ReadOnly Property ToListScaled() As ListScaled
      Get
        Return MyListOfFilterTotalGainValueLast
      End Get
    End Property

    Public Function ToArray() As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray
    End Function

    Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray(ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Property Tag As String

    Public Overrides Function ToString() As String
      Return Me.FilterLast.ToString
    End Function

    Public Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
      Return Me
    End Function

    Private Function FilterPrediction(NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
      Throw New NotImplementedException
    End Function

    Private Function FilterPrediction(NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
      Throw New NotImplementedException
    End Function

    Private Function FilterPrediction(Index As Integer, NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
      Throw New NotImplementedException
    End Function

    Private Function FilterPrediction(Index As Integer, NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
      Throw New NotImplementedException
    End Function

    Private ReadOnly Property IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
      Get
        Return True
      End Get
    End Property

    Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
      Get
        Return DirectCast(MyFilterPrediction, IFilterPrediction).ToListOfGainPerYear
      End Get
    End Property

    Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
      Get
        Return DirectCast(MyFilterPrediction, IFilterPrediction).ToListOfGainPerYearDerivative
      End Get
    End Property

    Public ReadOnly Property ToListOfGainPerYearRMS As System.Collections.Generic.IList(Of Double)
      Get
        Return MyListOfGainRMS
      End Get
    End Property

    Public ReadOnly Property ToListOfGainPerYearRMSAverage As System.Collections.Generic.IList(Of Double)
      Get
        Return MyListOfGainRMS
      End Get
    End Property

    Public ReadOnly Property ToListOfGainPerYear As System.Collections.Generic.IList(Of Double)
      Get
        Return IFilterPrediction_ToListOfGainPerYear
      End Get
    End Property

    Public ReadOnly Property ToListOfGainPerYearAverage As System.Collections.Generic.IList(Of Double)
      Get
        Return MyFilterGainAverage.ToList
      End Get
    End Property

    Public ReadOnly Property ToListOfGainTransactionPerformance As System.Collections.Generic.IList(Of Double)
      Get
        Return MyListOfGainTransactionPerformance
      End Get
    End Property

    Public Function GainAverage(ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
      Return Mean(Me.ToListOfGainPerYear, StartPoint, StopPoint, IsValidatePointRange:=True)
    End Function

    Public Function GainPerYearRMSAverage(ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
      Return Mean(MyListOfGainRMS, StartPoint, StopPoint, IsValidatePointRange:=True)
    End Function

    Public Function GainTransactionPerformance(ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
      Dim ThisGainAverage As Double = GainAverage(StartPoint, StopPoint)
      Dim ThisGainSquaredAverage As Double = GainPerYearRMSAverage(StartPoint, StopPoint)
      Dim ThisGainTransactionPerformance As Double

      If ThisGainSquaredAverage > 0.0 Then
        ThisGainTransactionPerformance = ThisGainAverage / (ThisGainSquaredAverage)
      Else
        ThisGainTransactionPerformance = 0.0
      End If
      Return ThisGainTransactionPerformance
    End Function
  End Class

  <Serializable()>
  Public Class FilterTransactionGain
    Public Enum enumControlBuySell
      Buy
      Sell
      Hold
    End Enum
    Private Enum enuControlStart
      Start
      WaitForRate
      WaitForTransaction
      Run
    End Enum

    Private MyRate As Integer
    Private MyFilterTotalGainValueLast As Double
    Private ValueLast As IPriceVol
    Private ValueFirst As IPriceVol
    Private ValueLastStop As Double
    Private MySwitchRef As Double
    Private MyFilterValueRef As Double
    Private MyListOfFilterTotalGainValueLast As ListScaled
    Private MySwitchControlLast As enumControlBuySell
    Private MyNumberOfTransaction As Integer
    Private MyTransactionStart As Integer
    Private IsValueChanging As Boolean
    Private MyNumberOfValueWithChange As Integer
    Private MyValueOfSwitchControl As Double
    Private MyWeightControlSumSquared As Double
    Private MyWeightControlRMS As Double
    Private MyWeightControlCount As Integer
    Private MyValueOfSwitchControlLast As Double
    Private MyTransaction As ITransaction
    Private MyControlStart As enuControlStart
    Private MyTransactionCost As Double
    Private MyStartPoint As Integer

    Public Sub New()
      Me.New(1, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer)
      Me.New(FilterRate, 0.0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal TransactionCostPerCent As Double)
      Me.New(FilterRate, TransactionCostPerCent, 0)
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal TransactionCostPerCent As Double, ByVal StartPoint As Integer)
      MyListOfFilterTotalGainValueLast = New ListScaled
      MyRate = FilterRate
      MyFilterTotalGainValueLast = 0
      ValueLast = New PriceVol(0)
      MySwitchRef = 0
      MyFilterValueRef = 0
      MyValueOfSwitchControl = 0
      Me.IsStopReverse = True
      Me.IsStopEnabled = True
      Me.IsTransactionCanReStartAfterStop = True
      MyTransactionCost = TransactionCostPerCent / 100
      MyStartPoint = StartPoint
    End Sub

    ''' <summary>
    '''   Calculate the gain base on the following linear weighted value:
    '''              n
    '''   Gain(n)= Sum[(Weight(n)*(Price(n) - Price(n-1))]
    '''              0
    ''' </summary>
    ''' <param name="Value">
    '''   The price
    ''' </param>
    ''' <param name="WeightControl">
    '''   The weight generally between -1 to +1 
    ''' </param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal WeightControl As Double) As Double

      Dim ThisSwitchControl As enumControlBuySell = FilterTransactionGain.ToSwitchScale(WeightControl, 0, 0)
      'If Value.DateDay = #1/26/2015# Then
      '  Value = Value
      'End If
      MyValueOfSwitchControl = WeightControl
      If MyListOfFilterTotalGainValueLast.Count = 0 Then
        MyControlStart = enuControlStart.Start
        MyFilterTotalGainValueLast = 0
        MyFilterValueRef = MyFilterTotalGainValueLast
        ValueLast = Value
        ValueFirst = Nothing
        IsValueChanging = False
        MyNumberOfValueWithChange = 0
        MyValueOfSwitchControl = 0
        MyWeightControlSumSquared = 0
        MyWeightControlCount = 0
      End If
      Select Case MyControlStart
        Case enuControlStart.Start
          If ValueLast.Last <> Value.Last Then
            'this is when the data started to change
            IsValueChanging = True
            MyNumberOfValueWithChange = 1
            MyWeightControlCount = MyWeightControlCount + 1
            MyWeightControlSumSquared = MyWeightControlSumSquared + MyValueOfSwitchControl ^ 2
            MyWeightControlRMS = Math.Sqrt(MyWeightControlSumSquared / MyWeightControlCount)
            MyControlStart = enuControlStart.WaitForRate
          End If
        Case enuControlStart.WaitForRate
          MyWeightControlCount = MyWeightControlCount + 1
          MyWeightControlSumSquared = MyWeightControlSumSquared + MyValueOfSwitchControl ^ 2
          MyWeightControlRMS = Math.Sqrt(MyWeightControlSumSquared / MyWeightControlCount)
          If MyListOfFilterTotalGainValueLast.Count >= MyStartPoint Then
            If MyNumberOfValueWithChange >= MyRate Then
              MyControlStart = enuControlStart.WaitForTransaction
            Else
              MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
            End If
          Else
            MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
          End If
        Case enuControlStart.WaitForTransaction
          MyWeightControlCount = MyWeightControlCount + 1
          MyWeightControlSumSquared = MyWeightControlSumSquared + MyValueOfSwitchControl ^ 2
          MyWeightControlRMS = Math.Sqrt(MyWeightControlSumSquared / MyWeightControlCount)
          If MyRate = 0 Then
            'do not wait for a transaction
            MyControlStart = enuControlStart.Run
            ValueFirst = ValueLast
          Else
            If MySwitchControlLast <> ThisSwitchControl Then
              MyControlStart = enuControlStart.Run
              ValueFirst = ValueLast
            End If
          End If
        Case enuControlStart.Run
          'Try
          MyWeightControlCount = MyWeightControlCount + 1
          MyWeightControlSumSquared = MyWeightControlSumSquared + MyValueOfSwitchControl ^ 2
          MyWeightControlRMS = Math.Sqrt(MyWeightControlSumSquared / MyWeightControlCount)
          'MyFilterTotalGainValueLast = MyFilterValueRef + MyValueOfSwitchControl * (Value.Last - ValueLast.Last)
          'MyFilterTotalGainValueLast = MyFilterValueRef + MyValueOfSwitchControl * (Value.Last - Value.Open)
          'MyFilterTotalGainValueLast = MyFilterValueRef + MyValueOfSwitchControl * (Value.Last - Value.Open)
          'MyFilterTotalGainValueLast = MyFilterValueRef + MyValueOfSwitchControl * ((Value.Last - ValueLast.Last) / ValueLast.Last)
          If MyWeightControlRMS > 0 Then
            Dim ThisSampleDeltaGain As Double
            Dim ThisWeightDeltaGain As Double

            ThisSampleDeltaGain = MyValueOfSwitchControlLast * ((Value.Last - ValueLast.Last) / ValueFirst.Last) / MyWeightControlRMS
            'this is the log version
            'ThisSampleDeltaGain = MyValueOfSwitchControlLast * (Math.Log(Value.Last / ValueLast.Last) / MyWeightControlRMS)

            ThisWeightDeltaGain = ((MyValueOfSwitchControl - MyValueOfSwitchControlLast) / MyWeightControlRMS)
            'Dim ThisTransactionCost As Double = 0.1 * Math.Abs(YahooAccessData.MathPlus.WaveForm.SignalLimit(ThisWeightDeltaGain / MyWeightControlRMS, 1))

            'MyFilterTotalGainValueLast = MyFilterValueRef + (1 - ThisTransactionCost) * (ThisSampleDeltaGain / MyWeightControlRMS)
            MyFilterTotalGainValueLast = MyFilterValueRef + ThisSampleDeltaGain

            If Math.Abs(ThisWeightDeltaGain) > 0.5 Then
              'If MySwitchControlLast <> ThisSwitchControl Then
              'add some transaction cost
              MyFilterTotalGainValueLast = MyFilterTotalGainValueLast - MyTransactionCost / 2
              MyNumberOfTransaction = MyNumberOfTransaction + 1
              'If MyListOfFilterTotalGainValueLast.Count = 300 Then
              'Debugger.Break()
              'End If
            End If
          End If
          'Catch ex As Exception
          '  ex = ex
          'End Try
          MyFilterValueRef = MyFilterTotalGainValueLast
      End Select
      MyListOfFilterTotalGainValueLast.Add(MyFilterTotalGainValueLast)
      ValueLastStop = Value.Last
      MyValueOfSwitchControlLast = MyValueOfSwitchControl
      MySwitchControlLast = ThisSwitchControl
      ValueLast = Value
      Return MyFilterTotalGainValueLast
    End Function

    Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal ValueStop As IPriceVol,
      ByVal ValueSwitchControl As Double,
      ByVal ThresholdLevelLow As Double,
      ByVal ThresholdLevelHigh As Double) As Double

      MyValueOfSwitchControl = ValueSwitchControl
      Return Me.Filter(Value, ValueStop, FilterTransactionGain.ToSwitchScale(MyValueOfSwitchControl, ThresholdLevelLow, ThresholdLevelHigh))
    End Function

    Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal ValueStop As IPriceVol,
      ByVal SwitchControl As enumControlBuySell) As Double

      Dim IsValueStopped As Boolean = False
      Dim ThisTransactionNew As ITransaction = Nothing
      Dim ThisSwitchControl As enumControlBuySell = SwitchControl

      If MyListOfFilterTotalGainValueLast.Count = 0 Then
        MyFilterTotalGainValueLast = 0
        MyFilterValueRef = MyFilterTotalGainValueLast
        ValueLast = Value
        IsValueChanging = False
        MyNumberOfValueWithChange = 0
      End If
      If IsValueChanging = False Then
        If ValueLast.Last <> Value.Last Then
          'this is when the data started to change
          IsValueChanging = True
          MyNumberOfValueWithChange = 1
        End If
      Else
        MyNumberOfValueWithChange = MyNumberOfValueWithChange + 1
      End If
      If MyNumberOfValueWithChange >= MyRate Then
        If MySwitchControlLast <> ThisSwitchControl Then
          'take a new reference for gain calculation
          Select Case ThisSwitchControl
            Case enumControlBuySell.Hold
              'close the transaction
              If MyTransaction IsNot Nothing Then
                MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value)
                MyFilterValueRef = MyFilterTotalGainValueLast
              End If
              MyTransaction = Nothing
            Case enumControlBuySell.Buy
              ThisTransactionNew = New TransactionStockBuy With {.IsStopReverse = Me.IsStopReverse, .IsStopEnabled = Me.IsStopEnabled}
              If MyNumberOfTransaction = 0 Then
                MyTransactionStart = Me.Count + 1
              End If
              MyNumberOfTransaction = MyNumberOfTransaction + 1
              ThisTransactionNew.Filter(Value, ValueStop)
              If MyTransaction IsNot Nothing Then
                'MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value, ValueStop)
                MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value)
              End If
              'next time we will work on the new transaction using a new ref
              MyFilterValueRef = MyFilterTotalGainValueLast
              MyTransaction = ThisTransactionNew
            Case enumControlBuySell.Sell
              ThisTransactionNew = New TransactionStockSell With {.IsStopReverse = Me.IsStopReverse, .IsStopEnabled = Me.IsStopEnabled}
              If MyNumberOfTransaction = 0 Then
                MyTransactionStart = Me.Count + 1
              End If
              MyNumberOfTransaction = MyNumberOfTransaction + 1
              ThisTransactionNew.Filter(Value, ValueStop)
              If MyTransaction IsNot Nothing Then
                'If MyTransaction.IsStop = False Then
                'MyTransaction.Tag = MyTransaction.Tag
                'End If
                'do not use the stop for the final calculation
                MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value)
              End If
              'next time we will work on the new transaction using a new ref
              MyFilterValueRef = MyFilterTotalGainValueLast
              MyTransaction = ThisTransactionNew
          End Select
        Else
          If MyTransaction IsNot Nothing Then
            'check if the transaction was previously stopped
            If MyTransaction.IsStop = False Then
              'no stop, continue measuring the performance
              MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value, ValueStop)
            Else
              'If Me.Tag = "AAPL" Then
              'Me.Tag = Me.Tag
              'End If
              'transaction was previosuly stopped
              'keep mesuring performance but check if the stop was a false flag
              'and the stock is a again a buy
              MyFilterTotalGainValueLast = MyFilterValueRef + MyTransaction.Filter(Value, ValueStop)
              If MyTransaction.CountStop >= 1 Then
                If Me.IsTransactionCanReStartAfterStop Then
                  'do not use this for now
                  Select Case MyTransaction.Type
                    Case ITransaction.enuTransactionType.StockBuy
                      If MyValueOfSwitchControl > MyValueOfSwitchControlLast Then
                        If Value.Last > MyTransaction.PriceStopValue.High Then
                          'stock rebound above the stop high
                          'create a new transaction buy putting the current stopped one to nothing
                          MyTransaction = Nothing
                          ThisSwitchControl = enumControlBuySell.Hold
                        End If
                      End If
                    Case ITransaction.enuTransactionType.StockOptionCall
                      If MyValueOfSwitchControl > MyValueOfSwitchControlLast Then
                        If Value.Last > MyTransaction.PriceStopValue.High Then
                          'stock rebound above the stop high
                          'create a new transaction buy putting the current stopped one to nothing
                          MyTransaction = Nothing
                          ThisSwitchControl = enumControlBuySell.Hold
                        End If
                      End If
                    Case ITransaction.enuTransactionType.StockSell
                      If MyValueOfSwitchControl < MyValueOfSwitchControlLast Then
                        If Value.Last < MyTransaction.PriceStopValue.Low Then
                          'stock crashed below the stop low
                          'create a new sell transaction buy putting the current stopped one to nothing
                          MyTransaction = Nothing
                          ThisSwitchControl = enumControlBuySell.Hold
                        End If
                      End If
                    Case ITransaction.enuTransactionType.StockOptionPut
                      If MyValueOfSwitchControl < MyValueOfSwitchControlLast Then
                        If Value.Last < MyTransaction.PriceStopValue.Low Then
                          'stock crashed below the stop low
                          'create a new sell transaction buy putting the current stopped one to nothing
                          MyTransaction = Nothing
                          ThisSwitchControl = enumControlBuySell.Hold
                        End If
                      End If
                  End Select
                End If
              End If
            End If
          End If
        End If
        MyListOfFilterTotalGainValueLast.Add(MyFilterTotalGainValueLast)
      Else
        MyListOfFilterTotalGainValueLast.Add(MyFilterTotalGainValueLast)
      End If
      'If Me.Tag = "AAPL" Then
      '  Me.Tag = Me.Tag

      'End If
      If MyTransaction IsNot Nothing Then
        ValueLastStop = MyTransaction.PriceStop
      Else
        ValueLastStop = ValueStop.Last
      End If
      MyValueOfSwitchControlLast = MyValueOfSwitchControl
      MySwitchControlLast = ThisSwitchControl
      ValueLast = Value
      Return MyFilterTotalGainValueLast
    End Function

    Public Function Filter(
      ByVal Value As IPriceVol,
      ByVal ValueStop As IPriceVol,
      ByVal ValueSwitchControl As Double) As Double

      Return Me.Filter(Value, ValueStop, ValueSwitchControl, 0.5, 0.5)
    End Function

    Public ReadOnly Property TransactionCost As Double
      Get
        Return MyTransactionCost
      End Get
    End Property

    Public ReadOnly Property TransactionNumber() As Integer
      Get
        Return MyNumberOfTransaction
      End Get
    End Property

    Public ReadOnly Property TransactionStart() As Integer
      Get
        Return MyTransactionStart
      End Get
    End Property

    Public Shared Function ToSwitchScale(ByVal Value As Double) As enumControlBuySell
      Return FilterTransactionGain.ToSwitchScale(Value, 0.5, 0.5)
    End Function

    Public Shared Function ToSwitchScale(ByVal Value As Double, ByVal ThresholdLevelLow As Double, ByVal ThresholdLevelHigh As Double) As enumControlBuySell
      If Value > ThresholdLevelHigh Then
        Return enumControlBuySell.Buy
      ElseIf Value < ThresholdLevelLow Then
        Return enumControlBuySell.Sell
      Else
        Return enumControlBuySell.Hold
      End If
    End Function

    ''' <summary>
    ''' not finish yet
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="ResultLast"></param>
    ''' <param name="ThresholdLevelLow"></param>
    ''' <param name="ThresholdLevelHigh"></param>
    ''' <param name="HysteresisLevel"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function ToSwitchScale(ByVal Value As Double, ByVal ResultLast As enumControlBuySell, ByVal ThresholdLevelLow As Double, ByVal ThresholdLevelHigh As Double, ByVal HysteresisLevel As Double) As enumControlBuySell
      Throw New NotSupportedException("ToSwitchScale")
      Select Case ResultLast
        Case enumControlBuySell.Buy
          If Value > (ThresholdLevelHigh - HysteresisLevel) Then
            Return enumControlBuySell.Buy
          ElseIf Value <= (ThresholdLevelHigh - HysteresisLevel) Then
            'If Value>
            Return enumControlBuySell.Sell
          End If
        Case enumControlBuySell.Sell

        Case enumControlBuySell.Hold

      End Select

    End Function

    Public Shared Function ToSwitchScale(ByVal Value() As Double, ByVal ThresholdLevelLow As Double, ByVal ThresholdLevelHigh As Double) As enumControlBuySell()
      Dim I As Integer
      Dim ThisResult() As enumControlBuySell
      ReDim ThisResult(0 To Value.Length - 1)

      For I = 0 To Value.Length - 1
        If Value(I) > ThresholdLevelHigh Then
          ThisResult(I) = enumControlBuySell.Buy
        ElseIf Value(I) < ThresholdLevelLow Then
          ThisResult(I) = enumControlBuySell.Sell
        Else
          ThisResult(I) = enumControlBuySell.Hold
        End If
      Next
      Return ThisResult
    End Function

    Public Shared Function ToSwitchBinaryAND(ByVal ValueA As enumControlBuySell, ByVal ValueB As enumControlBuySell) As enumControlBuySell
      Select Case ValueA
        Case enumControlBuySell.Buy
          If ValueB = enumControlBuySell.Buy Then
            Return enumControlBuySell.Buy
          Else
            Return enumControlBuySell.Hold
          End If
        Case enumControlBuySell.Sell
          If ValueB = enumControlBuySell.Sell Then
            Return enumControlBuySell.Sell
          Else
            Return enumControlBuySell.Hold
          End If
        Case Else
          Return enumControlBuySell.Hold
      End Select
    End Function

    Public Function FilterLast() As Double
      Return MyFilterTotalGainValueLast
    End Function

    Public Function Last() As Double
      Return ValueLast.Last
    End Function

    Public Function LastStop() As Double
      Return ValueLastStop
    End Function

    Public Property IsStopEnabled As Boolean
    Public Property IsStopReverse As Boolean
    Public Property IsTransactionCanReStartAfterStop As Boolean

    Public ReadOnly Property Rate As Integer
      Get
        Return MyRate
      End Get
    End Property

    Public ReadOnly Property Count As Integer
      Get
        Return MyListOfFilterTotalGainValueLast.Count
      End Get
    End Property

    Public ReadOnly Property ScaleRange As Double
      Get
        If Me.Max > Math.Abs(Me.Min) Then
          Return Me.Max
        Else
          Return Math.Abs(Me.Min)
        End If
      End Get
    End Property

    Public Property Max As Double
      Set(value As Double)
        MyListOfFilterTotalGainValueLast.Max = value
      End Set
      Get
        Return MyListOfFilterTotalGainValueLast.Max
      End Get
    End Property

    Public Property Min As Double
      Set(value As Double)
        MyListOfFilterTotalGainValueLast.Min = value
      End Set
      Get
        Return MyListOfFilterTotalGainValueLast.Min
      End Get
    End Property

    Public ReadOnly Property ToList() As IList(Of Double)
      Get
        Return MyListOfFilterTotalGainValueLast
      End Get
    End Property

    Public ReadOnly Property ToListScaled() As ListScaled
      Get
        Return MyListOfFilterTotalGainValueLast
      End Get
    End Property

    Public Function ToArray() As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray
    End Function

    Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray(ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Return MyListOfFilterTotalGainValueLast.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Property Tag As String

    Public Overrides Function ToString() As String
      Return Me.FilterLast.ToString
    End Function
  End Class
End Namespace