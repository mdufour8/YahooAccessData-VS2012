#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region

Namespace MathPlus.Filter
  ''' <summary>
  ''' The standard definition is for N=14 and an output filter of 3 given by
  ''' read more at: http://www.investopedia.com/terms/s/stochasticoscillator.asp#ixzz2MnmOFLnS
  ''' %K = 100[(C - L14)/(H14 - L14)] 
  ''' Where
  ''' C = the most recent closing price
  ''' L14 = the low of the 14 previous trading sessions
  ''' H14 = the highest price traded during the same 14-day period.
  ''' %D = 3-period moving average of %K
  ''' </summary>
  ''' <remarks>The current implementation support any rate or output filter</remarks>
  <Serializable()>
  Public Class FilterStochastic
    Implements IStochastic

    Private Const FILTER_ATTACK_DECAY_RATIO As Double = 0.5
    Private MyValueMax As IPriceVolLarge
    Private MyValueMin As IPriceVolLarge
    Private MyListOfValueWindows As ListWindow(Of PriceVolLargeAsClass)
    Private MyListOfValueWindowsHigh As ListWindow(Of PriceVolLargeAsClass)
    Private MyListOfValueWindowsLow As ListWindow(Of PriceVolLargeAsClass)
    Private MyListOfWindowFrameHigh As ListWindowFrame
    Private MyListOfWindowFrameLow As ListWindowFrame
		'Private MyListOfWindowFrameHighHalf As ListWindowFrame
		'Private MyListOfWindowFrameLowHalf As ListWindowFrame
		Private MyRangeStatistic As FilterStatistical
		Private MyRate As Integer
    Private MyRateOutput As Integer
    Private MyRatePreFilter As Integer
    Private MyValueSum As Double
    Private MyValueSumSquare As Double
    Private MyStocFastLastHisteresis As Double
    Private MyStocFastSlowLast As Double
    Private MyFilterLPOfRange As FilterLowPassExp
    Private MyPriceVolForRangeLast As IPriceVolLarge

    Private MyValueLast As Double
    Private MyFilter As IFilter
    Private MyFilterHigh As IFilter
    Private MyFilterLow As IFilter
    Private MyFilterLPOfStochasticSlow As IFilter
    Private MyFilterLPOfFilterBackTo As FilterLowPassExp
    Private MyListOfStochasticFast As ListScaled
    Private MyListOfStochasticFastSlow As ListScaled
    Private MyListOfPriceBandHigh As List(Of Double)
    Private MyListOfPriceBandLow As List(Of Double)
    Private MyListOfPriceRangeVolatility As ListScaled
    Private MyRangeLast As Double
    Private MyStocFastLast As Double
    Private MyStocRangeVolatility As Double

    'Private MyFilterBollinger As FilterBollingerBand

#Region "New"
    Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
      Me.New(FilterRate, CDbl(FilterOutputRate))
    End Sub

    Public Sub New(ByVal IsPreFilter As Boolean, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
      Me.New(FilterRate, FilterRate, FilterOutputRate)
    End Sub

    Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
      Me.New(PreFilterRate, FilterRate, CDbl(FilterOutputRate))
    End Sub

    Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
      Me.New(FilterRate, FilterOutputRate)
      MyRatePreFilter = PreFilterRate
      If MyRatePreFilter < 1 Then MyRatePreFilter = 1
      If MyRatePreFilter > 1 Then
        MyFilter = New FilterLowPassExpHull(MyRatePreFilter)
      End If
    End Sub

    Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
      If FilterRate < 1 Then FilterRate = 1
      If FilterOutputRate < 1 Then FilterOutputRate = 1
      MyRatePreFilter = 1
      MyRate = FilterRate
      MyRateOutput = CInt(FilterOutputRate)
      MyListOfValueWindows = New ListWindow(Of PriceVolLargeAsClass)(FilterRate)
      MyListOfValueWindowsHigh = New ListWindow(Of PriceVolLargeAsClass)(FilterRate)
      MyListOfValueWindowsLow = New ListWindow(Of PriceVolLargeAsClass)(FilterRate)
      MyListOfWindowFrameHigh = New ListWindowFrame(FilterRate)
      MyListOfWindowFrameLow = New ListWindowFrame(FilterRate)

			MyRangeStatistic = New FilterStatistical(
				FilterRate:=FilterRate,
				StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)

			MyFilterLPOfRange = New FilterLowPassExp(FilterRate)
			'MyListOfValueWindows = New ListWindowOfPriceVol(Of Double)(FilterRate)
			'MyListOfValueWindows = New ListWindow(Of IPriceVolLarge)(FilterRate)
			MyListOfStochasticFast = New ListScaled
      'MyFilterLPOfStochasticSlow = New FilterLowPassExpHull(FilterOutputRate)
      MyFilterLPOfStochasticSlow = New FilterLowPassExp(FilterOutputRate)
      'MyFilterHigh = New FilterAttackDecayExp(1, FILTER_ATTACK_DECAY_RATIO * MyRate)
      'MyFilterLow = New FilterAttackDecayExp(FILTER_ATTACK_DECAY_RATIO * MyRate, 1)
      'MyFilterHigh = New FilterLowPassExpHull(MyRate)
      'MyFilterLow = New FilterLowPassExpHull(MyRate)
      MyFilterHigh = New FilterLowPassPLL(MyRate)
      MyFilterLow = New FilterLowPassPLL(MyRate)
      MyListOfStochasticFastSlow = New ListScaled
      MyListOfPriceBandHigh = New List(Of Double)
      MyListOfPriceBandLow = New List(Of Double)
      MyListOfPriceRangeVolatility = New ListScaled
      'MyFilterBollinger = New FilterBollingerBand(FilterRate)
    End Sub
#End Region

    Private Function Filter(ByVal Value As Single) As Double Implements IStochastic.Filter
      Return Me.Filter(CDbl(Value))
    End Function

    Private Function Filter(ByRef Value As Double) As Double Implements IStochastic.Filter
      If Me.IsFilterRange Then
        'simulate the range using a low pass filtering technique
        Dim ThisPriceVol As IPriceVolLarge = New PriceVolLarge(Value)

        MyFilterLPOfRange.Filter(Value)
        With ThisPriceVol
          If MyPriceVolForRangeLast IsNot Nothing Then
            .LastPrevious = MyPriceVolForRangeLast.Last
          End If
          If MyFilterLPOfRange.FilterLast > .Last Then
            .High = MyFilterLPOfRange.FilterLast
            .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
          ElseIf MyFilterLPOfRange.FilterLast < .Last Then
            .Low = MyFilterLPOfRange.FilterLast
            .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
          End If
        End With
        MyPriceVolForRangeLast = ThisPriceVol
        Return Me.FilterLocal(ThisPriceVol.Last, ThisPriceVol.Range)
      Else
        Return Me.FilterLocal(Value, 0)
      End If
    End Function

    Private Function Filter(ByRef Value As IPriceVol, FilterRate As Integer) As Double Implements IStochastic.Filter
      Return Me.FilterLocal(Value.Last, Value.Range)
    End Function

    Public Function Filter(ByRef Value As IPriceVol) As Double Implements IStochastic.Filter
      Return Me.FilterLocal(Value)
    End Function

    Private Function Filter(ByRef Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double Implements IStochastic.Filter
      'Return Me.FilterLocal(Value, ValueExpectedMin, ValueExpectedMax)
      Return Me.FilterLocal(Value)
    End Function

    'Public Function Filter(ByRef Value As IPriceVol, ByVal ValueToAddAndAverage As Single) As Double Implements IStochastic.Filter
    '  Return Me.Filter(CDbl((Value.Last + ValueToAddAndAverage) / 2))
    'End Function

    Public Function Filter(ByRef Value() As Double) As Double() Implements IStochastic.Filter
      Dim ThisValue As Double
      For Each ThisValue In Value
        Me.Filter(ThisValue)
      Next
      Return Me.ToArray
    End Function

    ''' <summary>
    ''' old calculation method 
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="Range"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FilterLocal(ByVal Value As Double, ByVal Range As Double) As Double
      Dim ThisPriceDelta As Double
      Dim ThisValueRemoved As IPriceVolLarge = Nothing
      Dim ThisVariance As Double
      Dim ThisMean As Double
      Dim ThisMeanSquared As Double
      Dim ThisSigmaMean As Double
      Dim ThisValueHigh As Double
      Dim ThisValueLow As Double
      Dim ThisStocRangeVolatility As Double
      Dim ThisStocFastLast As Double
      Dim ThisStocFastSlowLast As Double

      Dim ThisData As PriceVolLargeAsClass

#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual <> ThisResultPrediction Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
      MyValueLast = Value
      If MyFilter Is Nothing Then
        ThisData = New PriceVolLargeAsClass(Value) With {.FilterLast = Value, .Range = Range}
      Else
        ThisData = New PriceVolLargeAsClass(Value) With {.FilterLast = MyFilter.Filter(Value), .Range = Range}
      End If
      If MyListOfValueWindows.Count = 0 Then
        MyValueSum = 0
        MyValueSumSquare = 0
        ThisStocFastLast = 0.5
        ThisStocFastSlowLast = 0.5
      End If
      MyListOfValueWindows.Add(ThisData)
      ThisValueRemoved = MyListOfValueWindows.ItemRemoved
      MyValueMax = MyListOfValueWindows.ItemHigh
      MyValueMin = MyListOfValueWindows.ItemLow
      MyFilterHigh.Filter(MyValueMax.FilterLast)
      MyFilterLow.Filter(MyValueMin.FilterLast)
      If Me.IsFilterPeak Then
        ThisValueHigh = MyFilterHigh.FilterLast
        ThisValueLow = MyFilterLow.FilterLast
      Else
        ThisValueHigh = MyValueMax.FilterLast
        ThisValueLow = MyValueMin.FilterLast
      End If
      If Me.IsFilterRange Then
        If ThisValueRemoved IsNot Nothing Then
          MyValueSum = MyValueSum - ThisValueRemoved.Range
          MyValueSumSquare = MyValueSumSquare - (ThisValueRemoved.Range * ThisValueRemoved.Range)
        End If
        MyValueSum = MyValueSum + ThisData.Range
        MyValueSumSquare = MyValueSumSquare + (ThisData.Range * ThisData.Range)
        ThisMean = MyValueSum / MyListOfValueWindows.Count
        ThisMeanSquared = ThisMean * ThisMean
        ThisSigmaMean = MyValueSumSquare / MyListOfValueWindows.Count
        If ThisSigmaMean > ThisMeanSquared Then
          ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
        Else
          ThisVariance = 0
        End If
        MyRangeLast = ThisMean + ThisVariance

        ThisPriceDelta = ThisValueHigh - ThisValueLow + MyRangeLast
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (ThisData.FilterLast - (ThisValueLow - MyRangeLast / 2)) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      Else
        'calculate the variance using the variation on the Last value
        MyValueSum = MyValueSum + ThisData.Last
        MyValueSumSquare = MyValueSumSquare + (ThisData.Last * ThisData.Last)
        ThisMean = MyValueSum / MyListOfValueWindows.Count
        ThisMeanSquared = ThisMean * ThisMean
        ThisSigmaMean = MyValueSumSquare / MyListOfValueWindows.Count
        If ThisSigmaMean > ThisMeanSquared Then
          ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
        Else
          ThisVariance = 0
        End If
        'only use the variance in this case 
        MyRangeLast = ThisVariance

        ThisPriceDelta = ThisValueHigh - ThisValueLow
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (ThisData.FilterLast - ThisValueLow) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      End If
      Me.ListDataUpdate(ThisValueHigh, ThisValueLow, ThisStocRangeVolatility, ThisStocFastLast)
      Return MyFilterLPOfStochasticSlow.FilterLast
    End Function

    Private Function FilterLocal(ByVal Value As IPriceVol) As Double
      Dim ThisPriceDelta As Double
      Dim ThisValueRemoved As IPriceVolLarge = Nothing
      Dim ThisVariance As Double
      Dim ThisMean As Double
      Dim ThisValueHigh As Double
      Dim ThisValueLow As Double
      Dim ThisStocRangeVolatility As Double
      Dim ThisStocFastLast As Double


      With Me.UpdatePeak(Value)
        ThisValueLow = .Item1
        ThisValueHigh = .Item2
      End With
      If Me.IsFilterRange Then
        With MyRangeStatistic.Filter(Value.Range)
          ThisMean = .Mean
          ThisVariance = .StandardDeviation
        End With
        MyRangeLast = ThisMean + ThisVariance
        ThisPriceDelta = ThisValueHigh - ThisValueLow + MyRangeLast
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (Value.Last - (ThisValueLow - MyRangeLast / 2)) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      Else
        'calculate the variance using the variation on the Last value
        With MyRangeStatistic.Filter(Value.Last)
          ThisMean = .Mean
          ThisVariance = .StandardDeviation
        End With
        'only use the variance in this case 
        MyRangeLast = ThisVariance
        ThisPriceDelta = ThisValueHigh - ThisValueLow
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (Value.Last - ThisValueLow) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      End If
      Return Me.ListDataUpdate(ThisValueHigh, ThisValueLow, ThisStocRangeVolatility, ThisStocFastLast)
    End Function

    Friend Function UpdatePeak(ByVal ValueLow As Double, ByVal ValueHigh As Double) As Tuple(Of Double, Double)
      'MyValueLast = Value.Last
      If MyListOfWindowFrameHigh.Count = 0 Then
        MyValueSum = 0
        MyValueSumSquare = 0
        MyStocFastLast = 0.5
        MyStocFastSlowLast = 0.5
      End If
      MyListOfWindowFrameHigh.Add(ValueHigh)
      MyListOfWindowFrameLow.Add(ValueLow)
      MyFilterHigh.Filter(MyListOfWindowFrameHigh.ItemHigh.Value)
      MyFilterLow.Filter(MyListOfWindowFrameLow.ItemLow.Value)
      If Me.IsFilterPeak Then
        Return Tuple.Create(MyFilterLow.FilterLast, MyFilterHigh.FilterLast)
      Else
        Return Tuple.Create(MyListOfWindowFrameLow.ItemLow.Value, MyListOfWindowFrameHigh.ItemHigh.Value)
      End If
    End Function

    Friend Function UpdatePeak(ByRef Value As IPriceVol) As Tuple(Of Double, Double)
      Return Me.UpdatePeak(Value.Low, Value.High)
    End Function

    Friend Function ListDataUpdate(ByVal ValueHigh As Double, ByVal ValueLow As Double, ByVal Volatility As Double, ByVal StochasticRawValue As Double) As Double
      MyListOfPriceBandHigh.Add(ValueHigh)
      MyListOfPriceBandLow.Add(ValueLow)
      MyListOfPriceRangeVolatility.Add(Volatility)
      MyFilterLPOfStochasticSlow.Filter(StochasticRawValue)
      MyListOfStochasticFast.Add(StochasticRawValue)
      MyStocFastSlowLast = StochasticRawValue - MyFilterLPOfStochasticSlow.FilterLast
      MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
      MyStocRangeVolatility = Volatility
      MyStocFastLast = StochasticRawValue
      Return MyFilterLPOfStochasticSlow.FilterLast
    End Function

    Private Function FilterLocal(ByVal Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double
      Dim ThisPriceDelta As Double
      Dim ThisValueRemoved As IPriceVolLarge = Nothing
      Dim ThisVariance As Double
      Dim ThisMean As Double
      Dim ThisMeanSquared As Double
      Dim ThisSigmaMean As Double
      Dim ThisValueHigh As Double
      Dim ThisValueLow As Double
      Dim ThisStocRangeVolatility As Double
      Dim ThisStocFastLast As Double

      Dim ThisData As PriceVolLargeAsClass
      Dim ThisDataHigh As PriceVolLargeAsClass
      Dim ThisDataLow As PriceVolLargeAsClass

      MyValueLast = Value.Last
      If MyFilter Is Nothing Then
        ThisData = New PriceVolLargeAsClass(Value) With {.FilterLast = Value.Last, .Range = .Range}
      Else
        'ThisData = New PriceVolLargeAsClass(Value) With {.FilterLast = MyFilter.Filter(Value.Last), .Range = .Range}
        MyFilter.Filter(Value.Last)
        ThisData = New PriceVolLargeAsClass(Value) With {.FilterLast = Value.Last, .Range = .Range}
      End If
      ThisDataHigh = New PriceVolLargeAsClass(Value) With {.FilterLast = ValueExpectedMax, .Range = .Range}
      ThisDataLow = New PriceVolLargeAsClass(Value) With {.FilterLast = ValueExpectedMin, .Range = .Range}
      If MyListOfValueWindows.Count = 0 Then
        MyValueSum = 0
        MyValueSumSquare = 0
        MyStocFastLast = 0.5
        MyStocFastSlowLast = 0.5
      End If
      MyListOfValueWindows.Add(ThisData)
      MyListOfValueWindowsHigh.Add(ThisDataHigh)
      MyListOfValueWindowsLow.Add(ThisDataLow)

      ThisValueRemoved = MyListOfValueWindows.ItemRemoved
      MyValueMax = MyListOfValueWindowsHigh.ItemHigh
      MyValueMin = MyListOfValueWindowsLow.ItemLow
      MyFilterHigh.Filter(MyValueMax.FilterLast)
      MyFilterLow.Filter(MyValueMin.FilterLast)
      If Me.IsFilterPeak Then
        ThisValueHigh = MyFilterHigh.FilterLast
        ThisValueLow = MyFilterLow.FilterLast
      Else
        ThisValueHigh = MyValueMax.FilterLast
        ThisValueLow = MyValueMin.FilterLast
      End If
      If Me.IsFilterRange Then
        If ThisValueRemoved IsNot Nothing Then
          MyValueSum = MyValueSum - ThisValueRemoved.Range
          MyValueSumSquare = MyValueSumSquare - (ThisValueRemoved.Range * ThisValueRemoved.Range)
        End If
        MyValueSum = MyValueSum + ThisData.Range
        MyValueSumSquare = MyValueSumSquare + (ThisData.Range * ThisData.Range)
        ThisMean = MyValueSum / MyListOfValueWindows.Count
        ThisMeanSquared = ThisMean * ThisMean
        ThisSigmaMean = MyValueSumSquare / MyListOfValueWindows.Count
        If ThisSigmaMean > ThisMeanSquared Then
          ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
        Else
          ThisVariance = 0
        End If
        MyRangeLast = ThisMean + ThisVariance

        'ThisPriceDelta = ThisValueHigh - ThisValueLow + MyRangeLast
        ThisPriceDelta = ThisValueHigh - ThisValueLow
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (ThisData.FilterLast - ThisValueLow) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      Else
        'calculate the variance using the variation on the Last value
        MyValueSum = MyValueSum + ThisData.Last
        MyValueSumSquare = MyValueSumSquare + (ThisData.Last * ThisData.Last)
        ThisMean = MyValueSum / MyListOfValueWindows.Count
        ThisMeanSquared = ThisMean * ThisMean
        ThisSigmaMean = MyValueSumSquare / MyListOfValueWindows.Count
        If ThisSigmaMean > ThisMeanSquared Then
          ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
        Else
          ThisVariance = 0
        End If
        'only use the variance in this case 
        MyRangeLast = ThisVariance

        ThisPriceDelta = ThisValueHigh - ThisValueLow
        If ThisPriceDelta <> 0 Then
          ThisStocRangeVolatility = MyRangeLast / ThisPriceDelta
          ThisStocFastLast = (ThisData.FilterLast - ThisValueLow) / ThisPriceDelta
        Else
          ThisStocRangeVolatility = MyStocRangeVolatility
          ThisStocFastLast = MyStocFastLast
        End If
      End If
      MyListOfPriceBandHigh.Add(ThisValueHigh)
      MyListOfPriceBandLow.Add(ThisValueLow)
      MyListOfPriceRangeVolatility.Add(ThisStocRangeVolatility)
      MyFilterLPOfStochasticSlow.Filter(MyStocFastLast)
      MyStocFastSlowLast = MyStocFastLast - MyFilterLPOfStochasticSlow.FilterLast
      MyListOfStochasticFast.Add(MyStocFastLast)
      MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
      Return MyFilterLPOfStochasticSlow.FilterLast
    End Function

    Private Function FilterPredictionNext(ByRef Value As Double) As Double Implements IStochastic.FilterPredictionNext
      Dim ThisPriceDelta As Double
      Dim ThisData As Double

      Dim ThisValueMax As IPriceVolLarge
      Dim ThisValueMin As IPriceVolLarge
      Dim ThisStocFastLast As Double = MyStocFastLast
      Dim ThisStocFastSlowLast As Double = MyStocFastSlowLast

      Dim ThisVariance As Double
      Dim ThisMeanSquared As Double
      Dim ThisSigmaMean As Double
      Dim ThisFilterHighAttackDecay As Double
      Dim ThisFilterLowAttackDecay As Double


      Dim ThisValueSum = MyValueSum
      Dim ThisValueSumSquare = MyValueSumSquare

      If MyFilter Is Nothing Then
        ThisData = Value
      Else
        ThisData = MyFilter.FilterPredictionNext(Value)
      End If
      If MyListOfValueWindows.Count = 0 Then
        ThisValueSum = 0
        ThisValueSumSquare = 0
        ThisStocFastLast = 0.5
        ThisStocFastSlowLast = 0.5
      End If
      MyListOfValueWindows.Add(New PriceVolLargeAsClass(Value) With {.FilterLast = ThisData})
      ThisValueMax = MyListOfValueWindows.ItemHigh
      ThisValueMin = MyListOfValueWindows.ItemLow
      ThisFilterHighAttackDecay = MyFilterHigh.FilterPredictionNext(ThisValueMax.FilterLast)
      ThisFilterLowAttackDecay = MyFilterLow.FilterPredictionNext(ThisValueMin.FilterLast)

      ThisValueSum = ThisValueSum + ThisData
      ThisValueSumSquare = ThisValueSumSquare + (ThisData * ThisData)
      ThisMeanSquared = ThisValueSum / MyListOfValueWindows.Count
      ThisMeanSquared = ThisMeanSquared * ThisMeanSquared
      ThisSigmaMean = ThisValueSumSquare / MyListOfValueWindows.Count
      If ThisSigmaMean > ThisMeanSquared Then
        ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
      Else
        ThisVariance = 0
      End If

      If Me.IsFilterPeak Then
        ThisPriceDelta = ThisFilterHighAttackDecay - ThisFilterLowAttackDecay
        If ThisPriceDelta <> 0 Then
          ThisStocFastLast = (ThisData - ThisFilterLowAttackDecay) / ThisPriceDelta
        End If
      Else
        ThisPriceDelta = ThisValueMax.FilterLast - ThisValueMin.FilterLast
        If ThisPriceDelta <> 0 Then
          ThisStocFastLast = (ThisData - ThisValueMin.FilterLast) / ThisPriceDelta
        End If
      End If
      Debug.Assert(False)
      Dim ThisFilterLPOfStochasticSlow = MyFilterLPOfStochasticSlow.FilterPredictionNext(ThisStocFastLast)
      'MyStocFastSlowLast = MyStocFastLast - MyFilterLPOfStochasticSlow.FilterLast
      'MyListOfStochasticFast.Add(MyStocFastLast)
      'MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
      'restore the list 
      MyListOfValueWindows.RemoveLast()
      Return ThisFilterLPOfStochasticSlow
    End Function

    Public Function FilterBackTo(ByRef Value As Double, Optional ByVal IsPreFilter As Boolean = True) As Double Implements IStochastic.FilterBackTo
      Dim ThisStocFastLast As Double
      Dim ThisValue As Double

      Debug.Assert(False)
      ThisStocFastLast = MyFilterLPOfStochasticSlow.FilterBackTo(Value)
      If ThisStocFastLast > 1.0 Then
        ThisStocFastLast = 1
      ElseIf ThisStocFastLast < 0 Then
        ThisStocFastLast = 0
      End If
      ThisValue = ((MyListOfPriceBandHigh.Last - MyListOfPriceBandLow.Last) * ThisStocFastLast) + (MyListOfPriceBandLow.Last)
      If MyFilter IsNot Nothing Then
        If IsPreFilter Then
          ThisValue = MyFilter.FilterBackTo(ThisValue)
        End If
      End If
      Return ThisValue
    End Function

    Public Function FilterPriceBandHigh() As Double Implements IStochastic.FilterPriceBandHigh
      Return MyListOfPriceBandHigh.Last
    End Function

    Public Function FilterPriceBandLow() As Double Implements IStochastic.FilterPriceBandLow
      Return MyListOfPriceBandLow.Last
    End Function

    Public Function FilterLast() As Double Implements IStochastic.FilterLast
      Return MyFilterLPOfStochasticSlow.FilterLast
    End Function

    ''' <summary>
    ''' return the last result of the fast-slow stochastic calculation
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FilterLast(ByVal Type As IStochastic.enuStochasticType) As Double Implements IStochastic.FilterLast
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
        Case Else
          Return MyStocFastSlowLast
      End Select
    End Function

    Public Function Last() As Double Implements IStochastic.Last
      Return MyValueLast
    End Function

    Public Property Rate(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Integer Implements IStochastic.Rate
      Get
        Select Case Type
          Case IStochastic.enuStochasticType.FastSlow
            Return MyRate
          Case IStochastic.enuStochasticType.Fast
            Return MyRatePreFilter
          Case IStochastic.enuStochasticType.Slow
            Return MyRateOutput
          Case Else
            Return 1
        End Select
      End Get
      Set(value As Integer)
        'do not set the rate here
      End Set
    End Property

    Public ReadOnly Property Count As Integer Implements IStochastic.Count
      Get
        Return MyFilterLPOfStochasticSlow.Count
      End Get
    End Property

    Public ReadOnly Property Max(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Max
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

    Public ReadOnly Property Min(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Min
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

    Public ReadOnly Property ToList() As IList(Of Double) Implements IStochastic.ToList
      Get
        Return MyFilterLPOfStochasticSlow.ToList
      End Get
    End Property

    Public Overridable ReadOnly Property ToList(ByVal Type As IStochastic.enuStochasticType) As IList(Of Double) Implements IStochastic.ToList
      Get
        Select Case Type
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
          Case IStochastic.enuStochasticType.RangeVolatility
            Return MyListOfPriceRangeVolatility
          Case IStochastic.enuStochasticType.ProbabilityHigh, IStochastic.enuStochasticType.ProbabilityLow
            Throw New NotSupportedException
          Case Else
            Return MyListOfStochasticFastSlow
        End Select
      End Get
    End Property

    Public Function ToArray(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
      Select Case Type
        Case IStochastic.enuStochasticType.FastSlow
          Return MyListOfStochasticFastSlow.ToArray
        Case IStochastic.enuStochasticType.Fast
          Return MyListOfStochasticFast.ToArray
        Case IStochastic.enuStochasticType.Slow
          Return MyFilterLPOfStochasticSlow.ToArray
        Case Else
          Return MyListOfStochasticFastSlow.ToArray
      End Select
    End Function

    Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
      Return Me.ToArray(Me.Min(Type), Me.Max(Type), ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
      Select Case Type
        Case IStochastic.enuStochasticType.FastSlow
          Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
        Case IStochastic.enuStochasticType.Fast
          Return MyListOfStochasticFast.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
        Case IStochastic.enuStochasticType.Slow
          Return MyFilterLPOfStochasticSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
        Case Else
          Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
      End Select
    End Function

    Public Property Tag As String Implements IStochastic.Tag

    Public Overrides Function ToString() As String Implements IStochastic.ToString
      Return Me.FilterLast.ToString
    End Function

    Public Property IsFilterPeak As Boolean Implements IStochastic.IsFilterPeak
    Public Property IsFilterRange As Boolean Implements IStochastic.IsFilterRange
  End Class
End Namespace