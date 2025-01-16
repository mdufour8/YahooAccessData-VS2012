Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions

Namespace MathPlus.Filter
  ''' <summary>
  ''' see the photo of: Price Probability Zero Stochatic Crossing PLL Filter.jpg
  ''' </summary>
  ''' <remarks></remarks>
  Friend Class FilterPLLDetectorForCDFToZero
    Implements IFilterPLLDetector

    Private MyPriceForMidStochastic As Double
    Private MyStochasticProbabilityHisteresis As Double
    Private IsStochasticProbabilityHisteresis As Boolean
    Private MyVolatility As Double
    Private MyGain As Double
    Private MyGainDerivative As Double
    Private MyPriceEndHigh As Double
    Private MyPriceEndLow As Double
    Private MyRate As Integer
    Private MyCount As Integer
    Private MyErrorLast As Double
    Private MyToCountLimit As Integer
    Private MyToErrorLimit As Double
    Private MyValueInit As Double
    Private MyValueOutput As Double
    Private MyFilterPLL As FilterLowPassPLL
    Private MyListOfConvergence As IList(Of Double)
    Private MyListOfPriceMedianNextDayLow As IList(Of Double)
    Private MyListOfPriceMedianNextDayHigh As IList(Of Double)
    Private MyFilterVolatilityForPriceStochasticMedian As FilterVolatilityYangZhang
    Private MyRateForVolatility As Integer
    Private MyFilterPLLForGain As FilterLowPassPLL


    Public Sub New(
      ByVal Rate As Integer,
      Optional ByVal ToCountLimit As Integer = 1,
      Optional ToErrorLimit As Double = 0.001,
      Optional StochasticProbabilityHisteresisLevel As Double = 0.0)
      MyRate = Rate
      MyToCountLimit = ToCountLimit
      MyToErrorLimit = ToErrorLimit
      MyCount = 0
      MyValueInit = 0
      MyStochasticProbabilityHisteresis = 0.5
      If StochasticProbabilityHisteresisLevel <> 0 Then
        IsStochasticProbabilityHisteresis = True
        MyStochasticProbabilityHisteresis = MyStochasticProbabilityHisteresis + StochasticProbabilityHisteresisLevel
      End If
      'this is working but leave it to false for now
      IsStochasticProbabilityHisteresis = False
			MyFilterPLL = New FilterLowPassPLL(FilterRate:=7, DampingFactor:=1.0, NumberOfPredictionOutput:=0)
			MyListOfConvergence = New List(Of Double)
      MyRateForVolatility = CInt(FilterPLLDetectorForVolatilitySigma.VolatilityRate)
      'MyFilterVolatilityForPriceStochasticMedian = New FilterVolatilityYangZhang(MyRateForVolatility, FilterVolatility.enuVolatilityStatisticType.Exponential)
      MyFilterVolatilityForPriceStochasticMedian = New FilterVolatilityYangZhang(MyRate, FilterVolatility.enuVolatilityStatisticType.Exponential)
      MyListOfPriceMedianNextDayLow = New List(Of Double)
      MyListOfPriceMedianNextDayHigh = New List(Of Double)
      MyFilterPLLForGain = New FilterLowPassPLL(MyRate, IsPredictionEnabled:=True)
      Me.Tag = TypeName(Me)  'by default
    End Sub

    Public Function RunErrorDetector(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.RunErrorDetector
      Dim ThisProbHigh As Tuple(Of Double, Double)
      Dim ThisProbLow As Tuple(Of Double, Double)
      Dim ThisCDFDetectorSlope As Double

      Dim ThisValueStart As Double = Me.ValueOutput(Input, InputFeedback)

      'If Math.Abs(InputFeedback / Input) > 0.2 Then
      '  ThisValueStart = ThisValueStart
      'End If
      If MyVolatility > 0 Then
        ThisProbHigh = Me.StockPricePredictionInverse(MyRate, ThisValueStart, MyGain, MyGainDerivative, MyVolatility, MyPriceEndHigh)
        ThisProbLow = Me.StockPricePredictionInverse(MyRate, ThisValueStart, MyGain, MyGainDerivative, MyVolatility, MyPriceEndLow)
        If IsStochasticProbabilityHisteresis Then
          ThisCDFDetectorSlope = (MyStochasticProbabilityHisteresis * ThisProbLow.Item2 + (1 - MyStochasticProbabilityHisteresis) * ThisProbHigh.Item2) / 2
          If ThisCDFDetectorSlope < 0 Then
            MyErrorLast = ((MyStochasticProbabilityHisteresis * ThisProbLow.Item1) - ((1 - MyStochasticProbabilityHisteresis) * (1 - ThisProbHigh.Item1))) / ThisCDFDetectorSlope
          Else
            MyErrorLast = ((MyStochasticProbabilityHisteresis * ThisProbLow.Item1) - ((1 - MyStochasticProbabilityHisteresis) * (1 - ThisProbHigh.Item1))) / -1
          End If
        Else
          ThisCDFDetectorSlope = (ThisProbLow.Item2 + ThisProbHigh.Item2) / 2
          If ThisCDFDetectorSlope < 0 Then
            MyErrorLast = ((ThisProbLow.Item1) - ((1 - ThisProbHigh.Item1))) / ThisCDFDetectorSlope
          Else
            MyErrorLast = ((ThisProbLow.Item1) - ((1 - ThisProbHigh.Item1))) / -1
          End If
        End If
      Else
        MyErrorLast = 0.0
      End If
      MyErrorLast = ThisValueStart * YahooAccessData.MathPlus.WaveForm.SignalLimit(MyErrorLast / ThisValueStart, 0.5)
      'MyErrorLast = ThisValueStart * YahooAccessData.MathPlus.WaveForm.SignalLimit(MyErrorLast / ThisValueStart, 0.5)
      MyCount = MyCount + 1
      Return MyErrorLast
    End Function


    ''' <summary>
    ''' Update all the function parameters
    ''' </summary>
    ''' <param name="Volatility"></param>
    ''' <param name="Gain"></param>
    ''' <param name="GainDerivative"></param>
    ''' <param name="PriceEndHigh"></param>
    ''' <param name="PriceEndLow"></param>
    ''' <remarks></remarks>
    Public Sub Update(
                     ByVal Volatility As Double,
                     ByVal Gain As Double,
                     ByVal GainDerivative As Double,
                     ByVal PriceEndHigh As Double,
                     ByVal PriceEndLow As Double)


      Dim ThisVolatility As Double
      Dim ThisGainPerYear As Double
      Dim ThisGainPerYearDerivative As Double
      Dim ThisTimeInYear As Double = MyRate / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisGain As Double = Gain + (ThisTimeInYear * GainDerivative)
      Dim ThisMu As Double = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear

      MyVolatility = Volatility
      MyGain = Gain
      MyGainDerivative = GainDerivative
      MyPriceEndHigh = PriceEndHigh
      MyPriceEndLow = PriceEndLow
      'MyFilterPLL.Filter((MyPriceEndLow + MyPriceEndHigh) / 2, Me)
      'If MyFilterPLL.Count = 0 Then
      '  ThisPriceStart = (MyPriceEndLow + MyPriceEndHigh) / 2
      'Else
      '  ThisPriceStart = MyFilterPLL.FilterLast
      '  'ThisPriceStart = (MyPriceEndLow + MyPriceEndHigh) / 2
      'End If
      'ThisPriceStart = (MyPriceEndLow + MyPriceEndHigh) / 2
      'the exact solution is:

      'Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)
      MyPriceForMidStochastic = Math.Sqrt(MyPriceEndLow * MyPriceEndHigh) * Math.Exp(-ThisMu)
      'MyPriceForMidStochastic = Math.Sqrt(MyPriceEndLow * MyPriceEndHigh)
      MyFilterPLLForGain.Filter(MyPriceForMidStochastic)
      ThisVolatility = MyFilterVolatilityForPriceStochasticMedian.Filter(MyPriceForMidStochastic)
      ThisGainPerYear = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYear.Last
      ThisGainPerYearDerivative = MyFilterPLLForGain.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last

      Dim ThisStockPriceHighValue = StockOption.StockPricePrediction(
        NumberTradingDays:=1,
        StockPrice:=MyPriceForMidStochastic,
        Gain:=ThisGainPerYear,
        GainDerivative:=ThisGainPerYearDerivative,
        Volatility:=ThisVolatility,
        Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)

      Dim ThisStockPriceLowValue = StockOption.StockPricePrediction(
        NumberTradingDays:=1,
        StockPrice:=MyPriceForMidStochastic,
        Gain:=ThisGainPerYear,
        GainDerivative:=ThisGainPerYearDerivative,
        Volatility:=ThisVolatility,
        Probability:=GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)

      MyListOfPriceMedianNextDayLow.Add(ThisStockPriceLowValue)
      MyListOfPriceMedianNextDayHigh.Add(ThisStockPriceHighValue)
      If IsStochasticProbabilityHisteresis Then
        'note IsStochasticProbabilityHisteresis is not normally being use yet
        MyFilterPLL.Filter(MyPriceForMidStochastic, Me)
      Else
        MyFilterPLL.ToList.Add(MyPriceForMidStochastic)
      End If
    End Sub

    ''' <summary>
    ''' Update some function parameters with teh Gain and GainDerivative fixed to zero
    ''' </summary>
    ''' <param name="Volatility"></param>
    ''' <param name="PriceEndHigh"></param>
    ''' <param name="PriceEndLow"></param>
    ''' <remarks></remarks>
    Public Sub Update(
                     ByVal Volatility As Double,
                     ByVal PriceEndHigh As Double,
                     ByVal PriceEndLow As Double)

      Me.Update(Volatility:=Volatility, Gain:=0, GainDerivative:=0, PriceEndHigh:=PriceEndHigh, PriceEndLow:=PriceEndLow)
    End Sub

    Public ReadOnly Property ToErrorLimit As Double Implements IFilterPLLDetector.ToErrorLimit
      Get
        Return MyToErrorLimit
      End Get
    End Property

    Public ReadOnly Property ToCount As Integer Implements IFilterPLLDetector.ToCount
      Get
        Return MyToCountLimit
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
      Return Input - InputFeedback
    End Function

    Private Function StockPricePredictionInverse(
                                                ByVal NumberTradingDays As Integer,
                                                ByVal StockPriceStart As Double,
                                                ByVal Gain As Double,
                                                ByVal GainDerivative As Double,
                                                ByVal Volatility As Double,
                                                ByVal StockPriceEnd As Double) As Tuple(Of Double, Double)
      Dim ThisResultOfCDF As Double
      Dim ThisResultForSlope As Double
      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisGain As Double = Gain + (ThisTimeInYear * GainDerivative)
      'Dim ThisPricePrediction As Double = StockOption.StockPricePrediction(NumberTradingDays, StockPrice, Gain)
      'Dim ThisPricePredictionMedian As Double = StockOption.StockPricePredictionMedian(NumberTradingDays, StockPrice, Gain, Volatility)

      'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
      'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
      'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
      Dim ThisMu As Double = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
      Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)

      If ThisSigma > 0 Then
        ThisResultOfCDF = Distributions.LogNormal.CDF(ThisMu, ThisSigma, StockPriceEnd / StockPriceStart)
        ThisResultForSlope = Distributions.LogNormal.PDF(ThisMu, ThisSigma, StockPriceEnd / StockPriceStart)
        'now calculate the final slope
        ThisResultForSlope = -ThisResultForSlope * StockPriceEnd / (StockPriceStart ^ 2)
      Else
        ThisResultOfCDF = 0.5
        ThisResultForSlope = 0
      End If
      Return New Tuple(Of Double, Double)(ThisResultOfCDF, ThisResultForSlope)
    End Function

    Public ReadOnly Property Status As Boolean Implements IFilterPLLDetector.Status
      Get
        Return True     'always true for this object
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
        Return 0.0
      End Get
    End Property

    Public ReadOnly Property Minimum As Double Implements IFilterPLLDetector.Minimum
      Get
        Return 0.0
      End Get
    End Property

    Public ReadOnly Property DetectorBalance As Double Implements IFilterPLLDetector.DetectorBalance
      Get
        Throw New NotImplementedException
      End Get
    End Property

    Public ReadOnly Property ToListOfVolatility As IList(Of Double) Implements IFilterPLLDetector.ToListOfVolatility
      Get
        Return MyFilterVolatilityForPriceStochasticMedian.ToList
      End Get
    End Property

    Public ReadOnly Property ToList As IList(Of Double) Implements IFilterPLLDetector.ToList
      Get
        Return MyFilterPLL.ToList
      End Get
    End Property

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

    Public ReadOnly Property ToListOfPriceMedianNextDayHigh As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayHigh
      Get
        Return MyListOfPriceMedianNextDayHigh
      End Get
    End Property

    Public ReadOnly Property ToListOfPriceMedianNextDayLow As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayLow
      Get
        Return MyListOfPriceMedianNextDayLow
      End Get
    End Property
  End Class
End Namespace