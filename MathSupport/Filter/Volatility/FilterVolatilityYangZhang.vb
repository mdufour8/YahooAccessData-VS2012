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
#Region "FilterVolatilityYangZhang"
  ''' <summary>
  ''' This class implement the Yang-Zhang volatility measurements
  ''' see definition:
  ''' http://en.wikipedia.org/wiki/Rate_of_return#Logarithmic_or_continuously_compounded_return
  ''' http://en.wikipedia.org/wiki/Volatility_%28finance%29
  ''' https://www.youtube.com/watch?v=eiTCTibH010
  ''' http://en.wikipedia.org/wiki/Volatility_(finance)
  ''' https://en.wikipedia.org/wiki/Stochastic_volatility
  ''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
  ''' </summary>
  ''' <remarks>
  ''' In 2000 Yang-Zhang created the most powerful volatility
  ''' measure that handles both opening jumps and drift. It is the sum of the overnight
  ''' volatility (close to open volatility) and a weighted average of the Rogers-Satchell
  ''' volatility and the open to close volatility. The assumption of continuous prices does
  ''' mean the measure tends to slightly underestimate the volatility. The class use 
  ''' the normalized logarithm compounded return to measure the volatility and
  ''' is only valid for positive value of signal
  ''' </remarks>
  <Serializable()>
  Public Class FilterVolatilityYangZhang
    Implements IFilter
    Implements IRegisterKey(Of String)

    Public Enum enuVolatilityDailyPeriodType
      FullDay
      PreviousCloseToOpen
      OpenToClose
      OpenToHighClose
      OpenToLowClose
      OpenToHighToLowCloseRatio
			OpenToHighToLowCloseRatioFiltered
			Rogers_Satchell_Yoon_Vrs
			Parkison_Vp
		End Enum

    Private Const FILTER_RATE_DEFAULT As Integer = MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
    Private MyFilterDirection As FilterRSI.SlopeDirection
    Private MyRate As Integer
    Private MyFilterValueLastK1 As Double
    Private MyFilterValueLast As Double
    Private MyValueLast As YahooAccessData.IPriceVol
    Private MyValueLastK1 As YahooAccessData.IPriceVol
    Private MyFilterVolatilityYearlyCorrection As Double
    Private MyListOfPreviousCloseToOpenHighLowClose As List(Of Double)
    Private MyListOfPreviousCloseToOpen As List(Of Double)
    Private MyListOfOpenToClose As List(Of Double)
    Private MyListOfOpenHighAsClose As List(Of Double)
    Private MyListOfOpenLowAsClose As List(Of Double)
    Private MyListOfOpenToHighToLowAsCloseRatio As List(Of Double)
    Private MyFilterOfOpenToHighToLowAsCloseRatio As Filter.FilterLowPassPLL

    'Private MyCountOfVolNotNull As Integer
    Private MyReturnLogForOpenToPreviousClose As Double
    Private MyReturnLogForCloseToOpen As Double
    Private MyReturnLogForHighToOpen As Double
    Private MyReturnLogForLowToOpen As Double
    Private MyReturnLogForHighToPreviousClose As Double
    Private MyReturnLogForLowToPreviousClose As Double

    'Private MyReturnLogForCloseToHigh As Double
    'Private MyReturnLogForCloseToLow As Double

    Private IsUseLastSampleHighLowTrailLocal As Boolean

    Private MyStatisticalForOpen As IFilter(Of IStatistical)
    Private MyStatisticalForClose As IFilter(Of IStatistical)
    Private MyStatisticalForOpenToHigh As IFilter(Of IStatistical)
    Private MyStatisticalForOpenToLow As IFilter(Of IStatistical)
    Private MyStatisticalForPreviousCloseToHigh As IFilter(Of IStatistical)
    Private MyStatisticalForPreviousCloseToLow As IFilter(Of IStatistical)
    Private MyFilterExpForPositiveVariance As Filter.FilterLowPassExp
    Private MyFilterExpForNegativeVariance As Filter.FilterLowPassExp

		'Rogers-Satchell is an estimator for measuring the volatility of securities
		'with an average return not equal to zero. Unlike Parkinson and Garman-Klass estimators,
		'Rogers-Satchell incorporates a drift term (mean return not equal to zero).2022
		Private MyStatisticalForVRSHighAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyStatisticalForVRSLowAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSHigh As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSLow As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSTotal As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyValueForK As Double
    Private MyStatisticType As FilterVolatility.enuVolatilityStatisticType
    Private MyVariancePositifSumLast As Double
    Private MyVarianceNegatifSumLast As Double
    Private MyFilterOfVolatilityPositif As Filter.FilterLowPassPLL
    Private MyFilterOfVolatilityNegatif As Filter.FilterLowPassPLL
    Private MyPriceNextDailyHighPreviousCloseToOpenSigma2 As Double


    Public Sub New()
      Me.New(FILTER_RATE_DEFAULT, Math.Sqrt(NUMBER_TRADINGDAY_PER_YEAR))
    End Sub

    Public Sub New(ByVal FilterRate As Integer, Optional StatisticType As FilterVolatility.enuVolatilityStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard, Optional IsUseLastSampleHighLowTrail As Boolean = False)
      Me.New(FilterRate, Math.Sqrt(NUMBER_TRADINGDAY_PER_YEAR), StatisticType, IsUseLastSampleHighLowTrail)
    End Sub

    Public Sub New(
      ByVal FilterRate As Integer,
      ByVal ScaleCorrection As Double,
      Optional StatisticType As FilterVolatility.enuVolatilityStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard,
      Optional IsUseLastSampleHighLowTrail As Boolean = False)

      IsUseLastSampleHighLowTrailLocal = IsUseLastSampleHighLowTrail
      MyFilterVolatilityYearlyCorrection = ScaleCorrection
      MyStatisticType = StatisticType
      MyListOfPreviousCloseToOpenHighLowClose = New List(Of Double)
      MyListOfPreviousCloseToOpen = New List(Of Double)
      MyListOfOpenToClose = New List(Of Double)
      MyListOfOpenHighAsClose = New List(Of Double)
      MyListOfOpenLowAsClose = New List(Of Double)
      MyListOfOpenToHighToLowAsCloseRatio = New List(Of Double)
      MyFilterOfOpenToHighToLowAsCloseRatio = New Filter.FilterLowPassPLL(FilterRate:=FilterRate)

      If FilterRate < 2 Then FilterRate = 2
      MyRate = CInt(FilterRate)
      'MyStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard
      Select Case MyStatisticType
        Case FilterVolatility.enuVolatilityStatisticType.Exponential
					MyStatisticalForOpen = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForOpenToLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForOpenToHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForPreviousCloseToHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForPreviousCloseToLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSTotal = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSHighAsClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSLowAsClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				Case Else
					'the default is a windows statistic
					MyStatisticalForOpen = New FilterStatistical(FilterRate)
          MyStatisticalForClose = New FilterStatistical(FilterRate)
          MyStatisticalForOpenToLow = New FilterStatistical(FilterRate)
          MyStatisticalForOpenToHigh = New FilterStatistical(FilterRate)
          MyStatisticalForPreviousCloseToHigh = New FilterStatistical(FilterRate)
          MyStatisticalForPreviousCloseToLow = New FilterStatistical(FilterRate)
          MyStatisticalForVRSTotal = New FilterStatistical(FilterRate)
          MyStatisticalForVRSHigh = New FilterStatistical(FilterRate)
          MyStatisticalForVRSLow = New FilterStatistical(FilterRate)
          MyStatisticalForVRSHighAsClose = New FilterStatistical(FilterRate)
          MyStatisticalForVRSLowAsClose = New FilterStatistical(FilterRate)
      End Select
      MyFilterExpForPositiveVariance = New Filter.FilterLowPassExp(FilterRate)
      MyFilterExpForNegativeVariance = New Filter.FilterLowPassExp(FilterRate)

      'see paper
      MyValueForK = 0.34 / (1.34 + ((MyRate + 1) / (MyRate - 1)))

      MyFilterValueLast = 0
      MyFilterValueLastK1 = 0
      MyValueLast = New YahooAccessData.PriceVol(0)
      MyValueLastK1 = MyValueLast
      MyFilterDirection = FilterRSI.SlopeDirection.Zero
    End Sub

    ''' <summary>
    ''' True for ignoring the volatility jump due to an open price exceeding the expected 2x sigma price peak value.
    ''' This can be usuful for reducing the impact of unexpected news on the stock volatility. 
    ''' </summary>
    ''' <returns>The current state</returns>
    Public Property IsFilterVolatilityJump As Boolean

    Public Function Filter(ByVal Value As YahooAccessData.IPriceVol, ByVal IsVolatityHoldToLast As Boolean) As Double
      Dim ThisResult As Double
      If Value.IsSpecialDividendPayout Or IsVolatityHoldToLast Then
				ThisResult = Me.CalculateFilterLocal(MyValueLast, False)
				'restore the correct value for the last parameters
				MyValueLast = Value
        MyValueLastK1 = MyValueLast
      Else
        ThisResult = Me.CalculateFilterLocal(Value, False)
      End If
      Return ThisResult
    End Function
    Private Function CalculateFilterLocal(ByRef Value As YahooAccessData.IPriceVol, ByVal IsVolatityHoldToLast As Boolean) As Double
      Dim ThisOpenToHighToLowAsCloseRatio As Double
      Dim ThisVRSTotalOpenToHighLow As Double
      Dim ThisVRSPartialOpenToHigh As Double
      Dim ThisVRSPartialOpenToLow As Double
      Dim ThisVRSTotalMeanForOpenToHighLow_Vrs As Double
      Dim ThisVarianceForPreviousCloseToOpen_Vo As Double
      Dim ThisVarianceForOpenToClose_Vc As Double
      Dim ThisVariancePreviousCloseToOpenHighLowClose_V As Double
      Dim ThisVarianceOpenToHighLowClose_KValue As Double
      Dim ThisValueLow As Single
      Dim ThisValueHigh As Single
      Dim ThisVRSMeanUp As Double
      Dim ThisVRSMeanDown As Double
      Dim ThisVarianceForOpenToHigh_Vh As Double
      Dim ThisVarianceForOpenToLow_Vl As Double
      Dim ThisVariancePositifSum As Double
      Dim ThisVarianceNegatifSum As Double
      Dim ThisVRSUp2 As Double
      Dim ThisVRSDown2 As Double


      ThisValueLow = Value.Low
      ThisValueHigh = Value.High
      If MyListOfPreviousCloseToOpenHighLowClose.Count = 0 Then
        'measure volatility from initial value only at start
        MyFilterValueLast = 0
        MyValueLast = New PriceVol(Value.Open)
        MyReturnLogForOpenToPreviousClose = LogPriceReturn(Value.Open, MyValueLast.Last)
        MyReturnLogForLowToOpen = LogPriceReturn(ThisValueLow, Value.Open)
        'MyReturnLogForLowToOpen = LogPriceReturn(Value.Open, ThisValueLow)
        MyReturnLogForHighToOpen = LogPriceReturn(ThisValueHigh, Value.Open)
        MyReturnLogForCloseToOpen = LogPriceReturn(Value.Last, Value.Open)
        IsVolatityHoldToLast = False  'always false for the first data
      End If
      If IsUseLastSampleHighLowTrailLocal Then
        If MyValueLast.Low < ThisValueLow Then
          ThisValueLow = MyValueLast.Low
        End If
        If MyValueLast.High > ThisValueHigh Then
          ThisValueHigh = MyValueLast.High
        End If
      End If
      'filter for value less than zero
      'If DirectCast(Value, PriceVol).IsNull = False Then
      '        MyCountOfVolNotNull = MyCountOfVolNotNull + 1
      'End If

      'ln(Open1/Close0)
      MyReturnLogForOpenToPreviousClose = LogPriceReturn(Value.Open, MyValueLast.Last)
      'ln(Low1/Open1)
      MyReturnLogForLowToOpen = LogPriceReturn(ThisValueLow, Value.Open)
      'ln(High1/Open1)
      MyReturnLogForHighToOpen = LogPriceReturn(ThisValueHigh, Value.Open)
      'ln(Close1/Open1)
      MyReturnLogForCloseToOpen = LogPriceReturn(Value.Last, Value.Open)

      Dim ThisReturnLogForPreviousHighToOpen = LogPriceReturn(Value.Open, MyValueLast.High)
      Dim ThisReturnLogForPreviousLowToOpen = LogPriceReturn(Value.Open, MyValueLast.Low)

      'VRS calculation
      'ThisVRSUp2 = (ln(High/Open)^ 2) + ((ln(Last/Low) ^ 2))


      'this is a local calcul to measure and compare the volatility Up and down of the stock
      'it is an intraday calculation that is very predictive and related to the comportment of the stock in the future
      'but it is not related and is not needed for the VolatilityYangZhang calculation
      'Note that some observation show the opening to be a bit more predictive than the closing
      'that is the reason we divide the close by 2
      'however other trader seem to indicate that the close is more predictive
      'an investigation is needed to clear this aspect
      'to do: calculate both for testing in the future
      'ThisVRSUp2 = (MyReturnLogForHighToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.Low) ^ 2) / 2)
      'ThisVRSDown2 = (MyReturnLogForLowToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.High) ^ 2) / 2)
      'modified mars 2024 taking into account the previous close
      ThisVRSUp2 = ((ThisReturnLogForPreviousLowToOpen ^ 2) / 2) + (MyReturnLogForHighToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.Low) ^ 2) / 2)
      ThisVRSDown2 = ((ThisReturnLogForPreviousHighToOpen ^ 2) / 2) + (MyReturnLogForLowToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.High) ^ 2) / 2)

      ThisVRSPartialOpenToHigh = MyReturnLogForHighToOpen * (MyReturnLogForHighToOpen - MyReturnLogForCloseToOpen)
      ThisVRSPartialOpenToLow = MyReturnLogForLowToOpen * (MyReturnLogForLowToOpen - MyReturnLogForCloseToOpen)

      MyReturnLogForHighToPreviousClose = LogPriceReturn(ThisValueHigh, MyValueLast.Last)
      MyReturnLogForLowToPreviousClose = LogPriceReturn(ThisValueLow, MyValueLast.Last)

      'If Me.Count = 1576 Then
      'ThisVarianceNegatifSum = ThisVarianceNegatifSum
      'End If

      'ThisVRSPartialOpenToHigh = MyReturnLogForHighToOpen * (MyReturnLogForHighToOpen - 0)
      'ThisVRSPartialOpenToLow = MyReturnLogForLowToOpen * (MyReturnLogForLowToOpen - 0)

      ThisVRSTotalOpenToHighLow = ThisVRSPartialOpenToHigh + ThisVRSPartialOpenToLow

      'calculate the variance for the open and close
      'This is Vo in the ref. paper
      ThisVarianceForPreviousCloseToOpen_Vo = MyStatisticalForOpen.Filter(MyReturnLogForOpenToPreviousClose).Variance
      'This is Vc in the ref. paper
      ThisVarianceForOpenToClose_Vc = MyStatisticalForClose.Filter(MyReturnLogForCloseToOpen).Variance
      ThisVarianceForOpenToHigh_Vh = MyStatisticalForOpenToHigh.Filter(MyReturnLogForHighToOpen).Variance
      ThisVarianceForOpenToLow_Vl = MyStatisticalForOpenToLow.Filter(MyReturnLogForLowToOpen).Variance

      'not used right now
      'ThisVarianceForPreviousCloseToHigh_Vh = MyStatisticalForOpenToLow.Filter(MyReturnLogForHighToPreviousClose).Variance
      'ThisVarianceForPreviousCloseToLow_Vl = MyStatisticalForOpenToLow.Filter(MyReturnLogForLowToPreviousClose).Variance

      'calculate le mean for the total high and low variation to close
      'this is Vrs in the ref. paper
      ThisVRSTotalMeanForOpenToHighLow_Vrs = MyStatisticalForVRSTotal.Filter(ThisVRSTotalOpenToHighLow).Mean

      ThisVarianceOpenToHighLowClose_KValue = MyValueForK * ThisVarianceForOpenToClose_Vc + (1 - MyValueForK) * ThisVRSTotalMeanForOpenToHighLow_Vrs
      'this is the final volatility
      ThisVariancePreviousCloseToOpenHighLowClose_V = ThisVarianceForPreviousCloseToOpen_Vo + ThisVarianceOpenToHighLowClose_KValue

      'separate the variance in positif and negatif value for calculation of the PriceVolatilityPositif measurement
      'ThisVariancePositifSum = (1 - MyValueForK) * (ThisVRSMeanUp / ThisVariancePreviousCloseToOpenHighLowClose_V)
      'ThisVarianceNegatifSum = (1 - MyValueForK) * (ThisVRSMeanDown / ThisVariancePreviousCloseToOpenHighLowClose_V)
      ThisVRSMeanUp = MyStatisticalForVRSHigh.Filter(ThisVRSUp2).Mean
      ThisVRSMeanDown = MyStatisticalForVRSLow.Filter(ThisVRSDown2).Mean
      ThisVariancePositifSum = (1 - MyValueForK) * ThisVRSMeanUp
      ThisVarianceNegatifSum = (1 - MyValueForK) * ThisVRSMeanDown
      'correct the value for the yearly variation

      MyFilterValueLastK1 = MyFilterValueLast
      MyFilterValueLast = ToYearCorrected(ThisVariancePreviousCloseToOpenHighLowClose_V)
      MyListOfPreviousCloseToOpenHighLowClose.Add(MyFilterValueLast)
      MyListOfPreviousCloseToOpen.Add(ToYearCorrected(ThisVarianceForPreviousCloseToOpen_Vo))
      MyListOfOpenToClose.Add(ToYearCorrected(ThisVarianceOpenToHighLowClose_KValue))

      '~~~~~~~~~~~~~~~
      'the filter is not used right now
      MyFilterExpForPositiveVariance.Filter(ThisVariancePositifSum)
      MyFilterExpForNegativeVariance.Filter(ThisVarianceNegatifSum)
      MyListOfOpenHighAsClose.Add(ToYearCorrected(ThisVariancePositifSum))
      MyListOfOpenLowAsClose.Add(ToYearCorrected(ThisVarianceNegatifSum))
      Dim ThisSumOfVolatilityPositifNegatif = ThisVariancePositifSum + ThisVarianceNegatifSum
      'the balance value of this ratio is 0.5
      If ThisSumOfVolatilityPositifNegatif > 0 Then
        ThisOpenToHighToLowAsCloseRatio = ThisVariancePositifSum / ThisSumOfVolatilityPositifNegatif
      Else
        ThisOpenToHighToLowAsCloseRatio = 0.5
      End If
      If MyListOfOpenToHighToLowAsCloseRatio.Count > 0 Then
        'this is much better and predictive
        Select Case ThisOpenToHighToLowAsCloseRatio
          Case > 0.5
            MyFilterDirection = FilterRSI.SlopeDirection.Positive
          Case < 0.5
            MyFilterDirection = FilterRSI.SlopeDirection.Negative
          Case Else
            MyFilterDirection = FilterRSI.SlopeDirection.Zero
        End Select
        'this method below is not predictive and is rather useless
        'Select Case MyListOfOpenToHighToLowAsCloseRatio.Last
        '  Case < ThisOpenToHighToLowAsCloseRatio
        '    MyFilterDirection = FilterRSI.SlopeDirection.Positive
        '  Case > ThisOpenToHighToLowAsCloseRatio
        '    MyFilterDirection = FilterRSI.SlopeDirection.Negative
        '  Case Else
        '    MyFilterDirection = FilterRSI.SlopeDirection.Zero
        'End Select
      Else
        MyFilterDirection = FilterRSI.SlopeDirection.Zero
      End If
      MyListOfOpenToHighToLowAsCloseRatio.Add(ThisOpenToHighToLowAsCloseRatio)
      MyFilterOfOpenToHighToLowAsCloseRatio.Filter(ThisOpenToHighToLowAsCloseRatio)
      '~~~~~~~~~~~~~~~

      MyValueLastK1 = MyValueLast
      MyValueLast = Value
      Return MyFilterValueLast
    End Function

    '''' <summary>
    '''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
    '''' This method expect positive value of asset price and use the YangZhang method

    '''' </summary>
    '''' <param name="Value">The current positive value of the asset</param>
    '''' <remarks>
    '''' The function assume by default a daily data input and scale the result to a yearly period. The scale factor may need to be adjusted if the data
    '''' is not at the daily sample rate or another type of volatility is needed.
    '''' </remarks>
    '''' 


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    Public Function Filter(ByVal Value As YahooAccessData.IPriceVol) As Double Implements IFilter.Filter
      Return Me.Filter(Value, IsVolatityHoldToLast:=False)
    End Function

    'Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
    '  Dim ThisValue As YahooAccessData.IPriceVol
    '  For Each ThisValue In Value
    '    Me.Filter(ThisValue)
    '  Next
    '  Return Me.ToArray
    'End Function

    Public Function FilterLast() As Double Implements IFilter.FilterLast
      Return MyFilterValueLast
    End Function

    Public Function FilterDirection() As FilterRSI.SlopeDirection
      Return MyFilterDirection
    End Function

    Public Function Last() As Double Implements IFilter.Last
      Return MyValueLast.Last
    End Function

    Public ReadOnly Property Rate As Integer Implements IFilter.Rate
      Get
        Return MyRate
      End Get
    End Property

    Public ReadOnly Property Count As Integer Implements IFilter.Count
      Get
        Return MyListOfPreviousCloseToOpenHighLowClose.Count
      End Get
    End Property

    Public ReadOnly Property Max As Double Implements IFilter.Max
      Get
        Return MyListOfPreviousCloseToOpenHighLowClose.Max
      End Get
    End Property

    Public ReadOnly Property Min As Double Implements IFilter.Min
      Get
        Return MyListOfPreviousCloseToOpenHighLowClose.Min
      End Get
    End Property

    Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
      Get
        Return MyListOfPreviousCloseToOpenHighLowClose
      End Get
    End Property

    Public ReadOnly Property ToList(ByVal Type As enuVolatilityDailyPeriodType) As IList(Of Double)
      Get
        Select Case Type
          Case enuVolatilityDailyPeriodType.FullDay
            Return MyListOfPreviousCloseToOpenHighLowClose
          Case enuVolatilityDailyPeriodType.OpenToClose
            Return MyListOfOpenToClose
          Case enuVolatilityDailyPeriodType.PreviousCloseToOpen
            Return MyListOfPreviousCloseToOpen
          Case enuVolatilityDailyPeriodType.OpenToHighClose
            Return MyListOfOpenHighAsClose
          Case enuVolatilityDailyPeriodType.OpenToLowClose
            Return MyListOfOpenLowAsClose
          Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio
            Return MyListOfOpenToHighToLowAsCloseRatio
          Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatioFiltered
            Return MyFilterOfOpenToHighToLowAsCloseRatio.ToList
          Case Else
            Return MyListOfPreviousCloseToOpenHighLowClose
        End Select
      End Get
    End Property

    Private ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
      Get
        Throw New NotImplementedException
      End Get
    End Property

    Public Function ToArray() As Double() Implements IFilter.ToArray
      Return MyListOfPreviousCloseToOpenHighLowClose.ToArray
    End Function

    Private Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Throw New NotImplementedException
    End Function

    Private Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Throw New NotImplementedException
    End Function

    Public Property Tag As String Implements IFilter.Tag

    Public Overrides Function ToString() As String Implements IFilter.ToString
      Return Me.FilterLast.ToString
    End Function

    Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
      Dim ThisValue As Double
      For Each ThisValue In Value
        Me.Filter(ThisValue)
      Next
      Return Me.ToArray
    End Function

    Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
      Throw New NotImplementedException
    End Function

    Public Function Filter(Value As Double) As Double Implements IFilter.Filter
      Return Me.Filter(New PriceVol(CSng(Value)))
    End Function

    Public Function Filter(Value As Single) As Double Implements IFilter.Filter
      Return Me.Filter(New PriceVol(Value))
    End Function

    Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
      Throw New NotImplementedException
    End Function

    Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
      Throw New NotImplementedException
    End Function

    Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
      Throw New NotImplementedException
    End Function

    Public Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
      Throw New NotImplementedException
    End Function

    Public Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
      Throw New NotImplementedException
    End Function

    Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
      Return MyValueLast
    End Function

    Public ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of Double) Implements IFilter.ToListOfError
      Get
        Throw New NotImplementedException
      End Get
    End Property

    Private Function ToYearCorrected(ByVal VolatilityPerDay As Double) As Double
      Return MyFilterVolatilityYearlyCorrection * Math.Sqrt(VolatilityPerDay)
    End Function
#Region "IRegisterKey"
    Public Function AsIRegisterKey() As IRegisterKey(Of String)
      Return Me
    End Function
    Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
    Dim MyKeyValue As String
    Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
      Get
        Return MyKeyValue
      End Get
      Set(value As String)
        MyKeyValue = value
      End Set
    End Property
#End Region
  End Class
#End Region
End Namespace

