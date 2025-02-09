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
#Region "FilterVolatility"
  ''' <summary>
  ''' This class implement the close to close standard volatility measurements.
  ''' 
  ''' see definition:
  ''' http://en.wikipedia.org/wiki/Rate_of_return#Logarithmic_or_continuously_compounded_return
  ''' http://en.wikipedia.org/wiki/Volatility_%28finance%29
  ''' https://www.youtube.com/watch?v=eiTCTibH010
  ''' http://en.wikipedia.org/wiki/Volatility_(finance)
  ''' https://en.wikipedia.org/wiki/Stochastic_volatility
  ''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
  ''' </summary>
  ''' <remarks>
  ''' The simplest and most common type of calculation that
  ''' benefits from only using reliable prices from closing auctions. We note that the
  ''' volatility should be the standard deviation multiplied by √N/(N-1) to take into
  ''' account the fact we are sampling a smaller subset of the population.
  ''' This class also use the normalized logarithm compounded return to measure the volatility and
  ''' is only valid for positive value
  ''' </remarks>
  <Serializable()>
  Public Class FilterVolatility
    Implements IFilter
    Implements IRegisterKey(Of String)

    Public Enum enuVolatilityStatisticType
      Standard
      Exponential
    End Enum


    Public Const VOLATILITY_FILTER_RATE_DEFAULT As Integer = MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
    Private MyRate As Integer
    Private FilterValueLastK1 As Double
    Private FilterValueLast As Double
    Private ValueLast As Double
    Private ValueLastK1 As Double
    Private MyFilterVolatilityYearlyCorrection As Double
    Private MyListOfValue As ListScaled
    Private MyStatistical As IFilter(Of IStatistical)
    Private MyStatisticType As enuVolatilityStatisticType
		'Private MyPriceNextDailyHighPreviousCloseToOpenSigma2 As Double
		Private IsSpecialDividendPayoutLocal As Boolean

    Public Sub New()
      Me.New(VOLATILITY_FILTER_RATE_DEFAULT, Math.Sqrt(NUMBER_TRADINGDAY_PER_YEAR))
    End Sub

    Public Sub New(ByVal FilterRate As Integer, Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard, Optional SamplingRatePerDay As Double = 1.0)
      Me.New(FilterRate, Math.Sqrt(SamplingRatePerDay * NUMBER_TRADINGDAY_PER_YEAR), StatisticType)
    End Sub

    Public Sub New(
      ByVal FilterRate As Integer,
      ByVal ScaleCorrection As Double,
      Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard)

      MyFilterVolatilityYearlyCorrection = ScaleCorrection
      MyListOfValue = New ListScaled
      MyStatisticType = StatisticType
      If FilterRate < 1 Then FilterRate = 1
      MyRate = CInt(FilterRate)
      Select Case MyStatisticType
        Case enuVolatilityStatisticType.Exponential
          MyStatistical = New FilterStatisticalExp(FilterRate)
        Case Else
          MyStatistical = New FilterStatistical(FilterRate)
      End Select
      FilterValueLast = 0
      FilterValueLastK1 = 0
      ValueLast = 0
      ValueLastK1 = 0
      IsSpecialDividendPayoutLocal = False
    End Sub


    ''' <summary>
    ''' True for ignoring the volatility jump due to an open price exceeding the expected 2x sigma price peak value.
    ''' This can be usuful for reducing the impact of unexpected news on the stock volatility. 
    ''' </summary>
    ''' <returns>The current state</returns>
    Public Property IsFilterVolatilityJump As Boolean

    ''' <summary>
    ''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
    ''' This method expect positive value of asset price.
    ''' </summary>
    ''' <param name="Value">The current positive value of the asset</param>
    ''' <param name="ValueRef">
    ''' The reference value for the asset return calculation normally the last value of the sample. This function may likely require to adjust
    ''' the scale factor to unity for raw volatility measurement
    ''' </param>
    ''' <remarks>
    ''' The function assume by default a daily data input. The scale factor may need to be adjusted if the data
    ''' is not at the daily sample rate and the yearly volatility is needed.
    ''' </remarks>
    Public Function Filter(ByVal Value As Double, ByVal ValueRef As Double) As Double
      Dim ThisReturnLog As Double

      If MyListOfValue.Count = 0 Then
        'assume volatility of zero at start
        FilterValueLast = 0
      Else
        If IsFilterVolatilityJump Then
          If FilterValueLast > 0 Then
            'nothing to calculate if the volatility is zero
            'calculate the 2 sigma range for the current stock and volatility
            Dim ThisPriceNextDailyHighPreviousCloseToClose = StockOption.StockPricePrediction(
              NumberTradingDays:=1,
              Me.ValueLast,
              Gain:=0.0,
              GainDerivative:=0.0,
              Me.FilterValueLast,
              GAUSSIAN_PROBABILITY_SIGMA3)

            If Value > ThisPriceNextDailyHighPreviousCloseToClose Then
              ThisPriceNextDailyHighPreviousCloseToClose = ThisPriceNextDailyHighPreviousCloseToClose
            End If
          End If
        End If
      End If
      If IsSpecialDividendPayoutLocal Then
        'ignore the current data and use the previous calculation
        IsSpecialDividendPayoutLocal = False
        ThisReturnLog = MyStatistical.Last
      Else
				'same thing shoudl be fixed
				'ThisReturnLog = GainLog(Value, ValueRef)  
				ThisReturnLog = LogPriceReturn(Value, ValueRef)
			End If
      MyStatistical.Filter(ThisReturnLog)
      FilterValueLastK1 = FilterValueLast
      'correct the value for the yearly variation
      FilterValueLast = MyFilterVolatilityYearlyCorrection * MyStatistical.FilterLast.StandardDeviation
      MyListOfValue.Add(FilterValueLast)
      'calculate the next sample 2 sigma normal last price range
      If MyListOfValue.Count > 0 Then
        If IsFilterVolatilityJump Then
					'MyPriceNextDailyHighPreviousCloseToOpenSigma2 = OptionValuation.StockOption.StockPricePrediction(
					'  NumberTradingDays:=TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY,
					'  StockPrice:=Value,
					'  Gain:=0.0,
					'  GainDerivative:=0.0,
					'  Volatility:=FilterValueLast,
					'  Probability:=GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2)
				End If
      End If
      ValueLastK1 = ValueLast
      ValueLast = Value
      Return FilterValueLast
    End Function

    ''' <summary>
    ''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
    ''' This method expect positive value of asset price.
    ''' </summary>
    ''' <param name="Value">The current positive value of the asset</param>
    ''' <returns>The current volatility corrected by default to a yearly period assuming a daily sample rate.</returns>
    ''' <remarks>
    ''' The function assume by default a daily data input. The scale factor may need to be adjusted if the data
    ''' is not at the daily sample rate and the yearly volatility is needed.
    ''' </remarks>
    Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
      Return Me.Filter(Value, ValueLast)
    End Function


    Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
      IsSpecialDividendPayoutLocal = Value.IsSpecialDividendPayout
      Return Me.Filter(CDbl(Value.Last))
    End Function

    ' ''' <summary>
    ' ''' Compute the yearly Volatility based on the Garman and Klauss estimator taking in account the high and low
    ' ''' </summary>
    ' ''' <param name="Value"></param>
    ' ''' <param name="ValueHigh"></param>
    ' ''' <param name="ValueLow"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function Filter(ByVal Value As YahooAccessData.IPriceVol) As Double
    '  Dim ThisReturnLog As Double
    '  Dim ThisReturnLogHighLow As Double


    '  Throw New NotImplementedException
    '  'If MyListOfValue.Count = 0 Then
    '  '  FilterValueLast = Value.Last
    '  'End If
    '  'If ValueLast <= 0 Then
    '  '  ThisReturnLog = 0
    '  '  ThisReturnLogHighLow = 0
    '  'Else
    '  '  If Value <= 0 Then
    '  '    ThisReturnLog = 0
    '  '    ThisReturnLogHighLow = 0
    '  '  Else
    '  '    ThisReturnLog = Math.Log(Value / ValueLast)
    '  '  End If
    '  'End If
    '  'MyStatistical.Filter(ThisReturnLog)
    '  'FilterValueLastK1 = FilterValueLast
    '  ''correct the value for the yearly variation
    '  'FilterValueLast = MyFilterVolatilityYearlyCorrection * MyStatistical.FilterLast.StandardDeviation
    '  'MyListOfValue.Add(FilterValueLast)
    '  'ValueLastK1 = ValueLast
    '  'ValueLast = Value
    '  'Return FilterValueLast
    'End Function

    Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
      Dim ThisValue As Double
      For Each ThisValue In Value
        Me.Filter(ThisValue)
      Next
      Return Me.ToArray
    End Function

    Public Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
      Throw New NotSupportedException
    End Function

    Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
      Throw New NotSupportedException
    End Function

    Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
      Throw New NotSupportedException
    End Function

    Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
      Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
      With ThisPriceVol
        .LastPrevious = CSng(FilterValueLastK1)
        If Me.FilterLast > .Last Then
          .High = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        ElseIf Me.Last < .Last Then
          .Low = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        End If
      End With
      Return ThisPriceVol
    End Function

    Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
      Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
      With ThisPriceVol
        .LastPrevious = CSng(ValueLastK1)
        If Me.FilterLast > .Last Then
          .High = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        ElseIf Me.FilterLast < .Last Then
          .Low = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        End If
      End With
      Return ThisPriceVol
    End Function

    Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
      Return CSng(Me.Filter(CDbl(Value)))
    End Function

    Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
      Throw New NotSupportedException
    End Function

    Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
      Return Me.FilterPredictionNext(CDbl(Value))
    End Function

    Public Function FilterLast() As Double Implements IFilter.FilterLast
      Return FilterValueLast
    End Function

    Public Function Last() As Double Implements IFilter.Last
      Return ValueLast
    End Function

    Public ReadOnly Property Rate As Integer Implements IFilter.Rate
      Get
        Return MyRate
      End Get
    End Property

    Public ReadOnly Property Count As Integer Implements IFilter.Count
      Get
        Return MyListOfValue.Count
      End Get
    End Property

    Public ReadOnly Property Max As Double Implements IFilter.Max
      Get
        Return MyListOfValue.Max
      End Get
    End Property

    Public ReadOnly Property Min As Double Implements IFilter.Min
      Get
        Return MyListOfValue.Min
      End Get
    End Property

    Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
      Get
        Return MyListOfValue
      End Get
    End Property

    private ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
      Get
        Throw New NotSupportedException
      End Get
    End Property

    Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
      Get
        Return MyListOfValue
      End Get
    End Property

    Public Function ToArray() As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray
    End Function

    Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Property Tag As String Implements IFilter.Tag

    Public Overrides Function ToString() As String Implements IFilter.ToString
      Return Me.FilterLast.ToString
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
