Imports MathNet.Numerics

Namespace MathProcess
  ''' <summary>
  ''' see: https://en.wikipedia.org/wiki/Wiener_process
  ''' </summary>
  ''' <remarks>
  ''' The Wiener process Wt is characterised by the following properties:[1]
  ''' 1. Wt(0)   = 0   
  ''' 2. Wt(t) has independent increments for every t > 0 and the future increments Wt(t+u)− W(t) u ≥ 0, are independent of the past values
  ''' 3. Wt(t+u) has Gaussian increments: Wt(t + u) − Wt(t) is normally distributed with mean 0 and variance u.   
  ''' 4. Wt has continuous paths with probability 1 and is continuous in t
  ''' </remarks>
  Public Class WeinerProcess
    Private MyGaussian As Distributions.Normal
    Private MyWeinerProcessValueLast As Double
    Private MyCount As Integer


    Public Sub New(ByVal Variance As Double)
      MyGaussian = New Distributions.Normal(mean:=0.0, stddev:=Math.Sqrt(Variance))
      MyWeinerProcessValueLast = 0.0
    End Sub

    Public Function Sample() As Double
      MyWeinerProcessValueLast = MyWeinerProcessValueLast + MyGaussian.Sample
      MyCount = MyCount + 1
      Return MyWeinerProcessValueLast
    End Function

    Public ReadOnly Property SampleLast() As Double
      Get
        Return MyWeinerProcessValueLast
      End Get
    End Property

    Public Function Samples(ByVal NumberPoints As Integer) As Double()
      Dim ThisArray(0 To NumberPoints - 1) As Double
      Dim I As Integer

      ThisArray(0) = MyWeinerProcessValueLast
      For I = 1 To NumberPoints - 1
        ThisArray(I) = Me.Sample
      Next
      Return ThisArray
    End Function

    Public ReadOnly Property Count As Integer
      Get
        Return MyCount
      End Get
    End Property

    Public Sub Reset()
      MyWeinerProcessValueLast = 0
    End Sub
  End Class


  ''' <summary>
  ''' Simulate a stochastic process that represent the solution do a Geometric Brownian Motion of a stock price variation.
  ''' For details see PDF file:
  ''' C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData VS2012\MathSupport\Geometric_Brownian_motion.pdf
  ''' See equation 3 in this papier:
  ''' C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData VS2012\MathSupport\Geometric Brownian Motion Option Pricing and Simulation.pdf
  ''' 
  ''' or:
  ''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
  ''' </summary>
  ''' <remarks></remarks>
  Public Class StochasticProcess
    Private Const TIME_DAILY_PERIOD_IN_YEAR As Double = 1 / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
    Private Const TIME_DAILY_FROM_LAST_TO_OPEN_PERIOD_IN_YEAR As Double = 2 / 3 * TIME_DAILY_PERIOD_IN_YEAR
    Private Const TIME_DAILY_FROM_OPEN_TO_LAST_PERIOD_IN_YEAR As Double = 1 / 3 * TIME_DAILY_PERIOD_IN_YEAR


    Private MyGaussian As Distributions.Normal
    Private MyStochasticProcessLast As Double
    Private MyValueStart As Double
    Private MyVolatility As Double
    Private MyGain As Double
    Private MyMuSigma As Double
    Private MyCount As Integer
    Private MySquareRootOfTimeDailyPeriod As Double

    ''' <summary>
    ''' Create the solution for a Stochastic process representing the daily variation of a process.
    ''' Each sample represent the daily expected variation of the random process
    ''' </summary>
    ''' <param name="ValueStart">The inital value of the process</param>
    ''' <param name="Gain">The gain of the process per year</param>
    ''' <param name="Volatility">The volatility of the process per year''' </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal ValueStart As Double, ByVal Gain As Double, ByVal Volatility As Double)
			MyGaussian = New Distributions.Normal(0, 1)
			MyValueStart = ValueStart
      MyStochasticProcessLast = ValueStart
      MyVolatility = Volatility
      MyMuSigma = Gain - (Volatility ^ 2 / 2)
      MyCount = 0
      MySquareRootOfTimeDailyPeriod = Math.Sqrt(TIME_DAILY_PERIOD_IN_YEAR)
    End Sub

    ''' <summary>
    ''' Given all the standard parameters of the log normal stochastic process this function 
    ''' retunr a random value for time T
    ''' </summary>
    ''' <param name="TimeInDay"></param>
    ''' <param name="ValueStart"></param>
    ''' <param name="Gain"></param>
    ''' <param name="Volatility"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function Run(ByVal TimeInDay As Double, ByVal ValueStart As Double, ByVal Gain As Double, ByVal Volatility As Double) As Double
      Dim ThisResult As Double
      Dim ThisTimeInYear As Double
      Dim ThisMuSigma As Double = Gain - (Volatility ^ 2 / 2)
      Dim ThisGaussian As New Distributions.Normal(0, 1)

      ThisMuSigma = Gain - (Volatility ^ 2 / 2)
      ThisTimeInYear = TimeInDay / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      ThisResult = ValueStart * Math.Exp((ThisMuSigma * ThisTimeInYear) + (Volatility * Math.Sqrt(ThisTimeInYear) * ThisGaussian.Sample))
      Return ThisResult
    End Function

    Public Function Sample(ByVal TimeInDay As Double) As Double
      Dim ThisResult As Double
      Dim ThisTimeInYear As Double
      ThisTimeInYear = TimeInDay / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      ThisResult = ValueStart * Math.Exp((MyMuSigma * ThisTimeInYear) + (MyVolatility * Math.Sqrt(ThisTimeInYear) * MyGaussian.Sample))
      Return ThisResult
    End Function

    Public Function Sample(ByVal TimeInDay As Double, ByVal ValueStartAtTimeZero As Double) As Double
      Dim ThisResult As Double
      Dim ThisTimeInYear As Double
      ThisTimeInYear = TimeInDay / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      ThisResult = ValueStartAtTimeZero * Math.Exp((MyMuSigma * ThisTimeInYear) + (MyVolatility * Math.Sqrt(ThisTimeInYear) * MyGaussian.Sample))
      Return ThisResult
    End Function

    ''' <summary>
    ''' Recurrent function that return the value of the stochastic process for every day. 
    ''' see equation 43 of paper by Kevin D. Brewer
    ''' </summary>
    ''' <returns>The latest trace value of the process</returns>
    ''' <remarks>This recurrent function can be used to obtain a typical data trace of the stochastic process with increasing time</remarks>
    Public Function Sample() As Double
      MyStochasticProcessLast = MyStochasticProcessLast * Math.Exp(MyMuSigma * TIME_DAILY_PERIOD_IN_YEAR + MyVolatility * MySquareRootOfTimeDailyPeriod * MyGaussian.Sample)
      Return MyStochasticProcessLast
    End Function

    ''' <summary>
    ''' Recurrent function that return the value of the stochastic process for every day including the open high low and last value of the trading day
    ''' see equation 43 of paper by Kevin D. Brewer
    ''' </summary>
    ''' <returns>The latest trace value of the process for a typical daily trade</returns>
    ''' <remarks>This recurrent function can be used to obtain a typical data trace of the stochastic process with increasing time</remarks>
    Public Function SampleAsDailyTrade() As IPriceVol

      Throw New NotImplementedException
      'Dim ThisPriceVol As PriceVol
      'With ThisPriceVol
      '  MyStochasticProcessLast = MyStochasticProcessLast * Math.Exp(MyMuSigma * TIME_DAILY_PERIOD_IN_YEAR + MyVolatility * MySquareRootOfTimeDailyPeriod * MyGaussian.Sample)

      'End With


      'Return MyStochasticProcessLast
    End Function

    Public Function SampleToPriceVol(ByVal Value As Double) As IPriceVol
      Dim ThisPriceVol As New PriceVol(CSng(Value))
      Dim ThisTemp As Single
      Const TIME_OF_MARKET_TRADING_IN_DAY As Double = 1 / 3

      Throw New NotImplementedException
      With ThisPriceVol
        'get a random variation over hour
        .Last = CSng(Me.Sample(TIME_OF_MARKET_TRADING_IN_DAY, Value))
        .High = CSng(Me.Sample(TIME_OF_MARKET_TRADING_IN_DAY, Value))
        .Low = CSng(Me.Sample(TIME_OF_MARKET_TRADING_IN_DAY, Value))

        'validation
        If .High < .Low Then
          ThisTemp = .Low
          .Low = .High
          .High = .Low
        End If
        If .Last > .High Then
          ThisTemp = .High
        End If
      End With
    End Function

    Private Sub Swap(Of T)(ByRef Value1 As T, ByRef Value2 As T)
      Dim ThisT = Value1

      Value1 = Value2
      Value2 = ThisT
    End Sub

    'return the last sample trace sample
    Public ReadOnly Property SampleLast() As Double
      Get
        Return MyStochasticProcessLast
      End Get
    End Property

    ''' <summary>
    ''' Provide a typical trace of the process with increasing time value. 
    ''' The time is always reset to zero before the data collection is started
    ''' </summary>
    ''' <param name="NumberPoints">The Number of points to collect</param>
    ''' <returns>The data array with a typical trace of the process</returns>
    ''' <remarks></remarks>
    Public Function Samples(ByVal NumberPoints As Integer) As Double()
      Dim ThisArray(0 To NumberPoints - 1) As Double
      Dim I As Integer

      Me.Reset()
      For I = 0 To NumberPoints - 1
        ThisArray(I) = Me.Sample
      Next
      Return ThisArray
    End Function

    'Public Function TransformAsIPriceVol(ByVal Value As Double) As IPriceVol

    'End Function
    
    Public ReadOnly Property ValueStart As Double
      Get
        Return MyValueStart
      End Get
    End Property

    Public ReadOnly Property Gain As Double
      Get
        Return MyGain
      End Get
    End Property

    Public ReadOnly Property Volatility As Double
      Get
        Return MyVolatility
      End Get
    End Property

    Public ReadOnly Property Count As Integer
      Get
        Return MyCount
      End Get
    End Property

    Public Sub Reset()
      MyCount = 0
      MyStochasticProcessLast = ValueStart
    End Sub
  End Class
End Namespace
