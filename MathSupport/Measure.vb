Imports MathNet.Numerics

Namespace MathPlus
#Region "Measure"
  Namespace Measure
    Public Class Measure
      Public Enum enuOptionType
        _Call
        _Put
      End Enum


      'See: https://en.wikipedia.org/wiki/68%E2%80%9395%E2%80%9399.7_rule
      Public Const GAUSSIAN_PROBABILITY_SIGMA1 As Double = 0.682689492137086   'note: inside probability
      Public Const GAUSSIAN_PROBABILITY_SIGMA2 As Double = 0.954499736103642
      Public Const GAUSSIAN_PROBABILITY_SIGMA3 As Double = 0.99730020393674
      Public Const GAUSSIAN_PROBABILITY_SIGMA4 As Double = 0.999936657516334
      Public Const GAUSSIAN_PROBABILITY_SIGMA5 As Double = 0.999999426696856
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1 As Double = 0.5 + GAUSSIAN_PROBABILITY_SIGMA1 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1 As Double = 0.5 - GAUSSIAN_PROBABILITY_SIGMA1 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2 As Double = 0.5 + GAUSSIAN_PROBABILITY_SIGMA2 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2 As Double = 0.5 - GAUSSIAN_PROBABILITY_SIGMA2 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3 As Double = 0.5 + GAUSSIAN_PROBABILITY_SIGMA3 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3 As Double = 0.5 - GAUSSIAN_PROBABILITY_SIGMA3 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA4 As Double = 0.5 + GAUSSIAN_PROBABILITY_SIGMA4 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA4 As Double = 0.5 - GAUSSIAN_PROBABILITY_SIGMA4 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA5 As Double = 0.5 + GAUSSIAN_PROBABILITY_SIGMA5 / 2
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA5 As Double = 0.5 - GAUSSIAN_PROBABILITY_SIGMA5 / 2

      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1_LABEL As String = "+σ1"
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1_LABEL As String = "-σ1"
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA2_LABEL As String = "+σ2"
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA2_LABEL As String = "-σ2"
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA3_LABEL As String = "+σ3"
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA3_LABEL As String = "-σ3"
      Public Const GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA4_LABEL As String = "+σ4"
      Public Const GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA4_LABEL As String = "-σ4"

      Public Const TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY As Double = (YahooAccessData.ReportDate.MARKET_OPEN_TO_CLOSE_PERIOD_HOUR_DEFAULT) / 24  'market open for 6.5 hours
      Public Const TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY As Double = 1 - TIME_TO_MARKET_FROM_OPEN_TO_CLOSE_IN_DAY  'market open for 6.5 hours


      Public Shared Function Mean(ByVal Value As IEnumerable(Of Double)) As Double
        Dim ThisMean As Double = 0
        Dim ThisElement As Double

        For Each ThisElement In Value
          ThisMean = ThisMean + ThisElement
        Next
        Return ThisMean / Value.Count
      End Function

      Public Shared Function Mean(ByRef Value() As Double) As Double
        Return Measure.Mean(Value, 0, Value.Length - 1)
      End Function

      Public Shared Function Mean(ByRef Value() As Double, ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
        Dim I As Integer
        Dim ThisMean As Double = 0

        For I = StartPoint To StopPoint
          ThisMean = ThisMean + Value(I)
        Next
        Return ThisMean / (StopPoint - StartPoint + 1)
      End Function


      Public Shared Function ValidatePointRange(ByVal Value As IEnumerable(Of Double), ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Tuple(Of Integer, Integer, Integer)
        If StartPoint < 0 Then StartPoint = 0
        If IsNothing(Value) = False Then
          If StartPoint > Value.Count - 1 Then
            StartPoint = Value.Count - 1
          End If
          If StopPoint < 0 Then StopPoint = 0
          If StopPoint > Value.Count - 1 Then
            StopPoint = Value.Count - 1
          End If
        Else
          StartPoint = 0
          StopPoint = 0
        End If
        If StopPoint < StartPoint Then
          StopPoint = StartPoint
        End If
        Return New Tuple(Of Integer, Integer, Integer)(StartPoint, StopPoint, StopPoint - StartPoint + 1)
      End Function

      Public Shared Function ValidatePointRange(ByVal Count As Integer, ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Tuple(Of Integer, Integer, Integer)
        If StartPoint < 0 Then StartPoint = 0
        If Count > 0 Then
          If StartPoint > Count - 1 Then
            StartPoint = Count - 1
          End If
          If StopPoint < 0 Then StopPoint = 0
          If StopPoint > Count - 1 Then
            StopPoint = Count - 1
          End If
        Else
          StartPoint = 0
          StopPoint = 0
        End If
        If StopPoint < StartPoint Then
          StopPoint = StartPoint
        End If
        Return New Tuple(Of Integer, Integer, Integer)(StartPoint, StopPoint, StopPoint - StartPoint + 1)
      End Function


      Public Shared Function Mean(ByVal Value As IEnumerable(Of Double), ByVal StartPoint As Integer, ByVal StopPoint As Integer, ByVal IsValidatePointRange As Boolean) As Double
        If IsValidatePointRange Then
          With ValidatePointRange(Value, StartPoint, StopPoint)
            StartPoint = .Item1
            StopPoint = .Item2
          End With
        End If
        Return Mean(Value, StartPoint, StopPoint)
      End Function


      Public Shared Function Mean(ByVal Value As IEnumerable(Of Double), ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
        Dim I As Integer
        Dim ThisMean As Double = 0

        For I = StartPoint To StopPoint
          ThisMean = ThisMean + Value(I)
        Next
        Return ThisMean / (StopPoint - StartPoint + 1)
      End Function

      Public Shared Function RMS(ByRef Value() As Double, ByVal Mean As Double) As Double
        Return RMS(Value, Mean, 0, Value.Length - 1)
      End Function

      Public Shared Function RMS(ByVal Value As IEnumerable(Of Double), ByVal Mean As Double) As Double
        Dim I As Integer
        Dim ThisRMSSum As Double = 0
        Dim ThisElement As Double

        For Each ThisElement In Value
          ThisRMSSum = ThisRMSSum + (ThisElement - Mean) ^ 2
        Next
        Return Math.Sqrt(ThisRMSSum / Value.Count)
      End Function

      Public Shared Function RMS(ByRef Value() As Double) As Double
        Return RMS(Value, Measure.Mean(Value), 0, Value.Length - 1)
      End Function

      Public Shared Function RMS(ByVal Value As IEnumerable(Of Double)) As Double
        Return RMS(Value, Measure.Mean(Value))
      End Function

      Public Shared Function RMS(ByRef Value() As Double, ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
        Return RMS(Value, Measure.Mean(Value, StartPoint, StopPoint), StartPoint, StopPoint)
      End Function

      ''' <summary>
      ''' return the standard deviation around the mean
      ''' </summary>
      ''' <param name="Value"></param>
      ''' <param name="Mean"></param>
      ''' <param name="StartPoint"></param>
      ''' <param name="StopPoint"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>

      Public Shared Function RMS(ByRef Value() As Double, ByVal Mean As Double, ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
        Dim I As Integer
        Dim ThisRMSSum As Double = 0

        For I = StartPoint To StopPoint
          ThisRMSSum = ThisRMSSum + (Value(I) - Mean) ^ 2
        Next
        Return Math.Sqrt(ThisRMSSum / (StopPoint - StartPoint + 1))
      End Function

      Public Shared Function RMS(ByVal Value As IEnumerable(Of Double), ByVal Mean As Double, ByVal StartPoint As Integer, ByVal StopPoint As Integer) As Double
        Dim I As Integer
        Dim ThisRMSSum As Double = 0

        For I = StartPoint To StopPoint
          ThisRMSSum = ThisRMSSum + (Value(I) - Mean) ^ 2
        Next
        Return Math.Sqrt(ThisRMSSum / (StopPoint - StartPoint + 1))
      End Function

      ''' <summary>
      ''' Calculate the true statistical variance taking into account
      ''' the finite number of element in the array
      ''' </summary>
      ''' <param name="Value"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Shared Function Variance(ByRef Value() As Double) As Double
        Dim I As Integer
        Dim ThisRMSSum As Double
        Dim ThisMean As Double

        If Value.Length <= 1 Then
          Throw New ArgumentOutOfRangeException
        End If
        ThisMean = Measure.Mean(Value)
        ThisRMSSum = 0.0
        For I = 0 To Value.Length - 1
          ThisRMSSum = ThisRMSSum + (Value(I) - ThisMean) ^ 2
        Next
        Return ThisRMSSum / Value.Length
      End Function


      ''' <summary>
      ''' The Black and Scholes (1973) europeen Stock option formula from:
      ''' http://www.espenhaug.com/black_scholes.html
      ''' </summary>
      ''' <param name="OptionType">call or put</param>
      ''' <param name="StockPrice">actual stock price</param>
      ''' <param name="OptionStrikePrice">Strike Price of the option</param>
      ''' <param name="TimeToExpirationInYear">Number of days to expiration in years</param>
      ''' <param name="RiskFreeRate">Rate of risk free investment</param>
      ''' <param name="VolatilityPerYear">Standard deviation Volatility per year</param>
      ''' <returns>The fair value of the option</returns>
      ''' <remarks>See The Complete Guide to Option Pricing Formulas by Espen Gaarder Haug</remarks>
      Public Shared Function BlackScholes(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal VolatilityPerYear As Double) As Double

        Return BlackScholes(
          OptionType,
          StockPrice,
          OptionStrikePrice,
          TimeToExpirationInYear,
          RiskFreeRate,
          0.0,
          VolatilityPerYear)
      End Function

      Public Shared Function BlackScholes(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As Double

        Dim d1 As Double, d2 As Double
        Dim ThisNormal = New MathNet.Numerics.Distributions.Normal(mean:=0, stddev:=1)
        'DividendRate = 0


        d1 = (Math.Log(StockPrice / OptionStrikePrice) + (RiskFreeRate - DividendRate + VolatilityPerYear ^ 2 / 2) * TimeToExpirationInYear) / (VolatilityPerYear * Math.Sqrt(TimeToExpirationInYear))
        d2 = d1 - VolatilityPerYear * Math.Sqrt(TimeToExpirationInYear)

        'd1 = (Log(S / x) + ( b + v ^ 2 / 2) * T) / (v * Sqr(T))
        'd2 = d1 - v * Sqr(T)

        'If CallPutFlag = "c" Then
        '  GBlackScholes = S * Exp((b - r) * T) * CND(d1) - x * Exp(-r * T) * CND(d2)
        'ElseIf CallPutFlag = "p" Then
        '  GBlackScholes = x * Exp(-r * T) * CND(-d2) - S * Exp((b - r) * T) * CND(-d1)
        'End If
        'C = Se−qTN(d1)−Ke−rTN(d2)
        'P = Ke−rTN(−d2)−Se−qTN(d1)
        'Where q Is the known dividend yield And
        'd1 = ln(SK) + T(r−q+σ22)σT−−√
        'd2 = ln(SK) + T(r−q−σ22)σT−−√=d1−σT−−√
        'One possible implementation (a rough And unoptimized 

        'Note:  b = RiskFreeRate - DividendRate

        Dim ThisCDFOf_d1 As Double = ThisNormal.CumulativeDistribution(d1)
        Dim ThisCDFOf_Negd1 As Double = ThisNormal.CumulativeDistribution(-d1)
        Dim ThisCDFOf_d2 As Double = ThisNormal.CumulativeDistribution(d2)
        Dim ThisCDFOf_Negd2 As Double = ThisNormal.CumulativeDistribution(-d2)

        If OptionType = enuOptionType._Call Then
          Return StockPrice * Math.Exp(-DividendRate * TimeToExpirationInYear) * ThisCDFOf_d1 - OptionStrikePrice * Math.Exp(-RiskFreeRate * TimeToExpirationInYear) * ThisCDFOf_d2
        Else
          Return OptionStrikePrice * Math.Exp(-RiskFreeRate * TimeToExpirationInYear) * ThisCDFOf_Negd2 - StockPrice * Math.Exp(-DividendRate * TimeToExpirationInYear) * ThisCDFOf_Negd1
        End If
      End Function

      ''' <summary>
      ''' Delta is the amount an option price is expected to move based on a $1 change 
      ''' in the underlying stock. Calls have positive delta, between 0 and 1.
      ''' </summary>
      ''' <param name="OptionType"></param>
      ''' <param name="StockPrice"></param>
      ''' <param name="OptionStrikePrice"></param>
      ''' <param name="TimeToExpirationInYear"></param>
      ''' <param name="RiskFreeRate"></param>
      ''' <param name="DividendRate"></param>
      ''' <param name="VolatilityPerYear"></param>
      ''' <returns></returns>
      Public Shared Function BlackScholesOptionDelta(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As Double

        Dim d1 As Double
        Dim ThisNormal = New MathNet.Numerics.Distributions.Normal(mean:=0, stddev:=1)
        'DividendRate = 0

        d1 = (Math.Log(StockPrice / OptionStrikePrice) + (RiskFreeRate - DividendRate + VolatilityPerYear ^ 2 / 2) * TimeToExpirationInYear) / (VolatilityPerYear * Math.Sqrt(TimeToExpirationInYear))

        Dim ThisCDFOf_d1 As Double = ThisNormal.CumulativeDistribution(d1)

        If OptionType = enuOptionType._Call Then
          Return Math.Exp(-DividendRate * TimeToExpirationInYear) * ThisCDFOf_d1
        Else
          Return Math.Exp(-DividendRate * TimeToExpirationInYear) * (ThisCDFOf_d1 - 1)
        End If
      End Function

      ''' <summary>
      ''' The option vega expresses the change in the price of the option 
      ''' for every 1% change in the volatility.
      ''' 
      ''' note:
      ''' From book and excel : GVega = S * Exp((b - r) * T) * ND(d1) * Sqr(T)
      ''' S=StockPrice
      ''' x =Strike Price
      ''' T=Time to expiration in years
      ''' r=Risk free interest rate
      ''' q=Dividend yields
      ''' b=Cost of carry (r-q)
      ''' v= volatility
      ''' </summary>
      ''' <param name="StockPrice"></param>
      ''' <param name="OptionStrikePrice"></param>
      ''' <param name="TimeToExpirationInYear"></param>
      ''' <param name="RiskFreeRate"></param>
      ''' <param name="DividendRate"></param>
      ''' <param name="VolatilityPerYear"></param>
      ''' <returns></returns>
      Public Shared Function BlackScholesOptionVega(
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As Double


        Dim d1 As Double
        'DividendRate = 0
        Dim ThisNormal = New MathNet.Numerics.Distributions.Normal(mean:=0, stddev:=1)

        d1 = (Math.Log(StockPrice / OptionStrikePrice) + (RiskFreeRate - DividendRate + VolatilityPerYear ^ 2 / 2) * TimeToExpirationInYear) / (VolatilityPerYear * Math.Sqrt(TimeToExpirationInYear))
        Dim ThisCDFOf_d1 As Double = ThisNormal.CumulativeDistribution(d1)

        Return StockPrice * Math.Exp(-DividendRate * TimeToExpirationInYear) * ThisCDFOf_d1 * Math.Sqrt(TimeToExpirationInYear)
      End Function


      Public Shared Function BlackScholesOptionImpliedVolatility(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal OptionMarketPrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal ErrorEpsilon As Double,
        ByVal Optional MaximumNumberOfIteration As Integer = 20) As Double

        Dim vi As Double
        Dim OptionMarketPriceEstimate As Double
        Dim vegai As Double
        Dim ThisOptionPriceEstimateDeltaAbsolute As Double
        Dim ThisOptionPriceEstimateDeltaAbsoluteLast As Double
        Dim ThisCount As Integer
        'DividendRate = 0
        'vi volatility estimate based on Manaster and Koehler seed value (vi)   (See Haug book P.P. 458 eq. 12.8)
        vi = Math.Sqrt(Math.Abs(Math.Log(StockPrice / OptionStrikePrice) + RiskFreeRate * TimeToExpirationInYear) * 2 / TimeToExpirationInYear)
        'if the estimate for volatility is too low the algorithm do not converge
        If vi < 0.05 Then
          vi = 0.05
        End If
        OptionMarketPriceEstimate = BlackScholes(OptionType, StockPrice, OptionStrikePrice, TimeToExpirationInYear, RiskFreeRate, DividendRate, vi)
        vegai = BlackScholesOptionVega(StockPrice, OptionStrikePrice, TimeToExpirationInYear, RiskFreeRate, DividendRate, vi)
        ThisOptionPriceEstimateDeltaAbsolute = Math.Abs(OptionMarketPrice - OptionMarketPriceEstimate)
        ThisOptionPriceEstimateDeltaAbsoluteLast = ThisOptionPriceEstimateDeltaAbsolute
        Try
          'keep going until lhe error si sufficantly small and keep going down from the last result
          ThisCount = 0
          While _
            (ThisOptionPriceEstimateDeltaAbsolute >= ErrorEpsilon) And
            (ThisOptionPriceEstimateDeltaAbsolute <= ThisOptionPriceEstimateDeltaAbsoluteLast)

            vi = vi - (OptionMarketPriceEstimate - OptionMarketPrice) / (vegai + Double.Epsilon)
            OptionMarketPriceEstimate = BlackScholes(OptionType, StockPrice, OptionStrikePrice, TimeToExpirationInYear, RiskFreeRate, DividendRate, vi)
            vegai = BlackScholesOptionVega(StockPrice, OptionStrikePrice, TimeToExpirationInYear, RiskFreeRate, DividendRate, vi)
            ThisOptionPriceEstimateDeltaAbsoluteLast = ThisOptionPriceEstimateDeltaAbsolute
            ThisOptionPriceEstimateDeltaAbsolute = Math.Abs(OptionMarketPrice - OptionMarketPriceEstimate)
            ThisCount += 1
            If ThisCount > MaximumNumberOfIteration Then Exit While
          End While
        Catch ex As Exception
          'should Not happen 
          Debugger.Break()   'want to know when it happen if the debugger is running
          Throw ex
        End Try
        If (ThisOptionPriceEstimateDeltaAbsolute < ErrorEpsilon) Or (ThisCount > MaximumNumberOfIteration) Then
          Return vi
        Else
          Debugger.Break()   'want to know when it happen if the debugger is running
          Return Double.NaN
        End If
      End Function


      Public Shared Function HaugHaugDividendVolatilityCorrection(
        ByVal StockPrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByRef DividendPaymentValues() As Double,
        ByRef DividendTimesInYear() As Double,
        ByVal VolatilityPerYear As Double) As Double

        Return General.HaugHaugDividendVolatilityCorrection(
          StockPrice,
          TimeToExpirationInYear,
          RiskFreeRate,
          DividendPaymentValues,
          DividendTimesInYear,
          VolatilityPerYear)
      End Function


      ''' <summary>
      ''' The Bjerksund and Stensland (2002) American approximation american Stock option formula from:
      ''' http://www.espenhaug.com/black_scholes.html
      ''' </summary>
      ''' <param name="OptionType">call or put</param>
      ''' <param name="StockPrice">actual stock price</param>
      ''' <param name="OptionStrikePrice">Strike Price of the option</param>
      ''' <param name="TimeToExpirationInYear">Number of days to expiration in years</param>
      ''' <param name="RiskFreeRate">Rate of risk free investment</param>
      ''' <param name="VolatilityPerYear">Standard deviation Volatility per year</param>
      ''' <returns>The fair value of the option</returns>
      ''' <remarks>See The Complete Guide to Option Pricing Formulas by Espen Gaarder Haug</remarks>
      Public Shared Function BSAmericanOption(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As Double

        'TimeToExpirationInYear = 0.5
        '// The Bjerksund and Stensland (2002) American approximation
        Dim b = RiskFreeRate - DividendRate   'cost to carry
        If OptionType = enuOptionType._Call Then
          'BSAmericanCallApprox2002(S, X, T, r, b, v)
          Return BSAmericanCallApprox2002(StockPrice, OptionStrikePrice, TimeToExpirationInYear, RiskFreeRate, b, VolatilityPerYear)
        Else
          'BSAmericanCallApprox2002(X, S, T, r - b, -b, v)
          Return BSAmericanCallApprox2002(OptionStrikePrice, StockPrice, TimeToExpirationInYear, RiskFreeRate - b, -b, VolatilityPerYear)
        End If
      End Function

      ''' <summary>
      ''' The Bjerksund and Stensland (2002) American approximation american Stock option formula from:
      ''' http://www.espenhaug.com/black_scholes.html
      ''' </summary>
      ''' <param name="OptionType">call or put</param>
      ''' <param name="StockPrice">actual stock price</param>
      ''' <param name="OptionStrikePrice">Strike Price of the option</param>
      ''' <param name="TimeToExpirationInYear">Number of days to expiration in years</param>
      ''' <param name="RiskFreeRate">Rate of risk free investment</param>
      ''' <param name="VolatilityPerYear">Standard deviation Volatility per year</param>
      ''' <returns>The fair value of the option</returns>
      ''' <remarks>See The Complete Guide to Option Pricing Formulas by Espen Gaarder Haug</remarks>
      Public Shared Function BSAmericanOption(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As IPriceVol,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As IPriceVol

        Dim ThisPriceVol As IPriceVol = New PriceVolAsClass

        ThisPriceVol.Open = CSng(Measure.BSAmericanOption(
          OptionType:=OptionType,
          StockPrice:=StockPrice.Open,
          OptionStrikePrice:=OptionStrikePrice,
          TimeToExpirationInYear:=TimeToExpirationInYear,
          RiskFreeRate:=RiskFreeRate,
          DividendRate:=DividendRate,
          VolatilityPerYear:=VolatilityPerYear))
        ThisPriceVol.Low = CSng(Measure.BSAmericanOption(
          OptionType:=OptionType,
          StockPrice:=StockPrice.Low,
          OptionStrikePrice:=OptionStrikePrice,
          TimeToExpirationInYear:=TimeToExpirationInYear,
          RiskFreeRate:=RiskFreeRate,
          DividendRate:=DividendRate,
          VolatilityPerYear:=VolatilityPerYear))
        ThisPriceVol.High = CSng(Measure.BSAmericanOption(
          OptionType:=OptionType,
          StockPrice:=StockPrice.High,
          OptionStrikePrice:=OptionStrikePrice,
          TimeToExpirationInYear:=TimeToExpirationInYear,
          RiskFreeRate:=RiskFreeRate,
          DividendRate:=DividendRate,
          VolatilityPerYear:=VolatilityPerYear))
        ThisPriceVol.Last = CSng(Measure.BSAmericanOption(
          OptionType:=OptionType,
          StockPrice:=StockPrice.Last,
          OptionStrikePrice:=OptionStrikePrice,
          TimeToExpirationInYear:=TimeToExpirationInYear,
          RiskFreeRate:=RiskFreeRate,
          DividendRate:=DividendRate,
          VolatilityPerYear:=VolatilityPerYear))
        ThisPriceVol.Vol = 1
        Return ThisPriceVol
      End Function

      ''' <summary>
      ''' Generate a sample for a gaussian distribution
      ''' </summary>
      ''' <param name="Mean"></param>
      ''' <param name="StandardDeviation"></param>
      ''' <returns></returns>
      Public Shared Function Gaussian(ByVal Mean As Double, ByVal StandardDeviation As Double) As Double
        Dim NormalDist As Distributions.Normal = New MathNet.Numerics.Distributions.Normal(Mean, StandardDeviation)
        Return NormalDist.Sample
      End Function

      Public Shared Function Gaussian(ByVal Mean As Double, ByVal StandardDeviation As Double, ByVal NumberOfPoint As Integer) As Double()
        Dim ThisArray(0 To NumberOfPoint - 1) As Double
        Dim NormalDist As Distributions.Normal = New MathNet.Numerics.Distributions.Normal(Mean, StandardDeviation)
        NormalDist.Samples(ThisArray)
        Return ThisArray
      End Function

      ''' <summary>
      ''' Calculate the Inverse Normal distribution for a symetric range between 0.001 et 0.999
      ''' </summary>
      ''' <param name="ProbabilityValue"> The probability value</param>
      ''' <returns>the inverse normal X value from -3 to +3 range</returns>
      Public Shared Function InverseNormal(ByVal ProbabilityValue As Double) As Double
        Return InverseNormal(ProbabilityValue, 3.0)
      End Function

      ''' <summary>
      ''' Calculate the InverseNormal (Mean=0, Sigma=1) over a range of x value from -RangeX to +RangeX.
      ''' </summary>
      ''' <param name="ProbabilityValue"></param>
      ''' <param name="RangeOfX">The range of X from -RangeX to +RangeX</param>
      ''' <returns></returns>
      Public Shared Function InverseNormal(ByVal ProbabilityValue As Double, ByVal RangeOfX As Double) As Double
        Dim ThisResult As Double
        ThisResult = MathNet.Numerics.Distributions.Normal.InvCDF(mean:=0, stddev:=1.0, ProbabilityValue)
        If ThisResult > RangeOfX Then
          ThisResult = RangeOfX
        ElseIf ThisResult < -RangeOfX Then
          ThisResult = -RangeOfX
        End If
        Return ThisResult
      End Function

      Public Shared Function LogNormal(ByVal Mu As Double, ByVal Sigma As Double, ByVal NumberOfPoint As Integer) As Double()
        Dim ThisArray(0 To NumberOfPoint - 1) As Double
        Dim LogNormalDist As Distributions.LogNormal = New MathNet.Numerics.Distributions.LogNormal(Mu, Sigma)
        LogNormalDist.Samples(ThisArray)
        Return ThisArray
      End Function

      ''' <summary>
      ''' Calculate the approximative logarithm gain using and squared data approximation that eliminate the problem of negative value but
      ''' that require that values to be ideally > 1 for accurate result.
      ''' See:https://people.duke.edu/~rnau/411log.htm
      ''' </summary>
      ''' <param name="Value">should be > 1 for valid result</param>
      ''' <param name="ValueRef"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Shared Function GainLog(ByVal Value As Double, ValueRef As Double) As Double
        'gain limiting is ignored for negative value 
        Return GainLog(Value, ValueRef, ScaleValue:=1.0)
        'Return Math.Log(((Value ^ 2 + 1) / (ValueRef ^ 2 + 1))) / 2
      End Function

      ''' <summary>
      ''' Compress and expand the input probabillity using the gaussian scale transformation and return a value between 0 and 1
      ''' corresponging to the gaussian scale transformation.  
      ''' </summary>
      ''' <param name="ProbabilityValue">The imput probability between 0 and 1</param>
      ''' <param name="ScaleOfX">the positive range of X value</param>
      ''' <returns></returns>
      Public Shared Function ProbabilityToGaussianScale(ByVal ProbabilityValue As Double, ByVal ScaleOfX As Double) As Double
        Return (Measure.InverseNormal(ProbabilityValue, ScaleOfX) / (2 * ScaleOfX)) + 0.5
      End Function

      Public Shared Function ProbabilityToGaussianScale(ByVal ProbabilityValue As IEnumerable(Of Double), ByVal ScaleOfX As Double) As IEnumerable(Of Double)
        Dim ThisList As New List(Of Double)
        For Each ThisValue In ProbabilityValue
          ThisList.Add((Measure.InverseNormal(ThisValue, ScaleOfX) / (2 * ScaleOfX)) + 0.5)
        Next
        Return ThisList
      End Function

      ''' <summary>
      ''' This function return the log gain between two value with a multiplier scale value. 
      ''' The function return zero if it teh range is too large for evaluation
      ''' </summary>
      ''' <param name="Value"></param>
      ''' <param name="ValueRef"></param>
      ''' <param name="ScaleValue"></param>
      ''' <returns></returns>
      Public Shared Function GainLog(ByVal Value As Double, ValueRef As Double, ByVal ScaleValue As Double) As Double

        Dim ThisResult As Double
        If ValueRef <= 0.0 Then
          'ThisResult = Double.NaN
          ThisResult = 0.0
        ElseIf Value <= 0 Then
          ThisResult = 0.0
        Else
          ThisResult = ScaleValue * Math.Log(Value / ValueRef)
          If Double.IsNaN(ThisResult) Or Double.IsInfinity(ThisResult) Then
            'ThisResult = Double.NaN
            ThisResult = 0.0
          End If
        End If
        Return ThisResult
      End Function

      Public Shared Function GainLog(
        ByVal Value As Double,
        ByVal ValueRef As Double,
        ByVal ScaleValue As Double,
        ByVal LimitGainAbsolute As Double) As Double

        'limit exponentially the gain value between -LimitGainAbsolute and +LimitGainAbsolute
        Dim ThisResult = GainLog(Value, ValueRef, ScaleValue)
        ThisResult = MathPlus.WaveForm.SignalLimit(ThisResult, LimitGainAbsolute)
        Return ThisResult
      End Function
      ''' <summary>
      ''' The inverse of the GainLog defined function
      ''' </summary>
      ''' <param name="GainLog"></param>
      ''' <param name="ValueRef"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Shared Function GainLogInverse(ByVal GainLog As Double, ValueRef As Double) As Double
        If Double.IsNaN(GainLog) Then
          Return Double.NaN
        Else
          Return ValueRef * Math.Exp(GainLog)
        End If
        'Dim ThisResult As Double = ((ValueRef ^ 2 + 1) * Math.Exp(2 * GainLog)) - 1
        'If ThisResult < 0.0 Then
        '  ThisResult = 0
        'End If
        'Return Math.Sqrt(ThisResult)
      End Function

      Public Shared Function LogNormal(ByVal Mu As Double, ByVal Sigma As Double) As Double
        Dim LogNormalDist As Distributions.LogNormal = New MathNet.Numerics.Distributions.LogNormal(Mu, Sigma)
        Return LogNormalDist.Sample
      End Function

      Public Shared Function VolatilityPrediction(ByVal FilterRate As Integer, ByVal Gain As Double, ByVal StandardDeviation As Double) As Double
        Dim ThisFilter = New Filter.FilterVolatility(FilterRate, 1.0)
        Dim I As Integer
        Dim ThisGainLinear = Math.Exp(Gain)
        Dim ThisRandomData() = LogNormal(0, StandardDeviation, 365)
        Dim ThisRandomData1() = LogNormal(0, StandardDeviation, 365)

        For I = 0 To ThisRandomData.Length - 1
          ThisFilter.Filter(ThisRandomData(I))
          ThisFilter.Filter(ThisGainLinear * ThisRandomData1(I))
        Next
        Return ThisFilter.FilterLast / Math.Sqrt(Gain ^ 2 + StandardDeviation ^ 2)
      End Function

      ''' <summary>
      ''' The cumulative normal distribution function approimative implementation for mean=0 and sigma = 1
      ''' see http://en.wikipedia.org/wiki/Normal_distribution
      ''' </summary>
      ''' <param name="X"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Public Shared Function CDFGaussian(ByVal X As Double) As Double

        Dim L As Double, K As Double
        Dim ThisResult As Double

        Const a1 As Double = 0.31938153
        Const a2 As Double = -0.356563782
        Const a3 As Double = 1.781477937
        Const a4 As Double = -1.821255978
        Const a5 As Double = 1.330274429

        L = Math.Abs(X)

        K = 1 / (1 + 0.2316419 * L)

        ThisResult = 1 - 1 / Math.Sqrt(2 * Math.PI) * Math.Exp(-L ^ 2 / 2) * (a1 * K + a2 * K ^ 2 + a3 * K ^ 3 + a4 * K ^ 4 + a5 * K ^ 5)

        If X < 0 Then
          ThisResult = 1 - ThisResult
        End If
        Return ThisResult
      End Function

      ''' <summary>
      ''' The cumulative normal distribution function approimative implementation for mean=0 and sigma = 1
      ''' see http://en.wikipedia.org/wiki/Normal_distribution
      ''' </summary>
      ''' <param name="X">the variable point</param>
      ''' <returns></returns>
      ''' <remarks>This function MathNet </remarks>
      Public Shared Function CDFGaussian(ByVal Mean As Double, ByVal StandardDeviation As Double, ByVal X As Double) As Double
        Dim ThisNormalDist = New MathNet.Numerics.Distributions.Normal(Mean, StandardDeviation)

        Return ThisNormalDist.CumulativeDistribution(X)
      End Function

#Disable Warning BC42304 ' XML documentation parse error
      ''' <summary>
      ''' The cumulative normal distribution function approimative implementation for mean=0 and sigma = 1
      ''' see http://en.wikipedia.org/wiki/Normal_distribution
      ''' </summary>
      ''' <param name="X">the variable point</param>
      ''' <returns>The P(Normal Function < X)</returns>
      ''' <remarks>This function MathNet </remarks>
      Public Shared Function InverseCDFGaussian(ByVal Mean As Double, ByVal StandardDeviation As Double, ByVal Probability As Double) As Double
#Enable Warning BC42304 ' XML documentation parse error
        Dim ThisNormalDist = New MathNet.Numerics.Distributions.Normal(Mean, StandardDeviation)

        Return ThisNormalDist.InverseCumulativeDistribution(Probability)
      End Function

      ' ''' <summary>
      ' ''' Calculate the Probability Density Function (PDF) for a given range of data
      ' ''' </summary>
      ' ''' <param name="ValueLogScaled">
      ' ''' The input value scaled on a logarithm scale and quantisized to an integer
      ' ''' </param>
      ' ''' <param name="RangeMinimum">
      ' ''' The minimum value for the PDF calculation
      ' ''' </param>
      ' ''' <param name="RangeMaximum">
      ' ''' The maximum value for the PDF calculation
      ' ''' </param>
      ' ''' <returns>The PDF density function scaled in PerCent
      ' ''' </returns>
      ' ''' <remarks></remarks>
      'Public Shared Function PDF(ByRef ValueLogScaled() As Integer, ByVal RangeMinimum As Integer, ByVal RangeMaximum As Integer) As Double()
      '  Dim I As Integer
      '  Dim J As Integer
      '  Dim ThisValueScaledTodB As Integer
      '  Dim ThisValueScaledTodBLast As Integer
      '  Dim LCRDistribution(0 To (RangeMaximum - RangeMinimum)) As Double


      '  ThisValueScaledTodBLast = ValueLogScaled(0)
      '  'bound the input value to a valid range of LCRDistribution
      '  If ThisValueScaledTodBLast < RangeMinimum Then
      '    ThisValueScaledTodBLast = RangeMinimum
      '  ElseIf ThisValueScaledTodBLast > RangeMaximum Then
      '    ThisValueScaledTodBLast = RangeMaximum
      '  End If
      '  For I = 1 To ValueLogScaled.Length - 1
      '    ThisValueScaledTodB = ValueLogScaled(I)
      '    'bound the input value to a valid range of LCRDistribution
      '    If ThisValueScaledTodB < RangeMinimum Then
      '      ThisValueScaledTodB = RangeMinimum
      '    ElseIf ThisValueScaledTodB > RangeMaximum Then
      '      ThisValueScaledTodB = RangeMaximum
      '    End If
      '    'calculate the positive Level Crossing Rate
      '    If ThisValueScaledTodB > ThisValueScaledTodBLast Then
      '      For J = ThisValueScaledTodBLast To ThisValueScaledTodB - 1
      '        LCRDistribution(J - RangeMinimum) = LCRDistribution(J - RangeMinimum) + 1
      '      Next
      '    End If
      '    ThisValueScaledTodBLast = ThisValueScaledTodB
      '  Next
      '  For I = 0 To LCRDistribution.Length - 1
      '    LCRDistribution(I) = LCRDistribution(I) / ValueLogScaled.Length
      '  Next
      '  Return LCRDistribution
      'End Function

      ''' <summary>
      ''' Calculate the positive Level Crossing Rate (LCR) density function for a given range of data
      ''' </summary>
      ''' <param name="ValueLogScaled">
      ''' The input value scaled on a logarithm scale and quantisized to an integer
      ''' </param>
      ''' <param name="RangeMinimum">
      ''' The minimum value for the LCR calculation
      ''' </param>
      ''' <param name="RangeMaximum">
      ''' The maximum value for the LCR calculation
      ''' </param>
      ''' <returns>The LCR density function scaled in PerCent
      ''' </returns>
      ''' <remarks></remarks>
      Public Shared Function LCRLevel(ByRef ValueLogScaled() As Integer, ByVal RangeMinimum As Integer, ByVal RangeMaximum As Integer) As Double()
        Dim I As Integer
        Dim J As Integer
        Dim ThisValueScaledTodB As Integer
        Dim ThisValueScaledTodBLast As Integer
        Dim LCRDistribution(0 To (RangeMaximum - RangeMinimum)) As Double


        ThisValueScaledTodBLast = ValueLogScaled(0)
        'bound the input value to a valid range of LCRDistribution
        If ThisValueScaledTodBLast < RangeMinimum Then
          ThisValueScaledTodBLast = RangeMinimum
        ElseIf ThisValueScaledTodBLast > RangeMaximum Then
          ThisValueScaledTodBLast = RangeMaximum
        End If
        For I = 1 To ValueLogScaled.Length - 1
          ThisValueScaledTodB = ValueLogScaled(I)
          'bound the input value to a valid range of LCRDistribution
          If ThisValueScaledTodB < RangeMinimum Then
            ThisValueScaledTodB = RangeMinimum
          ElseIf ThisValueScaledTodB > RangeMaximum Then
            ThisValueScaledTodB = RangeMaximum
          End If
          'calculate the positive Level Crossing Rate
          If ThisValueScaledTodB > ThisValueScaledTodBLast Then
            For J = ThisValueScaledTodBLast To ThisValueScaledTodB - 1
              LCRDistribution(J - RangeMinimum) = LCRDistribution(J - RangeMinimum) + 1
            Next
          End If
          ThisValueScaledTodBLast = ThisValueScaledTodB
        Next
        For I = 0 To LCRDistribution.Length - 1
          LCRDistribution(I) = LCRDistribution(I) / ValueLogScaled.Length
        Next
        Return LCRDistribution

        '        ' Assign Min and Max values of PSD
        '        MinPSDValue = ThisValueMin
        '        MaxPSDValue = ThisValueMax

        '        Occupancy = CountHit / (StopPoint - StartPoint + 1)
        '        'find the noise peak
        '        CalculatedNoiseLevel = ThisValueMax      'by default
        '        For I = ThisValueMin To ThisValueMax - 1
        '          If LCR(I) > LCR(I + 1) Then
        '            'verify that this is not a local peak
        '            StartPoint = I - 3
        '            If StartPoint < ThisValueMin Then StartPoint = ThisValueMin
        '            StopPoint = I + 3
        '            If StopPoint > ThisValueMax Then StopPoint = ThisValueMax
        '            For J = StartPoint To StopPoint
        '              If LCR(J) > LCR(I) Then
        '                'this was not a global peak
        '                I = J - 1
        '                Exit For
        '              End If
        '            Next
        '            If J > StopPoint Then
        '              'this is the global peak
        '#If LCRInterpolateEnabled Then
        '        X0 = I - 1
        '        X2 = I + 1
        '        'if False Then
        '        If X0 >= ThisValueMin Then
        '          If X2 <= ThisValueMax Then
        '            'interpolate the result
        '            'interpolation based on Newton
        '            'see: Calcul Differentiel et Integral, N. Piskounov, 1974, tome I, pp. 266-269
        '            dy = LCR(I) - LCR(X0)
        '            DY2 = LCR(X2) - LCR(I) - LCR(I) + LCR(X0)
        '            If DY2 <> 0 Then
        '              CalculatedNoiseLevel = X0 + 0.5 - dy / DY2
        '              'Debug.Print Row; Format(.NoiseLevel - I, "0.00")
        '              If CalculatedNoiseLevel < X0 Then
        '                CalculatedNoiseLevel = X0
        '              ElseIf CalculatedNoiseLevel > X2 Then
        '                CalculatedNoiseLevel = X2
        '              End If
        '              CalculatedNoiseLevel = CalculatedNoiseLevel + LCRCorrection
        '            Else
        '              'not able to interpolate the data
        '              CalculatedNoiseLevel = I + LCRCorrection
        '            End If
        '          Else
        '            'not able to interpolate the data
        '            CalculatedNoiseLevel = I + LCRCorrection
        '          End If
        '        Else
        '          'not able to interpolate the data
        '          CalculatedNoiseLevel = I + LCRCorrection
        '        End If
        '#Else
        '              CalculatedNoiseLevel = I + LCRCorrection
        '#End If
        '              Exit For
        '            End If
        '          End If
        '        Next
        '        'Debug.Print Format(CalculatedNoiseLevel, "0.0")

        '#If PDFNoiseMeasurement Then
        '  'find where the LCR start to go up from the noise peak
        '  StartPoint = NoiseLevel - 3
        '  StopPoint = ThisValueMax
        '  If StartPoint < ThisValueMin Then StartPoint = ThisValueMin
        '  For J = NoiseLevel To ThisValueMax - 1
        '    If ((LCR(J) - LCR(J - 1)) >= -3) Then
        '      StopPoint = J
        '    End If
        '    StartPoint = StartPoint + 1
        '  Next
        '  'get the average power
        '  StartPoint = ThisValueMin
        '  PowerSum = StartPoint * LCR(StartPoint)
        '  I = LCR(StartPoint)
        '  For J = StartPoint + 1 To StopPoint
        '    PowerSum = PowerSum + J * LCR(J)
        '    I = I + LCR(J)
        '  Next
        '  PDFNoiseLevel = PowerSum / I
        '  'NoiseLevel = PowerSum / I
        '#End If
        '        '#If PDFNoiseMeasurement Then
        '        'searching the noise
        '        'this is not finish
        '        '  PDFNoiseLevel = ThisValueMin
        '        '  CDF(ThisValueMin) = 0
        '        '  For I = ThisValueMin To ThisValueMax - 1
        '        '    CDF(I) = CDF(I) + PDF(I)
        '        '    If PDF(I) > PDF(I + 1) Then
        '        '      'verify that this is not a local peak
        '        '      StartPoint = I - 3
        '        '      If StartPoint < ThisValueMin Then StartPoint = ThisValueMin
        '        '      StopPoint = I + 3
        '        '      If StopPoint > ThisValueMax Then StopPoint = ThisValueMax
        '        '      For J = StartPoint To StopPoint
        '        '        If PDF(J) > PDF(I) Then
        '        '          'this was not a global peak
        '        '          I = J - 1
        '        '          Exit For
        '        '        End If
        '        '      Next
        '        '      If J > StopPoint Then
        '        '        'this is the global peak
        '        '        PDFNoiseLevel = I + 3
        '        '        Exit For
        '        '      End If
        '        '    End If
        '        '  Next
        '        '  Occupancy = CDF(PDFNoiseLevel) / (StopPoint - StartPoint + 1)
        '        '#End If

        '        '~~~~~~~~~~~~~~``





        '        'Debug.Print (PDFNoiseLevel - NoiseLevel)

        '        NoiseLevel = CLng(CalculatedNoiseLevel)
      End Function
    End Class
  End Namespace
#End Region
End Namespace