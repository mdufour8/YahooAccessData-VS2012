Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics
Imports YahooAccessData.OptionValuation
Imports System.Runtime.CompilerServices
Imports YahooAccessData.MathPlus.Measure.Measure

Namespace OptionValuation
  <Serializable()>
  Public Class StockOption
    Implements ICloneable
    Implements IStockOption
    Implements YahooAccessData.IMessageInfoEvents

    Private MyDividendPaymentPeriodType As IStockOption.enuDividendPaymentPeriodType
    Private MySymbol As String
    Private MyStockPrice As Double
    Private MyOptionPriceDelta As Double
    Private MyOptionPriceVega As Double    'variation in function of volatility for 1% change of volatility
    Private MyOptionPriceTheta As Double    'the decay rate of the option
    Private MyStrikePrice As Double
    Private MyVolatility As Double
    Private MyVolatilityStandard As Double
    Private MyVolatilityStandardImplied As Double
    Private MyBasicRateRiskFree As Double
    Private MyDateExpiration As Date
    Private MyGain As Double
    Private MyGainAtSigma As Double
    Private MyDividend As Double
    Private MyOptionType As Measure.enuOptionType
    Private MyValueOptionAtSigma As Double
    Private MyValueOption As Double
    Private MyStockPriceToExpiration As Double
    Private MyStockPriceToExpirationAtSigma As Double
    Private MyStockPriceSigmaToExpiration As Double
    Private MyValueDeltaFromGain As Double
    Private MyValueOptionStandard As Double
    Private MyValueOptionFromGainAtSigma As Double
    Private MyVolatilityStandardYearlyType As IStockOption.enuVolatilityStandardYearlyType
    Private MyDateStart As Date
    Private MyDateBuy As Date
    Private MyDateClose As Date
    Private MyPriceBuy As Double
    Private MyPriceClose As Double
    Private MyExDividendDateEstimated As Date
    Private MyExDividendDateDeclared As Date
    Private MyDividendPaymentNumber As Integer
    Private MyNumberOfOptionContract As Integer
    Public Event Message(Message As String, MessageType As YahooAccessData.IMessageInfoEvents.EnuMessageType) Implements YahooAccessData.IMessageInfoEvents.Message

    Public Sub New()
      Me.New("")
    End Sub

    Public Sub New(ByVal Symbol As String)
      Me.DateStart = Now.Date
      Me.DateExpiration = Me.DateStart
      Me.DateBuy = Me.DateStart
      Me.DateClose = Me.DateStart
      Me.Symbol = Symbol
      Me.VolatilityStandardYearlyType = IStockOption.enuVolatilityStandardYearlyType.YearlyMonthly
      Me.OptionType = Measure.enuOptionType._Call
      Me.OptionStyle = IStockOption.enuOptionStyle.American
      Me.ExDividendDateDeclared = YahooAccessData.ReportDate.DateNullValue
      Me.ExDividendDateEstimated = YahooAccessData.ReportDate.DateNullValue
      Me.IsExDividendDateEnabled = True
      Me.VolatilityMeasurementMethod = VolatilityMeasurementMethodDefault()
      Me.NumberOptionContract = 1
      MyDividendPaymentPeriodType = DividendPaymentPeriodTypeDefault()
    End Sub

    Public Sub New(ByVal StockOption As StockOption)
      Me.CopyTo(StockOption)
    End Sub

    Public Property Symbol As String Implements IStockOption.Symbol
      Get
        Return MySymbol
      End Get
      Set(value As String)
        MySymbol = value
      End Set
    End Property

    ''' <summary>
    ''' Represent the Number of day of the option since the contract bought time. This value is 
    ''' a fixed value use for portfolio tracking over time and is changing only when when the DateStart parameters is changed.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property NumberDayOfContract As Double
      Get
        Return (Me.DateStart - Me.DateBuy).TotalDays
      End Get
    End Property

    ''' <summary>
    ''' Represent the Number of option contract at the buy time. This value is 
    ''' a fixed value use for portfolio tracking over time and is not changing with any parameters of the option.
    ''' </summary>
    ''' <returns></returns>
    Public Property NumberOptionContract As Integer
      Get
        Return MyNumberOfOptionContract
      End Get
      Set(value As Integer)
        MyNumberOfOptionContract = value
      End Set
    End Property

    ''' <summary>
    ''' Represent the value of the stock or the option (at the user discretion) at the buy time. This value is 
    ''' a fixed value use for portfolio tracking over time and is not changing with any parameters of the option.
    ''' </summary>
    ''' <returns></returns>
    Public Property PriceBuy As Double Implements IStockOption.PriceBuy
      Get
        Return MyPriceBuy
      End Get
      Set(value As Double)
        MyPriceBuy = value
      End Set
    End Property

    ''' <summary>
    ''' Represent the value of the stock or the option (at the user discretion) at the sale time. This value is 
    ''' a fixed value use for portfolio tracking over time and is not changing with any parameters of the option.
    ''' </summary>
    ''' <returns></returns>
    Public Property PriceCLose As Double Implements IStockOption.PriceCLose
      Get
        Return MyPriceClose
      End Get
      Set(value As Double)
        MyPriceClose = value
      End Set
    End Property

    ''' <summary>
    ''' Represent the date of the stock or the option (at the user discretion) at the buy time. This value is 
    ''' a fixed value use for portfolio tracking over time and is not changing with any parameters of the option.
    ''' </summary>
    ''' <returns></returns>
    Public Property DateBuy As Date Implements IStockOption.DateBuy
      Get
        If MyDateBuy = Date.MinValue Then
          MyDateBuy = Now
        End If
        Return MyDateBuy
      End Get
      Set(value As Date)
        MyDateBuy = value
      End Set
    End Property

    ''' <summary>
    ''' Represent the value of the stock or the option (at the user discretion) at the sale time. This value is 
    ''' a fixed value use for portfolio tracking over time and is not changing with any parameters of the option.
    ''' </summary>
    ''' <returns></returns>
    Public Property DateClose As Date Implements IStockOption.DateClose
      Get
        If MyDateClose = Date.MinValue Then
          MyDateClose = Now
        End If
        Return MyDateClose
      End Get
      Set(value As Date)
        MyDateClose = value
      End Set
    End Property

    Public Property DateExpiration As Date Implements IStockOption.DateExpiration
      Get
        If MyDateExpiration = Date.MinValue Then
          MyDateExpiration = Now
        End If
        Return MyDateExpiration
      End Get
      Set(value As Date)
        MyDateExpiration = value.Date
      End Set
    End Property

    Public Property OptionType As YahooAccessData.MathPlus.Measure.Measure.enuOptionType Implements IStockOption.OptionType
      Get
        Return MyOptionType
      End Get
      Set(value As YahooAccessData.MathPlus.Measure.Measure.enuOptionType)
        MyOptionType = value
      End Set
    End Property

    Public Property RateBase As Double Implements IStockOption.RateBase
      Get
        Return MyBasicRateRiskFree
      End Get
      Set(value As Double)
        MyBasicRateRiskFree = value
      End Set
    End Property

    Public Property Gain As Double Implements IStockOption.Gain
      Get
        Return MyGain
      End Get
      Set(value As Double)
        MyGain = value
      End Set
    End Property

    ''' <summary>
    ''' The current stock price value
    ''' </summary>
    ''' <returns></returns>
    Public Property Price As Double Implements IStockOption.Price
      Get
        Return MyStockPrice
      End Get
      Set(value As Double)
        MyStockPrice = value
      End Set
    End Property

    ''' <summary>
    ''' The strike price value of the option
    ''' </summary>
    ''' <returns></returns>
    Public Property StrikePrice As Double Implements IStockOption.StrikePrice
      Get
        Return MyStrikePrice
      End Get
      Set(value As Double)
        MyStrikePrice = value
      End Set
    End Property

    Public Property VolatilityMeasurementMethod As IStockOption.enuVolatilityMeasurementMethod Implements IStockOption.VolatilityMeasurementMethod

    ''' <summary>
    ''' The volatility of the stock
    ''' </summary>
    ''' <returns></returns>
    Public Property Volatility As Double Implements IStockOption.Volatility
      Get
        Return MyVolatility
      End Get
      Set(value As Double)
        MyVolatility = value
      End Set
    End Property

    Public Function Refresh(ByVal StockPrice As Double) As Double
      Me.Price = StockPrice
      Return Me.Refresh()
    End Function

    Public Function Refresh(DateToday As Date) As Double Implements IStockOption.Refresh
      Me.DateStart = DateToday
      Return Me.Refresh()
    End Function

    Public Function Refresh(DateToday As Date, ByVal Price As Double) As Double Implements IStockOption.Refresh
      Me.DateStart = DateToday
      Me.Price = Price
      Return Me.Refresh()
    End Function

    Public Function Refresh() As Double Implements IStockOption.Refresh
      Dim ThisValueDeltaFromGainAtSigma As Double
      Dim ThisDividendPriceReduction As Double
      Dim ThisDividendPayment As Double
      Dim ThisTimeToExDividendExpiration As Double
      Dim ThisRateDividend As Double = 0
      Dim ThisVolatilityAdjustedStandard As Double
      Dim ThisVolatilityAdjustedGain As Double
      Dim ThisVolatilityAdjustedRatio As Double
      Dim ThisDividendPayments As New List(Of Double)
      Dim ThisDividendTimeToExDividend As New List(Of Double)
      Dim ThisValueOptionDeltaPerDay As Double
      Dim ThisTimeToExpiration As Double
      Dim ThisTimeToExpirationPlusOneDay As Double
      'Dim ThisLogNormalSigma As Double
      'Dim ThisLogNormalMean As Double
      'Dim ThisGainAtSigma As Double

      ThisTimeToExpiration = Me.TimeToExpiration(IStockOption.enuTimeToExpirationScale.Year)
      'add 1 day
      ThisTimeToExpirationPlusOneDay = Me.TimeToExpiration(Me.DateStart.AddDays(+1), IStockOption.enuTimeToExpirationScale.Year)
      'check the date Buy and close
      If Me.DateBuy > Me.DateExpiration Then
        'the current evalution date
        Me.DateBuy = Me.DateStart
      End If
      If Me.DateBuy > Me.DateClose Then
        Me.DateClose = Me.DateBuy
      End If
      If Me.DateClose > Me.DateExpiration Then
        Me.DateClose = Me.DateExpiration
      End If

      ThisDividendPayment = Me.DividendPaymentValue
      If Me.VolatilityStandard = 0 Then
        'message is not needed because it can keep blocking the user if it an old stock that has no data
        'RaiseEvent Message("Standard Volatility is zero!", YahooAccessData.IMessageInfoEvents.enuMessageType.Warning)
        Return 0.0
      End If
      If ThisDividendPayment <> 0 Then
        'validate the dividend date
        If Me.IsExDividendDateEnabled Then
          If Me.IsExDividendDateDeclaredEnabled Then
            If Me.ExDividendDateDeclared <= YahooAccessData.ReportDate.DateNullValue Then
              RaiseEvent Message("Invalid exdividend date!", YahooAccessData.IMessageInfoEvents.EnuMessageType.Warning)
              Return 0.0
            End If
          Else
            If Me.ExDividendDateEstimated <= YahooAccessData.ReportDate.DateNullValue Then
              RaiseEvent Message("Invalid exdividend date!", YahooAccessData.IMessageInfoEvents.EnuMessageType.Warning)
              Return 0.0
            End If
          End If
        End If
      End If
      MyDividendPaymentNumber = 0
      If Me.IsExDividendDateEnabled Then
        'following the discrete Escrowed Dividend Model (Mer73)
        Dim ThisExDividendDate As Date
        Dim ThisNumberOfDay As Integer

        If Me.IsExDividendDateDeclaredEnabled Then
          ThisExDividendDate = Me.ExDividendDateDeclared
        Else
          ThisExDividendDate = Me.ExDividendDateEstimated
        End If
        'check the number of day between the datestart and the exdividend and quickly adjust if the date is less than 
        Dim ThisNumberOfDaysFromDateStart = ThisExDividendDate.Subtract(Me.DateStart).Days
        Dim ThisDividendPaymentPeriod = Me.DividendPaymentPeriod
        If ThisDividendPaymentPeriod > 0 Then
          If ThisNumberOfDaysFromDateStart < 0 Then
            'quickly readjust the date to be close to DateStart and make sure the following loop execute quickly
            Dim ThisNumberOfPeriodBeforeDateStart = -ThisNumberOfDaysFromDateStart \ ThisDividendPaymentPeriod
            ThisNumberOfDay = ThisNumberOfPeriodBeforeDateStart * ThisDividendPaymentPeriod
            ThisExDividendDate = ThisExDividendDate.AddDays(ThisNumberOfDay)
          End If
          ThisNumberOfDay = 0
          ThisDividendPriceReduction = 0
          Do
            ThisExDividendDate = ThisExDividendDate.AddDays(ThisNumberOfDay)
            If ThisExDividendDate > Me.DateExpiration Then Exit Do
            If ThisExDividendDate > Me.DateStart Then
              ThisTimeToExDividendExpiration = StockOption.TimeToExDividend(Me.DateStart, ThisExDividendDate, IStockOption.enuTimeToExpirationScale.Year)
              ThisDividendTimeToExDividend.Add(ThisTimeToExDividendExpiration)
              ThisDividendPayments.Add(ThisDividendPayment)
              ThisDividendPriceReduction = ThisDividendPriceReduction + ThisDividendPayment * Math.Exp(-Me.RateBase * ThisTimeToExDividendExpiration)
              MyDividendPaymentNumber = MyDividendPaymentNumber + 1
            End If
            ThisNumberOfDay = ThisNumberOfDay + ThisDividendPaymentPeriod
          Loop
          ThisRateDividend = 0.0
        End If
      Else
        ThisRateDividend = Me.RateDividend
        ThisDividendPriceReduction = 0
        MyDividendPaymentNumber = 0
      End If
      If Me.IsVolatilityStandardImpliedEnabled Then
        ThisVolatilityAdjustedStandard = Me.VolatilityStandardImplied
      Else
        ThisVolatilityAdjustedStandard = Me.VolatilityStandard
      End If
      ThisVolatilityAdjustedGain = Me.Volatility

      'from : 
      'https://en.wikipedia.org/wiki/Geometric_Brownian_motion
      'for lab experiment see this excellent demo: 
      'http://www.math.uah.edu/stat/apps/GeometricBrownianMotion.html
      'See also for more general interest on statistic analysis and definition:
      'http://www.math.uah.edu/stat/brown/index.html

      'ThisLogNormalSigma = ThisVolatilityAdjustedGain * Math.Sqrt(ThisTimeToExpiration)
      'ThisLogNormalMean = ThisTimeToExpiration * (MyGain - ThisVolatilityAdjustedGain ^ 2 / 2)
      MyStockPriceToExpiration = MyStockPrice * Math.Exp(ThisTimeToExpiration * MyGain)
      MyStockPriceSigmaToExpiration = Math.Sqrt((MyStockPrice ^ 2) * (Math.Exp(2 * MyGain * ThisTimeToExpiration) * (Math.Exp(ThisTimeToExpiration * ThisVolatilityAdjustedGain ^ 2) - 1)))
      MyValueDeltaFromGain = MyStockPriceToExpiration - MyStockPrice
      If Me.OptionType = Measure.enuOptionType._Call Then
        MyStockPriceToExpirationAtSigma = MyStockPriceToExpiration - MyStockPriceSigmaToExpiration
      Else
        MyStockPriceToExpirationAtSigma = MyStockPriceToExpiration + MyStockPriceSigmaToExpiration
      End If
      ThisValueDeltaFromGainAtSigma = MyStockPriceToExpirationAtSigma - MyStockPrice
      'MyGainAtSigma is not needed
      'If ThisTimeToExpiration > 0 Then
      '  MyGainAtSigma = MyGain - (Math.Log(MyStockPriceToExpiration / MyStockPriceToExpirationAtSigma)) / ThisTimeToExpiration
      'Else
      '  MyGainAtSigma = MyGain
      'End If

      'Dim LogNormalDist As Distributions.LogNormal = New Distributions.LogNormal(Mu, Sigma)


      If ThisDividendPayments.Count > 0 Then
        ThisVolatilityAdjustedRatio = (Measure.HaugHaugDividendVolatilityCorrection(
          MyStockPrice,
          ThisTimeToExpiration,
          MyBasicRateRiskFree,
          ThisDividendPayments.ToArray,
          ThisDividendTimeToExDividend.ToArray,
          ThisVolatilityAdjustedStandard)) / ThisVolatilityAdjustedStandard
      Else
        ThisVolatilityAdjustedRatio = 1.0
      End If
      Select Case Me.OptionStyle
        Case IStockOption.enuOptionStyle.American
          'changing the strike price seem to be the best method to evaluate the effect of the gain
          MyValueOption = Measure.BSAmericanOption(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice - MyValueDeltaFromGain,
            ThisTimeToExpiration,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedGain)

          MyValueOptionAtSigma = Measure.BSAmericanOption(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice - ThisValueDeltaFromGainAtSigma,
            ThisTimeToExpiration,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedGain)


          MyValueOptionStandard = Measure.BSAmericanOption(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice,
            ThisTimeToExpiration,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard)

          'mesure the with the standard method using no directional gain
          ThisValueOptionDeltaPerDay = Measure.BSAmericanOption(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice,
            ThisTimeToExpirationPlusOneDay,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard)
        Case IStockOption.enuOptionStyle.Europeen
          'MyValueOption = Measure.BlackScholes(
          '  MyOptionType,
          '  MyStockPrice - ThisDividendPriceReduction,
          '  MyStrikePrice - MyValueDeltaFromGain,
          '  ThisTimeToExpiration,
          '  MyBasicRateRiskFree,
          '  ThisRateDividend,
          '  Me.Volatility)

          MyValueOption = Measure.BlackScholes(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice,
            ThisTimeToExpiration,
            MyBasicRateRiskFree + MyGain,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedGain)

          MyValueOptionStandard = Measure.BlackScholes(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice,
            ThisTimeToExpiration,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard)

          'mesure the with the standard method using no directional gain
          ThisValueOptionDeltaPerDay = Measure.BlackScholes(
            MyOptionType,
            MyStockPrice - ThisDividendPriceReduction,
            MyStrikePrice,
            ThisTimeToExpirationPlusOneDay,
            MyBasicRateRiskFree,
            ThisRateDividend,
            ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard)
      End Select
      MyOptionPriceTheta = ThisValueOptionDeltaPerDay - MyValueOptionStandard
      'Use the BlackScholes estimate for these value
      'the difference in value should not be very large
      MyOptionPriceDelta = Measure.BlackScholesOptionDelta(
        MyOptionType,
        StockPrice:=(MyStockPrice - ThisDividendPriceReduction),
        OptionStrikePrice:=MyStrikePrice,
        TimeToExpirationInYear:=ThisTimeToExpiration,
        RiskFreeRate:=MyBasicRateRiskFree,
        DividendRate:=ThisRateDividend,
        VolatilityPerYear:=(ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard))

      MyOptionPriceVega = Measure.BlackScholesOptionVega(
        StockPrice:=(MyStockPrice - ThisDividendPriceReduction),
        OptionStrikePrice:=MyStrikePrice,
        TimeToExpirationInYear:=ThisTimeToExpiration,
        RiskFreeRate:=MyBasicRateRiskFree,
        DividendRate:=ThisRateDividend,
        VolatilityPerYear:=(ThisVolatilityAdjustedRatio * ThisVolatilityAdjustedStandard))

      Return MyValueOption
    End Function

    Public Function AsIStockOption() As IStockOption Implements IStockOption.AsIStockOption
      Return Me
    End Function



    ''' <summary>
    ''' Return by default the time to expiration in years
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TimeToExpiration(Optional TimeToExpirationScale As IStockOption.enuTimeToExpirationScale = IStockOption.enuTimeToExpirationScale.Year) As Double Implements IStockOption.TimeToExpiration
      Select Case TimeToExpirationScale
        Case IStockOption.enuTimeToExpirationScale.Year
          Return MyDateExpiration.Date.Subtract(Me.DateStart).TotalDays / 365
        Case IStockOption.enuTimeToExpirationScale.Day
          Return MyDateExpiration.Subtract(Me.DateStart).TotalDays
        Case IStockOption.enuTimeToExpirationScale.Hour
          Return MyDateExpiration.Subtract(Me.DateStart).TotalHours
        Case IStockOption.enuTimeToExpirationScale.TradingDay
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR * MyDateExpiration.Subtract(Me.DateStart.Date).TotalDays / 365
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Function TimeToExpiration(ByVal DateStart As Date, Optional TimeToExpirationScale As IStockOption.enuTimeToExpirationScale = IStockOption.enuTimeToExpirationScale.Year) As Double
      Select Case TimeToExpirationScale
        Case IStockOption.enuTimeToExpirationScale.Year
          Return MyDateExpiration.Subtract(DateStart).TotalDays / 365
        Case IStockOption.enuTimeToExpirationScale.Day
          Return MyDateExpiration.Subtract(DateStart).TotalDays
        Case IStockOption.enuTimeToExpirationScale.Hour
          Return MyDateExpiration.Subtract(DateStart).TotalHours
        Case IStockOption.enuTimeToExpirationScale.TradingDay
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR * Me.DateExpiration.Subtract(DateStart.Date).TotalDays / 365
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Shared Function TimeToExpiration(ByVal DateStart As Date, ByVal DateOfExpiration As Date, Optional TimeToExpirationScale As IStockOption.enuTimeToExpirationScale = IStockOption.enuTimeToExpirationScale.Year) As Double
      Select Case TimeToExpirationScale
        Case IStockOption.enuTimeToExpirationScale.Year
          Return DateOfExpiration.Subtract(DateStart).TotalDays / 365
        Case IStockOption.enuTimeToExpirationScale.Day
          Return DateOfExpiration.Subtract(DateStart).TotalDays
        Case IStockOption.enuTimeToExpirationScale.Hour
          Return DateOfExpiration.Subtract(DateStart).TotalHours
        Case IStockOption.enuTimeToExpirationScale.TradingDay
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR * DateOfExpiration.Subtract(DateStart).TotalDays / 365
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Function TimeToExDividend(Optional TimeToExpirationScale As IStockOption.enuTimeToExpirationScale = IStockOption.enuTimeToExpirationScale.Day) As Double Implements IStockOption.TimeToExDividend
      If Me.IsExDividendDateDeclaredEnabled Then
        Return StockOption.TimeToExDividend(Me.DateStart, Me.ExDividendDateDeclared, TimeToExpirationScale)
      Else
        Return StockOption.TimeToExDividend(Me.DateStart, Me.ExDividendDateEstimated, TimeToExpirationScale)
      End If
    End Function

    Public Shared Function TimeToExDividend(ByVal DateStart As Date, ByVal ExDividendDate As Date, Optional TimeToExpirationScale As IStockOption.enuTimeToExpirationScale = IStockOption.enuTimeToExpirationScale.Day) As Double
      Dim ThisTimeToExDividend As Double
      ThisTimeToExDividend = StockOption.TimeToExpiration(DateStart.Date, ExDividendDate.Date, TimeToExpirationScale)
      If ThisTimeToExDividend < 0 Then
        ThisTimeToExDividend = 0
      End If
      Return ThisTimeToExDividend
    End Function

    Public Shared Function VolatilityTypeToNumberDays(ByVal DateStart As Date, ByVal DateOfExpiration As Date, ByVal VolatilityType As IStockOption.enuVolatilityStandardYearlyType) As Integer
      Select Case VolatilityType
        Case IStockOption.enuVolatilityStandardYearlyType.Daily10
          Return 10
        Case IStockOption.enuVolatilityStandardYearlyType.Daily15
          Return 15
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 6
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 4
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 2
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly, IStockOption.enuVolatilityStandardYearlyType.YearlyMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          Return CInt(TimeToExpiration(DateStart, DateOfExpiration, IStockOption.enuTimeToExpirationScale.Day))
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Shared Function VolatilityTypeToNumberDays(ByVal VolatilityType As IStockOption.enuVolatilityStandardYearlyType) As Integer
      Select Case VolatilityType
        Case IStockOption.enuVolatilityStandardYearlyType.YearlyDaily10, IStockOption.enuVolatilityStandardYearlyType.Daily10
          Return 10
        Case IStockOption.enuVolatilityStandardYearlyType.YearlyDaily15, IStockOption.enuVolatilityStandardYearlyType.Daily15
          Return 15
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly, IStockOption.enuVolatilityStandardYearlyType.YearlyMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 6
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 4
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 2
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Shared Function VolatilityTypeToNumberDays(ByVal VolatilityType As IStockOption.enuVolatilityYearlyType) As Integer
      Select Case VolatilityType
        Case IStockOption.enuVolatilityYearlyType.YearlyDaily10
          Return 10
        Case IStockOption.enuVolatilityYearlyType.YearlyDaily15
          Return 15
        Case IStockOption.enuVolatilityYearlyType.Monthly
          Return StockOption.VolatilityTypeToNumberDays(IStockOption.enuVolatilityStandardYearlyType.Monthly)
        Case IStockOption.enuVolatilityYearlyType.Quaterly
          Return StockOption.VolatilityTypeToNumberDays(IStockOption.enuVolatilityStandardYearlyType.Quaterly)
        Case IStockOption.enuVolatilityYearlyType.BiAnnual
          Return StockOption.VolatilityTypeToNumberDays(IStockOption.enuVolatilityStandardYearlyType.BiAnnual)
        Case IStockOption.enuVolatilityYearlyType.Yearly, IStockOption.enuVolatilityYearlyType.YearlyMonthly
          Return StockOption.VolatilityTypeToNumberDays(IStockOption.enuVolatilityStandardYearlyType.Yearly)
        Case Else
          Throw New NotSupportedException
      End Select
    End Function

    Public Shared Function VolatilityTypeFromNumberDays(ByVal NumberDays As Integer) As IStockOption.enuVolatilityStandardYearlyType
      Dim ThisNumberDaysPerMonth As Integer = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
      Select Case NumberDays
        Case Is < 15
          Return IStockOption.enuVolatilityStandardYearlyType.Daily10
        Case 15 To 20
          Return IStockOption.enuVolatilityStandardYearlyType.Daily15
        Case Is < 2 * ThisNumberDaysPerMonth
          Return IStockOption.enuVolatilityStandardYearlyType.Monthly
        Case Is < 5 * ThisNumberDaysPerMonth
          Return IStockOption.enuVolatilityStandardYearlyType.Quaterly
        Case Is < 8 * ThisNumberDaysPerMonth
          Return IStockOption.enuVolatilityStandardYearlyType.BiAnnual
        Case Else
          Return IStockOption.enuVolatilityStandardYearlyType.Yearly
      End Select
    End Function
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="DateStart">The current date</param>
    ''' <param name="OptionStyle"></param>
    ''' <param name="NumberOfMonth">Number of Month to add to teh current date</param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Traditional monthly American options expire the third Saturday of every month. They are closed for trading the Friday prior. *Expire the third Friday if the first of the month begins on a Saturday.
    ''' European options expire the Friday prior to the third Saturday of every month. Therefore, they are closed for trading the Thursday prior to the third Saturday of every month.
    ''' INVESTOPEDIA EXPLAINS 'Expiration Date (Derivatives)
    ''' The expiration date for all listed stock options in the United States is normally the third Friday of the contract month, 
    ''' which is the month when the contract expires. However, when that Friday falls on a holiday, 
    ''' the expiration date is on the Thursday immediately before the third Friday. 
    ''' Once an options or futures contract passes its expiration date, the contract is invalid.
    ''' </remarks>
    Public Shared Function DateOfExpirationFromMonth(ByVal DateStart As Date, ByVal OptionStyle As IStockOption.enuOptionStyle, ByVal NumberOfMonth As Integer) As Date
      Dim ThisDateEnd As Date

      ThisDateEnd = New DateTime(DateStart.Year, DateStart.Month, 1).AddMonths(NumberOfMonth)
      Select Case OptionStyle
        Case IStockOption.enuOptionStyle.American
          Select Case ThisDateEnd.DayOfWeek
            Case DayOfWeek.Monday
              ThisDateEnd = ThisDateEnd.AddDays(18)
            Case DayOfWeek.Tuesday
              ThisDateEnd = ThisDateEnd.AddDays(17)
            Case DayOfWeek.Wednesday
              ThisDateEnd = ThisDateEnd.AddDays(16)
            Case DayOfWeek.Thursday
              ThisDateEnd = ThisDateEnd.AddDays(15)
            Case DayOfWeek.Friday
              ThisDateEnd = ThisDateEnd.AddDays(14)
            Case DayOfWeek.Saturday
              'Note:Expire the third Friday if the first of the month begins on a Saturday.
              ThisDateEnd = ThisDateEnd.AddDays(20)
            Case DayOfWeek.Sunday
              ThisDateEnd = ThisDateEnd.AddDays(19)
          End Select
        Case IStockOption.enuOptionStyle.Europeen
          Select Case ThisDateEnd.DayOfWeek
            Case DayOfWeek.Monday
              ThisDateEnd = ThisDateEnd.AddDays(18)
            Case DayOfWeek.Tuesday
              ThisDateEnd = ThisDateEnd.AddDays(17)
            Case DayOfWeek.Wednesday
              ThisDateEnd = ThisDateEnd.AddDays(16)
            Case DayOfWeek.Thursday
              ThisDateEnd = ThisDateEnd.AddDays(15)
            Case DayOfWeek.Friday
              ThisDateEnd = ThisDateEnd.AddDays(14)
            Case DayOfWeek.Saturday
              'Note:Expire the third Friday if the first of the month begins on a Saturday.
              ThisDateEnd = ThisDateEnd.AddDays(20)
            Case DayOfWeek.Sunday
              ThisDateEnd = ThisDateEnd.AddDays(19)
          End Select
      End Select
      ThisDateEnd = ThisDateEnd.Date.AddSeconds(ReportDate.MARKET_CLOSE_TIME_SEC_DEFAULT)
      Dim ThisDateEndResult = YahooAccessData.ReportDate.DayOfTrade(ThisDateEnd)
      Return ThisDateEndResult.Date
    End Function

    Public Shared Function DateOfExpirationFromMonth(ByVal DateStart As Date, ByVal OptionStyle As IStockOption.enuOptionStyle, ByVal Month As YahooAccessData.ReportDate.MonthsOfYear) As Date
      Dim ThisMonth = DateStart.Month
      Dim ThisMonthOfExpiration = CInt(Month)
      Dim ThisMonthToAdd As Integer

      ThisMonthToAdd = ThisMonthOfExpiration - ThisMonth
      If ThisMonthToAdd <= 0 Then
        ThisMonthToAdd = ThisMonthToAdd + 12
      Else
        ThisMonthToAdd = ThisMonthOfExpiration - ThisMonth
      End If
      Return DateOfExpirationFromMonth(DateStart, OptionStyle, ThisMonthToAdd)
    End Function

    Public Property VolatilityStandardImplied As Double
      Get
        Return MyVolatilityStandardImplied
      End Get
      Set(value As Double)
        MyVolatilityStandardImplied = value
      End Set
    End Property

    Public Property IsVolatilityStandardImpliedEnabled As Boolean Implements IStockOption.IsVolatilityStandardImpliedEnabled

    Public Shared Function CalculateVolatilityStandardImplied(ByVal PriceOfOption As Double) As Double
      Throw New NotImplementedException
    End Function


    Public Function VolatilityTotal() As Double Implements IStockOption.VolatilityTotal
      Dim ThisVolatilityWithGain = Math.Sqrt(MyVolatility ^ 2 + MyGain ^ 2)

      Dim ThisVolatilityPrediction = YahooAccessData.MathPlus.Measure.Measure.VolatilityPrediction(Me.VolatilityFilterRate, Me.Gain, Me.Volatility)

      If Me.OptionType = Measure.enuOptionType._Call Then
        If Me.Gain > 0 Then
          Return ThisVolatilityWithGain
        Else
          Return MyVolatility
        End If
      Else
        If Me.Gain < 0 Then
          Return ThisVolatilityWithGain
        Else
          Return MyVolatility
        End If
      End If
    End Function

    Public Function Copy() As StockOption
      Dim ThisStockOption As New StockOption
      With ThisStockOption
        .DateStart = Me.DateStart
        .DateExpiration = Me.DateExpiration
        .DateBuy = Me.DateBuy
        .DateClose = Me.DateClose
        .PriceBuy = Me.PriceBuy
        .PriceCLose = Me.PriceCLose
        .NumberOptionContract = Me.NumberOptionContract
        .OptionType = Me.OptionType
        .OptionStyle = Me.OptionStyle
        .RateBase = Me.RateBase
        .Gain = Me.Gain
        .Price = Me.Price
        .StrikePrice = Me.StrikePrice
        .Volatility = Me.Volatility
        .VolatilityStandard = Me.VolatilityStandard
        .VolatilityStandardImplied = Me.VolatilityStandardImplied
        .IsVolatilityStandardImpliedEnabled = Me.IsVolatilityStandardImpliedEnabled
        .VolatilityStandardYearlyType = Me.VolatilityStandardYearlyType
        .RateDividend = Me.RateDividend
        .GainSigmaError = Me.GainSigmaError
        .GainSigmaErrorPeriod = Me.GainSigmaErrorPeriod
        .IsGainSigmaErrorEnabled = Me.IsGainSigmaErrorEnabled
        .NumberOptionContract = Me.NumberOptionContract

        .DividendPaymentPeriodType = Me.DividendPaymentPeriodType
        .DividendPaymentPeriodDetected = Me.DividendPaymentPeriodDetected

        .DividendPaymentPeriodType = Me.DividendPaymentPeriodType
        .ExDividendDateDeclared = Me.ExDividendDateDeclared
        .ExDividendDateEstimated = Me.ExDividendDateEstimated
        .IsExDividendDateEnabled = Me.IsExDividendDateEnabled
      End With
      Return ThisStockOption
    End Function

    Public Sub CopyTo(ByVal StockOption As StockOption)
      If StockOption Is Nothing Then
        StockOption = New StockOption
      End If
      With Me
        .DateStart = StockOption.DateStart
        .DateBuy = StockOption.DateBuy
        .DateClose = StockOption.DateClose
        .PriceBuy = StockOption.PriceBuy
        .PriceCLose = StockOption.PriceCLose
        .NumberOptionContract = StockOption.NumberOptionContract
        .DateExpiration = StockOption.DateExpiration
        .OptionType = StockOption.OptionType
        .OptionStyle = StockOption.OptionStyle
        .RateBase = StockOption.RateBase
        .Gain = StockOption.Gain
        .Price = StockOption.Price
        .StrikePrice = StockOption.StrikePrice
        .VolatilityStandardImplied = StockOption.VolatilityStandardImplied
        .IsVolatilityStandardImpliedEnabled = StockOption.IsVolatilityStandardImpliedEnabled
        .Volatility = StockOption.Volatility
        .VolatilityStandard = StockOption.VolatilityStandard
        .VolatilityStandardYearlyType = StockOption.VolatilityStandardYearlyType
        .RateDividend = StockOption.RateDividend
        .GainSigmaError = StockOption.GainSigmaError
        .GainSigmaErrorPeriod = StockOption.GainSigmaErrorPeriod
        .IsGainSigmaErrorEnabled = StockOption.IsGainSigmaErrorEnabled
        .NumberOptionContract = StockOption.NumberOptionContract
        .DividendPaymentPeriodType = StockOption.DividendPaymentPeriodType
        .DividendPaymentPeriodDetected = StockOption.DividendPaymentPeriodDetected

        .ExDividendDateDeclared = StockOption.ExDividendDateDeclared
        .ExDividendDateEstimated = StockOption.ExDividendDateEstimated
        .IsExDividendDateEnabled = StockOption.IsExDividendDateEnabled
      End With
    End Sub

    Public Function Clone() As Object Implements ICloneable.Clone
      Return Me.Copy
    End Function

    Public Property VolatilityStandard As Double Implements IStockOption.VolatilityStandard
      Get
        Return MyVolatilityStandard
      End Get
      Set(value As Double)
        MyVolatilityStandard = value
      End Set
    End Property

    Public Shared Function VolatilityStandardYearlyTypeDefault() As IStockOption.enuVolatilityStandardYearlyType
      Return IStockOption.enuVolatilityStandardYearlyType.Yearly
    End Function

    Public Shared Function OptionStyleDefault() As IStockOption.enuOptionStyle
      Return IStockOption.enuOptionStyle.American
    End Function

    Public Property VolatilityStandardYearlyType As IStockOption.enuVolatilityStandardYearlyType Implements IStockOption.VolatilityStandardYearlyType
      Get
        Return MyVolatilityStandardYearlyType
      End Get
      Set(value As IStockOption.enuVolatilityStandardYearlyType)
        MyVolatilityStandardYearlyType = value
      End Set
    End Property

    Public Function VolatilityFilterRate() As Integer Implements IStockOption.VolatilityFilterRate
      Select Case Me.VolatilityStandardYearlyType
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 2
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 6
        Case IStockOption.enuVolatilityStandardYearlyType.Daily10
          Return 10
        Case IStockOption.enuVolatilityStandardYearlyType.Daily15
          Return 15
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 4
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          Dim ThisNumberOfDayToExpiration As Integer = CInt(Me.TimeToExpiration(IStockOption.enuTimeToExpirationScale.TradingDay))
          If ThisNumberOfDayToExpiration < YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12 Then
            ThisNumberOfDayToExpiration = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
          End If
          Return ThisNumberOfDayToExpiration
        Case Else
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      End Select
    End Function

    Public Shared Function VolatilityImpliedApproximatedEstimate(ByVal PriceOfOption As Double) As Double
			Throw New NotImplementedException
		End Function

    Public Shared Function VolatilityImpliedCloseEstimate(ByVal PriceOfOption As Double) As Double
      Throw New NotImplementedException
    End Function

    Public Shared Function VolatilityFilterRate(VolatilityStandardYearlyType As IStockOption.enuVolatilityStandardYearlyType) As Integer
      Select Case VolatilityStandardYearlyType
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 2
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 6
        Case IStockOption.enuVolatilityStandardYearlyType.Daily10
          Return 10
        Case IStockOption.enuVolatilityStandardYearlyType.Daily15
          Return 15
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 4
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          Throw New NotSupportedException
        Case Else
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      End Select
    End Function

    Public Shared Function VolatilityFilterRate(ByVal DateStart As Date, ByVal DateOfExpiration As Date, VolatilityStandardYearlyType As IStockOption.enuVolatilityStandardYearlyType) As Integer
      Select Case VolatilityStandardYearlyType
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 2
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 6
        Case IStockOption.enuVolatilityStandardYearlyType.Daily10
          Return 10
        Case IStockOption.enuVolatilityStandardYearlyType.Daily15
          Return 15
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 4
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          Dim ThisNumberOfDayToExpiration As Integer = CInt(TimeToExpiration(DateStart, DateOfExpiration, IStockOption.enuTimeToExpirationScale.TradingDay))
          If ThisNumberOfDayToExpiration < YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12 Then
            ThisNumberOfDayToExpiration = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
          End If
          Return ThisNumberOfDayToExpiration
        Case Else
          Return YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      End Select
    End Function

    Public Property RateDividend As Double Implements IStockOption.RateDividend
      Get
        Return MyDividend
      End Get
      Set(value As Double)
        MyDividend = value
      End Set
    End Property

    ''' <summary>
    ''' The current date 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DateStart As Date Implements IStockOption.DateStart
      Get
        If MyDateStart = Date.MinValue Then
          MyDateStart = Now.Date
        End If
        Return MyDateStart
      End Get
      Set(value As Date)
        MyDateStart = value.Date
      End Set
    End Property

    Public Property OptionStyle As IStockOption.enuOptionStyle Implements IStockOption.OptionStyle

    ''' <summary>
    ''' The number of days use to calculate the gain sigma error
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property GainSigmaErrorPeriod As Integer Implements IStockOption.GainSigmaErrorPeriod

    Private MyGainSigmaError As Double
    Public Property GainSigmaError As Double Implements IStockOption.GainSigmaError
      Get
        Return MyGainSigmaError
      End Get
      Set(value As Double)
        MyGainSigmaError = value
      End Set
    End Property

    Public Property IsGainSigmaErrorEnabled As Boolean Implements IStockOption.IsGainSigmaErrorEnabled

    ''' <summary>
    ''' Stock dividend payment period in days
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>The default is 91 days</remarks>
    Public ReadOnly Property DividendPaymentPeriod As Integer Implements IStockOption.DividendPaymentPeriod
      Get
        Select Case Me.DividendPaymentPeriodType
          Case IStockOption.enuDividendPaymentPeriodType.AutoDetect
            Return Me.DividendPaymentPeriodDetected
          Case IStockOption.enuDividendPaymentPeriodType.BiAnnual
            Return 365 \ 2
          Case IStockOption.enuDividendPaymentPeriodType.Monthly
            Return 365 \ 12
          Case IStockOption.enuDividendPaymentPeriodType.Quaterly
            Return 365 \ 4
          Case IStockOption.enuDividendPaymentPeriodType.Yearly
            Return 365
          Case Else
            Throw New NotSupportedException
        End Select
      End Get
    End Property

    Public ReadOnly Property DividendPaymentNumber As Integer Implements IStockOption.DividendPaymentNumber
      Get
        Return MyDividendPaymentNumber
      End Get
    End Property

    Public Property DividendPaymentPeriodType As IStockOption.enuDividendPaymentPeriodType Implements IStockOption.DividendPaymentPeriodType
      Get
        Return MyDividendPaymentPeriodType
      End Get
      Set(value As IStockOption.enuDividendPaymentPeriodType)
        MyDividendPaymentPeriodType = value
      End Set
    End Property

    Public Shared Function DividendPaymentPeriodTypeDefault() As IStockOption.enuDividendPaymentPeriodType
      Return IStockOption.enuDividendPaymentPeriodType.AutoDetect
    End Function

    Public Shared Function VolatilityMeasurementMethodDefault() As IStockOption.enuVolatilityMeasurementMethod
      Return IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
    End Function

    Public Property ExDividendDateDeclared As Date Implements IStockOption.ExDividendDateDeclared
      Get
        Return MyExDividendDateDeclared
      End Get
      Set(value As Date)
        MyExDividendDateDeclared = value.Date
      End Set
    End Property

    Public Property ExDividendDateEstimated As Date Implements IStockOption.ExDividendDateEstimated
      Get
        Return MyExDividendDateEstimated
      End Get
      Set(value As Date)
        MyExDividendDateEstimated = value.Date
      End Set
    End Property

    Public Property IsExDividendDateDeclaredEnabled As Boolean Implements IStockOption.IsExDividendDateDeclaredEnabled

    Public ReadOnly Property DividendPaymentValue As Double Implements IStockOption.DividendPaymentValue
      Get
        If Me.TimeToExDividend > 0 Then
          If Me.DividendPaymentPeriod > 0 Then
            Return (Me.DividendPaymentPeriod / 365) * (Me.RateDividend * Me.Price)
          Else
            Return Me.RateDividend * Me.Price
          End If
        Else
          Return 0
        End If
      End Get
    End Property

    Public ReadOnly Property RateDividendToPayment As Double Implements IStockOption.RateDividendToPayment
      Get
        Dim ThisTimeToExDividend As Integer
        If Me.IsExDividendDateEnabled Then
          ThisTimeToExDividend = CInt(Me.TimeToExDividend)
          If ThisTimeToExDividend > 0 Then
            Return (365 / ThisTimeToExDividend) * Me.DividendPaymentValue / Me.Price
          Else
            Return 0.0
          End If
        Else
          Return Me.RateDividend
        End If
      End Get
    End Property

    Public Property IsExDividendDateEnabled As Boolean Implements IStockOption.IsExDividendDateEnabled
    Public Property DividendPaymentPeriodDetected As Integer Implements IStockOption.DividendPaymentPeriodDetected

#Region "IStockOptionPrice"
    Public Function AsIStockOptionPrice() As IStockOptionPrice Implements IStockOptionPrice.AsIStockOptionPrice
      Return Me
    End Function

    ''' <summary>
    ''' Return the current value of the option with the expected stock gain
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property IStockOptionPrice_ValueFromGainPrediction As Double Implements IStockOptionPrice.ValueFromGainPrediction
      Get
        Return MyValueOption
      End Get
    End Property

    ''' <summary>
    ''' Return the current value with no expected gain. 
    ''' It correspond to the standard value that the average
    ''' participant to the market (On average the market market participant reflect an average gain of zero, 
    ''' otherwise the option value would change) .
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property IStockOptionPrice_ValueStandard As Double Implements IStockOptionPrice.ValueStandard
      Get
        Return MyValueOptionStandard
      End Get
    End Property

    Private ReadOnly Property IStockOptionPrice_ValueAtSigma As Double Implements IStockOptionPrice.ValueAtSigma
      Get
        Return MyValueOptionAtSigma
      End Get
    End Property

    Private ReadOnly Property IStockOptionPrice_ValueDelta As Double Implements IStockOptionPrice.ValueDelta
      Get
        Return MyOptionPriceDelta
      End Get
    End Property
#End Region
#Region "IStockPrice"
    Public Function AsIStockPrice() As IStockPrice Implements IStockPrice.AsIStockPrice
      Return Me
    End Function

    Private ReadOnly Property IStockPrice_Value As Double Implements IStockPrice.Value
      Get
        Return MyStockPriceToExpiration
      End Get
    End Property

    Private ReadOnly Property IStockPrice_ValueOfDeltaPrice As Double Implements IStockPrice.ValueOfDeltaPrice
      Get
        Return MyValueDeltaFromGain
      End Get
    End Property

    Private ReadOnly Property IStockPrice_ValueOfSigma As Double Implements IStockPrice.ValueOfSigma
      Get
        Return MyStockPriceSigmaToExpiration
      End Get
    End Property
#End Region
#Region "IStockPricePrediction"
    Public Function AsIStockPricePrediction() As IStockPricePrediction Implements IStockPricePrediction.AsIStockPricePrediction
      Return Me
    End Function

    Private Function IStockPricePrediction_Value(NumberTradingDays As Integer) As Double Implements IStockPricePrediction.Value
      Return MyStockPrice * Math.Exp(NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR * MyGain)
    End Function

    Private Function IStockPricePrediction_ValueOfSigma(NumberTradingDays As Integer) As Double Implements IStockPricePrediction.ValueOfSigma
      Return Math.Sqrt(IStockPricePrediction_Value(NumberTradingDays) ^ 2 * (Math.Exp((NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR) * Me.Volatility ^ 2) - 1))
    End Function

    ''' <summary>
    ''' This return the mean price
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="StockPrice"></param>
    ''' <param name="Gain"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StockPricePrediction(
                                               ByVal NumberTradingDays As Integer,
                                               ByVal StockPrice As Double,
                                               ByVal Gain As Double) As Double

      Return StockPricePrediction(CDbl(NumberTradingDays), StockPrice, Gain)
    End Function

    ''' <summary>
    ''' This return the mean price
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="StockPrice"></param>
    ''' <param name="Gain"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StockPricePrediction(
                                               ByVal NumberTradingDays As Double,
                                               ByVal StockPrice As Double,
                                               ByVal Gain As Double) As Double

      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Return StockPrice * Math.Exp(ThisTimeInYear * Gain)
    End Function

    ''' <summary>
    ''' see:  
    ''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
    ''' for lab experiment see this excellent demo: 
    ''' http://www.math.uah.edu/stat/apps/GeometricBrownianMotion.html
    ''' See also for more general interest on statistic analysis and definition:
    ''' http://www.math.uah.edu/stat/brown/index.html
    ''' https://en.wikipedia.org/wiki/Log-normal_distribution
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="StockPrice"></param>
    ''' <param name="Gain"></param>
    ''' <param name="Volatility"></param>
    ''' <param name="Probability"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StockPricePrediction(
                                               ByVal NumberTradingDays As Integer,
                                               ByVal StockPrice As Double,
                                               ByVal Gain As Double,
                                               ByVal Volatility As Double,
                                               ByVal Probability As Double) As Double

      Return StockPricePrediction(CDbl(NumberTradingDays), StockPrice, Gain, 0.0, Volatility, Probability)
    End Function

    Public Shared Function StockPricePrediction(
                                               ByVal NumberTradingDays As Double,
                                               ByVal StockPrice As Double,
                                               ByVal Gain As Double,
                                               ByVal Volatility As Double,
                                               ByVal Probability As Double) As Double

      Return StockPricePrediction(NumberTradingDays, StockPrice, Gain, 0.0, Volatility, Probability)
    End Function


		''' <summary>
		''' Calculate the expected stock price for a given number of trading day and probability
		''' see: 
		''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
		''' for lab experiment see this excellent demo: 
		''' http://www.math.uah.edu/stat/apps/GeometricBrownianMotion.html
		''' See also for more general interest on statistic analysis and definition:
		''' http://www.math.uah.edu/stat/brown/index.html
		''' https://en.wikipedia.org/wiki/Log-normal_distribution
		''' </summary>
		''' <param name="NumberTradingDays"></param>
		''' <param name="StockPrice"></param>
		''' <param name="Gain"></param>
		''' <param name="Volatility"></param>
		''' <param name="Probability"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function StockPricePrediction(
																							 ByVal NumberTradingDays As Double,
																							 ByVal StockPrice As Double,
																							 ByVal Gain As Double,
																							 ByVal GainDerivative As Double,
																							 ByVal Volatility As Double,
																							 ByVal Probability As Double) As Double

			Dim ThisResult As Double
			Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
			Dim ThisGain As Double = Gain * (1 + ThisTimeInYear * GainDerivative)
			'Dim ThisPricePrediction As Double = StockOption.StockPricePrediction(NumberTradingDays, StockPrice, Gain)
			'Dim ThisPricePredictionMedian As Double = StockOption.StockPricePredictionMedian(NumberTradingDays, StockPrice, Gain, Volatility)

			'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
			'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
			'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
			Dim ThisMu As Double = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
			Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)

			If ThisSigma > 0 Then
				ThisResult = StockPrice * Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, Probability)
			Else
				ThisResult = StockPrice
			End If
			Return ThisResult
		End Function

		''' <summary>
		''' Simulates the next stock price sample using a Geometric Brownian Motion (GBM) model.
		''' This method assumes that the underlying asset follows a log-normal distribution,
		''' where the logarithmic returns are normally distributed.
		''' 
		''' The formula used is based on the stochastic differential equation:
		''' dS = μS dt + σS dW
		''' which leads to the solution:
		''' S(t) = S(0) * exp[(μ - σ²/2)t + σ√t * Z]
		''' where Z is a standard normal random variable.
		''' 
		''' Simulates the next stock price sample using a Geometric Brownian Motion (GBM) model.
		''' 
		''' This model assumes that the **logarithm of returns** follows a normal (Gaussian) distribution.
		''' Therefore, the resulting stock prices themselves follow a **log-normal** distribution.
		''' 
		''' ⚠ Why not use a normal distribution directly?
		''' - Directly applying a Gaussian distribution to price (e.g., S + N(0, σ)) can result in negative prices, 
		'''   which are not realistic in financial models.
		''' - Financial returns are **multiplicative**, not additive; they compound over time.
		''' 
		''' ✅ Why use log-normal instead:
		''' - Ensures all simulated prices remain positive.
		''' - Captures the correct behavior of compounded returns.
		''' - Mathematically derived from the SDE solution of GBM:
		'''       S(t) = S(0) * exp[(μ - σ²/2)t + σ√t * Z],  where Z ~ N(0,1)
		''' 
		''' In this implementation, we simulate the process by drawing a log-normal sample with parameters:
		'''   - μ' = (Gain - ½·Volatility²) · Time
		'''   - σ' = Volatility · sqrt(Time)
		'''   These match the mean and standard deviation of the normal variable inside the exponent.
		''' 
		''' References:
		''' - https://en.wikipedia.org/wiki/Geometric_Brownian_motion
		''' - Geometric Brownian motion simulation applet:
		'''   http://www.math.uah.edu/stat/apps/GeometricBrownianMotion.html
		''' - General overview of Brownian motion and statistical foundations:
		'''   http://www.math.uah.edu/stat/brown/index.html
		''' - Log-normal distribution:
		'''   https://en.wikipedia.org/wiki/Log-normal_distribution
		''' </summary>
		''' <param name="NumberTradingDays">Forecast horizon in trading days.</param>
		''' <param name="StockPrice">Current stock price (S₀).</param>
		''' <param name="Gain">Annualized expected return (μ).</param>
		''' <param name="GainDerivative">Slope of the expected return (μ') adjustment over time.</param>
		''' <param name="Volatility">Annualized volatility (σ), as a decimal (e.g., 0.2 for 20%).</param>
		''' <returns>A predicted future stock price simulated from the log-normal distribution.</returns>
		Public Shared Function StockPricePredictionSample(
																							 ByVal NumberTradingDays As Double,
																							 ByVal StockPrice As Double,
																							 ByVal Gain As Double,
																							 ByVal GainDerivative As Double,
																							 ByVal Volatility As Double) As Double


			' Convert trading days into fraction of a year
			Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR

			' Adjust expected return to account for slope (derivative of gain)
			Dim ThisGain As Double = Gain * (1 + ThisTimeInYear * GainDerivative)

			' Calculate μ and σ parameters for the log-normal distribution
			Dim ThisMu As Double = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
			Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)
			Dim ThisResult As Double
			' Generate a random sample from the log-normal distribution
			''Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
			''it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
			''These result for the PDF then follow directly from the corresponding results for the lognormal PDF.

			If ThisSigma > 0 Then
				ThisResult = StockPrice * Distributions.LogNormal.Sample(ThisMu, ThisSigma)
			Else
				ThisResult = StockPrice
			End If
			Return ThisResult
		End Function

		Public Shared Function StockPricePrediction(
                                               ByVal NumberTradingDays As Integer,
                                               ByVal StockPrice As Double,
                                               ByVal Gain As Double,
                                               ByVal GainDerivative As Double,
                                               ByVal Volatility As Double,
                                               ByVal Probability As Double) As Double
      Return StockPricePrediction(CDbl(NumberTradingDays), StockPrice, Gain, GainDerivative, Volatility, Probability)
    End Function


    ''' <summary>
    ''' Return the probability that the signal is lower than the StockPriceEnd value after the Number of trading days specified.
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="StockPriceStart"></param>
    ''' <param name="Gain"></param>
    ''' <param name="Volatility"></param>
    ''' <param name="StockPriceEnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StockPricePredictionInverse(
                                                      ByVal NumberTradingDays As Double,
                                                      ByVal StockPriceStart As Double,
                                                      ByVal Gain As Double,
                                                      ByVal Volatility As Double,
                                                      ByVal StockPriceEnd As Double) As Double
      Dim ThisResult As Double
      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      'Dim ThisPricePrediction As Double = StockOption.StockPricePrediction(NumberTradingDays, StockPrice, Gain)
      'Dim ThisPricePredictionMedian As Double = StockOption.StockPricePredictionMedian(NumberTradingDays, StockPrice, Gain, Volatility)

      'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
      'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
      'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
      Dim ThisMu As Double = (Gain - Volatility ^ 2 / 2) * ThisTimeInYear
      Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)

      If ThisSigma > 0 Then
        ThisResult = Distributions.LogNormal.CDF(ThisMu, ThisSigma, StockPriceEnd / StockPriceStart)
      Else
        ThisResult = 0.5
      End If
      Return ThisResult
    End Function

    ''' <summary>
    ''' Return the probability that the signal is lower than the StockPriceEnd value after the Number of trading days specified.
    ''' </summary>
    Public Shared Function StockPricePredictionInverse(
                                                      ByVal NumberTradingDays As Integer,
                                                      ByVal StockPriceStart As Double,
                                                      ByVal Gain As Double,
                                                      ByVal Volatility As Double,
                                                      ByVal StockPriceEnd As Double) As Double

      Return StockPricePredictionInverse(CDbl(NumberTradingDays), StockPriceStart, Gain, Volatility, StockPriceEnd)
    End Function

		''' <summary>
		''' Return the probability that the signal is lower than the StockPriceEnd value after the Number of trading days specified.
		''' </summary>
		''' <param name="NumberTradingDays"></param>
		''' <param name="StockPriceStart"></param>
		''' <param name="Gain"></param>
		''' <param name="GainDerivative"></param>
		''' <param name="Volatility"></param>
		''' <param name="StockPriceEnd"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function StockPricePredictionInverse(
																											ByVal NumberTradingDays As Double,
																											ByVal StockPriceStart As Double,
																											ByVal Gain As Double,
																											ByVal GainDerivative As Double,
																											ByVal Volatility As Double,
																											ByVal StockPriceEnd As Double) As Double
			Dim ThisResult As Double
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
				'Should be the Log(StockPriceEnd / StockPriceStart) but the result is not correct
				ThisResult = Distributions.LogNormal.CDF(ThisMu, ThisSigma, StockPriceEnd / StockPriceStart)
			Else
				ThisResult = 0.5
			End If
			'Dim ThisMu1 As Double = (-ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
			'ThisStockPriceStartAtMedian = StockPriceEnd * Distributions.LogNormal.InvCDF(ThisMu1, ThisSigma, 0.5)
			Return ThisResult
		End Function


		''' <summary>
		''' Return the probability that the signal is lower than the StockPriceEnd value after the Number of trading days specified.
		''' </summary>
		''' <param name="NumberTradingDays"></param>
		''' <param name="StockPriceStart"></param>
		''' <param name="Gain"></param>
		''' <param name="GainDerivative"></param>
		''' <param name="Volatility"></param>
		''' <param name="StockPriceEnd"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function StockPricePredictionInverse(
                                                      ByVal NumberTradingDays As Integer,
                                                      ByVal StockPriceStart As Double,
                                                      ByVal Gain As Double,
                                                      ByVal GainDerivative As Double,
                                                      ByVal Volatility As Double,
                                                      ByVal StockPriceEnd As Double) As Double

      Return StockPricePredictionInverse(
        CDbl(NumberTradingDays),
        StockPriceStart,
        Gain,
        GainDerivative,
        Volatility,
        StockPriceEnd)
    End Function


    ''' <summary>
    ''' Return the 1 Sigma Volatility ratio defined as the volatility relative to the input volatility needed for a 1 sigma probability
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="StockPriceStart"></param>
    ''' <param name="Gain"></param>
    ''' <param name="GainDerivative"></param>
    ''' <param name="Volatility">The input volatility</param>
    ''' <param name="StockPriceEnd"></param>
    ''' <returns>Return the 1 Sigma Volatility ratio</returns>
    ''' <remarks>the function always return a positive ratio between 0 and ~3 </remarks>
    Public Shared Function StockPricePredictionInverseToVolatilityRatio(
                                                      ByVal NumberTradingDays As Double,
                                                      ByVal StockPriceStart As Double,
                                                      ByVal Gain As Double,
                                                      ByVal GainDerivative As Double,
                                                      ByVal Volatility As Double,
                                                      ByVal StockPriceEnd As Double) As Double
      Dim ThisResult As Double
      Dim ThisVolatilityRatio As Double
      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisGain As Double = Gain + (ThisTimeInYear * GainDerivative)
      'Dim ThisPricePrediction As Double = StockOption.StockPricePrediction(NumberTradingDays, StockPrice, Gain)
      'Dim ThisPricePredictionMedian As Double = StockOption.StockPricePredictionMedian(NumberTradingDays, StockPrice, Gain, Volatility)

      'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
      'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
      'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
      Dim ThisMu As Double = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
      Dim ThisSigma As Double = Volatility * Math.Sqrt(ThisTimeInYear)

      'this is an exact mathematic solution
      If ThisSigma > 0 Then
        ThisResult = Distributions.LogNormal.CDF(ThisMu, ThisSigma, StockPriceEnd / StockPriceStart)
        'limit the calculation range
        If ThisResult > 0.999 Then
          ThisResult = 0.999
        End If
				If ThisResult < 0.001 Then
					ThisResult = 0.001
				End If
				Dim ThisProbStart = Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, ThisResult)
				Dim ThisProbStop = Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, 1 - ThisResult)
				Dim ThisProbStartAtSigma1 = Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, Measure.GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1)
				Dim ThisProbStopAtSigma1 = Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, Measure.GAUSSIAN_PROBABILITY_MEAN_MINUS_SIGMA1)

				ThisVolatilityRatio = (ThisProbStart - ThisProbStop) / (ThisProbStartAtSigma1 - ThisProbStopAtSigma1)
			Else
        ThisVolatilityRatio = 1.0
      End If
      If ThisVolatilityRatio < 0.0 Then
        ThisVolatilityRatio = -ThisVolatilityRatio
      End If
      Return ThisVolatilityRatio
    End Function

    Public Shared Function BlackScholesOptionProbabilityOfInTheMoney(
        ByVal OptionType As enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal VolatilityPerYear As Double) As Double

      Return Measure.BlackScholesOptionProbabilityOfInTheMoney(
        OptionType,
        StockPrice,
        OptionStrikePrice,
        TimeToExpirationInYear,
        RiskFreeRate,
        DividendRate,
        VolatilityPerYear)


    End Function
    ''' <summary>
    '''   Return the stock price needed for a succesful trade given the option price paid. Note that the method
    '''   use an iterative technique that can be time expensive and it should be use with caution. The TimeLimitInYear 
    '''   take into account the decay of the option over the period of interest and should be less 
    '''   than the expiration time of the option
    ''' </summary>
    ''' <param name="OptionPrice">The price pay for the option</param>
    ''' <param name="TimeLimitInYear">
    '''   Use to take into account the decay of the option over the period of interest and should be less 
    '''   than the expiration time of the option
    ''' </param>
    ''' <returns>
    '''   The Stock Price required for an In The Money (ITM) succesful option trade
    ''' </returns>
    Public Function PriceOfOptionToPriceOfStock(ByVal OptionPrice As Double, ByVal TimeLimitInYear As Double) As Double
      'add some protection 
      Dim ThisCount As Integer = 0
      If OptionPrice <= 0 Then Return 0.0
      If MyOptionPriceDelta = 0 Then Return 0.0
      If Double.IsNaN(MyValueOptionStandard) Then Return 0.0

      'initialize the first iteration step
      Dim ThisStockOption = Me.Copy
      Dim ThisTimeLimitInDays As Double = 365 * TimeLimitInYear

      'reduce the time to expiration to take into account the decay over that period
      ThisStockOption.DateExpiration = ThisStockOption.DateExpiration.AddDays(-ThisTimeLimitInDays)
      'for convergence check and make sure that DateStart is less than the DateExpiration by at least one day
      If ThisStockOption.DateExpiration <= ThisStockOption.DateStart Then
        'add 1 day to teh date expiration
        ThisStockOption.DateExpiration = ThisStockOption.DateStart.AddDays(1)
      End If
      'the stock price is use as a seed value. The final result is not dependant on that value
      Dim ThisStockPrice As Double = Me.Price

      Dim ThisValueOptionStandard = Me.AsIStockOptionPrice.ValueStandard
      Dim ThisOptionPriceDelta = Me.AsIStockOptionPrice.ValueDelta
      Dim ThisOptionPriceDeltaLarge = OptionPrice - ThisValueOptionStandard
      Dim ThisStockPriceDeltaEstimate As Double = ThisOptionPriceDeltaLarge / ThisOptionPriceDelta
      Do
        ThisStockPrice = ThisStockPrice + ThisStockPriceDeltaEstimate
        ThisStockOption.Refresh(ThisStockPrice)
        ThisValueOptionStandard = ThisStockOption.AsIStockOptionPrice.ValueStandard
        ThisOptionPriceDelta = ThisStockOption.AsIStockOptionPrice.ValueDelta
        ThisOptionPriceDeltaLarge = OptionPrice - ThisValueOptionStandard
        ThisStockPriceDeltaEstimate = ThisOptionPriceDeltaLarge / ThisOptionPriceDelta
        ThisCount = ThisCount + 1
        If ThisCount > 5 Then
          Exit Do
        End If
      Loop Until Math.Abs(ThisStockPriceDeltaEstimate) < 0.1
      ThisStockPrice = ThisStockPrice + ThisStockPriceDeltaEstimate
      Return ThisStockPrice
    End Function




    Public Shared Function StockPricePredictionMedian(
                                                     ByVal NumberTradingDays As Double,
                                                     ByVal StockPrice As Double,
                                                     ByVal Gain As Double,
                                                     ByVal Volatility As Double) As Double

      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisStockPriceMedian As Double = StockPrice * Math.Exp((Gain - Volatility ^ 2 / 2) * ThisTimeInYear)
      Return ThisStockPriceMedian
    End Function

    Public Shared Function StockPricePredictionMedian(
                                                     ByVal NumberTradingDays As Integer,
                                                     ByVal StockPrice As Double,
                                                     ByVal Gain As Double,
                                                     ByVal Volatility As Double) As Double
      Return StockPricePredictionMedian(CDbl(NumberTradingDays), StockPrice, Gain, Volatility)
    End Function

    Public Shared Function StockPricePredictionMedian(
                                                     ByVal NumberTradingDays As Double,
                                                     ByVal StockPrice As Double,
                                                     ByVal Gain As Double,
                                                     ByVal GainDerivative As Double,
                                                     ByVal Volatility As Double) As Double

      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisGain As Double = Gain + (ThisTimeInYear * GainDerivative)
      Dim ThisStockPriceMedian As Double = StockPrice * Math.Exp((ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear)
      Return ThisStockPriceMedian
    End Function

    Public Shared Function StockPricePredictionMedian(
                                                     ByVal NumberTradingDays As Integer,
                                                     ByVal StockPrice As Double,
                                                     ByVal Gain As Double,
                                                     ByVal GainDerivative As Double,
                                                     ByVal Volatility As Double) As Double
      Return StockPricePredictionMedian(CDbl(NumberTradingDays), StockPrice, Gain, GainDerivative, Volatility)
    End Function

    ' ''' <summary>
    ' ''' This function estimate the partial derivative Δprobability/Δsigma for a small sigma variation at a fixed probabilty threshold.
    ' ''' </summary>
    ' ''' <param name="NumberTradingDays"></param>
    ' ''' <param name="Volatility">the volatility level</param>
    ' ''' <param name="Probability">the probability point</param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Shared Function StockPricePredictionPartialDerivateOfProbabilityToVolatility(
    '                                            ByVal NumberTradingDays As Double,
    '                                            ByVal Volatility As Double,
    '                                            ByVal Probability As Double) As Double

    '  Const STEP_RATIO As Double = 0.01

    '  Dim ThisResult As Double
    '  Dim ThisResult1 As Double
    '  Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
    '  Dim ThisGain As Double = 0.0 'the gain has no effect on the calculation
    '  Dim ThisMu As Double
    '  Dim ThisSigma As Double
    '  Dim ThisProbability1 As Double
    '  Dim ThisVolatility1 As Double


    '  'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
    '  'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
    '  'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
    '  If Volatility < STEP_RATIO Then
    '    Volatility = STEP_RATIO
    '  End If
    '  ThisMu = (ThisGain - Volatility ^ 2 / 2) * ThisTimeInYear
    '  ThisSigma = Volatility * Math.Sqrt(ThisTimeInYear)
    '  ThisVolatility1 = Volatility + STEP_RATIO

    '  Dim ThisMu1 As Double = (ThisGain - ThisVolatility1 ^ 2 / 2) * ThisTimeInYear
    '  Dim ThisSigma1 As Double = ThisVolatility1 * Math.Sqrt(ThisTimeInYear)
    '  Dim ThisGradient As Double

    '  'this measurement give the partial derivative or gradient of the probability relative to a fixed sigma.
    '  If ThisSigma > 0 Then
    '    ThisResult = Distributions.LogNormal.InvCDF(ThisMu, ThisSigma, Probability)
    '    ThisResult1 = Distributions.LogNormal.InvCDF(ThisMu1, ThisSigma1, Probability)
    '    ThisProbability1 = Distributions.LogNormal.CDF(ThisMu, ThisSigma, ThisResult1)
    '    ThisGradient = (ThisProbability1 - Probability) / (STEP_RATIO)
    '  Else
    '    ThisResult = 1
    '  End If
    '  Return ThisGradient
    'End Function

    ''' <summary>
    ''' This function estimate the partial derivative Δprobability/Δsigma for a small sigma variation at a fixed probabilty threshold.
    ''' </summary>
    ''' <param name="NumberTradingDays"></param>
    ''' <param name="Volatility">the volatility level</param>
    ''' <param name="Probability">the probability point</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function StockPricePredictionPartialDerivateOfProbabilityToVolatility(
                                                ByVal NumberTradingDays As Double,
                                                ByVal Volatility As Double,
                                                ByVal Probability As Double) As Double

      Const STEP_RATIO As Double = 0.001

      Dim ThisResult As Double
      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisGain As Double = 0.0 'the gain has no effect on the calculation
      Dim ThisMu(0 To 2) As Double
      Dim ThisSigma(0 To 2) As Double
      Dim ThisProbability(0 To 2) As Double
      Dim ThisVolatility(0 To 2) As Double
      Dim I As Integer

      Dim ThisCubicSpline As Interpolation.CubicSpline

      'Since the variable Ut=(μ−σ2/2)t+σZt has the normal distribution with mean (μ−σ2/2)t and standard deviation σ√t, 
      'it follows that Xt=exp(Ut) has the lognormal distribution with these parameters. 
      'These result for the PDF then follow directly from the corresponding results for the lognormal PDF.
      If Volatility < 0 Then Volatility = 0
      If Volatility < STEP_RATIO Then
        Volatility = STEP_RATIO
      End If
      ThisVolatility(0) = Volatility - STEP_RATIO
      ThisVolatility(1) = Volatility
      ThisVolatility(2) = Volatility + STEP_RATIO
      ThisProbability(1) = Probability
      For I = 0 To 2
        ThisMu(I) = (ThisGain - ThisVolatility(I) ^ 2 / 2) * ThisTimeInYear
        ThisSigma(I) = ThisVolatility(I) * Math.Sqrt(ThisTimeInYear)
      Next
      ThisResult = Distributions.LogNormal.InvCDF(ThisMu(1), ThisSigma(1), ThisProbability(1))
      'for a fix ThisResult position calculate the partial derivate of delta probability over delta sigma
      ThisProbability(0) = Distributions.LogNormal.CDF(ThisMu(0), ThisSigma(0), ThisResult)
      ThisProbability(2) = Distributions.LogNormal.CDF(ThisMu(2), ThisSigma(2), ThisResult)
      'get the cubic spline of the 3 points
      ThisCubicSpline = Interpolation.CubicSpline.InterpolateNaturalSorted(ThisSigma, ThisProbability)
      ThisResult = ThisCubicSpline.Differentiate(ThisSigma(1))
      Return ThisResult
    End Function

    Public Shared Function StockPricePredictionMeanToMedian(
                                                     ByVal NumberTradingDays As Double,
                                                     ByVal Volatility As Double) As Double

      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisMeanToMedianRatio As Double = Math.Exp((Volatility ^ 2 / 2) * ThisTimeInYear)
      Return ThisMeanToMedianRatio
    End Function

    Public Shared Function StockPricePredictionMeanToMedian(
                                                     ByVal NumberTradingDays As Integer,
                                                     ByVal Volatility As Double) As Double

      Return StockPricePredictionMeanToMedian(CDbl(NumberTradingDays), Volatility)
    End Function

    Public Shared Function StockPriceSigmaPrediction(
                                                    ByVal NumberTradingDays As Double,
                                                    ByVal StockPrice As Double,
                                                    ByVal Gain As Double,
                                                    ByVal Volatility As Double) As Double

      Dim ThisTimeInYear As Double = NumberTradingDays / YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_YEAR
      Dim ThisPricePrediction As Double = StockOption.StockPricePrediction(NumberTradingDays, StockPrice, Gain)

      'Dim ThisBandLow As Double = StockOption.StockPriceProbability(NumberTradingDays, StockPrice, Gain, Volatility, 0.5)
      Return Math.Sqrt(ThisPricePrediction ^ 2 * (Math.Exp(ThisTimeInYear * Volatility ^ 2) - 1))
    End Function

    Public Shared Function StockPriceSigmaPrediction(
                                                    ByVal NumberTradingDays As Integer,
                                                    ByVal StockPrice As Double,
                                                    ByVal Gain As Double,
                                                    ByVal Volatility As Double) As Double
      Return StockPriceSigmaPrediction(CDbl(NumberTradingDays), StockPrice, Gain, Volatility)
    End Function

    Public Function SymbolOption(Optional OptionSymbolType As IStockOption.enuOptionSymbolType = IStockOption.enuOptionSymbolType.Yahoo) As String Implements IStockOption.SymbolOption
      Return StockOption.SymbolOption(IStockOption.enuOptionSymbolType.Yahoo, Me.OptionType, Me.Symbol, Me.DateExpiration, Me.StrikePrice)
    End Function

    Public Property ID As Integer Implements IStockOption.ID

    Private Property IStockOption_VolatilityStandardImplied As Double Implements IStockOption.VolatilityStandardImplied
      Get
        Return Me.VolatilityStandardImplied
      End Get
      Set(value As Double)
        Me.VolatilityStandardImplied = value
      End Set
    End Property

    ''' <summary>
    ''' The time decay of an option per day. Theta is expressed here as 
    ''' a negative number and can be thought of the amount by which the option's value declines per day.
    ''' </summary>
    ''' <returns>The time decay of an option per day</returns>
    Public Function ValueTheta() As Double Implements IStockOptionPrice.ValueTheta
      Return MyOptionPriceTheta
    End Function

    ''' <summary>
    ''' The option price variation with a 1% volatility variation
    ''' </summary>
    ''' <returns></returns>
    Public Function BlackScholeValueVega() As Double
      Return MyOptionPriceVega
    End Function

    Public Shared Function BlackScholesOptionImpliedVolatility(
        ByVal OptionType As Measure.enuOptionType,
        ByVal StockPrice As Double,
        ByVal OptionStrikePrice As Double,
        ByVal OptionMarketPrice As Double,
        ByVal TimeToExpirationInYear As Double,
        ByVal RiskFreeRate As Double,
        ByVal DividendRate As Double,
        ByVal ErrorEpsilon As Double) As Double

      Return Measure.BlackScholesOptionImpliedVolatility(OptionType,
        StockPrice,
        OptionStrikePrice,
        OptionMarketPrice,
        TimeToExpirationInYear,
        RiskFreeRate,
        DividendRate,
        ErrorEpsilon)
    End Function

    Public Shared Function SymbolOption(
      ByVal OptionSymbolType As IStockOption.enuOptionSymbolType,
      ByVal OptionType As Measure.enuOptionType,
      ByVal Symbol As String,
      ByVal DateExpiration As Date,
      ByVal StrikePrice As Double) As String


      Return SymbolOption(
        OptionSymbolType,
        OptionType,
        Symbol,
        DateExpiration,
        StrikePrice,
        SignatureOfCall:="C",
        SignatureOfPut:="P")
    End Function

    Public Shared Function SymbolOption(
      ByVal OptionSymbolType As IStockOption.enuOptionSymbolType,
      ByVal OptionType As Measure.enuOptionType,
      ByVal Symbol As String,
      ByVal DateExpiration As Date,
      ByVal StrikePrice As Double,
      ByVal SignatureOfCall As String,
      ByVal SignatureOfPut As String) As String

      Dim ThisSymbolOption As String
      Dim ThisStrikePriceValue As String

      If Strings.Left(Symbol, 1) = "~" Then
        'this is a user indice
        'remove the special character identification from the symbol
        Symbol = Replace(Symbol, "~", "")
        Symbol = Replace(Symbol, ".I", "")
      End If
      ThisStrikePriceValue = Strings.Right(String.Format("{0}{1}", "00000000", Int(1000 * StrikePrice).ToString), 8)
      If OptionType = Measure.enuOptionType._Call Then
        ThisSymbolOption = String.Format("{0}{1:yyMMdd}{2}{3}", Symbol, DateExpiration.Date, SignatureOfCall, ThisStrikePriceValue)
      Else
        ThisSymbolOption = String.Format("{0}{1:yyMMdd}{2}{3}", Symbol, DateExpiration.Date, SignatureOfPut, ThisStrikePriceValue)
      End If
      Return ThisSymbolOption
    End Function

    Protected Overrides Sub Finalize()
      MyBase.Finalize()
    End Sub
#End Region
  End Class
End Namespace