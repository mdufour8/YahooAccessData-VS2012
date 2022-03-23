Imports YahooAccessData.MathPlus.Measure

Namespace OptionValuation
  Public Interface IStockOption
    Inherits IStockOptionPrice
    Inherits IStockPrice
    Inherits IStockPricePrediction

    Enum enuOptionSymbolType
      Yahoo
    End Enum
    Enum enuDividendPaymentPeriodType
      Monthly
      Quaterly
      BiAnnual
      Yearly
      AutoDetect
    End Enum
    Enum enuTimeToExpirationScale
      Year
      Day
      TradingDay
      Hour
    End Enum
    Enum enuVolatilityStandardYearlyType
      Daily10
      Daily15
      Monthly
      BiMonthly
      Quaterly
      BiAnnual
      Yearly
      YearlyMonthly
      YearlyDaily10
      YearlyDaily15
      ToExpiration
    End Enum
    Enum enuVolatilityYearlyType
      Monthly
      Quaterly
      BiAnnual
      Yearly
      YearlyMonthly
      YearlyDaily10
      YearlyDaily15
    End Enum
    Enum enuVolatilityMeasurementMethod
      Standard
      StandardExp
      YangZhang
      YangZhangExp
      StandardYangZhangMerged
      StandardYangZhangMergedExp
      YangZhangExpRegulated
    End Enum

    Enum enuOptionStyle
      American
      Europeen
    End Enum

    Property ID As Integer
    Property Symbol As String
    Property Price As Double
    Property StrikePrice As Double
    Property PriceBuy As Double
    Property PriceCLose As Double
    Property Volatility As Double
    Property VolatilityStandard As Double
    Property RateBase As Double
    Property RateDividend As Double
    Property DateStart As Date
    Property DateExpiration As Date
    Property DateBuy As Date
    Property DateClose As Date
    Property Gain As Double
    Property GainSigmaError As Double
    Property GainSigmaErrorPeriod As Integer
    Property IsGainSigmaErrorEnabled As Boolean
    Property OptionType As Measure.enuOptionType
    Property OptionStyle As enuOptionStyle
    ReadOnly Property DividendPaymentPeriod As Integer
    Property DividendPaymentPeriodDetected As Integer
    Property DividendPaymentPeriodType As enuDividendPaymentPeriodType
    Property ExDividendDateEstimated As Date
    Property ExDividendDateDeclared As Date
    Property IsExDividendDateDeclaredEnabled As Boolean
    Property IsExDividendDateEnabled As Boolean
    Property VolatilityStandardYearlyType As enuVolatilityStandardYearlyType
    Property VolatilityMeasurementMethod As enuVolatilityMeasurementMethod
    ReadOnly Property RateDividendToPayment As Double
    ReadOnly Property DividendPaymentValue As Double
    ReadOnly Property DividendPaymentNumber As Integer
    Property VolatilityStandardImplied As Double
    Property IsVolatilityStandardImpliedEnabled As Boolean
    Function Refresh() As Double
    Function Refresh(ByVal DateToday As Date) As Double
    Function Refresh(ByVal DateToday As Date, ByVal Price As Double) As Double
    Function AsIStockOption() As IStockOption
    Function TimeToExpiration(Optional ByVal TimeToExpirationScale As enuTimeToExpirationScale = enuTimeToExpirationScale.Year) As Double
    Function TimeToExDividend(Optional ByVal TimeToExpirationScale As enuTimeToExpirationScale = enuTimeToExpirationScale.Day) As Double
    Function VolatilityTotal() As Double
    Function VolatilityFilterRate() As Integer
    Function SymbolOption(Optional OptionSymbolType As IStockOption.enuOptionSymbolType = enuOptionSymbolType.Yahoo) As String
  End Interface

  Public Interface IStockOptionPriceGain
    Function AsIStockOptionPriceGain() As IStockOptionPriceGain
    Property Gain As Double
    Property GainDerivative As Double
    Property PriceStockMedian As Double
    Property VolatilityStandardToOpen As Double

    Property VolatilityStandardToClose As Double

    Property VolatilityStandardFromOpenToClose As Double

    Property VolatilityRegulatedToOpen As Double

    Property VolatilityRegulatedToClose As Double

    Property VolatilityRegulatedFromOpenToClose As Double
  End Interface

  Public Interface IStockPrice
    Function AsIStockPrice() As IStockPrice
    ReadOnly Property Value As Double
    'ReadOnly Property ValueCDF(ByVal X As Double) As Double
    ReadOnly Property ValueOfSigma As Double
    ReadOnly Property ValueOfDeltaPrice As Double
  End Interface

  Public Interface IStockPricePrediction
    Function AsIStockPricePrediction() As IStockPricePrediction
    Function Value(ByVal NumberTradingDays As Integer) As Double
    Function ValueOfSigma(ByVal NumberTradingDays As Integer) As Double
  End Interface

  Public Interface IStockOptionPrice
    Function AsIStockOptionPrice() As IStockOptionPrice
    ReadOnly Property ValueFromGainPrediction As Double
    ReadOnly Property ValueStandard As Double
    ReadOnly Property ValueAtSigma As Double
    ReadOnly Property ValueDelta As Double
    Function ValueTheta() As Double
  End Interface
End Namespace