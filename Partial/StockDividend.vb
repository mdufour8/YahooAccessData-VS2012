Public Class StockDividend
  Implements IStockDividend

  'the most frequent in north america
  Const PAYMENT_FREQUENCY_DEFAULT_PER_YEAR As Double = 4.0


  Private MyDateReference As Date
  Private MyPriceValue As Single
  Private MyDividendShare As Single
  Private MyDividendShareToExDividendFuture As Single
  Private MyExDividendDateDeclared As Date
  Private MyExDividendDateEstimateFuture As Date
  Private MyExDividendDateDeclaredLast As Date
  Private MyPaymentFrequencyEstimatePerYear As Double
  Private MyPaymentFrequencyEstimatePerYearDefault As Double
  Private MyPaymentPeriodEstimateInDays As Double

  Private MyNumberDaysToExDividendFuture As Integer

  Public Sub New(ByRef PriceVol As PriceVol)
    Me.New(PriceVol, PAYMENT_FREQUENCY_DEFAULT_PER_YEAR)
  End Sub

  Public Sub New(ByRef PriceVol As PriceVol, ByVal PaymentFrequencyEstimatePerYear As Double)
    MyDateReference = PriceVol.DateLastTrade
    MyPriceValue = PriceVol.Last
    MyDividendShare = PriceVol.DividendShare
    'extimate the frequency of the payment
    MyPaymentFrequencyEstimatePerYearDefault = PaymentFrequencyEstimatePerYear
    MyPaymentFrequencyEstimatePerYear = PaymentFrequencyEstimatePerYear
    MyPaymentPeriodEstimateInDays = 365 / MyPaymentFrequencyEstimatePerYear
    MyExDividendDateDeclaredLast = PriceVol.ExDividendDate
    'set by default MyExDividendDateDeclared to the last declared dividend
    'it with be updated to a future date when it become known by the user
    MyExDividendDateDeclared = MyExDividendDateDeclaredLast
    'calculate the estimate of number of day left until the next payment
    MyExDividendDateEstimateFuture = MyExDividendDateDeclaredLast.AddDays(MyPaymentPeriodEstimateInDays)
    MyNumberDaysToExDividendFuture = MyExDividendDateEstimateFuture.Subtract(MyDateReference).Days
  End Sub

  Public Function AsIStockDividend() As IStockDividend Implements IStockDividend.AsIStockDividend
    Return Me
  End Function

  ''' <summary>
  ''' The current date
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property DateReference As Date Implements IStockDividend.DateReference
    Get
      Return MyDateReference
    End Get
  End Property

  ''' <summary>
  ''' The yearly divident/share payment
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property DividendShare As Single Implements IStockDividend.DividendShare
    Get
      Return MyDividendShare
    End Get
  End Property
  ''' <summary>
  ''' The yearly divident rate payment in Percent
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property DividendYield As Single Implements IStockDividend.DividendYield
    Get
      If MyPriceValue > 0 Then
        Return 100 * MyDividendShare / MyPriceValue
      Else
        Return 0.0
      End If
    End Get
  End Property

  ''' <summary>
  ''' The payment/share at the next ex dividend date including the effect of the payment period
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property DividendShareToExDividend As Single Implements IStockDividend.DividendShareToExDividend
    Get
      Return CSng(MyDividendShare / MyPaymentFrequencyEstimatePerYear)
    End Get
  End Property
  ''' <summary>
  ''' The divident rate payment per year in Percent calculated based on the future ex dividend date payment
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property DividendYieldToExDividend As Single Implements IStockDividend.DividendYieldToExDividend
    Get
      Return CSng(100 * (MyNumberDaysToExDividendFuture / 365) * ((MyDividendShare / MyPaymentFrequencyEstimatePerYear) / MyPriceValue))
    End Get
  End Property

  ''' <summary>
  ''' the future date declared of the ex dividend
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ExDividendDateDeclared As Date Implements IStockDividend.ExDividendDateDeclared
    Get
      Return MyExDividendDateDeclared
    End Get
  End Property
  ''' <summary>
  ''' the future date estimate of the ex dividend
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ExDividendDateEstimate As Date Implements IStockDividend.ExDividendDateEstimate
    Get
      Return MyExDividendDateEstimateFuture
    End Get
  End Property
  ''' <summary>
  ''' The future date of the ex dividend estimated or declared
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ExDividendDateFuture As Date Implements IStockDividend.ExDividendDateFuture
    Get
      If Me.IsExDividendDateDeclaredValid Then
        Return MyExDividendDateDeclared
      Else
        Return MyExDividendDateEstimateFuture
      End If
    End Get
  End Property

  Public Function IsExDividendDateDeclaredValid() As Boolean Implements IStockDividend.IsExDividendDateDeclaredValid
    If Me.ExDividendDateDeclared > MyExDividendDateDeclaredLast Then
      Return True
    Else
      Return False
    End If
  End Function

  Public ReadOnly Property PriceValue As Single Implements IStockDividend.PriceValue
    Get
      Return MyPriceValue
    End Get
  End Property

  ''' <summary>
  ''' The number of days to the ex dividend payment
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property NumberDaysToExDividend As Integer Implements IStockDividend.NumberDaysToExDividend
    Get
      Return MyNumberDaysToExDividendFuture
    End Get
  End Property

  ''' <summary>
  ''' The period estimate between dividend payment in year
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property PaymentPeriodEstimate As Double Implements IStockDividend.PaymentPeriodEstimate
    Get
      Return 1 / MyPaymentFrequencyEstimatePerYear
    End Get
  End Property

  ''' <summary>
  ''' The frequency estimate of the dividend payment per Year
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property PaymentFrequencyEstimate As Double Implements IStockDividend.PaymentFrequencyEstimate
    Get
      Return MyPaymentFrequencyEstimatePerYear
    End Get
  End Property

  Public Sub RefreshDividendDateDeclared(ValueOfExDividendDateDeclared As Date) Implements IStockDividend.RefreshDividendDateDeclared
    If ValueOfExDividendDateDeclared < MyExDividendDateDeclaredLast Then
      ValueOfExDividendDateDeclared = MyExDividendDateDeclaredLast
    End If
    If ExDividendDateDeclared > MyExDividendDateDeclaredLast Then
      MyExDividendDateDeclared = ExDividendDateDeclared
      MyPaymentFrequencyEstimatePerYear = 365 / MyExDividendDateDeclared.Subtract(MyExDividendDateDeclaredLast).Days
    Else
      MyPaymentFrequencyEstimatePerYear = MyPaymentFrequencyEstimatePerYearDefault
    End If
    MyPaymentPeriodEstimateInDays = 365 / MyPaymentFrequencyEstimatePerYear
    'set by default MyExDividendDateDeclared to the last declared dividend
    'it with be updated to a future date when it become known by the user
    MyExDividendDateDeclared = MyExDividendDateDeclaredLast
    'calculate the estimate of number of day left until the next payment
    MyExDividendDateEstimateFuture = MyExDividendDateDeclaredLast.AddDays(MyPaymentPeriodEstimateInDays)
    MyNumberDaysToExDividendFuture = MyExDividendDateEstimateFuture.Subtract(MyDateReference).Days
  End Sub

  Public Sub RefreshDividendDateEstimate(ValueOfExDividendDateEstimate As Date) Implements IStockDividend.RefreshDividendDateEstimate
    'If ValueOfExDividendDateEstimate < MyExDividendDateDeclaredLast Then
    '  ValueOfExDividendDateEstimate = MyExDividendDateDeclaredLast
    'End If
    'If ExDividendDateDeclared > MyExDividendDateDeclaredLast Then
    '  MyExDividendDateDeclared = ExDividendDateDeclared
    '  MyPaymentFrequencyEstimatePerYear = 365 / MyExDividendDateDeclared.Subtract(MyExDividendDateDeclaredLast).Days
    'Else
    '  MyPaymentFrequencyEstimatePerYear = MyPaymentFrequencyEstimatePerYearDefault
    'End If
    'MyPaymentPeriodEstimateInDays = 365 / MyPaymentFrequencyEstimatePerYear
    ''set by default MyExDividendDateDeclared to the last declared dividend
    ''it with be updated to a future date when it become known by the user
    'MyExDividendDateDeclared = MyExDividendDateDeclaredLast
    ''calculate the estimate of number of day left until the next payment
    'MyExDividendDateEstimateFuture = MyExDividendDateDeclaredLast.AddDays(MyPaymentPeriodEstimateInDays)
    'MyNumberDaysToExDividendFuture = MyExDividendDateEstimateFuture.Subtract(MyDateReference).Days

  End Sub

  Public ReadOnly Property ExDividendDateDeclaredLast As Date Implements IStockDividend.ExDividendDateDeclaredLast
    Get
      Return MyExDividendDateDeclaredLast
    End Get
  End Property
End Class
