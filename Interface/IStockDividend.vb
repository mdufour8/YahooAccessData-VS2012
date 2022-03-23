Public Interface IStockDividend
  ReadOnly Property ExDividendDateEstimate As Date
  ReadOnly Property ExDividendDateDeclared As Date
  ReadOnly Property ExDividendDateDeclaredLast As Date
  ReadOnly Property PaymentFrequencyEstimate As Double
  ReadOnly Property PaymentPeriodEstimate As Double
  ReadOnly Property DividendShare As Single
  ReadOnly Property DividendYield As Single
  ReadOnly Property DividendShareToExDividend As Single
  ReadOnly Property DividendYieldToExDividend As Single
  ReadOnly Property NumberDaysToExDividend As Integer
  ReadOnly Property DateReference As Date
  ReadOnly Property PriceValue As Single
  ReadOnly Property ExDividendDateFuture As Date

  Sub RefreshDividendDateEstimate(ByVal ExDividendDateEstimate As Date)
  Sub RefreshDividendDateDeclared(ByVal ExDividendDateDeclared As Date)

  Function IsExDividendDateDeclaredValid() As Boolean
  Function AsIStockDividend() As IStockDividend
End Interface
