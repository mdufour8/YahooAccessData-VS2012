Public Interface IStockDividendSinglePayout
  Property DateOfExDividend As Date
  Property DateReference As Date
  Property PricePayoutValue As Single
  Function AsIStockDividendSinglePayout() As IStockDividendSinglePayout
End Interface
