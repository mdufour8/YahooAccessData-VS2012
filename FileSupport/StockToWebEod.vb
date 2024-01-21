Public Class StockToWebEod
  Private MyWebStockCode As String
  Private MyWebExchangeCode As String
  Private MyWebExchangeName As String
  Private MyWebExchangeCountry As String
  Sub New(ByVal StockData As Stock)
    Dim ThisResultOfSplit() As String = StockData.Exchange.Split(":"c)

    MyWebExchangeCountry = ThisResultOfSplit(0)
    MyWebExchangeCode = ThisResultOfSplit(1)
    MyWebExchangeName = ThisResultOfSplit(2)

    'for the symbol code
    '
    MyWebStockCode = StockData.Symbol
    If MyWebExchangeCode <> "US" Then
      'need to remove the exchange code added at the end of the symbol
      'should work in all case but may be subject to problem if the exchange code 
      Dim ThisLastIndexOfDot As Integer = MyWebStockCode.LastIndexOf("."c)

      If ThisLastIndexOfDot > 0 Then
        'replace from that position
        MyWebStockCode = MyWebStockCode.Substring(0, ThisLastIndexOfDot + 1)
      End If
    End If
  End Sub

  Public ReadOnly Property StockCode As String
    Get
      Return MyWebStockCode
    End Get
  End Property

  Public ReadOnly Property ExchangeCode As String
    Get
      Return MyWebExchangeCode
    End Get
  End Property

  Public ReadOnly Property ExchangeName As String
    Get
      Return MyWebExchangeName
    End Get
  End Property

  Public ReadOnly Property ExchangeCountry As String
    Get
      Return MyWebExchangeName
    End Get
  End Property
End Class