#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Public Class StockDividendSinglePayout
  Implements IStockDividendSinglePayout

  Private MyDateOfExDividend As Date
  Private MyDateReference As Date
  Private MyPricePayoutValue As Single

  Public Sub New(
    ByVal DateReference As Date,
    ByVal DateOfExDividend As Date,
    ByVal PricePayoutValue As Single)

    MyDateOfExDividend = DateOfExDividend
    MyDateReference = DateReference
    MyPricePayoutValue = PricePayoutValue
  End Sub

  Public Sub New()
    Me.New(Now.Date, Now.Date, 0.0)
  End Sub

  Public Property DateOfExDividend As Date Implements IStockDividendSinglePayout.DateOfExDividend
    Get
      Return MyDateOfExDividend
    End Get
    Set(value As Date)
      MyDateOfExDividend = value.Date
    End Set
  End Property

  Public Property DateReference As Date Implements IStockDividendSinglePayout.DateReference
    Get
      Return MyDateReference
    End Get
    Set(value As Date)
      MyDateReference = value.Date
    End Set
  End Property

  Public Property PricePayoutValue As Single Implements IStockDividendSinglePayout.PricePayoutValue
    Get
      Return MyPricePayoutValue
    End Get
    Set(value As Single)
      MyPricePayoutValue = value
    End Set
  End Property

  Public Function AsIStockDividendSinglePayout() As IStockDividendSinglePayout Implements IStockDividendSinglePayout.AsIStockDividendSinglePayout
    Return Me
  End Function
End Class


