Public Class StockBasicInfo
	Implements IStockBasicInfo

  Private MySymbol As String
  Private MyName As String
  Private IsOptionLocal As Boolean
  Private MyDateStart As Date
  Private MyDateStop As Date
  Private IsSymbolErrorLocal As Boolean
  Private MyExchange As String
  Private MyErrorDescription As String
  Private MySectorName As String
  Private MyIndustryName As String

  Public Sub New(ByVal Stock As Stock)
    'copy the basic information
    With Stock
      MySymbol = .Symbol
      MyName = .Name
      IsOptionLocal = .IsOption
      MyDateStart = .DateStart
      MyDateStop = .DateStop
      IsSymbolErrorLocal = .IsSymbolError
      MyExchange = .Exchange
      MyErrorDescription = .ErrorDescription
      MySectorName = .SectorName
      MyIndustryName = .IndustryName
    End With
  End Sub

  Private ReadOnly Property AsStockBasicInfo As IStockBasicInfo Implements IStockBasicInfo.AsStockBasicInfo
    Get
      Return Me
    End Get
  End Property

  Public ReadOnly Property Symbol As String Implements IStockBasicInfo.Symbol
    Get
      Return MySymbol
    End Get
  End Property

  Public ReadOnly Property Name As String Implements IStockBasicInfo.Name
    Get
      Return MyName
    End Get
  End Property

  Public ReadOnly Property IsOption As Boolean Implements IStockBasicInfo.IsOption
    Get
      Return IsOptionLocal
    End Get
  End Property

  Public ReadOnly Property DateStart As Date Implements IStockBasicInfo.DateStart
    Get
      Return MyDateStart
    End Get
  End Property

  Public ReadOnly Property DateStop As Date Implements IStockBasicInfo.DateStop
    Get
      Return MyDateStop
    End Get
  End Property

  Public ReadOnly Property IsSymbolError As Boolean Implements IStockBasicInfo.IsSymbolError
    Get
      Return IsSymbolErrorLocal
    End Get
  End Property

  Public ReadOnly Property Exchange As String Implements IStockBasicInfo.Exchange
    Get
      Return MyExchange
    End Get
  End Property

  Public ReadOnly Property ErrorDescription As String Implements IStockBasicInfo.ErrorDescription
    Get
      Return MyErrorDescription
    End Get
  End Property

  Public ReadOnly Property SectorName As String Implements IStockBasicInfo.SectorName
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Public ReadOnly Property IndustryName As String Implements IStockBasicInfo.IndustryName
    Get
      Throw New NotImplementedException()
    End Get
  End Property
End Class

#Region "IStockBasicInfo"
Public Interface IStockBasicInfo
  ReadOnly Property AsStockBasicInfo As IStockBasicInfo
  ReadOnly Property Symbol As String
  ReadOnly Property Name As String
  ReadOnly Property IsOption As Boolean
  ReadOnly Property DateStart As Date
  ReadOnly Property DateStop As Date
  ReadOnly Property IsSymbolError As Boolean
  ReadOnly Property Exchange As String
  ReadOnly Property ErrorDescription As String
  ReadOnly Property SectorName As String
  ReadOnly Property IndustryName As String
End Interface
#End Region
