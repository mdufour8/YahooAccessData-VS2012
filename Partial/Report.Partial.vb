#Region "Imports"
Imports System
Imports System.Collections.Generic
  Imports System.Collections.Concurrent
  Imports System.Runtime.CompilerServices
  Imports System.Threading.Tasks
  Imports System.Threading
  Imports System.Net
  Imports System.Runtime.Serialization.Formatters.Binary
  Imports System.IO
  Imports System.Text
Imports YahooAccessData.ExtensionService
  Imports System.Reflection
  Imports Ionic.Zip
#End Region

<Serializable()>
Partial Public Class Report
  Implements IDisposable
  Implements IEquatable(Of Report)
  Implements IRegisterKey(Of Date)
  Implements IComparable(Of Report)
  Implements IMemoryStream
  Implements IFormatData
  Implements IDateRange
  Implements IDateUpdate
  Implements IMessageInfoEvents
  Implements ISystemEvent(Of BondRate)
  Implements ISystemEvent(Of BondRate1)
  Implements ISystemEvent(Of SplitFactorFuture)
  Implements IRecordControlInfo
  Implements IStockRecordEvent


#Region "Main"
  Public Enum enuTimeFormat
    Sample     'default raw data
    Daily
    Weekly
  End Enum

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  Private colVb As Microsoft.VisualBasic.Collection
  Private MyFillDataRate As Double
  Private MyStream As Stream = Nothing
  'Private MyStreamBinaryReader As BinaryReader
  Private MyFileType As IMemoryStream.enuFileType = IMemoryStream.enuFileType.Standard
  Private IsFileAccessReadOnly As Boolean
  Private _IsFileOpen As Boolean
  Private MyFileName As String
  Private MyDateRangeStart As Date?
  Private MyDateRangeStop As Date?
  Private MyDictionaryOfStockRecordLoaded As Dictionary(Of String, String)
  Private MyStockRecordQueue As Queue(Of String)
  Private MyStockRecordQueueCount As Integer
  Private MySyncLockForRecordLoading As Object = New Object
  Private MyTaskOfLoadCache As Task(Of Boolean)
  Private MyConcurrentQueue As ConcurrentQueue(Of String)
  Private MyCancellationTokenSource As New Threading.CancellationTokenSource()
  Private MySymbolRecordLoading As String
  Private IsRecordLoading As Boolean
  Private IsRecordCancel As Boolean
  Private MyLoadToCacheLatestTick As Integer
  Private _IsFileReadEndOfDayEnabled As Boolean

  'Need to use these name to correctly capture the data from the net old object serialization
  'this object serialization is not use anymore but we may have old file
  'that require these variable for a succesful serialization load
  Private _DateStart As Date
  Private _DateStop As Date
  Private _Industries As ICollection(Of Industry) = New HashSet(Of Industry)
  Private _Sectors As ICollection(Of Sector) = New HashSet(Of Sector)
  Private _Stocks As ICollection(Of Stock) = New HashSet(Of Stock)
  Private _SplitFactorFutures As ICollection(Of SplitFactorFuture) = New HashSet(Of SplitFactorFuture)
  Private _BondRates As ICollection(Of BondRate) = New HashSet(Of BondRate)
  Private _BondRates1 As ICollection(Of BondRate1) = New LinkedHashSet(Of BondRate1, String)

  Private Const STOCK_RECORD_QUEUE_SIZE_DEFAULT As Integer = 10
#End Region
#Region "New"
  Public Sub New(ByVal Name As String)
    Me.New(Name, Now)
  End Sub

  Public Sub New(ByVal Name As String, ByVal DateStart As Date)
    MyDictionaryOfStockRecordLoaded = New Dictionary(Of String, String)
    MyStockRecordQueue = New Queue(Of String)
    MyStockRecordQueueCount = STOCK_RECORD_QUEUE_SIZE_DEFAULT
    IStockRecordInfo_Enabled = False  'by default
    With Me
      .ID = 1
      .DateStart = DateStart
      .DateStop = Me.DateStart
      .Name = Name
      .Industries = New LinkedHashSet(Of Industry, String)
      With CType(.Industries, IRegisterKey(Of String))
        .KeyID = Me.KeyID
        .KeyValue = Me.KeyValue.ToString
      End With
      .Sectors = New LinkedHashSet(Of Sector, String)
      With CType(.Sectors, IRegisterKey(Of String))
        .KeyID = Me.KeyID
        .KeyValue = Me.KeyValue.ToString
      End With
      .Stocks = New LinkedHashSet(Of Stock, String)
      With CType(.Stocks, IRegisterKey(Of String))
        .KeyID = Me.KeyID
        .KeyValue = Me.KeyValue.ToString
      End With
      .BondRates = New LinkedHashSet(Of BondRate, String)
      .SplitFactorFutures = New LinkedHashSet(Of SplitFactorFuture, String)
    End With
    If MyListHeaderInfo Is Nothing Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    End If
  End Sub

  Public Sub New()
    Me.New("", Now)
  End Sub
#End Region
#Region "Basic Function"
  Public Property Tag As String

  Public Overrides Function ToString() As String
    Dim ThisCountRecord As Integer = 0
    Dim ThisCountRecordTotal As Integer = 0
    Dim ThisCountRecordDaily As Integer = 0
    Dim ThisCountRecordDailyTotal As Integer = 0
    Dim ThisBondRateRecord As Integer = 0
    Dim ThisBondRateRecordTotal As Integer = 0
    Dim ThisSplitFactorFutureRecord As Integer = 0
    Dim ThisSplitFactorFutureRecordTotal As Integer = 0
    Dim ThisSymbol As String = ""
    Dim ThisStockIndex As Integer
    Dim I As Integer

    If Me.Exception Is Nothing Then
      If Me.Stocks.Count > 0 Then
        I = 0
        ThisStockIndex = 0
        Do
          With Me.Stocks(ThisStockIndex)
            'do not lod the record
            ThisCountRecord = .Records(IsLoadEnabled:=False).Count
            ThisCountRecordTotal = CType(.Records(IsLoadEnabled:=False), IRecordInfo).CountTotal
            ThisCountRecordDaily = .RecordsDaily(IsLoadEnabled:=False).Count
            ThisCountRecordDailyTotal = CType(.RecordsDaily(IsLoadEnabled:=False), IRecordInfo).CountTotal
            ThisSymbol = .Symbol
            'If ThisCountRecord > 0 Then Exit Do
            Exit Do
          End With
          I = I + 1
          If I >= 10 Then Exit Do
          ThisStockIndex = CInt(Math.Floor((Me.Stocks.Count) * Rnd()))
        Loop
        ThisStockIndex = ThisStockIndex + 1
      End If
      'using the _collection does not cause the data to automatically load
      Return String.Format("{0},ID:{1},Key:{2},Sector:{3},Industry:{4},BondRate:{5} of {6},SplitFactorFuture:{7} of {8},Stock:{9} ({10} of {11}).Record:({12} of {13}),.RecordDaily:{14} of {15}",
        TypeName(Me),
        Me.KeyID,
        Me.KeyValue.ToString,
        Me.Sectors.Count,
        Me.Industries.Count,
        Me.BondRates(IsLoadEnabled:=False).Count,
        CType(Me.BondRates(IsLoadEnabled:=False), IRecordInfo).CountTotal,
        Me.SplitFactorFutures(IsLoadEnabled:=False).Count,
        CType(Me.SplitFactorFutures(IsLoadEnabled:=False), IRecordInfo).CountTotal,
        ThisSymbol,
        ThisStockIndex,
        Me.Stocks.Count,
        ThisCountRecord,
        ThisCountRecordTotal,
        ThisCountRecordDaily,
        ThisCountRecordDailyTotal)
    Else
      'return the error
      Return String.Format("{0},ID:{1},Key:{2},Error:{3}", TypeName(Me), Me.KeyID, Me.KeyValue.ToString, Me.Exception.MessageAll)
    End If
  End Function

  Public Property TimeFormat As enuTimeFormat

  Public ReadOnly Property IsFileReadEndOfDay As Boolean
    Get
      Dim ThisFilePathRecordNameForEndOfDay As String

      If Me.IsFileOpen Then
        If Me.FileType = IMemoryStream.enuFileType.RecordIndexed Then
          ThisFilePathRecordNameForEndOfDay = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Me.FileName), "Stock\RecordEndOfDay")
          If My.Computer.FileSystem.DirectoryExists(ThisFilePathRecordNameForEndOfDay) Then
            Return True
          End If
        End If
      End If
      Return False
    End Get
  End Property

  Public Property IsFileReadEndOfDayEnabled As Boolean
    Get
      Return _IsFileReadEndOfDayEnabled
    End Get
    Set(value As Boolean)
      If value = True Then
        'some condition apply for true
        If Me.IsFileReadEndOfDay Then
          _IsFileReadEndOfDayEnabled = True
        Else
          _IsFileReadEndOfDayEnabled = False
        End If
      Else
        _IsFileReadEndOfDayEnabled = False
      End If
    End Set
  End Property
#End Region
#Region "Collection Properties"
  Public Overridable Property Industries As ICollection(Of Industry)
    Get
      Return _Industries
    End Get
    Set(value As ICollection(Of Industry))
      _Industries = value
    End Set
  End Property
  Public Overridable Property Sectors As ICollection(Of Sector)
    Get
      Return _Sectors
    End Get
    Set(value As ICollection(Of Sector))
      _Sectors = value
    End Set
  End Property

  Public Overridable Property Stocks As ICollection(Of Stock)
    Get
      Return _Stocks
    End Get
    Set(value As ICollection(Of Stock))
      _Stocks = value
    End Set
  End Property

  Public Overridable Property SplitFactorFutures As ICollection(Of SplitFactorFuture)
    Get
      If TypeOf _SplitFactorFutures Is IDataVirtual Then
        DirectCast(_SplitFactorFutures, IDataVirtual).Load()
      End If
      Return _SplitFactorFutures
    End Get
    Set(value As ICollection(Of SplitFactorFuture))
      _SplitFactorFutures = value
    End Set
  End Property

  Public Overridable Property SplitFactorFutures(ByVal IsLoadEnabled As Boolean) As ICollection(Of SplitFactorFuture)
    Get
      If IsLoadEnabled Then
        If TypeOf _SplitFactorFutures Is IDataVirtual Then
          DirectCast(_SplitFactorFutures, IDataVirtual).Load()
        End If
      End If
      Return _SplitFactorFutures
    End Get
    Set(value As ICollection(Of SplitFactorFuture))
      _SplitFactorFutures = value
    End Set
  End Property

  Public Overridable Property BondRates As ICollection(Of BondRate)
    Get
      'it is important to test for that interface because the old file format did not support it
      If TypeOf _BondRates Is IDataVirtual Then
        DirectCast(_BondRates, IDataVirtual).Load()
      End If
      Return _BondRates
    End Get
    Set(value As ICollection(Of BondRate))
      _BondRates = value
    End Set
  End Property

  Public Overridable Property BondRates(ByVal IsLoadEnabled As Boolean) As ICollection(Of BondRate)
    Get
      'it is important to test for that interface because the old file format did not support it
      If IsLoadEnabled Then
        If TypeOf _BondRates Is IDataVirtual Then
          DirectCast(_BondRates, IDataVirtual).Load()
        End If
      End If
      Return _BondRates
    End Get
    Set(value As ICollection(Of BondRate))
      _BondRates = value
    End Set
  End Property

  Public Overridable Property BondRates1 As ICollection(Of BondRate1)
    Get
      If TypeOf _BondRates Is IDataVirtual Then
        DirectCast(_BondRates, IDataVirtual).Load()
      End If
      'it is important to test for that interface because the old file format did not support it
      If TypeOf _BondRates1 Is IDataVirtual Then
        DirectCast(_BondRates1, IDataVirtual).Load()
      End If
      Return _BondRates1
    End Get
    Set(value As ICollection(Of BondRate1))
      _BondRates1 = value
    End Set
  End Property

  Public Overridable Property BondRates1(ByVal IsLoadEnabled As Boolean) As ICollection(Of BondRate1)
    Get
      If IsLoadEnabled Then
        If TypeOf _BondRates Is IDataVirtual Then
          DirectCast(_BondRates, IDataVirtual).Load()
        End If
        'it is important to test for that interface because the old file format did not support it
        If TypeOf _BondRates1 Is IDataVirtual Then
          DirectCast(_BondRates1, IDataVirtual).Load()
        End If
      End If
      Return _BondRates1
    End Get
    Set(value As ICollection(Of BondRate1))
      _BondRates1 = value
    End Set
  End Property

  Public ReadOnly Property StockExchanges As IEnumerable(Of String)
    Get
      Dim ThisDictionaryOfExchanges As New Dictionary(Of String, String)
      For Each ThisStock In Me.Stocks
        If ThisDictionaryOfExchanges.ContainsKey(ThisStock.Exchange) = False Then
          ThisDictionaryOfExchanges.Add(ThisStock.Exchange, ThisStock.Exchange)
        End If
      Next
      Return ThisDictionaryOfExchanges.Values
    End Get
  End Property
#End Region
#Region "Maintenance Functions"
  Friend ReadOnly Property ToStream As Stream
    Get
      Return MyStream
    End Get
  End Property

  Friend ReadOnly Property FileType As IMemoryStream.enuFileType
    Get
      Return MyFileType
    End Get
  End Property

  Friend ReadOnly Property IsReadOnly As Boolean
    Get
      Return IsFileAccessReadOnly
    End Get
  End Property

  Public Sub Clear()
    With Me
      .ID = 1
      .DateStart = Now
      .DateStop = Me.DateStart
      .Name = ""
      .Exception = Nothing
    End With
    _Stocks.Clear()
    _Sectors.Clear()
    _Industries.Clear()
    _BondRates.Clear()
    _SplitFactorFutures.Clear()
  End Sub

  Public Property DateStart As Date
    Get
      Return _DateStart
    End Get
    Set(value As Date)
      _DateStart = value
      If _DateStart > _DateStop Then
        Me.DateStop = Me.DateStart
      End If
    End Set
  End Property

  Public Property DateStop As Date
    Get
      Return _DateStop
    End Get
    Set(value As Date)
      _DateStop = value
      If _DateStop < _DateStart Then
        Me.DateStart = _DateStop
      End If
    End Set
  End Property

  Public Function NumberOfMarketTradingDays() As Integer
    Return YahooAccessData.ReportDate.MarketTradingDeltaDays(Me.DateStart, Me.DateStop) + 1
  End Function

  Public Function Add(ByVal Report As YahooAccessData.Report) As Boolean
    Dim ThisStockLocal As YahooAccessData.Stock
    Dim ThisStockNew As YahooAccessData.Stock
    Dim ThisSectorLocal As YahooAccessData.Sector
    Dim ThisIndustryLocal As YahooAccessData.Industry
    Dim ThisSplitFactorFutureLocal As YahooAccessData.SplitFactorFuture
    Dim ThisBondRateLocal As YahooAccessData.BondRate
    Dim ThisResult As Boolean = False

    Try
      With Report
        'If Me.BondRates.Count = 0 Then Return True
        'do not add report that contain an error
        If .Exception IsNot Nothing Then
          'Throw New Exception(String.Format("Unable to add a report with error:{0}{1}", vbCr, .Exception.MessageAll))
          Return ThisResult
        End If
        'make sure we can add only newer data
        If Me.DateStart = Me.DateStop Then
          'first time the date are initialized
          Me.DateStart = .DateStart.Date    'this set the date a the beginnning of the day
          Me.DateStop = .DateStop
        End If
        If .DateStart < Me.DateStart Then
          Me.DateStart = .DateStart
        End If
        If .DateStop > Me.DateStop Then
          Me.DateStop = .DateStop
        End If
        'check the sector list
        For Each ThisSector In .Sectors
          ThisSectorLocal = Me.Sectors.ToSearch.Find(ThisSector.KeyValue)
          If ThisSectorLocal Is Nothing Then
            ThisSectorLocal = ThisSector.CopyDeep(Me, IsIgnoreID:=True)
          End If
        Next
        'check the Industry list
        For Each ThisIndustry In .Industries
          ThisIndustryLocal = Me.Industries.ToSearch.Find(ThisIndustry.KeyValue)
          If ThisIndustryLocal Is Nothing Then
            ThisIndustryLocal = ThisIndustry.CopyDeep(Me, IsIgnoreID:=True)
          End If
        Next
        For Each ThisSplitFactorFuture In .SplitFactorFutures
          ThisSplitFactorFutureLocal = Me.SplitFactorFutures.ToSearch.Find(ThisSplitFactorFuture.KeyValue)
          If ThisSplitFactorFutureLocal Is Nothing Then
            ThisSplitFactorFutureLocal = ThisSplitFactorFuture.CopyDeep(Me, IsIgnoreID:=True)
          End If
        Next
        For Each ThisBondRate In .BondRates
          ThisBondRateLocal = Me.BondRates.ToSearch.Find(ThisBondRate.KeyValue)
          If ThisBondRateLocal Is Nothing Then
            ThisBondRateLocal = ThisBondRate.CopyDeep(Me, IsIgnoreID:=True)
          End If
        Next
        For Each ThisStockNew In .Stocks
          ThisStockLocal = Me.Stocks.ToSearch.Find(ThisStockNew.Symbol)
          If ThisStockLocal Is Nothing Then
            'this stock does not exist yet
            ThisStockLocal = ThisStockNew.CopyDeep(Me, IsIgnoreID:=True)
          Else
            'stock already exist
            ThisStockLocal.Add(ThisStockNew)
          End If
        Next
      End With
      ThisResult = True
    Catch ex As Exception
      Me.Exception = New Exception(String.Format("Error adding report..."), ex)
      ThisResult = False
    End Try
    Return ThisResult
  End Function

  Public Function StockAdd(ByVal StockSymbol As String, ByVal SectorName As String, ByVal IndustryName As String) As YahooAccessData.Stock
    Dim ThisStock As YahooAccessData.Stock = Nothing
    Dim ThisSector As YahooAccessData.Sector = Nothing
    Dim ThisIndustry As YahooAccessData.Industry = Nothing

    'get rid of the empty string
    If StockSymbol Is Nothing Then StockSymbol = ""
    If SectorName Is Nothing Then SectorName = ""
    If IndustryName Is Nothing Then IndustryName = ""

    If SectorName = "Indices" Then
      SectorName = "Indices"
    End If
    ThisSector = Me.Sectors.ToSearch.Find(SectorName)
    ThisIndustry = Me.Industries.ToSearch.Find(IndustryName)
    ThisStock = Me.Stocks.ToSearch.Find(StockSymbol)

    If ThisSector Is Nothing Then
      ThisSector = New YahooAccessData.Sector(Me, SectorName)
    End If
    If ThisIndustry Is Nothing Then
      ThisIndustry = New YahooAccessData.Industry(Me, ThisSector, IndustryName)
    End If
    If ThisStock Is Nothing Then
      If StockSymbol.Length > 0 Then
        'note passing the reference 'Me' automatically add the stock in the base report collection
        ThisStock = New YahooAccessData.Stock(Me, ThisSector, ThisIndustry, Symbol:=StockSymbol)
      End If
    End If
    If ThisStock IsNot Nothing Then
      With ThisStock
        If (.Sector IsNot ThisSector) Or (.Industry IsNot ThisIndustry) Then
          'break all relation with the actual industry and sector if it exist
          If .Sector IsNot Nothing Then
            If .Sector IsNot ThisSector Then
              'remove the stock relation for this sector
              .Sector.Stocks.Remove(ThisStock)
              .SectorID = 0
            End If
          End If
          If .Industry IsNot Nothing Then
            If .Industry IsNot ThisIndustry Then
              'remove the stock link with this industry
              With .Industry.Stocks
                .Remove(ThisStock)
                ThisStock.IndustryID = 0
                If .Count = 0 Then
                  'no more stock
                  'we can remove the link to sector
                  ThisStock.Sector.Industries.Remove(ThisStock.Industry)
                End If
              End With
            End If
          End If
          'all previous link have been removed
          'we can update with the new link
          If .Sector IsNot ThisSector Then
            .Sector = ThisSector
            .Sector.Stocks.Add(ThisStock)
            .SectorID = .Sector.ID
          End If
          If .Industry IsNot ThisIndustry Then
            .Industry = ThisIndustry
            .Industry.Stocks.Add(ThisStock)
            .IndustryID = .Industry.ID
          End If
        End If
      End With
    End If
    Return ThisStock
  End Function

  Public Function StockAdd(ByVal StockNew As YahooAccessData.Stock) As YahooAccessData.Stock
    Dim ThisStock As YahooAccessData.Stock
    Dim ThisSector As YahooAccessData.Sector = Nothing
    Dim ThisIndustry As YahooAccessData.Industry = Nothing


    If StockNew.Sector IsNot Nothing Then
      ThisSector = Me.Sectors.ToSearch.Find(StockNew.Sector.KeyValue)
      If ThisSector Is Nothing Then
        ThisSector = StockNew.Sector.CopyDeep(Me, IsIgnoreID:=True)
      End If
      If ThisSector.Exception IsNot Nothing Then
        Me.Exception = New Exception("Sector addition error...", ThisSector.Exception)
        Return Nothing
      End If
    End If
    If StockNew.Industry IsNot Nothing Then
      ThisIndustry = Me.Industries.ToSearch.Find(StockNew.Industry.KeyValue)
      If ThisIndustry Is Nothing Then
        ThisIndustry = StockNew.Industry.CopyDeep(Me, IsIgnoreID:=True)
      End If
      If ThisIndustry.Exception IsNot Nothing Then
        Me.Exception = New Exception("Industry addition error...", Me.Exception)
        Return Nothing
      End If
    End If
    ThisStock = Me.Stocks.ToSearch.Find(StockNew.KeyValue)
    If ThisStock Is Nothing Then
      'stock does not exist yet locally
      ThisStock = StockNew.CopyDeep(Me, IsIgnoreID:=True)
      If ThisStock.Exception IsNot Nothing Then
        Me.Exception = New Exception("Stock addition error...", Me.Exception)
        Return Nothing
      End If
    Else
      'stock already exist in the list
      With ThisStock
        If StockNew.DateStop > .DateStop Then
          'this is a new record
          'check if it match the current relation with sector and industry
          If ThisStock.Sector.KeyValue <> ThisSector.KeyValue Then
            'change to the new relation
            'remove the stock relation for this sector
            Debug.Assert(False)
            .Sector.Stocks.Remove(ThisStock)
            .Sector = ThisSector
            .Sector.Stocks.Add(ThisStock)
            .SectorID = .Sector.ID
            'Me.Exception = New Exception(String.Format("Invalid sector for stock {0}", ThisStock.KeyValue))
            'Return Nothing
          End If
          If ThisStock.Industry.KeyValue <> ThisIndustry.KeyValue Then
            'remove the stock link with this industry
            Debug.Assert(False)
            With .Industry.Stocks
              .Remove(ThisStock)
              ThisStock.IndustryID = 0
              If .Count = 0 Then
                'no more stock
                'we can remove the link to sector
                ThisStock.Sector.Industries.Remove(ThisStock.Industry)
              End If
            End With
            .Industry = ThisIndustry
            .Industry.Stocks.Add(ThisStock)
            .IndustryID = .Industry.ID
          End If
          If .Sector.Industries.Contains(.Industry) = False Then
            .Sector.Industries.Add(.Industry)
          End If
          .Name = StockNew.Name
          .IsOption = StockNew.IsOption
          .Exchange = StockNew.Exchange
          .IsSymbolError = StockNew.IsSymbolError
          .ErrorDescription = StockNew.ErrorDescription
          If StockNew.IsSymbolError Then
            'add a new error
            Dim ThisStockError = New YahooAccessData.StockError
            With ThisStockError
              .DateUpdate = StockNew.DateStop
              .Description = StockNew.ErrorDescription
              .Symbol = StockNew.Symbol
              .Stock = ThisStock
              .StockID = .Stock.ID
            End With
            .StockErrors.Add(ThisStockError)
          End If
        End If
      End With
    End If
    Return ThisStock
  End Function

  Public Sub StockSymbolChange(ByVal StockSymbolOld As String, ByVal StockSymbolNew As String)
    Dim ThisStock = Me.Stocks.Item(StockSymbolOld)
    If ThisStock IsNot Nothing Then
      ThisStock.Symbol = StockSymbolNew
    End If
  End Sub

  Public Sub StockSymbolUpdateChange()
    For Each ThisStock In Me.StockErrorList
      With ThisStock
        Select Case .ErrorDescription
          Case "Ticker symbol has changed to:"

          Case "No such ticker symbol."



        End Select
      End With
    Next
  End Sub

  Public Sub Remove(ByVal StockSymbol As String)
    Dim ThisStock = Me.Stocks.Item(StockSymbol)
    If ThisStock IsNot Nothing Then
      Dim ThisSector As Sector = ThisStock.Sector
      Dim ThisIndustry As Industry = ThisStock.Industry

      If ThisSector IsNot Nothing Then
        'remove the stock link relation for this sector
        ThisSector.Stocks.Remove(ThisStock)
      End If
      If ThisIndustry IsNot Nothing Then
        'remove the stock link with this industry
        With ThisIndustry.Stocks
          .Remove(ThisStock)
          If .Count = 0 Then
            'no more stock
            'we can remove the link to sector
            ThisSector.Industries.Remove(ThisIndustry)
          End If
        End With
      End If
    End If
  End Sub

  Public Async Function CopyDeepAsync() As Task(Of YahooAccessData.Report)
    Dim ThisTask = New Task(Of YahooAccessData.Report)(
      Function()
        Return Me.CopyDeep
      End Function)
    'start the task immediately
    ThisTask.Start()
    Await ThisTask
    Return ThisTask.Result
  End Function

  Public Function CopyDeep(Optional IsThreaded As Boolean = False) As YahooAccessData.Report
    Dim ThisStream As Stream = New MemoryStream

    Me.SerializeSaveTo(ThisStream)
    Return YahooAccessData.ReportFrom.Load(ThisStream)
  End Function

  ''' <summary>
  ''' check if the report is empty and contain no data
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function IsEmpty() As Boolean
    With Me
      If .Sectors.Count = 0 Then
        If .Industries.Count = 0 Then
          If .Stocks.Count = 0 Then
            Return True
          End If
        End If
      End If
    End With
    Return False
  End Function

  Private Function CopyDeepLocal() As YahooAccessData.Report
    Dim ThisReport = New YahooAccessData.Report

    With ThisReport
      .ID = Me.ID
      .Name = Me.Name
      .DateStart = Me.DateStart
      .DateStop = Me.DateStop
    End With
    Me.Sectors.CopyDeep(ThisReport)
    'update the collection dictionary for quick element search
    Me.Industries.CopyDeep(ThisReport)
    Me.SplitFactorFutures.CopyDeep(ThisReport)
    Me.BondRates.CopyDeep(ThisReport)
    Me.Stocks.CopyDeep(ThisReport)
    Return ThisReport
  End Function

  Public ReadOnly Property StockErrorList As IEnumerable(Of Stock)
    Get
      Try
        Dim ThisStockListWithError = Me.Stocks.Where(Function(ThisStock As YahooAccessData.Stock) ThisStock.IsSymbolError)
        Return ThisStockListWithError
      Catch ex As Exception
        Return New List(Of Stock)
      End Try
    End Get
  End Property

  ''' <summary>
  ''' Percentage of stock with valid updated record
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks>Can be use to estimate the success rate of the downloading software</remarks>
  Public ReadOnly Property FillDataRate As Double
    Get
      Dim ThisZero As Double = 0
      Dim ThisStockCount As Integer
      Dim ThisStockErrorCount As Integer
      Try
        ThisStockCount = Me.Stocks.Count
      Catch
      End Try
      If ThisStockCount > 0 Then
        'count the number of stock in error
        ThisStockErrorCount = Me.StockErrorList.Count
        Return 100 * (ThisStockCount - ThisStockErrorCount) / ThisStockCount
      Else
        Return ThisZero
      End If
    End Get
  End Property
#End Region
#Region "Date Calculation"
  ''' <summary>
  ''' Automatically calculate all the market day trading holiday
  ''' see: http://www.rightline.net/calendar/market-holidays.html
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>

  Public Function DayOfTradeNext(ByVal DateValue As Date) As Date
    'Static DateValueInLast As Date
    'Static DateValueOutLast As Date

    'If DateValueInLast = DateValue Then
    '	Return DateValueOutLast
    'End If
    'DateValueInLast = DateValue
    'DateValueOutLast = DayOfTradeNextLocal(DateValue, True)
    'Return DateValueOutLast
    Return ReportDate.DayOfTradeNext(DateValue)
  End Function

  Public Function DayOfTradeNext(
    ByVal DateValue As Date,
    ByVal MarketOpenTime As Date,
    ByVal MarketCloseTime As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    ''execute all the time here
    'Return DayOfTradeNextLocal(
    '	DateValue,
    '	False,
    '	MarketOpenTime,
    '	MarketCloseTime,
    '	MarketDelayMinutes)
    Return ReportDate.DayOfTradeNext(
      DateValue,
      MarketOpenTime,
      MarketCloseTime,
      MarketDelayMinutes)
  End Function

  'Private Function DayOfTradeNextLocal(
  '	ByVal DateValue As Date,
  '	ByVal IsMarketTimeToDefault As Boolean,
  '	Optional ByVal MarketOpenTime As Date = #9:30:00 AM#,
  '	Optional ByVal MarketCloseTime As Date = #4:00:00 PM#,
  '	Optional ByVal MarketDelayMinutes As Integer = 15) As Date

  '	Dim ThisDateValueIn As Date = DateValue
  '	'adjust the open and closing time as needed
  '	'set the default value
  '	Dim ThisMarketOpenTimeSpan As New TimeSpan(Hour(MarketOpenTime), (Minute(MarketOpenTime)) + MarketDelayMinutes, Second(MarketOpenTime))
  '	Dim ThisMarketCloseTimeSpan As New TimeSpan(Hour(MarketCloseTime), (Minute(MarketCloseTime)) + MarketDelayMinutes, Second(MarketCloseTime))
  '	If IsMarketTimeToDefault Then
  '		'the user did not provide any time for closing
  '		'check for: The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on Friday, the day after Thanksgiving.
  '		If DateValue.Month = 11 Then
  '			If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '				'The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on Friday, 
  '				'the day after Thanksgiving.
  '				If DateValue.Date = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday).AddDays(1) Then
  '					MarketCloseTime = #1:00:00 PM#
  '					ThisMarketCloseTimeSpan = New TimeSpan(Hour(MarketCloseTime), (Minute(MarketCloseTime)) + MarketDelayMinutes, Second(MarketCloseTime))
  '				End If
  '			End If
  '		End If
  '	End If
  '	'adjust for normal daily trading close time
  '	If DateValue.TimeOfDay > ThisMarketCloseTimeSpan Then
  '		'the trading is for the next day
  '		DateValue = DateValue.AddDays(1)
  '	End If
  '	'remove the time component
  '	DateValue = DateValue.Date
  '	'correct for weekend
  '	Select Case DateValue.DayOfWeek
  '	Case System.DayOfWeek.Sunday
  '		DateValue = DateValue.AddDays(1)
  '	Case System.DayOfWeek.Saturday
  '		DateValue = DateValue.AddDays(2)
  '	End Select
  '	'correct for the holiday
  '	Select Case DateValue.Month
  '	Case 1
  '		'check for new year i.e.
  '		'New Years' Day (January 1) in 2011 falls on a Saturday. 
  '		'The rules of the applicable exchanges state that when a holiday falls on a Saturday, 
  '		'the preceding Friday is observed unless the Friday is the end of a monthly or yearly 
  '		'accounting period. In this case, Friday, December 31, 2010 is the end of both a monthly 
  '		'and yearly accounting period; therefore the exchanges will be open that day 
  '		'and the following Monday.
  '		'When any stock market holiday falls on a Saturday, 
  '		'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
  '		'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '		Select Case DateValue.Day
  '		Case 1
  '			'this is January first
  '			If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '				DateValue = DateValue.AddDays(3)
  '			Else
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		Case 2
  '			'Jan 2 could be an holiday if the new year did fall on Sunday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'new year was on Sunday and in this case the following rule apply:
  '				'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		Case Else
  '			'check for other Holiday happening on Monday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'Martin Luther King is observed on the third Monday of January each year
  '				If DateValue = DateFromDayOfWeek(DateValue.Year, 1, 3, FirstDayOfWeek.Monday) Then
  '					'this is the Martin Luther King holiday
  '					DateValue = DateValue.AddDays(1)
  '				End If
  '			End If
  '		End Select
  '	Case 2
  '		'Washington's Birthday is a United States federal holiday 
  '		'celebrated on the third Monday of February in honor of George Washington, the first President of the United States.
  '		'check for Holliday
  '		'Martin Luther King is observed on the third Monday of January each year
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 2, 3, FirstDayOfWeek.Monday) Then
  '				'this is Washington's Birthday holiday
  '				'the trading is for the next day
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		End If
  '	Case 3, 4
  '		'check for good Friday
  '		If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '			'check if this is good Friday
  '			'the market is not open on good Friday in the US and Canada
  '			'EasterDate is Sunday
  '			If DateValue = EasterDate(DateValue).AddDays(-2) Then
  '				'this is good Friday
  '				'next Monday is not an holiday in the US but it is in Canada
  '				'for now assume the market is the US and the market is trading on Monday
  '				DateValue = DateValue.AddDays(3)
  '			End If
  '		End If
  '	Case 5
  '		'Memorial Day is a United States federal holiday observed on the last Monday of May.
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			'check if this is good Friday
  '			'the market is not open on good Friday in the US and Canada
  '			'EasterDate is Sunday
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 5, -1, FirstDayOfWeek.Monday) Then
  '				'this is the Memorial Day holiday
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		End If
  '	Case 7
  '		'check for Independence day always on July 4
  '		'When any stock market holiday falls on a Saturday, 
  '		'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
  '		'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '		If DateValue.Day = 4 Then
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				DateValue = DateValue.AddDays(1)
  '			Else
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		Else
  '			'Friday July 3 is also and holiday for Independence day on Saturday
  '			If DateValue.Day = 3 Then
  '				If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '					DateValue = DateValue.AddDays(3)
  '				End If
  '			ElseIf DateValue.Day = 5 Then
  '				'Monday July 5 is also and holiday for Independence day on Sunday
  '				If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '					DateValue = DateValue.AddDays(1)
  '				End If
  '			End If
  '		End If
  '	Case 9
  '		'Labor Day:
  '		'The first Monday in September, observed as a holiday in the United States and Canada
  '		'in honor of working people.
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 9, 1, FirstDayOfWeek.Monday) Then
  '				'this is labor day
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		End If
  '	Case 11
  '		'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
  '		'Traditionally, it has been a time to give thanks to God, friends, and family.
  '		'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
  '		'and in the United States, it is celebrated on the fourth Thursday of November. 
  '		'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
  '		If DateValue.DayOfWeek = DayOfWeek.Thursday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday) Then
  '				'this is the Thanksgiving Day holiday
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		End If
  '	Case 12
  '		'Christmas
  '		Select Case DateValue.Day
  '		Case 25
  '			'this is Christmas
  '			If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '				DateValue = DateValue.AddDays(3)
  '			Else
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		Case 26
  '			'Dec 26 could be an holiday if Christmas did fall on Sunday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'Christmas was on Sunday and in this case the following rule apply:
  '				'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '				DateValue = DateValue.AddDays(1)
  '			End If
  '		End Select
  '	End Select
  '	If DateValue.Date > ThisDateValueIn.Date Then
  '		DateValue = DateValue.Add(ThisMarketOpenTimeSpan)
  '	ElseIf DateValue.Date = ThisDateValueIn.Date Then
  '		If ThisDateValueIn.TimeOfDay < ThisMarketOpenTimeSpan Then
  '			DateValue = DateValue.Add(ThisMarketOpenTimeSpan)
  '		Else
  '			DateValue = ThisDateValueIn
  '		End If
  '	Else
  '		'should never happen
  '		Debug.Assert(False)
  '	End If
  '	Return DateValue
  'End Function

  Public Function DayOfTrade(ByVal DateValue As Date) As Date
    'Static DateValueInLast As Date
    'Static DateValueOutLast As Date

    'If DateValueInLast = DateValue Then
    '	Return DateValueOutLast
    'End If
    'DateValueInLast = DateValue
    'DateValueOutLast = DayOfTradeLocal(DateValue, True)
    'Return DateValueOutLast
    Return ReportDate.DayOfTrade(DateValue)
  End Function

  Public Function DayOfTrade(
    ByVal DateValue As Date,
    ByVal MarketOpenTime As Date,
    ByVal MarketCloseTime As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    ''execute all the time here
    'Return DayOfTradeLocal(
    '	DateValue,
    '	False,
    '	MarketOpenTime,
    '	MarketCloseTime,
    '	MarketDelayMinutes)
    Return ReportDate.DayOfTrade(
      DateValue,
      MarketOpenTime,
      MarketCloseTime,
      MarketDelayMinutes)
  End Function

  'Private Function DayOfTradeLocal(
  '	ByVal DateValue As Date,
  '	ByVal IsMarketTimeToDefault As Boolean,
  '	Optional ByVal MarketOpenTime As Date = #9:30:00 AM#,
  '	Optional ByVal MarketCloseTime As Date = #4:00:00 PM#,
  '	Optional ByVal MarketDelayMinutes As Integer = 15) As Date

  '	Dim ThisDateValueIn As Date = DateValue
  '	'adjust the open and closing time as needed
  '	'set the default value
  '	Dim ThisMarketOpenTimeSpan As New TimeSpan(Hour(MarketOpenTime), (Minute(MarketOpenTime)) + MarketDelayMinutes, Second(MarketOpenTime))
  '	Dim ThisMarketCloseTimeSpan As New TimeSpan(Hour(MarketCloseTime), (Minute(MarketCloseTime)) + MarketDelayMinutes, Second(MarketCloseTime))
  '	'adjust for normal daily trading time
  '	If DateValue.TimeOfDay <
  '		MarketOpenTime.TimeOfDay.Add(New TimeSpan(hours:=0, minutes:=MarketDelayMinutes, seconds:=0)) Then
  '		'the trading is for the previous day
  '		DateValue = DateValue.AddDays(-1)
  '	End If
  '	'remove the time component
  '	DateValue = DateValue.Date
  '	'correct for weekend
  '	Select Case DateValue.DayOfWeek
  '	Case System.DayOfWeek.Sunday
  '		DateValue = DateValue.AddDays(-2)
  '	Case System.DayOfWeek.Saturday
  '		DateValue = DateValue.AddDays(-1)
  '	End Select
  '	'correct for the holiday
  '	Select Case DateValue.Month
  '	Case 1
  '		'check for new year i.e.
  '		'New Years' Day (January 1) in 2011 falls on a Saturday. 
  '		'The rules of the applicable exchanges state that when a holiday falls on a Saturday, 
  '		'the preceding Friday is observed unless the Friday is the end of a monthly or yearly 
  '		'accounting period. In this case, Friday, December 31, 2010 is the end of both a monthly 
  '		'and yearly accounting period; therefore the exchanges will be open that day 
  '		'and the following Monday.
  '		'The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on 
  '		'Friday, November 25, 2011 (the day after Thanksgiving). 

  '		'Although the day after Thanksgiving (Friday) is not an official holiday, 
  '		'the market has a tradition of closing at 1:00 p.m. ET. 
  '		'When any stock market holiday falls on a Saturday, 
  '		'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
  '		'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '		Select Case DateValue.Day
  '		Case 1
  '			'this is January first
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				DateValue = DateValue.AddDays(-3)
  '			Else
  '				DateValue = DateValue.AddDays(-1)
  '			End If
  '		Case 2
  '			'Jan 2 could be an holiday if the new year did fall on Sunday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'new year was on Sunday and in this case the following rule apply:
  '				'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '				DateValue = DateValue.AddDays(-3)
  '			End If
  '		Case Else
  '			'check for other Holiday happening on Monday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'Martin Luther King is observed on the third Monday of January each year
  '				If DateValue = DateFromDayOfWeek(DateValue.Year, 1, 3, FirstDayOfWeek.Monday) Then
  '					'this is the Martin Luther King holiday
  '					DateValue = DateValue.AddDays(-3)
  '				End If
  '			End If
  '		End Select
  '	Case 2
  '		'Washington's Birthday is a United States federal holiday 
  '		'celebrated on the third Monday of February in honor of George Washington, the first President of the United States.
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 2, 3, FirstDayOfWeek.Monday) Then
  '				'this is the Washington's Birthday holiday
  '				DateValue = DateValue.AddDays(-3)
  '			End If
  '		End If
  '	Case 3, 4
  '		'check for good Friday
  '		If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '			If DateValue = EasterDate(DateValue).AddDays(-2) Then
  '				'this is good Friday
  '				'next Monday is not an holiday in the US but it is in Canada
  '				DateValue = DateValue.AddDays(-1)
  '			End If
  '		End If
  '	Case 5
  '		'Memorial Day is a United States federal holiday observed on the last Monday of May.
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 5, -1, FirstDayOfWeek.Monday) Then
  '				'this is the Memorial Day holiday
  '				'subtract 3 days to get to the previous Friday
  '				DateValue = DateValue.AddDays(-3)
  '			End If
  '		End If
  '	Case 7
  '		'check for Independence day always on July 4
  '		'When any stock market holiday falls on a Saturday, 
  '		'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
  '		'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '		If DateValue.Day = 4 Then
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				DateValue = DateValue.AddDays(-3)
  '			Else
  '				DateValue = DateValue.AddDays(-1)
  '			End If
  '		Else
  '			'Friday July 3 is also and holiday for Independence day on Saturday
  '			If DateValue.Day = 3 Then
  '				If DateValue.DayOfWeek = DayOfWeek.Friday Then
  '					DateValue = DateValue.AddDays(-1)
  '				End If
  '			ElseIf DateValue.Day = 5 Then
  '				'Monday July 5 is also and holiday for Independence day on Sunday
  '				If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '					DateValue = DateValue.AddDays(-3)
  '				End If
  '			End If
  '		End If
  '	Case 9
  '		'Labor Day:
  '		'The first Monday in September, observed as a holiday in the United States and Canada
  '		'in honor of working people.
  '		If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '			If DateValue = DateFromDayOfWeek(DateValue.Year, 9, 1, FirstDayOfWeek.Monday) Then
  '				'this is labor day
  '				'subtract 3 days to get to the previous Friday
  '				DateValue = DateValue.AddDays(-3)
  '			End If
  '		End If
  '	Case 11
  '		'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
  '		'Traditionally, it has been a time to give thanks to God, friends, and family.
  '		'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
  '		'and in the United States, it is celebrated on the fourth Thursday of November. 
  '		'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
  '		Dim ThisDateForThanksGiving As Date
  '		ThisDateForThanksGiving = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday)
  '		If DateValue.DayOfWeek = DayOfWeek.Thursday Then
  '			If DateValue = ThisDateForThanksGiving Then
  '				'this is the Thanksgiving Day holiday
  '				'subtract 1 days to get to the previous day
  '				DateValue = DateValue.AddDays(-1)
  '			End If
  '		Else
  '			If IsMarketTimeToDefault Then
  '				'the user did not provide any time for closing
  '				'check for: The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on Friday, the day after Thanksgiving.
  '				'The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on Friday, 
  '				'the day after Thanksgiving.
  '				If DateValue = ThisDateForThanksGiving.AddDays(1) Then
  '					MarketCloseTime = #1:00:00 PM#
  '					ThisMarketCloseTimeSpan = New TimeSpan(Hour(MarketCloseTime), (Minute(MarketCloseTime)) + MarketDelayMinutes, Second(MarketCloseTime))
  '				End If
  '			End If
  '		End If
  '	Case 12
  '		'Christmas
  '		Select Case DateValue.Day
  '		Case 25
  '			'this is Christmas
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				DateValue = DateValue.AddDays(-3)
  '			Else
  '				DateValue = DateValue.AddDays(-1)
  '			End If
  '		Case 26
  '			'Dec 26 could be an holiday if Christmas did fall on Sunday
  '			If DateValue.DayOfWeek = DayOfWeek.Monday Then
  '				'Christmas was on Sunday and in this case the following rule apply:
  '				'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
  '				DateValue = DateValue.AddDays(-3)
  '			End If
  '		End Select
  '	End Select
  '	If DateValue.Date < ThisDateValueIn.Date Then
  '		DateValue = DateValue.Add(ThisMarketCloseTimeSpan)
  '	ElseIf DateValue.Date = ThisDateValueIn.Date Then
  '		If ThisDateValueIn.TimeOfDay > ThisMarketCloseTimeSpan Then
  '			'DateValue does not contain the time value
  '			DateValue = DateValue.Add(ThisMarketCloseTimeSpan)
  '		Else
  '			DateValue = ThisDateValueIn
  '		End If
  '	Else
  '		'should never happen
  '		Debug.Assert(False)
  '	End If
  '	Return DateValue
  'End Function

  'Private Function DateFromDayOfWeek(
  '	ByVal ThisYear As Integer,
  '	ByVal ThisMonth As Integer, _
  '	ByVal NumberOfFirstDayOfWeek As Integer,
  '	ByVal ThisFirstDayOfWeek As Microsoft.VisualBasic.FirstDayOfWeek) As Date

  '	Dim ThisNumberDeltaDay As Integer
  '	Dim ThisDate As Date

  '	'note: The value returned by the Weekday function corresponds to the values 
  '	'of the FirstDayOfWeek enumeration; that is, 1 indicates Sunday and 7 indicates Saturday
  '	Select Case NumberOfFirstDayOfWeek
  '	Case 0
  '		ThisDate = DateSerial(ThisYear, ThisMonth, 1)
  '	Case Is > 0
  '		Dim ThisFirstDateOfMonth As Date
  '		ThisFirstDateOfMonth = DateSerial(ThisYear, ThisMonth, 1)
  '		'Number of day to add before we reach the ThisFirstDayOfWeek for the given the year and month
  '		ThisNumberDeltaDay = (7 - ((Weekday(ThisFirstDateOfMonth, ThisFirstDayOfWeek) - 1))) Mod 7
  '		ThisDate = ThisFirstDateOfMonth.AddDays(ThisNumberDeltaDay).AddDays(7 * (NumberOfFirstDayOfWeek - 1))
  '	Case Is < 0
  '		'measure from the last day of month
  '		Dim ThisLastDateOfMonth As Date
  '		ThisLastDateOfMonth = DateSerial(ThisYear, ThisMonth, 1).AddMonths(1).AddDays(-1)
  '		'Number of day to add before we reach the ThisFirstDayOfWeek for the given the next month
  '		ThisNumberDeltaDay = (7 - ((Weekday(ThisLastDateOfMonth, ThisFirstDayOfWeek) - 1))) Mod 7
  '		ThisDate = ThisLastDateOfMonth.AddDays(ThisNumberDeltaDay - 7).AddDays(7 * (-NumberOfFirstDayOfWeek - 1))
  '	End Select
  '	Return ThisDate
  'End Function


  '' ===================================================================
  '' Easter Calculator 0.1
  '' -------------------------------------------------------------------
  ''Easter Date Calculator in Visual Basic.NET
  ''This class calculates the date of Easter using an algorithm that 
  ''was first published in Butcher's Ecclesiastical Calendar in 1876.  
  ''It is valid for all years in the Gregorian Calendar (1583+).  
  ''The algorithm used in this class was adapted from pseudo-code 
  ''in Practical Astronomy for the Calculator, 3rd Edition by 
  ''Peter Duffett-Smith. 
  '' Copyright (c) 2007 David Pinch.
  '' Download from http://www.thoughtproject.com/Snippets/Easter/
  '' ===================================================================
  'Private Function EasterDate(ByVal DateValue As Date) As Date
  '	' ===============================================================
  '	' Easter
  '	' ---------------------------------------------------------------
  '	' Calculates the date of Easter using an algorithm that was first
  '	' published in Butcher's Ecclesiastical Calendar (1876).  It is
  '	' valid for all years in the Gregorian calendar (1583+).  The
  '	' code is based on an implementation by Peter Duffett-Smith in
  '	' Practical Astronomy with your Calculator (3rd Edition).
  '	' ===============================================================

  '	Dim Year As Integer
  '	Static EasterDateLast As Date
  '	Static YearLast As Integer

  '	Year = DateValue.Year
  '	If Year = YearLast Then
  '		Return EasterDateLast
  '	End If
  '	YearLast = Year
  '	If Year < 1583 Then
  '		'this calculation is not valid but e do not want to retunr an error
  '		'Err.Raise(5)
  '		EasterDateLast = DateSerial(Year:=1583, Month:=3, Day:=22)
  '	Else
  '		Dim a As Integer
  '		Dim b As Integer
  '		Dim c As Integer
  '		Dim d As Integer
  '		Dim e As Integer
  '		Dim f As Integer
  '		Dim g As Integer
  '		Dim h As Integer
  '		Dim i As Integer
  '		Dim k As Integer
  '		Dim l As Integer
  '		Dim m As Integer
  '		Dim n As Integer
  '		Dim p As Integer
  '		' Step 1: Divide the year by 19 and store the
  '		' remainder in variable A.  Example: If the year
  '		' is 2000, then A is initialized to 5.

  '		a = Year Mod 19

  '		' Step 2: Divide the year by 100.  Store the integer
  '		' result in B and the remainder in C.

  '		b = Year \ 100
  '		c = Year Mod 100

  '		' Step 3: Divide B (calculated above).  Store the
  '		' integer result in D and the remainder in E.

  '		d = b \ 4
  '		e = b Mod 4

  '		' Step 4: Divide (b+8)/25 and store the integer
  '		' portion of the result in F.

  '		f = (b + 8) \ 25

  '		' Step 5: Divide (b-f+1)/3 and store the integer
  '		' portion of the result in G.

  '		g = (b - f + 1) \ 3

  '		' Step 6: Divide (19a+b-d-g+15)/30 and store the
  '		' remainder of the result in H.

  '		h = (19 * a + b - d - g + 15) Mod 30

  '		' Step 7: Divide C by 4.  Store the integer result
  '		' in I and the remainder in K.

  '		i = c \ 4
  '		k = c Mod 4

  '		' Step 8: Divide (32+2e+2i-h-k) by 7.  Store the
  '		' remainder of the result in L.

  '		l = (32 + 2 * e + 2 * i - h - k) Mod 7

  '		' Step 9: Divide (a + 11h + 22l) by 451 and
  '		' store the integer portion of the result in M.

  '		m = (a + 11 * h + 22 * l) \ 451

  '		' Step 10: Divide (h + l - 7m + 114) by 31.  Store
  '		' the integer portion of the result in N and the
  '		' remainder in P.

  '		n = (h + l - 7 * m + 114) \ 31
  '		p = (h + l - 7 * m + 114) Mod 31

  '		' At this point p+1 is the day on which Easter falls.
  '		' n is 3 for March or 4 for April.

  '		EasterDateLast = DateSerial(Year, n, p + 1)
  '	End If
  '	Return EasterDateLast
  'End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As Report, Optional ByVal IsIgnoreID As Boolean = False) As Boolean

    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
      End If
      If .Name <> Me.Name Then Return False
      If .DateStart <> Me.DateStart Then Return False
      If .DateStop <> Me.DateStop Then Return False
      If Me.Sectors.EqualsDeep(.Sectors, IsIgnoreID) = False Then Return False
      If Me.Industries.EqualsDeep(.Industries, IsIgnoreID) = False Then Return False
      If Me.BondRates.EqualsDeep(.BondRates, IsIgnoreID) = False Then Return False
      If Me.SplitFactorFutures.EqualsDeep(.SplitFactorFutures, IsIgnoreID) = False Then Return False
      If Me.Stocks.EqualsDeep(.Stocks, IsIgnoreID) = False Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As Report) As Boolean Implements IEquatable(Of Report).Equals
    If other Is Nothing Then Return False
    If Me.Name = other.Name Then
      If (Me.DateStart = other.DateStart) And (Me.DateStop = other.DateStop) Then
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If TypeOf obj Is Report Then
      Return Me.Equals(DirectCast(obj, Report))
    Else
      Return False
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return (Me.Name + Me.DateStart.ToString + Me.DateStop.ToString).GetHashCode
  End Function
#End Region
#Region "Error functions"
  ''' <summary>
  ''' If an exception occurs, the exception object will be stored here. If no exception occurs, this property is null/Nothing.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Exception() As Exception
    Get
      Return MyException
    End Get
    Set(Exception As Exception)
      MyException = Exception
      If MyException IsNot Nothing Then
        MyException = MyException
        RaiseEvent Message(MyException.MessageAll, IMessageInfoEvents.enuMessageType.InError)
      End If
    End Set
  End Property
#End Region
#Region "File"
  ''' <summary>
  ''' Load a zip file with all the standard rep file
  ''' </summary>
  ''' <param name="FileZipName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FileLoadZip(ByVal FileZipName As String) As YahooAccessData.Report
    Dim ThisReportRead As YahooAccessData.Report = Nothing
    Dim ThisReport = New YahooAccessData.Report
    Dim I As Integer = 0

    Try
      Using ThisZip As ZipFile = New ZipFile(FileZipName)

        For Each ThisZipEntry In ThisZip.Entries
          'Console.WriteLine(ThisZipEntry.FileName)
          'process only the file with an *.rep extension name
          If UCase(System.IO.Path.GetExtension(ThisZipEntry.FileName)) = ".REP" Then
            I = I + 1
            Dim ThisStream As Stream = New MemoryStream

            ThisZipEntry.Extract(ThisStream)
            ThisStream.Position = 0
            ThisReportRead = Me.FileLoad(ThisStream)
            ThisStream.Dispose()
            If I = 1 Then
              'assign the first report as a reference
              If ThisZip.Entries.Count = 1 Then
                'we do not need to copy 
                'this will be faster 
                ThisReport = ThisReportRead
              Else
                ThisReport = ThisReportRead.CopyDeep
              End If
            Else
              'add all subsequent report
              ThisReport.Add(ThisReportRead)
            End If
            If ThisReport.Exception IsNot Nothing Then
              'thrown the exception again
              Throw New Exception(String.Format("Error adding zip entry: {0}", ThisZipEntry.FileName), Me.Exception)
            End If
          End If
        Next
        'there is a possibility that the symbol in those file terminate with a dot.
        'the error originate from an initial erronous download from eodData that is already on zip format
        'correct for this error here
        Dim ThisSymbolNew As String
        For Each ThisStock In ThisReport.Stocks
          If Right(ThisStock.Symbol, 1) = "." Then
            'change the symbol
            ThisSymbolNew = ThisStock.Symbol.Substring(0, ThisStock.Symbol.Length - 1)
            ThisReport.StockSymbolChange(ThisStock.Symbol, ThisSymbolNew)
          End If
        Next
      End Using
    Catch Ex As Exception
      ThisReport.Exception = New Exception(String.Format("Error while loading zip file: {0}", FileZipName), Ex)
    End Try
    Return ThisReport
  End Function

  Public Sub FileSave()
    'save to the attached file in append mode
    If Me.IsFileOpen Then
      Me.SerializeSaveTo(MyStream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed)
    End If
  End Sub

  ''' <summary>
  ''' This functin can be used to upgrade the current opened file to support the end of day record structure
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Async Function FileSaveConvertToEndOfDay() As Task(Of Exception)
    'Dim ThisQueue As New Queue(Of Task(Of RecordIndex(Of Record, Date)))
    Dim ThisDictionaryOfTask As Dictionary(Of String, Task(Of RecordIndex(Of Record, Date)))
    Dim ThisFilePathRecordNameForEndOfDay As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Me.FileName), "Stock\RecordEndOfDay")
    Dim I As Integer
    Dim ThisCommpletion As Single
    Dim ThisCommpletionLast As Single
    Dim ThisStockCount = Me.Stocks.Count
    Dim ThisTaskWhenAny As Task(Of Task(Of RecordIndex(Of Record, Date)))
    Dim ThisTaskEnd As Task(Of RecordIndex(Of Record, Date))
    Dim ThisStopWatch As New Stopwatch

    Const QUEUE_SIZE As Integer = 16

    Dim ThisTaskOfDirectoryDelete = New Task(Of Exception)(
      Function()
        Try
          My.Computer.FileSystem.DeleteDirectory(ThisFilePathRecordNameForEndOfDay, showUI:=FileIO.UIOption.AllDialogs, onUserCancel:=FileIO.UICancelOption.DoNothing, recycle:=FileIO.RecycleOption.DeletePermanently)
          Return Nothing
        Catch ex As Exception
          Return ex
        End Try
      End Function)

    If Me.IsFileOpen = False Then
      Return New FileNotFoundException
    End If
    If Me.FileType <> IMemoryStream.enuFileType.RecordIndexed Then
      Return New InvalidExpressionException("File operation not supported for this file type...")
    End If
    If My.Computer.FileSystem.DirectoryExists(ThisFilePathRecordNameForEndOfDay) Then
      'delete the directory with all the subfolder
      RaiseEvent Message(String.Format("Start deleting the end of day files directory structure..."), IMessageInfoEvents.enuMessageType.Information)
      ThisTaskOfDirectoryDelete.Start()
      Await ThisTaskOfDirectoryDelete
      If ThisTaskOfDirectoryDelete.Result IsNot Nothing Then
        Return New Exception("Error deleting the end of day files directory structure...", ThisTaskOfDirectoryDelete.Result)
      End If
      RaiseEvent Message(String.Format("Completed the deleting of the end of day files directory structure."), IMessageInfoEvents.enuMessageType.Information)
    End If

    RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.Information)

    ThisDictionaryOfTask = New Dictionary(Of String, Task(Of RecordIndex(Of Record, Date)))
    For Each ThisStock As Stock In Me.Stocks
      Await ThisStock.RecordLoadAsync
      Dim ThisTaskForRecordsEndOfDay = CreateTaskForRecordIndexSave(Of Record, Date)(
        Me.FileName,
        FileMode.Open,
        Me.AsDateRange,
        "Stock\RecordEndOfDay",
        ThisStock.Symbol,
        ".rec",
        ThisStock.Records,
        IsSaveAtEndOfDay:=True)


      ThisDictionaryOfTask.Add(ThisStock.Symbol, ThisTaskForRecordsEndOfDay)

      If ThisDictionaryOfTask.Count = QUEUE_SIZE Then
        ThisTaskWhenAny = Task.WhenAny(ThisDictionaryOfTask.Values.ToList)
        ThisStopWatch.Reset()
        Await ThisTaskWhenAny
        ThisStopWatch.Stop()
        ThisTaskEnd = ThisTaskWhenAny.Result
        'release the memory when it is not needed anymore
        Me.Stocks.ToSearch.Find(ThisTaskEnd.Result.KeyName).RecordReleaseAll()
        ThisDictionaryOfTask.Remove(ThisTaskEnd.Result.KeyName)
        ThisTaskEnd.Dispose()
        ThisTaskEnd = Nothing
        I = I + 1
        ThisCommpletion = CSng(100 * (I / ThisStockCount))
        If (ThisCommpletion - ThisCommpletionLast) > 0.1 Then
          'send a message
          ThisCommpletionLast = ThisCommpletion
          RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.InformationUpdate)
        End If
      End If
    Next
    Do Until ThisDictionaryOfTask.Count = 0
      ThisTaskWhenAny = Task.WhenAny(ThisDictionaryOfTask.Values.ToList)
      Await ThisTaskWhenAny
      ThisTaskEnd = ThisTaskWhenAny.Result
      'release the memory when it is not needed anymore
      Me.Stocks.ToSearch.Find(ThisTaskEnd.Result.KeyName).RecordReleaseAll()
      ThisDictionaryOfTask.Remove(ThisTaskEnd.Result.KeyName)
      ThisTaskEnd.Dispose()
      ThisTaskEnd = Nothing
      I = I + 1
      ThisCommpletion = CSng(100 * (I / ThisStockCount))
      'send a message
      ThisCommpletionLast = ThisCommpletion
      RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.InformationUpdate)
    Loop
    Return Nothing
  End Function

  Public Function FileSaveZip(ByVal FileName As String) As Exception
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisStream As Stream = Nothing
    Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
    Try
      'create the directory if it does not exist
      With My.Computer.FileSystem
        If .DirectoryExists(ThisPath) = False Then
          Try
            .CreateDirectory(ThisPath)
          Catch ex As Exception
            Return New Exception(String.Format("Unable to create directory {0},ThisPath"), ex)
          End Try
        End If
      End With
      'make sure any previous file is deleted so that the creation date is updated
      Try
        System.IO.File.Delete(FileName)
      Catch DeleteError As IOException
        Return New Exception("Unable to save file...", DeleteError)
      End Try
      Try
        Using ThisZip As Ionic.Zip.ZipFile = New ZipFile
          'compress the serialization stream
          ThisStream = New MemoryStream
          Me.SerializeSaveTo(ThisStream)
          ThisStream.Seek(0, SeekOrigin.Begin)
          ThisZip.AddEntry(System.IO.Path.GetFileNameWithoutExtension(FileName) & ".rep", ThisStream)
          ThisZip.Save(FileName)
        End Using
      Catch Ex As Exception
        Return New Exception(String.Format("Error while packaging zip file: {0}", FileName), Ex)
      End Try
    Catch e As Exception
      If ThisStream IsNot Nothing Then
        ThisStream.Dispose()
      End If
      Return New Exception("Serialization writing error...", e)
    End Try
    Return Nothing
  End Function

  Public Async Function FileSaveZipAsync(ByVal FileName As String) As Task(Of Exception)
    Dim ThisTask As Task(Of Exception)

    ThisTask = New Task(Of Exception)(
      Function()
        Return Me.FileSaveZip(FileName)
      End Function)

    ThisTask.Start()
    Await ThisTask
    Return ThisTask.Result
  End Function

  Public Function FileSave(ByVal FileName As String) As Boolean
    Return Me.FileSave(Me, FileName, IMemoryStream.enuFileType.Standard)
  End Function

  Public Function FileSave(ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType) As Boolean
    Return Me.FileSave(Me, FileName, FileType)
  End Function

  Public Async Function FileSaveAsync(Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)
    'save to the attached file in append mode
    If Me.IsFileOpen Then
      'this is much slower
      'Dim ThisTask As New Task(Of Boolean)(
      '  Function()
      '    Me.SerializeSaveTo(MyStream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed)
      '    Return True
      '  End Function)
      'ThisTask.Start()
      'Await ThisTask
      'This is very fast
      Await SerializeSaveToAsync(MyStream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed, IsSingleThread)
      Return True
    Else
      Return False
    End If
  End Function

  Public Async Function FileSaveAsync(ByVal FileName As String, Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)
    Await Me.FileSaveAsync(Me, FileName, IMemoryStream.enuFileType.RecordIndexed, IsSingleThread)
    Return True
  End Function

  Public Async Function FileSaveAsync(ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType, Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)
    Await Me.FileSaveAsync(Me, FileName, FileType, IsSingleThread)
    Return True
  End Function

  Public Async Function FileSaveAsync(ByVal Report As Report, ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType, Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisStream As Stream = Nothing
    Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
    Try
      'create the directory if it does not exist
      With My.Computer.FileSystem
        If .DirectoryExists(ThisPath) = False Then
          Try
            .CreateDirectory(ThisPath)
          Catch ex As Exception
            Throw New Exception(String.Format("Unable to create directory {0},ThisPath"), ex)
          End Try
        End If
      End With
      Report.Exception = Nothing
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          'make sure any previous file is deleted so that the creation date is updated
          Try
            System.IO.File.Delete(FileName)
          Catch DeleteError As IOException
            Report.Exception = New Exception("Unable to save file...", DeleteError)
            Return False
          End Try
          ThisStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None)
          Report.SerializeSaveTo(ThisStream, FileType)
          ThisStream.Dispose()
          ThisStream = Nothing
          Return True
        Case IMemoryStream.enuFileType.RecordIndexed
          'make sure any previous file is deleted so that the creation date is updated
          Try
            System.IO.File.Delete(FileName)
            System.IO.File.Delete(FileName & ".xml")
            With My.Computer.FileSystem
              If .DirectoryExists(ThisPath & "\Stock") Then
                .DeleteDirectory(ThisPath & "\Stock", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
              If .DirectoryExists(ThisPath & "\SplitFactorFuture") Then
                .DeleteDirectory(ThisPath & "\SplitFactorFuture", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
              If .DirectoryExists(ThisPath & "\BondRate") Then
                .DeleteDirectory(ThisPath & "\BondRate", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
            End With
          Catch DeleteError As IOException
            Report.Exception = New Exception("Unable to save file...", DeleteError)
            Return False
          End Try
          ThisStream = New FileStream(FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None)
          Await Report.SerializeSaveToAsync(ThisStream, FileType, IsSingleThread)
          ThisStream.Dispose()
          Return True
      End Select
    Catch e As Exception
      Me.Exception = New Exception("Serialization writing error...", e)
      If ThisStream IsNot Nothing Then
        ThisStream.Dispose()
      End If
      Return False
    End Try
    Return True
  End Function

  Public Function FileSave(ByRef Report As Report, ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType) As Boolean
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisStream As Stream = Nothing
    Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
    Dim ThisEx As Exception
    Try
      'create the directory if it does not exist
      With My.Computer.FileSystem
        If .DirectoryExists(ThisPath) = False Then
          Try
            .CreateDirectory(ThisPath)
          Catch ex As Exception
            Throw New Exception(String.Format("Unable to create directory {0}", ThisPath), ex)
          End Try
        End If
      End With
      Report.Exception = Nothing
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          'make sure any previous file is deleted so that the creation date is updated
          Try
            System.IO.File.Delete(FileName)
          Catch DeleteError As IOException
            Report.Exception = New Exception("Unable to save file...", DeleteError)
            Return False
          End Try
          ThisStream = New FileStream(FileName, FileMode.Create, FileAccess.Write, FileShare.None)
          Report.SerializeSaveTo(ThisStream, FileType)
          ThisStream.Dispose()
          ThisStream = Nothing
          Return True
        Case IMemoryStream.enuFileType.RecordIndexed
          'make sure any previous file is deleted so that the creation date is updated
          Try
            System.IO.File.Delete(FileName)
            System.IO.File.Delete(FileName & ".xml")
            With My.Computer.FileSystem
              If .DirectoryExists(ThisPath & "\Stock") Then
                .DeleteDirectory(ThisPath & "\Stock", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
              If .DirectoryExists(ThisPath & "\SplitFactorFuture") Then
                .DeleteDirectory(ThisPath & "\SplitFactorFuture", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
              If .DirectoryExists(ThisPath & "\BondRate") Then
                .DeleteDirectory(ThisPath & "\BondRate", FileIO.DeleteDirectoryOption.DeleteAllContents)
              End If
            End With
          Catch DeleteError As IOException
            Report.Exception = New Exception("Unable to save file...", DeleteError)
            Return False
          End Try
          ThisStream = New FileStream(FileName, FileMode.Create, FileAccess.ReadWrite, FileShare.None)
          Report.SerializeSaveTo(ThisStream, FileType)
          ThisStream.Dispose()
          Return True
      End Select
    Catch e As Exception
      ThisEx = e
      Me.Exception = New Exception("Serialization writing error...", ThisEx)
      If ThisStream IsNot Nothing Then
        ThisStream.Dispose()
      End If
      Return False
    End Try
    Return True
  End Function

  ''' <summary>
  ''' indicate if the file has been successfully opened
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property IsFileOpen As Boolean
    Get
      Return _IsFileOpen
    End Get
    Friend Set(value As Boolean)
      _IsFileOpen = value
      Me.IsFileReadEndOfDayEnabled = False
    End Set
  End Property

  ''' <summary>
  ''' Return the current active file name
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property FileName As String
    Get
      Return MyFileName
    End Get
  End Property

  Public Async Function FileOpenAsync(ByVal FileName As String, Optional IsReadOnly As Boolean = False) As Task(Of Boolean)
    Dim ThisBinaryReader As BinaryReader = Nothing
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisHeader As String

    If Me.IsFileOpen Then
      Me.FileClose()
    End If
    Try
      IsFileAccessReadOnly = IsReadOnly
      If IsReadOnly Then
        MyStream = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
      Else
        MyStream = New FileStream(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.Read)
      End If
      ThisBinaryReader = New BinaryReader(MyStream, New System.Text.UTF8Encoding(), leaveOpen:=True)
      MyFileName = FileName
      Me.IsFileOpen = True
      ThisHeader = ThisBinaryReader.ReadString
      If ThisHeader = "Report" Then
        MyFileType = IMemoryStream.enuFileType.Standard    'by default
        Call SerializeLoadFromMainLocal(ThisBinaryReader, True)
        Await SerializeLoadFromStockLocalAsync(ThisBinaryReader, True)
      Else
        'old serialization net format
        ThisBinaryReader.Dispose()
        Throw New Exception("Invalid file format...")
      End If
      ThisBinaryReader.Dispose()
    Catch e As Exception
      If Me.IsFileOpen Then
        Me.FileClose()
      End If
      If ThisBinaryReader IsNot Nothing Then
        ThisBinaryReader.Dispose()
      End If
      Throw New Exception(String.Format("{0} Serialization reading error...", FileName), e)
    End Try
    If Me.FileType = IMemoryStream.enuFileType.Standard Then
      Me.FileClose()
      Throw New FileLoadException("FileOpen is only supported for Record indexed file type...")
    End If
    Return Me.IsFileOpen
  End Function

  Public Async Function FileOpenAsync(ByVal FileName As String, ByVal DateStart As Date, ByVal DateStop As Date, Optional IsReadOnly As Boolean = False) As Task(Of Boolean)
    Me.FileClose()
    Me.IDateRange_DateStart = DateStart
    Me.IDateRange_DateStop = DateStop
    Await Me.FileOpenAsync(FileName)
    Return Me.IsFileOpen
  End Function

  ''' <summary>
  ''' Open and attach a file to a Report
  ''' </summary>
  ''' <param name="FileName"></param>
  ''' <remarks></remarks>
  Public Function FileOpen(ByVal FileName As String, Optional IsReadOnly As Boolean = False) As Boolean
    Dim ThisStream As Stream = Nothing
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisBinaryReader As BinaryReader = Nothing
    Dim ThisHeader As String

    If Me.IsFileOpen Then
      Me.FileClose()
    End If
    Try
      IsFileAccessReadOnly = IsReadOnly
      If IsReadOnly Then
        ThisStream = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
      Else
        ThisStream = New FileStream(FileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
      End If
      System.IO.File.SetAttributes(FileName, IO.FileAttributes.Normal)
      ThisBinaryReader = New BinaryReader(ThisStream, New System.Text.UTF8Encoding(), leaveOpen:=True)
      ThisHeader = ThisBinaryReader.ReadString
      ThisBinaryReader.Dispose()
      ThisStream.Position = 0
      If ThisHeader = "Report" Then
        Me.SerializeLoadFrom(ThisStream, True)
      Else
        'old serialization net format
        Throw New Exception("Invalid file format...")
      End If
      ThisStream = Nothing
    Catch e As Exception
      If ThisBinaryReader IsNot Nothing Then
        ThisBinaryReader.Dispose()
      End If
      If ThisStream IsNot Nothing Then
        ThisStream.Dispose()
      End If
      Me.FileClose()
      Me.Exception = New FileLoadException(String.Format("{0} file open error...:{1}{2}", FileName, vbCr, e.Message), e)
      Return False
    End Try
    If Me.FileType = IMemoryStream.enuFileType.Standard Then
      'Me.FileClose()
      'Me.Exception = New FileLoadException("FileOpen is only supported for Record indexed file type...")
      'Return False
    End If
    MyFileName = FileName
    Me.IsFileOpen = True
    Return Me.IsFileOpen
  End Function

  Public Function FileOpen(ByVal FileName As String, ByVal DateStart As Date, ByVal DateStop As Date, Optional IsReadOnly As Boolean = False) As Boolean
    Me.FileClose()
    Me.IDateRange_DateStart = DateStart
    Me.IDateRange_DateStop = DateStop
    Me.FileOpen(FileName, IsReadOnly)
    Return Me.IsFileOpen
  End Function

  Public Sub FileClose()
    'close the file
    If MyStream IsNot Nothing Then
      MyStream.Dispose()
    End If
    Me.Clear()
    MyFileName = ""
    Me.IsFileOpen = False
    IsFileAccessReadOnly = False
  End Sub

  Public Function FileLoad(ByVal FileName As String, ByVal IsRecordVirtual As Boolean, ByVal DateStart As Date, ByVal DateStop As Date) As YahooAccessData.Report
    Me.IDateRange_DateStart = DateStart
    Me.IDateRange_DateStop = DateStop
    Return FileLoad(FileName, IsRecordVirtual)
  End Function

  Public Function FileLoad(ByVal FileName As String, ByVal IsRecordVirtual As Boolean) As YahooAccessData.Report
    Dim ThisReport As YahooAccessData.Report = Nothing
    Dim ThisStream As Stream = Nothing
    Try
      If My.Computer.FileSystem.FileExists(FileName) Then
        ThisStream = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
        ThisReport = Me.FileLoad(ThisStream, IsRecordVirtual)
        If IsRecordVirtual = False Then
          'close the file immediately
          ThisStream.Dispose()
        End If
        ThisStream = Nothing
      Else
        ThisReport = New YahooAccessData.Report
        ThisReport.Exception = New Exception(String.Format("File not found: {0}", FileName))
      End If
    Catch e As Exception
      ThisReport = New YahooAccessData.Report
      ThisReport.Exception = New Exception(String.Format("{0} Serialization reading error...", FileName), e)
      If ThisStream IsNot Nothing Then
        ThisStream.Dispose()
      End If
    End Try
    Return ThisReport
  End Function

  Public Function FileLoad(ByVal ThisFileName As String) As YahooAccessData.Report
    Return FileLoad(ThisFileName, False)
  End Function

  Public Function FileLoad(ByRef Stream As Stream, ByVal IsRecordVirtual As Boolean, ByVal DateStart As Date, ByVal DateStop As Date) As YahooAccessData.Report
    Me.IDateRange_DateStart = DateStart
    Me.IDateRange_DateStop = DateStop
    Return FileLoad(Stream, IsRecordVirtual)
  End Function

  Public Function FileLoad(ByRef Stream As Stream, ByVal IsRecordVirtual As Boolean) As YahooAccessData.Report
    Dim ThisReportRead As YahooAccessData.Report = Nothing
    Dim ThisReport As YahooAccessData.Report = Nothing
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisBinaryReader As BinaryReader = Nothing
    Dim ThisHeader As String


    Try
      Stream.Seek(0, SeekOrigin.Begin)
      ThisBinaryReader = New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
      ThisHeader = ThisBinaryReader.ReadString
      ThisBinaryReader.Dispose()
      Stream.Position = 0
      If ThisHeader = "Report" Then
        'new version
        ThisReport = New YahooAccessData.Report
        ThisReport.SerializeLoadFrom(Stream, IsRecordVirtual)
      Else
        'old serialization net format
        ThisReportRead = CType(ThisFormatter.Deserialize(Stream), Report)
        'need to validate with the old file version
        ThisReport = Me.ReportValidate(ThisReportRead)
      End If
    Catch e As Exception
      If ThisBinaryReader IsNot Nothing Then
        ThisBinaryReader.Dispose()
      End If
      ThisReport = New YahooAccessData.Report
      ThisReport.Exception = New Exception(String.Format("Serialization streaming reading error..."), e)
    End Try
    ThisBinaryReader = Nothing
    Return ThisReport
  End Function
  Public Function FileLoad(ByRef Stream As Stream) As YahooAccessData.Report
    Return FileLoad(Stream, False)
  End Function

  Private Function ReportValidate(ByVal Report As YahooAccessData.Report) As YahooAccessData.Report
    Dim ThisReport As YahooAccessData.Report
    Dim ThisSplitFactorFuture As YahooAccessData.SplitFactorFuture
    Dim ThisStock As YahooAccessData.Stock
    Dim ThisStockSymbol As YahooAccessData.StockSymbol

    If Report.Stocks.Count > 0 Then
      If Report.Stocks.ElementAt(0).ReportID = 0 Then
        'older file do not have their ID and link saved correctly
        For Each ThisStock In Report.Stocks
          With ThisStock
            'older file may not have their ID saved correctly
            'but the object stock contain all the link information
            .ReportID = Report.ID
            .SectorID = .Sector.ID
            .IndustryID = .Industry.ID
            .Sector.ReportID = Report.ID
            .Industry.ReportID = Report.ID
            .Industry.SectorID = .Sector.ID
            .Industry.Sector = .Sector
            If .StockErrors Is Nothing Then
              'this will happen if we have loaded an old version (<= V2.0.0)
              .StockErrors = New LinkedHashSet(Of StockError, Date)
            End If
            If .StockSymbols Is Nothing Then
              'this will happen if we have loaded an old version (<= V2.0.0)
              .StockSymbols = New LinkedHashSet(Of StockSymbol, Date)
            Else
              For Each ThisStockSymbol In .StockSymbols
                ThisStockSymbol.Exchange = ThisStockSymbol.Stock.Exchange
              Next
            End If
          End With
        Next
      End If
    End If
    If Report.SplitFactorFutures Is Nothing Then
      'this will happen if we have loaded an old version (<= V1.7.9) where the
      'the SplitFactorFutures did not exist
      'in that case restore the SplitFactorFutures collection with a count of zero
      'that make this current version compatible with the older version
      Report.SplitFactorFutures = New LinkedHashSet(Of SplitFactorFuture, String)
    Else
      For Each ThisSplitFactorFuture In Report.SplitFactorFutures
        With ThisSplitFactorFuture
          If .DateUpdate.HasValue = False Then
            'this could happen for older file
            'assign the report start value by default
            .DateUpdate = Report.DateStart
          End If
        End With
      Next
    End If
    If Report.BondRates Is Nothing Then
      'this will happen if we have loaded an old version (<= V1.8.0) where the
      'the BondRates did not exist
      'in that case restore the SplitFactorFutures collection with a count of zero
      'that make this current version compatible with the older version
      Report.BondRates = New LinkedHashSet(Of BondRate, String)
    End If
    'report may be an older version 
    If TypeOf (Report.Stocks) Is ISearchKey(Of Stock, String) Then
      'the report version is current
      ThisReport = Report
    Else
      'This is an older report version that did not have yet the ISearchKey interface
      'copydeep will update the report to the latest report version using linked hashset
      ThisReport = Report.CopyDeepLocal
    End If
    Return ThisReport
  End Function

  Public ReadOnly Property FileVersion As String
    Get
      Try
        Dim ThisAssemblyName = System.Reflection.AssemblyName.GetAssemblyName("YahooAccessData.dll")
        Return ThisAssemblyName.Version.ToString
      Catch Ex As Exception
        Return "unknown"
      End Try
    End Get
  End Property
#End Region
#Region "IComparable"
  ''' <summary>
  ''' </summary>
  ''' <param name="other"></param>
  ''' <returns>
  ''' Less than zero: This object is less than the other parameter. 
  ''' Zero : This object is equal to other. 
  ''' Greater than zero : This object is greater than other. 
  ''' </returns>
  ''' <remarks></remarks>
  Public Function CompareTo(other As Report) As Integer Implements System.IComparable(Of Report).CompareTo
    Return Me.KeyValue.CompareTo(other.KeyValue)
  End Function
#End Region
#Region "Register Key"
  Public Property KeyID As Integer Implements IRegisterKey(Of Date).KeyID
    Get
      Return Me.ID
    End Get
    Set(value As Integer)
      Me.ID = value
    End Set
  End Property

  Public Property KeyValue As Date Implements IRegisterKey(Of Date).KeyValue
    Get
      Return Me.DateStop
    End Get
    Set(value As Date)

    End Set
  End Property
#End Region
#Region "IMemoryStream"
  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
    Me.SerializeLoadFrom(Stream, False)
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As BinaryReader = Nothing
    Dim ThisHeader As String

    MyStream = Stream

    ThisBinaryReader = New BinaryReader(MyStream, New System.Text.UTF8Encoding(), leaveOpen:=True)
    MyStream.Position = 0
    MyFileType = IMemoryStream.enuFileType.Standard    'by default
    ThisHeader = ThisBinaryReader.ReadString
    If ThisHeader = "Report" Then
      Call SerializeLoadFromMainLocal(ThisBinaryReader, IsRecordVirtual)
      Call SerializeLoadFromStockLocal(ThisBinaryReader, IsRecordVirtual)
    Else
      'old serialization net format
      Dim ThisReportRead As Report
      Dim ThisReport As Report
      Dim ThisFormatter = New BinaryFormatter
      MyStream.Seek(0, SeekOrigin.Begin)
      'old serialization net format
      ThisReportRead = CType(ThisFormatter.Deserialize(Stream), Report)
      'need to validate with the old file version
      ThisReport = Me.ReportValidate(ThisReportRead)
      Me.Add(ThisReport)
    End If
    ThisBinaryReader.Dispose()
  End Sub

  Private Sub SerializeLoadFromMainLocal(ByVal BinaryReader As BinaryReader, ByRef IsRecordVirtual As Boolean)
    Dim ThisFileVersion As String
    Dim ThisVersion As Single
    Dim I As Integer
    Dim Count As Integer
    Dim ThisMaxID As Integer
    Dim ThisDateStart As Date
    Dim ThisDateStop As Date

    With BinaryReader
      ThisFileVersion = .ReadString
      ThisVersion = .ReadSingle
      MyFileType = CType(.ReadInt32, IMemoryStream.enuFileType)
      If MyFileType = IMemoryStream.enuFileType.Standard Then
        'Virtual record is not supported for the standard file
        IsRecordVirtual = False
      Else
        IsRecordVirtual = True
        Me.BondRates = New LinkedHashSet(Of BondRate, String)(Me)
        Me.SplitFactorFutures = New LinkedHashSet(Of SplitFactorFuture, String)(Me)
      End If
      Me.ID = .ReadInt32
      Me.Exception = Nothing
      For I = 1 To .ReadInt32
        Me.Exception = New Exception(.ReadString, Me.Exception)
      Next
      Me.Name = .ReadString

      Me.DateStart = DateTime.FromBinary(.ReadInt64)
      Me.DateStop = DateTime.FromBinary(.ReadInt64)

      Count = .ReadInt32
      Select Case Me.FileType
        Case IMemoryStream.enuFileType.Standard
          If Count > 4 Then
            CType(_BondRates, LinkedHashSet(Of BondRate, String)).Capacity = Count
          End If
          For I = 1 To Count
            Dim ThisBondRate As New BondRate(Me, MyStream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          ThisMaxID = .ReadInt32
          ThisDateStart = DateTime.FromBinary(.ReadInt64)
          ThisDateStop = DateTime.FromBinary(.ReadInt64)
          With CType(_BondRates, IRecordInfo)
            .CountTotal = Count
            .MaximumID = ThisMaxID
          End With
          With DirectCast(_BondRates, IDateUpdate)
            .DateStart = ThisDateStart
            .DateStop = ThisDateStop
          End With
      End Select
      Count = .ReadInt32
      Select Case Me.FileType
        Case IMemoryStream.enuFileType.Standard
          If Count > 4 Then
            CType(_SplitFactorFutures, LinkedHashSet(Of SplitFactorFuture, String)).Capacity = Count
          End If
          For I = 1 To Count
            Dim ThisSplitFactorFuture As New SplitFactorFuture(Me, MyStream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          ThisMaxID = .ReadInt32
          ThisDateStart = DateTime.FromBinary(.ReadInt64)
          ThisDateStop = DateTime.FromBinary(.ReadInt64)
          With CType(_SplitFactorFutures, IRecordInfo)
            .CountTotal = Count
            .MaximumID = ThisMaxID
          End With
          With DirectCast(_SplitFactorFutures, IDateUpdate)
            .DateStart = ThisDateStart
            .DateStop = ThisDateStop
          End With
      End Select
      Count = .ReadInt32
      If Count > 4 Then
        CType(Me.Sectors, LinkedHashSet(Of Sector, String)).Capacity = Count
      End If
      For I = 1 To Count
        Dim ThisSector As New Sector(Me, MyStream)
      Next
      Count = .ReadInt32
      If Count > 4 Then
        CType(Me.Industries, LinkedHashSet(Of Industry, String)).Capacity = Count
      End If
      For I = 1 To Count
        Dim ThisIndustry As New Industry(Me, MyStream)
      Next
    End With
  End Sub

  Private Sub SerializeLoadFromStockLocal(ByRef BinaryReader As BinaryReader, IsRecordVirtual As Boolean)
    Dim I As Integer
    Dim Count As Integer

    With BinaryReader
      Count = .ReadInt32
      If Count > 4 Then
        CType(Me.Stocks, LinkedHashSet(Of Stock, String)).Capacity = Count
      End If
      For I = 1 To Count
        Dim ThisStock As New Stock(Me, MyStream, IsRecordVirtual)
      Next
    End With
  End Sub

  Private Async Function SerializeLoadFromStockLocalAsync(ByVal BinaryReader As BinaryReader, ByVal IsRecordVirtual As Boolean) As Task(Of Boolean)
    Dim I As Integer
    Dim Count As Integer
    Dim ThisStopWatch As New Stopwatch
    Dim ThisProgress As New Progress(Of Integer)
    Dim ThisProgressCaptureEvent =
      Sub(Sender As Object, value As Integer)
        RaiseEvent Message(String.Format("Saving stock list {0:P1}...", I / Count), IMessageInfoEvents.enuMessageType.Information)
        RaiseEvent Message(value.ToString, IMessageInfoEvents.enuMessageType.Information)
      End Sub

    RaiseEvent Message(String.Format("Saving stock list..."), IMessageInfoEvents.enuMessageType.Information)
    AddHandler ThisProgress.ProgressChanged, ThisProgressCaptureEvent

    'main code
    Count = BinaryReader.ReadInt32()
    If Count > 4 Then
      'CType(Me.Stocks, LinkedHashSet(Of Stock, String)).Capacity = Count
    End If
    Dim ThisTask As New Task(Of Boolean)(
      Function()
        Dim Progress As IProgress(Of Integer) = ThisProgress
        With BinaryReader
          ThisStopWatch.Restart()
          For I = 1 To Count
            Dim ThisStock As New Stock(Me, MyStream, IsRecordVirtual)
            If ThisStopWatch.ElapsedMilliseconds > 500 Then
              ThisStopWatch.Restart()
              Progress.Report(I)
            End If
          Next
        End With
        Return True
      End Function)
    ThisTask.Start()
    Await ThisTask
    RemoveHandler ThisProgress.ProgressChanged, ThisProgressCaptureEvent
    Return True
  End Function

  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
    Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
  End Sub

  Private Async Function SerializeSaveToAsync(ByVal Stream As System.IO.Stream, ByVal FileType As IMemoryStream.enuFileType, Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)
    Dim ThisListException As List(Of Exception)
    Dim ThisException As Exception
    Dim ThisZero As Integer = 0
    Dim ThisTaskForBondRates As Task(Of RecordIndex(Of BondRate, String)) = Nothing
    Dim ThisTaskForSplitFactorFutures As Task(Of RecordIndex(Of SplitFactorFuture, String)) = Nothing
    Dim ThisStream As Stream
    'store the count because saving on a thread may clear the value before we have time to used it
    Dim ThisBondRateCount As Integer = 0
    Dim ThisSplitFactorFutureCount As Integer = 0
    Dim ThisQueueSize As Integer

    Const PARALLEL_STOCK_SAVE_DEFAULT As Integer = 8

    If IsSingleThread Then
      ThisQueueSize = 1
    Else
      ThisQueueSize = PARALLEL_STOCK_SAVE_DEFAULT
    End If

    ThisStream = Stream

    Dim ThisBinaryWriter As New BinaryWriter(ThisStream)

    If (FileType = IMemoryStream.enuFileType.RecordIndexed) Then
      If TypeOf _BondRates Is IDataVirtual Then
        'make sure the data has been loaded before we save
        With DirectCast(_BondRates, IDataVirtual)
          Dim IsLoaded As Boolean = .IsLoaded
          If IsLoaded = False Then .Load()
        End With
      End If
      If _BondRates.Count > 0 Then
        ThisBondRateCount = _BondRates.Count
        'start the process of saving the BondRate
        ThisTaskForBondRates = CreateTaskForRecordIndexSave(Of BondRate, String)(
          DirectCast(ThisStream, FileStream).Name(),
          FileMode.Open,
          Me.AsDateRange,
          "BondRate",
          "",
          ".bra",
          _BondRates)

        If IsSingleThread Then
          Await ThisTaskForBondRates
        End If
      End If
      If TypeOf _SplitFactorFutures Is IDataVirtual Then
        'make sure the data has been loaded before we save
        With DirectCast(_SplitFactorFutures, IDataVirtual)
          Dim IsLoaded As Boolean = .IsLoaded
          If IsLoaded = False Then .Load()
        End With
      End If
      If _SplitFactorFutures.Count > 0 Then
        ThisSplitFactorFutureCount = _SplitFactorFutures.Count
        'start the process of saving the SplitFactorFuture
        ThisTaskForSplitFactorFutures = CreateTaskForRecordIndexSave(Of SplitFactorFuture, String)(
          DirectCast(ThisStream, FileStream).Name(),
          FileMode.Open,
          Me.AsDateRange,
          "SplitFactorFuture",
          "",
          ".sff",
          _SplitFactorFutures)

        If IsSingleThread Then
          Await ThisTaskForSplitFactorFutures
        End If
      End If
    End If
    With ThisBinaryWriter
      .Seek(0, SeekOrigin.Begin)
      .Write("Report")  'File identification
      .Write(Me.FileVersion)
      .Write(VERSION_MEMORY_STREAM)
      .Write(FileType)
      .Write(Me.ID)
      ThisListException = Me.Exception.ToList
      ThisListException.Reverse()
      .Write(ThisListException.Count)
      For Each ThisException In ThisListException
        .Write(ThisException.Message)
      Next
      .Write(Me.Name)
      .Write(Me.DateStart.ToBinary)
      .Write(Me.DateStop.ToBinary)
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          .Write(Me.BondRates.Count)
          For Each ThisBondRate In Me.BondRates
            ThisBondRate.SerializeSaveTo(ThisStream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          If ThisBondRateCount > 0 Then
            Dim ThisRecordIndex As RecordIndex(Of BondRate, String)
            Await ThisTaskForBondRates
            ThisRecordIndex = ThisTaskForBondRates.Result
            .Write(ThisRecordIndex.FileCount)
            .Write(ThisRecordIndex.MaxID)
            .Write(ThisRecordIndex.DateStart.ToBinary)
            .Write(ThisRecordIndex.DateStop.ToBinary)
            With DirectCast(_BondRates, IDateUpdate)
              .DateStart = ThisRecordIndex.DateStart
              .DateStop = ThisRecordIndex.DateStop
            End With
            With DirectCast(_BondRates, IRecordInfo)
              .CountTotal = ThisRecordIndex.FileCount
              .MaximumID = ThisRecordIndex.MaxID
            End With
            ThisRecordIndex.Dispose()
            ThisRecordIndex = Nothing
            ThisTaskForBondRates = Nothing
          Else
            Dim ThisRecordInfo = DirectCast(_BondRates, IRecordInfo)
            Dim ThisDateUpdate = DirectCast(_BondRates, IDateUpdate)
            .Write(ThisRecordInfo.CountTotal)
            .Write(ThisRecordInfo.MaximumID)
            .Write(ThisDateUpdate.DateStart.ToBinary)
            .Write(ThisDateUpdate.DateStop.ToBinary)
          End If
      End Select
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          .Write(Me.SplitFactorFutures.Count)
          For Each ThisSplitFactorFuture In Me.SplitFactorFutures
            ThisSplitFactorFuture.SerializeSaveTo(ThisStream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          If ThisSplitFactorFutureCount > 0 Then
            Dim ThisRecordIndex As RecordIndex(Of SplitFactorFuture, String)
            Await ThisTaskForSplitFactorFutures
            ThisRecordIndex = ThisTaskForSplitFactorFutures.Result
            .Write(ThisRecordIndex.FileCount)
            .Write(ThisRecordIndex.MaxID)
            .Write(ThisRecordIndex.DateStart.ToBinary)
            .Write(ThisRecordIndex.DateStop.ToBinary)
            With DirectCast(_SplitFactorFutures, IDateUpdate)
              .DateStart = ThisRecordIndex.DateStart
              .DateStop = ThisRecordIndex.DateStop
            End With
            With DirectCast(_SplitFactorFutures, IRecordInfo)
              .CountTotal = ThisRecordIndex.FileCount
              .MaximumID = ThisRecordIndex.MaxID
            End With
            ThisRecordIndex.Dispose()
            ThisRecordIndex = Nothing
            ThisTaskForSplitFactorFutures = Nothing
          Else
            Dim ThisRecordInfo = DirectCast(_SplitFactorFutures, IRecordInfo)
            Dim ThisDateUpdate = DirectCast(_SplitFactorFutures, IDateUpdate)
            .Write(ThisRecordInfo.CountTotal)
            .Write(ThisRecordInfo.MaximumID)
            .Write(ThisDateUpdate.DateStart.ToBinary)
            .Write(ThisDateUpdate.DateStop.ToBinary)
          End If
      End Select
      .Write(Me.Sectors.Count)
      For Each ThisSector In Me.Sectors
        ThisSector.SerializeSaveTo(ThisStream)
      Next
      .Write(Me.Industries.Count)
      For Each ThisIndustry In Me.Industries
        ThisIndustry.SerializeSaveTo(ThisStream)
      Next
      Dim Count As Integer = Me.Stocks.Count
      .Write(Count)

      'Dim I As Integer
      'Dim ThisStopWatch As New Stopwatch
      'ThisStopWatch.Restart()
      'RaiseEvent Message(String.Format("Saving stock list..."), IMessageInfoEvents.enuMessageType.Information)

      Dim ThisQueue As New Queue(Of Task(Of MemoryStreamWatch))
      Dim ThisStreamFileName As String = DirectCast(ThisStream, FileStream).Name()
      Dim I As Integer
      Dim ThisStockCount = Me.Stocks.Count
      Dim ThisListOfWatch As New List(Of YahooAccessData.IProcessTimeMeasurement)
      'Dim ThisStopWatch As New Stopwatch
      Dim ThisTaskForStock As New Task(Of Boolean)(
        Function()
          Dim ThisTaskEnd As Task(Of MemoryStreamWatch)
          Dim ThisStopWatch As New Stopwatch

          For Each ThisStock As Stock In Me.Stocks
            'ThisStock.SerializeSaveTo(ThisStream, FileType)
            'Console.WriteLine(String.Format("Processing {0}", I))
            Dim ThisTask As Task(Of MemoryStreamWatch)
            ThisTask = ThisStock.FileSaveWatchAsync(ThisStreamFileName, IsSingleThread)
            ThisQueue.Enqueue(ThisTask)
            If ThisQueue.Count = ThisQueueSize Then
              ThisTaskEnd = ThisQueue.Dequeue
              ThisTaskEnd.Wait()
              ThisStopWatch.Restart()
              ThisTaskEnd.Result.WriteTo(ThisStream)
              ThisStopWatch.Stop()
              For Each ProcessWatch In ThisTaskEnd.Result.ToListOfProcessWatch
                ThisListOfWatch.Add(ProcessWatch)
              Next
              ThisListOfWatch.Add(New YahooAccessData.ProcessTimeMeasurement(ThisTaskEnd.Result.KeyValue, "Writing at the end of queue", ThisStopWatch.ElapsedMilliseconds))
              ThisTaskEnd = Nothing
              'ThisStopWatch.Restart()
              I = I + 1
            End If
          Next
          'Console.WriteLine(String.Format("Processing final stage"))
          'finish with the last set of task
          For Each ThisTaskEnd In ThisQueue
            'Console.WriteLine(String.Format("Processing {0}", I))
            ThisTaskEnd.Wait()
            ThisStopWatch.Restart()
            ThisTaskEnd.Result.WriteTo(ThisStream)
            ThisStopWatch.Stop()
            For Each ProcessWatch In ThisTaskEnd.Result.ToListOfProcessWatch
              ThisListOfWatch.Add(ProcessWatch)
            Next
            ThisListOfWatch.Add(New YahooAccessData.ProcessTimeMeasurement(ThisTaskEnd.Result.KeyValue, "Writing at the end of queue", ThisStopWatch.ElapsedMilliseconds))
            ThisTaskEnd = Nothing
            'I = I + 1
          Next
          Return True
        End Function)
      If ThisStockCount > 0 Then
        ThisTaskForStock.Start()

        Dim ThisCommpletion As Single
        Dim ThisCommpletionLast As Single
        RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.Information)

        Do
          Dim ThisTaskWait As New Task(Of Boolean)(
            Function()
              Thread.Sleep(1000)
              Return True
            End Function)

          ThisTaskWait.Start()
          Await ThisTaskWait
          ThisCommpletion = CSng(100 * (I / ThisStockCount))
          If (ThisCommpletion - ThisCommpletionLast) > 1.0 Then
            'send a message
            ThisCommpletionLast = ThisCommpletion
            RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.InformationUpdate)
          End If
        Loop Until ThisTaskForStock.IsCompleted
        ThisCommpletion = 100
        RaiseEvent Message(String.Format("Processing stock: {0:#.#}%", ThisCommpletion), IMessageInfoEvents.enuMessageType.InformationUpdate)
      End If
    End With
    ThisBinaryWriter = Nothing
    Return True
  End Function

  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
    Dim ThisBinaryWriter As New BinaryWriter(Stream)
    Dim ThisListException As List(Of Exception)
    Dim ThisException As Exception
    Dim ThisZero As Integer = 0

    With ThisBinaryWriter
      .Seek(0, SeekOrigin.Begin)
      .Write("Report")  'File identification
      .Write(Me.FileVersion)
      .Write(VERSION_MEMORY_STREAM)
      .Write(FileType)
      .Write(Me.ID)
      ThisListException = Me.Exception.ToList
      ThisListException.Reverse()
      .Write(ThisListException.Count)
      For Each ThisException In ThisListException
        .Write(ThisException.Message)
      Next
      .Write(Me.Name)
      .Write(Me.DateStart.ToBinary)
      .Write(Me.DateStop.ToBinary)
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          .Write(Me.BondRates.Count)
          For Each ThisBondRate In Me.BondRates
            ThisBondRate.SerializeSaveTo(Stream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          If TypeOf _BondRates Is IDataVirtual Then
            'make sure the data has been loaded before we save
            With DirectCast(_BondRates, IDataVirtual)
              Dim IsLoaded As Boolean = .IsLoaded
              If IsLoaded = False Then .Load()
            End With
          End If
          If _BondRates.Count > 0 Then
            Dim ThisRecordIndex As RecordIndex(Of BondRate, String)
            ThisRecordIndex = New RecordIndex(Of BondRate, String)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.AsDateRange, "BondRate", "", ".bra")
            ThisRecordIndex.Save(_BondRates)
            If (Me.FileType = IMemoryStream.enuFileType.RecordIndexed) Then
              With DirectCast(_BondRates, IDataVirtual)
                If .Enabled Then
                  .Release()
                End If
              End With
            End If
            .Write(ThisRecordIndex.FileCount)
            .Write(ThisRecordIndex.MaxID)
            .Write(ThisRecordIndex.DateStart.ToBinary)
            .Write(ThisRecordIndex.DateStop.ToBinary)
            With DirectCast(_BondRates, IDateUpdate)
              .DateStart = ThisRecordIndex.DateStart
              .DateStop = ThisRecordIndex.DateStop
            End With
            With DirectCast(_BondRates, IRecordInfo)
              .CountTotal = ThisRecordIndex.FileCount
              .MaximumID = ThisRecordIndex.MaxID
            End With
            ThisRecordIndex.Dispose()
            ThisRecordIndex = Nothing
          Else
            Dim ThisRecordInfo = DirectCast(_BondRates, IRecordInfo)
            Dim ThisDateUpdate = DirectCast(_BondRates, IDateUpdate)
            .Write(ThisRecordInfo.CountTotal)
            .Write(ThisRecordInfo.MaximumID)
            .Write(ThisDateUpdate.DateStart.ToBinary)
            .Write(ThisDateUpdate.DateStop.ToBinary)
          End If
      End Select
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          .Write(Me.SplitFactorFutures.Count)
          For Each ThisSplitFactorFuture In Me.SplitFactorFutures
            ThisSplitFactorFuture.SerializeSaveTo(Stream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          If TypeOf _SplitFactorFutures Is IDataVirtual Then
            'make sure the data has been loaded before we save
            With DirectCast(_SplitFactorFutures, IDataVirtual)
              Dim IsLoaded As Boolean = .IsLoaded
              If IsLoaded = False Then .Load()
            End With
          End If
          If _SplitFactorFutures.Count > 0 Then
            Dim ThisRecordIndex As RecordIndex(Of SplitFactorFuture, String)
            ThisRecordIndex = New RecordIndex(Of SplitFactorFuture, String)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.AsDateRange, "SplitFactorFuture", "", ".sff")
            ThisRecordIndex.Save(_SplitFactorFutures)
            If (Me.FileType = IMemoryStream.enuFileType.RecordIndexed) Then
              With DirectCast(_SplitFactorFutures, IDataVirtual)
                If .Enabled Then
                  .Release()
                End If
              End With
            End If
            .Write(ThisRecordIndex.FileCount)
            .Write(ThisRecordIndex.MaxID)
            .Write(ThisRecordIndex.DateStart.ToBinary)
            .Write(ThisRecordIndex.DateStop.ToBinary)
            With DirectCast(_SplitFactorFutures, IDateUpdate)
              .DateStart = ThisRecordIndex.DateStart
              .DateStop = ThisRecordIndex.DateStop
            End With
            With DirectCast(_SplitFactorFutures, IRecordInfo)
              .CountTotal = ThisRecordIndex.FileCount
              .MaximumID = ThisRecordIndex.MaxID
            End With
            ThisRecordIndex.Dispose()
            ThisRecordIndex = Nothing
          Else
            Dim ThisRecordInfo = DirectCast(_SplitFactorFutures, IRecordInfo)
            Dim ThisDateUpdate = DirectCast(_SplitFactorFutures, IDateUpdate)
            .Write(ThisRecordInfo.CountTotal)
            .Write(ThisRecordInfo.MaximumID)
            .Write(ThisDateUpdate.DateStart.ToBinary)
            .Write(ThisDateUpdate.DateStop.ToBinary)
          End If
      End Select

      .Write(Me.Sectors.Count)
      For Each ThisSector In Me.Sectors
        ThisSector.SerializeSaveTo(Stream)
      Next
      .Write(Me.Industries.Count)
      For Each ThisIndustry In Me.Industries
        ThisIndustry.SerializeSaveTo(Stream)
      Next
      Dim Count As Integer = Me.Stocks.Count
      .Write(Count)

      'Dim I As Integer
      'Dim ThisStopWatch As New Stopwatch
      'ThisStopWatch.Restart()
      'RaiseEvent Message(String.Format("Saving stock list..."), IMessageInfoEvents.enuMessageType.Information)
      For Each ThisStock As Stock In Me.Stocks
        ThisStock.SerializeSaveTo(Stream, FileType)
        'I = I + 1
        'If ThisStopWatch.ElapsedMilliseconds > 500 Then
        '	ThisStopWatch.Restart()
        '	RaiseEvent Message(String.Format("Saving stock list {0:P1}...", I / Count), IMessageInfoEvents.enuMessageType.Information)
        'End If
      Next
    End With
    ThisBinaryWriter = Nothing
  End Sub

  Public Sub SerializeLoadFrom(ByRef Data() As Byte) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisStream As Stream = New System.IO.MemoryStream(Data, writable:=True)
    Me.SerializeLoadFrom(ThisStream)
  End Sub

  Public Function SerializeSaveTo() As Byte() Implements IMemoryStream.SerializeSaveTo
    Dim ThisStream As Stream = New System.IO.MemoryStream
    Dim ThisBinaryReader As New BinaryReader(ThisStream)
    Dim ThisData As Byte()

    Me.SerializeSaveTo(ThisStream)
    ThisData = ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
    ThisBinaryReader.Dispose()
    ThisStream.Dispose()
    Return ThisData
  End Function
#End Region
#Region "IFormatData Implementation"
  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Return Extensions.ToStingOfData(Of Report)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = False})
      .Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateStart", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateStop", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IDateRange access and implementation"
  Public ReadOnly Property AsDateRange As IDateRange
    Get
      Return Me
    End Get
  End Property

  Private Property IDateRange_DateStart As Date Implements IDateRange.DateStart
    Get
      If MyDateRangeStart.HasValue Then
        Return CDate(MyDateRangeStart)
      Else
        Return Me.DateStart
      End If
    End Get
    Set(value As Date)
      MyDateRangeStart = value
    End Set
  End Property

  Private Property IDateRange_DateStop As Date Implements IDateRange.DateStop
    Get
      If MyDateRangeStop.HasValue Then
        Return CDate(MyDateRangeStop)
      Else
        Return Me.DateStop
      End If
    End Get
    Set(value As Date)
      MyDateRangeStop = value
    End Set
  End Property

  Private Sub IDateRange_MoveBegin() Implements IDateRange.MoveBegin
    MyDateRangeStart = Me.DateStart
    MyDateRangeStop = Me.DateStart
    Me.IDateRange_Refresh()
  End Sub

  Private Sub IDateRange_MoveLast() Implements IDateRange.MoveLast
    MyDateRangeStart = Me.DateStart
    MyDateRangeStop = Me.DateStop
    Me.IDateRange_Refresh()
  End Sub

  Private Sub IDateRange_MoveNext() Implements IDateRange.MoveNext
    If MyDateRangeStart.HasValue And MyDateRangeStop.HasValue Then
      MyDateRangeStart = CDate(MyDateRangeStart).AddDays(1)
      MyDateRangeStop = CDate(MyDateRangeStop).AddDays(1)
      Me.IDateRange_Refresh()
    End If
  End Sub

  Private Sub IDateRange_MovePrevious() Implements IDateRange.MovePrevious
    If MyDateRangeStart.HasValue And MyDateRangeStop.HasValue Then
      MyDateRangeStart = CDate(MyDateRangeStart).AddDays(-1)
      MyDateRangeStop = CDate(MyDateRangeStop).AddDays(-1)
      Me.IDateRange_Refresh()
    End If
  End Sub

  Private Function IDateRange_NumberDays() As Integer Implements IDateRange.NumberDays
    If MyDateRangeStart.HasValue And MyDateRangeStop.HasValue Then
      Return CDate(MyDateRangeStop).Subtract(CDate(MyDateRangeStart)).Days
    Else
      Return Me.DateStop.Subtract(Me.DateStart).Days
    End If
  End Function

  Private Sub IDateRange_Refresh() Implements IDateRange.Refresh
    If Me.IsFileOpen Then
      _BondRates.Clear()
      'Dim ThisRecordIndexBondRate As New RecordIndex(Of BondRate, String)(MyStream, FileMode.Open, Me.AsDateRange, "BondRate", "", ".bra")
      'With ThisRecordIndexBondRate
      '	For Each ThisPosition In .ToListPosition
      '		Dim ThisBondRate As New BondRate(Me, .BaseStream(ThisPosition))
      '	Next
      'End With
      'With CType(Me.BondRates, IRecordInfo)
      '	.CountTotal = ThisRecordIndexBondRate.FileCount
      '	.MaximumID = ThisRecordIndexBondRate.MaxID
      'End With
      'ThisRecordIndexBondRate.Dispose()
      'ThisRecordIndexBondRate = Nothing

      _SplitFactorFutures.Clear()
      'Dim ThisRecordIndexSplitFactorFuture As New RecordIndex(Of SplitFactorFuture, String)(MyStream, FileMode.Open, Me.AsDateRange, "SplitFactorFuture", "", ".sff")
      'With ThisRecordIndexSplitFactorFuture
      '	For Each ThisPosition In .ToListPosition
      '		Dim ThisSplitFactorFuture As New SplitFactorFuture(Me, .BaseStream(ThisPosition))
      '	Next
      '	With CType(Me.SplitFactorFutures, IRecordInfo)
      '		.CountTotal = ThisRecordIndexSplitFactorFuture.FileCount
      '		.MaximumID = ThisRecordIndexSplitFactorFuture.MaxID
      '	End With
      'End With
      'ThisRecordIndexSplitFactorFuture.Dispose()
      'ThisRecordIndexSplitFactorFuture = Nothing

      Dim I As Integer
      Dim Count As Integer = Me.Stocks.Count
      Dim ThisStopWatch As New Stopwatch
      ThisStopWatch.Restart()
      RaiseEvent Message(String.Format("Stock list refreshing..."), IMessageInfoEvents.enuMessageType.Information)
      For Each ThisStock As Stock In Me.Stocks
        ThisStock.RefreshDate()
        I = I + 1
        If ThisStopWatch.ElapsedMilliseconds > 500 Then
          ThisStopWatch.Restart()
          RaiseEvent Message(String.Format("Stock list refresh {0:P1}...", I / Count), IMessageInfoEvents.enuMessageType.Information)
        End If
      Next
    End If
  End Sub
  Private Sub IDateRange_Refresh(DateStart As Date, DateStop As Date) Implements IDateRange.Refresh
    MyDateRangeStart = DateStart
    MyDateRangeStop = DateStop
    Me.IDateRange_Refresh()
  End Sub

  Private Sub IDateRange_Refresh(DateStart As Date, NumberDays As Integer) Implements IDateRange.Refresh
    MyDateRangeStart = DateStart
    MyDateRangeStop = CDate(MyDateRangeStart).AddDays(NumberDays)
    Me.IDateRange_Refresh()
  End Sub

  Private Sub IDateRange_Refresh(NumberDays As Integer, DateStop As Date) Implements IDateRange.Refresh
    MyDateRangeStop = DateStop
    MyDateRangeStart = CDate(MyDateRangeStop).AddDays(-NumberDays)
    Me.IDateRange_Refresh()
  End Sub
#End Region
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return Me.DateStart
    End Get
    Set(value As Date)
    End Set
  End Property
  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return Me.DateStop
    End Get
    Set(value As Date)
    End Set
  End Property
  Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return IDateUpdate_DateStop
    End Get
  End Property

  Private ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return IDateUpdate_DateStop
    End Get
  End Property

  Private ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return IDateUpdate_DateStop.Date
    End Get
  End Property
#End Region 'IDateUpdate
#Region "IMessageInfoEvents"
  Public Sub RaiseMessage(ByVal Message As String, MessageType As IMessageInfoEvents.enuMessageType)
    RaiseMessage(Message, MessageType)
  End Sub

  Public Event Message(ByVal Message As String, MessageType As IMessageInfoEvents.enuMessageType) Implements IMessageInfoEvents.Message
#End Region
#Region "ISystemEventOfBondRate1"
  Private Sub ISystemEventOfBondRate1_Add(item As BondRate1) Implements ISystemEvent(Of BondRate1).Add

  End Sub

  Private Sub ISystemEventOfBondRate1_Clear() Implements ISystemEvent(Of BondRate1).Clear

  End Sub

  Private Sub ISystemEventOfBondRate1_Load() Implements ISystemEvent(Of BondRate1).Load
    'proceed with reading only if the time overlap
    Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me, IDateRange))

    If ThisTimePeriodOverlap.IsOverlap(DirectCast(_BondRates1, IDateUpdate)) Then
      'the time overlap
      Dim ThisRecordIndexBondRate As New RecordIndex(Of BondRate1, String)(DirectCast(Me.ToStream, FileStream).Name(), FileMode.Open, Me.AsDateRange, "BondRate1", "", ".bra")
      With ThisRecordIndexBondRate
        For Each ThisPosition In .ToListPosition
          Dim ThisBondRate1 As New BondRate1(Me, .BaseStream(ThisPosition))
        Next
      End With
      With DirectCast(_BondRates1, IDateUpdate)
        .DateStart = ThisRecordIndexBondRate.DateStart
        .DateStop = ThisRecordIndexBondRate.DateStop
      End With
      With CType(_BondRates1, IRecordInfo)
        .CountTotal = ThisRecordIndexBondRate.FileCount
        .MaximumID = ThisRecordIndexBondRate.MaxID
      End With
      ThisRecordIndexBondRate.Dispose()
      ThisRecordIndexBondRate = Nothing
    End If
  End Sub

  Private Function ISystemEventOfBondRate1_Remove(item As BondRate1) As Boolean Implements ISystemEvent(Of BondRate1).Remove

  End Function
#End Region
#Region "ISystemEventOfBondRate"
  Private Sub ISystemEventOfBondRate_Add(item As BondRate) Implements ISystemEvent(Of BondRate).Add
    Dim ThisBondRate1 As BondRate1
    'specifically use the symbol for BondRate for the search and not the KeyValue 
    'which combine the symbol and the date as a key
    'the new version Bondrate1 use the symbol only
    ThisBondRate1 = Me.BondRates1.ToSearch.Find(item.Symbol)
    If ThisBondRate1 Is Nothing Then
      ThisBondRate1 = New BondRate1
      With item
        ThisBondRate1.ID = .ID
        ThisBondRate1.Symbol = .Symbol
        ThisBondRate1.Name = .Name
        ThisBondRate1.Exchange = ""
        ThisBondRate1.MaturityDays = .MaturityDays
        ThisBondRate1.Security = CType(.Security, YahooAccessData.BondRate1.enuBondSecurity)
        ThisBondRate1.Type = CType(.Type, YahooAccessData.BondRate1.enuBondType)
        ThisBondRate1.DateStart = .DateUpdate
        ThisBondRate1.DateStop = .DateUpdate
        ThisBondRate1.ErrorDescription = ""
        ThisBondRate1.IsError = False
        ThisBondRate1.Exception = .Exception
        ThisBondRate1.Report = Me
        Dim ThisBondRateRecord As New BondRateRecord
        ThisBondRateRecord.DateUpdate = .DateUpdate
        ThisBondRateRecord.Interest = .Interest
        ThisBondRate1.Records.Add(ThisBondRateRecord)
        Me.BondRates1.Add(ThisBondRate1)
      End With
    Else
      With item
        ThisBondRate1.Name = .Name
        ThisBondRate1.Exchange = ""
        ThisBondRate1.Security = CType(.Security, YahooAccessData.BondRate1.enuBondSecurity)
        ThisBondRate1.Type = CType(.Type, YahooAccessData.BondRate1.enuBondType)
        If .DateUpdate < ThisBondRate1.DateStart Then
          ThisBondRate1.DateStart = .DateUpdate
        End If
        If .DateUpdate > DateSerial(2012, 9, 30) Then
          .DateUpdate = .DateUpdate
        End If
        If .DateUpdate > ThisBondRate1.DateStop Then
          ThisBondRate1.DateStop = .DateUpdate
        End If
        ThisBondRate1.Exception = .Exception
        'If .Exception IsNot Nothing Then
        '  .Exception = .Exception

        'End If
        Dim ThisBondRateRecord As New BondRateRecord
        ThisBondRateRecord.DateUpdate = .DateUpdate
        ThisBondRateRecord.Interest = .Interest
        ThisBondRate1.Records.Add(ThisBondRateRecord)
      End With
    End If
  End Sub

  Private Sub ISystemEventOfBondRate_Clear() Implements ISystemEvent(Of BondRate).Clear
    Me.BondRates1.Clear()
  End Sub

  Private Sub ISystemEventOfBondRate_Load() Implements ISystemEvent(Of BondRate).Load
    'proceed with reading only if the time overlap
    Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me, IDateRange))

    If ThisTimePeriodOverlap.IsOverlap(DirectCast(_BondRates, IDateUpdate)) Then
      'the time overlap
      Dim ThisRecordIndexBondRate As New RecordIndex(Of BondRate, String)(DirectCast(Me.ToStream, FileStream).Name(), FileMode.Open, Me.AsDateRange, "BondRate", "", ".bra")
      With ThisRecordIndexBondRate
        For Each ThisPosition In .ToListPosition
          Dim ThisBondRate As New BondRate(Me, .BaseStream(ThisPosition))
        Next
      End With
      With DirectCast(_BondRates, IDateUpdate)
        .DateStart = ThisRecordIndexBondRate.DateStart
        .DateStop = ThisRecordIndexBondRate.DateStop
      End With
      With CType(_BondRates, IRecordInfo)
        .CountTotal = ThisRecordIndexBondRate.FileCount
        .MaximumID = ThisRecordIndexBondRate.MaxID
      End With
      ThisRecordIndexBondRate.Dispose()
      ThisRecordIndexBondRate = Nothing
    End If
  End Sub

  Private Function ISystemEventOfBondRate_Remove(item As BondRate) As Boolean Implements ISystemEvent(Of BondRate).Remove
    Return Me.BondRates1.Remove(Me.BondRates1.ToSearch.Find(item.KeyValue))
  End Function
#End Region
#Region "ISystemEventOfSplitFactorFuture"
  Private Sub ISystemEventOfSplitFactorFuture_Add(item As SplitFactorFuture) Implements ISystemEvent(Of SplitFactorFuture).Add

  End Sub

  Private Sub ISystemEventOfSplitFactorFuture_Clear() Implements ISystemEvent(Of SplitFactorFuture).Clear

  End Sub

  Private Sub ISystemEventOfSplitFactorFuture_Load() Implements ISystemEvent(Of SplitFactorFuture).Load
    'proceed with reading only if the time overlap
    Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me, IDateRange))

    If ThisTimePeriodOverlap.IsOverlap(DirectCast(_SplitFactorFutures, IDateUpdate)) Then
      Dim ThisRecordIndexSplitFactorFuture As New RecordIndex(Of SplitFactorFuture, String)(DirectCast(Me.ToStream, FileStream).Name(), FileMode.Open, Me.AsDateRange, "SplitFactorFuture", "", ".sff")
      With ThisRecordIndexSplitFactorFuture
        For Each ThisPosition In .ToListPosition
          Dim ThisSplitFactorFuture As New SplitFactorFuture(Me, .BaseStream(ThisPosition))
        Next
        With DirectCast(_SplitFactorFutures, IDateUpdate)
          .DateStart = ThisRecordIndexSplitFactorFuture.DateStart
          .DateStop = ThisRecordIndexSplitFactorFuture.DateStop
        End With
        With CType(_SplitFactorFutures, IRecordInfo)
          .CountTotal = ThisRecordIndexSplitFactorFuture.FileCount
          .MaximumID = ThisRecordIndexSplitFactorFuture.MaxID
        End With
      End With
      ThisRecordIndexSplitFactorFuture.Dispose()
      ThisRecordIndexSplitFactorFuture = Nothing
    End If
  End Sub

  Private Function ISystemEventOfSplitFactorFuture_Remove(item As SplitFactorFuture) As Boolean Implements ISystemEvent(Of SplitFactorFuture).Remove

  End Function
#End Region
#Region "IStockRecordEvent"
  Public ReadOnly Property IsLoading As Boolean Implements IStockRecordEvent.IsLoading
    Get
      Return IsRecordLoading
    End Get
  End Property

  Public ReadOnly Property SymbolLoading As String Implements IStockRecordEvent.SymbolLoading
    Get
      Return MySymbolRecordLoading
    End Get
  End Property

  Public ReadOnly Property IsLoadCancel As Boolean Implements IStockRecordEvent.IsLoadCancel
    Get
      Return IsRecordCancel
    End Get
  End Property


  Public Sub LoadAfter(Symbol As String) Implements IStockRecordEvent.LoadAfter
    SyncLock MySyncLockForRecordLoading

      If IStockRecordInfo_Enabled Then
        If (Me.FileType = IMemoryStream.enuFileType.RecordIndexed) Then
          'make sure this code section execute only one at a time
          'this is needed to support properly the threading use in loading to cache 

          MyStockRecordQueue.Enqueue(Symbol)
          'keep a record of which stock is loaded
          If MyDictionaryOfStockRecordLoaded.ContainsKey(Symbol) = False Then
            MyDictionaryOfStockRecordLoaded.Add(Symbol, Symbol)
          End If
          Do While MyStockRecordQueue.Count > MyStockRecordQueueCount
            Dim ThisStockSymbol As String = MyStockRecordQueue.Dequeue
            If ThisStockSymbol <> Symbol Then
              'We can release the data only if the symbol is different
              'than the one we just load
              With Me.Stocks.ToSearch.Find(ThisStockSymbol)
                .RecordReleaseAll()
              End With
              'remove this symbol from the list of loaded stock
              If MyDictionaryOfStockRecordLoaded.ContainsKey(ThisStockSymbol) Then
                MyDictionaryOfStockRecordLoaded.Remove(ThisStockSymbol)
              End If
            End If
          Loop
        End If
      End If
      IsRecordLoading = False
      MySymbolRecordLoading = ""
    End SyncLock
  End Sub

  Public Sub LoadBefore(Symbol As String) Implements IStockRecordEvent.LoadBefore
    SyncLock MySyncLockForRecordLoading
      IsRecordCancel = False
      IsRecordLoading = True
      MySymbolRecordLoading = Symbol
    End SyncLock
  End Sub

  Public Sub LoadCancel(Symbol As String) Implements IStockRecordEvent.LoadCancel
    SyncLock MySyncLockForRecordLoading
      IsRecordCancel = True
      IsRecordLoading = False
      MySymbolRecordLoading = ""
    End SyncLock
  End Sub
#End Region
#Region "IRecordControlInfo"
  ReadOnly Property AsRecordControlInfo As IRecordControlInfo
    Get
      Return Me
    End Get
  End Property

  Private Property IStockRecordInfo_ControlType As enuStockRecordLoadType Implements IRecordControlInfo.ControlType

  Private Property IStockRecordInfo_Count As Integer Implements IRecordControlInfo.Count
    Get
      Return MyStockRecordQueueCount
    End Get
    Set(value As Integer)
      MyStockRecordQueueCount = value
    End Set
  End Property

  Private Property IStockRecordInfo_Enabled As Boolean Implements IRecordControlInfo.Enabled

  ''' <summary>
  ''' Load the data immediatly for this list of stock
  ''' </summary>
  ''' <param name="StockList"></param>
  ''' <remarks></remarks>
  Public Sub LoadToCache(ByVal StockList As IEnumerable(Of String)) Implements IRecordControlInfo.LoadToCache
    Dim ThisListOfStock As New List(Of YahooAccessData.Stock)

    'take only the symbol that are not yet loaded
    SyncLock MySyncLockForRecordLoading
      For Each ThisSymbol In StockList
        If MyDictionaryOfStockRecordLoaded.ContainsKey(ThisSymbol) = False Then
          ThisListOfStock.Add(Me.Stocks.ToSearch.Find(ThisSymbol))
        End If
      Next
      If ThisListOfStock.Count = 0 Then Exit Sub
    End SyncLock
    Me.LoadToCache(ThisListOfStock)
  End Sub

  Public Sub LoadToCache(StockSymbol As String) Implements IRecordControlInfo.LoadToCache
    Dim ThisListOfStock As New List(Of YahooAccessData.Stock)

    SyncLock MySyncLockForRecordLoading
      If MyDictionaryOfStockRecordLoaded.ContainsKey(StockSymbol) = False Then
        ThisListOfStock.Add(Me.Stocks.ToSearch.Find(StockSymbol))
        If ThisListOfStock.Count = 0 Then Exit Sub
      End If
    End SyncLock
    Me.LoadToCache(ThisListOfStock)
  End Sub

  Public Sub LoadToCache(Stock As Stock) Implements IRecordControlInfo.LoadToCache
    Dim ThisListOfStock As New List(Of YahooAccessData.Stock)

    SyncLock MySyncLockForRecordLoading
      If MyDictionaryOfStockRecordLoaded.ContainsKey(Stock.Symbol) = False Then
        ThisListOfStock.Add(Stock)
        If ThisListOfStock.Count = 0 Then Exit Sub
      End If
    End SyncLock
    Me.LoadToCache(ThisListOfStock)
  End Sub

  Public Async Sub LoadToCache(ByVal StockList As IEnumerable(Of YahooAccessData.Stock)) Implements IRecordControlInfo.LoadToCache
    Dim ThisList As New List(Of YahooAccessData.Stock)
    Dim ThisTaskOfLoadCache As Task(Of LoadToCacheStatus)
    Dim ThisQueue As Queue(Of YahooAccessData.Stock)
    Dim ThisStock As YahooAccessData.Stock = Nothing
    Dim ThisTick As Integer



    Exit Sub

    'take only the symbol that are not yet loaded
    SyncLock MySyncLockForRecordLoading
      For Each ThisStock In StockList
        If MyDictionaryOfStockRecordLoaded.ContainsKey(ThisStock.KeyValue) = False Then
          ThisList.Add(ThisStock)
        End If
      Next
      If ThisList.Count = 0 Then Exit Sub
    End SyncLock

    'ThisQueue = New Queue(Of YahooAccessData.Stock)(ThisList)
    ThisTick = Environment.TickCount
    MyLoadToCacheLatestTick = ThisTick
    For Each ThisStock In ThisList
      If ThisStock IsNot Nothing Then
        Await ThisStock.RecordLoadAsync
      End If
      If MyLoadToCacheLatestTick <> ThisTick Then Exit For
    Next

    'Do Until ThisQueue.Count = 0
    '  ThisStock = ThisQueue.Dequeue


    '  SyncLock MySyncLockForRecordLoading

    '  End SyncLock



    '        If MyLoadToCacheLatestTick <> ThisTick Then
    '          ThisLoadToCacheStatus.NumberOfLoadToCacheCancelled = ThisQueue.Count
    '          ThisLoadToCacheStatus.IsCancelled = True
    '        End If
    '      End SyncLock
    '      If ThisLoadToCacheStatus.IsCancelled = True Then Exit Do
    '      ThisStock = ThisQueue.Dequeue
    '      If ThisStock IsNot Nothing Then
    '        'this command load the data in memory
    '        ThisStock.RecordLoad()
    '        ThisLoadToCacheStatus.NumberOfLoadToCacheCompleted = ThisLoadToCacheStatus.NumberOfLoadToCacheCompleted + 1
    '      End If
    'Loop

    'Dim progress = New Progress(Of LoadToCacheStatus)(
    '  Sub(Status As LoadToCacheStatus)
    '    'RaiseEvent Message(String.Format(
    '  End Sub)

    'ThisTaskOfLoadCache = New Task(Of LoadToCacheStatus)(
    '  Function()
    '    Dim ThisStock As YahooAccessData.Stock = Nothing
    '    Dim ThisLoadToCacheStatus As LoadToCacheStatus

    '    ThisLoadToCacheStatus = New LoadToCacheStatus(ThisQueue.Count)

    '    Do Until ThisQueue.Count = 0
    '      SyncLock MySyncLockForRecordLoading
    '        If MyLoadToCacheLatestTick <> ThisTick Then
    '          ThisLoadToCacheStatus.NumberOfLoadToCacheCancelled = ThisQueue.Count
    '          ThisLoadToCacheStatus.IsCancelled = True
    '        End If
    '      End SyncLock
    '      If ThisLoadToCacheStatus.IsCancelled = True Then Exit Do
    '      ThisStock = ThisQueue.Dequeue
    '      If ThisStock IsNot Nothing Then
    '        'this command load the data in memory
    '        ThisStock.RecordLoad()
    '        ThisLoadToCacheStatus.NumberOfLoadToCacheCompleted = ThisLoadToCacheStatus.NumberOfLoadToCacheCompleted + 1
    '      End If
    '    Loop
    '    Return ThisLoadToCacheStatus
    '  End Function)

    'SyncLock MySyncLockForRecordLoading
    '  ThisTick = Environment.TickCount
    '  MyLoadToCacheLatestTick = ThisTick
    '  ThisTaskOfLoadCache.Start()
    'End SyncLock
    'Await ThisTaskOfLoadCache
  End Sub

  ' Notify when task is cancelled
  Private Sub LoadToCacheCancelledMessage()
    RaiseEvent Message("Cancellation request made!!", IMessageInfoEvents.enuMessageType.Information)
  End Sub

  Private Class LoadToCacheStatus
    Private MyNumberOfLoadToCache As Integer

    Public Sub New(ByVal NumberOfLoadToCache As Integer)
      MyNumberOfLoadToCache = NumberOfLoadToCache
    End Sub

    Public Property NumberOfLoadToCacheCompleted As Integer
    Public Property NumberOfLoadToCacheCancelled As Integer
    Public Property IsCancelled As Boolean
    Public ReadOnly Property NumberOfLoadToCache As Integer
      Get
        Return MyNumberOfLoadToCache
      End Get
    End Property
  End Class
#End Region
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' TODO: dispose managed state (managed objects).
        Me.Clear()
        If MyStream IsNot Nothing Then
          MyStream.Dispose()
          MyStream = Nothing
        End If
      End If

      ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' TODO: set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

  ''TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub
#End Region
End Class

#Region "Report Interface"
Public Interface IReport
  Property ID As Integer
  Property Name As String
  Property DateStart As Date
  Property DateStop As Date

  Property Industries As ICollection(Of Industry)
  Property Sectors As ICollection(Of Sector)
  Property Stocks As ICollection(Of Stock)
  Property SplitFactorFutures As ICollection(Of SplitFactorFuture)
  Property BondRates As ICollection(Of BondRate)
End Interface

Public Interface IReportHeader
  ReadOnly Property ID As Integer
  ReadOnly Property Name As String
  ReadOnly Property FileVersion As String
  ReadOnly Property DateStart As Date
  ReadOnly Property DateStop As Date
  Property Exception As Exception
End Interface

Public Class ReportHeader
  Implements IReportHeader

  Private MyDateStart As Date
  Private MyDateStop As Date
  Private MyID As Integer
  Private MyName As String
  Private MyFileVersion As String

  Public Sub New()
    MyFileVersion = ""
    MyDateStart = YahooAccessData.ReportDate.DATE_NULL_VALUE
    MyDateStop = YahooAccessData.ReportDate.DATE_NULL_VALUE
    MyID = 0
    MyName = ""
    Me.Exception = Nothing
  End Sub

  'use if you are only interested in date
  Public Sub New(ByVal DateStart As Date, ByVal DateStop As Date)
    MyDateStart = DateStart
    MyDateStop = DateStop
  End Sub

  Public Sub New(ByVal FileVersion As String, ByVal ID As Integer, ByVal Name As String, ByVal DateStart As Date, ByVal DateStop As Date)
    MyFileVersion = FileVersion
    MyDateStart = DateStart
    MyDateStop = DateStop
    MyID = ID
    MyName = Name
    Me.Exception = Nothing
  End Sub

  Public Sub New(ByVal Exception As Exception)
    Me.New()
    Me.Exception = Exception
  End Sub

  Public Sub New(ByVal ReportHeader As IReportHeader)
    Me.New()
    If ReportHeader IsNot Nothing Then
      With ReportHeader
        MyFileVersion = .FileVersion
        MyDateStart = .DateStart
        MyDateStop = .DateStop
        MyID = .ID
        MyName = .Name
        Me.Exception = .Exception
      End With
    End If
  End Sub

  Public ReadOnly Property DateStart As Date Implements IReportHeader.DateStart
    Get
      Return MyDateStart
    End Get
  End Property

  Public ReadOnly Property DateStop As Date Implements IReportHeader.DateStop
    Get
      Return MyDateStop
    End Get
  End Property

  Public ReadOnly Property ID As Integer Implements IReportHeader.ID
    Get
      Return MyID
    End Get
  End Property

  Public ReadOnly Property Name As String Implements IReportHeader.Name
    Get
      Return MyName
    End Get
  End Property

  Public ReadOnly Property FileVersion As String Implements IReportHeader.FileVersion
    Get
      Return MyFileVersion
    End Get
  End Property

  Public Property Exception As Exception Implements IReportHeader.Exception

  ''' <summary>
  ''' Use to update only the date range. 
  ''' The update is done on condition the ReportHeader object does not report on an exception.
  ''' </summary>
  ''' <param name="ReportHeader"></param>
  ''' <remarks>Can be use to keep track of the date range (minimum and maximum) for a set of ReportHeader</remarks>
  Public Sub UpdateDateRange(ByVal ReportHeader As IReportHeader)
    With ReportHeader
      If .Exception Is Nothing Then
        If .DateStart < MyDateStart Then
          MyDateStart = .DateStart
        End If
        If .DateStop > MyDateStop Then
          MyDateStop = .DateStop
        End If
      End If
    End With
  End Sub
End Class
#End Region
'#Region "IStockExchange"
'  Public Interface IStockExchange
'    Property Name As String
'    Property Stocks As ICollection(Of Stock)
'  End Interface
'#End Region

#Region "FileAccess"
Public Class ReportFrom
  ''' <summary>
  ''' Use to read only the Header of the report file.
  ''' This function minimize the footprint needed to read the header
  ''' </summary>
  ''' <param name="FileName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function LoadHeader(ByVal FileName As String) As IReportHeader
    Dim ThisStream As Stream = Nothing
    Dim ThisFormatter = New BinaryFormatter
    Dim ThisBinaryReader As BinaryReader = Nothing
    Dim ThisFileVersionIdentification As String
    Dim ThisVersion As Single
    Dim ThisFileVersion As String
    Dim ThisID As Integer
    Dim I As Integer
    Dim ThisDateStart As Date
    Dim ThisDateStop As Date
    Dim ThisFileType As IMemoryStream.enuFileType
    Dim ThisName As String
    Dim ThisReportHeader As IReportHeader

    Try
      ThisStream = New FileStream(FileName, FileMode.Open, FileAccess.Read, FileShare.Read)
      System.IO.File.SetAttributes(FileName, IO.FileAttributes.Normal)
      ThisBinaryReader = New BinaryReader(ThisStream, New System.Text.UTF8Encoding(), leaveOpen:=True)
      ThisFileVersionIdentification = ThisBinaryReader.ReadString
      If ThisFileVersionIdentification = "Report" Then
        'this is report file
        ThisFileVersion = ThisBinaryReader.ReadString
        ThisVersion = ThisBinaryReader.ReadSingle
        ThisFileType = CType(ThisBinaryReader.ReadInt32, IMemoryStream.enuFileType)
        ThisID = ThisBinaryReader.ReadInt32
        For I = 1 To ThisBinaryReader.ReadInt32
        Next
        ThisName = ThisBinaryReader.ReadString
        ThisDateStart = DateTime.FromBinary(ThisBinaryReader.ReadInt64)
        ThisDateStop = DateTime.FromBinary(ThisBinaryReader.ReadInt64)
        ThisReportHeader = New ReportHeader(ThisFileVersion, ThisID, ThisName, ThisDateStart, ThisDateStop)
      Else
        ThisReportHeader = New ReportHeader(New Exception("Invalid file format..."))
      End If
    Catch e As Exception
      ThisReportHeader = New ReportHeader(e)
    End Try
    If ThisBinaryReader IsNot Nothing Then
      ThisBinaryReader.Dispose()
    End If
    If ThisStream IsNot Nothing Then
      ThisStream.Dispose()
    End If
    ThisBinaryReader = Nothing
    ThisStream = Nothing
    Return ThisReportHeader
  End Function

  Public Shared Function CreateFile(ByRef Report As Report, ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType) As Boolean
    Dim ThisReport As New YahooAccessData.Report

    Return ThisReport.FileSave(Report, FileName, FileType)
  End Function

  Public Shared Function CreateFile(ByVal FileName As String, ByVal FileType As IMemoryStream.enuFileType) As Boolean
    Dim ThisReport As New YahooAccessData.Report

    Return CreateFile(ThisReport, FileName, FileType)
  End Function

  Public Shared Function Load(ByVal FileName As String, ByVal IsRecordVirtual As Boolean) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report
    Return ThisReport.FileLoad(FileName, IsRecordVirtual)
  End Function

  Public Shared Function Load(ByVal FileName As String, ByVal IsRecordVirtual As Boolean, ByVal DateStart As Date, ByVal DateStop As Date) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report
    Return ThisReport.FileLoad(FileName, IsRecordVirtual, DateStart, DateStop)
  End Function

  Public Shared Function Load(ByVal FileName As String) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report

    Return ThisReport.FileLoad(FileName)
  End Function

  Public Shared Function Load(ByRef Stream As System.IO.Stream, ByVal IsRecordVirtual As Boolean, ByVal DateStart As Date, ByVal DateStop As Date) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report
    Return ThisReport.FileLoad(Stream, IsRecordVirtual, DateStart, DateStop)
  End Function

  Public Shared Function Load(ByRef Stream As System.IO.Stream, ByVal IsRecordVirtual As Boolean) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report
    Return ThisReport.FileLoad(Stream, IsRecordVirtual)
  End Function

  Public Shared Function Load(ByRef Stream As System.IO.Stream) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report

    Return ThisReport.FileLoad(Stream)
  End Function

  Public Shared Function LoadZip(ByVal FileZipName As String) As YahooAccessData.Report
    Dim ThisReport As New YahooAccessData.Report

    Return ThisReport.FileLoadZip(FileZipName)
  End Function

  Public Shared Function Version() As String
    Dim ThisReport As New Report
    Return ThisReport.FileVersion
  End Function
End Class
#End Region

#Region "ReportFileStatus"
'<Serializable()>
'Public Class ReportFileStatus
'  Public Enum enuFileStatus
'    OpenToRead
'    OpenToReadWrite
'    Close
'  End Enum

'  Public Sub New()
'    Me.DateUpdate = YahooAccessData.ReportDate.DateNullValue
'  End Sub

'  Public Property DateUpdate As Date
'End Class
#End Region

#Region "Template"


'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class Report
  Public Property ID As Integer
  Public Property Name As String = ""
  'Public Property DateStart As Date
  'Public Property DateStop As Date

  'Public Overridable Property Industries As ICollection(Of Industry) = New HashSet(Of Industry)
  'Public Overridable Property Sectors As ICollection(Of Sector) = New HashSet(Of Sector)
  'Public Overridable Property Stocks As ICollection(Of Stock) = New HashSet(Of Stock)
  'Public Overridable Property SplitFactorFutures As ICollection(Of SplitFactorFuture) = New HashSet(Of SplitFactorFuture)
  'Public Overridable Property BondRates As ICollection(Of BondRate) = New HashSet(Of BondRate)

End Class
#End Region





