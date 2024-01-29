#Region "Imports"
'Imports System
'Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
Imports System.Threading.Tasks
Imports YahooAccessData.ExtensionService.Extensions
Imports StockViewInterface
#End Region

<Serializable()>
Partial Public Class Stock
	Implements IEquatable(Of Stock)
	Implements IRegisterKey(Of String)
	Implements ISystemEvent(Of Record)
	Implements ISystemEvent(Of RecordDaily)
	Implements ISystemEvent(Of StockError)
	Implements ISystemEvent(Of StockSymbol)
	Implements ISystemEvent(Of SplitFactor)
	Implements IComparable(Of Stock)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate
	Implements IDisposable
	Implements IStockProcess
	Implements IStockInfo
	Implements IStockRank
	Implements IWebYahooDescriptor


#Region "Definition"
#Const IsLogSavingAction = False

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Private MyRecordQuoteValues As LinkedHashSet(Of RecordQuoteValue, Date)
	Private MyStreamPositionWhereRecordStart As Long
	Private MyStreamPositionWhereRecordStop As Long
	Private MyFileNumberRecord As Integer = 0
	'Private MyRecordIndexOfSplitFactor As RecordIndex(Of SplitFactor, Date)
	Private MyRecordIndexOfRecordDaily As RecordIndex(Of RecordDaily, Date)
	Private MySyncLockForRecordLoading As Object = New Object
	Private MySyncLockForRecordDailyLoading As Object = New Object
	Private Shared MySyncLockForStockRecordLoadingAll As Object = New Object
	Private Shared MyLastTimeOfRecordQuoteValuesAsync As Integer
	Private Shared MyLastTimeOfRecordLoadAsync As Integer
	'Private Shared MyStockSymbol

	Private _Records As ICollection(Of Record) = New HashSet(Of Record)
	Private _RecordsDaily As ICollection(Of RecordDaily)
	Private _SplitFactors As ICollection(Of SplitFactor) = New HashSet(Of SplitFactor)
	Private _StockErrors As ICollection(Of StockError) = New HashSet(Of StockError)
	Private _StockSymbols As ICollection(Of StockSymbol) = New HashSet(Of StockSymbol)
	Dim MyFiscalYearEnd As Date = YahooAccessData.ReportDate.DateNullValue
	Dim MyFiscalYearEndMeasured As Date = YahooAccessData.ReportDate.DateNullValue
#End Region
#Region "New"
	Public Sub New(ByRef Parent As YahooAccessData.Report, ByVal Symbol As String, ByVal Name As String, ByVal Exchange As String)
		Me.New(Symbol, Name, Exchange)
		With Me
			.Report = Parent
			If .Report IsNot Nothing Then
				.ReportID = .Report.ID
				.Report.Stocks.Add(Me)
			End If
		End With
	End Sub

	Public Sub New(ByVal Symbol As String, ByVal Name As String, ByVal Exchange As String)
		Me.New(Symbol, Name, Exchange, Nothing, False)
	End Sub

	Public Sub New(
		ByVal Symbol As String,
		ByVal Name As String,
		ByVal Exchange As String,
		ByRef Stream As Stream,
		ByVal IsRecordVirtual As Boolean)

		With Me
			.DateStart = Now
			.DateStop = Me.DateStart
			.Name = Name
			.Symbol = Symbol
			.ErrorDescription = ""
			.Exchange = Exchange
			'even if IsRecordVirtual is false 
			'activate the sink event ISystemEvent to get the update for IRecordQuoteValue
			'the Me parameters will activate the automatic loading of the data when needed
			'note that for Web update this event should be turned off i.e. IsRecordVirtual==false
			If IsRecordVirtual Then
				'the Me parameters will activate the automatic loading of the data when needed
				.Records = New LinkedHashSet(Of Record, Date)(Me)
				.RecordsDaily = New LinkedHashSet(Of RecordDaily, Date)(Me)
				.SplitFactors = New LinkedHashSet(Of SplitFactor, Date)(Me)
				.StockErrors = New LinkedHashSet(Of StockError, Date)(Me)
				.StockSymbols = New LinkedHashSet(Of StockSymbol, Date)(Me)
			Else
				.Records = New LinkedHashSet(Of Record, Date)
				.RecordsDaily = New LinkedHashSet(Of RecordDaily, Date)
				.SplitFactors = New LinkedHashSet(Of SplitFactor, Date)
				.StockErrors = New LinkedHashSet(Of StockError, Date)
				.StockSymbols = New LinkedHashSet(Of StockSymbol, Date)
			End If

			MyRecordQuoteValues = New LinkedHashSet(Of RecordQuoteValue, Date)
		End With
		If MyListHeaderInfo Is Nothing Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
			MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		End If
		Me.IsSplitEnabled = True
	End Sub

	Public Sub New(ByRef Parent As YahooAccessData.Report, ByVal Symbol As String, ByVal Name As String)
		Me.New(Symbol, Name, "")
		With Me
			.Report = Parent
			If .Report IsNot Nothing Then
				.ReportID = .Report.ID
				.Report.Stocks.Add(Me)
			End If
		End With
	End Sub

	Public Sub New(ByRef Parent As Report, ByRef Stream As Stream, Optional ByVal IsRecordVirtual As Boolean = False)
		Me.New(Stream, IsRecordVirtual)
		With Me
			.Report = Parent
			.ReportID = .Report.ID
			.SerializeLoadFrom(Stream, IsRecordVirtual)
			.Sector = .Report.Sectors.ToSearch.Find(.SectorID)
			.Industry = .Report.Industries.ToSearch.Find(.IndustryID)
			If .Sector IsNot Nothing Then
				If .Industry IsNot Nothing Then
					If .Sector.Industries.ToSearch.Find(.Industry.ID) Is Nothing Then
						.Sector.Industries.Add(.Industry)
					End If
				End If
				.Sector.Stocks.Add(Me)
			End If
			If .Industry IsNot Nothing Then
				.Industry.Stocks.Add(Me)
			End If
			.Report.Stocks.Add(Me)
		End With
	End Sub

	Public Sub New(
		ByRef Parent As Report,
		ByRef Sector As Sector,
		ByRef Industry As Industry,
		ByVal Symbol As String)

		Me.New(Symbol, "", "")
		With Me
			.Report = Parent
			If .Report IsNot Nothing Then
				.ReportID = .Report.ID
				.Report.Stocks.Add(Me)
			End If
			.Sector = Sector
			If .Sector IsNot Nothing Then
				.SectorID = .Sector.ID
				.Sector.Stocks.Add(Me)
			End If
			.Industry = Industry
			If .Industry IsNot Nothing Then
				.IndustryID = .Industry.ID
				.Industry.Stocks.Add(Me)
			End If
		End With
	End Sub

	Public Sub New(ByVal Symbol As String, ByVal Name As String)
		Me.New(Symbol, Name, "")
	End Sub

	Public Sub New(ByRef Parent As YahooAccessData.Report, ByVal Symbol As String)
		Me.New(Symbol, "", "")
		With Me
			.Report = Parent
			If .Report IsNot Nothing Then
				.ReportID = .Report.ID
				.Report.Stocks.Add(Me)
			End If
		End With
	End Sub

	Public Sub New(ByVal Symbol As String)
		Me.New(Symbol, "", "")
	End Sub

	Public Sub New(ByRef Stream As Stream, ByVal IsRecordVirtual As Boolean)
		Me.New("", "", "", Stream, IsRecordVirtual)
	End Sub

	Public Sub New()
		Me.New("", "", "", Nothing, False)
	End Sub
#End Region
#Region "Main Control Function"
	Public Function WebRefreshRecord(ByVal RecordDateStop As Date) As Date

		Dim ThisTaskOfWebRefreshRecord = New Task(Of Date)(
		 Function()
			 Dim ThisTask = WebRefreshRecordAsync(Now)
			 Return ThisTask.Result
		 End Function)

		ThisTaskOfWebRefreshRecord.Start()
		ThisTaskOfWebRefreshRecord.Wait()

		Debug.Print($"WebRefreshRecord: {Me.Symbol}")
		Return ThisTaskOfWebRefreshRecord.Result
	End Function



	''' <summary>
	'''   Use to refresh the data record via the buid-in web interface
	''' </summary>
	''' <param name="RecordDateStop"></param>
	''' <returns>
	'''   return the date for the last record after the update.
	''' </returns>
	Public Async Function WebRefreshRecordAsync(ByVal RecordDateStop As Date) As Task(Of Date)
		Dim ThisWebDataSource = Me.Report.WebDataSource
		If ThisWebDataSource Is Nothing Then
			Return Me.DateStop
		End If
		'always remove the automatic splitting adjustment
		'when connected to teh web.
		'the data is already adjusted to reflect the share splitting
		Me.IsSplitEnabled = False
		Dim ThisWebEodStockDescriptor As IWebEodDescriptor = New WebEODData.WebStockDescriptor(Me)
		Dim ThisExchangeCode = ThisWebEodStockDescriptor.ExchangeCode
		Dim ThisLastTradingDay = ThisWebDataSource.DayTimeOfLastTrading(ThisExchangeCode)
		If RecordDateStop > ThisLastTradingDay Then
			RecordDateStop = ThisLastTradingDay
		End If
		If Me.DateStop < RecordDateStop Then
			'try the web update
			'get just the data that is needed for an update
			Dim ThisWebDateStart As Date
			'note this function work at very low level and should not call the .records interface directly
			'to make sure there is no issue of stack overflow call the object _Records instead
			If _Records.Count = 0 Then
				ThisWebDateStart = Me.DateStart
				'make sure there is nothing in the MyRecordQuoteValues
				MyRecordQuoteValues.Clear()
			Else
				ThisWebDateStart = Me.DateStop.Date.AddDays(1)
			End If
			Dim ThisResponseQuery = Await ThisWebDataSource.LoadStockQuoteAsync(
				ThisWebEodStockDescriptor.ExchangeCode,
				ThisWebEodStockDescriptor.SymbolCode,
				DateStart:=Me.DateStop.Date,
				RecordDateStop.Date)

			If ThisResponseQuery.IsSuccess Then
				Dim ThisDictionaryOfStockQuote = ThisResponseQuery.Result
				If ThisDictionaryOfStockQuote.Count > 0 Then
					'only one stock at a time and it content is element 0 of the dictionary
					Dim ThisListOfStockQuote As List(Of WebEODData.IStockQuote) = ThisDictionaryOfStockQuote.Values.First
					If ThisListOfStockQuote.Count > 0 Then
						'new data record availaible for this stock 
						For Each ThisRecord In ThisListOfStockQuote.ToListOfRecord
							ThisRecord.Stock = Me
							ThisRecord.StockID = ThisRecord.Stock.ID
							If _Records.Count > 0 Then
								If ThisRecord.DateDay > Me.DateStop Then
									'work diectly with the collection
									_Records.Add(ThisRecord)
									MyRecordQuoteValues.Add(New RecordQuoteValue(ThisRecord))
									Me.DateStop = ThisRecord.DateDay
								End If
							Else
								'special case for the first record
								_Records.Add(ThisRecord)
								MyRecordQuoteValues.Add(New RecordQuoteValue(ThisRecord))
								Me.DateStop = ThisRecord.DateDay
							End If
						Next
					End If
					If Me.DateStop > Me.Report.DateStop Then
						Me.Report.DateStop = Me.DateStop
					End If
				End If
			End If
		End If
		Return Me.DateStop
	End Function




	''' <summary>
	''' Use to modify the current stock with new parameters 
	''' </summary>
	''' <param name="Stock">the new stock parameter</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Friend Function Add(ByRef Stock As Stock) As Boolean
		Dim ThisStockSymbol As YahooAccessData.StockSymbol
		Dim ThisStockSymbolFromList As YahooAccessData.StockSymbol
		Dim ThisSplitFactor As YahooAccessData.SplitFactor
		Dim ThisSplitFactorFromList As YahooAccessData.SplitFactor

		With Stock
			Me.Name = .Name
			Me.IsOption = .IsOption
			'make sure we can add only newer data
			If Me.DateStart = Me.DateStop Then
				'first time the date are initialized
				Me.DateStart = .DateStart
				Me.DateStop = .DateStop
			End If
			If .DateStart < Me.DateStart Then
				Me.DateStart = .DateStart
			End If
			If .DateStop > Me.DateStop Then
				Me.DateStop = .DateStop
			End If
			Me.IsSymbolError = .IsSymbolError
			Me.RankGain = .RankGain
			Me.Exchange = .Exchange
			Me.ErrorDescription = .ErrorDescription
			.StockErrors.CopyDeep(Me, IsIgnoreID:=True)
			For Each ThisStockSymbol In .StockSymbols
				ThisStockSymbolFromList = Me.StockSymbols.ToSearch.Find(ThisStockSymbol.KeyValue)
				If ThisStockSymbolFromList Is Nothing Then
					'add to the current list
					ThisStockSymbol.CopyDeep(Me, IsIgnoreID:=True)
				Else
					'check if there are the same symbol and update if needed
					With ThisStockSymbolFromList
						If (.Symbol = ThisStockSymbol.Symbol) And (.SymbolNew = ThisStockSymbol.SymbolNew) Then
							.DateUpdate = ThisStockSymbol.DateUpdate
							.Exception = ThisStockSymbol.Exception
							.Exchange = ThisStockSymbol.Exchange
							.Name = ThisStockSymbol.Name
						Else
							Throw New Exception(String.Format("Error: Multiple key value in object {0}:{1}:{2} ...", ThisStockSymbol.Symbol, ThisStockSymbol.ToString, ThisStockSymbol.KeyValue))
						End If
					End With
				End If
			Next
			For Each ThisSplitFactor In .SplitFactors
				ThisSplitFactorFromList = Me.SplitFactors.ToSearch.Find(ThisSplitFactor.KeyValue)
				If ThisSplitFactorFromList Is Nothing Then
					'add to the current list
					ThisSplitFactor.CopyDeep(Me, IsIgnoreID:=True)
				Else
					'check if there are the same symbol and update if needed
					With ThisSplitFactorFromList
						If .Ratio = ThisSplitFactor.Ratio Then
							.Exception = ThisSplitFactor.Exception
							.SharesLast = ThisSplitFactor.SharesLast
							.SharesNew = ThisSplitFactor.SharesNew
						Else
							Throw New Exception(String.Format("Error: Multiple key value in object {0}:{1}:{2} ...", Stock.Symbol, ThisSplitFactor.ToString, ThisSplitFactor.KeyValue))
						End If
					End With
				End If
			Next
			.Records.CopyDeep(Me, IsIgnoreID:=True)
			.RecordsDaily.CopyDeep(Me, IsIgnoreID:=True)
		End With
		Return True
	End Function

	Friend Sub RefreshDate()
		If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
			_StockErrors.Clear()
			_StockSymbols.Clear()
			_SplitFactors.Clear()
			_Records.Clear()
			If CType(_Records, IDataVirtual).Enabled = False Then
				ISystemEventOfRecord_Load()
			End If
		End If
	End Sub

	Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As Stock
		Dim ThisStock As Stock
		If Parent.FileType = IMemoryStream.enuFileType.RecordIndexed Then
			ThisStock = New Stock(Nothing, IsRecordVirtual:=True)
		Else
			ThisStock = New Stock(Nothing, IsRecordVirtual:=False)
		End If
		With ThisStock
			If IsIgnoreID = False Then .ID = Me.ID
			.Symbol = Me.Symbol
			.Name = Me.Name
			.IsOption = Me.IsOption
			.DateStart = Me.DateStart
			.DateStop = Me.DateStop
			.IsSymbolError = Me.IsSymbolError
			.RankGain = Me.RankGain
			.Exchange = Me.Exchange
			.ErrorDescription = Me.ErrorDescription
			.Report = Parent
			.ReportID = .Report.ID
			.Report.Stocks.Add(ThisStock)
			If Me.Sector IsNot Nothing Then
				.Sector = .Report.Sectors.ToSearch.Find(Me.Sector.KeyValue)
			End If
			If .Sector Is Nothing Then
				.Exception = New Exception("Invalid sector...", .Exception)
			Else
				.SectorID = .Sector.ID
				.Sector.Stocks.Add(ThisStock)
			End If
			If Me.Industry IsNot Nothing Then
				.Industry = .Report.Industries.ToSearch.Find(Me.Industry.KeyValue)
			End If
			If .Industry Is Nothing Then
				.Exception = New Exception("Invalid industry...", .Exception)
			Else
				.IndustryID = .Industry.ID
				.Industry.Stocks.Add(ThisStock)
			End If
			'ThisStopWatch.Restart()
			Me.Records.CopyDeep(ThisStock, IsIgnoreID)
			Me.StockErrors.CopyDeep(ThisStock, IsIgnoreID)
			Me.StockSymbols.CopyDeep(ThisStock, IsIgnoreID)
			Me.SplitFactors.CopyDeep(ThisStock, IsIgnoreID)
			'ThisStopWatch.Stop()
		End With
		Return ThisStock
	End Function

	Public Async Function FileSaveAsync(ByVal Stream As Stream) As Task(Of Boolean)
		'save to the attached file in append mode
		Await SerializeSaveToAsync(Stream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed, DirectCast(Stream, FileStream).Name())
		Return True
	End Function

	Public Async Function FileSaveAsync(ByVal Stream As Stream, ByVal StreamBaseName As String) As Task(Of Boolean)
		'save to the attached file in append mode
		Await SerializeSaveToAsync(Stream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed, StreamBaseName)
		Return True
	End Function

	Public Async Function FileSaveAsync(ByVal StreamBaseName As String, Optional ByVal IsSingleThread As Boolean = False) As Task(Of MemoryStream)
		'save to the attached file in append mode

		Try
			Dim ThisMemoryStream As MemoryStream = New MemoryStream
			Await SerializeSaveToAsync(ThisMemoryStream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed, StreamBaseName, IsSingleThread)
			Return ThisMemoryStream
		Catch ex As Exception
			Return Nothing
		End Try
	End Function

	Public Async Function FileSaveWatchAsync(ByVal StreamBaseName As String, Optional ByVal IsSingleThread As Boolean = False) As Task(Of MemoryStreamWatch)
		'save to the attached file in append mode
		Dim ThisTimer As New Stopwatch
		Try
			Dim ThisMemoryStream As MemoryStreamWatch = New MemoryStreamWatch With {.KeyValue = Me.KeyValue}
			ThisTimer.Restart()
			Await SerializeSaveToAsync(ThisMemoryStream, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed, StreamBaseName, IsSingleThread)
			ThisTimer.Stop()
			ThisMemoryStream.ToListOfProcessWatch.Add(New ProcessTimeMeasurement(Me.KeyValue, "Total Time", ThisTimer.ElapsedMilliseconds))
			Return ThisMemoryStream
		Catch ex As Exception
			Return Nothing
		End Try
	End Function

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
			'If MyException IsNot Nothing Then
			'	If Me.Report IsNot Nothing Then
			'		Me.Report.Exception = New Exception(Me.ToString, MyException)
			'	End If
			'End If
		End Set
	End Property

	Public Async Function RecordQuoteValuesAsync() As Task(Of IEnumerable(Of RecordQuoteValue))
		Dim ThisTask As Task(Of IEnumerable(Of RecordQuoteValue))

		ThisTask = Me.RecordQuoteValuesAsync(Me.Report.TimeFormat)
		Await ThisTask
		Return ThisTask.Result
	End Function

	''' <summary>
	''' </summary>
	''' <param name="TimeFormat"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Async Function RecordQuoteValuesAsync(ByVal TimeFormat As YahooAccessData.Report.enuTimeFormat) As Task(Of IEnumerable(Of RecordQuoteValue))
		Dim ThisTaskOfRecordQuoteValue As Task(Of IEnumerable(Of RecordQuoteValue))
		Dim ThisTick As Integer

		ThisTaskOfRecordQuoteValue = New Task(Of IEnumerable(Of RecordQuoteValue))(
			Function()
				Dim ThisRecord As IEnumerable(Of RecordQuoteValue) = Nothing
				Dim IsLoaded As Boolean
				'Dim IsThisTheLastRecordReadingRequest As Boolean
				'this command load the data in memory only one thread at the time

				SyncLock MySyncLockForStockRecordLoadingAll
					'SyncLock MySyncLockForRecordLoading
					With DirectCast(_Records, IDataVirtual)
						IsLoaded = .IsLoaded
						'If MyLastTimeOfRecordQuoteValuesAsync = ThisTick Then
						'  IsThisTheLastRecordReadingRequest = True
						'Else
						'  IsThisTheLastRecordReadingRequest = False
						'End If
						If IsLoaded = False Then
							'always read one at the time to support FTP web reading
							Try
								.Load()
								IsLoaded = .IsLoaded
								ThisRecord = Me.RecordQuoteValues(TimeFormat)

								'    ThisRecord = Me.RecordQuoteValues(TimeFormat)
								'If IsThisTheLastRecordReadingRequest Then
								'  'only load the most recent request
								'  .Load()
								'  IsLoaded = .IsLoaded
								'  If IsLoaded Then
								'    ThisRecord = Me.RecordQuoteValues(TimeFormat)
								'  Else
								'    'the loading have been cancelled
								'    ThisRecord = Nothing
								'  End If
								'Else
								'  ThisRecord = Nothing
								'End If
							Catch ex As Exception
								'how to fix because this may be a problem on a thread:seem to work
								Me.Report.Exception = ex
								ThisRecord = Nothing
							End Try
						Else
							ThisRecord = Me.RecordQuoteValues(TimeFormat)
						End If
					End With
					'End SyncLock
				End SyncLock
				Return ThisRecord
			End Function)

		SyncLock MySyncLockForRecordLoading
			ThisTick = Environment.TickCount
			MyLastTimeOfRecordQuoteValuesAsync = ThisTick
			With DirectCast(Me.Report, IStockRecordEvent)
				Dim IsLoading As Boolean = .IsLoading
				If IsLoading Then
					'cancel the current loading unless it is the current symbol
					If .SymbolLoading <> Me.Symbol Then
						'not sure yet about this
						'cancel the loading
						.LoadCancel(.SymbolLoading)
					End If
				End If
			End With
			ThisTaskOfRecordQuoteValue.Start()
		End SyncLock
		Await ThisTaskOfRecordQuoteValue
		Return ThisTaskOfRecordQuoteValue.Result
	End Function

	''' <summary>
	''' Force a complete release of all data of the stock object including the 
	''' record, SplitFactors, StockErrors and StockSymbols change.
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function RecordReleaseAll() As Boolean
		SyncLock MySyncLockForStockRecordLoadingAll
			DirectCast(_Records, IDataVirtual).Release()
			'do not release the splitfactor 
			'the number of splitfactor is small and has little impast on memory
			'in addition it is important to keep the splitfactor data intact
			'that is loaded only one from a file at the beginning of the program
			'DirectCast(_SplitFactors, IDataVirtual).Release()
			DirectCast(_StockErrors, IDataVirtual).Release()
			DirectCast(_StockSymbols, IDataVirtual).Release()
			DirectCast(_RecordsDaily, IDataVirtual).Release()
		End SyncLock
		Return True
	End Function

	Public Function IsRecordLoaded() As Boolean
		Dim IsLoaded As Boolean
		SyncLock MySyncLockForStockRecordLoadingAll
			With DirectCast(_Records, IDataVirtual)
				IsLoaded = .IsLoaded
			End With
		End SyncLock
		Return IsLoaded
	End Function

	Public Function RecordRelease() As Boolean
		Dim IsLoaded As Boolean
		SyncLock MySyncLockForStockRecordLoadingAll
			With DirectCast(_Records, IDataVirtual)
				IsLoaded = .IsLoaded
				If IsLoaded = True Then
					Try
						.Release()
						IsLoaded = .IsLoaded
					Catch ex As Exception
						'how to fix because this may be a problem on a thread
						Me.Report.Exception = ex
						IsLoaded = False
					End Try
				End If
			End With
		End SyncLock
		Return IsLoaded
	End Function

	Public Function RecordLoad() As Boolean
		Dim IsLoaded As Boolean
		SyncLock MySyncLockForStockRecordLoadingAll
			If Me.Report.WebDataSource Is Nothing Then
				With DirectCast(_Records, IDataVirtual)
					IsLoaded = .IsLoaded
					If IsLoaded = False Then
						'always read one at the time to support FTP web reading
						Try
							.Load()
							IsLoaded = .IsLoaded
						Catch ex As Exception
							'how to fix because this may be a problem on a thread
							Me.Report.Exception = ex
							IsLoaded = False
						End Try
					End If
				End With
			Else
				Me.WebRefreshRecord(Now)
				IsLoaded = True
			End If
		End SyncLock
		Return IsLoaded
	End Function

	Public Async Function RecordLoadAsync() As Task(Of Boolean)
		Dim ThisTaskOfRecordLoad As Task(Of Boolean)
		Dim ThisTick As Integer

		ThisTaskOfRecordLoad = New Task(Of Boolean)(
			Function()
				Dim ThisResult As Boolean = False
				Dim IsLoaded As Boolean
				Dim IsThisTheLastRecordReadingRequest As Boolean
				'this command load the data in memory only one thread at the time

				SyncLock MySyncLockForStockRecordLoadingAll
					With DirectCast(_Records, IDataVirtual)
						IsLoaded = .IsLoaded
						If MyLastTimeOfRecordLoadAsync = ThisTick Then
							IsThisTheLastRecordReadingRequest = True
						Else
							IsThisTheLastRecordReadingRequest = False
						End If
						If IsLoaded = False Then
							'always read one at the time to support FTP web reading
							Try
								If IsThisTheLastRecordReadingRequest Then
									'only load the most recent request
									.Load()
									IsLoaded = .IsLoaded
									If IsLoaded Then
										ThisResult = True
									Else
										'the loading have been cancelled
										ThisResult = False
									End If
								Else
									ThisResult = False
								End If
							Catch ex As Exception
								'how to fix because this may be a problem on a thread:seem to work
								Me.Report.Exception = ex
								ThisResult = False
							End Try
						Else
							ThisResult = True
						End If
					End With
				End SyncLock
				Return ThisResult
			End Function)

		SyncLock MySyncLockForRecordLoading
			ThisTick = Environment.TickCount
			MyLastTimeOfRecordLoadAsync = ThisTick
			ThisTaskOfRecordLoad.Start()
		End SyncLock
		Await ThisTaskOfRecordLoad
		Return ThisTaskOfRecordLoad.Result
	End Function


	Public ReadOnly Property RecordQuoteValues As IEnumerable(Of RecordQuoteValue)
		Get
			Return Me.RecordQuoteValues(Me.Report.TimeFormat)
		End Get
	End Property

	Public ReadOnly Property RecordQuoteValues(ByVal TimeFormat As YahooAccessData.Report.enuTimeFormat) As IEnumerable(Of RecordQuoteValue)
		Get
			If Me.RecordLoad Then
				Select Case TimeFormat
					Case YahooAccessData.Report.enuTimeFormat.Sample
						Return MyRecordQuoteValues
					Case YahooAccessData.Report.enuTimeFormat.Daily
						Return MyRecordQuoteValues.ToDaily
					Case YahooAccessData.Report.enuTimeFormat.Weekly
						'not implemented yet
						Throw New NotImplementedException("Weekly time data is not yet implemented at this level...")
						'Return MyRecordQuoteValues.ToDaily
					Case Else
						Return MyRecordQuoteValues
				End Select
			Else
				Return Nothing
			End If
		End Get
	End Property

	Public Overrides Function ToString() As String
		Dim ThisRecordsCountTotal As Integer = DirectCast(_Records, IRecordInfo).CountTotal
		Dim ThisRecordsDailyCountTotal As Integer = DirectCast(_RecordsDaily, IRecordInfo).CountTotal

		Return String.Format("{0},ID:{1},Key:{2},Record:{3} of {4},RecordDaily:{5} of {6}", TypeName(Me), Me.KeyID, Me.KeyValue.ToString, Me.Records.Count, ThisRecordsCountTotal, Me.RecordsDaily.Count, ThisRecordsDailyCountTotal)
	End Function
#End Region
#Region "IStockGain"
	Public Property RankGain As Double Implements IStockRank.RankGain
#End Region
#Region "Properties"
	Public Property FiscalYearEnd As Date
		Get
			Return MyFiscalYearEnd
		End Get
		Set(value As Date)
			MyFiscalYearEnd = value
		End Set
	End Property

	Public Property FiscalYearMeasured As Date
		Get
			Return MyFiscalYearEndMeasured
		End Get
		Set(value As Date)
			MyFiscalYearEndMeasured = value
		End Set
	End Property
#End Region
#Region "Collection Definition"
	Public Overridable Property Records As ICollection(Of Record)
		Get
			If Me.Report.WebDataSource Is Nothing Then
				Me.RecordLoad()
				Return _Records
			Else
				Me.WebRefreshRecord(Now)
				Return _Records
			End If
		End Get
		Set(value As ICollection(Of Record))
			_Records = value
		End Set
	End Property

	Public Overridable Property Records(ByVal IsLoadEnabled As Boolean) As ICollection(Of Record)
		Get
			If IsLoadEnabled Then Me.RecordLoad()
			Return _Records
		End Get
		Set(value As ICollection(Of Record))
			_Records = value
		End Set
	End Property

	Public Overridable Property RecordsDaily As ICollection(Of RecordDaily)
		Get
			'synclock ensure that we load the record only from one thread at a time
			SyncLock MySyncLockForRecordDailyLoading
				If TypeOf _Records Is IDataVirtual Then
					With DirectCast(_Records, IDataVirtual)
						Dim IsLoaded As Boolean = .IsLoaded
						If IsLoaded = False Then
							Try
								.Load()
							Catch ex As Exception
								Me.Report.Exception = ex
							End Try
						End If
					End With
				End If
			End SyncLock
			Return _RecordsDaily
		End Get
		Set(value As ICollection(Of RecordDaily))
			_RecordsDaily = value
		End Set
	End Property

	Public Overridable Property RecordsDaily(ByVal IsLoadEnabled As Boolean) As ICollection(Of RecordDaily)
		Get
			'synclock ensure that we load the record only from one thread at a time
			SyncLock MySyncLockForRecordDailyLoading
				If IsLoadEnabled Then
					If TypeOf _Records Is IDataVirtual Then
						With DirectCast(_Records, IDataVirtual)
							Dim IsLoaded As Boolean = .IsLoaded
							If IsLoaded = False Then
								Try
									.Load()
								Catch ex As Exception
									Me.Report.Exception = ex
								End Try
							End If
						End With
					End If
				End If
			End SyncLock
			Return _RecordsDaily
		End Get
		Set(value As ICollection(Of RecordDaily))
			_RecordsDaily = value
		End Set
	End Property


	Public Overridable Property SplitFactors As ICollection(Of SplitFactor)
		Get
			If TypeOf _SplitFactors Is IDataVirtual Then
				With DirectCast(_SplitFactors, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			Return _SplitFactors
		End Get
		Set(value As ICollection(Of SplitFactor))
			_SplitFactors = value
		End Set
	End Property
	Public Overridable Property SplitFactors(ByVal IsLoadEnabled As Boolean) As ICollection(Of SplitFactor)
		Get
			If IsLoadEnabled Then
				If TypeOf _SplitFactors Is IDataVirtual Then
					With DirectCast(_SplitFactors, IDataVirtual)
						Dim IsLoaded As Boolean = .IsLoaded
						If IsLoaded = False Then
							Try
								.Load()
							Catch ex As Exception
								Me.Report.Exception = ex
							End Try
						End If
					End With
				End If
			End If
			Return _SplitFactors
		End Get
		Set(value As ICollection(Of SplitFactor))
			_SplitFactors = value
		End Set
	End Property
	Public Overridable Property StockErrors As ICollection(Of StockError)
		Get
			If TypeOf _StockErrors Is IDataVirtual Then
				With DirectCast(_StockErrors, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			Return _StockErrors
		End Get
		Set(value As ICollection(Of StockError))
			_StockErrors = value
		End Set
	End Property
	Public Overridable Property StockErrors(ByVal IsLoadEnabled As Boolean) As ICollection(Of StockError)
		Get
			If IsLoadEnabled Then
				If TypeOf _StockErrors Is IDataVirtual Then
					With DirectCast(_StockErrors, IDataVirtual)
						Dim IsLoaded As Boolean = .IsLoaded
						If IsLoaded = False Then
							Try
								.Load()
							Catch ex As Exception
								Me.Report.Exception = ex
							End Try
						End If
					End With
				End If
			End If
			Return _StockErrors
		End Get
		Set(value As ICollection(Of StockError))
			_StockErrors = value
		End Set
	End Property
	Public Overridable Property StockSymbols As ICollection(Of StockSymbol)
		Get
			If TypeOf _StockSymbols Is IDataVirtual Then
				With DirectCast(_StockSymbols, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			Return _StockSymbols
		End Get
		Set(value As ICollection(Of StockSymbol))
			_StockSymbols = value
		End Set
	End Property
	Public Overridable Property StockSymbols(ByVal IsLoadEnabled As Boolean) As ICollection(Of StockSymbol)
		Get
			If IsLoadEnabled Then
				If TypeOf _StockSymbols Is IDataVirtual Then
					With DirectCast(_StockSymbols, IDataVirtual)
						Dim IsLoaded As Boolean = .IsLoaded
						If IsLoaded = False Then
							Try
								.Load()
							Catch ex As Exception
								Me.Report.Exception = ex
							End Try
						End If
					End With
				End If
			End If
			Return _StockSymbols
		End Get
		Set(value As ICollection(Of StockSymbol))
			_StockSymbols = value
		End Set
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
	Public Function CompareTo(other As Stock) As Integer Implements System.IComparable(Of Stock).CompareTo
		Return Me.KeyValue.CompareTo(other.KeyValue)
	End Function
#End Region
#Region "Register Key"
	Public Property KeyID As Integer Implements IRegisterKey(Of String).KeyID
		Get
			Return Me.ID
		End Get
		Set(value As Integer)
			Me.ID = value
		End Set
	End Property

	Public Property KeyValue As String Implements IRegisterKey(Of String).KeyValue
		Get
			Return Me.Symbol.ToString
		End Get
		Set(value As String)
			Me.Symbol = value
		End Set
	End Property
#End Region
#Region "IMemoryStream"
	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
		Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
	End Sub

	Private Async Function SerializeSaveToAsync(
			ByVal Stream As Stream,
			ByVal FileType As IMemoryStream.enuFileType,
			ByVal StreamBaseName As String,
			Optional ByVal IsSingleThread As Boolean = False) As Task(Of Boolean)

		Dim ThisStream As Stream = Stream
		Dim ThisBinaryWriter As New BinaryWriter(ThisStream)
		Dim ThisListException As List(Of Exception)
		Dim ThisException As Exception
		Dim ThisStreamPositionWhereRecordsEnd As Long
		Dim ThisPosition As Long
		Dim ThisTaskForRecords As Task(Of RecordIndex(Of Record, Date)) = Nothing
		Dim ThisTaskForRecordsEndOfDay As Task(Of RecordIndex(Of Record, Date)) = Nothing
		Dim ThisTaskForRecordsDaily As Task(Of RecordIndex(Of RecordDaily, Date)) = Nothing
		Dim ThisTaskForStockErrors As Task(Of RecordIndex(Of StockError, Date)) = Nothing
		Dim ThisTaskForStockSymbols As Task(Of RecordIndex(Of StockSymbol, Date)) = Nothing
		Dim ThisTaskForSplitFactors As Task(Of RecordIndex(Of SplitFactor, Date)) = Nothing
		Dim ThisRecordsCopy As ICollection(Of Record)
		Dim ThisRecordIndexOfRecord As RecordIndex(Of Record, Date)

		'store the count because saving on a thread may clear the value before we have time to used it
		Dim ThisRecordCount As Integer = 0
		Dim ThisRecordDailyCount As Integer = 0
		Dim ThisStockErrorCount As Integer = 0
		Dim ThisStockSymbolCount As Integer = 0
		Dim ThisSplitFactorCount As Integer = 0
		Dim ThisListOfProcessWatch As IList(Of IProcessTimeMeasurement)

		If TypeOf (Stream) Is MemoryStreamWatch Then
			ThisListOfProcessWatch = DirectCast(Stream, MemoryStreamWatch).ToListOfProcessWatch
		Else
			ThisListOfProcessWatch = Nothing
		End If

		If (FileType = IMemoryStream.enuFileType.RecordIndexed) Then
			'create and start the task of saving to file
			SyncLock MySyncLockForRecordLoading
				If TypeOf _Records Is IDataVirtual Then
					'make sure the data has been loaded before we save
					With DirectCast(_Records, IDataVirtual)
						Dim IsLoaded As Boolean = .IsLoaded
						If IsLoaded = False Then
							Try
								.Load()
							Catch ex As Exception
								Me.Report.Exception = ex
							End Try
						End If
					End With
				End If
			End SyncLock
			If _Records.Count > 0 Then
				ThisRecordCount = _Records.Count
				'start the process of saving 

				ThisTaskForRecords = CreateTaskForRecordIndexSave(Of Record, Date)(
					StreamBaseName,
					FileMode.Open,
					Me.Report.AsDateRange,
					"Stock\Record",
					"_" & Me.Symbol,
					".rec",
					_Records)

				If IsSingleThread Then
					Await ThisTaskForRecords
				End If

				'Dim ThisFilePathRecordNameForEndOfDay = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Me.Report.FileName), "Stock\RecordEndOfDay")
				'If My.Computer.FileSystem.DirectoryExists(ThisFilePathRecordNameForEndOfDay) Then
				If Me.Report.IsFileReadEndOfDay Then
					'save also the only the record end of day data (EOD) in a different directory 
					'first make a new copy of the data because the collection is not thread safe
					ThisRecordsCopy = New List(Of Record)(_Records)

					ThisTaskForRecordsEndOfDay = CreateTaskForRecordIndexSave(Of Record, Date)(
						StreamBaseName,
						FileMode.Open,
						Me.Report.AsDateRange,
						"Stock\RecordEndOfDay",
						"_" & Me.Symbol,
						".rec",
						ThisRecordsCopy,
						IsSaveAtEndOfDay:=True)

					If IsSingleThread Then
						Await ThisTaskForRecordsEndOfDay
					End If
				End If
			End If
			'create the task
			If TypeOf _RecordsDaily Is IDataVirtual Then
				'make sure the data has been virtually loaded before we save
				With DirectCast(_RecordsDaily, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			If _RecordsDaily.Count > 0 Then
				ThisRecordDailyCount = _RecordsDaily.Count
				'start the process of saving 
				ThisTaskForRecordsDaily = CreateTaskForRecordIndexSave(Of RecordDaily, Date)(
					StreamBaseName,
					FileMode.Open,
					Me.Report.AsDateRange,
					"Stock\Record",
					"_" & Me.Symbol,
					".hist.rec",
					_RecordsDaily)

				If IsSingleThread Then
					Await ThisTaskForRecordsDaily
				End If
			End If
			If TypeOf _StockErrors Is IDataVirtual Then
				'make sure the data has been loaded before we save
				With DirectCast(_StockErrors, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			If _StockErrors.Count > 0 Then
				ThisStockErrorCount = _StockErrors.Count
				ThisTaskForStockErrors = CreateTaskForRecordIndexSave(Of StockError, Date)(
					StreamBaseName,
					FileMode.Open,
					Me.Report.AsDateRange,
					"Stock\StockError",
					"_" & Me.Symbol,
					".ser",
					_StockErrors)

				If IsSingleThread Then
					Await ThisTaskForStockErrors
				End If
			End If
			If TypeOf _StockSymbols Is IDataVirtual Then
				'make sure the data has been loaded before we save
				With DirectCast(_StockSymbols, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			If _StockSymbols.Count > 0 Then
				ThisStockSymbolCount = _StockSymbols.Count
				ThisTaskForStockSymbols = CreateTaskForRecordIndexSave(Of StockSymbol, Date)(
					StreamBaseName,
					FileMode.Open,
					Me.Report.AsDateRange,
					"Stock\StockSymbol",
					"_" & Me.Symbol,
					".ssy",
					_StockSymbols)

				If IsSingleThread Then
					Await ThisTaskForStockSymbols
				End If
			End If
			If TypeOf _SplitFactors Is IDataVirtual Then
				'make sure the data has been loaded before we save
				With DirectCast(_SplitFactors, IDataVirtual)
					Dim IsLoaded As Boolean = .IsLoaded
					If IsLoaded = False Then
						Try
							.Load()
						Catch ex As Exception
							Me.Report.Exception = ex
						End Try
					End If
				End With
			End If
			If _SplitFactors.Count > 0 Then
				ThisSplitFactorCount = _SplitFactors.Count
				ThisTaskForSplitFactors = CreateTaskForRecordIndexSave(Of SplitFactor, Date)(
					StreamBaseName,
					FileMode.Open,
					Me.Report.AsDateRange,
					"Stock\SplitFactor",
					"_" & Me.Symbol,
					".sfa",
					_SplitFactors)

				If IsSingleThread Then
					Await ThisTaskForSplitFactors
				End If
			End If
		End If
		'all the needed tasks are started start writing to the main file
		With ThisBinaryWriter
			.Write(VERSION_MEMORY_STREAM)
			.Write(Me.ID)
			ThisListException = Me.Exception.ToList
			ThisListException.Reverse()
			.Write(ThisListException.Count)
			For Each ThisException In ThisListException
				.Write(ThisException.Message)
			Next
			.Write(Me.Symbol)
			.Write(Me.Name)
			.Write(Me.IndustryID)
			.Write(Me.SectorID)
			.Write(Me.IsOption)
			.Write(Me.DateStart.ToBinary)
			.Write(Me.DateStop.ToBinary)
			.Write(Me.IsSymbolError)
			.Write(Me.ReportID)
			.Write(Me.Exchange)
			.Write(Me.ErrorDescription)

			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.StockErrors.Count)
					For Each ThisStockError In Me.StockErrors
						ThisStockError.SerializeSaveTo(ThisStream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If ThisStockErrorCount > 0 Then
						Await ThisTaskForStockErrors
						Dim ThisRecordIndexOfStockError As RecordIndex(Of StockError, Date)
						ThisRecordIndexOfStockError = ThisTaskForStockErrors.Result
						.Write(ThisRecordIndexOfStockError.FileCount)
						.Write(ThisRecordIndexOfStockError.MaxID)
						.Write(ThisRecordIndexOfStockError.DateStart.ToBinary)
						.Write(ThisRecordIndexOfStockError.DateStop.ToBinary)
						With DirectCast(_StockErrors, IDateUpdate)
							.DateStart = ThisRecordIndexOfStockError.DateStart
							.DateStop = ThisRecordIndexOfStockError.DateStop
						End With
						With DirectCast(_StockErrors, IRecordInfo)
							.CountTotal = ThisRecordIndexOfStockError.FileCount
							.MaximumID = ThisRecordIndexOfStockError.MaxID
						End With
						ThisRecordIndexOfStockError.Dispose()
						ThisRecordIndexOfStockError = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_StockErrors, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_StockErrors, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select

			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.StockSymbols.Count)
					For Each ThisStockSymbol In Me.StockSymbols
						ThisStockSymbol.SerializeSaveTo(ThisStream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If ThisStockSymbolCount > 0 Then
						Await ThisTaskForStockSymbols
						Dim ThisRecordIndexOfStockSymbol = ThisTaskForStockSymbols.Result
						.Write(ThisRecordIndexOfStockSymbol.FileCount)
						.Write(ThisRecordIndexOfStockSymbol.MaxID)
						.Write(ThisRecordIndexOfStockSymbol.DateStart.ToBinary)
						.Write(ThisRecordIndexOfStockSymbol.DateStop.ToBinary)
						With DirectCast(_StockSymbols, IDateUpdate)
							.DateStart = ThisRecordIndexOfStockSymbol.DateStart
							.DateStop = ThisRecordIndexOfStockSymbol.DateStop
						End With
						With DirectCast(_StockSymbols, IRecordInfo)
							.CountTotal = ThisRecordIndexOfStockSymbol.FileCount
							.MaximumID = ThisRecordIndexOfStockSymbol.MaxID
						End With
						ThisRecordIndexOfStockSymbol.Dispose()
						ThisRecordIndexOfStockSymbol = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_StockSymbols, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_StockSymbols, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select
			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.SplitFactors.Count)
					For Each ThisSplitFactor In Me.SplitFactors
						ThisSplitFactor.SerializeSaveTo(ThisStream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If ThisSplitFactorCount > 0 Then
						Await ThisTaskForSplitFactors
						Dim ThisRecordIndexOfSplitFactor = ThisTaskForSplitFactors.Result
						.Write(ThisRecordIndexOfSplitFactor.FileCount)
						.Write(ThisRecordIndexOfSplitFactor.MaxID)
						.Write(ThisRecordIndexOfSplitFactor.DateStart.ToBinary)
						.Write(ThisRecordIndexOfSplitFactor.DateStop.ToBinary)
						With DirectCast(_SplitFactors, IDateUpdate)
							.DateStart = ThisRecordIndexOfSplitFactor.DateStart
							.DateStop = ThisRecordIndexOfSplitFactor.DateStop
						End With
						With DirectCast(_SplitFactors, IRecordInfo)
							.CountTotal = ThisRecordIndexOfSplitFactor.FileCount
							.MaximumID = ThisRecordIndexOfSplitFactor.MaxID
						End With
						ThisRecordIndexOfSplitFactor.Dispose()
						ThisRecordIndexOfSplitFactor = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_SplitFactors, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_SplitFactors, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select
			ThisStreamPositionWhereRecordsEnd = ThisStream.Position
			ThisPosition = -1    'by default
			.Write(ThisPosition)
			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.Records.Count)
					For Each ThisRecord In Me.Records
						ThisRecord.SerializeSaveTo(ThisStream, FileType)
					Next
					'save the recordsdaily
					.Write(Me.RecordsDaily.Count)
					For Each ThisRecordDaily In Me.RecordsDaily
						ThisRecordDaily.SerializeSaveTo(ThisStream, FileType)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If ThisRecordCount > 0 Then
						If ThisTaskForRecordsEndOfDay IsNot Nothing Then
							Await Task.WhenAll(ThisTaskForRecords, ThisTaskForRecordsEndOfDay)
						Else
							Await ThisTaskForRecords
						End If
						ThisRecordIndexOfRecord = ThisTaskForRecords.Result
						If ThisRecordIndexOfRecord.Exception Is Nothing Then
							.Write(ThisRecordIndexOfRecord.FileCount)
							.Write(ThisRecordIndexOfRecord.MaxID)
							.Write(ThisRecordIndexOfRecord.DateStart.ToBinary)
							.Write(ThisRecordIndexOfRecord.DateStop.ToBinary)
							With DirectCast(_Records, IDateUpdate)
								.DateStart = ThisRecordIndexOfRecord.DateStart
								.DateStop = ThisRecordIndexOfRecord.DateStop
							End With
							With DirectCast(_Records, IRecordInfo)
								.CountTotal = ThisRecordIndexOfRecord.FileCount
								.MaximumID = ThisRecordIndexOfRecord.MaxID
							End With
						Else
							Me.Report.RaiseMessage(String.Format("File saving error with {0}: {1}", Me.Symbol, Me.Exception.Message), IMessageInfoEvents.EnuMessageType.Warning)
							Dim ThisRecordInfo = DirectCast(_Records, IRecordInfo)
							Dim ThisDateUpdate = DirectCast(_Records, IDateUpdate)
							.Write(ThisRecordInfo.CountTotal)
							.Write(ThisRecordInfo.MaximumID)
							.Write(ThisDateUpdate.DateStart.ToBinary)
							.Write(ThisDateUpdate.DateStop.ToBinary)
						End If
						ThisRecordIndexOfRecord.Dispose()
						ThisRecordIndexOfRecord = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_Records, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_Records, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
					'continue with the RecordsDaily
					If ThisRecordDailyCount > 0 Then
						Await ThisTaskForRecordsDaily
						MyRecordIndexOfRecordDaily = ThisTaskForRecordsDaily.Result
						.Write(MyRecordIndexOfRecordDaily.FileCount)
						.Write(MyRecordIndexOfRecordDaily.MaxID)
						.Write(MyRecordIndexOfRecordDaily.DateStart.ToBinary)
						.Write(MyRecordIndexOfRecordDaily.DateStop.ToBinary)
						With DirectCast(_RecordsDaily, IDateUpdate)
							.DateStart = MyRecordIndexOfRecordDaily.DateStart
							.DateStop = MyRecordIndexOfRecordDaily.DateStop
						End With
						With DirectCast(_RecordsDaily, IRecordInfo)
							.CountTotal = MyRecordIndexOfRecordDaily.FileCount
							.MaximumID = MyRecordIndexOfRecordDaily.MaxID
						End With
						MyRecordIndexOfRecordDaily.Dispose()
						MyRecordIndexOfRecordDaily = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_RecordsDaily, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_RecordsDaily, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
					Try
						'rewrite the next ThisStream position value
						ThisPosition = ThisStream.Position
						ThisStream.Position = ThisStreamPositionWhereRecordsEnd
						.Write(ThisPosition)
						're-store the position to the end
						ThisStream.Position = ThisPosition
					Catch ex As Exception
						Throw New Exception("Writing next record position file failure", ex)
					End Try
			End Select
		End With
		Return True
	End Function

	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
		Dim ThisBinaryWriter As New BinaryWriter(Stream)
		Dim ThisListException As List(Of Exception)
		Dim ThisException As Exception
		Dim ThisStreamPositionWhereRecordsEnd As Long
		Dim ThisPosition As Long

		With ThisBinaryWriter
			.Write(VERSION_MEMORY_STREAM)
			.Write(Me.ID)
			ThisListException = Me.Exception.ToList
			ThisListException.Reverse()
			.Write(ThisListException.Count)
			For Each ThisException In ThisListException
				.Write(ThisException.Message)
			Next
			.Write(Me.Symbol)
			.Write(Me.Name)
			.Write(Me.IndustryID)
			.Write(Me.SectorID)
			.Write(Me.IsOption)
			.Write(Me.DateStart.ToBinary)
			.Write(Me.DateStop.ToBinary)
			.Write(Me.IsSymbolError)
			.Write(Me.ReportID)
			.Write(Me.Exchange)
			.Write(Me.ErrorDescription)

			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.StockErrors.Count)
					For Each ThisStockError In Me.StockErrors
						ThisStockError.SerializeSaveTo(Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If TypeOf _StockErrors Is IDataVirtual Then
						'make sure the data has been loaded before we save
						With DirectCast(_StockErrors, IDataVirtual)
							Dim IsLoaded As Boolean = .IsLoaded
							If IsLoaded = False Then
								Try
									.Load()
								Catch ex As Exception
									Me.Report.Exception = ex
									Exit Sub
								End Try
							End If
						End With
					End If
					If _StockErrors.Count > 0 Then
						Dim ThisRecordIndexOfStockError = New RecordIndex(Of StockError, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\StockError", "_" & Me.Symbol, ".ser")
						ThisRecordIndexOfStockError.Save(_StockErrors)
						If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
							With DirectCast(_StockErrors, IDataVirtual)
								If .Enabled Then
									.Release()
								End If
							End With
						End If
						.Write(ThisRecordIndexOfStockError.FileCount)
						.Write(ThisRecordIndexOfStockError.MaxID)
						.Write(ThisRecordIndexOfStockError.DateStart.ToBinary)
						.Write(ThisRecordIndexOfStockError.DateStop.ToBinary)
						With DirectCast(_StockErrors, IDateUpdate)
							.DateStart = ThisRecordIndexOfStockError.DateStart
							.DateStop = ThisRecordIndexOfStockError.DateStop
						End With
						With DirectCast(_StockErrors, IRecordInfo)
							.CountTotal = ThisRecordIndexOfStockError.FileCount
							.MaximumID = ThisRecordIndexOfStockError.MaxID
						End With
						ThisRecordIndexOfStockError.Dispose()
						ThisRecordIndexOfStockError = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_StockErrors, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_StockErrors, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select

			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.StockSymbols.Count)
					For Each ThisStockSymbol In Me.StockSymbols
						ThisStockSymbol.SerializeSaveTo(Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If TypeOf _StockSymbols Is IDataVirtual Then
						'make sure the data has been loaded before we save
						With DirectCast(_StockSymbols, IDataVirtual)
							Dim IsLoaded As Boolean = .IsLoaded
							If IsLoaded = False Then
								Try
									.Load()
								Catch ex As Exception
									Me.Report.Exception = ex
									Exit Sub
								End Try
							End If
						End With
					End If
					If _StockSymbols.Count > 0 Then
						Dim ThisRecordIndexOfStockSymbol = New RecordIndex(Of StockSymbol, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\StockSymbol", "_" & Me.Symbol, ".ssy")
						ThisRecordIndexOfStockSymbol.Save(_StockSymbols)
						If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
							With DirectCast(_StockSymbols, IDataVirtual)
								If .Enabled Then
									.Release()
								End If
							End With
						End If
						.Write(ThisRecordIndexOfStockSymbol.FileCount)
						.Write(ThisRecordIndexOfStockSymbol.MaxID)
						.Write(ThisRecordIndexOfStockSymbol.DateStart.ToBinary)
						.Write(ThisRecordIndexOfStockSymbol.DateStop.ToBinary)
						With DirectCast(_StockSymbols, IDateUpdate)
							.DateStart = ThisRecordIndexOfStockSymbol.DateStart
							.DateStop = ThisRecordIndexOfStockSymbol.DateStop
						End With
						With DirectCast(_StockSymbols, IRecordInfo)
							.CountTotal = ThisRecordIndexOfStockSymbol.FileCount
							.MaximumID = ThisRecordIndexOfStockSymbol.MaxID
						End With
						ThisRecordIndexOfStockSymbol.Dispose()
						ThisRecordIndexOfStockSymbol = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_StockSymbols, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_StockSymbols, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select
			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.SplitFactors.Count)
					For Each ThisSplitFactor In Me.SplitFactors
						ThisSplitFactor.SerializeSaveTo(Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					If TypeOf _SplitFactors Is IDataVirtual Then
						'make sure the data has been loaded before we save
						With DirectCast(_SplitFactors, IDataVirtual)
							Dim IsLoaded As Boolean = .IsLoaded
							If IsLoaded = False Then
								Try
									.Load()
								Catch ex As Exception
									Me.Report.Exception = ex
									Exit Sub
								End Try
							End If
						End With
					End If
					If _SplitFactors.Count > 0 Then
						Dim ThisRecordIndexOfSplitFactor = New RecordIndex(Of SplitFactor, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\SplitFactor", "_" & Me.Symbol, ".sfa")
						ThisRecordIndexOfSplitFactor.Save(_SplitFactors)
						If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
							With DirectCast(_SplitFactors, IDataVirtual)
								If .Enabled Then
									.Release()
								End If
							End With
						End If
						.Write(ThisRecordIndexOfSplitFactor.FileCount)
						.Write(ThisRecordIndexOfSplitFactor.MaxID)
						.Write(ThisRecordIndexOfSplitFactor.DateStart.ToBinary)
						.Write(ThisRecordIndexOfSplitFactor.DateStop.ToBinary)
						With DirectCast(_SplitFactors, IDateUpdate)
							.DateStart = ThisRecordIndexOfSplitFactor.DateStart
							.DateStop = ThisRecordIndexOfSplitFactor.DateStop
						End With
						With DirectCast(_SplitFactors, IRecordInfo)
							.CountTotal = ThisRecordIndexOfSplitFactor.FileCount
							.MaximumID = ThisRecordIndexOfSplitFactor.MaxID
						End With
						ThisRecordIndexOfSplitFactor.Dispose()
						ThisRecordIndexOfSplitFactor = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_SplitFactors, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_SplitFactors, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
			End Select
			ThisStreamPositionWhereRecordsEnd = Stream.Position
			ThisPosition = -1    'by default
			.Write(ThisPosition)
			Select Case FileType
				Case IMemoryStream.enuFileType.Standard
					.Write(Me.Records.Count)
					For Each ThisRecord In Me.Records
						ThisRecord.SerializeSaveTo(Stream, FileType)
					Next
					'save the recordsdaily
					.Write(Me.RecordsDaily.Count)
					For Each ThisRecordDaily In Me.RecordsDaily
						ThisRecordDaily.SerializeSaveTo(Stream, FileType)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					SyncLock MySyncLockForRecordLoading
						If TypeOf _Records Is IDataVirtual Then
							'make sure the data has been loaded before we save
							With DirectCast(_Records, IDataVirtual)
								Dim IsLoaded As Boolean = .IsLoaded
								If IsLoaded = False Then
									Try
										.Load()
									Catch ex As Exception
										Me.Report.Exception = ex
										Exit Sub
									End Try
								End If
							End With
						End If
					End SyncLock
					If _Records.Count > 0 Then
						Dim ThisRecordIndexOfRecord As RecordIndex(Of Record, Date)
						Try
							ThisRecordIndexOfRecord = New RecordIndex(Of Record, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\Record", "_" & Me.Symbol, ".rec")
						Catch ex As Exception
							Me.Report.RaiseMessage(String.Format("File saving error with stock {0}:{1}", Me.Symbol, ex.Message), IMessageInfoEvents.EnuMessageType.Warning)
							ThisRecordIndexOfRecord = Nothing
						End Try
						If ThisRecordIndexOfRecord IsNot Nothing Then
							ThisRecordIndexOfRecord.Save(_Records)
							If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
								With DirectCast(_Records, IDataVirtual)
									If .Enabled Then
										.Release()
									End If
								End With
							End If
							.Write(ThisRecordIndexOfRecord.FileCount)
							.Write(ThisRecordIndexOfRecord.MaxID)
							.Write(ThisRecordIndexOfRecord.DateStart.ToBinary)
							.Write(ThisRecordIndexOfRecord.DateStop.ToBinary)
							With DirectCast(_Records, IDateUpdate)
								.DateStart = ThisRecordIndexOfRecord.DateStart
								.DateStop = ThisRecordIndexOfRecord.DateStop
							End With
							With DirectCast(_Records, IRecordInfo)
								.CountTotal = ThisRecordIndexOfRecord.FileCount
								.MaximumID = ThisRecordIndexOfRecord.MaxID
							End With
							ThisRecordIndexOfRecord.Dispose()
							ThisRecordIndexOfRecord = Nothing
						Else
							Dim ThisRecordInfo = DirectCast(_Records, IRecordInfo)
							Dim ThisDateUpdate = DirectCast(_Records, IDateUpdate)
							.Write(ThisRecordInfo.CountTotal)
							.Write(ThisRecordInfo.MaximumID)
							.Write(ThisDateUpdate.DateStart.ToBinary)
							.Write(ThisDateUpdate.DateStop.ToBinary)
						End If
					Else
						Dim ThisRecordInfo = DirectCast(_Records, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_Records, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
					'save the RecordsDaily
					If TypeOf _RecordsDaily Is IDataVirtual Then
						'make sure the data has been loaded before we save
						With DirectCast(_RecordsDaily, IDataVirtual)
							Dim IsLoaded As Boolean = .IsLoaded
							If IsLoaded = False Then
								Try
									.Load()
								Catch ex As Exception
									Me.Report.Exception = ex
									Exit Sub
								End Try
							End If
						End With
					End If
					If _RecordsDaily.Count > 0 Then
						If MyRecordIndexOfRecordDaily Is Nothing Then
							MyRecordIndexOfRecordDaily = New RecordIndex(Of RecordDaily, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\Record", "_" & Me.Symbol, ".hist.rec")
						End If
						MyRecordIndexOfRecordDaily.Save(_RecordsDaily)
						If Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed Then
							With DirectCast(_RecordsDaily, IDataVirtual)
								If .Enabled Then
									.Release()
								End If
							End With
						End If
						.Write(MyRecordIndexOfRecordDaily.FileCount)
						.Write(MyRecordIndexOfRecordDaily.MaxID)
						.Write(MyRecordIndexOfRecordDaily.DateStart.ToBinary)
						.Write(MyRecordIndexOfRecordDaily.DateStop.ToBinary)
						With DirectCast(_RecordsDaily, IDateUpdate)
							.DateStart = MyRecordIndexOfRecordDaily.DateStart
							.DateStop = MyRecordIndexOfRecordDaily.DateStop
						End With
						With DirectCast(_RecordsDaily, IRecordInfo)
							.CountTotal = MyRecordIndexOfRecordDaily.FileCount
							.MaximumID = MyRecordIndexOfRecordDaily.MaxID
						End With
						MyRecordIndexOfRecordDaily.Dispose()
						MyRecordIndexOfRecordDaily = Nothing
					Else
						Dim ThisRecordInfo = DirectCast(_RecordsDaily, IRecordInfo)
						Dim ThisDateUpdate = DirectCast(_RecordsDaily, IDateUpdate)
						.Write(ThisRecordInfo.CountTotal)
						.Write(ThisRecordInfo.MaximumID)
						.Write(ThisDateUpdate.DateStart.ToBinary)
						.Write(ThisDateUpdate.DateStop.ToBinary)
					End If
					Try
						'rewrite the next stream position value
						ThisPosition = Stream.Position
						Stream.Position = ThisStreamPositionWhereRecordsEnd
						.Write(ThisPosition)
						're-store the position to the end
						Stream.Position = ThisPosition
					Catch ex As Exception
						Throw New Exception("Writing next record position file failure", ex)
					End Try
			End Select
		End With
	End Sub

	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
		Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
		Dim ThisVersion As Single
		Dim Count As Integer
		Dim ThisFileNumberRecordsDaily As Integer
		Dim ThisPosition As Long = 0
		Dim ThisMaxID As Integer
		Dim ThisDateStart As Date
		Dim ThisDateStop As Date

		With ThisBinaryReader
			ThisVersion = .ReadSingle
			Me.ID = .ReadInt32
			Me.Exception = Nothing
			Dim I As Integer

			For I = 1 To .ReadInt32
				MyException = New Exception(.ReadString, MyException)
			Next
			Me.Symbol = .ReadString
			Me.Name = .ReadString
			Me.IndustryID = .ReadInt32
			Me.SectorID = .ReadInt32
			Me.IsOption = .ReadBoolean
			Me.DateStart = DateTime.FromBinary(.ReadInt64)
			Me.DateStop = DateTime.FromBinary(.ReadInt64)
			Me.IsSymbolError = .ReadBoolean
			Me.ReportID = .ReadInt32
			Me.Exchange = .ReadString
			Me.ErrorDescription = .ReadString
			Count = .ReadInt32
			Select Case Me.Report.FileType
				Case IMemoryStream.enuFileType.Standard
					If Count > 4 Then
						'CType(Me.StockErrors, LinkedHashSet(Of StockError, Date)).Capacity = Count
					End If
					For I = 1 To Count
						Dim ThisStockError As New StockError(Me, Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					ThisMaxID = .ReadInt32
					ThisDateStart = DateTime.FromBinary(.ReadInt64)
					ThisDateStop = DateTime.FromBinary(.ReadInt64)
					With CType(_StockErrors, IRecordInfo)
						.CountTotal = Count
						.MaximumID = ThisMaxID
					End With
					With DirectCast(_StockErrors, IDateUpdate)
						.DateStart = ThisDateStart
						.DateStop = ThisDateStop
					End With
			End Select

			Count = .ReadInt32
			Select Case Me.Report.FileType
				Case IMemoryStream.enuFileType.Standard
					If Count > 4 Then
						'CType(Me.StockSymbols, LinkedHashSet(Of StockSymbol, Date)).Capacity = Count
					End If
					For I = 1 To Count
						Dim ThisStockSymbol As New StockSymbol(Me, Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					ThisMaxID = .ReadInt32
					ThisDateStart = DateTime.FromBinary(.ReadInt64)
					ThisDateStop = DateTime.FromBinary(.ReadInt64)
					With CType(_StockSymbols, IRecordInfo)
						.CountTotal = Count
						.MaximumID = ThisMaxID
					End With
					With DirectCast(_StockSymbols, IDateUpdate)
						.DateStart = ThisDateStart
						.DateStop = ThisDateStop
					End With
			End Select

			Count = .ReadInt32
			Select Case Me.Report.FileType
				Case IMemoryStream.enuFileType.Standard
					If Count > 4 Then
						'CType(Me.SplitFactors, LinkedHashSet(Of SplitFactor, Date)).Capacity = Count
					End If
					For I = 1 To Count
						Dim ThisSplitFactor As New SplitFactor(Me, Stream)
					Next
				Case IMemoryStream.enuFileType.RecordIndexed
					ThisMaxID = .ReadInt32
					ThisDateStart = DateTime.FromBinary(.ReadInt64)
					ThisDateStop = DateTime.FromBinary(.ReadInt64)
					With CType(_SplitFactors, IRecordInfo)
						.CountTotal = Count
						.MaximumID = ThisMaxID
					End With
					With DirectCast(_SplitFactors, IDateUpdate)
						.DateStart = ThisDateStart
						.DateStop = ThisDateStop
					End With
			End Select
			MyStreamPositionWhereRecordStop = .ReadInt64()
			MyFileNumberRecord = .ReadInt32
			Select Case Me.Report.FileType
				Case IMemoryStream.enuFileType.Standard
					'load the record data immediately calling the system event interface
					MyStreamPositionWhereRecordStart = Stream.Position
					'load all the data record in the file
					If MyFileNumberRecord > 4 Then
						'CType(Me.Records, LinkedHashSet(Of Record, Date)).Capacity = MyFileNumberRecord
					End If
					For I = 1 To MyFileNumberRecord
						Dim ThisRecord = New Record(Me, Stream)
					Next
					If MyFileNumberRecord > 0 Then
						With _Records.First
							If .DateLastTrade <> YahooAccessData.ReportDate.DateNullValue Then
								If .DateLastTrade < Me.DateStart Then
									Me.DateStart = .DateLastTrade
								End If
							End If
						End With
						With _Records.Last
							If .DateLastTrade <> YahooAccessData.ReportDate.DateNullValue Then
								If .DateLastTrade > Me.DateStop Then
									Me.DateStop = .DateLastTrade
								End If
							End If
						End With
					End If
					If ThisVersion >= VERSION_MEMORY_STREAM_FOR_RECORDS_DAILY Then
						'load all the data record in the file
						ThisFileNumberRecordsDaily = .ReadInt32
						If ThisFileNumberRecordsDaily > 4 Then
							'CType(Me.RecordsDaily, LinkedHashSet(Of RecordDaily, Date)).Capacity = ThisFileNumberRecordsDaily
						End If
						For I = 1 To ThisFileNumberRecordsDaily
							Dim ThisRecordDaily = New RecordDaily(Me, Stream)
						Next
						If ThisFileNumberRecordsDaily > 0 Then
							With _RecordsDaily.First
								If .DateDay <> YahooAccessData.ReportDate.DateNullValue Then
									If .DateDay < Me.DateStart Then
										Me.DateStart = .DateDay
									End If
								End If
							End With
							With _RecordsDaily.Last
								If .DateDay <> YahooAccessData.ReportDate.DateNullValue Then
									If .DateDay > Me.DateStop Then
										Me.DateStop = .DateDay
									End If
								End If
							End With
						End If
					End If
				Case IMemoryStream.enuFileType.RecordIndexed
					'RecordIndexed file are always virtual
					ThisMaxID = .ReadInt32
					ThisDateStart = DateTime.FromBinary(.ReadInt64)
					ThisDateStop = DateTime.FromBinary(.ReadInt64)
					With CType(_Records, IRecordInfo)
						.CountTotal = MyFileNumberRecord
						.MaximumID = ThisMaxID
					End With
					With _Records.AsDateUpdate
						.DateStart = ThisDateStart
						.DateStop = ThisDateStop
					End With
					If ThisVersion >= VERSION_MEMORY_STREAM_FOR_RECORDS_DAILY Then
						Dim ThisFileNumberRecordDaily As Integer = .ReadInt32
						ThisMaxID = .ReadInt32
						ThisDateStart = DateTime.FromBinary(.ReadInt64)
						ThisDateStop = DateTime.FromBinary(.ReadInt64)
						With CType(_RecordsDaily, IRecordInfo)
							.CountTotal = ThisFileNumberRecordDaily
							.MaximumID = ThisMaxID
						End With
						With _RecordsDaily.AsDateUpdate
							.DateStart = ThisDateStart
							.DateStop = ThisDateStop
						End With
					End If
			End Select
		End With
		ThisBinaryReader.Dispose()
	End Sub

	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
		Call SerializeLoadFrom(Stream, False)
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
#Region "IFormatData"
	Public Function ToStockBasicInfo() As IStockBasicInfo
		Dim ThisStockBasicInfo = New StockBasicInfo(Me)
		Return ThisStockBasicInfo
	End Function

	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of Stock)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

	Public Shared Function ListOfHeader() As List(Of HeaderInfo)
		Dim ThisListHeaderInfo As New List(Of HeaderInfo)
		With ThisListHeaderInfo
			.Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "Exchange", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "Symbol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "IsOption", .Title = .Name, .Format = "{0}", .Visible = False, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "DateStart", .Title = .Name, .Format = "{0:g}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "DateStop", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "RankGain", .Title = .Name, .Format = "{0:n2}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "IsSymbolError", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "ErrorDescription", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "SectorName", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
			.Add(New HeaderInfo With {.Name = "IndustryName", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
		End With
		Return ThisListHeaderInfo
	End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As Stock, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .IndustryID <> Me.IndustryID Then Return False
				If .SectorID <> Me.SectorID Then Return False
				If .ReportID <> Me.ReportID Then Return False
			End If
			If .Symbol <> Me.Symbol Then Return False
			If .Name <> Me.Name Then Return False
			If .IsOption <> Me.IsOption Then Return False
			If .DateStart <> Me.DateStart Then Return False
			If .DateStop <> Me.DateStop Then Return False
			If .IsSymbolError <> Me.IsSymbolError Then Return False
			If .Exchange <> Me.Exchange Then Return False
			If .ErrorDescription <> Me.ErrorDescription Then Return False
			If Me.Records.EqualsDeep(.Records, IsIgnoreID) = False Then Return False
			If Me.RecordsDaily.EqualsDeep(.RecordsDaily, IsIgnoreID) = False Then Return False
			If Me.SplitFactors.EqualsDeep(.SplitFactors, IsIgnoreID) = False Then Return False
			If Me.StockErrors.EqualsDeep(.StockErrors, IsIgnoreID) = False Then Return False
			If Me.StockSymbols.EqualsDeep(.StockSymbols, IsIgnoreID) = False Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As Stock) As Boolean Implements IEquatable(Of Stock).Equals
		If other Is Nothing Then Return False
		If Me.Symbol = other.Symbol Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If (TypeOf obj Is Stock) Then
			Return Me.Equals(DirectCast(obj, Stock))
		Else
			Return (False)
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		Return Me.Symbol.GetHashCode()
	End Function
#End Region
#Region "ISystemEventOfSplitFactor"
	Private Sub ISystemEvent_Add(item As SplitFactor) Implements ISystemEvent(Of SplitFactor).Add

	End Sub

	Private Sub ISystemEvent_Clear() Implements ISystemEvent(Of SplitFactor).Clear

	End Sub

	Private Sub ISystemEvent_Load() Implements ISystemEvent(Of SplitFactor).Load
		Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me.Report, IDateRange))

		If ThisTimePeriodOverlap.IsOverlap(DirectCast(_SplitFactors, IDateUpdate)) Then
			Dim ThisRecordIndexOfSplitFactor = New RecordIndex(Of SplitFactor, Date)(DirectCast(Me.Report.ToStream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\SplitFactor", "_" & Me.Symbol, ".sfa")
			If ThisRecordIndexOfSplitFactor.Exception Is Nothing Then
				With DirectCast(_SplitFactors, IDateUpdate)
					.DateStart = ThisRecordIndexOfSplitFactor.DateStart
					.DateStop = ThisRecordIndexOfSplitFactor.DateStop
				End With
				With CType(_SplitFactors, IRecordInfo)
					.CountTotal = ThisRecordIndexOfSplitFactor.FileCount
					.MaximumID = ThisRecordIndexOfSplitFactor.MaxID
				End With
				If ThisRecordIndexOfSplitFactor.ToListPosition.Count > 4 Then
					'CType(_SplitFactors, LinkedHashSet(Of Record, Date)).Capacity = ThisRecordIndexOfSplitFactor.ToListPosition.Count
				End If
				For Each ThisPosition In ThisRecordIndexOfSplitFactor.ToListPosition
					Try
						Dim ThisSplitFactor = New SplitFactor(Me, ThisRecordIndexOfSplitFactor.BaseStream(ThisPosition))
					Catch ex As Exception
						ThisRecordIndexOfSplitFactor.Dispose()
						ThisRecordIndexOfSplitFactor = Nothing
						Throw New Exception(String.Format("Record SplitFactor reading error with stock {0}", Me.Symbol))
					End Try
				Next
			Else
				Me.Report.RaiseMessage(String.Format("File SplitFactor reading error with stock {0}:{1}", Me.Symbol, ThisRecordIndexOfSplitFactor.Exception.Message), IMessageInfoEvents.EnuMessageType.Warning)
			End If
			ThisRecordIndexOfSplitFactor.Dispose()
			ThisRecordIndexOfSplitFactor = Nothing
		End If
	End Sub

	Private Function ISystemEvent_Remove(item As SplitFactor) As Boolean Implements ISystemEvent(Of SplitFactor).Remove

	End Function
#End Region
#Region "ISystemEventOfStockSymbol"
	Private Sub ISystemEventOfStockSymbol_Add(item As StockSymbol) Implements ISystemEvent(Of StockSymbol).Add

	End Sub

	Private Sub ISystemEventOfStockSymbol_Clear() Implements ISystemEvent(Of StockSymbol).Clear

	End Sub

	Private Sub ISystemEventOfStockSymbol_Load() Implements ISystemEvent(Of StockSymbol).Load
		Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me.Report, IDateRange))

		If ThisTimePeriodOverlap.IsOverlap(DirectCast(_StockSymbols, IDateUpdate)) Then
			Dim ThisRecordIndexOfStockSymbol = New RecordIndex(Of StockSymbol, Date)(DirectCast(Me.Report.ToStream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\StockSymbol", "_" & Me.Symbol, ".ssy")

			If ThisRecordIndexOfStockSymbol.Exception Is Nothing Then
				With DirectCast(_StockSymbols, IDateUpdate)
					.DateStart = ThisRecordIndexOfStockSymbol.DateStart
					.DateStop = ThisRecordIndexOfStockSymbol.DateStop
				End With
				With CType(_StockSymbols, IRecordInfo)
					.CountTotal = ThisRecordIndexOfStockSymbol.FileCount
					.MaximumID = ThisRecordIndexOfStockSymbol.MaxID
				End With
				If ThisRecordIndexOfStockSymbol.ToListPosition.Count > 4 Then
					'CType(_StockSymbols, LinkedHashSet(Of StockSymbol, Date)).Capacity = ThisRecordIndexOfStockSymbol.ToListPosition.Count
				End If
				For Each ThisPosition In ThisRecordIndexOfStockSymbol.ToListPosition
					Try
						Dim ThisStockSymbol = New StockSymbol(Me, ThisRecordIndexOfStockSymbol.BaseStream(ThisPosition))
					Catch ex As Exception
						ThisRecordIndexOfStockSymbol.Dispose()
						ThisRecordIndexOfStockSymbol = Nothing
						Throw New Exception(String.Format("Record StockSymbol reading error with stock {0}", Me.Symbol))
					End Try
				Next
			Else
				Me.Report.RaiseMessage(String.Format("File StockSymbol reading error with stock {0}:{1}", Me.Symbol, ThisRecordIndexOfStockSymbol.Exception.Message), IMessageInfoEvents.EnuMessageType.Warning)
			End If
			ThisRecordIndexOfStockSymbol.Dispose()
			ThisRecordIndexOfStockSymbol = Nothing
		End If
	End Sub

	Private Function ISystemEventOfStockSymbol_Remove(item As StockSymbol) As Boolean Implements ISystemEvent(Of StockSymbol).Remove

	End Function
#End Region
#Region "ISystemEventOfStockError"
	Private Sub ISystemEventOfStockError_Add(item As StockError) Implements ISystemEvent(Of StockError).Add

	End Sub

	Private Sub ISystemEventOfStockError_Clear() Implements ISystemEvent(Of StockError).Clear

	End Sub

	Private Sub ISystemEventOfStockError_Load() Implements ISystemEvent(Of StockError).Load
		Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me.Report, IDateRange))

		If ThisTimePeriodOverlap.IsOverlap(DirectCast(_StockErrors, IDateUpdate)) Then
			Dim ThisRecordIndexOfStockError = New RecordIndex(Of StockError, Date)(DirectCast(Me.Report.ToStream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\StockError", "_" & Me.Symbol, ".ser")
			If ThisRecordIndexOfStockError.Exception Is Nothing Then
				With DirectCast(_StockErrors, IDateUpdate)
					.DateStart = ThisRecordIndexOfStockError.DateStart
					.DateStop = ThisRecordIndexOfStockError.DateStop
				End With
				With CType(_StockErrors, IRecordInfo)
					.CountTotal = ThisRecordIndexOfStockError.FileCount
					.MaximumID = ThisRecordIndexOfStockError.MaxID
				End With
				If ThisRecordIndexOfStockError.ToListPosition.Count > 4 Then
					'CType(_StockErrors, LinkedHashSet(Of StockError, Date)).Capacity = ThisRecordIndexOfStockError.ToListPosition.Count
				End If
				For Each ThisPosition In ThisRecordIndexOfStockError.ToListPosition
					Try
						Dim ThisStockError = New StockError(Me, ThisRecordIndexOfStockError.BaseStream(ThisPosition))
					Catch ex As Exception
						ThisRecordIndexOfStockError.Dispose()
						ThisRecordIndexOfStockError = Nothing
						Throw New Exception(String.Format("StockError reading error with stock {0}", Me.Symbol))
					End Try
				Next
				ThisRecordIndexOfStockError.Dispose()
				ThisRecordIndexOfStockError = Nothing
			Else
				Me.Report.RaiseMessage(String.Format("File StockError reading error with stock {0}:{1}", Me.Symbol, ThisRecordIndexOfStockError.Exception.Message), IMessageInfoEvents.EnuMessageType.Warning)
			End If
		End If
	End Sub

	Private Function ISystemEventOfStockError_Remove(item As StockError) As Boolean Implements ISystemEvent(Of StockError).Remove

	End Function
#End Region
#Region "ISystemEventOfRecord"
	Private Sub ISystemEventOfRecord_Add(item As Record) Implements ISystemEvent(Of Record).Add
		MyRecordQuoteValues.Add(New RecordQuoteValue(item))
	End Sub

	Private Sub ISystemEventOfRecord_Clear() Implements ISystemEvent(Of Record).Clear
		MyRecordQuoteValues.Clear()
	End Sub

	Private Function ISystemEventOfRecord_Remove(item As Record) As Boolean Implements ISystemEvent(Of Record).Remove
		With MyRecordQuoteValues
			Return .Remove(.ToSearch.Find(item.KeyValue))
		End With
	End Function

	Private Sub ISystemEventOfRecord_Load() Implements ISystemEvent(Of Record).Load
		If Me.Report Is Nothing Then Exit Sub
		Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me.Report, IDateRange))

		'SyncLock MySyncLockForRecordLoading
		Select Case Me.Report.AsRecordControlInfo.ControlType
			Case enuStockRecordLoadType.FixedCount
				If ThisTimePeriodOverlap.IsOverlap(DirectCast(_Records, IDateUpdate)) Then
					'Task.Delay(2000).Wait()   'just for testing a web delay
					Dim ThisChildPath As String
					If Me.Report.IsFileReadEndOfDayEnabled Then
						ThisChildPath = "Stock\RecordEndOfDay"
					Else
						ThisChildPath = "Stock\Record"
					End If
					Dim ThisRecordIndexOfRecord As RecordIndex(Of Record, Date)
					ThisRecordIndexOfRecord = New RecordIndex(Of Record, Date)(
						DirectCast(Me.Report.ToStream, FileStream).Name(),
						FileMode.Open,
						Me.Report.AsDateRange,
						ThisChildPath,
						"_" & Me.Symbol,
						".rec",
						Me.Report.IsReadOnly)

					If ThisRecordIndexOfRecord.Exception Is Nothing Then
						If Me.Report.AsRecordControlInfo.Enabled Then
							DirectCast(Me.Report, IStockRecordEvent).LoadBefore(Me.Symbol)
						End If
						With DirectCast(_Records, IDateUpdate)
							.DateStart = ThisRecordIndexOfRecord.DateStart
							.DateStop = ThisRecordIndexOfRecord.DateStop
						End With
						With CType(_Records, IRecordInfo)
							.CountTotal = ThisRecordIndexOfRecord.FileCount
							.MaximumID = ThisRecordIndexOfRecord.MaxID
						End With
						Dim IsLoadCancelled As Boolean = False
						If ThisRecordIndexOfRecord.ToListPosition.Count > 4 Then
							'CType(_Records, LinkedHashSet(Of Record, Date)).Capacity = .ToListPosition.Count
						End If
						'Dim ThisRecordLast As Record = Nothing
						For Each ThisPosition In ThisRecordIndexOfRecord.ToListPosition
							Try
								Dim ThisRecord = New Record(Me, ThisRecordIndexOfRecord.BaseStream(ThisPosition))
								'If ThisRecordLast IsNot Nothing Then
								'  If ThisRecord.QuoteValues(0).ExDividendDate < ThisRecordLast.QuoteValues(0).ExDividendDate Then
								'    ThisRecordLast = ThisRecordLast
								'  End If
								'End If
								'ThisRecordLast = ThisRecord
							Catch ex As Exception
								ThisRecordIndexOfRecord.Dispose()
								ThisRecordIndexOfRecord = Nothing
								Throw New Exception(String.Format("Record reading error with stock {0}", Me.Symbol))
							End Try
							IsLoadCancelled = DirectCast(Me.Report, IStockRecordEvent).IsLoadCancel
							If IsLoadCancelled Then
								'request to cancel the loading is activated
								'clear the data currently accumulated in the records
								DirectCast(_Records, IDataVirtual).Release()
								Exit For
							End If
						Next
						If Me.Report.AsRecordControlInfo.Enabled Then
							If IsLoadCancelled = False Then
								DirectCast(Me.Report, IStockRecordEvent).LoadAfter(Me.Symbol)
							End If
						End If
					Else
						Me.Report.RaiseMessage(String.Format("File record reading error with stock {0}:{1}", Me.Symbol, ThisRecordIndexOfRecord.Exception.Message), IMessageInfoEvents.EnuMessageType.Warning)
					End If
					ThisRecordIndexOfRecord.Dispose()
					ThisRecordIndexOfRecord = Nothing
				End If
			Case enuStockRecordLoadType.WebEodHistorical
				'Dim ThisTaskRun = New Task(Of Date)(
				'  Function()
				'    Return WebRefreshRecordAsync(Now)
				'  End Function)
				'ThisTaskRun.Start()
				'Await ThisTaskRun
				'Dim ThisDate = ThisTaskRun.Result.Result
				'Dim ThisDate = Await WebRefreshRecordAsync(Now)
				'Dim ThisDate = Await WebRefreshRecordAsync(Now)
				'Dim ThisnumberOfRecord = Me.Records.Count
				'this interface will not work well since it is only fired one and only if the record has nort been loaded
				'it does not expect to update every time we request a read
		End Select
	End Sub
#End Region
#Region "ISystemEventOfRecordDaily"
	Private Sub ISystemEventOfRecordDaily_Add(item As RecordDaily) Implements ISystemEvent(Of RecordDaily).Add
	End Sub

	Private Sub ISystemEventOfRecordDaily_Clear() Implements ISystemEvent(Of RecordDaily).Clear
	End Sub

	Private Function ISystemEventOfRecordDaily_Remove(item As RecordDaily) As Boolean Implements ISystemEvent(Of RecordDaily).Remove
		Return False
	End Function

	Private Sub ISystemEventOfRecordDaily_Load() Implements ISystemEvent(Of RecordDaily).Load
		'proceed with reading only if the time overlap
		Dim ThisTimePeriodOverlap = New TimePeriodOverlap(DirectCast(Me.Report, IDateRange))

		If ThisTimePeriodOverlap.IsOverlap(DirectCast(_RecordsDaily, IDateUpdate)) Then
			Try
				If Me.Report.AsRecordControlInfo.Enabled Then
					DirectCast(Me.Report, IStockRecordEvent).LoadBefore(Me.Symbol)
				End If
				If MyRecordIndexOfRecordDaily Is Nothing Then
					MyRecordIndexOfRecordDaily = New RecordIndex(Of RecordDaily, Date)(DirectCast(Me.Report.ToStream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "Stock\Record", "_" & Me.Symbol, ".hist.rec")
				End If
				With DirectCast(_RecordsDaily, IDateUpdate)
					.DateStart = MyRecordIndexOfRecordDaily.DateStart
					.DateStop = MyRecordIndexOfRecordDaily.DateStop
				End With
				With CType(_RecordsDaily, IRecordInfo)
					.CountTotal = MyRecordIndexOfRecordDaily.FileCount
					.MaximumID = MyRecordIndexOfRecordDaily.MaxID
				End With
				With MyRecordIndexOfRecordDaily
					If .ToListPosition.Count > 4 Then
						'CType(_RecordsDaily, LinkedHashSet(Of RecordDaily, Date)).Capacity = .ToListPosition.Count
					End If
					For Each ThisPosition In .ToListPosition
						Try
							Dim ThisRecordDaily = New RecordDaily(Me, .BaseStream(ThisPosition))
						Catch ex As Exception
							MyRecordIndexOfRecordDaily.Dispose()
							MyRecordIndexOfRecordDaily = Nothing
							Throw New Exception(String.Format("RecordDaily reading error with stock {0}", Me.Symbol))
						End Try
					Next
				End With
				MyRecordIndexOfRecordDaily.Dispose()
				MyRecordIndexOfRecordDaily = Nothing
				If Me.Report.AsRecordControlInfo.Enabled Then
					DirectCast(Me.Report, IStockRecordEvent).LoadAfter(Me.Symbol)
				End If
			Catch ex As Exception
				Throw New Exception("ISystemEvent(Of RecordDaily).Load reading file failure...", ex)
			End Try
		End If
	End Sub
#End Region
#Region "IDateUpdate"
	Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
		Get
			Dim ThisRecordsDailyDateStart As Date = DirectCast(_RecordsDaily, IDateUpdate).DateStart
			Dim ThisRecordsDateStart As Date = DirectCast(_Records, IDateUpdate).DateStart
			If ThisRecordsDailyDateStart = YahooAccessData.ReportDate.DateNullValue Then
				Return ThisRecordsDateStart
			ElseIf ThisRecordsDateStart = YahooAccessData.ReportDate.DateNullValue Then
				Return ThisRecordsDailyDateStart
			Else
				If ThisRecordsDailyDateStart < ThisRecordsDateStart Then
					Return ThisRecordsDailyDateStart
				Else
					Return ThisRecordsDateStart
				End If
			End If
		End Get
		Set(value As Date)
		End Set
	End Property
	Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
		Get
			Dim ThisRecordsDailyDateStop As Date = DirectCast(_RecordsDaily, IDateUpdate).DateStop
			Dim ThisRecordsDateStop As Date = DirectCast(_Records, IDateUpdate).DateStop
			If ThisRecordsDailyDateStop = YahooAccessData.ReportDate.DateNullValue Then
				Return ThisRecordsDateStop
			ElseIf ThisRecordsDateStop = YahooAccessData.ReportDate.DateNullValue Then
				Return ThisRecordsDailyDateStop
			Else
				If ThisRecordsDailyDateStop > ThisRecordsDateStop Then
					Return ThisRecordsDailyDateStop
				Else
					Return ThisRecordsDateStop
				End If
			End If
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
#End Region  'IDateUpdate
#Region "IDisposable Support"
	Private disposedValue As Boolean ' To detect redundant calls

	' IDisposable
	Protected Overridable Sub Dispose(disposing As Boolean)
		If Not Me.disposedValue Then
			If disposing Then
				' dispose managed state (managed objects).
			End If
			' free unmanaged resources (unmanaged objects) and override Finalize() below.
			' set large fields to null.
		End If
		Me.disposedValue = True
	End Sub

	'override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
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
#Region "IStockProcess"
	Public ReadOnly Property AsStockProcess As IStockProcess Implements IStockProcess.AsStockProcess
		Get
			Return Me
		End Get
	End Property

	Private Function IStockProcess_CalculateFiscalYearEnd() As Date Implements IStockProcess.CalculateFiscalYearEnd
		Return IStockProcess_CalculateFiscalYearEnd(Me.RecordQuoteValues.Last.DateDay.Year)
	End Function

	Private Function IStockProcess_CalculateFiscalYearEnd(Year As Integer) As Date Implements IStockProcess.CalculateFiscalYearEnd
		Dim ThisRecordDateStart As Date = Me.RecordQuoteValues.First.DateDay
		Dim ThisRecordDateStop As Date = Me.RecordQuoteValues.Last.DateDay

		If ThisRecordDateStop > DateSerial(Year, 12, 31) Then
			'enough data on top for the indicated Year
			ThisRecordDateStop = DateSerial(Year, 12, 31)
			If ThisRecordDateStart < DateSerial(Year, 1, 1) Then
				'enough data on the bottom for the indicated Year
				'limit the date to the full year
				ThisRecordDateStart = DateSerial(Year, 1, 1)
			End If
		Else
			'not enough data on top 
			'extend the botton on the previous year
			Dim ThisDate = ThisRecordDateStop.AddYears(-1)
			If ThisDate > ThisRecordDateStart Then
				ThisRecordDateStart = ThisDate
			End If
		End If

		Dim ThisRecordDaily = Me.RecordQuoteValues.ToDaily(ThisRecordDateStart, ThisRecordDateStop)


		Return ThisRecordDateStop
	End Function

	Private Function IStockProcess_CalculateFiscalQuarterEndLast() As Date Implements IStockProcess.CalculateFiscalQuarterEndLast
		Dim ThisRecordDateStart As Date = Me.RecordQuoteValues.First.DateDay
		Dim ThisRecordDateStop As Date = Me.RecordQuoteValues.Last.DateDay
		Dim ThisEarning As Single
		Dim ThisRecordDaily = Me.RecordQuoteValues.ToDaily.Reverse

		ThisEarning = ThisRecordDaily.First.EarningsShare
		For Each ThisRecord In ThisRecordDaily
			If ThisRecord.EarningsShare <> ThisEarning Then
				Return ThisRecord.DateDay
			End If
		Next
		Return ThisRecordDaily.Last.DateDay
	End Function

	Private Function IStockProcess_CalculateFiscalQuarterEnd(Year As Integer) As IEnumerable(Of Date) Implements IStockProcess.CalculateFiscalQuarterEnd

	End Function

	Private Function IStockProcess_CalculateFiscalQuarterEndLast(DateValue As Date) As Date Implements IStockProcess.CalculateFiscalQuarterEndLast

	End Function

	Private Function IStockProcess_CalculateFiscalQuarterEndLast(DateValue As Date, NumberOfQuarter As Integer) As IEnumerable(Of Date) Implements IStockProcess.CalculateFiscalQuarterEndLast

	End Function

	Private Function IStockProcess_CalculateFiscalQuarterEndLast(NumberOfQuarter As Integer) As IEnumerable(Of Date) Implements IStockProcess.CalculateFiscalQuarterEndLast
		Dim ThisRecordDateStart As Date = Me.RecordQuoteValues.First.DateDay
		Dim ThisRecordDateStop As Date = Me.RecordQuoteValues.Last.DateDay
		Dim ThisEarning As Single
		Dim ThisDate As Date
		Dim ThisRecordDaily = Me.RecordQuoteValues.ToDaily.Reverse
		Dim ThisList As New List(Of Date)
		Dim ThisState As Integer = 0
		'Dim NumberDayBetweenQuarterAverage As Integer

		ThisEarning = ThisRecordDaily.First.EarningsShare
		For Each ThisRecord In ThisRecordDaily
			Select Case ThisState
				Case 0
					If ThisRecord.EarningsShare <> ThisEarning Then
						ThisEarning = ThisRecord.EarningsShare
						ThisDate = ThisRecord.DateDay
						ThisState = 1
					End If
				Case 1
					If ThisRecord.EarningsShare = ThisEarning Then
						ThisList.Add(ThisDate)
						If ThisList.Count >= NumberOfQuarter Then
							Return ThisList
						End If
					End If
					ThisState = 0
			End Select
		Next
		Return ThisList
	End Function
#End Region 'IStockProcess
#Region "IStockInfo"
	Public ReadOnly Property AsStockInfo As IStockInfo Implements IStockInfo.AsStockInfo
		Get
			Return Me
		End Get
	End Property

	Private Property IStockInfo_Exception As Exception Implements IStockInfo.Exception
		Get
			Return Me.Exception
		End Get
		Set(value As Exception)
			Me.Exception = value
		End Set
	End Property

	Private Property IStockInfo_FullTimeEmployees As Integer Implements IStockInfo.FullTimeEmployees
		Get
			Return 0
		End Get
		Set(value As Integer)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property

	Private Property IStockInfo_IndustryName As String Implements IStockInfo.IndustryName
		Get
			Return Me.Industry.Name
		End Get
		Set(value As String)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property

	Private Property IStockInfo_Name As String Implements IStockInfo.Name
		Get
			Return Me.Name
		End Get
		Set(value As String)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property

	Private Property IStockInfo_SectorName As String Implements IStockInfo.SectorName
		Get
			Return Me.Sector.Name
		End Get
		Set(value As String)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property

	Private ReadOnly Property IStockInfo_Symbol As String Implements IStockInfo.Symbol
		Get
			Return Me.Symbol
		End Get
	End Property

	Private Property IStockInfo_TradingStartDate As Date Implements IStockInfo.TradingStartDate
		Get
			Return Me.DateStart
		End Get
		Set(value As Date)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property
	Private Property IStockInfo_TradingStopDate As Date Implements IStockInfo.TradingStopDate
		Get
			Return Me.DateStop
		End Get
		Set(value As Date)
			'not allowed at this level
			Throw New NotSupportedException
		End Set
	End Property

	Private ReadOnly Property IWebYahooDescriptor_Symbol As String Implements StockViewInterface.IWebYahooDescriptor.Symbol
		Get
			Return Me.Symbol
		End Get
	End Property

	Private ReadOnly Property IWebYahooDescriptor_Exchange As String Implements StockViewInterface.IWebYahooDescriptor.Exchange
		Get
			Return Me.Exchange
		End Get
	End Property
#End Region

End Class

#Region "IStockRank"
Public Interface IStockRank
	Property RankGain As Double
End Interface
#End Region
#Region "IStockInfo"
Public Interface IStockInfo
	ReadOnly Property AsStockInfo As IStockInfo
	Property Exception As Exception
	Property Name As String
	ReadOnly Property Symbol As String
	''' <summary>
	''' The first trading day of the company's stock
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property TradingStartDate() As Date
	''' <summary>
	''' The last trading day of the company's stock
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Property TradingStopDate() As Date
	''' <summary>
	''' The full time employees in this company
	''' </summary>
	''' <value></value>
	''' <returns>The number of employees</returns>
	''' <remarks></remarks>
	Property FullTimeEmployees() As Integer
	Property SectorName() As String
	Property IndustryName() As String
End Interface
#End Region
#Region "IStockProcess"
Public Interface IStockProcess
	ReadOnly Property AsStockProcess As IStockProcess
	'Function DetectBookEarningGlitch() As IEnumerable(Of Date)
	Function CalculateFiscalYearEnd(ByVal Year As Integer) As Date
	Function CalculateFiscalYearEnd() As Date
	Function CalculateFiscalQuarterEnd(ByVal Year As Integer) As IEnumerable(Of Date)
	Function CalculateFiscalQuarterEndLast() As Date
	Function CalculateFiscalQuarterEndLast(ByVal NumberOfQuarter As Integer) As IEnumerable(Of Date)
	Function CalculateFiscalQuarterEndLast(ByVal DateValue As Date) As Date
	Function CalculateFiscalQuarterEndLast(ByVal DateValue As Date, ByVal NumberOfQuarter As Integer) As IEnumerable(Of Date)
End Interface
#End Region
#Region "StockValue"
Public Class StockValue




End Class
#End Region
#Region "IStockValue"
Public Interface IStockValue
	ReadOnly Property Name As String
	ReadOnly Property Symbol As String
	ReadOnly Property Sector As String
	ReadOnly Property Industry As String
	ReadOnly Property DateStart As Date
	ReadOnly Property DateStop As Date
	ReadOnly Property IsSymbolError As Boolean
	ReadOnly Property ErrorDescription As String
	ReadOnly Property Exchange As String
	ReadOnly Property Price As Single
	ReadOnly Property DividendYield As Single
	ReadOnly Property PriceYieldToOneYear As Single
	ReadOnly Property EarningYield As Single
	ReadOnly Property EarningYieldToOneYear As Single


	ReadOnly Property MarketCapitalization As Single
End Interface
#End Region
#Region "EqualityComparerOfStock"
<Serializable()>
Friend Class EqualityComparerOfStock
	Implements IEqualityComparer(Of Stock)

	Public Overloads Function Equals(x As Stock, y As Stock) As Boolean Implements IEqualityComparer(Of Stock).Equals
		If (x Is Nothing) And (y Is Nothing) Then
			Return True
		ElseIf (x Is Nothing) Xor (y Is Nothing) Then
			Return False
		Else
			If x.Symbol = y.Symbol Then
				Return True
			Else
				Return False
			End If
		End If
	End Function

	Public Overloads Function GetHashCode(obj As Stock) As Integer Implements IEqualityComparer(Of Stock).GetHashCode
		If obj IsNot Nothing Then
			Return obj.Symbol.GetHashCode
		Else
			Return obj.GetHashCode
		End If
	End Function
End Class
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

Partial Public Class Stock
	Public Property ID As Integer
	Public Property Symbol As String
	Public Property Name As String
	Public Property IndustryID As Integer
	Public Property SectorID As Integer
	Public Property IsOption As Boolean
	Public Property IsSplitEnabled As Boolean
	Public Property DateStart As Date
	Public Property DateStop As Date
	Public Property IsSymbolError As Boolean
	Public Property ReportID As Integer
	Public Property Exchange As String
	Public Property ErrorDescription As String

	Public ReadOnly Property SectorName As String
		Get
			Return Me.Sector.Name
		End Get
	End Property

	Public ReadOnly Property IndustryName As String
		Get
			Return Me.Industry.Name
		End Get
	End Property

	Public Overridable Property Industry As Industry
	'Public Overridable Property Records As ICollection(Of Record) = New HashSet(Of Record)
	Public Overridable Property Report As Report
	Public Overridable Property Sector As Sector
	'Public Overridable Property SplitFactors As ICollection(Of SplitFactor) = New HashSet(Of SplitFactor)
	'Public Overridable Property StockErrors As ICollection(Of StockError) = New HashSet(Of StockError)
	'Public Overridable Property StockSymbols As ICollection(Of StockSymbol) = New HashSet(Of StockSymbol)

	'''' <summary>
	'''' Class needed to split the exchange and symbol info in stock to a similar info but modified back 
	'''' to be compatibale with WebEODHistorical web Interface
	'''' </summary>
	'Private Class StockWebEODInfo
	'  Sub New(ByVal Stock As Stock)


	'  End Sub
	'End Class
End Class
#End Region

