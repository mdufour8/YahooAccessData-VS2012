#Region "Imports"
Imports System
Imports System.IO
Imports System.Collections.Generic
'Imports System.Data.Entity
'Imports System.Data.Entity.Infrastructure
Imports System.Linq
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Net
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
Imports System.Runtime.CompilerServices
Imports System.ComponentModel
Imports System.Globalization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

#End Region
#Region "Extension"
Namespace ExtensionService
	Public Module Extensions
		Public Const LIST_OF_HEADER_FILE_ENABLED = False
		Public Enum EnumFileType
			Binary
			XML
			Json
		End Enum

		'Public Const VERSION_MEMORY_STREAM As Single = 1.1   'original file 
		Public Const VERSION_MEMORY_STREAM_FOR_BASE As Single = 1.1    'changed file to include RecordsDaily
		Public Const VERSION_MEMORY_STREAM_FOR_RECORDS_DAILY As Single = 1.2    'changed file to include RecordsDaily
		'Public Const VERSION_MEMORY_STREAM_FOR_DATETRADE As Single = 1.3    'changed file to include the IDateTrade interface information
		Public Const VERSION_MEMORY_STREAM As Single = 1.2    'changed file to include RecordsDaily

		''' <summary>
		''' Conversion of a generic primitive variable of type FromValue to another primitive generic variable of type ToValue
		''' </summary>
		''' <typeparam name="FromValue"></typeparam>
		''' <typeparam name="ToValue"></typeparam>
		''' <param name="value"></param>
		''' <returns></returns>
		''' <remarks>all primitive variable include the IConvertible interface</remarks>
		Public Function ConvertValue(Of FromValue As IConvertible, ToValue As IConvertible)(value As FromValue) As ToValue
			Return DirectCast(Convert.ChangeType(value, GetType(ToValue)), ToValue)
		End Function
#Region "Function Generic "
		'Public Function TryParse(Of T, U)(inValue As U) As T
		'  Dim converter As TypeConverter = TypeDescriptor.GetConverter(GetType(T))

		'  'Return DirectCast(converter.ConvertFromString(Nothing, CultureInfo.InvariantCulture, inValue), T)
		'  'converter.ConvertFrom(U
		'  'Return DirectCast(converter.ConvertFromString(inValue), T)
		'  Return DirectCast(converter.ConvertFrom(inValue), T)
		'End Function

		''' <summary>
		''' This function can be use for an in place shift of the data in a List.
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="colData"></param>
		''' <param name="NumberShift">
		''' Indicate the amount of shift, positive or negative
		''' </param>
		''' <returns>
		''' return the list shifted
		''' </returns>
		''' <remarks></remarks>
		<Extension()>
		Public Function ShiftTo(Of T As {New})(colData As IList(Of T), ByVal NumberShift As Integer) As IList(Of T)
			Dim ThisListOfCopy As New List(Of T)(colData)

			Dim I As Integer
			If colData.Count > 0 Then
				If NumberShift > 0 Then
					For I = 1 To NumberShift
						colData.RemoveAt(colData.Count - 1)
						'colData.Insert(0, New T)
						colData.Insert(0, colData.First)
					Next
				Else
					For I = 1 To -NumberShift
						colData.RemoveAt(0)
						colData.Add(colData.Last)
					Next
				End If
			End If
			Return colData
		End Function

		''' <summary>
		''' This function can be use for to copy and shift the data
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="colData"></param>
		''' <param name="NumberShift">
		''' Indicate the amount of shift, positive or negative
		''' </param>
		''' <param name="IsCopy">
		''' indicate is the data is copied before shifting
		''' </param>
		''' <returns></returns>
		''' <remarks></remarks>
		<Extension()>
		Public Function ShiftTo(Of T As {New})(colData As IList(Of T), ByVal NumberShift As Integer, ByVal IsCopy As Boolean) As IList(Of T)
			If IsCopy Then
				Dim ThisListOfCopy As New List(Of T)(colData)
				Return ThisListOfCopy.ShiftTo(NumberShift)
			Else
				Return colData.ShiftTo(NumberShift)
			End If
		End Function


		''' <summary>
		''' Will shift end extend the number of element of the array to accomodate the number of shift
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="colData"></param>
		''' <param name="NumberShift"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		<Extension()>
		Public Function ShiftToExtend(Of T As {New})(colData As IList(Of T), ByVal NumberShift As Integer) As IList(Of T)
			Dim ThisListOfCopy As New List(Of T)(colData)

			Dim I As Integer
			If ThisListOfCopy.Count > 0 Then
				If NumberShift > 0 Then
					For I = 1 To NumberShift
						ThisListOfCopy.Insert(0, ThisListOfCopy.First)
					Next
				Else
					For I = 1 To -NumberShift
						ThisListOfCopy.Add(ThisListOfCopy.Last)
					Next
				End If
			End If
			Return ThisListOfCopy
		End Function

		<Extension()>
		Public Function ToDataTable(Of T As {IFormatData})(
			colData As IEnumerable(Of T)) As DataTable

			'create the data table coloumn
			Dim ThisDataTable As New DataTable
			Dim ThisType As Type
			Dim ThisPropertyInfo As System.Reflection.PropertyInfo
			Dim ThisListOfHeader As List(Of HeaderInfo)
			Dim ThisPropertyInfoes() As System.Reflection.PropertyInfo
			Dim I As Integer

			ThisListOfHeader = colData.ToListOfHeader
			ReDim ThisPropertyInfoes(0 To ThisListOfHeader.Count - 1)
			ThisType = GetType(T)
			ThisDataTable.TableName = ThisType.FullName
			For Each ThisHeaderInfo In colData.ToListOfHeader
				If ThisHeaderInfo.Visible Then
					ThisPropertyInfo = ThisType.GetProperty(ThisHeaderInfo.Name)
					ThisDataTable.Columns.Add(ThisHeaderInfo.Name, ThisPropertyInfo.PropertyType)
					'store the propertyinfo for use later
					ThisPropertyInfoes(ThisDataTable.Columns.Count - 1) = ThisPropertyInfo
				End If
			Next
			For Each ThisT In colData
				Dim ThisRow = ThisDataTable.NewRow
				For I = 0 To ThisDataTable.Columns.Count - 1
					ThisRow(I) = ThisPropertyInfoes(I).GetValue(ThisT, Nothing)
				Next
				ThisDataTable.Rows.Add(ThisRow)
			Next
			Return ThisDataTable
		End Function

		Public Function IsEvenNumber(ByVal Number As Integer) As Boolean
			Return (Number And 1) = 0
		End Function

		Public Function ToDaily(Of T As {New, Class, IRegisterKey(Of Date), IDateUpdate, IFormatData})(
			colData As IEnumerable(Of T),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of T)

			Dim colDataDaily As New LinkedHashSet(Of T, Date)
			Dim ThisRecordQuoteValue As IRecordQuoteValue
			Dim ThisRecordQuoteValueLast As IRecordQuoteValue

			If colData.Count > 0 Then
				'record data is expected to always be by increasing date
				Dim ThisDataLast As T = Nothing
				For Each ThisData In colData
					If ThisData.DateDay >= DateStart Then
						If ThisData.DateDay <= DateStop Then
							If ThisDataLast IsNot Nothing Then
								'check if the trading date change
								If ThisData.DateLastTrade.Date <> ThisDataLast.DateLastTrade.Date Then
									'new trading day
									'save last record from the previous day in the collection
									colDataDaily.Add(ThisDataLast)
								Else
									If TypeOf ThisData Is YahooAccessData.IRecordQuoteValue Then
										'there is some know bug with yahoo where the BookValue and Earning would change during the day
										'eliminate these glitch and correct the data
										ThisRecordQuoteValue = DirectCast(ThisData, YahooAccessData.IRecordQuoteValue)
										ThisRecordQuoteValueLast = DirectCast(ThisDataLast, YahooAccessData.IRecordQuoteValue)
										If (ThisRecordQuoteValue.BookValue <> ThisRecordQuoteValueLast.BookValue) Or (ThisRecordQuoteValue.EarningsShare <> ThisRecordQuoteValueLast.EarningsShare) Then
											ThisRecordQuoteValue.RemoveBookEarningGlitch(ThisRecordQuoteValueLast)
										End If
									End If
								End If
							End If
							'update the last data
							ThisDataLast = ThisData
						Else
							'the record are always by increasing date
							'so there is no more record in the date range
							Exit For
						End If
					End If
				Next
				If ThisDataLast IsNot Nothing Then
					'add the last day to whatever time we are
					colDataDaily.Add(ThisDataLast)
				End If
				Return colDataDaily
			Else
				'no data to process
				'return the same collection???
				'Return colData
				Return colDataDaily
			End If
		End Function

		Public Function ToDaily(Of T As {New, Class, IRegisterKey(Of Date), IDateUpdate, IFormatData})(
			colData As IEnumerable(Of IEnumerable(Of T))) As IEnumerable(Of T)

			Dim colDataDaily As New LinkedHashSet(Of T, Date)
			For Each ThisList In colData
				colDataDaily.Add(ThisList.Last)
			Next
			Return colDataDaily
		End Function

		''' <summary>
		''' Transform the daily data to an enumarable array of dailey PriceVol
		''' </summary>
		''' <param name="PriceVolDailyData">
		''' The data need to be on daily time format and increasing time
		''' </param>
		''' <returns></returns>
		<Extension>
		Public Function ToDaily(ByRef PriceVolDailyData As PriceVol()) As IEnumerable(Of IPriceVol)
			Dim ThisListOfPriceVolDaily As New List(Of IPriceVol)
			Dim I As Integer

			If PriceVolDailyData.Count = 0 Then Return ThisListOfPriceVolDaily
			For I = 0 To PriceVolDailyData.Length - 1
				ThisListOfPriceVolDaily.Add(PriceVolDailyData(I).CopyFromAsClass)
			Next
			Return ThisListOfPriceVolDaily
		End Function


		''' <summary>
		''' This funtion return true if the data are in the same work week from Monday to Friday included.
		''' the function return false if any of the data fall on a weekend. The date can be of any order realtive to each other. 
		''' </summary>
		''' <param name="DateValueRef"></param>
		''' <param name="DateValue"></param>
		''' <returns></returns>
		<Extension>
		Public Function IsSameTradingWeek(ByVal DateValueRef As Date, DateValue As Date) As Boolean
			Return ReportDate.IsSameTradingWeek(DateValueRef, DateValue)
		End Function


		''' <summary>
		''' Transform the daily data to an enumarable array of weekly PriceVol
		''' </summary>
		''' <param name="PriceVolDailyData">
		''' The data need to be on daily time format and increasing time
		''' </param>
		''' <returns></returns>
		<Extension>
		Public Function ToWeekly(ByRef PriceVolDailyData As PriceVol()) As IEnumerable(Of IPriceVol)
			Dim ThisListOfPriceVolWeekly As New List(Of IPriceVol)
			Dim ThisWekklyData As IPriceVol = Nothing
			Dim I As Integer
			Dim ThisArraySize As Integer

			If PriceVolDailyData.Count = 0 Then Return ThisListOfPriceVolWeekly
			'it is easier and more efficient to work with a class here
			'only the reference will be copied in the enumaration
			ThisArraySize = PriceVolDailyData.Length - 1
			ThisWekklyData = PriceVolDailyData(0).CopyFromAsClass
			For I = 1 To ThisArraySize
				If ReportDate.IsSameTradingWeek(PriceVolDailyData(I).DateLastTrade, PriceVolDailyData(I - 1).DateLastTrade) Then
					'update weeklydata but do not save yet in the list
					If PriceVol.AddToWeeklyMerge(PriceVolIn:=PriceVolDailyData(I).AsIPriceVol, PriceVolResult:=ThisWekklyData) = False Then
						Throw New InvalidCastException("Invalid weekly data transformation...")
					End If
				Else
					ThisListOfPriceVolWeekly.Add(ThisWekklyData)
					're-initialize the new weekly data with this new week date and keep going
					ThisWekklyData = PriceVolDailyData(I).CopyFromAsClass
				End If
			Next
			ThisListOfPriceVolWeekly.Add(ThisWekklyData)
			Return ThisListOfPriceVolWeekly
		End Function

		Public Function ToDaily(Of T As {New, Class, IRegisterKey(Of Date), IDateUpdate, IFormatData})(
			colData As IEnumerable(Of T)) As IEnumerable(Of T)

			Dim colDataDaily As New LinkedHashSet(Of T, Date)
			Dim ThisRecordQuoteValue As IRecordQuoteValue
			Dim ThisRecordQuoteValueLast As IRecordQuoteValue
			Dim ThisData As T
			Dim ThisDataLast As T

			If colData.Count > 0 Then
				'data is always by increasing date
				ThisDataLast = Nothing
				For Each ThisData In colData
					If ThisDataLast IsNot Nothing Then
						'check if the trading date change
						If ThisData.DateLastTrade.Date <> ThisDataLast.DateLastTrade.Date Then
							'new trading day
							'save the previous day in the collection
							colDataDaily.Add(ThisDataLast)
						Else
							If TypeOf ThisData Is YahooAccessData.IRecordQuoteValue Then
								'there is some know bug with yahoo where the BookValue and Earning would change during the day
								'eliminate these glitch and correct the data
								ThisRecordQuoteValue = DirectCast(ThisData, YahooAccessData.IRecordQuoteValue)
								ThisRecordQuoteValueLast = DirectCast(ThisDataLast, YahooAccessData.IRecordQuoteValue)
								If (ThisRecordQuoteValue.BookValue <> ThisRecordQuoteValueLast.BookValue) Or (ThisRecordQuoteValue.EarningsShare <> ThisRecordQuoteValueLast.EarningsShare) Then
									ThisRecordQuoteValue.RemoveBookEarningGlitch(ThisRecordQuoteValueLast)
								End If
							End If
						End If
					End If
					'update the last data
					ThisDataLast = ThisData
				Next
				If ThisDataLast IsNot Nothing Then
					'add the last day to whatever time we are
					colDataDaily.Add(ThisDataLast)
				End If
				Return colDataDaily
			Else
				'no data to process
				'return the same collection???
				'Return colData
				Return colDataDaily
			End If
		End Function

		Public Function ToDailyIntraDay(Of T As {New, Class, IRegisterKey(Of Date), IDateUpdate, IFormatData})(
			colData As IEnumerable(Of T),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of IEnumerable(Of T))

			Dim colDataDailyIntraDay As New List(Of IEnumerable(Of T))
			Dim colDataForOneDay As LinkedHashSet(Of T, Date)
			Dim ThisRecordQuoteValue As IRecordQuoteValue = Nothing
			Dim ThisRecordQuoteValueLast As IRecordQuoteValue = Nothing
			Dim ThisData As T
			Dim ThisDataLast As T = Nothing
			'Dim ThisPriceLowDaily As Single
			'Dim ThisPriceHighDaily As Single
			'Dim ThisHighLowRange As Single
			'Dim ThisHighLowRangeLast As Single
			'Dim ThisPriceOpen As Single
			'Dim IsPriceHighLowError As Boolean

			If colData.Count > 0 Then
				'record data is expected to always be by increasing date
				colDataForOneDay = New LinkedHashSet(Of T, Date)
				Dim ThisDataLastItem = colData.Last
				Dim ThisDataLastItemPrevious = colData.ElementAt(colData.Count - 2)
				For Each ThisData In colData
					If ThisData.DateDay >= DateStart Then
						If ThisData.DateDay <= DateStop Then
							ThisRecordQuoteValue = DirectCast(ThisData, YahooAccessData.IRecordQuoteValue)
							If ThisDataLast IsNot Nothing Then
								'check if the trading date change
								If ThisData.DateLastTrade.Date <> ThisDataLast.DateLastTrade.Date Then
									'new trading day
									'save last record from the previous day in the collection
									colDataDailyIntraDay.Add(colDataForOneDay)
									colDataForOneDay = New LinkedHashSet(Of T, Date)
								Else
									If TypeOf ThisData Is YahooAccessData.IRecordQuoteValue Then
										'there is some know bug with yahoo where the BookValue and Earning would change during the day
										'eliminate these glitch and correct the data
										If (ThisRecordQuoteValue.BookValue <> ThisRecordQuoteValueLast.BookValue) Or (ThisRecordQuoteValue.EarningsShare <> ThisRecordQuoteValueLast.EarningsShare) Then
											ThisRecordQuoteValue.RemoveBookEarningGlitch(ThisRecordQuoteValueLast)
										End If
									End If
								End If
							End If
							''record the high low from the last value just in case there is an error on the yahoo high low 
							'If colDataForOneDay.Count = 0 Then
							'  'initialization 
							'  ThisHighLowRange = 0
							'  IsPriceHighLowError = False
							'End If
							'ThisHighLowRangeLast = ThisHighLowRange
							'ThisHighLowRange = ThisRecordQuoteValue.High - ThisRecordQuoteValue.Low
							''check for anomaly in the high low value
							''the high low range should always be increasing
							'If ThisHighLowRange < ThisHighLowRangeLast Then
							'  'this is an anomaly
							'  IsPriceHighLowError = True
							'End If
							'update the last data
							ThisDataLast = ThisData
							ThisRecordQuoteValueLast = ThisRecordQuoteValue
							colDataForOneDay.Add(ThisDataLast)
						Else
							'the record are always by increasing date
							'so there is no more record in the date range
							Exit For
						End If
					End If
				Next
				If colDataForOneDay.Count > 0 Then
					'If IsPriceHighLowError Then
					'  'there is an error in the High low value
					'  'correct everything using the last value
					'  ThisRecordQuoteValue = DirectCast(colDataForOneDay.First, YahooAccessData.IRecordQuoteValue)
					'  With ThisRecordQuoteValue
					'    ThisPriceLowDaily = ThisRecordQuoteValue.Last
					'    ThisPriceHighDaily = ThisRecordQuoteValue.Last
					'    ThisPriceOpen = ThisRecordQuoteValue.Last
					'  End With
					'  For Each ThisRecordQuoteValue In colDataForOneDay
					'    ThisRecordQuoteValue.PriceChange(ThisPriceOpen, ThisPriceLowDaily, ThisPriceHighDaily, ThisRecordQuoteValue.Last)
					'    'update the high low value from the last
					'    If ThisRecordQuoteValue.Last < ThisPriceLowDaily Then
					'      ThisPriceLowDaily = ThisRecordQuoteValue.Last
					'    ElseIf ThisRecordQuoteValue.Last > ThisPriceHighDaily Then
					'      ThisPriceHighDaily = ThisRecordQuoteValue.Last
					'    End If
					'  Next
					'End If
					'add the last day to whatever time we are
					colDataDailyIntraDay.Add(colDataForOneDay)
				End If
				Return colDataDailyIntraDay
			Else
				'no data to process
				'return the same collection???
				'Return colData
				Return colDataDailyIntraDay
			End If
		End Function

		Public Function ToDailyIntraDay(Of T As {New, Class, IRegisterKey(Of Date), IDateUpdate, IFormatData})(
			colData As IEnumerable(Of T)) As IEnumerable(Of IEnumerable(Of T))

			Dim colDataDailyIntraDay As New List(Of IEnumerable(Of T))
			Dim colDataForOneDay As LinkedHashSet(Of T, Date)
			Dim ThisRecordQuoteValue As IRecordQuoteValue
			Dim ThisRecordQuoteValueLast As IRecordQuoteValue

			If colData.Count > 0 Then
				'data is always by increasing date
				colDataForOneDay = New LinkedHashSet(Of T, Date)
				Dim ThisDataLast As T = Nothing
				For Each ThisData In colData
					If ThisDataLast IsNot Nothing Then
						'check if the trading date change
						If ThisData.DateLastTrade.Date <> ThisDataLast.DateLastTrade.Date Then
							'new trading day
							'save the previous day in the collection
							colDataDailyIntraDay.Add(colDataForOneDay)
							colDataForOneDay = New LinkedHashSet(Of T, Date)
						Else
							If TypeOf ThisData Is YahooAccessData.IRecordQuoteValue Then
								'there is some know bug with yahoo where the BookValue and Earning would change during the day
								'eliminate these glitch and correct the data
								ThisRecordQuoteValue = DirectCast(ThisData, YahooAccessData.IRecordQuoteValue)
								ThisRecordQuoteValueLast = DirectCast(ThisDataLast, YahooAccessData.IRecordQuoteValue)
								If (ThisRecordQuoteValue.BookValue <> ThisRecordQuoteValueLast.BookValue) Or (ThisRecordQuoteValue.EarningsShare <> ThisRecordQuoteValueLast.EarningsShare) Then
									ThisRecordQuoteValue.RemoveBookEarningGlitch(ThisRecordQuoteValueLast)
								End If
							End If
						End If
					End If
					'update the last data
					ThisDataLast = ThisData
					colDataForOneDay.Add(ThisDataLast)
				Next
				If colDataForOneDay.Count > 0 Then
					'add the last day to whatever time we are
					colDataDailyIntraDay.Add(colDataForOneDay)
				End If
				Return colDataDailyIntraDay
			Else
				'no data to process
				'return the same collection???
				'Return colData
				Return colDataDailyIntraDay
			End If
		End Function
#End Region
#Region "Serialize HeaderInfo"
		<Extension()>
		Public Function ToListOfHeader(Of T)(colData As IEnumerable(Of T)) As List(Of HeaderInfo)

			If TypeOf colData Is IHeaderListInfo Then
				Return TryCast(colData, IHeaderListInfo).HeaderInfo
			Else
				Return New List(Of HeaderInfo)
			End If
		End Function

		<Extension()>
		Public Function ToHeaderDataItem(Of T As {IFormatData})(colData As IEnumerable(Of T), ByVal Item As Integer) As String
			Dim ThisFormatData As IFormatData
			Dim ThisData As String()
			Dim ThisResult As New System.Text.StringBuilder(1000)
			Dim ThisListOfHeader As List(Of YahooAccessData.HeaderInfo)
			Dim I As Integer

			If Item < colData.Count Then
				ThisFormatData = colData(Item)
				With ThisFormatData
					ThisListOfHeader = .ToListOfHeader
					ThisData = .ToStingOfData
				End With
				For I = 0 To ThisListOfHeader.Count - 2
					ThisResult.Append(ThisListOfHeader(I).Name)
					ThisResult.Append(":")
					ThisResult.Append(ThisData(I).ToString)
					ThisResult.Append(",")
				Next
				'last one
				ThisResult.Append(ThisListOfHeader(I))
				ThisResult.Append(":")
				ThisResult.Append(ThisData(I).ToString)
			End If
			Return Trim(ThisResult.ToString)
		End Function

		Public Function ToStingOfData(Of T As {IFormatData})(ByRef Data As T) As String()
			Dim ThisResult As New List(Of String)
			Dim ThisItem As PropertyInfo

			For Each ThisHeaderInfo In DirectCast(Data, IFormatData).ToListOfHeader
				With ThisHeaderInfo
					If .Visible Then
						Try
							ThisItem = Data.GetType.GetProperty(.Name)
						Catch ex As Exception
							ThisItem = Nothing
						End Try
						If ThisItem Is Nothing Then
							ThisResult.Add("")
						Else
							ThisResult.Add(String.Format(.Format, ThisItem.GetValue(Data, Nothing)))
						End If
					End If
				End With
			Next
			Return ThisResult.ToArray
		End Function

		Public Function CreateInstance(Of T As {New})() As T
			Dim ThisInstance As New T
			Return ThisInstance
			'this will also work but seem more elaborate and slower than necessary
			'Return CType(GetType(T).GetConstructor(New System.Type() {}).Invoke(New Object() {}), T)
		End Function
#End Region
#Region "Reports"
		'<Extension()>
		'Public Function ItemTakeLasts(
		'  colReports As DbSet(Of YahooAccessData.Report),
		'  colSectors As DbSet(Of YahooAccessData.Sector),
		'  ByVal ReportName As String,
		'  ByVal SectorName As String,
		'  ByVal NumberToTake As Integer) As IEnumerable(Of Report)

		'  Dim ThisResult As List(Of Report) = Nothing
		'  Try
		'    Dim ThisSectorResults As IQueryable(Of YahooAccessData.Report) =
		'      From ThisS As YahooAccessData.Sector In colSectors
		'      Join ThisR As YahooAccessData.Report In colReports
		'      On ThisS.ID Equals ThisR.ID
		'      Where (ThisS.Name = SectorName) AndAlso (ThisR.Name = ReportName)
		'      Order By ThisR.DateStart Descending
		'      Select ThisR
		'      Take NumberToTake

		'    ThisResult = ThisSectorResults.ToList
		'    If ThisResult Is Nothing Then
		'      ThisResult = New List(Of Report)
		'    End If
		'  Catch ex As Exception
		'    Throw (New Exception("Database Error executing function ItemTakeLasts...", ex))
		'    ThisResult = New List(Of Report)
		'  End Try
		'  Return ThisResult
		'End Function

		'<Extension()>
		'Public Function ItemTakeLasts(
		'  colReports As DbSet(Of YahooAccessData.Report),
		'  colSectors As DbSet(Of YahooAccessData.Sector),
		'  ByVal SectorName As String,
		'  ByVal NumberToTake As Integer) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report) = Nothing
		'  Try
		'    Dim ThisSectorResults As IQueryable(Of YahooAccessData.Report) =
		'      From ThisS As YahooAccessData.Sector In colSectors
		'      Join ThisR As YahooAccessData.Report In colReports
		'      On ThisS.ID Equals ThisR.ID
		'      Where (ThisS.Name = SectorName)
		'      Order By ThisR.DateStart Descending
		'      Select ThisR
		'      Take NumberToTake

		'    ThisResult = ThisSectorResults.AsEnumerable
		'    If ThisResult Is Nothing Then
		'      ThisResult = New List(Of Report)
		'    End If
		'  Catch ex As Exception
		'    Throw (New Exception("Database Error executing function ItemTakeLasts...", ex))
		'    ThisResult = New List(Of Report)
		'  End Try
		'  Return ThisResult
		'End Function

		'<Extension()>
		'Public Function ItemTakeLasts(
		'  colReports As DbSet(Of YahooAccessData.Report),
		'  ByVal ReportName As String,
		'  ByVal NumberToTake As Integer) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report) = Nothing
		'  Try
		'    Dim ThisSectorResults As IQueryable(Of YahooAccessData.Report) =
		'      From ThisR As YahooAccessData.Report In colReports
		'      Where (ThisR.Name = ReportName)
		'      Order By ThisR.DateStart Descending
		'      Select ThisR
		'      Take NumberToTake

		'    ThisResult = ThisSectorResults.AsEnumerable
		'    If ThisResult Is Nothing Then
		'      ThisResult = New List(Of Report)
		'    End If
		'  Catch ex As Exception
		'    Throw (New Exception("Database Error executing function ItemTakeLasts...", ex))
		'    ThisResult = New List(Of Report)
		'  End Try
		'  Return ThisResult
		'End Function

		'<Extension()>
		'Public Function ItemTakeLasts(
		'  colReports As DbSet(Of YahooAccessData.Report),
		'  ByVal NumberToTake As Integer) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report) = Nothing
		'  Try
		'    Dim ThisSectorResults As IQueryable(Of YahooAccessData.Report) =
		'      From ThisR As YahooAccessData.Report In colReports
		'      Order By ThisR.DateStart Descending
		'      Select ThisR
		'      Take NumberToTake

		'    ThisResult = ThisSectorResults.AsEnumerable
		'    If ThisResult Is Nothing Then
		'      ThisResult = New List(Of Report)
		'    End If
		'  Catch ex As Exception
		'    Throw (New Exception("Database Error executing function ItemTakeLasts...", ex))
		'    ThisResult = New List(Of Report)
		'  End Try
		'  Return ThisResult
		'End Function

		'<Extension()>
		'Public Function Items(
		'  colData As DbSet(Of YahooAccessData.Report),
		'  ByVal Name As String,
		'  ByVal DateStart As Date,
		'  ByVal DateStop As Date,
		'  Optional ByVal IsCreateIfEmpty As Boolean = False) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report) = Nothing
		'  Dim ThisCount As Integer

		'  Try
		'    ThisResult = colData.
		'      Where(Function(ThisR) (ThisR.Name = Name) And (ThisR.DateStart >= DateStart) And (ThisR.DateStop <= DateStop)).
		'      OrderBy(Function(ThisR) ThisR.DateStart)
		'    If ThisResult Is Nothing Then
		'      Debug.Assert(False)
		'      ThisResult = New List(Of Report)
		'      ThisCount = 0
		'    Else
		'      '.Count may not be accessible in case the query or access to the database failed
		'      ThisCount = ThisResult.Count
		'    End If
		'  Catch ex As Exception
		'    Debug.Assert(False)
		'    ThisResult = New List(Of Report)
		'    ThisCount = 0
		'  End Try
		'  If ThisCount = 0 Then
		'    If IsCreateIfEmpty Then
		'      Dim ThisReport As Report = colData.Create()
		'      With ThisReport
		'        .Name = Name
		'        .DateStart = DateStart
		'        .DateStop = DateStop
		'      End With
		'      colData.Add(ThisReport)
		'      Return colData.Items(Name, DateStart, DateStop)
		'    Else
		'      Return ThisResult
		'    End If
		'  Else
		'    Return ThisResult
		'  End If
		'End Function

		'<Extension()>
		'Public Function Items(
		'  colData As DbSet(Of YahooAccessData.Report),
		'  ByVal Name As String,
		'  Optional ByVal IsCreateIfEmpty As Boolean = False) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report) = Nothing
		'  Dim ThisCount As Integer
		'  Try
		'    ThisResult = colData.Where(Function(ThisR As YahooAccessData.Report) (ThisR.Name = Name)).OrderBy(Function(ThisR As YahooAccessData.Report) ThisR.DateStart)
		'    If ThisResult Is Nothing Then
		'      ThisResult = New List(Of Report)
		'      ThisCount = 0
		'    Else
		'      '.Count may not be accessible in case the query or access to the database failed
		'      'this will generate an error that will be trap below
		'      ThisCount = ThisResult.Count
		'    End If
		'  Catch ex As Exception
		'    ThisResult = New List(Of Report)
		'    ThisCount = 0
		'  End Try
		'  If ThisCount = 0 Then
		'    If IsCreateIfEmpty Then
		'      Dim ThisReport As Report = colData.Create()
		'      With ThisReport
		'        .Name = Name
		'        .DateStart = Now
		'        .DateStop = .DateStart
		'      End With
		'      colData.Add(ThisReport)
		'      Return colData.Items(Name, ThisReport.DateStart, ThisReport.DateStop)
		'    Else
		'      Return ThisResult
		'    End If
		'  Else
		'    Return ThisResult
		'  End If
		'End Function

		'<Extension()>
		'Public Function Items(
		'  colData As DbSet(Of YahooAccessData.Report),
		'  ByVal ThisReport As Report,
		'  Optional ByVal IsCreateIfEmpty As Boolean = False) As IEnumerable(Of Report)

		'  Dim ThisResult As IEnumerable(Of Report)
		'  With ThisReport
		'    ThisResult = colData.Items(.Name, .DateStart, .DateStop)
		'  End With
		'  If ThisResult.Count = 0 Then
		'    If IsCreateIfEmpty Then
		'      colData.Add(ThisReport)
		'      With ThisReport
		'        ThisResult = colData.Items(.Name, .DateStart, .DateStop)
		'      End With
		'      Return ThisResult
		'    Else
		'      Return ThisResult
		'    End If
		'  Else
		'    Return ThisResult
		'  End If
		'End Function
#End Region
#Region "Stocks"
		<Extension()>
		Public Function ToListOfExchange(
			colData As ICollection(Of YahooAccessData.Stock)) As IEnumerable(Of String)

			Dim ThisDictionary As New Dictionary(Of String, String)
			Dim ThisList As New List(Of String)

			For Each ThisStock In colData
				If ThisStock.Sector.Name = "Indices" Then
					If ThisDictionary.ContainsKey("Indices") = False Then
						ThisDictionary.Add("Indices", "Indices")
						ThisList.Add("Indices")
					End If
				Else
					If ThisDictionary.ContainsKey(ThisStock.Exchange) = False Then
						ThisDictionary.Add(ThisStock.Exchange, ThisStock.Exchange)
						ThisList.Add(ThisStock.Exchange)
					End If
				End If
			Next
			Return ThisList.OrderBy(Function(ThisExchange As String) (ThisExchange))
		End Function

		<Extension()>
		Public Function ToGroupByExchange(
			colData As ICollection(Of YahooAccessData.Stock),
			ByVal ExchangeName As String) As IEnumerable(Of Stock)

			Dim ThisList As New LinkedHashSet(Of Stock, String)

			For Each ThisStock In colData.Where(Function(ThisS As YahooAccessData.Stock) (ThisS.Exchange = ExchangeName))
				ThisList.Add(ThisStock)
			Next
			Return ThisList
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.Stock),
			ByVal Key As String,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.Stock

			Dim ThisStock As YahooAccessData.Stock
			ThisStock = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisStock Is Nothing Then
					ThisStock = New YahooAccessData.Stock(Symbol:=Key)
					colData.Add(ThisStock)
				End If
			End If
			Return ThisStock
		End Function

		'<Extension()>
		'Public Function Search(
		'  colData As ICollection(Of YahooAccessData.Stock)) As YahooAccessData.ISearchKey(Of Stock, String)
		'  Return TryCast(colData, YahooAccessData.ISearchKey(Of Stock, String))
		'End Function

		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.Stock)) As YahooAccessData.ISearchKey(Of Stock, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of Stock, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.Stock)) As YahooAccessData.ISort(Of Stock)
			Return TryCast(colData, YahooAccessData.ISort(Of Stock))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.Stock),
			ByRef Parent As Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			'Dim ThisStopWatch = New System.Diagnostics.Stopwatch

			'ThisStopWatch.Restart()
			For Each ThisStockData In colData
				ThisStockData.CopyDeep(Parent, IsIgnoreID)
			Next
			'ThisStopWatch.Stop()
		End Sub

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.Stock),
			ByRef Parent As Report,
			ByVal StartElement As Integer,
			ByVal StopElement As Integer,
			Optional ByVal IsIgnoreID As Boolean = False)

			'Dim ThisStopWatch = New System.Diagnostics.Stopwatch
			Dim I As Integer
			'ThisStopWatch.Restart()

			If colData.Count > StopElement Then
				For I = StartElement To StopElement
					colData(I).CopyDeep(Parent, IsIgnoreID)
				Next
			End If
			'ThisStopWatch.Stop()
			'Debug.Print(ThisStopWatch.ElapsedMilliseconds.ToString)
		End Sub
		<Extension()>
		Friend Sub CopyDeep(
			colData As IEnumerable(Of YahooAccessData.Stock),
			ByRef Parent As Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			'Dim ThisStopWatch = New System.Diagnostics.Stopwatch

			'ThisStopWatch.Restart()
			For Each ThisStockData In colData
				ThisStockData.CopyDeep(Parent, IsIgnoreID)
			Next
			'ThisStopWatch.Stop()
			'Debug.Print(ThisStopWatch.ElapsedMilliseconds.ToString)
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.Stock),
			other As ICollection(Of Stock),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then
					Return False
				End If
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.Stock),
			other As ICollection(Of Stock)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "SplitFactorFutures"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.SplitFactorFuture)) As YahooAccessData.ISearchKey(Of YahooAccessData.SplitFactorFuture, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.SplitFactorFuture, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.SplitFactorFuture)) As YahooAccessData.ISort(Of YahooAccessData.SplitFactorFuture)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.SplitFactorFuture))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.SplitFactorFuture),
			ByRef Parent As Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisSplitFactorFuture In colData
				ThisSplitFactorFuture.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.SplitFactorFuture),
			other As ICollection(Of YahooAccessData.SplitFactorFuture),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.SplitFactorFuture),
			other As ICollection(Of YahooAccessData.SplitFactorFuture)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "BondRates"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.BondRate)) As YahooAccessData.ISearchKey(Of YahooAccessData.BondRate, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.BondRate, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.BondRate)) As YahooAccessData.ISort(Of YahooAccessData.BondRate)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.BondRate))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.BondRate),
			ByRef Parent As Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			Dim ThisBondRate As YahooAccessData.BondRate

			For Each ThisBondRate In colData
				ThisBondRate.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub


		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.BondRate),
			other As ICollection(Of YahooAccessData.BondRate),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.BondRate),
			other As ICollection(Of YahooAccessData.BondRate)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "BondRate1"
		<Extension()>
		Public Function ToSearch(
		colData As ICollection(Of YahooAccessData.BondRate1)) As YahooAccessData.ISearchKey(Of YahooAccessData.BondRate1, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.BondRate1, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.BondRate1)) As YahooAccessData.ISort(Of YahooAccessData.BondRate1)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.BondRate1))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.BondRate1),
			ByRef Parent As Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			Dim ThisBondRate1 As YahooAccessData.BondRate1

			For Each ThisBondRate1 In colData
				ThisBondRate1.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub


		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.BondRate1),
			other As ICollection(Of YahooAccessData.BondRate1),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.BondRate1),
			other As ICollection(Of YahooAccessData.BondRate1)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "StockErrors"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.StockError)) As YahooAccessData.ISearchKey(Of YahooAccessData.StockError, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.StockError, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.StockError)) As YahooAccessData.ISort(Of YahooAccessData.StockError)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.StockError))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.StockError),
			ByRef Parent As Stock,
			Optional ByVal IsIgnoreID As Boolean = False)

			Dim ThisStockError As StockError

			For Each ThisStockError In colData
				ThisStockError.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.StockError),
			other As ICollection(Of YahooAccessData.StockError),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.StockError),
			other As ICollection(Of YahooAccessData.StockError)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "IStockQuote"
		<Extension()>
		Public Function ToListOfRecord(colData As List(Of WebEODData.IStockQuote)) As List(Of YahooAccessData.Record)
			Dim ThisList = New List(Of YahooAccessData.Record)
			Dim ThisRecord As YahooAccessData.Record
			For Each ThisStockQuote In colData
				ThisRecord = New YahooAccessData.Record
				With ThisRecord
					.DateDay = ThisStockQuote.DateTime
					.DateLastTrade = .DateDay
					.DateUpdate = .DateDay
					.High = ThisStockQuote.High.ToSingleSafe(RoundingDigit:=3)
					.Open = ThisStockQuote.Open.ToSingleSafe(RoundingDigit:=3)
					.Low = ThisStockQuote.Low.ToSingleSafe(RoundingDigit:=3)
					.Last = ThisStockQuote.Close.ToSingleSafe(RoundingDigit:=3)
					.Vol = ThisStockQuote.Volume.ToIntegerSafe
					If ThisStockQuote.IsLiveUpdate Then
						.AsIRecordType.RecordType = IRecordType.enuRecordType.LiveUpdate
					Else
						.AsIRecordType.RecordType = IRecordType.enuRecordType.EndOfDay
					End If
				End With
				ThisList.Add(ThisRecord)
			Next
			Return ThisList
		End Function

		<Extension()>
		Public Function ToRecord(StockQuote As WebEODData.IStockQuote, IsLiveUpdate As Boolean) As Record
			Dim ThisRecord As New Record
			With ThisRecord
				.DateDay = StockQuote.DateTime
				.DateLastTrade = .DateDay
				.DateUpdate = .DateDay
				.High = StockQuote.High.ToSingleSafe(RoundingDigit:=3)
				.Open = StockQuote.Open.ToSingleSafe(RoundingDigit:=3)
				.Low = StockQuote.Low.ToSingleSafe(RoundingDigit:=3)
				.Last = StockQuote.Close.ToSingleSafe(RoundingDigit:=3)
				.Vol = StockQuote.Volume.ToIntegerSafe
				If IsLiveUpdate Then
					.AsIRecordType.RecordType = IRecordType.enuRecordType.LiveUpdate
				Else
					.AsIRecordType.RecordType = IRecordType.enuRecordType.EndOfDay
				End If
			End With
			Return ThisRecord
		End Function
#End Region
#Region "Support Function"
		<Extension>
		Private Function ToSingleSafe(ByVal Value As Double) As Single
			If Value > Single.MaxValue Then
				Return Single.MaxValue
			ElseIf Value < Single.MinValue Then
				Return Single.MinValue
			Else
				Return CSng(Value)
			End If
		End Function

		<Extension>
		Private Function ToSingleSafe(ByVal Value As Double, ByVal RoundingDigit As Integer) As Single
			If Value > Single.MaxValue Then
				Return Single.MaxValue
			ElseIf Value < Single.MinValue Then
				Return Single.MinValue
			Else
				Return CSng(Math.Round((Value), RoundingDigit))
			End If
		End Function

		<Extension>
		Private Function ToIntegerSafe(ByVal Value As Long) As Integer
			If Value > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Value < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Value)
			End If
		End Function
#End Region

#Region "StockSymbols"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.StockSymbol)) As YahooAccessData.ISearchKey(Of YahooAccessData.StockSymbol, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.StockSymbol, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.StockSymbol)) As YahooAccessData.ISort(Of YahooAccessData.StockSymbol)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.StockSymbol))
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.StockSymbol),
			ByRef Parent As Stock,
			Optional ByVal IsIgnoreID As Boolean = False)

			Dim ThisStockSymbol As StockSymbol

			For Each ThisStockSymbol In colData
				ThisStockSymbol.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.StockSymbol),
			other As ICollection(Of YahooAccessData.StockSymbol),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.StockSymbol),
			other As ICollection(Of YahooAccessData.StockSymbol)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "Sectors"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.Sector)) As YahooAccessData.ISearchKey(Of YahooAccessData.Sector, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.Sector, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.Sector)) As YahooAccessData.ISort(Of YahooAccessData.Sector)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.Sector))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.Sector),
			ByVal Key As String,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.Sector

			Dim ThisSector As YahooAccessData.Sector

			'this is a slow linear search
			'ThisSector = colData.SingleOrDefault(Function(ThisR As YahooAccessData.Sector) ThisR.Name = Key)
			'this is a very fast search (O(1))
			ThisSector = colData.ToSearch.Find(Key)

			If IsCreateIfEmpty Then
				If ThisSector Is Nothing Then
					ThisSector = New YahooAccessData.Sector(Name:=Key)
					colData.Add(ThisSector)
				End If
			End If
			Return ThisSector
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.Sector),
			ByRef Parent As YahooAccessData.Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisData In colData
				ThisData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.Sector),
			other As ICollection(Of Sector),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.Sector),
			other As ICollection(Of Sector)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "BondRateRecord"
		<Extension()>
		Public Function ToDaily(
			colData As IEnumerable(Of YahooAccessData.BondRateRecord),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of YahooAccessData.BondRateRecord)

			Return Extensions.ToDaily(Of YahooAccessData.BondRateRecord)(colData, DateStart, DateStop)
		End Function

		<Extension()>
		Public Function ToDaily(
			colData As IEnumerable(Of YahooAccessData.BondRateRecord)) As IEnumerable(Of YahooAccessData.BondRateRecord)

			Return Extensions.ToDaily(Of YahooAccessData.BondRateRecord)(colData)
		End Function

		<Extension()>
		Public Function ToDailyBondInterests(
			colData As IEnumerable(Of YahooAccessData.BondRateRecord),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As BondInterests

			Return New BondInterests(colData, DateStart, DateStop)
		End Function

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.BondRateRecord),
			other As ICollection(Of BondRateRecord),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "Industries"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.Industry)) As YahooAccessData.ISearchKey(Of YahooAccessData.Industry, String)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.Industry, String))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.Industry)) As YahooAccessData.ISort(Of YahooAccessData.Industry)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.Industry))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.Industry),
			ByVal Key As String,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.Industry

			Dim ThisIndustry As YahooAccessData.Industry
			'ThisIndustry = colData.SingleOrDefault(Function(ThisR) ThisR.Name = Key)
			ThisIndustry = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisIndustry Is Nothing Then
					ThisIndustry = New YahooAccessData.Industry(Name:=Key)
					colData.Add(ThisIndustry)
				End If
			End If
			Return ThisIndustry
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.Industry),
			ByRef Parent As YahooAccessData.Report,
			Optional ByVal IsIgnoreID As Boolean = False)

			Dim ThisIndustry As Industry
			For Each ThisIndustry In colData
				ThisIndustry.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.Industry),
			other As ICollection(Of Industry),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.Industry),
			other As ICollection(Of Industry)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "RecordQuoteValue"
		<Extension()>
		Public Function CopyFrom(colData() As PriceVol) As PriceVol()
			Dim ThisPriceVols() As PriceVol
			Dim I As Integer

			ReDim ThisPriceVols(0 To colData.Length - 1)
			For I = 0 To colData.Length - 1
				ThisPriceVols(I) = colData(I).CopyFrom
			Next
			Return ThisPriceVols
		End Function

		<Extension()>
		Public Function ToDaily(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of YahooAccessData.RecordQuoteValue)

			Return Extensions.ToDaily(Of YahooAccessData.RecordQuoteValue)(colData, DateStart, DateStop)
		End Function

		<Extension()>
		Public Function ToDaily(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue)) As IEnumerable(Of YahooAccessData.RecordQuoteValue)

			Return Extensions.ToDaily(Of YahooAccessData.RecordQuoteValue)(colData)
		End Function

		<Extension()>
		Public Function ToDailyIntraDay(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of IEnumerable(Of YahooAccessData.RecordQuoteValue))

			Return Extensions.ToDailyIntraDay(Of YahooAccessData.RecordQuoteValue)(colData, DateStart, DateStop)
		End Function

		<Extension()>
		Public Function ToDailyIntraDay(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue)) As IEnumerable(Of IEnumerable(Of YahooAccessData.RecordQuoteValue))

			Return Extensions.ToDailyIntraDay(Of YahooAccessData.RecordQuoteValue)(colData)
		End Function

		<Extension()>
		Public Function ToDaily(
			colData As IEnumerable(Of IEnumerable(Of YahooAccessData.RecordQuoteValue))) As IEnumerable(Of YahooAccessData.RecordQuoteValue)

			Return Extensions.ToDaily(Of YahooAccessData.RecordQuoteValue)(colData)
		End Function

		''' <summary>
		''' Return the daily record prices from the collection of RecordQuoteValue
		''' This is used to create a RecordPrices object that contains the daily prices layout has a complete set of daily prices
		''' for the given date range. The missing data is extrapolated from the previous last day of trading.
		''' If the collection is empty, an exception is thrown.
		''' </summary>
		<Extension()>
		Public Function ToDailyRecordPrices(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As RecordPrices

			Return New RecordPrices(colData, DateStart, DateStop)
		End Function

		<Extension()>
		Public Function ToDailyRecordPrices(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
			ByVal DateStop As Date) As RecordPrices

			Return New RecordPrices(colData, DateStop)
		End Function

		<Extension()>
		Public Function ToDailyIntradayRecordPrices(
			colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As RecordPrices

			Dim ThisResult = colData.ToDailyIntraDay(DateStart, DateStop)
			If ThisResult.Count = 0 Then
				Throw New ArgumentException("Exception Occured in ToDailyIntradayRecordPrices with no data in collection!")
			End If
			Return New RecordPrices(ThisResult(0), DateStart, DateStop)
		End Function


		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.RecordQuoteValue)) As YahooAccessData.ISearchKey(Of YahooAccessData.RecordQuoteValue, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.RecordQuoteValue, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.RecordQuoteValue)) As YahooAccessData.ISort(Of YahooAccessData.RecordQuoteValue)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.RecordQuoteValue))
		End Function
#End Region
#Region "Records"

		<Extension()>
		Public Function RemoveAt(colData As ICollection(Of Record), Index As Integer) As Boolean
			Return DirectCast(colData, LinkedHashSet(Of Record, Date)).RemoveAt(Index)
		End Function

		<Extension()>
		Public Function TryAdd(colData As ICollection(Of Record), Item As Record) As Boolean
			Return DirectCast(colData, LinkedHashSet(Of Record, Date)).TryAdd(Item)
		End Function

		<Extension()>
		Public Function AsDateUpdate(Of T As {IDateUpdate})(colData As ICollection(Of T)) As YahooAccessData.IDateUpdate
			Return DirectCast(colData, YahooAccessData.IDateUpdate)
		End Function

		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.Record)) As YahooAccessData.ISearchKey(Of YahooAccessData.Record, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.Record, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.Record)) As YahooAccessData.ISort(Of YahooAccessData.Record)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.Record))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.Record),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.Record

			Dim ThisRecord As YahooAccessData.Record
			'ThisRecord = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisRecord = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisRecord Is Nothing Then
					ThisRecord = New YahooAccessData.Record(DateUpdate:=Key)
					colData.Add(ThisRecord)
				End If
			End If
			Return ThisRecord
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.Record), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of Record)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.Record),
			ByRef Parent As YahooAccessData.Stock,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisData In colData
				ThisData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.Record),
			other As ICollection(Of Record),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.Record),
			other As ICollection(Of Record)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "RecordsDaily"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.RecordDaily)) As YahooAccessData.ISearchKey(Of YahooAccessData.RecordDaily, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.RecordDaily, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.RecordDaily)) As YahooAccessData.ISort(Of YahooAccessData.RecordDaily)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.RecordDaily))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.RecordDaily),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.RecordDaily

			Dim ThisRecord As YahooAccessData.RecordDaily
			'ThisRecord = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisRecord = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisRecord Is Nothing Then
					ThisRecord = New YahooAccessData.RecordDaily(DateUpdate:=Key)
					colData.Add(ThisRecord)
				End If
			End If
			Return ThisRecord
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.RecordDaily), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of RecordDaily)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.RecordDaily),
			ByRef Parent As YahooAccessData.Stock,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisData In colData
				ThisData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.RecordDaily),
			other As ICollection(Of RecordDaily),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

		''' <summary>
		''' perform a standard shallow test of equality on the collection
		''' </summary>
		''' <param name="colData"></param>
		''' <param name="other"></param>
		''' <returns>True if equals</returns>
		''' <remarks></remarks>
		<Extension()>
		Friend Function EqualsShallow(
			colData As ICollection(Of YahooAccessData.RecordDaily),
			other As ICollection(Of RecordDaily)) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).Equals(other(I)) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "SplitFactors"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.SplitFactor)) As YahooAccessData.ISearchKey(Of YahooAccessData.SplitFactor, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.SplitFactor, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.SplitFactor)) As YahooAccessData.ISort(Of YahooAccessData.SplitFactor)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.SplitFactor))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.SplitFactor),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.SplitFactor

			Dim ThisSplitFactor As YahooAccessData.SplitFactor
			'ThisSplitFactor = colData.SingleOrDefault(Function(ThisR) ThisR.DateDay = Key.Date)
			ThisSplitFactor = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisSplitFactor Is Nothing Then
					ThisSplitFactor = New YahooAccessData.SplitFactor(DateUpdate:=Key.Date)
					colData.Add(ThisSplitFactor)
				End If
			End If
			Return ThisSplitFactor
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.SplitFactor), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of SplitFactor)
			Return colData.
				Where(Function(ThisR) (ThisR.DateDay >= DateStart) And (ThisR.DateDay <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateDay)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.SplitFactor),
			ByRef Parent As YahooAccessData.Stock,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisSplitFactorData In colData
				ThisSplitFactorData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.SplitFactor),
			other As ICollection(Of SplitFactor),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

#End Region
#Region "FinancialHighlights"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.FinancialHighlight)) As YahooAccessData.ISearchKey(Of YahooAccessData.FinancialHighlight, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.FinancialHighlight, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.FinancialHighlight)) As YahooAccessData.ISort(Of YahooAccessData.FinancialHighlight)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.FinancialHighlight))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.FinancialHighlight),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.FinancialHighlight

			Dim ThisFinancialHighlight As YahooAccessData.FinancialHighlight
			'ThisFinancialHighlight = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisFinancialHighlight = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisFinancialHighlight Is Nothing Then
					ThisFinancialHighlight = New YahooAccessData.FinancialHighlight(DateUpdate:=Key)
					colData.Add(ThisFinancialHighlight)
				End If
			End If
			Return ThisFinancialHighlight
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.FinancialHighlight), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of FinancialHighlight)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.FinancialHighlight),
			ByRef Parent As YahooAccessData.Record,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisFinancialHighlightData In colData
				ThisFinancialHighlightData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.FinancialHighlight),
			other As ICollection(Of FinancialHighlight),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function
#End Region
#Region "MarketQuoteDatas"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.MarketQuoteData)) As YahooAccessData.ISearchKey(Of YahooAccessData.MarketQuoteData, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.MarketQuoteData, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.MarketQuoteData)) As YahooAccessData.ISort(Of YahooAccessData.MarketQuoteData)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.MarketQuoteData))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.MarketQuoteData),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.MarketQuoteData

			Dim ThisMarketQuoteData As YahooAccessData.MarketQuoteData
			'ThisMarketQuoteData = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisMarketQuoteData = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisMarketQuoteData Is Nothing Then
					ThisMarketQuoteData = New YahooAccessData.MarketQuoteData(DateUpdate:=Key)
					colData.Add(ThisMarketQuoteData)
				End If
			End If
			Return ThisMarketQuoteData
		End Function

		<Extension()>
		Public Function Items(
			colData As ICollection(Of YahooAccessData.MarketQuoteData),
			ByVal DateStart As Date,
			ByVal DateStop As Date) As IEnumerable(Of MarketQuoteData)

			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.MarketQuoteData),
			ByRef Parent As YahooAccessData.Record,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisMarketQuoteDataData In colData
				ThisMarketQuoteDataData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.MarketQuoteData),
			other As ICollection(Of MarketQuoteData),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

#End Region
#Region "QuoteValues"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.QuoteValue)) As YahooAccessData.ISearchKey(Of YahooAccessData.QuoteValue, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.QuoteValue, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.QuoteValue)) As YahooAccessData.ISort(Of YahooAccessData.QuoteValue)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.QuoteValue))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.QuoteValue),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.QuoteValue

			Dim ThisQuoteValue As YahooAccessData.QuoteValue
			'ThisQuoteValue = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisQuoteValue = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisQuoteValue Is Nothing Then
					ThisQuoteValue = New YahooAccessData.QuoteValue(DateUpdate:=Key)
					colData.Add(ThisQuoteValue)
				End If
			End If
			Return ThisQuoteValue
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.QuoteValue), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of QuoteValue)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.QuoteValue),
			ByRef Parent As YahooAccessData.Record,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisQuoteValueData In colData
				ThisQuoteValueData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.QuoteValue),
			other As ICollection(Of QuoteValue),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

#End Region
#Region "TradeInfoes"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.TradeInfo)) As YahooAccessData.ISearchKey(Of YahooAccessData.TradeInfo, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.TradeInfo, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.TradeInfo)) As YahooAccessData.ISort(Of YahooAccessData.TradeInfo)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.TradeInfo))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.TradeInfo),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.TradeInfo

			Dim ThisTradeInfo As YahooAccessData.TradeInfo
			'ThisTradeInfo = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisTradeInfo = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisTradeInfo Is Nothing Then
					ThisTradeInfo = New YahooAccessData.TradeInfo(DateUpdate:=Key)
					colData.Add(ThisTradeInfo)
				End If
			End If
			Return ThisTradeInfo
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.TradeInfo), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of TradeInfo)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.TradeInfo),
			ByRef Parent As YahooAccessData.Record,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisTradeInfoData In colData
				ThisTradeInfoData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.TradeInfo),
			other As ICollection(Of TradeInfo),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

#End Region
#Region "ValuationMeasures"
		<Extension()>
		Public Function ToSearch(
			colData As ICollection(Of YahooAccessData.ValuationMeasure)) As YahooAccessData.ISearchKey(Of YahooAccessData.ValuationMeasure, Date)
			Return TryCast(colData, YahooAccessData.ISearchKey(Of YahooAccessData.ValuationMeasure, Date))
		End Function

		<Extension()>
		Public Function ToSort(
			colData As ICollection(Of YahooAccessData.ValuationMeasure)) As YahooAccessData.ISort(Of YahooAccessData.ValuationMeasure)
			Return TryCast(colData, YahooAccessData.ISort(Of YahooAccessData.ValuationMeasure))
		End Function

		<Extension()>
		Public Function Item(
			colData As ICollection(Of YahooAccessData.ValuationMeasure),
			ByVal Key As Date,
			Optional ByVal IsCreateIfEmpty As Boolean = False) As YahooAccessData.ValuationMeasure

			Dim ThisValuationMeasure As YahooAccessData.ValuationMeasure
			'ThisValuationMeasure = colData.SingleOrDefault(Function(ThisR) ThisR.DateUpdate = Key)
			ThisValuationMeasure = colData.ToSearch.Find(Key)
			If IsCreateIfEmpty Then
				If ThisValuationMeasure Is Nothing Then
					ThisValuationMeasure = New YahooAccessData.ValuationMeasure(DateUpdate:=Key)
					colData.Add(ThisValuationMeasure)
				End If
			End If
			Return ThisValuationMeasure
		End Function

		<Extension()>
		Public Function Items(colData As ICollection(Of YahooAccessData.ValuationMeasure), ByVal DateStart As Date, ByVal DateStop As Date) As IEnumerable(Of ValuationMeasure)
			Return colData.
				Where(Function(ThisR) (ThisR.DateUpdate >= DateStart) And (ThisR.DateUpdate <= DateStop)).
				OrderBy(Function(ThisR) ThisR.DateUpdate)
		End Function

		<Extension()>
		Friend Sub CopyDeep(
			colData As ICollection(Of YahooAccessData.ValuationMeasure),
			ByRef Parent As YahooAccessData.Record,
			Optional ByVal IsIgnoreID As Boolean = False)

			For Each ThisValuationMeasureData In colData
				ThisValuationMeasureData.CopyDeep(Parent, IsIgnoreID)
			Next
		End Sub

		<Extension()>
		Friend Function EqualsDeep(
			colData As ICollection(Of YahooAccessData.ValuationMeasure),
			other As ICollection(Of ValuationMeasure),
			Optional ByVal IsIgnoreID As Boolean = False) As Boolean

			Dim I As Integer

			If colData.Count <> other.Count Then Return False
			For I = 0 To colData.Count - 1
				If colData(I).EqualsDeep(other(I), IsIgnoreID) = False Then Return False
			Next
			Return True
		End Function

#End Region
#Region "clsEqualsDeep"
		'Private Class clsEqualsDeep(Of T)
		'	Function EqualsDeep(
		'		colData As ICollection(Of T),
		'		other As ICollection(Of T)) As Boolean

		'		Dim I As Integer
		'		Dim ThisDuplicate = New Duplicate(Of T)
		'		Dim ThisX As T = Nothing
		'		Dim ThisY As T = Nothing

		'		If colData.Count <> other.Count Then Return False
		'		For I = 0 To colData.Count
		'			ThisX = colData.ElementAt(I)
		'			ThisY = other.ElementAt(I)
		'			If ThisDuplicate.Equals(ThisX, ThisY) = False Then Return False
		'		Next
		'		Return True
		'	End Function
		'End Class
#End Region
#Region "Exception"
		<Extension()>
		Public Function MessageAll(ByVal Exception As System.Exception, Optional InnerSeparator As String = vbCr) As String
			Dim ThisException As System.Exception
			Dim ThisMessage As String = ""
			Dim ThisMessageLast As String = ""

			ThisException = Exception
			Do Until ThisException Is Nothing
				If ThisException.Message <> ThisMessageLast Then
					ThisMessageLast = ThisException.Message
					If ThisMessage = "" Then
						ThisMessage = ThisException.Message
					Else
						ThisMessage = ThisMessage + InnerSeparator + ThisException.Message
					End If
				End If
				ThisException = ThisException.InnerException
			Loop
			ThisException = Nothing
			Return ThisMessage
		End Function

		<Extension()>
		Public Function ToList(ByVal Exception As Exception) As List(Of Exception)
			Dim ThisList As New List(Of Exception)
			Dim ThisException As Exception

			ThisException = Exception
			Do Until ThisException Is Nothing
				ThisList.Add(New Exception(ThisException.Message))
				ThisException = ThisException.InnerException
			Loop
			Return ThisList
		End Function
#End Region
		Public Function FileHeaderRead(
			ByVal FileName As String,
			ByRef HeaderInfoDefault As List(Of HeaderInfo),
			ByRef Exception As Exception) As List(Of HeaderInfo)

			Dim ThisListOfData As List(Of HeaderInfo) = Nothing

			'MyMutex.WaitOne()
			Try
				ThisListOfData = FileListRead(Of HeaderInfo)(FileName, HeaderInfoDefault, Exception)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				'MyMutex.ReleaseMutex()
			End Try
			Return ThisListOfData
		End Function

		Public Function FileListRead(Of T)(
			ByVal FileName As String,
			ByRef DataDefault As List(Of T),
			ByRef Exception As Exception) As List(Of T)

			Dim ThisListOfData As List(Of T) = Nothing

			'MyMutex.WaitOne()
			Try
				ThisListOfData = FileListReadLocal(Of T)(FileName, DataDefault, Exception)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				'MyMutex.ReleaseMutex()
			End Try
			Return ThisListOfData
		End Function

		''' <summary>
		''' Save the data in an xml or json file format based on the file extension *.json or *.xml
		''' If the extension is json but the file does not exist the function will automatically try to read the same file
		''' but with the xml extension and if successful will convert the file to Jason file. The xml file is not changed by the operation
		''' This characteristic is needed to help to move the xml file of this application to only the json standard providing a better visual
		''' for for inspection on web browser. 
		''' </summary>
		''' <typeparam name="T">The type in the list</typeparam>
		''' <param name="FileName"></param>
		''' <param name="DataDefault"></param>
		''' <param name="Exception"></param>
		''' <returns></returns>
		Private Function FileListReadLocal(Of T)(
			ByVal FileName As String,
			ByRef DataDefault As List(Of T),
			ByRef Exception As Exception) As List(Of T)

			Dim ThisException As Exception = Nothing
			Dim ThisXmlSerializer As XmlSerializer
			Dim ThisTextReader As System.IO.TextReader
			Dim ThisListInfo As List(Of T)
			Dim ThisFileExtension As String = Path.GetExtension(FileName).ToLower
			Dim ThisFileNameXML As String = Path.ChangeExtension(FileName, ".xml").ToLower
			Dim ThisFileData As String

			Select Case ThisFileExtension
				Case ".json"
					If My.Computer.FileSystem.FileExists(FileName) = False Then
						'check to read the file using an xml format
						'if it work save it under the correct format and continue
						ThisListInfo = FileListReadLocal(Of T)(ThisFileNameXML, DataDefault, Nothing)

						FileListSaveLocal(Of T)(FileName, ThisListInfo)
						Return ThisListInfo
					End If
					Try
						ThisFileData = File.ReadAllText(FileName)
						ThisListInfo = JsonConvert.DeserializeObject(Of List(Of T))(ThisFileData)
						Return ThisListInfo
					Catch ex As Exception
						'try to read as an xml file
						Try
							ThisListInfo = FileListRead(Of T)(ThisFileNameXML, DataDefault, Nothing)
							Return ThisListInfo
						Catch ex1 As Exception
							Throw ex1
						End Try
					End Try
				Case ".xml"
					If My.Computer.FileSystem.FileExists(FileName) = False Then
						Throw New FileNotFoundException
					End If
					Try
						ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
						'StreamReader is designed for character input in a particular encoding, whereas the Stream class
						'is designed for byte input and output. Use StreamReader for reading lines of information
						'From a standard text file.
						ThisTextReader = New StreamReader(FileName)
						ThisListInfo = CType(ThisXmlSerializer.Deserialize(ThisTextReader), List(Of T))
						ThisTextReader.Dispose()
						Return ThisListInfo
					Catch ex As Exception
						Throw ex
					End Try
				Case Else
					Throw New InvalidDataException($"File name extension is not supported:{vbCr}{FileName}")
			End Select
		End Function

		''' <summary>
		''' Save the file under and xml or json file format in function of the extension provided in the file name
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="FileName"></param>
		''' <param name="Data"></param>
		Private Sub FileListSaveLocal(Of T)(ByVal FileName As String, ByRef Data As List(Of T))
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing

			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0}", ThisPath), ex)
					End Try
				End If
			End With
			Dim ThisFileNameToCopy = Path.ChangeExtension(FileName, $"{Path.GetExtension(FileName)}.{"bak"}")
			If System.IO.File.Exists(FileName) Then
				'make a backup file
				Try
					System.IO.File.Copy(FileName, ThisFileNameToCopy, overwrite:=True)
				Catch ex As Exception
					Throw New Exception(String.Format("Unable to copy to a backup file {0}", ThisFileNameToCopy), ex)
				End Try
				Try
					System.IO.File.Delete(FileName)
				Catch ex As Exception
					Throw New Exception(String.Format("Unable to delete the file {0}", ThisFileNameToCopy), ex)
				End Try
			End If
			Select Case Path.GetExtension(FileName)
				Case ".json"
					Try
						Dim ThisJsonResult = JsonConvert.SerializeObject(Data, Formatting.Indented)
						File.WriteAllText(FileName, ThisJsonResult)
					Catch ex As Exception
						Throw ex
					End Try
					'Creates a new file, write the contents to the file, and then closes the file.
					'If the target file already exists, it is truncated and overwritten.
				Case ".Xml"
					Try
						ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create Xml Serializer."), ex)
					End Try
					Dim ThisTextWriter = New StreamWriter(FileName)
					Try
						ThisXmlSerializer.Serialize(ThisTextWriter, Data)
						'ThisSharpSerializer.Serialize(Data, filename:=FileName)
						ThisTextWriter.Dispose()
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to save to file {0}", FileName), ex)
						'restore the previous file from the backup file
					End Try
			End Select
		End Sub
	End Module
End Namespace
#End Region  'Extension
