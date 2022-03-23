#Region "Imports"
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System
Imports System.IO
#End Region
<Serializable()>
Partial Public Class QuoteValue
	Implements IEquatable(Of QuoteValue)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of QuoteValue)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New(ByVal DateUpdate As Date)
		With Me
			.DateUpdate = DateUpdate
			.LastTradeDate = .DateUpdate
			.LastTradeRealtimeWithTime = .DateUpdate
      .DividendPayDate = ReportDate.DateNullValue
			.ExDividendDate = ReportDate.DateNullValue
			.StockExchange = ""
		End With
		If MyListHeaderInfo Is Nothing Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		End If
	End Sub

	Public Sub New()
    Me.New(Now)
	End Sub

	Public Sub New(ByRef Parent As Record, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Record = Parent
			.RecordID = .Record.ID
			.Record.QuoteValues.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Record, Optional ByVal IsIgnoreID As Boolean = False) As QuoteValue
		Dim ThisQuoteValue = New QuoteValue

		With ThisQuoteValue
			If IsIgnoreID = False Then .ID = Me.ID
			.AfterHoursChangeRealtime = Me.AfterHoursChangeRealtime
			.AnnualizedGain = Me.AnnualizedGain
			.Ask = Me.Ask
			.AskRealTime = Me.AskRealTime
			.AskSize = Me.AskSize
			.AverageDailyVolume = Me.AverageDailyVolume
			.Bid = Me.Bid
			.BidRealtime = Me.BidRealtime
			.BidSize = Me.BidSize
			.BookValue = Me.BookValue
			.DividendPayDate = Me.DividendPayDate
			.DividendShare = Me.DividendShare
			.DividendYield = Me.DividendYield
			.EarningsShare = Me.EarningsShare
			.EBITDA = Me.EBITDA
			.EPSEstimateCurrentYear = Me.EPSEstimateCurrentYear
			.EPSEstimateNextQuarter = Me.EPSEstimateNextQuarter
			.EPSEstimateNextYear = Me.EPSEstimateNextYear
			.ExDividendDate = Me.ExDividendDate
			.FloatShares = Me.FloatShares
			.HoldingsValue = Me.HoldingsValue
			.HoldingsValueRealtime = Me.HoldingsValueRealtime
			.LastTradeDate = Me.LastTradeDate
			.LastTradePriceOnly = Me.LastTradePriceOnly
			.LastTradeRealtimeWithTime = Me.LastTradeRealtimeWithTime
			.LastTradeSize = Me.LastTradeSize
			.MarketCapitalization = Me.MarketCapitalization
			.MarketCapRealtime = Me.MarketCapRealtime
			.OneyrTargetPrice = Me.OneyrTargetPrice
			.OrderBookRealtime = Me.OrderBookRealtime
			.PEGRatio = Me.PEGRatio
			.PERatio = Me.PERatio
			.PERatioRealtime = Me.PERatioRealtime
			.PriceBook = Me.PriceBook
			.PriceEPSEstimateCurrentYear = Me.PriceEPSEstimateCurrentYear
			.PriceEPSEstimateNextYear = Me.PriceEPSEstimateNextYear
			.PriceSales = Me.PriceSales
			.SharesOwned = Me.SharesOwned
			.ShortRatio = Me.ShortRatio
			.TickerTrend = Me.TickerTrend
			.StockExchange = Me.StockExchange
			.DateUpdate = Me.DateUpdate
      .Record = Parent
      If .Record IsNot Nothing Then
        .RecordID = .Record.ID
        .Record.QuoteValues.Add(ThisQuoteValue)
      End If
		End With
		Return ThisQuoteValue
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
		End Set
	End Property

	Public Overrides Function ToString() As String
		Return String.Format("Object:{0}:{1}:{2}", TypeName(Me), Me.KeyID, Me.KeyValue)
	End Function

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
	Public Function CompareTo(other As QuoteValue) As Integer Implements System.IComparable(Of QuoteValue).CompareTo
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
			Return Me.DateUpdate
		End Get
		Set(value As Date)

		End Set
	End Property
#End Region
#Region "IMemoryStream"

	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
		Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
	End Sub

	Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
		Dim ThisBinaryWriter As New BinaryWriter(Stream)
		Dim ThisListException As List(Of Exception)
		Dim ThisException As Exception

		With ThisBinaryWriter
			.Write(VERSION_MEMORY_STREAM)
			.Write(Me.ID)
			ThisListException = Me.Exception.ToList
			ThisListException.Reverse()
			.Write(ThisListException.Count)
			For Each ThisException In ThisListException
				.Write(ThisException.Message)
			Next
			.Write(Me.AfterHoursChangeRealtime)
			.Write(Me.AnnualizedGain)
			.Write(Me.Ask)
			.Write(Me.AskRealTime)
			.Write(Me.AskSize)
			.Write(Me.AverageDailyVolume)
			.Write(Me.Bid)
			.Write(Me.BidRealtime)
			.Write(Me.BidSize)
			.Write(Me.BookValue)
			.Write(Me.DividendPayDate.ToBinary)
			.Write(Me.DividendShare)
			.Write(Me.DividendYield)
			.Write(Me.EarningsShare)
			.Write(Me.EBITDA)
			.Write(Me.EPSEstimateCurrentYear)
			.Write(Me.EPSEstimateNextQuarter)
			.Write(Me.EPSEstimateNextYear)
			.Write(Me.ExDividendDate.ToBinary)
			.Write(Me.FloatShares)
			.Write(Me.HoldingsValue)
			.Write(Me.HoldingsValueRealtime)
			.Write(Me.LastTradeDate.ToBinary)
			.Write(Me.LastTradePriceOnly)
			.Write(Me.LastTradeRealtimeWithTime.ToBinary)
			.Write(Me.LastTradeSize)
			.Write(Me.MarketCapitalization)
			.Write(Me.MarketCapRealtime)
			.Write(Me.OneyrTargetPrice)
			.Write(Me.OrderBookRealtime)
			.Write(Me.PEGRatio)
			.Write(Me.PERatio)
			.Write(Me.PERatioRealtime)
			.Write(Me.PriceBook)
			.Write(Me.PriceEPSEstimateCurrentYear)
			.Write(Me.PriceEPSEstimateNextYear)
			.Write(Me.PriceSales)
			.Write(Me.SharesOwned)
			.Write(Me.ShortRatio)
			.Write(Me.TickerTrend)
			.Write(Me.RecordID)
			.Write(Me.StockExchange)
			.Write(Me.DateUpdate.ToBinary)
		End With
	End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
    Dim ThisVersion As Single

    With ThisBinaryReader
      ThisVersion = .ReadSingle
      Me.ID = .ReadInt32
      Me.Exception = Nothing
      Dim I As Integer
      For I = 1 To .ReadInt32
        MyException = New Exception(.ReadString, MyException)
      Next
      Me.AfterHoursChangeRealtime = .ReadSingle
      Me.AnnualizedGain = .ReadSingle
      Me.Ask = .ReadSingle
      Me.AskRealTime = .ReadSingle
      Me.AskSize = .ReadInt32
      Me.AverageDailyVolume = .ReadInt32
      Me.Bid = .ReadSingle
      Me.BidRealtime = .ReadSingle
      Me.BidSize = .ReadInt32
      Me.BookValue = .ReadSingle
      Me.DividendPayDate = DateTime.FromBinary(.ReadInt64)
      Me.DividendShare = .ReadSingle
      Me.DividendYield = .ReadSingle
      Me.EarningsShare = .ReadSingle
      Me.EBITDA = .ReadSingle
      Me.EPSEstimateCurrentYear = .ReadSingle
      Me.EPSEstimateNextQuarter = .ReadSingle
      Me.EPSEstimateNextYear = .ReadSingle
      Me.ExDividendDate = DateTime.FromBinary(.ReadInt64)
      Me.FloatShares = .ReadInt32
      Me.HoldingsValue = .ReadSingle
      Me.HoldingsValueRealtime = .ReadSingle
      Me.LastTradeDate = DateTime.FromBinary(.ReadInt64)
      Me.LastTradePriceOnly = .ReadSingle
      Me.LastTradeRealtimeWithTime = DateTime.FromBinary(.ReadInt64)
      Me.LastTradeSize = .ReadInt32
      Me.MarketCapitalization = .ReadSingle
      Me.MarketCapRealtime = .ReadSingle
      Me.OneyrTargetPrice = .ReadSingle
      Me.OrderBookRealtime = .ReadSingle
      Me.PEGRatio = .ReadSingle
      Me.PERatio = .ReadSingle
      Me.PERatioRealtime = .ReadSingle
      Me.PriceBook = .ReadSingle
      Me.PriceEPSEstimateCurrentYear = .ReadSingle
      Me.PriceEPSEstimateNextYear = .ReadSingle
      Me.PriceSales = .ReadSingle
      Me.SharesOwned = .ReadInt32
      Me.ShortRatio = .ReadSingle
      Me.TickerTrend = .ReadSingle
      Me.RecordID = .ReadInt32
      Me.StockExchange = .ReadString
      Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
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
	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of QuoteValue)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = False})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = False})
      .Add(New HeaderInfo With {.Name = "AfterHoursChangeRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AnnualizedGain", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Ask", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AskRealTime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AskSize", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "AverageDailyVolume", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Bid", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "BidRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "BidSize", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "BookValue", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DividendPayDate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DividendShare", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DividendYield", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EarningsShare", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EBITDA", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EPSEstimateCurrentYear", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EPSEstimateNextQuarter", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "EPSEstimateNextYear", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ExDividendDate", .Title = .Name, .Format = "{0:d}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "FloatShares", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "HoldingsValue", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "HoldingsValueRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastTradeRealtimeWithTime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastTradeSize", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "MarketCapitalization", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "MarketCapRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastTradeDate", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LastTradePriceOnly", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OneyrTargetPrice", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "OrderBookRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PEGRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PERatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PERatioRealtime", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceBook", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceEPSEstimateCurrentYear", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceEPSEstimateNextYear", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceSales", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesOwned", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ShortRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "TickerTrend", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "StockExchange", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As QuoteValue, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .RecordID <> Me.RecordID Then Return False
			End If
			If .AfterHoursChangeRealtime <> Me.AfterHoursChangeRealtime Then Return False
			If .AnnualizedGain <> Me.AnnualizedGain Then Return False
			If .Ask <> Me.Ask Then Return False
			If .AskRealTime <> Me.AskRealTime Then Return False
			If .AskSize <> Me.AskSize Then Return False
			If .AverageDailyVolume <> Me.AverageDailyVolume Then Return False
			If .Bid <> Me.Bid Then Return False
			If .BidRealtime <> Me.BidRealtime Then Return False
			If .BidSize <> Me.BidSize Then Return False
			If .BookValue <> Me.BookValue Then Return False
			If .DividendPayDate <> Me.DividendPayDate Then Return False
			If .DividendShare <> Me.DividendShare Then Return False
			If .DividendYield <> Me.DividendYield Then Return False
			If .EarningsShare <> Me.EarningsShare Then Return False
			If .EBITDA <> Me.EBITDA Then Return False
			If .EPSEstimateCurrentYear <> Me.EPSEstimateCurrentYear Then Return False
			If .EPSEstimateNextQuarter <> Me.EPSEstimateNextQuarter Then Return False
			If .EPSEstimateNextYear <> Me.EPSEstimateNextYear Then Return False
			If .ExDividendDate <> Me.ExDividendDate Then Return False
			If .FloatShares <> Me.FloatShares Then Return False
			If .HoldingsValue <> Me.HoldingsValue Then Return False
			If .HoldingsValueRealtime <> Me.HoldingsValueRealtime Then Return False
			If .LastTradeDate <> Me.LastTradeDate Then Return False
			If .LastTradePriceOnly <> Me.LastTradePriceOnly Then Return False
			If .LastTradeRealtimeWithTime <> Me.LastTradeRealtimeWithTime Then Return False
			If .LastTradeSize <> Me.LastTradeSize Then Return False
			If .MarketCapitalization <> Me.MarketCapitalization Then Return False
			If .MarketCapRealtime <> Me.MarketCapRealtime Then Return False
			If .OneyrTargetPrice <> Me.OneyrTargetPrice Then Return False
			If .OrderBookRealtime <> Me.OrderBookRealtime Then Return False
			If .PEGRatio <> Me.PEGRatio Then Return False
			If .PERatio <> Me.PERatio Then Return False
			If .PERatioRealtime <> Me.PERatioRealtime Then Return False
			If .PriceBook <> Me.PriceBook Then Return False
			If .PriceEPSEstimateCurrentYear <> Me.PriceEPSEstimateCurrentYear Then Return False
			If .PriceEPSEstimateNextYear <> Me.PriceEPSEstimateNextYear Then Return False
			If .PriceSales <> Me.PriceSales Then Return False
			If .SharesOwned <> Me.SharesOwned Then Return False
			If .ShortRatio <> Me.ShortRatio Then Return False
			If .TickerTrend <> Me.TickerTrend Then Return False
			If .StockExchange <> Me.StockExchange Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As QuoteValue) As Boolean Implements IEquatable(Of QuoteValue).Equals
		If other Is Nothing Then Return False
		If Me.DateUpdate = other.DateUpdate Then
			If Me.Record IsNot Nothing Then
				Return Me.Record.Equals(other.Record)
			Else
				Return False
			End If
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If (TypeOf obj Is QuoteValue) Then
			Return Me.Equals(DirectCast(obj, QuoteValue))
		Else
			Return (False)
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		If Me.Record IsNot Nothing Then
			Return Me.Record.GetHashCode()
		Else
			Return Me.DateUpdate.GetHashCode()
		End If
	End Function
#End Region
#Region "IDateUpdate"
	Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
		Get
			Return Me.DateUpdate
		End Get
		Set(value As Date)
		End Set
	End Property

	Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
		Get
			Return Me.DateUpdate
		End Get
		Set(value As Date)
		End Set
	End Property

	Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
		Get
			Return Me.DateUpdate
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
#End Region	 'IDateUpdate
End Class

#Region "EqualityComparerOfQuoteValue"
	<Serializable()>
	Friend Class EqualityComparerOfQuoteValue
		Implements IEqualityComparer(Of QuoteValue)

		Public Overloads Function Equals(x As QuoteValue, y As QuoteValue) As Boolean Implements IEqualityComparer(Of QuoteValue).Equals
			If (x Is Nothing) And (y Is Nothing) Then
				Return True
			ElseIf (x Is Nothing) Xor (y Is Nothing) Then
				Return False
			Else
				If x.DateUpdate = y.DateUpdate Then
					If (x.Record Is Nothing) And (y.Record Is Nothing) Then
						Return True
					ElseIf (x.Record Is Nothing) Xor (y.Record Is Nothing) Then
						Return False
					Else
						If x.Record.Equals(y) Then
							Return True
						Else
							Return False
						End If
					End If
				Else
					Return False
				End If
			End If
		End Function

		Public Overloads Function GetHashCode(obj As QuoteValue) As Integer Implements IEqualityComparer(Of QuoteValue).GetHashCode
			If obj IsNot Nothing Then
				Return obj.DateUpdate.GetHashCode
			Else
				Return obj.GetHashCode
			End If
		End Function
	End Class
#End Region

#Region "template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class QuoteValue
  Public Property ID As Integer
  Public Property AfterHoursChangeRealtime As Single
  Public Property AnnualizedGain As Single
  Public Property Ask As Single
  Public Property AskRealTime As Single
  Public Property AskSize As Integer
  Public Property AverageDailyVolume As Integer
  Public Property Bid As Single
  Public Property BidRealtime As Single
  Public Property BidSize As Integer
  Public Property BookValue As Single
  Public Property DividendPayDate As Date
  Public Property DividendShare As Single
  Public Property DividendYield As Single
  Public Property EarningsShare As Single
  Public Property EBITDA As Single
  Public Property EPSEstimateCurrentYear As Single
  Public Property EPSEstimateNextQuarter As Single
  Public Property EPSEstimateNextYear As Single
  Public Property ExDividendDate As Date
  Public Property FloatShares As Integer
  Public Property HoldingsValue As Single
  Public Property HoldingsValueRealtime As Single
  Public Property LastTradeDate As Date
  Public Property LastTradePriceOnly As Single
  Public Property LastTradeRealtimeWithTime As Date
  Public Property LastTradeSize As Integer
  Public Property MarketCapitalization As Single
  Public Property MarketCapRealtime As Single
  Public Property OneyrTargetPrice As Single
  Public Property OrderBookRealtime As Single
  Public Property PEGRatio As Single
  Public Property PERatio As Single
  Public Property PERatioRealtime As Single
  Public Property PriceBook As Single
  Public Property PriceEPSEstimateCurrentYear As Single
  Public Property PriceEPSEstimateNextYear As Single
  Public Property PriceSales As Single
  Public Property SharesOwned As Integer
  Public Property ShortRatio As Single
  Public Property TickerTrend As Single
  Public Property RecordID As Integer
  Public Property StockExchange As String
  Public Property DateUpdate As Date

  Public Overridable Property Record As Record
End Class
#End Region