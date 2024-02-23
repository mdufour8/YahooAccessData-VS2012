#Region "Imports"
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System
Imports System.IO
#End Region

<Serializable()>
Partial Public Class MarketQuoteData
	Implements IEquatable(Of MarketQuoteData)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of MarketQuoteData)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)
	Private MyID As Integer

	Public Sub New(ByVal DateUpdate As Date)
		With Me
			.DateUpdate = DateUpdate
		End With
		If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
			MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		Else
			MyListHeaderInfo = ListOfHeader()
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
			.Record.MarketQuoteDatas.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Record, Optional ByVal IsIgnoreID As Boolean = False) As MarketQuoteData
		Dim ThisMarketQuoteData = New MarketQuoteData

		With ThisMarketQuoteData
			If IsIgnoreID = False Then .ID = Me.ID
			.DividendYieldPercent = Me.DividendYieldPercent
			.LongTermDeptToEquity = Me.LongTermDeptToEquity
			.MarketCapitalizationInMillion = Me.MarketCapitalizationInMillion
			.NetProfitMarginPercent = Me.NetProfitMarginPercent
			.PriceEarningsRatio = Me.PriceEarningsRatio
			.PriceToBookValue = Me.PriceToBookValue
			.PriceToFreeCashFlow = Me.PriceToFreeCashFlow
			.ReturnOnEquityPercent = Me.ReturnOnEquityPercent
			.DateUpdate = Me.DateUpdate
			.Record = Parent
			.RecordID = .Record.ID
			.Record.MarketQuoteDatas.Add(ThisMarketQuoteData)
		End With
		Return ThisMarketQuoteData
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
	Public Function CompareTo(other As MarketQuoteData) As Integer Implements System.IComparable(Of MarketQuoteData).CompareTo
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
#Region "IFormatData"
	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of MarketQuoteData)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DividendYieldPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "LongTermDeptToEquity", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "MarketCapitalizationInMillion", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "NetProfitMarginPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceEarningsRatio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceToBookValue", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "PriceToFreeCashFlow", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ReturnOnEquityPercent", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IMemoryStream"
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
			Me.DividendYieldPercent = .ReadSingle
			Me.LongTermDeptToEquity = .ReadSingle
			Me.MarketCapitalizationInMillion = .ReadSingle
			Me.NetProfitMarginPercent = .ReadSingle
			Me.PriceEarningsRatio = .ReadSingle
			Me.PriceToBookValue = .ReadSingle
			Me.PriceToFreeCashFlow = .ReadSingle
			Me.ReturnOnEquityPercent = .ReadSingle
			Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
			Me.RecordID = .ReadInt32
    End With
    ThisBinaryReader.Dispose()
	End Sub

	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
		Call SerializeLoadFrom(Stream, False)
	End Sub

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
			.Write(Me.DividendYieldPercent)
			.Write(Me.LongTermDeptToEquity)
			.Write(Me.MarketCapitalizationInMillion)
			.Write(Me.NetProfitMarginPercent)
			.Write(Me.PriceEarningsRatio)
			.Write(Me.PriceToBookValue)
			.Write(Me.PriceToFreeCashFlow)
			.Write(Me.ReturnOnEquityPercent)
			.Write(Me.DateUpdate.ToBinary)
			.Write(Me.RecordID)
		End With
	End Sub

	Public Sub SerializeLoadFrom(ByRef Data() As Byte) Implements IMemoryStream.SerializeLoadFrom
		Dim ThisStream As Stream = New System.IO.MemoryStream(Data, writable:=True)
		Me.SerializeLoadFrom(ThisStream)
	End Sub

	Public Function SerializeSaveTo() As Byte() Implements IMemoryStream.SerializeSaveTo
		Dim ThisStream As Stream = New System.IO.MemoryStream
		Dim ThisBinaryReader As New BinaryReader(ThisStream)
		Me.SerializeSaveTo(ThisStream)
		Return ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
	End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As MarketQuoteData, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .RecordID <> Me.RecordID Then Return False
			End If
			If .DividendYieldPercent <> Me.DividendYieldPercent Then Return False
			If .LongTermDeptToEquity <> Me.LongTermDeptToEquity Then Return False
			If .MarketCapitalizationInMillion <> Me.MarketCapitalizationInMillion Then Return False
			If .NetProfitMarginPercent <> Me.NetProfitMarginPercent Then Return False
			If .PriceEarningsRatio <> Me.PriceEarningsRatio Then Return False
			If .PriceToBookValue <> Me.PriceToBookValue Then Return False
			If .PriceToFreeCashFlow <> Me.PriceToFreeCashFlow Then Return False
			If .ReturnOnEquityPercent <> Me.ReturnOnEquityPercent Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As MarketQuoteData) As Boolean Implements IEquatable(Of MarketQuoteData).Equals
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
		If (TypeOf obj Is MarketQuoteData) Then
			Return Me.Equals(DirectCast(obj, MarketQuoteData))
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
#End Region
End Class

#Region "EqualityComparerOfMarketQuoteData"
	<Serializable()>
	Friend Class EqualityComparerOfMarketQuoteData
		Implements IEqualityComparer(Of MarketQuoteData)

		Public Overloads Function Equals(x As MarketQuoteData, y As MarketQuoteData) As Boolean Implements IEqualityComparer(Of MarketQuoteData).Equals
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

		Public Overloads Function GetHashCode(obj As MarketQuoteData) As Integer Implements IEqualityComparer(Of MarketQuoteData).GetHashCode
			If obj IsNot Nothing Then
				Return obj.DateUpdate.GetHashCode
			Else
				Return obj.GetHashCode
			End If
		End Function
	End Class
#End Region


#Region "MarketQuoteData template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class MarketQuoteData
  Public Property ID As Integer
  Public Property DividendYieldPercent As Single
  Public Property LongTermDeptToEquity As Single
  Public Property MarketCapitalizationInMillion As Single
  Public Property NetProfitMarginPercent As Single
  Public Property PriceEarningsRatio As Single
  Public Property PriceToBookValue As Single
  Public Property PriceToFreeCashFlow As Single
  Public Property ReturnOnEquityPercent As Single
  Public Property RecordID As Integer
  Public Property DateUpdate As Date

  Public Overridable Property Record As Record

End Class


#End Region