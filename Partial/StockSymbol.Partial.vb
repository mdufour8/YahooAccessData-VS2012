#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Public Class StockSymbol
	Implements IEquatable(Of StockSymbol)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of StockSymbol)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New()
		With Me
			.DateUpdate = Now
			.Exchange = ""
			.Name = ""
			.SymbolNew = ""
			.Symbol = ""
		End With
		If MyListHeaderInfo Is Nothing Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		End If
	End Sub

	Public Sub New(ByRef Parent As Stock, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Stock = Parent
			.StockID = .Stock.ID
			.Stock.StockSymbols.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Stock, Optional ByVal IsIgnoreID As Boolean = False) As StockSymbol
		Dim ThisStockSymbol = New StockSymbol

		With ThisStockSymbol
			If IsIgnoreID = False Then .ID = Me.ID
			.Name = Me.Name
			.Symbol = Me.Symbol
			.SymbolNew = Me.SymbolNew
			.DateUpdate = Me.DateUpdate
			.Exchange = Me.Exchange
			.Stock = Parent
			.IndustryID = .Stock.Industry.ID
			.SectorID = .Stock.Sector.ID
			.Stock.StockSymbols.Add(ThisStockSymbol)
		End With
		Return ThisStockSymbol
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
	Public Function CompareTo(other As StockSymbol) As Integer Implements System.IComparable(Of StockSymbol).CompareTo
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
		Return Extensions.ToStingOfData(Of StockSymbol)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Symbol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SymbolNew", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Exchange", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "IndustryID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SectorID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IMemoryStream"
	Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
		Me.SerializeLoadFrom(Stream, False)
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
			Me.Name = .ReadString
			Me.Symbol = .ReadString
			Me.SymbolNew = .ReadString
			Me.Exchange = .ReadString
			Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
			Me.IndustryID = .ReadInt32
			Me.SectorID = .ReadInt32
			Me.StockID = .ReadInt32
    End With
    ThisBinaryReader.Dispose()
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
			.Write(Me.Name)
			.Write(Me.Symbol)
			.Write(Me.SymbolNew)
			.Write(Me.Exchange)
			.Write(Me.DateUpdate.ToBinary)
			.Write(Me.IndustryID)
			.Write(Me.SectorID)
			.Write(Me.StockID)
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
	Public Function EqualsDeep(other As StockSymbol, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		If other Is Nothing Then Return False
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .IndustryID <> Me.IndustryID Then Return False
				If .SectorID <> Me.SectorID Then Return False
				If .SectorID <> Me.SectorID Then Return False
			End If
			If .Symbol <> Me.Symbol Then Return False
			If .SymbolNew <> Me.SymbolNew Then Return False
			If .Name <> Me.Name Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
			If .Exchange <> Me.Exchange Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As StockSymbol) As Boolean Implements IEquatable(Of StockSymbol).Equals
		If other Is Nothing Then Return False
		If Me.Symbol = other.Symbol Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If obj Is Nothing Then Return False
		If (TypeOf obj Is StockSymbol) Then
			Return Me.Equals(DirectCast(obj, StockSymbol))
		Else
			Return False
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		Return Me.Symbol.GetHashCode
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

#Region "Template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class StockSymbol
  Public Property ID As Integer
  Public Property DateUpdate As Date
  Public Property Exchange As String
  Public Property Name As String
  Public Property IndustryID As Integer
  Public Property SectorID As Integer
  Public Property Symbol As String
  Public Property SymbolNew As String
  Public Property StockID As Integer

  Public Overridable Property Stock As Stock

End Class
#End Region