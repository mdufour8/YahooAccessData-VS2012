#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Partial Public Class StockError
	Implements IEquatable(Of StockError)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of StockError)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New()
		With Me
			.DateUpdate = Now
			.Description = ""
			.Symbol = ""
		End With
		If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
			MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		Else
			MyListHeaderInfo = ListOfHeader()
		End If
	End Sub

	Public Sub New(ByRef Parent As Stock, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Stock = Parent
			.StockID = .Stock.ID
			.Stock.StockErrors.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Stock, Optional ByVal IsIgnoreID As Boolean = False) As StockError
		Dim ThisStockError = New StockError

		With ThisStockError
			If IsIgnoreID = False Then .ID = Me.ID
			.Symbol = Me.Symbol
			.DateUpdate = Me.DateUpdate
			.Description = Me.Description
			.Exception = Me.Exception
			.Stock = Parent
			.StockID = .Stock.ID
			.Stock.StockErrors.Add(ThisStockError)
		End With
		Return ThisStockError
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
	Public Function CompareTo(other As StockError) As Integer Implements System.IComparable(Of StockError).CompareTo
		Return Me.KeyValue.CompareTo(other.KeyValue)
	End Function
#End Region
#Region "IRegisterKey"
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
			.Write(Me.DateUpdate.ToBinary)
			.Write(Me.Symbol)
			.Write(Me.Description)
			.Write(Me.StockID)
			If Me.StockID = 0 Then
				Me.StockID = 0
			End If
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
			Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
			Me.Symbol = .ReadString
			Me.Description = .ReadString
			Me.StockID = .ReadInt32
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
		Me.SerializeSaveTo(ThisStream)
		Return ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
	End Function
#End Region
#Region "IFormatData"
	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of StockError)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = "DateUpdate", .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Symbol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Description", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As StockError, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		If other Is Nothing Then Return False
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .StockID <> Me.StockID Then Return False
			End If
			If .Symbol <> Me.Symbol Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
			If .Description <> Me.Description Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As StockError) As Boolean Implements IEquatable(Of StockError).Equals
		If other Is Nothing Then Return False
		If Me.Symbol = other.Symbol Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If obj Is Nothing Then Return False
		If (TypeOf obj Is StockError) Then
			Return Me.Equals(DirectCast(obj, StockError))
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

Partial Public Class StockError
  Public Property ID As Integer
  Public Property DateUpdate As Date
  Public Property Symbol As String
  Public Property Description As String
  Public Property StockID As Integer

  Public Overridable Property Stock As Stock

End Class


#End Region