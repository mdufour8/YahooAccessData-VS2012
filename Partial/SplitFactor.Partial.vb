#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Partial Public Class SplitFactor
	Implements IEquatable(Of SplitFactor)
	Implements IRegisterKey(Of Date)
	Implements IComparable(Of SplitFactor)
	Implements IMemoryStream
	Implements IFormatData
	Implements IDateUpdate

	Private MyException As Exception
	Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New(ByVal DateUpdate As Date)
		With Me
			.DateDay = DateUpdate.Date
		End With
		If MyListHeaderInfo Is Nothing Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		End If
	End Sub

	Public Sub New()
    Me.New(Now)
	End Sub

	Public Sub New(ByRef Parent As Stock, ByRef Stream As Stream)
		Me.New()
		With Me
			.SerializeLoadFrom(Stream)
			.Stock = Parent
			.StockID = .Stock.ID
			.Stock.SplitFactors.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Stock, Optional ByVal IsIgnoreID As Boolean = False) As SplitFactor
		Dim ThisSplitFactor = New SplitFactor

		With ThisSplitFactor
			If IsIgnoreID = False Then .ID = Me.ID
			.DateDay = Me.DateDay
			.SharesLast = Me.SharesLast
			.SharesNew = Me.SharesNew
			.Ratio = Me.Ratio
			.Stock = Parent
			.StockID = .Stock.ID
			.Stock.SplitFactors.Add(ThisSplitFactor)
		End With
		Return ThisSplitFactor
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
	Public Function CompareTo(other As SplitFactor) As Integer Implements System.IComparable(Of SplitFactor).CompareTo
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
			'restrict the symbol change to a maximum of one per day
			Return Me.DateDay.Date
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
			.Write(Me.StockID)
			.Write(Me.DateDay.ToBinary)
			.Write(Me.SharesLast)
			.Write(Me.SharesNew)
			.Write(Me.Ratio)
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
			Me.StockID = .ReadInt32
			Me.DateDay = DateTime.FromBinary(.ReadInt64)
			Me.SharesLast = .ReadInt32
			Me.SharesNew = .ReadInt32
			Me.Ratio = .ReadSingle
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
		Return Extensions.ToStingOfData(Of SplitFactor)(Me)
	End Function

	Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
		Return MyListHeaderInfo
	End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateDay", .Title = .Name, .Format = "{0:g}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesLast", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesNew", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Ratio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
	Public Function EqualsDeep(other As SplitFactor, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .StockID <> Me.StockID Then Return False
			End If
			If .DateDay <> Me.DateDay Then Return False
			If .SharesLast <> Me.SharesLast Then Return False
			If .SharesNew <> Me.SharesNew Then Return False
			If .Ratio <> Me.Ratio Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As SplitFactor) As Boolean Implements IEquatable(Of SplitFactor).Equals
		If other Is Nothing Then Return False
		If Me.DateDay = other.DateDay Then
			If Me.Stock IsNot Nothing Then
				Return Me.Stock.Equals(other.Stock)
			Else
				Return False
			End If
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If (TypeOf obj Is SplitFactor) Then
			Return Me.Equals(DirectCast(obj, SplitFactor))
		Else
			Return False
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		If Me.Stock IsNot Nothing Then
			Return Me.Stock.GetHashCode()
		Else
			Return Me.DateDay.GetHashCode()
		End If
	End Function

	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
#End Region
#Region "IDateUpdate"
	Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
		Get
			Return Me.DateDay
		End Get
		Set(value As Date)
		End Set
	End Property
	Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
		Get
			Return Me.DateDay
		End Get
		Set(value As Date)
		End Set
	End Property
	Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
		Get
			Return Me.DateDay
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
#End Region	'IDateUpdate
End Class

#Region "EqualityComparerOfSplitFactor"
	<Serializable()>
	Friend Class EqualityComparerOfSplitFactor
		Implements IEqualityComparer(Of SplitFactor)

		Public Overloads Function Equals(x As SplitFactor, y As SplitFactor) As Boolean Implements IEqualityComparer(Of SplitFactor).Equals
			If (x Is Nothing) And (y Is Nothing) Then
				Return True
			ElseIf (x Is Nothing) Xor (y Is Nothing) Then
				Return False
			Else
				If x.DateDay = y.DateDay Then
					If (x.Stock Is Nothing) And (y.Stock Is Nothing) Then
						Return True
					ElseIf (x.Stock Is Nothing) Xor (y.Stock Is Nothing) Then
						Return False
					Else
						If x.Stock.Equals(y) Then
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

		Public Overloads Function GetHashCode(obj As SplitFactor) As Integer Implements IEqualityComparer(Of SplitFactor).GetHashCode
			If obj IsNot Nothing Then
				Return obj.DateDay.GetHashCode
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

Partial Public Class SplitFactor
  Public Property ID As Integer
  Public Property StockID As Integer
  Public Property DateDay As Date
  Public Property SharesLast As Integer
  Public Property SharesNew As Integer
  Public Property Ratio As Single

  Public Overridable Property Stock As Stock

End Class

#End Region