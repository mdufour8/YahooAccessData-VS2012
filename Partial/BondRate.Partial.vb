#Region "Imports"
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System
Imports System.IO
Imports System.Xml.Serialization
#End Region


<Serializable()>
Partial Public Class BondRate
	Implements IEquatable(Of BondRate)
	Implements IRegisterKey(Of String)
	Implements IComparable(Of BondRate)
	Implements IMemoryStream
	Implements IFormatData
  Implements IDateUpdate

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)

	Public Sub New()
		With Me
			.DateUpdate = Now
		End With
		If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
			Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
			MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
		Else
			MyListHeaderInfo = ListOfHeader()
		End If
	End Sub

	Public Sub New(ByRef Parent As Report, ByRef Stream As Stream)
		Me.New()
		With (Me)
			.SerializeLoadFrom(Stream)
			.Report = Parent
			.ReportID = .Report.ID
			.Report.BondRates.Add(Me)
		End With
	End Sub

	Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As BondRate
		Dim ThisBondRate = New BondRate

		With ThisBondRate
			If IsIgnoreID = False Then .ID = Me.ID
			.Name = Me.Name
			.Symbol = Me.Symbol
			.Interest = Me.Interest
			.MaturityDays = Me.MaturityDays
			.Security = Me.Security
			.Type = Me.Type
			.DateUpdate = Me.DateUpdate
			.Exception = Me.Exception
			.Report = Parent
			.ReportID = .Report.ID
			.Report.BondRates.Add(ThisBondRate)
		End With
		Return ThisBondRate
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
	Public Function CompareTo(other As BondRate) As Integer Implements System.IComparable(Of BondRate).CompareTo
    Return Me.KeyValue.CompareTo(other.KeyValue)
  End Function
#End Region
#Region "Register Key"
	Dim ThisKeyValue As String = ""
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
			Return Me.Symbol & Me.DateUpdate.ToString
		End Get
		Set(value As String)

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
			.Write(Me.ReportID)
			.Write(Me.Symbol)
			.Write(Me.Name)
			.Write(Me.Type)
			.Write(Me.Security)
			.Write(Me.Interest)
			.Write(Me.MaturityDays)
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
      Me.ReportID = .ReadInt32
      Me.Symbol = .ReadString
      Me.Name = .ReadString
      Me.Type = .ReadInt16
      Me.Security = .ReadInt16
      Me.Interest = .ReadSingle
      Me.MaturityDays = .ReadInt32
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
#Region "IEquatable"
	Public Function EqualsDeep(other As BondRate, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
		If other Is Nothing Then Return False
		With other
			If IsIgnoreID = False Then
				If .ID <> Me.ID Then Return False
				If .ReportID <> Me.ReportID Then Return False
			End If
			If .Symbol <> Me.Symbol Then Return False
			If .Name <> Me.Name Then Return False
			If .Interest <> Me.Interest Then Return False
			If .MaturityDays <> Me.MaturityDays Then Return False
			If .Security <> Me.Security Then Return False
			If .Type <> Me.Type Then Return False
			If .DateUpdate <> Me.DateUpdate Then Return False
		End With
		Return True
	End Function

	Public Overloads Function Equals(other As BondRate) As Boolean Implements IEquatable(Of BondRate).Equals
		If other Is Nothing Then Return False
		If Me.Symbol = other.Symbol Then
			Return True
		Else
			Return False
		End If
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If obj Is Nothing Then Return False
		If (TypeOf obj Is BondRate) Then
			Return Me.Equals(DirectCast(obj, BondRate))
		Else
			Return False
		End If
	End Function

	Public Overrides Function GetHashCode() As Integer
		Return Me.Symbol.GetHashCode
	End Function
#End Region
#Region "IFormatData"
	Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
		Return Extensions.ToStingOfData(Of BondRate)(Me)
	End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Symbol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Interest", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "MaturityDays", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Security", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Type", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
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

  Public ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return Me.DateUpdate
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateUpdate.Date
    End Get
  End Property
#End Region
End Class

#Region "IBondRate"
Public Interface IBondRate
  Property ID As Integer
  Property ReportID As Integer
  Property Symbol As String
  Property Name As String
  Property Type As Short 'Short is an int16
  Property Security As Short
  Property Interest As Single
  Property MaturityDays As Integer
  Property DateUpdate As Date
  Property Report As Report
End Interface
#End Region


#Region "BondRate template"
Partial Public Class BondRate
  Public Property ID As Integer
  Public Property ReportID As Integer
  Public Property Symbol As String
  Public Property Name As String
  Public Property Type As Short 'Short is an int16
  Public Property Security As Short
  Public Property Interest As Single
  Public Property MaturityDays As Integer
  Public Property DateUpdate As Date

  Public Overridable Property Report As Report

End Class
#End Region