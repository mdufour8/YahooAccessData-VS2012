#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region
<Serializable()>
Partial Public Class Sector
	Implements IEquatable(Of Sector)
	Implements IRegisterKey(Of String)
	Implements IComparable(Of Sector)
	Implements IMemoryStream
	Implements IFormatData
  Implements IDateUpdate
  Implements IDisposable

  'Implements ICompareData(Of Sector)

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  'Private Shared MyCompareByName As CompareByName(Of Sector)

  Public Sub New(ByRef Parent As YahooAccessData.Report, ByVal Name As String)
    Me.New(Name)
    With Me
      .Report = Parent
      If .Report IsNot Nothing Then
        .ReportID = .Report.ID
        .Report.Sectors.Add(Me)
      End If
    End With
  End Sub

  Public Sub New(ByRef Parent As Report, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .Report = Parent
      .ReportID = .Report.ID
      .Report.Sectors.Add(Me)
    End With
  End Sub

  Public Sub New(ByVal Name As String)
    With Me
      .Name = Name
      .Industries = New LinkedHashSet(Of Industry, String)
      .Stocks = New LinkedHashSet(Of Stock, String)
    End With
    If MyListHeaderInfo Is Nothing And LIST_OF_HEADER_FILE_ENABLED Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    Else
      MyListHeaderInfo = ListOfHeader()
    End If
  End Sub

  Public Sub New()
    Me.New("")
  End Sub

  Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As Sector
    Dim ThisSector As Sector
    ThisSector = New YahooAccessData.Sector
    'add the sector
    With ThisSector
      If IsIgnoreID = False Then .ID = Me.ID
      .Name = Me.Name
      .DataSourceID = Me.DataSourceID
      .Report = Parent
      .ReportID = .Report.ID
      .Report.Sectors.Add(ThisSector)
    End With
    Return ThisSector
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
  Public Function CompareTo(other As Sector) As Integer Implements System.IComparable(Of Sector).CompareTo
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
      Return Me.Name
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
      .Write(Me.Name)
      .Write(Me.ReportID)
      .Write(Me.DataSourceID)
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
      Me.Name = .ReadString
      Me.ReportID = .ReadInt32
      Me.DataSourceID = .ReadInt32
    End With
    ThisBinaryReader.Dispose()
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
    Me.SerializeLoadFrom(Stream, False)
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
    Return Extensions.ToStingOfData(Of Sector)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DataSourceID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As Sector, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
        If .ReportID <> Me.ReportID Then Return False
      End If
      If .Name <> Me.Name Then Return False
      If .DataSourceID <> Me.DataSourceID Then Return False
      'we do not need to check for deep equality here
      'since the stocks and industries are only a reference to the main list
      If Me.Stocks.EqualsShallow(.Stocks) = False Then Return False
      If Me.Industries.EqualsShallow(.Industries) = False Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As Sector) As Boolean Implements IEquatable(Of Sector).Equals
    If other Is Nothing Then Return False
    If Me.Name = other.Name Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is Sector) Then
      Return Me.Equals(DirectCast(obj, Sector))
    Else
      Return (False)
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.Name.GetHashCode()
  End Function
#End Region
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return IDateUpdate_DateUpdate
    End Get
    Set(value As Date)
    End Set
  End Property

  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return IDateUpdate_DateUpdate
    End Get
    Set(value As Date)
    End Set
  End Property

  Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return Me.Report.DateStop
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
  '#Region "ICompareData"
  '  Private Function ICompareData_Comparer(ByVal TableName As String) As System.Comparison(Of Sector) Implements ICompareData(Of Sector).Comparer
  '    MyCompareByName.PropertyName = TableName
  '    Return AddressOf MyCompareByName.Compare
  '  End Function
  '#End Region

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
End Class

#Region "EqualityComparerOfSector"
	<Serializable()>
	Friend Class EqualityComparerOfSector
		Implements IEqualityComparer(Of Sector)

		Public Overloads Function Equals(x As Sector, y As Sector) As Boolean Implements IEqualityComparer(Of Sector).Equals
			If (x Is Nothing) And (y Is Nothing) Then
				Return True
			ElseIf (x Is Nothing) Xor (y Is Nothing) Then
				Return False
			Else
				If x.Name = y.Name Then
					Return True
				Else
					Return False
				End If
			End If
		End Function

		Public Overloads Function GetHashCode(obj As Sector) As Integer Implements IEqualityComparer(Of Sector).GetHashCode
			Return obj.Name.GetHashCode
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
Partial Public Class Sector
  Public Property ID As Integer
  Public Property Name As String
  Public Property ReportID As Integer
  Public Property DataSourceID As Integer

  Public Overridable Property Industries As ICollection(Of Industry) = New HashSet(Of Industry)
  Public Overridable Property Report As Report
  Public Overridable Property Stocks As ICollection(Of Stock) = New HashSet(Of Stock)

End Class
#End Region