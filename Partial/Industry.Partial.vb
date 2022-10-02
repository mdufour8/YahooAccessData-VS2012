#Region "Imports"
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System
Imports System.IO
#End Region

<Serializable()>
Partial Public Class Industry
	Implements IEquatable(Of Industry)
	Implements IRegisterKey(Of String)
	Implements IComparable(Of Industry)
	Implements IMemoryStream
	Implements IFormatData
  Implements IDateUpdate
  Implements IDisposable

  'Implements ICompareData(Of Industry)

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  'Private Shared MyCompareByName As CompareByName(Of Industry)

  Public Sub New(ByRef Parent As YahooAccessData.Report, ByVal Name As String)
    Me.New(Name)
    With Me
      .Report = Parent
      If .Report IsNot Nothing Then
        .ReportID = .Report.ID
        .Report.Industries.Add(Me)
      End If
    End With
  End Sub

  Public Sub New(ByRef Parent As YahooAccessData.Report, ByRef Sector As Sector, ByVal Name As String)
    Me.New(Name)
    With Me
      .Report = Parent
      If .Report IsNot Nothing Then
        .ReportID = .Report.ID
        .Report.Industries.Add(Me)
      End If
      .Sector = Sector
      If .Sector IsNot Nothing Then
        .SectorID = .Sector.ID
        .Sector.Industries.Add(Me)
      End If
    End With
  End Sub

  Public Sub New(ByRef Parent As Report, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .Report = Parent
      .Sector = .Report.Sectors.ToSearch.Find(.SectorID)
      If .Sector Is Nothing Then
        .Sector = .Sector
        .Sector.Industries.Add(Me)
      End If
      .ReportID = .Report.ID
      .Report.Industries.Add(Me)
    End With
  End Sub

  Public Sub New(ByVal Name As String)
    With Me
      .Name = Name
      .Stocks = New LinkedHashSet(Of Stock, String)
    End With
    If MyListHeaderInfo Is Nothing Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    End If
    'If MyCompareByName Is Nothing Then
    '  MyCompareByName = New CompareByName(Of Industry)
    'End If
  End Sub

  Public Sub New()
    Me.New("")
  End Sub

  Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As Industry
    Dim ThisIndustry As Industry

    ThisIndustry = New YahooAccessData.Industry
    With ThisIndustry
      If IsIgnoreID = False Then .ID = Me.ID
      .Name = Me.Name
      .DataSourceID = Me.DataSourceID
      .Report = Parent
      .ReportID = .Report.ID
      .Report.Industries.Add(ThisIndustry)
      .Sector = .Report.Sectors.ToSearch.Find(Me.Sector.KeyValue)
      If .Sector Is Nothing Then
        .Exception = New Exception("Invalid sector...", .Exception)
        Return Nothing
      Else
        .SectorID = .Sector.ID
        .Sector.Industries.Add(ThisIndustry)
      End If
    End With
    Return ThisIndustry
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
  Public Function CompareTo(other As Industry) As Integer Implements System.IComparable(Of Industry).CompareTo
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
      .Write(Me.SectorID)
      If Me.SectorID = 0 Then
        Me.SectorID = 0
      End If
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
      Me.SectorID = .ReadInt32
      Me.DataSourceID = .ReadInt32
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
    Return Extensions.ToStingOfData(Of Industry)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0:G}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DataSourceID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As Industry, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
        If .ReportID <> Me.ReportID Then Return False
        If .SectorID <> Me.SectorID Then Return False
      End If
      If .Name <> Me.Name Then Return False
      If .DataSourceID <> Me.DataSourceID Then Return False
      'we do not need to check for deep equality here
      'since the stocks are only a reference to the main list
      If Me.Stocks.EqualsShallow(.Stocks) = False Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As Industry) As Boolean Implements IEquatable(Of Industry).Equals
    If other Is Nothing Then Return False
    If Me.Name = other.Name Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is Industry) Then
      Return Me.Equals(DirectCast(obj, Industry))
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
  '  Private Function ICompareData_Comparer(ByVal TableName As String) As System.Comparison(Of Industry) Implements ICompareData(Of Industry).Comparer
  '    MyCompareByName.PropertyName = TableName
  '    Return CType(AddressOf MyCompareByName.Compare, System.Comparison(Of Industry))
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

  ' override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
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

#Region "EqualityComparer Class"
	<Serializable()>
	Friend Class EqualityComparerOfIndustry
		Implements IEqualityComparer(Of Industry)

		Public Overloads Function Equals(x As Industry, y As Industry) As Boolean Implements IEqualityComparer(Of Industry).Equals
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

		Public Overloads Function GetHashCode(obj As Industry) As Integer Implements IEqualityComparer(Of Industry).GetHashCode
			If obj IsNot Nothing Then
				Return obj.Name.GetHashCode
			Else
				Return obj.GetHashCode
			End If
		End Function
	End Class
#End Region

#Region "Industry template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class Industry
  Public Property ID As Integer
  Public Property Name As String
  Public Property ReportID As Integer
  Public Property SectorID As Integer
  Public Property DataSourceID As Integer

  Public Overridable Property Stocks As ICollection(Of Stock) = New HashSet(Of Stock)
  Public Overridable Property Report As Report
  Public Overridable Property Sector As Sector

End Class
#End Region

