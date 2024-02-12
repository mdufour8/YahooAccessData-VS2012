#Region "Imports"
Imports System
Imports System.Collections.Generic
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

<Serializable()>
Partial Public Class SplitFactorFuture
	Implements IEquatable(Of SplitFactorFuture)
	Implements IRegisterKey(Of String)
	Implements IComparable(Of SplitFactorFuture)
	Implements IMemoryStream
	Implements IFormatData
  Implements IDateUpdate
  Implements ICompareData(Of SplitFactorFuture)
  Implements IDisposable

  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  Private Shared MyCompareByName As CompareByName(Of SplitFactorFuture)

  Public Sub New()
    With Me
      .DateUpdate = Now
      .DateEx = CDate(.DateUpdate)
      .SharesNew = 1
      .SharesLast = 1
      .Ratio = CSng(.SharesNew / .SharesLast)
    End With
    If MyListHeaderInfo Is Nothing Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    End If
    'If MyCompareByName Is Nothing Then
    '  MyCompareByName = New CompareByName(Of SplitFactorFuture)
    'End If
  End Sub

  Public Sub New(ByRef Parent As Report, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .Report = Parent
      .ReportID = .Report.ID
      .Report.SplitFactorFutures.Add(Me)
    End With
  End Sub

  Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As SplitFactorFuture
    Dim ThisSplitFactorFuture = New SplitFactorFuture

    With ThisSplitFactorFuture
      If IsIgnoreID = False Then .ID = Me.ID
      .DateUpdate = Me.DateUpdate
      .DateAnnounce = Me.DateAnnounce
      .DateEx = Me.DateEx
      .DatePayable = Me.DatePayable
      .Exception = Me.Exception
      .Exchange = Me.Exchange
      .Name = Me.Name
      .Ratio = Me.Ratio
      .SharesLast = Me.SharesLast
      .SharesNew = Me.SharesNew
      .Symbol = Me.Symbol
      .Report = Parent
      .ReportID = .Report.ID
      .Report.SplitFactorFutures.Add(ThisSplitFactorFuture)
    End With
    Return ThisSplitFactorFuture
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
  Public Function CompareTo(other As SplitFactorFuture) As Integer Implements System.IComparable(Of SplitFactorFuture).CompareTo
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
      If Me.DateAnnounce Is Nothing Then
        .Write(False)
      Else
        .Write(True)
        .Write(CDate(Me.DateAnnounce).ToBinary)
      End If
      If Me.DatePayable Is Nothing Then
        .Write(False)
      Else
        .Write(True)
        .Write(CDate(Me.DatePayable).ToBinary)
      End If
      .Write(Me.Exchange)
      .Write(Me.DateEx.ToBinary)
      .Write(Me.SharesLast)
      .Write(Me.SharesNew)
      .Write(Me.Ratio)
      .Write(CDate(Me.DateUpdate).ToBinary)
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
      If .ReadBoolean Then
        Me.DateAnnounce = DateTime.FromBinary(.ReadInt64)
      Else
        Me.DateAnnounce = New Nullable(Of Date)
      End If
      If .ReadBoolean Then
        Me.DatePayable = DateTime.FromBinary(.ReadInt64)
      Else
        Me.DatePayable = New Nullable(Of Date)
      End If
      Me.Exchange = .ReadString
      Me.DateEx = DateTime.FromBinary(.ReadInt64)
      Me.SharesLast = .ReadInt32
      Me.SharesNew = .ReadInt32
      Me.Ratio = .ReadSingle
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
    Me.SerializeSaveTo(ThisStream)
    Return ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
  End Function
#End Region
#Region "IFormatData"
  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Return Extensions.ToStingOfData(Of SplitFactorFuture)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Name", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Exchange", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Symbol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateAnnounce", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DatePayable", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateEx", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesLast", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "SharesNew", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Ratio", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As SplitFactorFuture, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    If other Is Nothing Then Return False
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
        If .ReportID <> Me.ReportID Then Return False
      End If
      If .Symbol <> Me.Symbol Then Return False
      If .Name <> Me.Name Then Return False
      If .DateEx <> Me.DateEx Then Return False
      If .SharesLast <> Me.SharesLast Then Return False
      If .SharesNew <> Me.SharesNew Then Return False
      If .Ratio <> Me.Ratio Then Return False
      If .DatePayable <> Me.DatePayable Then Return False
      If .DateAnnounce <> Me.DateAnnounce Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As SplitFactorFuture) As Boolean Implements IEquatable(Of SplitFactorFuture).Equals
    If other Is Nothing Then Return False
    If Me.Symbol = other.Symbol Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If obj Is Nothing Then Return False
    If (TypeOf obj Is SplitFactorFuture) Then
      Return Me.Equals(DirectCast(obj, SplitFactorFuture))
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
      If Me.DateUpdate.HasValue Then
        Return CDate(Me.DateUpdate)
      Else
        Return Me.Report.DateStop
      End If
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
#End Region 'IDateUpdate
#Region "ICompareData"
  Private Function ICompareData_Comparer(ByVal TableName As String) As System.Comparison(Of SplitFactorFuture) Implements ICompareData(Of SplitFactorFuture).Comparer
    MyCompareByName.PropertyName = TableName
    Return AddressOf MyCompareByName.Compare
  End Function
#End Region
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


#Region "Template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Partial Public Class SplitFactorFuture
  Public Property ID As Integer
  Public Property ReportID As Integer
  Public Property Symbol As String
  Public Property Name As String
  Public Property DateAnnounce As Nullable(Of Date)
  Public Property DatePayable As Nullable(Of Date)
  Public Property Exchange As String
  Public Property DateEx As Date
  Public Property SharesLast As Integer
  Public Property SharesNew As Integer
  Public Property Ratio As Single
  Public Property DateUpdate As Nullable(Of Date)

  Public Overridable Property Report As Report

End Class
#End Region