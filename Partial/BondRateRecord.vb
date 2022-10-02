
#Region "Imports"
  Imports YahooAccessData.ExtensionService
  Imports System
  Imports System.IO
  Imports System.Xml.Serialization
#End Region

Public Class BondRateRecord
  Implements IEquatable(Of BondRateRecord)
  Implements IRegisterKey(Of Date)
  Implements IComparable(Of BondRateRecord)
  Implements IMemoryStream
  Implements IFormatData
  Implements IDateUpdate
  Implements IDisposable

#Region "Main"
  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  'Private Shared MyCompareByName As CompareByName(Of BondRateRecord)
#End Region
#Region "New"
  Public Sub New(ByVal DateUpdate As Date)
    With Me
      .DateUpdate = DateUpdate
    End With
    If MyListHeaderInfo Is Nothing Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.xml"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    End If
    'If MyCompareByName Is Nothing Then
    '  MyCompareByName = New CompareByName(Of BondRateRecord)
    'End If
  End Sub

  Public Sub New()
    Me.New(Now)
  End Sub

  Public Sub New(ByRef Parent As BondRate1, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .BondRate1 = Parent
      If .BondRate1 IsNot Nothing Then
        .BondRate1.Records.Add(Me)
      End If
    End With
  End Sub
#End Region
#Region "General Function"
  Friend Function CopyDeep(ByRef Parent As BondRate1, Optional ByVal IsIgnoreID As Boolean = False) As YahooAccessData.BondRateRecord
    Dim ThisRecord = New BondRateRecord
    With ThisRecord
      If IsIgnoreID = False Then .ID = Me.ID
      .DateUpdate = Me.DateUpdate
      .Interest = Me.Interest
      .BondRate1 = Parent
      .BondRate1.Records.Add(ThisRecord)
    End With
    Return ThisRecord
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
    Return String.Format("{0},ID:{1},Key:{2},Record:{3} of {4}", TypeName(Me), Me.KeyID, Me.KeyValue.ToString)
  End Function
#End Region 'General function
#Region "Properties Definition"
  Public Property ID As Integer
  Public Property DateUpdate As Date
  Public Property Interest As Single
  Public Property BondRate1 As BondRate1
#End Region
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
  Public Function CompareTo(other As BondRateRecord) As Integer Implements System.IComparable(Of BondRateRecord).CompareTo
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
      .Write(Me.DateUpdate.ToBinary)
      .Write(Me.Interest)
    End With
  End Sub

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
      Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
      Me.Interest = .ReadSingle
    End With
    ThisBinaryReader.Dispose()
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
    Return Extensions.ToStingOfData(Of BondRateRecord)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Interest", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As BondRateRecord, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
      End If
      If .DateUpdate <> Me.DateUpdate Then Return False
      If .Interest <> Me.Interest Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As BondRateRecord) As Boolean Implements IEquatable(Of BondRateRecord).Equals
    If other Is Nothing Then Return False
    If Me.KeyValue = other.KeyValue Then
      If Me.BondRate1 IsNot Nothing Then
        Return Me.BondRate1.Equals(other.BondRate1)
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is BondRateRecord) Then
      Return Me.Equals(DirectCast(obj, BondRateRecord))
    Else
      Return (False)
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.KeyValue.GetHashCode()
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

  Private ReadOnly Property DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return Me.DateUpdate
    End Get
  End Property

  Private ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateUpdate.Date
    End Get
  End Property
#End Region 'IDateUpdate

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

#Region "RecordBondInterest Structure Definition"
  Public Structure BondInterests
    Public Sub New(
      ByRef colData As IEnumerable(Of YahooAccessData.BondRateRecord),
      ByVal DateStartValue As Date,
      ByVal DateStopValue As Date)

      Dim I As Integer
      Dim colDataDaily As IEnumerable(Of YahooAccessData.BondRateRecord)

      'set the default value
      'adjust the date for the constraint of always starting on Monday and ignore the weekend 
      Me.DateStart = ReportDate.DateToMondayPrevious(DateStartValue.Date)
      'DateStop should not fall on a weekend 
      'if it does move it back to the previous friday
      Me.DateStop = ReportDate.DateToWeekEndRemovePrevious(DateStopValue.Date)
      If Me.DateStop < Me.DateStart Then
        Me.DateStop = Me.DateStart
      End If
      Me.NumberPoint = ReportDate.MarketTradingDeltaDays(Me.DateStart, Me.DateStop) + 1
      ReDim Me.Interests(0 To Me.NumberPoint - 1)
      'get only the result between start and stop date included
      colDataDaily = colData.ToDaily(Me.DateStart, Me.DateStop)
      'set the default value for that condition
      Me.StartPoint = -1
      Me.StopPoint = -1
      If colDataDaily.Count = 0 Then
        Me.IsNull = True
        For I = 0 To Me.NumberPoint - 1
          Me.Interests(I).IsNull = True
        Next
        Me.NumberNullPoint = Me.NumberPoint
        Me.InterestMax = 0
        Me.InterestMin = 0
        Return
      Else
        'initialize the range variable
        Me.InterestMax = 0
        Me.InterestMin = Single.MaxValue
      End If
      'adjust the data
      Dim ThisRecord = colDataDaily.First
      Dim ThisDateCurrent As Date = Me.DateStart
      If ThisRecord.BondRate1 IsNot Nothing Then
        Me.Symbol = ThisRecord.BondRate1.Symbol
      Else
        Me.Symbol = ""
      End If
      I = 0
      'note that colDataDaily contain only the record between DateStart and DateStop
      For Each ThisRecord In colDataDaily
        'make sure the record reach the current pointer data before processing
        'start process the data
        'date synchronization loop in case the record date is later than then the current pointer data
        'fill the data with null value
        Do Until ThisDateCurrent >= ThisRecord.DateUpdate.Date
          Call BondInterestUpdateToNull(Me.Interests(I), ThisRecord, ThisDateCurrent)
          'add a day and make sure we jump over the week end
          ThisDateCurrent = ReportDate.DateToWeekEndRemoveNext(ThisDateCurrent.AddDays(1))
          I = I + 1
        Loop
        If Me.StartPoint = -1 Then
          Me.StartPoint = I
        End If
        'pointer and record date match
        'update the interest data
        Call BondInterestUpdate(Me.Interests(I), ThisRecord, ThisDateCurrent)
        ThisDateCurrent = ReportDate.DateToWeekEndRemoveNext(ThisDateCurrent.AddDays(1))
        I = I + 1
      Next
      Me.StopPoint = I - 1
      'make sure the data is fill with null if necessary up to DateStop
      Do Until ThisDateCurrent > Me.DateStop
        Call BondInterestUpdateToNull(Me.Interests(I), Nothing, ThisDateCurrent)
        Me.NumberNullPointToEnd = Me.NumberNullPointToEnd + 1
        'add a day and make sure we jump over the week end
        ThisDateCurrent = ReportDate.DateToWeekEndRemoveNext(ThisDateCurrent.AddDays(1))
        I = I + 1
      Loop
      If Me.NumberNullPoint = Me.NumberPoint Then
        Me.IsNull = True
      Else
        Me.IsNull = False
      End If
    End Sub

    Private Sub BondInterestUpdateToNull(
      ByRef Interest As Interest,
      ByRef Record As YahooAccessData.BondRateRecord,
      ByRef DateValue As Date)

      If Record IsNot Nothing Then
        With Interest
          .DateUpdate = DateValue
          .Value = Record.Interest
          .IsNull = True
          .Record = Record
        End With
      Else
        With Interest
          .DateUpdate = DateValue
          .Value = 0
          .IsNull = True
          .Record = Nothing
        End With
      End If
      Me.NumberNullPoint = Me.NumberNullPoint + 1
    End Sub

    Private Sub BondInterestUpdate(
      ByRef Interest As Interest,
      ByRef Record As YahooAccessData.BondRateRecord,
      ByRef DateValue As Date)

      With Interest
        .DateUpdate = DateValue
        .Value = Record.Interest
        .IsNull = False
        .Record = Record
        'check the range
        If .Value > Me.InterestMax Then
          Me.InterestMax = .Value
        End If
        If .Value > 0 Then
          If .Value < Me.InterestMin Then
            Me.InterestMin = .Value
          End If
        End If
      End With
    End Sub

    Public Function ToIndex(ByVal DateValue As Date) As Integer
      Return ReportDate.MarketTradingDeltaDays(Me.DateStart, DateValue)
    End Function

    Public InterestLast As Interest
    Public Interests() As Interest
    Public InterestMin As Single
    Public InterestMax As Single
    Public DateStart As Date
    Public DateStop As Date
    Public NumberPoint As Integer
    Public NumberNullPoint As Integer
    Public NumberNullPointToEnd As Integer
    Public StartPoint As Integer
    Public StopPoint As Integer
    Public IsError As Boolean
    Public ErrorDescription As String
    Public Symbol As String
    Public IsNull As Boolean
  End Structure

  Public Structure Interest
    Public DateUpdate As Date
    Public Value As Single
    Public Record As BondRateRecord
    Public IsNull As Boolean
  End Structure
#End Region