#Region "Imports"
  Imports YahooAccessData.ExtensionService
  Imports System
  Imports System.IO
  Imports System.Xml.Serialization
#End Region

Public Class BondRate1
  Implements IEquatable(Of BondRate1)
  Implements IRegisterKey(Of String)
  Implements IComparable(Of BondRate1)
  Implements IMemoryStream
  Implements IFormatData
  Implements IDisposable
  Implements IDateUpdate

#Region "Main Definition"
#Region "Enum"
  Public Enum enuBondType
    USTreasury
    Municipal
    Corporate
  End Enum
  Public Enum enuBondSecurity
    AAA 'Best quality, with smallest degree of investment risk.
    AA  'High quality by all standards; together with the AAA group they comprise what are generally known as high-grade bonds.
    A   'Possess many favorable investment attributes. Considered as upper-medium-grade obligations.
    BAA 'Medium-grade obligations (neither highly protected nor poorly secured). Bonds rated Baa and above are considered investment grade.
    BA  'Have speculative elements; futures are not as well-assured. Bonds rated Ba and below are generally considered speculative.
    B   'Generally lack characteristics of a desirable investment.
    CAA 'Bonds of poor standing.
    C   'Lowest rated class of bonds, with extremely poor prospects of ever attaining any real investment standing.
  End Enum
#End Region   'Enum

  Private _DateStart As Date
  Private _DateStop As Date
  Private MyException As Exception
  Private _Records As ICollection(Of BondRateRecord) = New LinkedHashSet(Of BondRateRecord, Date)
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
  'Private Shared MyCompareByName As CompareByName(Of BondRate1)
#End Region
#Region "New"
  Public Sub New()
    With Me
      .DateStart = Now
      .DateStop = Me.DateStart
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
      .Report.BondRates1.Add(Me)
    End With
  End Sub
#End Region
#Region "Properties Definition"
  Public Property ID As Integer
  Public Property Symbol As String
  Public Property Name As String
  Public Property Exchange As String
  Public Property MaturityDays As Integer
  Public Property Security As enuBondSecurity
  Public Property Type As enuBondType
  Public Property DateStart As Date
    Get
      Return _DateStart
    End Get
    Set(value As Date)
      _DateStart = value
      If _DateStart > _DateStop Then
        Me.DateStop = Me.DateStart
      End If
    End Set
  End Property
  Public Property DateStop As Date
    Get
      Return _DateStop
    End Get
    Set(value As Date)
      _DateStop = value
      If _DateStop < _DateStart Then
        Me.DateStart = _DateStop
      End If
    End Set
  End Property
  Public Property IsError As Boolean
  Public Property ErrorDescription As String
  Public Property Report As Report

  Public Property Records As ICollection(Of BondRateRecord)
    Get
      If TypeOf _Records Is IDataVirtual Then
        DirectCast(_Records, IDataVirtual).Load()
      End If
      Select Case Me.Report.TimeFormat
        Case YahooAccessData.Report.enuTimeFormat.Sample
          Return _Records
        Case YahooAccessData.Report.enuTimeFormat.Daily
          Return _Records
        Case YahooAccessData.Report.enuTimeFormat.Weekly
          Return _Records
        Case Else
          Return _Records
      End Select
    End Get
    Set(value As ICollection(Of BondRateRecord))
      _Records = value
    End Set
  End Property
#End Region
#Region "General"
  Friend Function CopyDeep(ByRef Parent As Report, Optional ByVal IsIgnoreID As Boolean = False) As BondRate1
    Dim ThisBondRate1 = New BondRate1

    With ThisBondRate1
      If IsIgnoreID = False Then .ID = Me.ID
      .Name = Me.Name
      .Symbol = Me.Symbol
      .Exchange = Me.Exchange
      .MaturityDays = Me.MaturityDays
      .Security = Me.Security
      .Type = Me.Type
      .DateStart = Me.DateStart
      .DateStop = Me.DateStop
      .IsError = Me.IsError
      .ErrorDescription = Me.ErrorDescription
      .Exception = Me.Exception
      .Report = Parent
      'to do
      .Report.BondRates1.Add(ThisBondRate1)
    End With
    Return ThisBondRate1
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
  Public Function CompareTo(other As BondRate1) As Integer Implements System.IComparable(Of BondRate1).CompareTo
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
      Return Me.Symbol
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
    Dim ThisZero As Integer

    With ThisBinaryWriter
      .Write(VERSION_MEMORY_STREAM)
      .Write(Me.ID)
      ThisListException = Me.Exception.ToList
      ThisListException.Reverse()
      .Write(ThisListException.Count)
      For Each ThisException In ThisListException
        .Write(ThisException.Message)
      Next
      .Write(Me.Symbol)
      .Write(Me.Name)
      .Write(Me.Exchange)
      .Write(Me.DateStart.ToBinary)
      .Write(Me.DateStop.ToBinary)
      .Write(CType(Me.Type, Short))
      .Write(CType(Me.Security, Short))
      .Write(Me.MaturityDays)
      Select Case FileType
        Case IMemoryStream.enuFileType.Standard
          .Write(Me.Records.Count)
          For Each ThisBondRateRecord In Me.Records
            ThisBondRateRecord.SerializeSaveTo(Stream, FileType)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          If TypeOf _Records Is IDataVirtual Then
            'make sure the data has been loaded before we save
            With DirectCast(_Records, IDataVirtual)
              Dim IsLoaded As Boolean = .IsLoaded
              If IsLoaded = False Then .Load()
            End With
          End If
          If _Records.Count > 0 Then
            Dim ThisRecordIndex As RecordIndex(Of BondRateRecord, Date)
            ThisRecordIndex = New RecordIndex(Of BondRateRecord, Date)(DirectCast(Stream, FileStream).Name(), FileMode.Open, Me.Report.AsDateRange, "BondRate1\Record", "_" & Me.Symbol, ".brr")
            ThisRecordIndex.Save(_Records)
            If (Me.Report.FileType = IMemoryStream.enuFileType.RecordIndexed) Then
              With DirectCast(_Records, IDataVirtual)
                If .Enabled Then
                  .Release()
                End If
              End With
            End If
            .Write(ThisRecordIndex.FileCount)
            .Write(ThisRecordIndex.MaxID)
            .Write(ThisRecordIndex.DateStart.ToBinary)
            .Write(ThisRecordIndex.DateStop.ToBinary)
            With DirectCast(_Records, IDateUpdate)
              .DateStart = ThisRecordIndex.DateStart
              .DateStop = ThisRecordIndex.DateStop
            End With
            With DirectCast(_Records, IRecordInfo)
              .CountTotal = ThisRecordIndex.FileCount
              .MaximumID = ThisRecordIndex.MaxID
            End With
            ThisRecordIndex.Dispose()
            ThisRecordIndex = Nothing
          Else
            .Write(ThisZero)
            .Write(ThisZero)
            .Write(Me.DateStart.ToBinary)
            .Write(Me.DateStart.ToBinary)
            With DirectCast(_Records, IDateUpdate)
              .DateStart = Me.DateStart
              .DateStop = Me.DateStart
            End With
            With DirectCast(_Records, IRecordInfo)
              .CountTotal = ThisZero
              .MaximumID = ThisZero
            End With
          End If
      End Select
    End With
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
    Dim ThisVersion As Single
    Dim ThisMaxID As Integer
    Dim ThisDateStart As Date
    Dim ThisDateStop As Date
    Dim Count As Integer

    With ThisBinaryReader
      ThisVersion = .ReadSingle
      Me.ID = .ReadInt32
      Me.Exception = Nothing
      Dim I As Integer
      For I = 1 To .ReadInt32
        MyException = New Exception(.ReadString, MyException)
      Next
      Me.Symbol = .ReadString
      Me.Name = .ReadString
      Me.Exchange = .ReadString
      Me.DateStart = DateTime.FromBinary(.ReadInt64)
      Me.DateStop = DateTime.FromBinary(.ReadInt64)
      Me.Type = CType(.ReadInt16, enuBondType)
      Me.Security = CType(.ReadInt16, enuBondSecurity)
      Me.MaturityDays = .ReadInt32
      Count = .ReadInt32
      Select Case Me.Report.FileType
        Case IMemoryStream.enuFileType.Standard
          If Count > 4 Then
            CType(_Records, LinkedHashSet(Of BondRateRecord, Date)).Capacity = Count
          End If
          For I = 1 To Count
            Dim ThisBondRateRecord As New BondRateRecord(Me, Stream)
          Next
        Case IMemoryStream.enuFileType.RecordIndexed
          ThisMaxID = .ReadInt32
          ThisDateStart = DateTime.FromBinary(.ReadInt64)
          ThisDateStop = DateTime.FromBinary(.ReadInt64)
          With CType(_Records, IRecordInfo)
            .CountTotal = Count
            .MaximumID = ThisMaxID
          End With
          With DirectCast(_Records, IDateUpdate)
            .DateStart = ThisDateStart
            .DateStop = ThisDateStop
          End With
      End Select
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
#Region "IEquatable"
  Public Function EqualsDeep(other As BondRate1, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    If other Is Nothing Then Return False
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
      End If
      If .Symbol <> Me.Symbol Then Return False
      If .Name <> Me.Name Then Return False
      If .Exchange <> Me.Exchange Then Return False
      If .MaturityDays <> Me.MaturityDays Then Return False
      If .Security <> Me.Security Then Return False
      If .Type <> Me.Type Then Return False
      If .DateStart <> Me.DateStart Then Return False
      If .DateStop <> Me.DateStop Then Return False
      If .IsError <> Me.IsError Then Return False
      If .ErrorDescription <> Me.ErrorDescription Then Return False
      If Me.Records.EqualsDeep(.Records, IsIgnoreID) = False Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As BondRate1) As Boolean Implements IEquatable(Of BondRate1).Equals
    If other Is Nothing Then Return False
    If Me.KeyValue = other.KeyValue Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If obj Is Nothing Then Return False
    If (TypeOf obj Is BondRate1) Then
      Return Me.Equals(DirectCast(obj, BondRate1))
    Else
      Return False
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.KeyValue.GetHashCode
  End Function
#End Region
#Region "IFormatData"
  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Return Extensions.ToStingOfData(Of BondRate1)(Me)
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
      .Add(New HeaderInfo With {.Name = "MaturityDays", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Type", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Security", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateStart", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateStop", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "IsError", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "ErrorDescription", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return DirectCast(_Records, IDateUpdate).DateStart
    End Get
    Set(value As Date)
    End Set
  End Property

  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return DirectCast(_Records, IDateUpdate).DateStop
    End Get
    Set(value As Date)
    End Set
  End Property

  Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return IDateUpdate_DateStop
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return IDateUpdate_DateStop
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return IDateUpdate_DateStop.Date
    End Get
  End Property
#End Region
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls
  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' dispose managed state (managed objects).
      End If

      '~free unmanaged resources (unmanaged objects) and override Finalize() below.
      '~set large fields to null.
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

