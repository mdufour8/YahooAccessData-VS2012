#Region "Imports"
  Imports System
  Imports System.IO
  Imports System.Threading
  Imports System.Threading.Tasks
#End Region

#Region "modRecordIndex"
Public Module modRecordIndex
  Public Async Function CreateTaskForRecordIndexSave(Of T As {New, IRegisterKey(Of U), IDateUpdate, IMemoryStream}, U)(
      ByVal StreamBaseName As String,
      ByVal FileMode As FileMode,
      ByVal DateRange As IDateRange,
      ByVal ChildPath As String,
      ByVal KeyName As String,
      ByVal FileExtension As String,
      ByVal colData As System.Collections.Generic.ICollection(Of T),
      Optional ByVal IsSaveAtEndOfDay As Boolean = False) As Task(Of RecordIndex(Of T, U))


    Dim ThisTask As Task(Of RecordIndex(Of T, U))
    Dim ThisRecordIndex As RecordIndex(Of T, U)

    'create and wait on a local task for saving
    ThisTask = New Task(Of RecordIndex(Of T, U))(
      Function()
        ThisRecordIndex = New RecordIndex(Of T, U)(StreamBaseName, FileMode, DateRange, ChildPath, KeyName, FileExtension)
        ThisRecordIndex.Save(colData, IsSaveAtEndOfDay)
        Return ThisRecordIndex
      End Function)

    'ThisTask.RunSynchronously()
    ThisTask.Start()
    Await ThisTask
    Return ThisTask.Result
  End Function
End Module
#End Region
#Region "IRecordIndex(Of T)"
Public Interface IRecordIndex(Of T)
  ReadOnly Property MaxID As Integer
  ReadOnly Property FileCount As Integer
  Function ToPosition(ByVal Index As Integer) As Long
  ReadOnly Property ToListPosition As List(Of Long)
  Sub Close()
  Sub Save(ByRef colData As ICollection(Of T))
  Sub Save(ByRef colData As ICollection(Of T), IsSaveAtEndOfDay As Boolean)
  Sub Save(ByRef Data As T)
  Sub Save(ByRef Data As T, IsSaveAtEndOfDay As Boolean)
  ReadOnly Property BaseStreamIndex() As Stream
  ReadOnly Property BaseStreamIndex(ByVal Position As Long) As Stream
  ReadOnly Property BaseStream() As Stream
  ReadOnly Property BaseStream(ByVal Position As Long) As Stream
End Interface
#End Region
#Region "RecordIndex(Of T)"
Public Class RecordIndex(Of T As {New, IRegisterKey(Of U), IDateUpdate, IMemoryStream}, U)
  Implements IRecordIndex(Of T)

  Implements IDisposable
  Implements IDateUpdate
  Implements IDateTrade

  Private Structure strRecordIndex
    Public DateUpdate As Date
    Public Position As Long
  End Structure

  Private Const SIZE_OF_RECORD_INDEX As Long = 16
  Private Const POSITION_FILE_START_OFFSET As Long = 40
  Private Const FILE_GROW_SIZE_MINIMUM_INDEX As Integer = 200 * SIZE_OF_RECORD_INDEX
  Private Const FILE_GROW_SIZE_MINIMUM_RECORD As Integer = 64000

  Private MyKeyName As String
  Private MyDateRange As IDateRange
  Private MyDateStart As Date
  Private MyDateStop As Date
  Private MyPositionEndIndex As Long
  Private MyPositionEndRecord As Long
  Private MyBinaryReaderOfRecordIndex As BinaryReader
  Private MyFileCount As Integer
  Private MyMaxKeyID As Integer
  Private MyListPosition As List(Of Long)
  Private MyFileStreamOfRecord As Stream
  Private MyFileStreamOfRecordIndex As FileStream
  Private MyFileMode As FileMode
  Private IsPositionLoaded As Boolean

#Region "New"
  Public Sub New( _
      ByVal StreamBaseName As String,
      ByVal FileMode As FileMode,
      ByRef DateRange As IDateRange,
      ByVal ChildPath As String,
      ByVal KeyName As String,
      ByVal FileExtension As String,
      ByVal IsReadOnly As Boolean)

    Dim ThisFilePathRecordName As String
    Dim Zero As Integer = 0
    Dim IsFileExist As Boolean

    MyFileMode = FileMode
    Dim ThisStreamFileName As String = StreamBaseName
    Dim ThisFileBaseName As String = System.IO.Path.GetFileNameWithoutExtension(ThisStreamFileName)
    Dim ThisFilePathBase As String = System.IO.Path.GetDirectoryName(ThisStreamFileName)
    If Len(ChildPath) > 0 Then
      ThisFilePathRecordName = System.IO.Path.Combine(ThisFilePathBase, ChildPath)
    Else
      ThisFilePathRecordName = ThisFilePathBase
    End If
    If Mid$(KeyName, 1, 1) = "_" Then
      'remove the _ separator
      MyKeyName = Mid$(KeyName, 2)
    Else
      MyKeyName = KeyName
    End If
    If Len(MyKeyName) >= 1 Then
      'create a subdirectoy using the first letter of keyname
      ThisFilePathRecordName = System.IO.Path.Combine(ThisFilePathRecordName, Mid(MyKeyName, 1, 1))
    End If
    Dim ThisRecordBaseName As String = System.IO.Path.Combine(ThisFilePathRecordName, String.Format("{0}_{1}", ThisFileBaseName, MyKeyName))
    Dim ThisStreamRecordName As String = ThisRecordBaseName & FileExtension
    With My.Computer.FileSystem
      IsFileExist = .FileExists(ThisStreamRecordName)
      If FileMode = IO.FileMode.CreateNew Then
        'delete the current file
        If IsFileExist Then
          Try
            .DeleteFile(ThisStreamRecordName)
            .DeleteFile(ThisStreamRecordName & ".idx")
          Catch ex As Exception
            Me.Exception = ex
            Exit Sub
          End Try
        End If
      End If
      If IsFileExist = False Then
        'create the directory if it does not exist
        If .DirectoryExists(ThisFilePathRecordName) = False Then
          Try
            .CreateDirectory(ThisFilePathRecordName)
          Catch ex As Exception
            Me.Exception = New Exception(String.Format("Unable to create directory {0},ThisPath"), ex)
            Exit Sub
          End Try
        End If
      End If
      'create or open the file if IsReadonly is false
      Try
        If IsReadOnly = True Then
          MyFileStreamOfRecord = New FileStream(ThisStreamRecordName, IO.FileMode.Open, FileAccess.Read, FileShare.Read)
          MyFileStreamOfRecordIndex = New FileStream(ThisStreamRecordName & ".idx", FileMode.Open, FileAccess.Read, FileShare.Read)
        Else
          MyFileStreamOfRecord = New FileStream(ThisStreamRecordName, IO.FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
          MyFileStreamOfRecordIndex = New FileStream(ThisStreamRecordName & ".idx", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None)
        End If
      Catch ex As Exception
        If MyFileStreamOfRecord IsNot Nothing Then
          MyFileStreamOfRecord.Dispose()
        End If
        If MyFileStreamOfRecordIndex IsNot Nothing Then
          MyFileStreamOfRecordIndex.Dispose()
        End If
        Me.Exception = Exception
        Exit Sub
      End Try
      If IsFileExist = False Then
        'Init the file if it did not exist
        Try
          Dim ThisBinaryWriterOfRecordIndex As New BinaryWriter(MyFileStreamOfRecordIndex)
          MyFileStreamOfRecord.Seek(0, SeekOrigin.Begin)
          MyPositionEndRecord = 0
          MyFileStreamOfRecord.SetLength(FILE_GROW_SIZE_MINIMUM_RECORD)
          With ThisBinaryWriterOfRecordIndex
            .Seek(0, SeekOrigin.Begin)
            MyPositionEndIndex = 0

            .Write(MyPositionEndIndex)
            .Write(MyPositionEndRecord)
            'write the number of data in the main file
            .Write(Zero)
            'write the maximum keyID always in the last record
            .Write(Zero)
            'write the date update range in the file
            MyDateStart = Now
            MyDateStop = MyDateStart
            .Write(MyDateStart.ToBinary)
            .Write(MyDateStop.ToBinary)
            MyPositionEndIndex = .BaseStream.Position
            .BaseStream.SetLength(MyPositionEndIndex + FILE_GROW_SIZE_MINIMUM_INDEX)
            .Seek(0, SeekOrigin.Begin)
            .Write(MyPositionEndIndex)
          End With
        Catch ex As Exception
          If MyFileStreamOfRecord IsNot Nothing Then
            MyFileStreamOfRecord.Dispose()
          End If
          If MyFileStreamOfRecordIndex IsNot Nothing Then
            MyFileStreamOfRecordIndex.Dispose()
          End If
          Me.Exception = ex
        End Try
      End If
    End With
    'update the parameters
    MyListPosition = New List(Of Long)
    MyDateRange = DateRange
    MyBinaryReaderOfRecordIndex = New BinaryReader(MyFileStreamOfRecordIndex)
    With MyBinaryReaderOfRecordIndex
      .BaseStream.Seek(0, SeekOrigin.Begin)
      MyPositionEndIndex = .ReadInt64
      MyPositionEndRecord = .ReadInt64
      MyFileCount = .ReadInt32
      MyMaxKeyID = .ReadInt32
      MyDateStart = DateTime.FromBinary(.ReadInt64)
      MyDateStop = DateTime.FromBinary(.ReadInt64)
    End With
  End Sub

  Public Sub New(ByVal StreamBaseName As String, ByVal FileMode As FileMode, ByRef DateRange As IDateRange, ByVal ChildPath As String, ByVal KeyName As String, ByVal FileExtension As String)
    Me.New(StreamBaseName, FileMode, DateRange, ChildPath, KeyName, FileExtension, IsReadOnly:=False)
  End Sub
#End Region 'New

  Public Property Exception() As Exception

#Region "IRecordIndex implementation"
  Public ReadOnly Property MaxID As Integer Implements IRecordIndex(Of T).MaxID
    Get
      Return MyMaxKeyID
    End Get
  End Property

  Public ReadOnly Property FileCount As Integer Implements IRecordIndex(Of T).FileCount
    Get
      Return MyFileCount
    End Get
  End Property

  ''' <summary>
  ''' Return the stream position for the object at the specified index  
  ''' </summary>
  ''' <param name="Index">
  ''' zero base index of the object
  ''' </param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ToPosition(Index As Integer) As Long Implements IRecordIndex(Of T).ToPosition
    Return ReadRecordIndex(Index).Position
  End Function

  Public ReadOnly Property ToListPosition As System.Collections.Generic.List(Of Long) Implements IRecordIndex(Of T).ToListPosition
    Get
      If MyFileMode = FileMode.Open Then
        If IsPositionLoaded = False Then
          Dim ThisPosition As Integer
          Dim ThisRecordData As strRecordIndex

          IsPositionLoaded = True
          With MyBinaryReaderOfRecordIndex
            If MyFileCount > 0 Then
              'search for the new position
              'return nothing if the date range is outside the current file range 
              If MyDateRange.DateStart <= MyDateStop Then
                If MyDateRange.DateStop >= MyDateStart Then
                  ThisPosition = BinarySearch(MyDateRange.DateStart)
                  'the list size is not know a priori but still
                  'estimate the list size on the high side for maximum speed
                  MyListPosition.Capacity = MyFileCount - ThisPosition
                  With MyBinaryReaderOfRecordIndex
                    .BaseStream.Seek(SIZE_OF_RECORD_INDEX * ThisPosition + POSITION_FILE_START_OFFSET, SeekOrigin.Begin)
                    Do
                      ThisRecordData.DateUpdate = DateTime.FromBinary(.ReadInt64)
                      ThisRecordData.Position = .ReadInt64
                      If ThisRecordData.DateUpdate <= MyDateRange.DateStop Then
                        MyListPosition.Add(ThisRecordData.Position)
                      Else
                        Exit Do
                      End If
                      ThisPosition = ThisPosition + 1
                    Loop Until ThisPosition = MyFileCount
                  End With
                End If
              End If
            End If
          End With
        End If
      End If
      Return MyListPosition
    End Get
  End Property

  Public ReadOnly Property BaseStream As System.IO.Stream Implements IRecordIndex(Of T).BaseStream
    Get
      Return MyFileStreamOfRecord
    End Get
  End Property

  Public ReadOnly Property BaseStream(Position As Long) As System.IO.Stream Implements IRecordIndex(Of T).BaseStream
    Get
      MyFileStreamOfRecord.Position = Position
      Return MyFileStreamOfRecord
    End Get
  End Property

  Public ReadOnly Property BaseStreamIndex As System.IO.Stream Implements IRecordIndex(Of T).BaseStreamIndex
    Get
      Return MyFileStreamOfRecordIndex
    End Get
  End Property

  Public ReadOnly Property BaseStreamIndex(Position As Long) As System.IO.Stream Implements IRecordIndex(Of T).BaseStreamIndex
    Get
      MyFileStreamOfRecordIndex.Position = Position
      Return MyFileStreamOfRecordIndex
    End Get
  End Property

  Public Sub Close() Implements IRecordIndex(Of T).Close
    If MyBinaryReaderOfRecordIndex IsNot Nothing Then
      MyBinaryReaderOfRecordIndex.Dispose()
    End If
    If MyFileStreamOfRecord IsNot Nothing Then
      MyFileStreamOfRecord.Dispose()
    End If
    If MyFileStreamOfRecordIndex IsNot Nothing Then
      MyFileStreamOfRecordIndex.Dispose()
    End If
  End Sub

  Public Sub Save(ByRef colData As System.Collections.Generic.ICollection(Of T)) Implements IRecordIndex(Of T).Save
    Dim ThisT As T
    Dim ThisBinaryWriterOfRecordIndex As BinaryWriter
    Dim ThisPosition As Long
    Dim ThisFileLengthBeforeChange As Long

    Me.Exception = Nothing
    ThisBinaryWriterOfRecordIndex = New BinaryWriter(MyFileStreamOfRecordIndex)
    IsPositionLoaded = False
    'write the maximum keyID always in the last record
    If colData.Count > 0 Then
      If MyFileCount = 0 Then
        'initialize the date
        MyDateStart = colData.First.DateUpdate
        MyDateStop = MyDateStart
        MyMaxKeyID = 0
      End If
      'egual time is possible at the beginning of the file or in case of different KeyID but egual update time
      For Each ThisT In colData
        If ThisT.DateUpdate >= MyDateStop Then
          If ThisT.KeyID <= MyMaxKeyID Then
            'update the object KeyID to reflect the other object in the file
            ThisT.KeyID = MyMaxKeyID + 1
          End If
          'ready to update the index header to reflect the new object
          ThisFileLengthBeforeChange = MyFileStreamOfRecord.Length
          ThisPosition = MyPositionEndRecord
          MyFileStreamOfRecord.Position = ThisPosition
          Try
            'save the full record at the end of the record file 
            ThisT.SerializeSaveTo(MyFileStreamOfRecord)
          Catch ex As Exception
            'restore the file to the original position
            Me.Exception = New Exception(String.Format("Saving record error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
            Exit Sub
          End Try
          MyPositionEndRecord = MyFileStreamOfRecord.Position
          If MyPositionEndRecord > ThisFileLengthBeforeChange Then
            'increase the file size
            MyFileStreamOfRecord.SetLength(MyPositionEndRecord + FILE_GROW_SIZE_MINIMUM_RECORD)
          End If
          'finalize the saving to the record index
          MyDateStop = ThisT.DateUpdate
          With ThisBinaryWriterOfRecordIndex
            Try
              MyFileCount = MyFileCount + 1
              MyMaxKeyID = ThisT.KeyID
              ThisFileLengthBeforeChange = .BaseStream.Length
              MyPositionEndIndex = MyPositionEndIndex + SIZE_OF_RECORD_INDEX
              .Seek(0, SeekOrigin.Begin)
              .Write(MyPositionEndIndex)
              .Write(MyPositionEndRecord)
              .Write(MyFileCount)
              .Write(MyMaxKeyID)
              .Write(MyDateStart.ToBinary)
              .Write(MyDateStop.ToBinary)
              .BaseStream.Position = MyPositionEndIndex - SIZE_OF_RECORD_INDEX
              .Write(MyDateStop.ToBinary)
              .Write(ThisPosition)
              If MyPositionEndIndex > ThisFileLengthBeforeChange Then
                'increase the file size
                .BaseStream.SetLength(MyPositionEndIndex + FILE_GROW_SIZE_MINIMUM_INDEX)
              End If
            Catch ex As Exception
              Me.Exception = New Exception(String.Format("Saving record index error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
              Exit Sub
            End Try
          End With
        End If
      Next
    End If
  End Sub

  Public Sub Save(ByRef colData As ICollection(Of T), IsSaveAtEndOfDay As Boolean) Implements IRecordIndex(Of T).Save
    Dim ThisT As T
    Dim ThisTNext As T
    Dim ThisBinaryWriterOfRecordIndex As BinaryWriter
    Dim ThisPosition As Long
    Dim ThisFileLengthBeforeChange As Long
    Dim IsSaveRecordNew As Boolean
    Dim IsSaveRecordOverLast As Boolean
    Dim IsSaveRecordOverLastCompleted As Boolean

    If IsSaveAtEndOfDay = False Then
      Me.Save(colData)
      Exit Sub
    End If
    Me.Exception = Nothing
    ThisBinaryWriterOfRecordIndex = New BinaryWriter(MyFileStreamOfRecordIndex)

    IsPositionLoaded = False
    'write the maximum keyID always in the last record
    If colData.Count > 0 Then
      If MyFileCount = 0 Then
        'initialize the date
        MyDateStart = colData.First.DateUpdate
        MyDateStop = MyDateStart
        MyMaxKeyID = 0
      End If
      IsSaveRecordOverLastCompleted = False
      For I = 0 To (colData.Count - 1)
        ThisT = colData(I)
        If I < (colData.Count - 1) Then
          ThisTNext = colData(I + 1)
        Else
          ThisTNext = Nothing
        End If
        If ThisT.DateUpdate >= MyDateStop Then
          IsSaveRecordOverLast = False
          IsSaveRecordNew = False
          If MyFileCount = 0 Then
            'only a new record is possible in this case
            'need to check the DateUpdate of the next record to make a decision
            If ThisTNext Is Nothing Then
              'always save the last data
              IsSaveRecordNew = True
              IsSaveRecordOverLastCompleted = True
            Else
              If ThisT.DateDay <> ThisTNext.DateDay Then
                'this the last data for the day
                IsSaveRecordNew = True
                IsSaveRecordOverLastCompleted = True
              End If
            End If
          Else
            If IsSaveRecordOverLastCompleted = False Then
              'two conditions are possible 
              're-write over the record or a new file record?
              'check the date of the last record in the file
              If MyDateStop.Date = ThisT.DateDay Then
                're-write over last record on these conditions
                If ThisTNext Is Nothing Then
                  'always re-write over the last data
                  IsSaveRecordOverLast = True
                  IsSaveRecordOverLastCompleted = True
                Else
                  If ThisT.DateDay <> ThisTNext.DateDay Then
                    'this the last data for the day
                    IsSaveRecordOverLast = True
                    IsSaveRecordOverLastCompleted = True
                  End If
                End If
              Else
                'do not re-write over 
                'need to check the DateUpdate of the next record to make a decision
                If ThisTNext Is Nothing Then
                  'always save the last data
                  IsSaveRecordNew = True
                  IsSaveRecordOverLastCompleted = True
                Else
                  If ThisT.DateDay <> ThisTNext.DateDay Then
                    'this the last data for the day
                    IsSaveRecordNew = True
                    IsSaveRecordOverLastCompleted = True
                  End If
                End If
              End If
            Else
              'only a new record is possible at this point
              'need to check the DateUpdate of the next data to make a decision
              If ThisTNext Is Nothing Then
                'always save the last data
                IsSaveRecordNew = True
              Else
                If ThisT.DateDay <> ThisTNext.DateDay Then
                  'this the last data for the day
                  IsSaveRecordNew = True
                End If
              End If
            End If
          End If
          'proceed with the required action
          If IsSaveRecordOverLast Then
            'take the same KeyID than the last record
            ThisT.KeyID = MyMaxKeyID
            'proceed with saving over the last record
            ThisFileLengthBeforeChange = MyFileStreamOfRecord.Length
            'get the file position of the last record
            ThisPosition = Me.ReadRecordIndex(MyFileCount - 1).Position
            MyFileStreamOfRecord.Position = ThisPosition
            Try
              'save the full record over the last record
              ThisT.SerializeSaveTo(MyFileStreamOfRecord)
            Catch ex As Exception
              Me.Exception = New Exception(String.Format("Saving over record error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
              Exit Sub
            End Try
            MyPositionEndRecord = MyFileStreamOfRecord.Position
            If MyPositionEndRecord > ThisFileLengthBeforeChange Then
              'increase the file size
              MyFileStreamOfRecord.SetLength(MyPositionEndRecord + FILE_GROW_SIZE_MINIMUM_RECORD)
            End If
            'finalize the saving to the record index
            MyDateStop = ThisT.DateUpdate
            With ThisBinaryWriterOfRecordIndex
              Try
                'nothing to change here  except the date since we just re-write the last index 
                'MyFileCount = MyFileCount 
                'MyMaxKeyID = ThisT.KeyID
                'MyPositionEndIndex = MyPositionEndIndex 
                .Seek(0, SeekOrigin.Begin)
                .Write(MyPositionEndIndex)
                .Write(MyPositionEndRecord)
                .Write(MyFileCount)
                .Write(MyMaxKeyID)
                .Write(MyDateStart.ToBinary)
                .Write(MyDateStop.ToBinary)
                'position to save over the last record index
                .BaseStream.Position = MyPositionEndIndex - SIZE_OF_RECORD_INDEX
                .Write(MyDateStop.ToBinary)
                .Write(ThisPosition)
              Catch ex As Exception
                Me.Exception = New Exception(String.Format("Saving record index error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
                Exit Sub
              End Try
            End With
          ElseIf IsSaveRecordNew Then
            If ThisT.KeyID <= MyMaxKeyID Then
              'update the object KeyID to reflect the other object in the file
              ThisT.KeyID = MyMaxKeyID + 1
            End If
            'ready to update the index header to reflect the new object
            ThisFileLengthBeforeChange = MyFileStreamOfRecord.Length
            ThisPosition = MyPositionEndRecord
            MyFileStreamOfRecord.Position = ThisPosition
            Try
              'save the full record at the end of the record file 
              ThisT.SerializeSaveTo(MyFileStreamOfRecord)
            Catch ex As Exception
              'restore the file to the original position
              Me.Exception = New Exception(String.Format("Saving record error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
              Exit Sub
            End Try
            MyPositionEndRecord = MyFileStreamOfRecord.Position
            If MyPositionEndRecord > ThisFileLengthBeforeChange Then
              'increase the file size
              MyFileStreamOfRecord.SetLength(MyPositionEndRecord + FILE_GROW_SIZE_MINIMUM_RECORD)
            End If
            'finalize the saving to the record index
            MyDateStop = ThisT.DateUpdate
            With ThisBinaryWriterOfRecordIndex
              Try
                MyFileCount = MyFileCount + 1
                MyMaxKeyID = ThisT.KeyID
                ThisFileLengthBeforeChange = .BaseStream.Length
                MyPositionEndIndex = MyPositionEndIndex + SIZE_OF_RECORD_INDEX
                .Seek(0, SeekOrigin.Begin)
                .Write(MyPositionEndIndex)
                .Write(MyPositionEndRecord)
                .Write(MyFileCount)
                .Write(MyMaxKeyID)
                .Write(MyDateStart.ToBinary)
                .Write(MyDateStop.ToBinary)
                .BaseStream.Position = MyPositionEndIndex - SIZE_OF_RECORD_INDEX
                .Write(MyDateStop.ToBinary)
                .Write(ThisPosition)
                If MyPositionEndIndex > ThisFileLengthBeforeChange Then
                  'increase the file size
                  .BaseStream.SetLength(MyPositionEndIndex + FILE_GROW_SIZE_MINIMUM_INDEX)
                End If
              Catch ex As Exception
                Me.Exception = New Exception(String.Format("Saving record index error for object of type {0} with key value {1}", ThisT.GetType.ToString, ThisT.KeyValue))
                Exit Sub
              End Try
            End With
          Else
            'do nothing with this intra day record
          End If
        End If
      Next
    End If
  End Sub

  Public Sub Save(ByRef ThisT As T) Implements IRecordIndex(Of T).Save
    Dim ThisList As ICollection(Of T) = New List(Of T)

    ThisList.Add(ThisT)
    Me.Save(ThisList)
  End Sub

  Public Sub Save(ByRef ThisT As T, IsSaveOnlyEndOfDay As Boolean) Implements IRecordIndex(Of T).Save
    Dim ThisList As ICollection(Of T) = New List(Of T)

    ThisList.Add(ThisT)
    Me.Save(ThisList, IsSaveOnlyEndOfDay)
  End Sub
#End Region    'IRecordIndex
#Region "Private Local function"
  Private Function ReadRecordIndex(ByVal Index As Integer) As strRecordIndex
    Dim ThisRecordIndex As strRecordIndex

    With MyBinaryReaderOfRecordIndex
      .BaseStream.Seek(SIZE_OF_RECORD_INDEX * Index + POSITION_FILE_START_OFFSET, SeekOrigin.Begin)
      ThisRecordIndex.DateUpdate = DateTime.FromBinary(.ReadInt64)
      ThisRecordIndex.Position = .ReadInt64
    End With
    Return ThisRecordIndex
  End Function

  Private Function ReadRecordIndexDate(ByVal Index As Integer) As Date
    Dim ThisDataUpdate As Date

    With MyBinaryReaderOfRecordIndex
      .BaseStream.Seek(SIZE_OF_RECORD_INDEX * Index + POSITION_FILE_START_OFFSET, SeekOrigin.Begin)
      ThisDataUpdate = DateTime.FromBinary(.ReadInt64)
    End With
    Return ThisDataUpdate
  End Function

  Private Function BinarySearch(ByVal KeyDate As Date) As Integer
    Dim Low As Integer
    Dim High As Integer
    Dim Middle As Integer
    Dim LastElement As Integer
    Dim ThisDateMiddle As Date
    Dim ThisDateHigh As Date

    LastElement = MyFileCount - 1
    'protection
    If LastElement <= 0 Then
      Return LastElement
    End If
    Low = 0
    High = LastElement
    Do
      Middle = (Low + High) \ 2
      ThisDateMiddle = ReadRecordIndexDate(Middle)
      If ThisDateMiddle = KeyDate Then
        Return Middle
        Exit Do
      ElseIf ((ThisDateMiddle < KeyDate)) Then
        If Low = Middle Then
          'return the closest element
          ThisDateHigh = ReadRecordIndexDate(High)
          If Math.Abs(ThisDateMiddle.Subtract(KeyDate).Ticks) <= Math.Abs(ThisDateHigh.Subtract(KeyDate).Ticks) Then
            Return Low
          Else
            Return High
          End If
        Else
          Low = Middle + 1
        End If
      Else
        If Low = Middle Then
          'return the closest element
          ThisDateHigh = ReadRecordIndexDate(High)
          If Math.Abs(ThisDateMiddle.Subtract(KeyDate).Ticks) <= Math.Abs(ThisDateHigh.Subtract(KeyDate).Ticks) Then
            Return Low
          Else
            Return High
          End If
        Else
          High = Middle - 1
        End If
      End If
    Loop
  End Function
#End Region  'Private Local function
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' dispose managed state (managed objects).
        If MyBinaryReaderOfRecordIndex IsNot Nothing Then
          MyBinaryReaderOfRecordIndex.Dispose()
        End If
        If MyFileStreamOfRecord IsNot Nothing Then
          MyFileStreamOfRecord.Dispose()
        End If
        If MyFileStreamOfRecordIndex IsNot Nothing Then
          MyFileStreamOfRecordIndex.Dispose()
        End If
      End If
      ' free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

  'override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  Protected Overrides Sub Finalize()
    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    Dispose(False)
    MyBase.Finalize()
  End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub
#End Region
#Region "IDateUpdate"
  Public Property DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return MyDateStart
    End Get
    Set(value As Date)
    End Set
  End Property

  Public Property DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return MyDateStop
    End Get
    Set(value As Date)
    End Set
  End Property


  Public ReadOnly Property KeyName As String
    Get
      Return MyKeyName
    End Get
  End Property
  Public ReadOnly Property DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return Me.DateStop
    End Get
  End Property

  Public ReadOnly Property DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return Me.DateStop
    End Get
  End Property

  Public ReadOnly Property DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateStop.Date
    End Get
  End Property
#End Region  'IDateUpdate
#Region "IDateTrade"
  Public ReadOnly Property AsDateTrade As YahooAccessData.IDateTrade
    Get
      Return Me
    End Get
  End Property

  Private Property IDateTrade_DateStart As Date Implements IDateTrade.DateStart
    Get
      Return Me.DateStart
    End Get
    Set(value As Date)

    End Set
  End Property

  Private Property IDateTrade_DateStop As Date Implements IDateTrade.DateStop
    Get
      Return Me.DateStop
    End Get
    Set(value As Date)

    End Set
  End Property
#End Region 'IDateTrade
End Class
#End Region   'RecordIndex(Of T)
