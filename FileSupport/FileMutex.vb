Imports System.IO
Imports System.Threading
Imports System.Xml.Serialization
Imports ExcelDataReader
Imports ExcelDataReader.ExcelDataReaderExtensions

Namespace ExtensionServiceMutex
  Public Class FileMutex
    Implements IDisposable

    Private disposedValue As Boolean
    Private MyMutex As Mutex

    ''' <summary>
    ''' Create a file protection mutex for multiple interprocess access to the different local file
    ''' </summary>
    ''' <param name="initiallyOwned">see definition of system mutex</param>
    ''' <param name="FileMutexName">The name for the mutex (see definition of system mutex)</param>
    Public Sub New(ByVal initiallyOwned As Boolean, ByVal FileMutexName As String)
      Try
        MyMutex = New Mutex(initiallyOwned, FileMutexName)
      Catch ex As Exception
        Throw New InvalidProgramException("Error with Mutex creation...", ex)
      End Try
    End Sub

    Public Sub FileHeaderSave(ByVal FileName As String, ByRef HeaderInfo As List(Of HeaderInfo))
      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileHeaderSave(FileName, HeaderInfo)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Function FileHeaderRead(
      ByVal FileName As String,
      ByRef HeaderInfoDefault As List(Of HeaderInfo),
      ByRef Exception As Exception) As List(Of HeaderInfo)

      Dim ThisListOfData As List(Of HeaderInfo) = Nothing

      MyMutex.WaitOne()
      Try
        ThisListOfData = YahooAccessData.ExtensionService.Extensions.FileHeaderRead(FileName, HeaderInfoDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisListOfData
    End Function

    Public Sub FileListSave(Of T)(ByVal FileName As String, ByRef Data As List(Of T))
      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileListSave(Of T)(FileName, Data)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Sub FileListSave(Of T)(ByVal FileName As String, ByRef Data As T)
      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileListSave(Of T)(FileName, Data)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    ''' <summary>
    ''' Read an excel file of type extension format ".xls" (97-2003), ".xlsx (2007 format) and ".csv" format
    ''' </summary>
    ''' <param name="FileName">The file name</param>
    ''' <returns>The dataset extracted from teh file reading</returns>
    Public Function FileExcelRead(ByVal FileName As String, Optional IsHeaderRow As Boolean = False) As DataSet
      Dim ThisWorkBookDataSet As DataSet = Nothing

      If My.Computer.FileSystem.FileExists(FileName) = False Then
        Throw New FileNotFoundException(FileName)
      End If
      MyMutex.WaitOne()
      Select Case Path.GetExtension(FileName).ToLower
        Case ".xls"
          'Reading from a binary Excel file ('97-2003 format; *.xls)
          Try
            Using StreamReader = File.Open(FileName, FileMode.Open, FileAccess.Read)
              Using ExcelReader As ExcelDataReader.IExcelDataReader = ExcelDataReader.ExcelReaderFactory.CreateReader(StreamReader)
                ThisWorkBookDataSet = ExcelReader.AsDataSet(New ExcelDataSetConfiguration() With {.ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {.UseHeaderRow = IsHeaderRow}})
              End Using
            End Using
          Catch ex As Exception
            MyMutex.ReleaseMutex()
            Throw ex
          End Try
        Case ".xlsx"
          'Reading from a OpenXml Excel file (2007 format; *.xlsx)
          Try
            Using StreamReader = File.Open(FileName, FileMode.Open, FileAccess.Read)
              Using ExcelReader As ExcelDataReader.IExcelDataReader = ExcelDataReader.ExcelReaderFactory.CreateOpenXmlReader(StreamReader)
                ThisWorkBookDataSet = ExcelReader.AsDataSet(New ExcelDataSetConfiguration() With {.ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {.UseHeaderRow = True}})
              End Using
            End Using
          Catch ex As Exception
            MyMutex.ReleaseMutex()
            Throw ex
          End Try
        Case ".csv"
          'not tested
          'Reading from a csv Excel file (2007 format; *.csv)
          Debugger.Break()
          Try
            Using StreamReader = File.Open(FileName, FileMode.Open, FileAccess.Read)
              Using ExcelReader As ExcelDataReader.IExcelDataReader = ExcelDataReader.ExcelReaderFactory.CreateCsvReader(StreamReader)
                ThisWorkBookDataSet = ExcelReader.AsDataSet(New ExcelDataSetConfiguration() With {.ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {.UseHeaderRow = True}})
              End Using
            End Using
          Catch ex As Exception
            MyMutex.ReleaseMutex()
            Throw ex
          End Try
        Case Else
          Throw New Exception("Invalid excel extension name...")
      End Select
      MyMutex.ReleaseMutex()
      Return ThisWorkBookDataSet
    End Function

    Public Function FileListRead(Of TKey, TValue)(
      ByVal FileName As String,
      ByRef DataDefault As Dictionary(Of TKey, TValue),
      ByRef Exception As Exception) As Dictionary(Of TKey, TValue)

      Dim ThisListOfData As Dictionary(Of TKey, TValue) = Nothing

      MyMutex.WaitOne()
      Try
        ThisListOfData = YahooAccessData.ExtensionService.Extensions.FileListRead(Of TKey, TValue)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisListOfData
    End Function

    Public Sub FileListSave(Of TKey, TValue)(ByVal FileName As String, ByRef Data As Dictionary(Of TKey, TValue))
      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileListSave(Of TKey, TValue)(FileName, Data)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Sub FileDelete(ByVal FileName As String)
      MyMutex.WaitOne()
      Try
        My.Computer.FileSystem.DeleteFile(FileName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.DeletePermanently)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Function FileReadOfDictionary(Of TKey, TValue)(
      ByVal FileName As String,
      ByRef DataDefaultOnError As Dictionary(Of TKey, TValue),
      ByRef Exception As Exception,
      ByVal FileType As ExtensionService.EnumFileType) As Dictionary(Of TKey, TValue)

      Dim ThisData As Dictionary(Of TKey, TValue) = Nothing
      MyMutex.WaitOne()
      Try
        ThisData = YahooAccessData.ExtensionService.Extensions.FileReadOfDictionary(Of TKey, TValue)(FileName, DataDefaultOnError, Exception, FileType)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisData
    End Function

    Public Function FileListReadBinary(Of TKey, TValue)(
      ByVal FileName As String,
      ByRef DataDefault As Dictionary(Of TKey, TValue),
      ByRef Exception As Exception) As Dictionary(Of TKey, TValue)

      Dim ThisData As Dictionary(Of TKey, TValue) = Nothing

      MyMutex.WaitOne()
      Try
        ThisData = YahooAccessData.ExtensionService.Extensions.FileListReadBinary(Of TKey, TValue)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisData
    End Function

    Public Sub FileSaveOfDictionary(Of TKey, TValue)(
      ByVal FileName As String,
      ByRef Data As Dictionary(Of TKey, TValue),
      ByVal FileType As ExtensionService.EnumFileType)

      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileSaveOfDictionary(Of TKey, TValue)(FileName, Data, FileType:=FileType)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Sub FileListSaveBinary(Of TKey, TValue)(
      ByVal FileName As String,
      ByRef Data As Dictionary(Of TKey, TValue))

      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileListSaveBinary(Of TKey, TValue)(FileName, Data)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Function FileReadBinary(Of T)(
      ByVal FileName As String,
      ByRef DataDefault As T,
      ByRef Exception As Exception) As T

      Dim ThisData As T = Nothing

      MyMutex.WaitOne()
      Try
        ThisData = YahooAccessData.ExtensionService.Extensions.FileReadBinary(Of T)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisData
    End Function

    Public Sub FileSaveBinary(Of T)(
      ByVal FileName As String,
      ByRef Data As T)

      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileSaveBinary(Of T)(FileName, Data)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Function FileListRead(Of T)(
      ByVal FileName As String,
      ByRef DataDefault As T,
      ByRef Exception As Exception) As T

      Dim ThisData As T

      MyMutex.WaitOne()
      Try
        ThisData = YahooAccessData.ExtensionService.Extensions.FileListRead(Of T)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisData
    End Function

    Public Function FileListRead(Of T)(
      ByVal FileName As String,
      ByRef DataDefault As List(Of T),
      ByRef Exception As Exception) As List(Of T)

      Dim ThisListOfData As List(Of T) = Nothing

      MyMutex.WaitOne()
      Try
        ThisListOfData = YahooAccessData.ExtensionService.Extensions.FileListRead(Of T)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisListOfData
    End Function

    Public Function FileListRead(Of T)(ByVal FileName As String) As List(Of T)

      Dim ThisListOfData As List(Of T) = Nothing
      MyMutex.WaitOne()
      Try
        ThisListOfData = YahooAccessData.ExtensionService.Extensions.FileListRead(Of T)(FileName)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisListOfData
    End Function

    Public Sub FileListSave(Of T As {ITreeNode(Of U)}, U)(
      ByVal FileName As String,
      ByRef HeaderInfo As List(Of T))

      MyMutex.WaitOne()
      Try
        YahooAccessData.ExtensionService.Extensions.FileListSave(Of T, U)(FileName, HeaderInfo)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
    End Sub

    Public Function FileListRead(Of T, U)(
      ByVal FileName As String,
      ByRef DataDefault As List(Of T),
      ByRef Exception As Exception) As List(Of T)

      Dim ThisListOfData As List(Of T) = Nothing
      MyMutex.WaitOne()
      Try
        ThisListOfData = YahooAccessData.ExtensionService.Extensions.FileListRead(Of T, U)(FileName, DataDefault, Exception)
      Catch ex As Exception
        Throw ex
      Finally
        'always release
        MyMutex.ReleaseMutex()
      End Try
      Return ThisListOfData
    End Function


    Protected Overrides Sub Finalize()
      Me.Dispose()
      MyBase.Finalize()
    End Sub

#Region "Dispose"
    Protected Overridable Sub Dispose(disposing As Boolean)
      If Not disposedValue Then
        If disposing Then
          ' dispose managed state (managed objects)
        End If
        ' free unmanaged resources (unmanaged objects) and override finalizer
        ' set large fields to null
        If MyMutex IsNot Nothing Then
          MyMutex.Close()
          MyMutex.Dispose()
        End If
        disposedValue = True
      End If
    End Sub

    ' ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    ' Protected Overrides Sub Finalize()
    '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
    '     Dispose(disposing:=False)
    '     MyBase.Finalize()
    ' End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
      ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
      Dispose(disposing:=True)
      GC.SuppressFinalize(Me)
    End Sub
  End Class
#End Region
End Namespace