Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Xml.Serialization
Imports ExcelDataReader
Imports ExcelDataReader.ExcelDataReaderExtensions
Imports Newtonsoft.Json
Imports YahooAccessData.ExtensionService.Extensions

Namespace ExtensionServiceMutex
	Public Class FileMutex
		Implements IDisposable

		Private disposedValue As Boolean
		Private MyMutex As Mutex

		''' <summary>
		''' Create a file protection mutex for multiple inter-process access to the different local file
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

		'Public Sub FileHeaderSave(ByVal FileName As String, ByRef HeaderInfo As List(Of HeaderInfo))
		'	MyMutex.WaitOne()
		'	Try
		'		Me.FileHeaderSaveLocal(FileName, HeaderInfo)
		'	Catch ex As Exception
		'		Throw ex
		'	Finally
		'		'always release
		'		MyMutex.ReleaseMutex()
		'	End Try
		'End Sub

		'Public Function FileReadSettingOfDictionary(Of TKey, TValue)(
		'	ByVal FileName As String,
		'	ByRef DataDefaultOnError As Dictionary(Of TKey, TValue),
		'	ByRef Exception As Exception,
		'	ByVal FileType As ExtensionService.EnumFileType) As Dictionary(Of TKey, TValue)

		'	Dim ThisData As Dictionary(Of TKey, TValue) = Nothing
		'	MyMutex.WaitOne()
		'	Try
		'		ThisData = Me.FileReadOfDictionaryLocal(Of TKey, TValue)(FileName, DataDefaultOnError, Exception, FileType)
		'	Catch ex As Exception
		'		Throw ex
		'	Finally
		'		'always release
		'		MyMutex.ReleaseMutex()
		'	End Try
		'	Return ThisData
		'End Function

		Public Function FileHeaderRead(
			ByVal FileName As String,
			ByRef HeaderInfoDefault As List(Of HeaderInfo),
			ByRef Exception As Exception) As List(Of HeaderInfo)

			Dim ThisListOfData As List(Of HeaderInfo) = Nothing

			'MyMutex.WaitOne()
			Try
				ThisListOfData = FileListRead(Of HeaderInfo)(FileName, HeaderInfoDefault, Exception)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				'MyMutex.ReleaseMutex()
			End Try
			Return ThisListOfData
		End Function

		Public Sub FileListSave(Of T)(ByVal FileName As String, ByRef Data As List(Of T))
			MyMutex.WaitOne()
			Try
				Me.FileListSaveLocal(Of T)(FileName, Data)
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

		Public Function FileListRead(Of T)(
			ByVal FileName As String,
			ByRef DataDefault As List(Of T),
			ByRef Exception As Exception) As List(Of T)

			Dim ThisData As List(Of T)

			MyMutex.WaitOne()
			Try
				ThisData = Me.FileListReadLocal(Of T)(FileName, DataDefault, Exception)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				MyMutex.ReleaseMutex()
			End Try
			Return ThisData
		End Function

		Private Function FileListReadLocal(Of T, U)(ByVal FileName As String, ByRef DataDefault As List(Of T), ByRef Exception As Exception) As List(Of T)
			Dim ThisException As Exception = Nothing
			Dim ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
			Dim ThisTextReader As System.IO.TextReader
			Dim ThisListInfo As List(Of T)
			Const CREATE_FILE_ALWAYS As Boolean = False 'should be false by default

			If CREATE_FILE_ALWAYS Then
				If My.Computer.FileSystem.FileExists(FileName) Then
					My.Computer.FileSystem.DeleteFile(FileName)
				End If
			End If
			If My.Computer.FileSystem.FileExists(FileName) Then
				Try
					ThisTextReader = New StreamReader(FileName)
					ThisListInfo = CType(ThisXmlSerializer.Deserialize(ThisTextReader), List(Of T))
					ThisTextReader.Dispose()
				Catch ex As Exception
					Exception = ex
					ThisListInfo = DataDefault
				End Try
			Else
				'get the default header info 
				ThisListInfo = DataDefault
				'save the default file
				Try
					Me.FileListSaveLocal(Of T)(FileName, ThisListInfo)
				Catch ex As Exception
					Exception = ex
				End Try
			End If
			Return ThisListInfo
		End Function

		Public Function FileListRead(Of TKey, TValue)(
			ByVal FileName As String,
			ByRef DataDefault As Dictionary(Of TKey, TValue),
			ByRef Exception As Exception) As Dictionary(Of TKey, TValue)

			Dim ThisListOfData As Dictionary(Of TKey, TValue) = Nothing

			MyMutex.WaitOne()
			Try
				ThisListOfData = Me.FileListReadLocal(Of TKey, TValue)(FileName, DataDefault, Exception)
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
				Me.FileListSaveLocal(Of TKey, TValue)(FileName, Data)
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
			ByVal FileType As ExtensionService.EnumFileType) As Dictionary(Of TKey, TValue)

			Dim ThisData As Dictionary(Of TKey, TValue) = Nothing
			MyMutex.WaitOne()
			Try
				ThisData = Me.FileReadOfDictionaryLocal(Of TKey, TValue)(FileName, FileType)
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
				Return Me.FileReadOfDictionaryLocal(Of TKey, TValue)(
					FileName:=FileName,
					FileType:=EnumFileType.Binary)
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
				Me.FileSaveOfDictionaryLocal(Of TKey, TValue)(FileName, Data, FileType:=FileType)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				MyMutex.ReleaseMutex()
			End Try
		End Sub

		Public Sub FileListSave(Of T As {ITreeNode(Of U)}, U)(
			ByVal FileName As String,
			ByRef HeaderInfo As List(Of T))

			MyMutex.WaitOne()
			Try
				Me.FileListSaveLocal(Of T, U)(FileName, HeaderInfo)
			Catch ex As Exception
				Throw ex
			Finally
				'always release
				MyMutex.ReleaseMutex()
			End Try
		End Sub

		Public Sub FileHeaderSave(ByVal FileName As String, ByRef HeaderInfo As List(Of HeaderInfo))
			FileListSave(Of HeaderInfo)(FileName, HeaderInfo)
		End Sub

		Public Sub FileListSave(Of T)(ByVal FileName As String, ByRef Data As T)
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing
			Try
				ThisXmlSerializer = New XmlSerializer(GetType(T))
			Catch ex As Exception
				Throw New Exception(String.Format("Unable to create Xml Serializer."), ex)
			End Try

			Dim ThisTextWriter As System.IO.TextWriter

			Debugger.Break()
			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0}", ThisPath), ex)
					End Try
				End If
			End With
			System.IO.File.Delete(FileName)
			ThisTextWriter = New StreamWriter(FileName)
			Try
				ThisXmlSerializer.Serialize(ThisTextWriter, Data)
				ThisTextWriter.Dispose()
			Catch ex As Exception
				Throw New Exception(String.Format("Unable to save to file {0}", FileName), ex)
			End Try
		End Sub


		Public Function FileListRead(Of T, U)(
			ByVal FileName As String,
			ByRef DataDefault As List(Of T),
			ByRef Exception As Exception) As List(Of T)

			Dim ThisListOfData As List(Of T) = Nothing
			MyMutex.WaitOne()
			Try
				ThisListOfData = Me.FileListReadLocal(Of T, U)(FileName, DataDefault, Exception)
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

		'#Region "File access"
		'		Private Sub FileHeaderSaveLocal(ByVal FileName As String, ByRef HeaderInfo As List(Of HeaderInfo))
		'			FileListSave(Of HeaderInfo)(FileName, HeaderInfo)
		'		End Sub

		'Private Function FileHeaderReadLocal(ByVal FileName As String, ByRef HeaderInfoDefault As List(Of HeaderInfo), ByRef Exception As Exception) As List(Of HeaderInfo)
		'	Return Me.FileListRead(Of HeaderInfo)(FileName, HeaderInfoDefault, Exception)
		'End Function

		Private Sub FileListSaveLocal(Of T)(ByVal FileName As String, ByRef Data As T)
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing
			Try
				ThisXmlSerializer = New XmlSerializer(GetType(T))
			Catch ex As Exception
				Throw New Exception(String.Format("Unable to create Xml Serializer."), ex)
			End Try

			Dim ThisTextWriter As System.IO.TextWriter

			Debugger.Break()
			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0}", ThisPath), ex)
					End Try
				End If
			End With
			System.IO.File.Delete(FileName)
			ThisTextWriter = New StreamWriter(FileName)
			Try
				ThisXmlSerializer.Serialize(ThisTextWriter, Data)
				ThisTextWriter.Dispose()
			Catch ex As Exception
				Throw New Exception(String.Format("Unable to save to file {0}", FileName), ex)
			End Try
		End Sub

		''' <summary>
		''' Save the file under and xml or json file format in function of the extension provided in the file name
		''' </summary>
		''' <typeparam name="T"></typeparam>
		''' <param name="FileName"></param>
		''' <param name="Data"></param>
		Private Sub FileListSaveLocal(Of T)(ByVal FileName As String, ByRef Data As List(Of T))
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing

			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0}", ThisPath), ex)
					End Try
				End If
			End With
			Dim ThisFileNameToCopy = Path.ChangeExtension(FileName, $"{Path.GetExtension(FileName)}.{"bak"}")
			If System.IO.File.Exists(FileName) Then
				'make a backup file
				Try
					System.IO.File.Copy(FileName, ThisFileNameToCopy, overwrite:=True)
				Catch ex As Exception
					Throw New Exception(String.Format("Unable to copy to a backup file {0}", ThisFileNameToCopy), ex)
				End Try
				Try
					System.IO.File.Delete(FileName)
				Catch ex As Exception
					Throw New Exception(String.Format("Unable to delete the file {0}", ThisFileNameToCopy), ex)
				End Try
			End If
			Select Case Path.GetExtension(FileName)
				Case ".json"
					Try
						Dim ThisJsonResult = JsonConvert.SerializeObject(Data, Formatting.Indented)
						File.WriteAllText(FileName, ThisJsonResult)
					Catch ex As Exception
						Throw ex
					End Try
					'Creates a new file, write the contents to the file, and then closes the file.
					'If the target file already exists, it is truncated and overwritten.
				Case ".Xml"
					Try
						ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create Xml Serializer."), ex)
					End Try
					Dim ThisTextWriter = New StreamWriter(FileName)
					Try
						ThisXmlSerializer.Serialize(ThisTextWriter, Data)
						'ThisSharpSerializer.Serialize(Data, filename:=FileName)
						ThisTextWriter.Dispose()
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to save to file {0}", FileName), ex)
						'restore the previous file from the backup file
					End Try
			End Select
		End Sub

		Private Function FileListReadLocal(Of TKey, TValue)(
			ByVal FileName As String,
			ByRef DataDefault As Dictionary(Of TKey, TValue),
			ByRef Exception As Exception) As Dictionary(Of TKey, TValue)

			Dim ThisList As New List(Of IDictionaryKeyValuePair(Of TKey, TValue))
			Dim ThisDictionaryResult As Dictionary(Of TKey, TValue)
			Dim ThisException As Exception = Nothing
			Dim ThisXmlSerializer As XmlSerializer
			Dim ThisTextReader As System.IO.TextReader
			Const CREATE_FILE_ALWAYS As Boolean = False 'should be false by default

			ThisXmlSerializer = New XmlSerializer(GetType(List(Of IDictionaryKeyValuePair(Of TKey, TValue))))
			If CREATE_FILE_ALWAYS Then
				If My.Computer.FileSystem.FileExists(FileName) Then
					My.Computer.FileSystem.DeleteFile(FileName)
				End If
			End If
			If My.Computer.FileSystem.FileExists(FileName) Then
				Try
					ThisTextReader = New StreamReader(FileName)
					ThisList = CType(ThisXmlSerializer.Deserialize(ThisTextReader), List(Of IDictionaryKeyValuePair(Of TKey, TValue)))
					ThisTextReader.Dispose()
				Catch ex As Exception
					If Exception IsNot Nothing Then
						Exception = ex
					End If
					ThisList.Clear()
				End Try
				ThisDictionaryResult = New Dictionary(Of TKey, TValue)
				For Each ThisData In ThisList
					ThisDictionaryResult.Add(ThisData.Key, ThisData.Value)
				Next
			Else
				'get the default 
				'save the default file
				Try
					Me.FileListSaveLocal(Of TKey, TValue)(FileName, DataDefault)
				Catch ex As Exception
					If Exception IsNot Nothing Then
						Exception = ex
					End If
				End Try
				ThisDictionaryResult = DataDefault
			End If
			Return ThisDictionaryResult
		End Function

		Private Sub FileListSaveLocal(Of TKey, TValue)(ByVal FileName As String, ByRef Data As Dictionary(Of TKey, TValue))
			Dim ThisList As New List(Of IDictionaryKeyValuePair(Of TKey, TValue))
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing
			Dim ThisDictionaryKeyValuePair As DictionaryKeyValuePair(Of TKey, TValue)

			If Data Is Nothing Then Return
			Try
				ThisXmlSerializer = New XmlSerializer(GetType(List(Of IDictionaryKeyValuePair(Of TKey, TValue))))
			Catch ex As Exception
				Throw New Exception(String.Format("Unable to create Xml Serializer: {0}"), ex)
			End Try

			Dim ThisTextWriter As System.IO.TextWriter

			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0},ThisPath"), ex)
					End Try
				End If
			End With
			System.IO.File.Delete(FileName)
			ThisTextWriter = New StreamWriter(FileName)
			For Each ThisData In Data
				ThisDictionaryKeyValuePair = New DictionaryKeyValuePair(Of TKey, TValue) With {.Key = ThisData.Key, .Value = ThisData.Value}
				ThisList.Add(ThisDictionaryKeyValuePair)
			Next
			ThisXmlSerializer.Serialize(ThisTextWriter, ThisList)
			ThisTextWriter.Dispose()
		End Sub

#Region "FileReadOfDictionary"
		Private Function FileReadOfDictionaryLocal(Of TKey, TValue)(
			ByVal FileName As String,
			ByVal FileType As EnumFileType,
			Optional IsRecoverFromXMLAndOrBin As Boolean = False) As Dictionary(Of TKey, TValue)

			Dim ThisFileNameXml As String
			Dim ThisSharpSerializer As Polenter.Serialization.SharpSerializer = Nothing
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisDictionaryResult As Dictionary(Of TKey, TValue) = Nothing
			Dim ThisException As Exception = Nothing
			Dim ThisJasonResult As String

			'JSON is a language-independent data format. It was derived from JavaScript, but many modern programming
			'languages include code to generate and parse JSON-format data. JSON filenames use the extension .json
			Select Case FileType
				Case EnumFileType.Binary
					'Throw New NotSupportedException("Invalid Binary file type...")
					ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=True)
					Try
						If ThisDictionaryResult Is Nothing Then
							ThisDictionaryResult = CType(ThisSharpSerializer.Deserialize(filename:=FileName), Dictionary(Of TKey, TValue))
						End If
					Catch ex As Exception
					End Try
					Return ThisDictionaryResult
				Case EnumFileType.XML
					Try
						ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=False)
						ThisDictionaryResult = CType(ThisSharpSerializer.Deserialize(filename:=FileName), Dictionary(Of TKey, TValue))
						Return ThisDictionaryResult
					Catch ex As Exception
						Try
							ThisJasonResult = File.ReadAllText(FileName)
							ThisDictionaryResult = JsonConvert.DeserializeObject(Of Dictionary(Of TKey, TValue))(ThisJasonResult)
							Return ThisDictionaryResult
						Catch ex1 As Exception
							Throw ex1
						End Try
					End Try
				Case EnumFileType.Json
					'check if the file exist
					'special case use to move the file to a jason standard
					If File.Exists(FileName) = False Then
						If IsRecoverFromXMLAndOrBin Then
							'check if a binary file exist
							ThisDictionaryResult = Nothing
							Dim ThisFileNameBin = Path.ChangeExtension(FileName, ".Bin")
							If File.Exists(ThisFileNameBin) Then
								'try to read the binary file and save it in a json file
								'use the old version of the sharp serializer
								Try
									ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=True)
									ThisDictionaryResult = CType(ThisSharpSerializer.Deserialize(filename:=ThisFileNameBin), Dictionary(Of TKey, TValue))
								Catch ex As Exception
									'ignore the error and try in xml
									ThisFileNameXml = Path.ChangeExtension(FileName, ".xml")
									Try
										If File.Exists(ThisFileNameXml) Then
											ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=False)
											ThisDictionaryResult = CType(ThisSharpSerializer.Deserialize(filename:=ThisFileNameXml), Dictionary(Of TKey, TValue))
										Else
											'no file of any type found
										End If
									Catch ex1 As Exception
									End Try
								End Try
							Else
								ThisFileNameXml = Path.ChangeExtension(FileName, ".xml")
								If File.Exists(ThisFileNameXml) Then
									Try
										ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=False)
										ThisDictionaryResult = CType(ThisSharpSerializer.Deserialize(filename:=ThisFileNameXml), Dictionary(Of TKey, TValue))
									Catch ex As Exception

									End Try
								Else
									'no file of any type found
								End If
							End If
							If ThisDictionaryResult IsNot Nothing Then
								'save to other format to the json format
								Me.FileSaveOfDictionaryLocal(Of TKey, TValue)(FileName, ThisDictionaryResult, FileType:=EnumFileType.Json)
							End If
						End If
					End If
					'at this point a json file exist unless there is no file of any type with that name
					Try
						ThisJasonResult = File.ReadAllText(FileName)
						ThisDictionaryResult = JsonConvert.DeserializeObject(Of Dictionary(Of TKey, TValue))(ThisJasonResult)
					Catch ex As Exception
						Throw ex
					End Try
					Return ThisDictionaryResult
				Case Else
					Throw New InvalidCastException("Invalid file type...")
			End Select
		End Function
#End Region
#Region "FileSaveOfDictionary"
		Private Sub FileSaveOfDictionaryLocal(Of TKey, TValue)(
			ByVal FileName As String,
			ByRef Data As Dictionary(Of TKey, TValue),
			ByVal FileType As EnumFileType)

			Dim ThisSharpSerializer As Polenter.Serialization.SharpSerializer
			Dim ThisJsonResult As String = ""
			Dim ThisFileBackupName As String = FileName + ".bak"
			If Data Is Nothing Then Return
			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			'Debugger.Break()

			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory: {0}"), ex)
					End Try
				End If
			End With
			With My.Computer.FileSystem
				If .FileExists(FileName) Then
					'copy the existing file to a backup file
					If .FileExists(ThisFileBackupName) Then
						'delete the backup file
						.DeleteFile(ThisFileBackupName)
						'rename the file to the backup name
						.RenameFile(FileName, Path.GetFileName(ThisFileBackupName))
						'ready to write the new file
					End If
				End If
			End With
			Dim ThisException As Exception = Nothing
			Try
				Select Case FileType
					Case EnumFileType.Binary
						Debugger.Break()
						ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=True)
						ThisSharpSerializer.Serialize(Data, FileName)
					Case EnumFileType.XML
						'ThisJsonResult = JsonConvert.SerializeObject(Data, Formatting.Indented)
						'Creates a new file, write the contents to the file, and then closes the file.
						'If the target file already exists, it is truncated and overwritten.
						'File.WriteAllText(FileName, ThisJsonResult)
						ThisSharpSerializer = New Polenter.Serialization.SharpSerializer(binarySerialization:=False)
						ThisSharpSerializer.Serialize(Data, FileName)
					Case EnumFileType.Json
						ThisJsonResult = JsonConvert.SerializeObject(Data, Formatting.Indented)
						'Creates a new file, write the contents to the file, and then closes the file.
						'If the target file already exists, it is truncated and overwritten.
						File.WriteAllText(FileName, ThisJsonResult)
				End Select
			Catch ex As Exception
				Throw ex
			Finally
			End Try
		End Sub

		''' <summary>
		''' Save the data in an xml or json file format based on the file extension *.json or *.xml
		''' If the extension is json but the file does not exist the function will automatically try to read the same file
		''' but with the xml extension and if successful will convert the file to Jason file. The xml file is not changed by the operation
		''' This characteristic is needed to help to move the xml file of this application to only the json standard providing a better visual
		''' for for inspection on web browser. 
		''' </summary>
		''' <typeparam name="T">The type in the list</typeparam>
		''' <param name="FileName"></param>
		''' <param name="DataDefault"></param>
		''' <param name="Exception"></param>
		''' <returns></returns>
		Private Function FileListReadLocal(Of T)(
			ByVal FileName As String,
			ByRef DataDefault As List(Of T),
			ByRef Exception As Exception) As List(Of T)

			Dim ThisException As Exception = Nothing
			Dim ThisXmlSerializer As XmlSerializer
			Dim ThisTextReader As System.IO.TextReader
			Dim ThisListInfo As List(Of T)
			Dim ThisFileExtension As String = Path.GetExtension(FileName).ToLower
			Dim ThisFileNameXML As String = Path.ChangeExtension(FileName, ".xml").ToLower
			Dim ThisFileData As String

			Select Case ThisFileExtension
				Case ".json"
					If My.Computer.FileSystem.FileExists(FileName) = False Then
						'check to read the file using an xml format
						'if it work save it under the correct format and continue
						ThisListInfo = Me.FileListReadLocal(Of T)(ThisFileNameXML, DataDefault, Nothing)

						Me.FileListSaveLocal(Of T)(FileName, ThisListInfo)
						Return ThisListInfo
					End If
					Try
						ThisFileData = File.ReadAllText(FileName)
						ThisListInfo = JsonConvert.DeserializeObject(Of List(Of T))(ThisFileData)
						Return ThisListInfo
					Catch ex As Exception
						'try to read as an xml file
						Try
							ThisListInfo = FileListRead(Of T)(ThisFileNameXML, DataDefault, Nothing)
							Return ThisListInfo
						Catch ex1 As Exception
							Throw ex1
						End Try
					End Try
				Case ".xml"
					If My.Computer.FileSystem.FileExists(FileName) = False Then
						Throw New FileNotFoundException
					End If
					Try
						ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
						'StreamReader is designed for character input in a particular encoding, whereas the Stream class
						'is designed for byte input and output. Use StreamReader for reading lines of information
						'From a standard text file.
						ThisTextReader = New StreamReader(FileName)
						ThisListInfo = CType(ThisXmlSerializer.Deserialize(ThisTextReader), List(Of T))
						ThisTextReader.Dispose()
						Return ThisListInfo
					Catch ex As Exception
						Throw ex
					End Try
				Case Else
					Throw New InvalidDataException($"File name extension is not supported:{vbCr}{FileName}")
			End Select
		End Function

		'Public Function FileListRead(Of T)(ByVal FileName As String) As List(Of T)
		'	Dim ThisListDefauly As List(Of T) = Nothing
		'	Dim ThisException As Exception = Nothing

		'	Return FileListRead(Of T)(FileName, ThisListDefauly, ThisException)
		'End Function

		Private Sub FileListSaveLocal(Of T As {ITreeNode(Of U)}, U)(ByVal FileName As String, ByRef HeaderInfo As List(Of T))
			Dim ThisXmlSerializer As XmlSerializer = Nothing
			Dim ThisException As Exception = Nothing
			Try
				ThisXmlSerializer = New XmlSerializer(GetType(List(Of T)))
			Catch ex As Exception
				ex = ex
			End Try

			Dim ThisTextWriter As System.IO.TextWriter

			'create the directory if it does not exist
			Dim ThisPath As String = System.IO.Path.GetDirectoryName(FileName)
			With My.Computer.FileSystem
				If .DirectoryExists(ThisPath) = False Then
					Try
						.CreateDirectory(ThisPath)
					Catch ex As Exception
						Throw New Exception(String.Format("Unable to create directory {0},ThisPath"), ex)
					End Try
				End If
			End With
			System.IO.File.Delete(FileName)
			ThisTextWriter = New StreamWriter(FileName)
			ThisXmlSerializer.Serialize(ThisTextWriter, HeaderInfo)
			ThisTextWriter.Dispose()
		End Sub
#End Region

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