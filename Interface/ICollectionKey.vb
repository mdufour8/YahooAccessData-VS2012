#Region "Imports"
Imports System
Imports System.IO
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices
'Imports System.Data.Entity
'Imports System.Data.Entity.Infrastructure
Imports System.Linq
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Net
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.Xml.Serialization
#End Region
#Region "DuplicateObject"
'Public Class Duplicate(Of T)
'	Implements IEqualityComparer(Of T)
'	Implements IEquatable(Of T)

'	Private ThisCopyFrom As T

'	Public Sub New()
'		ThisCopyFrom = Nothing
'	End Sub
'	Public Sub New(ByRef CopyFrom As T)
'		ThisCopyFrom = CopyFrom
'	End Sub

'	''' <summary>
'	''' Provide deep copy of the object property
'	''' </summary>
'	''' <returns></returns>
'	''' <remarks></remarks>
'	Public Function CloneDeep() As T
'		If ThisCopyFrom Is Nothing Then Return Nothing
'		If ThisCopyFrom.GetType.IsValueType OrElse ThisCopyFrom.GetType Is Type.GetType("System.String") Then
'			Return ThisCopyFrom
'		Else
'			Dim ThisCopyTo = Activator.CreateInstance(Of T)()
'			For Each Item As PropertyInfo In ThisCopyTo.GetType.GetProperties
'				'do not try to copy the class object in this version
'				With Item
'					If .CanWrite Then

'						If ((.PropertyType.IsClass = False) And
'							 (.PropertyType.IsInterface = False)) Or
'							 (.PropertyType.Name = "String") Then
'							Item.SetValue(ThisCopyTo, Item.GetValue(ThisCopyFrom, Nothing), Nothing)
'						End If
'					End If
'				End With
'			Next
'			Return ThisCopyTo
'		End If
'	End Function

'#Region "Equality Interface"
'	''' <summary>
'	''' Provide deep equality of the object property
'	''' </summary>
'	''' <param name="other"></param>
'	''' <returns></returns>
'	''' <remarks></remarks>
'	Public Overloads Function Equals(other As T) As Boolean Implements System.IEquatable(Of T).Equals

'		For Each Item As PropertyInfo In other.GetType.GetProperties
'			With Item
'				If .CanRead Then
'					If ((.PropertyType.IsClass = False) And
'							 (.PropertyType.IsInterface = False)) Or
'							 (.PropertyType.Name = "String") Then

'						If .GetValue(ThisCopyFrom, Nothing).Equals(.GetValue(other, Nothing)) = False Then Return False
'					End If
'				End If
'			End With
'		Next
'		Return True
'	End Function

'	''' <summary>
'	''' Provide deep equality of the object property
'	''' </summary>
'	''' <param name="obj"></param>
'	''' <returns></returns>
'	''' <remarks></remarks>
'	Public Overrides Function Equals(obj As Object) As Boolean
'		If TypeOf obj Is Report Then
'			Return Me.Equals(DirectCast(obj, T))
'		Else
'			Return False
'		End If
'	End Function

'	Public Overrides Function GetHashCode() As Integer
'		If ThisCopyFrom Is Nothing Then
'			Return Me.GetHashCode
'		Else
'			Return ThisCopyFrom.GetHashCode
'		End If
'	End Function
'	''' <summary>
'	''' Provide deep equality using all the objects property
'	''' </summary>
'	''' <param name="x">First object</param>
'	''' <param name="y">Second object to compare</param>
'	''' <returns>True or false</returns>
'	''' <remarks></remarks>
'	Public Overloads Function Equals(x As T, y As T) As Boolean Implements System.Collections.Generic.IEqualityComparer(Of T).Equals
'		Dim ThisCopyFromSave = ThisCopyFrom
'		Dim IsEquals As Boolean
'		ThisCopyFrom = x
'		IsEquals = Me.Equals(y)
'		ThisCopyFrom = ThisCopyFromSave
'		Return IsEquals
'	End Function

'	Public Overloads Function GetHashCode(obj As T) As Integer Implements System.Collections.Generic.IEqualityComparer(Of T).GetHashCode
'		If obj Is Nothing Then
'			Return Me.GetHashCode
'		Else
'			Return obj.GetHashCode
'		End If
'	End Function
'#End Region
'End Class
#End Region

#Region "ISearchKey"
  Public Interface ISearchKey(Of T, U)
    ''' <summary>
    ''' Quickly search for an object with the ID value
    ''' </summary>
    ''' <param name="ID"></param>
    ''' <returns>The object T</returns>
    ''' <remarks>The search speed is O(1)</remarks>
    Function Find(ByVal ID As Integer) As T
    ''' <summary>
    ''' Quickly search for an object with the key value
    ''' </summary>
    ''' <param name="KeyValue"></param>
    ''' <returns>The object T</returns>
    ''' <remarks>The search speed is O(1)</remarks>
    Function Find(ByVal KeyValue As U) As T
    ''' <summary>
    ''' Searches the entire sorted List(Of T) for an element using the default comparer
    ''' </summary>
    ''' <param name="item"></param>
    ''' <returns>returns the zero-based index of the element. </returns>
    ''' <remarks>
    ''' If the List(Of T)does not contain the specified value, the method returns a negative integer. 
    ''' You can apply the bitwise complement operation (~) to this negative integer to get the index 
    ''' of the first element that is larger than the search value. When inserting the value into 
    ''' the List(Of T), this index should be used as the insertion point to maintain the sort order. 
    ''' This method is an O(log n) operation, where n is the number of elements in the range. 
    ''' </remarks>
    Function Find(item As T) As Integer
    ''' <summary>
    ''' Searches the entire sorted List(Of T)for an element using the specified comparer
    ''' </summary>
    ''' <param name="item"></param>
    ''' <param name="comparer"></param>
    ''' <returns>returns the zero-based index of the element.</returns>
    ''' <remarks></remarks>
    Function Find(item As T, comparer As IComparer(Of T)) As Integer
    ''' <summary>
    ''' Searches a range of elements in the sorted List(Of T)for an element using the specified comparer. 
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="count"></param>
    ''' <param name="item"></param>
    ''' <param name="comparer"></param>
    ''' <returns>returns the zero-based index of the element.</returns>
    ''' <remarks></remarks>
    Function Find(index As Integer, count As Integer, item As T, comparer As IComparer(Of T)) As Integer
    ' ''' <summary>
    ' ''' item is the element where the key value and ID relation may have change
    ' ''' </summary>
    ' ''' <param name="item"></param>
    ' ''' <remarks>This function should be call to update the relation between the ID and Key of the object</remarks>
    'Sub Refresh(item As T)
  End Interface
#End Region
#Region "IProcessTimeMeasurement"
Public Interface IProcessTimeMeasurement
  ReadOnly Property Key As String
  Property Name As String
  Property ElapsedMilliseconds As Long
End Interface
#End Region
#Region "IMemoryStream"
Public Interface IMemoryStream
  Enum enuFileType As Integer
    Standard
    RecordIndexed
    'RecordIndexedEndOfDay
  End Enum
  Sub SerializeSaveTo(ByRef Stream As Stream, ByVal FileType As enuFileType)
  Sub SerializeSaveTo(ByRef Stream As Stream)
  Function SerializeSaveTo() As Byte()
  Sub SerializeLoadFrom(ByRef Stream As Stream)
  Sub SerializeLoadFrom(ByRef Stream As Stream, ByVal IsRecordVirtual As Boolean)
  Sub SerializeLoadFrom(ByRef Data As Byte())
End Interface
#End Region
#Region "IDataPointer"
	Public Interface IDataPosition
		Property Current As Long
		Property ToNext As Long
		Property ToPrevious As Long
	End Interface
#End Region
#Region "IDataVirtual"
	Public Interface IDataVirtual
		Property Enabled As Boolean
		ReadOnly Property IsLoaded As Boolean
		Sub Load()
		Sub Release()
		Property Count As Integer
	End Interface
#End Region
#Region "IRecordInfo"
	Public Interface IRecordInfo
		Property CountTotal As Integer
		Property MaximumID As Integer
	End Interface
#End Region
#Region "IListCopy"
  Public Interface IListCopy(Of T)
    Function Copy() As ICollection(Of T)
  End Interface
#End Region
#Region "ReadVirtual"
	'not completed yet
	'Friend Class ReadVirtual(Of T As {IRegisterKey(Of U), IDateUpdate, IMemoryStream}, U)
	'	Private MyStreamBase As Stream
	'	Private MyDateRange As IDateRange
	'	Private MyChildPath As String
	'	Private MyKeyName As String
	'	Private MyFileExtension As String

	'	Public Sub New(ByRef StreamBase As Stream, ByRef DateRange As IDateRange, ByVal ChildPath As String, ByVal KeyName As String, ByVal FileExtension As String)
	'		MyStreamBase = StreamBase
	'		MyDateRange = DateRange
	'		MyChildPath = ChildPath
	'		MyKeyName = KeyName
	'		MyFileExtension = FileExtension
	'	End Sub

	'	Public Property Count As Integer Implements IDataVirtual.Count
	'		Get

	'		End Get
	'		Set(value As Integer)

	'		End Set
	'	End Property

	'	Public Property Enabled As Boolean Implements IDataVirtual.Enabled
	'		Get

	'		End Get
	'		Set(value As Boolean)

	'		End Set
	'	End Property

	'	Public ReadOnly Property IsLoaded As Boolean Implements IDataVirtual.IsLoaded
	'		Get

	'		End Get
	'	End Property

	'	Public Sub Load() Implements IDataVirtual.Load

	'End Sub


	'	End Sub

	'	Public Sub Release() Implements IDataVirtual.Release

	'	End Sub
	'End Class
#End Region
#Region "IRegisterKey"
  Public Interface IRegisterKey(Of T)
    Property KeyID As Integer
    Property KeyValue As T
  End Interface
#End Region
#Region "IFormatData"
	Public Interface IFormatData
		Function ToStingOfData() As String()
		Function ToListOfHeader() As List(Of HeaderInfo)
	End Interface

  Public Interface IHeaderInfo
    Property Title As String
    Property Name As String
    Property Format As String
    Property Width As Integer
    Property IsSortable As Boolean
    Property IsSortEnabled As Boolean
    Property Visible As Boolean
  End Interface

  Public Interface IHeaderInfoPlus
    Property Title As String
    Property Name As String
    Property Format As String
    Property Width As Integer
    Property IsSortable As Boolean
    Property IsSortEnabled As Boolean
    Property Visible As Boolean
    Function HasValue(ByVal Tablename As String) As Boolean
  End Interface

  'Public Interface IHeaderInfoPlus(Of T)
  '  Property Title As String
  '  Property Name As String
  '  Property Format As String
  '  Property Width As Integer
  '  Property IsSortable As Boolean
  '  Property IsSortEnabled As Boolean
  '  Property Visible As Boolean
  '  Function HasValue(x As T, ByVal Tablename As String) As Boolean
  'End Interface

  <Serializable()>
  Public Class HeaderInfo
    Implements IHeaderInfo

    Public Sub New()
      'by default
      Me.IsSortEnabled = True
    End Sub

    Public Property Format As String Implements IHeaderInfo.Format
    Public Property IsSortable As Boolean Implements IHeaderInfo.IsSortable
    Public Property IsSortEnabled As Boolean Implements IHeaderInfo.IsSortEnabled
    Public Property Name As String Implements IHeaderInfo.Name
    Public Property Title As String Implements IHeaderInfo.Title
    Public Property Visible As Boolean Implements IHeaderInfo.Visible
    Public Property Width As Integer Implements IHeaderInfo.Width
  End Class

  Public Interface IHeaderListInfo
    ReadOnly Property HeaderInfo As List(Of HeaderInfo)
    Function HasValue(ByVal Tablename As String) As Boolean
  End Interface

'  'Class HeaderInfoPlus(Of T)
#End Region  'IFormatData
#Region "ISort"
  Public Enum enuSortType
    None
    Ascending
    Descending
  End Enum
  Public Interface ISort(Of T)
    ''' <summary>
    ''' This method use the QuickSort algorithm
    ''' </summary>
    ''' <remarks>
    ''' On average, this method is an O( n log n) operation, where n is Count; 
    ''' in the worst case it is an O( n ^ 2) operation. 
    ''' </remarks>
    Sub Sort()
    Sub Sort(ByVal KeyName As String, ByVal SortType As enuSortType)
    Sub Sort(ByVal KeyPrimaryName As String, ByVal PrimarySortType As enuSortType, ByVal KeySecondaryName As String, ByVal SecondarySortType As enuSortType)
    ''' <summary>
    ''' Sorts the elements in the entire List(Of T)using the specified System.Comparison(Of T). 
    ''' </summary>
    ''' <param name="comparison"></param>
    ''' <remarks></remarks>
    Sub Sort(comparison As Comparison(Of T))
    ''' <summary>
    ''' Sorts the elements in the entire List(Of T)using the specified comparer. 
    ''' </summary>
    ''' <param name="comparer"></param>
    ''' <remarks></remarks>
    Sub Sort(comparer As IComparer(Of T))
    ''' <summary>
    ''' Sorts the elements in a range of elements in List(Of T)using the specified comparer.
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="count"></param>
    ''' <param name="comparer"></param>
    ''' <remarks></remarks>
    Sub Sort(index As Integer, count As Integer, comparer As IComparer(Of T))
    ''' <summary>
    ''' Clear the current sort
    ''' </summary>
    ''' <remarks></remarks>
    Sub Clear()
    ''' <summary>
    ''' Reverses the order of the elements in the entire List(Of T). 
    ''' </summary>
    ''' <remarks>This method is an O( n) operation</remarks>
    Sub Reverse()
    ''' <summary>
    ''' Reverses the order of the elements in the specified range.
    ''' </summary>
    ''' <param name="index"></param>
    ''' <param name="count"></param>
    ''' <remarks></remarks>
    Sub Reverse(index As Integer, count As Integer)
  End Interface
#End Region
#Region "ICompareData"
  Public Interface ICompareData(Of T)
    Function Comparer(ByVal TableName As String) As System.Comparison(Of T)
  End Interface
#End Region
#Region "ISystemInternal"
  Public Interface ISystemInternal(Of T)
    ReadOnly Property ToList() As List(Of T)
  End Interface
#End Region
#Region "ISystemEvent"
  Public Interface ISystemEvent(Of T)
    Sub Add(item As T)
    Function Remove(item As T) As Boolean
    Sub Clear()
    Sub Load()
  End Interface
#End Region
#Region "ISystemEventRegister"
  Public Interface ISystemEventRegister(Of T)
    Property SinkEvent As ISystemEvent(Of T)
  End Interface
#End Region
#Region "IDateRange"
  Public Interface IDateRange
    Property DateStart As Date
    Property DateStop As Date
    Function NumberDays() As Integer
    Sub MovePrevious()
    Sub MoveNext()
    Sub MoveBegin()
    Sub MoveLast()
    Sub Refresh()
    Sub Refresh(ByVal DateStart As Date, ByVal DateStop As Date)
    Sub Refresh(ByVal NumberDays As Integer, ByVal DateStop As Date)
    Sub Refresh(ByVal DateStart As Date, ByVal NumberDays As Integer)
  End Interface
#End Region
#Region "IDateUpdate"
  Public Interface IDateUpdate
    Property DateStart As Date
    Property DateStop As Date
    ReadOnly Property DateUpdate As Date
    ReadOnly Property DateLastTrade As Date
    ReadOnly Property DateDay As Date
  End Interface
#End Region
#Region "IDateTrade"
  Public Interface IDateTrade
    Property DateStart As Date
    Property DateStop As Date
  End Interface
#End Region
#Region "IStockRecord"
Public Enum enuStockRecordLoadType
  FixedCount
  WebEodHistorical
End Enum
Public Interface IStockRecordEvent
  Sub LoadBefore(ByVal Symbol As String)
  Sub LoadAfter(ByVal Symbol As String)
  Sub LoadCancel(ByVal Symbol As String)
  ReadOnly Property IsLoadCancel As Boolean
  ReadOnly Property IsLoading As Boolean
  ReadOnly Property SymbolLoading As String
End Interface
Public Interface IRecordControlInfo
  Property Enabled As Boolean
  Property Count As Integer
  Property ControlType As enuStockRecordLoadType
  Sub LoadToCache(ByVal StockList As IEnumerable(Of String))
  Sub LoadToCache(ByVal StockList As IEnumerable(Of YahooAccessData.Stock))
  Sub LoadToCache(ByVal StockSymbol As String)
  Sub LoadToCache(ByVal Stock As YahooAccessData.Stock)
End Interface
#End Region
#Region "Message Event Interface Definition"
Public Interface IMessageEvents
  Event Message(ByVal Message As String)
  Event MessageError(ByVal MessageError As String)
End Interface
Public Interface IMessageInfoEvents
  ''' <summary>
  ''' Error icon. The user interface (UI) is presenting an error or problem that has occurred.
  ''' Warning icon. The UI is presenting a condition that might cause a problem in the future.
  ''' Information icon. The UI is presenting useful information.
  ''' Question mark icon. The UI indicates a Help entry point.
  ''' </summary>
  Enum EnuMessageType
    Information
    Warning
    Question
    InError
    InformationUpdate
    Debug
  End Enum

  Event Message(ByVal Message As String, MessageType As EnuMessageType)
End Interface
#End Region





