  ''' <summary>
  ''' A thread safe linked Hashset of object T and searchable with a key of type U. 
  ''' </summary>
  ''' <typeparam name="T"></typeparam>
  '''	The object type
  ''' <typeparam name="U"></typeparam>
  '''	The key of type U use to search for the object
  ''' <remarks>
  ''' The interface ISearchKey(Of T, U) is implemented for searching or sort on T for the key U or an integer ID.
  ''' The object T need to implement the IRegisterKey(Of U) to provide an automatic access to the ID and key. The 
  ''' object will automatically update the ID if the ID is left at zero when adding the element to the collection.
  ''' The standard generic interface IComparable(Of In T) need to be implemented in the object 
  ''' T to support the ISearchKey(T, U).Sort method.
  '''</remarks>
<Serializable()>
Public Class LinkedHashSet(Of T As {New, Class, IRegisterKey(Of U),  IDateUpdate, IFormatData}, U)
  Implements ICollection(Of T)
  Implements ISearchKey(Of T, U)
  Implements IRegisterKey(Of String)
  Implements ISort(Of T)
  Implements ISystemInternal(Of T)
  Implements ISystemEventRegister(Of T)
  Implements IDataVirtual
  Implements IRecordInfo
  Implements IEnumerable(Of T)
  Implements IDateUpdate
  Implements IDateTrade
  Implements IListCopy(Of T)
  Implements IHeaderListInfo

#Region "Definition"
  Private MyList As List(Of T)
  Private MyIDMax As Integer
  Private MyIDMaxForZero As Integer
  Private MyDictionaryOfID As Dictionary(Of Integer, T)
  Private MyDictionaryOfKey As Dictionary(Of U, T)


  'use to make the collection thread safe
  Private ThisLock As New Object
  Private MySinkEvent As ISystemEvent(Of T)
  Private MyKeyID As Integer
  Private MyCountTotal As Integer
  Private MyKey As String = System.Guid.NewGuid.ToString()
  Private MyCountMax As Integer
  Private IsVirtualEnabled As Boolean
  Private IsVirtualLoaded As Boolean
  Private MyVirtualFirstT As T = Nothing
  Private Shared MyMaxKeyID As Integer
  Private MyDateUpdateStart As Date?
  Private MyDateUpdateStop As Date?
  Private MyDateTradeStart As Date?
  Private MyDateTradeStop As Date?
#End Region
#Region "New"
  Public Sub New()
    Me.New(False, 0, 4)
  End Sub

  Public Sub New(ByVal IsVirtualDataEnabled As Boolean)
    Me.New(IsVirtualDataEnabled, 0, 4)
  End Sub

  Public Sub New(ByVal IsVirtualDataEnabled As Boolean, ByVal IDMax As Integer)
    Me.New(IsVirtualDataEnabled, IDMax, 4)
  End Sub

  Public Sub New(ByVal Data As IEnumerable(Of T))
    Me.New()
    For Each ThiItem In Data
      Me.Add(ThiItem)
    Next
  End Sub

  Public Sub New(ByVal IsVirtualDataEnabled As Boolean, ByVal IDMax As Integer, ByVal Capacity As Integer)
    MyIDMax = IDMax
    MyIDMaxForZero = MyIDMax
    IsVirtualEnabled = IsVirtualDataEnabled
    IsVirtualLoaded = False
    KeyInit()
    MyList = New List(Of T)
    MyDictionaryOfID = New Dictionary(Of Integer, T)
    MyDictionaryOfKey = New Dictionary(Of U, T)
  End Sub

  Public Sub New(ByRef SinkEvent As ISystemEvent(Of T), ByVal IDMax As Integer, ByVal Capacity As Integer)
    If SinkEvent Is Nothing Then
      IsVirtualEnabled = False
    Else
      IsVirtualEnabled = True
      Me.SinkEvent = SinkEvent
    End If
    IsVirtualLoaded = False
    KeyInit()
    MyList = New List(Of T)
    MyDictionaryOfID = New Dictionary(Of Integer, T)
    MyDictionaryOfKey = New Dictionary(Of U, T)
  End Sub

  Public Sub New(ByRef SinkEvent As ISystemEvent(Of T), ByVal IDMax As Integer)
    Me.New(SinkEvent, IDMax, 4)
  End Sub

  Public Sub New(ByRef SinkEvent As ISystemEvent(Of T))
    Me.New(SinkEvent, 0, 4)
  End Sub

  Private Sub KeyInit()
    SyncLock ThisLock
      If MyMaxKeyID = Integer.MaxValue Then
        Throw New Exception(String.Format("Maximum unique KeyID reached for object {0}", Me.ToString))
      Else
        MyMaxKeyID = MyMaxKeyID + 1
      End If
      MyKeyID = MyMaxKeyID
      MyKey = MyKeyID.ToString
    End SyncLock
  End Sub
#End Region
#Region "Main"
  Public Overrides Function ToString() As String
    Return String.Format("{0},ID:{1},Key:{2},Record:{3} of {4}", TypeName(Me), Me.KeyID, Me.KeyValue.ToString, Me.Count, Me.CountTotal)
  End Function

  Public Property Capacity As Integer
    Get
      Return MyList.Capacity
    End Get
    Set(value As Integer)
      'ignore the capacity
      'MyList.Capacity = value
      'If MyDictionaryOfID.Count = 0 Then
      '  'there is no capacity property with the dictionary
      '  'just create a new object if there is no element 
      '  MyDictionaryOfID = New Dictionary(Of Integer, T)(Capacity)
      '  MyDictionaryOfKey = New Dictionary(Of U, T)(Capacity)
      'End If
    End Set
  End Property

  Public Function Copy() As ICollection(Of T) Implements IListCopy(Of T).Copy
    Dim ThisLinkedHashSet As New LinkedHashSet(Of T, U)
    For Each ThisT In Me
      ThisLinkedHashSet.Add(ThisT)
    Next
    Return ThisLinkedHashSet
  End Function
#End Region
#Region "IHeaderListInfo"
  Public Function HasValue(ColoumnName As String) As Boolean Implements IHeaderListInfo.HasValue
    Dim ThisValue As Object
    Dim ThisT As T
    Dim ThisDateNull As Date = ReportDate.DateNullValue

    If MyList.Count > 0 Then
      Dim ThisType As Type = GetType(T)
      Dim ThisProperty As System.Reflection.PropertyInfo = ThisType.GetProperty(ColoumnName)
      If ThisProperty Is Nothing Then Return False
      Try
        Select Case ThisProperty.PropertyType
          Case GetType([Enum])
            Return True
          Case GetType(Boolean)
            Return True
          Case GetType(Integer)
            For Each ThisT In MyList
              ThisValue = ThisProperty.GetValue(ThisT, Nothing)
              If CType(ThisValue, Integer) <> 0 Then Return True
            Next
          Case GetType(Single)
            For Each ThisT In MyList
              ThisValue = ThisProperty.GetValue(ThisT, Nothing)
              If CType(ThisValue, Single) <> 0 Then Return True
            Next
          Case GetType(Double)
            For Each ThisT In MyList
              ThisValue = ThisProperty.GetValue(ThisT, Nothing)
              If CType(ThisValue, Double) <> 0 Then Return True
            Next
          Case GetType(String)
            For Each ThisT In MyList
              ThisValue = ThisProperty.GetValue(ThisT, Nothing)
              If Len(CType(ThisValue, String)) > 0 Then Return True
            Next
          Case GetType(Date)
            For Each ThisT In MyList
              ThisValue = ThisProperty.GetValue(ThisT, Nothing)
              If CType(ThisValue, Date) <> ThisDateNull Then Return True
            Next
        End Select
        Return False
      Catch ex As Exception
        Debug.Assert(False)
        Return True
      End Try
    Else
      Return False
    End If
  End Function

  Public ReadOnly Property HeaderInfo As List(Of HeaderInfo) Implements IHeaderListInfo.HeaderInfo
    Get
      Dim ThisT As T
      If MyList.Count > 0 Then
        ThisT = MyList.First
      Else
        ThisT = ExtensionService.CreateInstance(Of T)()
      End If
      Return ThisT.ToListOfHeader
    End Get
  End Property
#End Region  'IHeaderListInfo
#Region "IRecordInfo"
  Public Property CountTotal As Integer Implements IRecordInfo.CountTotal
    Get
      Return MyCountTotal
    End Get
    Set(value As Integer)
      MyCountTotal = value
    End Set
  End Property

  Public Property MaximumID As Integer Implements IRecordInfo.MaximumID
    Get
      Return MyIDMax
    End Get
    Set(value As Integer)
      'update the MaxKeyID only if it is greater than the actual value
      If value > MyIDMax Then
        MyIDMax = value
        'do not initiate loading from here by reading the count directly 
        'from the list and not from the properties
        If MyList.Count = 0 Then
          MyIDMaxForZero = MyIDMax
        End If
      End If
    End Set
  End Property
#End Region  'IRecordInfo
#Region "ISystemInternal"
  ''' <summary>
  ''' Give access to the internal list for added flexibility
  ''' </summary>
  ''' <value></value>
  ''' <returns>The internal list</returns>
  ''' <remarks>
  ''' Use at your own risk. Any of the function called on the list should not change the list data
  ''' in a significant way, or the internal key will become out of sync
  ''' </remarks>
  Public ReadOnly Property ToList As List(Of T) Implements ISystemInternal(Of T).ToList
    Get
      SyncLock ThisLock
        Return MyList
      End SyncLock
    End Get
  End Property

  ''' <summary>
  ''' Use by an external thread to lock the LinkedHashset thread using SyncLock
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks>can be use with SyncLock to lock the list for multiple operation</remarks>
  Public ReadOnly Property Lock As Object
    Get
      Return ThisLock
    End Get
  End Property

  Public Property SinkEvent As ISystemEvent(Of T) Implements ISystemEventRegister(Of T).SinkEvent
    Get
      Return MySinkEvent
    End Get
    Set(value As ISystemEvent(Of T))
      MySinkEvent = value
    End Set
  End Property
#End Region
#Region "ICollection"
  ''' <summary>
  ''' This function will try to add the the item to the collection. 
  ''' If it fail the function return false otherwise true for success
  ''' The failure to add may indicate a duplicated key in the list
  ''' </summary>
  ''' <param name="item"></param>
  ''' <returns>True for success</returns>
  Public Function TryAdd(item As T) As Boolean
    SyncLock ThisLock
      If MyDictionaryOfKey.ContainsKey(item.KeyValue) Then
        'item already in collection 
        Return False
      End If
    End SyncLock
    Me.Add(item)
    Return True
  End Function

  Public Sub Add(item As T) Implements System.Collections.Generic.ICollection(Of T).Add
    SyncLock ThisLock
      'Dim ThisIdentity As IRegisterKey(Of U) = Nothing
      'If TypeOf item Is IRegisterKey(Of U) Then
      '	ThisIdentity = TryCast(item, IRegisterKey(Of U))
      'End If
      'If ThisIdentity Is Nothing Then
      '	Throw New Exception(String.Format("Error: Interface IRegisterKey is not supported for data object {0}...", item.ToString))
      'End If
      If MyDictionaryOfKey.ContainsKey(item.KeyValue) Then
        'item already in collection 
        'Throw New Exception(String.Format("Error: Multiple ID key value {0} in object {1} ...", item.KeyValue, item.ToString))
        Exit Sub
      End If
      If item.KeyID = 0 Then
        'new object need a new key
        MyIDMax = MyIDMax + 1
        item.KeyID = MyIDMax
      Else
        If item.KeyID > MyIDMax Then
          MyIDMax = item.KeyID
        End If
      End If
      If MyDictionaryOfID.ContainsKey(item.KeyID) Then
        'This should not happen but try to assign a new ID to the object
        'and keep going
        MyIDMax = MyIDMax + 1
        item.KeyID = MyIDMax
      End If
      MyList.Add(item)
      MyDictionaryOfKey.Add(item.KeyValue, item)
      MyDictionaryOfID.Add(item.KeyID, item)

      If MySinkEvent IsNot Nothing Then
        MySinkEvent.Add(item)
      End If
      If Me.CountTotal < Me.Count Then
        Me.CountTotal = Me.Count
      End If
      If MyMaxKeyID < Me.Count Then
        MyMaxKeyID = Me.Count
      End If
    End SyncLock
  End Sub

  Public Sub Clear() Implements System.Collections.Generic.ICollection(Of T).Clear
    SyncLock ThisLock
      MyList.Clear()
      MyDictionaryOfKey.Clear()
      MyDictionaryOfID.Clear()
      'to make sure that the object inside the list are recycle 
      'it is mecessary to create the new object again
      'otherwise it has been demonstrated that the system memory is slowly leaking
      MyList = New List(Of T)
      MyDictionaryOfID = New Dictionary(Of Integer, T)
      MyDictionaryOfKey = New Dictionary(Of U, T)
      MyIDMax = MyIDMaxForZero
      If MySinkEvent IsNot Nothing Then
        MySinkEvent.Clear()
      End If
      IsVirtualLoaded = False
    End SyncLock
  End Sub

  Public Function Contains(item As T) As Boolean Implements System.Collections.Generic.ICollection(Of T).Contains
    SyncLock ThisLock
      'Dim ThisIdentity As IRegisterKey(Of U) = Nothing
      'If TypeOf item Is IRegisterKey(Of U) Then
      '	ThisIdentity = TryCast(item, IRegisterKey(Of U))
      'End If
      'If ThisIdentity Is Nothing Then
      '	Return False
      'End If
      Return MyDictionaryOfKey.ContainsKey(item.KeyValue)
    End SyncLock
  End Function

  Public Sub CopyTo(array() As T, arrayIndex As Integer) Implements System.Collections.Generic.ICollection(Of T).CopyTo
    SyncLock ThisLock
      MyList.CopyTo(array, arrayIndex)
    End SyncLock
  End Sub

  Public ReadOnly Property Count As Integer Implements System.Collections.Generic.ICollection(Of T).Count
    Get
      SyncLock ThisLock
        Return MyList.Count
      End SyncLock
    End Get
  End Property

  Public ReadOnly Property IsReadOnly As Boolean Implements System.Collections.Generic.ICollection(Of T).IsReadOnly
    Get
      Return False
    End Get
  End Property

  Public Function AsSearchKey() As ISearchKey(Of T, U)
    Return Me
  End Function

  Public Function ContainKey(ByVal KeyValue As U) As Boolean
    Return MyDictionaryOfKey.ContainsKey(KeyValue)
  End Function

  ''' <summary>
  ''' the ICollection does not support this removal by index
  ''' </summary>
  ''' <param name="index"></param>
  ''' <returns>true if successful</returns>
  Public Function RemoveAt(ByVal index As Integer) As Boolean
    'use directly the MyList for item removal the ICollection
    SyncLock ThisLock
      If index >= 0 And index < Me.Count Then
        'When you call RemoveAt to remove an item, the remaining items in the list are
        'renumbered to replace the removed item. For example, if you remove the item
        'at index 3, the item at index 4 Is moved to the 3 position.
        'In addition, the number of items in the list (as represented by the Count property is
        'reduced by 1.
        'This method Is an O(n) operation, where n Is (Count - index).
        'i.e. removing the top item is very fast
        Dim ThisItemToRemove = MyList(index)
        MyList.RemoveAt(index)
        MyDictionaryOfID.Remove(ThisItemToRemove.KeyID)
        MyDictionaryOfKey.Remove(ThisItemToRemove.KeyValue)
        If MySinkEvent IsNot Nothing Then
          MySinkEvent.Remove(ThisItemToRemove)
        End If
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
  Public Function Remove(item As T) As Boolean Implements System.Collections.Generic.ICollection(Of T).Remove
    SyncLock ThisLock
      If MyList.Remove(item) = True Then
        'Dim ThisIdentity = TryCast(item, IRegisterKey(Of U))
        MyDictionaryOfID.Remove(item.KeyID)
        MyDictionaryOfKey.Remove(item.KeyValue)
        If MySinkEvent IsNot Nothing Then
          MySinkEvent.Remove(item)
        End If
        'Todo: datestop should also be updated because it would change if the last record is changed
        Return True
      Else
        Return False
      End If
    End SyncLock
  End Function
#End Region
#Region "Enumerator"
  Public Function GetEnumerator() As System.Collections.Generic.IEnumerator(Of T) Implements System.Collections.Generic.IEnumerable(Of T).GetEnumerator
    SyncLock ThisLock
      Return MyList.GetEnumerator
    End SyncLock
  End Function

  Private Function IEnumerator_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
    SyncLock ThisLock
      Return Me.GetEnumerator
    End SyncLock
  End Function
#End Region
#Region "ISearchKey Implementation"
  Public Function Find(ID As Integer) As T Implements ISearchKey(Of T, U).Find
    SyncLock ThisLock
      If MyDictionaryOfID.ContainsKey(ID) Then
        Return MyDictionaryOfID(ID)
      Else
        Return Nothing
      End If
    End SyncLock
  End Function

  Public Function Find(Value As U) As T Implements ISearchKey(Of T, U).Find
    SyncLock ThisLock
      If Value Is Nothing Then Return Nothing
      If MyDictionaryOfKey.ContainsKey(Value) Then
        Return MyDictionaryOfKey(Value)
      Else
        Return Nothing
      End If
    End SyncLock
  End Function

  Public Function Find(index As Integer, count As Integer, item As T, comparer As System.Collections.Generic.IComparer(Of T)) As Integer Implements ISearchKey(Of T, U).Find
    SyncLock ThisLock
      Return MyList.BinarySearch(index, count, item, comparer)
    End SyncLock
  End Function

  Public Function Find(item As T) As Integer Implements ISearchKey(Of T, U).Find
    SyncLock ThisLock
      Return MyList.BinarySearch(item)
    End SyncLock
  End Function

  Public Function Find(item As T, comparer As System.Collections.Generic.IComparer(Of T)) As Integer Implements ISearchKey(Of T, U).Find
    SyncLock ThisLock
      Return MyList.BinarySearch(item, comparer)
    End SyncLock
  End Function
#End Region
#Region "ISort Implementation"
#Region "ISort Implementation Definition"
  Private MyComparePrimaryKey As System.Comparison(Of T)
  Private MyCompareSecondaryKey As System.Comparison(Of T)

  Private Function SortSelect(
    ByRef CompareForPrimaryKey As System.Comparison(Of T),
    ByVal SortTypeForPrimaryKey As enuSortType,
    ByRef CompareForSecondaryKey As System.Comparison(Of T),
    ByVal SortTypeForSecondaryKey As enuSortType) As System.Comparison(Of T)

    If CompareForSecondaryKey Is Nothing Then
      Return SortSelect(CompareForPrimaryKey, SortTypeForPrimaryKey)
    Else
      If CompareForPrimaryKey Is Nothing Then
        Return SortSelect(CompareForSecondaryKey, SortTypeForSecondaryKey)
      End If
      MyComparePrimaryKey = CompareForPrimaryKey
      MyCompareSecondaryKey = CompareForSecondaryKey
      Select Case SortTypeForPrimaryKey
        Case enuSortType.Ascending
          Select Case SortTypeForSecondaryKey
            Case enuSortType.Ascending
              Return AddressOf ComparePrimarySecondaryKeyAscendingAscending
            Case enuSortType.Descending
              Return AddressOf ComparePrimarySecondaryKeyAscendingDescending
            Case enuSortType.None
              Return AddressOf ComparePrimaryKeyNone
          End Select
        Case enuSortType.Descending
          Select Case SortTypeForSecondaryKey
            Case enuSortType.Ascending
              Return AddressOf ComparePrimarySecondaryKeyDescendingAscending
            Case enuSortType.Descending
              Return AddressOf ComparePrimarySecondaryKeyDescendingDescending
            Case enuSortType.None
              Return AddressOf ComparePrimaryKeyNone
          End Select
        Case enuSortType.None
          Return AddressOf ComparePrimaryKeyNone
      End Select
    End If
    MyComparePrimaryKey = Nothing
    MyCompareSecondaryKey = Nothing
    Return Nothing
  End Function

  Private Function SortSelect(
    ByRef CompareForPrimaryKey As System.Comparison(Of T),
    ByVal SortTypeForPrimaryKey As enuSortType) As System.Comparison(Of T)

    MyComparePrimaryKey = CompareForPrimaryKey
    MyCompareSecondaryKey = Nothing
    Select Case SortTypeForPrimaryKey
      Case enuSortType.Ascending
        Return AddressOf ComparePrimaryKeyAscending
      Case enuSortType.Descending
        Return AddressOf ComparePrimaryKeyDescending
      Case enuSortType.None
        Return AddressOf ComparePrimaryKeyNone
    End Select
    MyComparePrimaryKey = Nothing
    Return Nothing
  End Function

  Private Function ComparePrimaryKeyAscending(x As T, y As T) As Integer
    Return MyComparePrimaryKey(x, y)
  End Function
  Private Function ComparePrimaryKeyDescending(x As T, y As T) As Integer
    Return -MyComparePrimaryKey(x, y)
  End Function
  Private Function ComparePrimaryKeyNone(x As T, y As T) As Integer
    Return 0
  End Function
  Private Function ComparePrimarySecondaryKeyDescendingDescending(x As T, y As T) As Integer
    Dim ThisResult As Integer = -MyCompareSecondaryKey(x, y)
    If ThisResult = 0 Then
      ThisResult = -MyComparePrimaryKey(x, y)
    End If
    Return ThisResult
  End Function
  Private Function ComparePrimarySecondaryKeyAscendingAscending(x As T, y As T) As Integer
    Dim ThisResult As Integer = MyCompareSecondaryKey(x, y)
    If ThisResult = 0 Then
      ThisResult = MyComparePrimaryKey(x, y)
    End If
    Return ThisResult
  End Function
  Private Function ComparePrimarySecondaryKeyDescendingAscending(x As T, y As T) As Integer
    Dim ThisResult As Integer = MyCompareSecondaryKey(x, y)
    If ThisResult = 0 Then
      ThisResult = -MyComparePrimaryKey(x, y)
    End If
    Return ThisResult
  End Function
  Private Function ComparePrimarySecondaryKeyAscendingDescending(x As T, y As T) As Integer
    Dim ThisResult As Integer = -MyCompareSecondaryKey(x, y)
    If ThisResult = 0 Then
      ThisResult = MyComparePrimaryKey(x, y)
    End If
    Return ThisResult
  End Function
#End Region
  Private Sub ISort_Sort() Implements ISort(Of T).Sort
    SyncLock ThisLock
      MyList.Sort()
    End SyncLock
  End Sub

  Private Sub ISort_Sort(index As Integer, count As Integer, comparer As System.Collections.Generic.IComparer(Of T)) Implements ISort(Of T).Sort
    SyncLock ThisLock
      MyList.Sort(index, count, comparer)
    End SyncLock
  End Sub

  Private Sub ISort_Sort(comparer As System.Collections.Generic.IComparer(Of T)) Implements ISort(Of T).Sort
    SyncLock ThisLock
      MyList.Sort(comparer)
    End SyncLock
  End Sub

  Private Sub ISort_Sort(comparison As System.Comparison(Of T)) Implements ISort(Of T).Sort
    SyncLock ThisLock
      If comparison IsNot Nothing Then
        MyList.Sort(comparison)
      End If
    End SyncLock
  End Sub

  Private Sub ISort_Clear() Implements ISort(Of T).Clear
    SyncLock ThisLock
      MyList.Sort(AddressOf ComparePrimaryKeyNone)
    End SyncLock
  End Sub

  Private Sub ISort_Sort(KeyName As String, SortType As enuSortType) Implements ISort(Of T).Sort
    SyncLock ThisLock
      If Me.Count > 0 Then
        Dim MyComparerT = New CompareByName(Of T)(KeyName)
        Me.ISort_Sort(SortSelect(AddressOf MyComparerT.Compare, SortType))
      End If
    End SyncLock
  End Sub

  Private Sub ISort_Sort(KeyPrimaryName As String, PrimarySortType As enuSortType, KeySecondaryName As String, SecondarySortType As enuSortType) Implements ISort(Of T).Sort
    SyncLock ThisLock
      If Me.Count > 0 Then
        Dim MyComparerPrimaryOfT = New CompareByName(Of T)(KeyPrimaryName)
        Dim MyComparerSecondaryOfT = New CompareByName(Of T)(KeySecondaryName)
        Me.ISort_Sort(SortSelect(AddressOf MyComparerPrimaryOfT.Compare, PrimarySortType, AddressOf MyComparerSecondaryOfT.Compare, SecondarySortType))
      End If
    End SyncLock
  End Sub

  Private Sub Reverse() Implements ISort(Of T).Reverse
    SyncLock ThisLock
      MyList.Reverse()
    End SyncLock
  End Sub

  Private Sub Reverse(index As Integer, count As Integer) Implements ISort(Of T).Reverse
    SyncLock ThisLock
      MyList.Reverse(index, count)
    End SyncLock
  End Sub
#End Region
#Region "IRegisterKey"
  Public Property KeyID As Integer Implements IRegisterKey(Of String).KeyID
    Get
      Return MyKeyID
    End Get
    Set(value As Integer)
      MyKeyID = value
    End Set
  End Property

  Public Property KeyValue As String Implements IRegisterKey(Of String).KeyValue
    Get
      Return MyKey
    End Get
    Set(value As String)
      MyKey = value
    End Set
  End Property
#End Region
#Region "IDataVirtual"
  Public Property IDataVirtual_Enabled As Boolean Implements IDataVirtual.Enabled
    Get
      Return IsVirtualEnabled
    End Get
    Set(value As Boolean)
      'IsVirtualEnabled = False
      'Return
      If IsVirtualEnabled = False Then
        If value = True Then
          IsVirtualEnabled = True
          Me.Clear()
        End If
      Else
        If value = False Then
          IDataVirtual_Load()
          IsVirtualEnabled = False
          IsVirtualLoaded = False
        End If
      End If
    End Set
  End Property

  Private ReadOnly Property IDataVirtual_IsLoaded As Boolean Implements IDataVirtual.IsLoaded
    Get
      If IsVirtualEnabled Then
        Return IsVirtualLoaded
      Else
        'always loaded if the virtual mode is not enabled
        Return True
      End If
    End Get
  End Property

  Private Sub IDataVirtual_Load() Implements IDataVirtual.Load
    If IsVirtualEnabled Then
      If IsVirtualLoaded = False Then
        If MySinkEvent IsNot Nothing Then
          Try
            IsVirtualLoaded = True
            MySinkEvent.Load()
          Catch ex As Exception
            Me.Clear()
            Throw ex
          End Try
        End If
      End If
    End If
  End Sub

  Private Sub IDataVirtual_Release() Implements IDataVirtual.Release
    If IsVirtualEnabled Then
      If IsVirtualLoaded = True Then
        Me.Clear()
      End If
    End If
  End Sub

  Private Property IDataVirtual_Count As Integer Implements IDataVirtual.Count
    Get
      Return MyCountMax
    End Get
    Set(value As Integer)
      If IDataVirtual_Enabled Then
        MyCountMax = value
        If Me.CountTotal < MyCountMax Then
          Me.CountTotal = MyCountMax
        End If
      End If
    End Set
  End Property
#End Region 'IDataVirtual
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      If MyDateUpdateStart.HasValue Then
        Return MyDateUpdateStart.Value
      Else
        If MyList.Count > 0 Then
          Return MyList.First.DateStart
        Else
          'return the default initialization date
          Return YahooAccessData.ReportDate.DateNullValue
        End If
      End If
    End Get
    Set(value As Date)
      MyDateUpdateStart = value
    End Set
  End Property

  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      If MyDateUpdateStop.HasValue Then
        Return MyDateUpdateStop.Value
      Else
        If MyList.Count > 0 Then
          Return MyList.Last.DateStop
        Else
          'return the default initialization date
          Return YahooAccessData.ReportDate.DateNullValue
        End If
      End If
    End Get
    Set(value As Date)
      MyDateUpdateStop = value
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
#End Region   'IDateUpdate
#Region "IDateTrade"
  Private Property IDateTrade_DateStart As Date Implements IDateTrade.DateStart
    Get
      If MyDateTradeStart.HasValue Then
        Return CDate(MyDateUpdateStart)
      Else
        If MyList.Count > 0 Then
          Dim ThisT = MyList.First
          If TypeOf ThisT Is YahooAccessData.IDateTrade Then
            Return DirectCast(ThisT, YahooAccessData.IDateTrade).DateStart
          Else
            'return the default initialization date
            Return YahooAccessData.ReportDate.DateNullValue
          End If
        Else
          'return the default initialization date
          Return YahooAccessData.ReportDate.DateNullValue
        End If
      End If
    End Get
    Set(value As Date)
      MyDateTradeStart = value
    End Set
  End Property

  Private Property IDateTrade_DateStop As Date Implements IDateTrade.DateStop
    Get
      If MyDateTradeStop.HasValue Then
        Return CDate(MyDateTradeStop)
      Else
        If MyList.Count > 0 Then
          Dim ThisT = MyList.Last
          If TypeOf ThisT Is YahooAccessData.IDateTrade Then
            Return DirectCast(ThisT, YahooAccessData.IDateTrade).DateStop
          Else
            'return the default initialization date
            Return YahooAccessData.ReportDate.DateNullValue
          End If
        Else
          'return the default initialization date
          Return YahooAccessData.ReportDate.DateNullValue
        End If
      End If
    End Get
    Set(value As Date)
      MyDateTradeStop = value
    End Set
  End Property
#End Region
End Class

#Region "LinkedHashSet(Of T)"
<Serializable()>
Public Class LinkedHashSet(Of T As {New, Class, IRegisterKey(Of String), IDateUpdate, IFormatData})
  Inherits LinkedHashSet(Of T, String)
End Class
#End Region

Public Class ListScaled
  Inherits List(Of Double)


  Private MyMax As Double
  Private MyMin As Double
  Private MyValueLast As Double

  Public Sub New()
    MyBase.New()
    MyMax = Double.MinValue
    MyMin = Double.MaxValue
  End Sub

  Public Sub New(capacity As Integer)
    MyBase.New(capacity)
    MyMax = Double.MinValue
    MyMin = Double.MaxValue
  End Sub

  Public Sub New(collection As IEnumerable(Of Double))
    MyBase.New(collection)
    MyMax = collection.Max
    MyMin = collection.Min
  End Sub

  Public Overloads Property Item(index As Integer, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double
    Get
      Dim ThisValue As Double = MyBase.Item(index)

      Dim ThisSlope As Double = Me.ScaleGainTo(ScaleToMinValue, ScaleToMaxValue)
      'check the slope denominator
      If ThisSlope = 0 Then
        'divide by zero
        'return the average of of the scale
        Return (ScaleToMinValue + ScaleToMaxValue) / 2
      Else
        'complete the scaling calculation
        Dim ThisOffset As Double = ScaleToMinValue + ThisSlope * Me.ScaleRange
        Return ThisSlope * ThisValue + ThisOffset
      End If
    End Get
    Set(value As Double)
      MyBase.Item(index) = value
    End Set
  End Property

  Public Overloads Sub Add(ByRef item As Double)
    MyBase.Add(item)
    If item > MyMax Then
      MyMax = item
    End If
    If item < MyMin Then
      MyMin = item
    End If
    MyValueLast = item
  End Sub

  Public Function Last() As Double
    Return MyValueLast
  End Function

  Public Property Max() As Double
    Get
      Return MyMax
    End Get
    Set(value As Double)
      MyMax = value
    End Set
  End Property

  Public Property Min() As Double
    Get
      Return MyMin
    End Get
    Set(value As Double)
      MyMin = value
    End Set
  End Property

  Public ReadOnly Property ScaleRange As Double
    Get
      If Me.Max > Math.Abs(Me.Min) Then
        Return Me.Max
      Else
        Return Math.Abs(Me.Min)
      End If
    End Get
  End Property

  Public ReadOnly Property ScaleGainTo(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double
    Get
      Dim ThisSlope As Double

      ThisSlope = 2 * Me.ScaleRange
      'check the slope denominator
      If ThisSlope = 0 Then
        Return 0
      Else
        'complete the slope scaling calculation
        Return (ScaleToMaxValue - ScaleToMinValue) / ThisSlope
      End If
    End Get
  End Property

  Public Overloads Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
    Return Me.ToArray(MyMin, MyMax, ScaleToMinValue, ScaleToMaxValue)
  End Function

  Public Overloads Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
    Dim ThisValues() As Double
    Dim ThisSlope As Double
    Dim ThisOffset As Double

    ThisValues = Me.ToArray
    If Me.Count > 0 Then
      'find the min and max for scaling
      ThisSlope = MaxValueInitial - MinValueInitial
      'check the slope denominator
      If ThisSlope = 0 Then
        'divide by zero
        'return all values as ScaleFromMinValue
        For I = 0 To ThisValues.Length - 1
          ThisValues(I) = ScaleToMinValue
        Next
      Else
        'complete the slope scaling calculation
        ThisSlope = (ScaleToMaxValue - ScaleToMinValue) / ThisSlope
        ThisOffset = ScaleToMinValue - ThisSlope * MinValueInitial
        For I = 0 To ThisValues.Length - 1
          ThisValues(I) = ThisSlope * ThisValues(I) + ThisOffset
        Next
      End If
    End If
    Return ThisValues
  End Function
End Class

#Region "SearchOfSubKey"
Public Class SearchSubKey(Of T As {New, Class})
  Implements IComparer(Of T)

  Private MyData As IEnumerable(Of T)
  Private MyDataAsSearch As YahooAccessData.ISearchKey(Of T, String)
  Private MyIndexLast As Integer
  Private MySubKeyLast As String

  Public Sub New(ByVal Data As IEnumerable(Of T))
    Dim ThisT = New T
    If TypeOf Data Is YahooAccessData.ISearchKey(Of T, String) = False Then
      Throw New InvalidCastException
    End If
    If TypeOf ThisT Is YahooAccessData.IRegisterKey(Of String) = False Then
      Throw New InvalidCastException
    End If
    MyData = Data
    MyDataAsSearch = DirectCast(MyData, YahooAccessData.ISearchKey(Of T, String))
    MyIndexLast = -1
  End Sub


  ''' <summary>
  ''' Seach a subkey value from the current index (zero base)
  ''' </summary>
  ''' <param name="SubKey">
  ''' A substring of the key
  ''' </param>
  ''' <param name="Index">
  ''' the current index array position. The seach will scan throw all the valid position 
  ''' from the current index. 
  ''' </param>
  ''' <param name="IsShiftKey">if True the search will assume that the subkey is the second element of the subkey
  ''' with the first element using the previous search subkey fist element </param>
  ''' <returns>The position of the new search</returns>
  Public Function Find(ByVal SubKey As String, ByVal Index As Integer, Optional ByVal IsShiftKey As Boolean = False) As Integer
    Dim ThisSearchResult As Integer = -1
    Dim ThisSubKey As String
    Dim ThisIndex As Integer

    Dim ThisT As T = New T

    ThisSubKey = UCase(Mid(SubKey, SubKey.Length, 1))
    If IsShiftKey Then
      ThisSubKey = String.Concat(Mid(MySubKeyLast, 1, 1), ThisSubKey)
    End If
    'Debug.Print(Index.ToString)
    'MyIndexLast = Index
    If Index > MyIndexLast + 1 Then
      MyIndexLast = Index
    ElseIf Index < MyIndexLast - 1 Then
      MyIndexLast = Index
    End If
    'the user manually moved up in the list
    'MyIndexLast = Index

    'End If
    If ThisSubKey.Length = 0 Then Return -1
    DirectCast(ThisT, YahooAccessData.IRegisterKey(Of String)).KeyValue = ThisSubKey
    If ThisSubKey = MySubKeyLast Then
      ThisIndex = MyIndexLast + 1
      If ThisIndex >= MyData.Count Then
        ThisIndex = MyData.Count - 1
        'reset the search to a new element
        MySubKeyLast = "~"   'set to an invalid character that wiil cause a new search
      End If
      If (FindKeyOfData(ThisIndex, MySubKeyLast.Length) = MySubKeyLast) Then
        'no search needed
        'increase the index by 1 
        ThisSearchResult = ThisIndex
      Else
        'search for first element of the subkey 
        ThisSearchResult = MyDataAsSearch.Find(ThisT, Me)
        If ThisSearchResult < 0 Then
          ThisSearchResult = (ThisSearchResult Xor -1) - 1
        End If
        'do a linear correction search from here
        Do
          ThisIndex = ThisSearchResult - 1
          If ThisIndex < 0 Then Exit Do
          Select Case Me.Compare(MyData(ThisIndex), ThisT)
            Case -1
              'this instance precede 
              Exit Do
            Case 0
              'this instance egual the subkey
              ThisSearchResult = ThisIndex
            Case 1
              Exit Do
          End Select
        Loop
      End If
    Else
      'new search
      ThisSearchResult = MyDataAsSearch.Find(ThisT, Me)
      If ThisSearchResult < 0 Then
        ThisSearchResult = (ThisSearchResult Xor -1) - 1
      End If
      'search for first element of the subkey 
      ThisSearchResult = MyDataAsSearch.Find(ThisT, Me)
      If ThisSearchResult < 0 Then
        ThisSearchResult = (ThisSearchResult Xor -1) - 1
      End If
      'do a linear correction search from here
      Do
        ThisIndex = ThisSearchResult - 1
        If ThisIndex < 0 Then Exit Do
        Select Case Me.Compare(MyData(ThisIndex), ThisT)
          Case -1
            'this instance precede 
            Exit Do
          Case 0
            'this instance egual the subkey
            ThisSearchResult = ThisIndex
          Case 1
            Exit Do
        End Select
      Loop
    End If
    MySubKeyLast = ThisSubKey
    MyIndexLast = ThisSearchResult
    Return MyIndexLast
  End Function

  Public ReadOnly Property Index As Integer
    Get
      Return MyIndexLast
    End Get
  End Property

  Public ReadOnly Property SubKey As String
    Get
      Return MySubKeyLast
    End Get
  End Property

#Region "Private Function"
  Private Function FindKeyOfData(ByVal Index As Integer, ByVal Length As Integer) As String
    Return UCase(Mid(DirectCast(MyData(Index), YahooAccessData.IRegisterKey(Of String)).KeyValue, 1, Length))
    'Return UCase(Mid(MyData(Index).KeyValue, 1, Length))
  End Function

  Private Function FindKeyOfData(ByVal Index As Integer) As String
    Return UCase(DirectCast(MyData(Index), YahooAccessData.IRegisterKey(Of String)).KeyValue)
    'Return UCase(MyData(Index).KeyValue)
  End Function

  Private Function FindKeyOfData(ByRef Data As T) As String
    Return UCase(DirectCast(Data, YahooAccessData.IRegisterKey(Of String)).KeyValue)
    'Return UCase(Data.KeyValue)
  End Function

  Private Function FindKeyOfData(ByRef Data As T, ByVal Length As Integer) As String
    Return UCase(Mid(DirectCast(Data, YahooAccessData.IRegisterKey(Of String)).KeyValue, 1, Length))
    'Return UCase(Mid(Data.KeyValue, 1, Length))
  End Function
#End Region
#Region "IComparer(of T)"
  Public Function Compare(x As T, y As T) As Integer Implements IComparer(Of T).Compare
    Dim ThisSubKeyForY = FindKeyOfData(y)
    Return FindKeyOfData(x, ThisSubKeyForY.Length).CompareTo(ThisSubKeyForY)
  End Function
#End Region
End Class
#End Region