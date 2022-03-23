#Region "TreeNode(Of T)"
#Region "Interface definition"
Public Interface ITreeNode(Of T)
  Property Item As T
  Property Name As String
  Property Key As String
  Property Tag As Object
  Property ToChildren As List(Of TreeNodeList(Of T))
  Property ToParent As ITreeNode(Of T)
  ReadOnly Property ToRootNode As ITreeNode(Of T)
  Function IsNodeExist(ByVal Key As String) As Boolean
  Function SearchChildren(ByVal Key As String) As ITreeNode(Of T)
  Function Search(ByVal Key As String) As ITreeNode(Of T)
  Function AddChildren(ByVal item As T) As ITreeNode(Of T)
  Function AddChildren(ByVal item As T, ByVal Name As String) As ITreeNode(Of T)
  Function AddChildren(ByVal item As T, ByVal Name As String, ByVal Key As String) As ITreeNode(Of T)
  Function AddChildren(TreeNode As TreeNodeList(Of T)) As ITreeNode(Of T)
  Function FullPath() As String
  Function FullPath(ByVal PathBase As String) As String
  Function Copy() As TreeNodeList(Of T)
  Sub RemoveChildren(ByVal Key As String)
End Interface

Public Interface ITreeTraverse(Of T)
  Function ToNode(ByVal Node As ITreeNode(Of T)) As ITreeNode(Of T)
End Interface
#End Region
<Serializable()>
Public Class TreeNodeList(Of T)
  Implements ITreeNode(Of T)
  Implements IEquatable(Of ITreeNode(Of T))
  Implements IComparable(Of ITreeNode(Of T))


  Private MyDictionaryOfChildren As Dictionary(Of String, ITreeNode(Of T))
  Private MyDictionaryOfFullPath As Dictionary(Of String, ITreeNode(Of T))
  Private MyBaseNode As TreeNodeList(Of T)
  Private MyParentFullPath As String
  Private MyParent As ITreeNode(Of T)
  Private MyRootNode As ITreeNode(Of T)

  Public Sub New()
    Me.ToChildren = New List(Of TreeNodeList(Of T))
    Me.ToParent = Nothing 'bydefault
    Me.Key = ""
    Me.Name = ""
    MyParentFullPath = ""
    MyDictionaryOfChildren = New Dictionary(Of String, ITreeNode(Of T))
  End Sub

  Public Sub New(item As T)
    Me.New(item, Nothing)
  End Sub

  Public Sub New(item As T, Parent As ITreeNode(Of T))
    Me.ToChildren = New List(Of TreeNodeList(Of T))
    Me.ToParent = Parent
    Me.Item = item
    Me.Key = ""
    Me.Name = ""
    MyParentFullPath = ""
    MyDictionaryOfChildren = New Dictionary(Of String, ITreeNode(Of T))
  End Sub

  Public Sub New(ByVal TreeNode As TreeNodeList(Of T))
    Me.New(TreeNode, IsCopyChildren:=True)
  End Sub

  Public Sub New(ByVal TreeNode As TreeNodeList(Of T), ByVal IsCopyChildren As Boolean)
    Me.New(TreeNode.Item, TreeNode.ToParent)
    Me.Key = TreeNode.Key
    Me.Name = TreeNode.Name

    If IsCopyChildren Then
      For Each ThisNode In TreeNode.ToChildren
        Me.ToChildren.Add(New TreeNodeList(Of T)(ThisNode))
      Next
    End If
  End Sub

  Public Sub New(ByVal TreeNode As TreeNodeList(Of T), ByVal TraverseEvent As ITreeTraverse(Of T))
    Me.New(TreeNode, TraverseEvent, IsCopyChildren:=True)
  End Sub

  Public Sub New(ByVal TreeNode As TreeNodeList(Of T), ByVal TraverseEvent As ITreeTraverse(Of T), ByVal IsCopyChildren As Boolean)
    Me.New(TreeNode.Item, TreeNode.ToParent)
    Dim ThisTreeNodeForChildren As TreeNodeList(Of T)
    Dim ThisNodeResult As ITreeNode(Of T)

    Me.Key = TreeNode.Key
    Me.Name = TreeNode.Name
    ThisNodeResult = TraverseEvent.ToNode(Me)

    If ThisNodeResult.Key Is Nothing Then
      Me.Key = Nothing
    Else
      Me.Key = ThisNodeResult.Key
    End If
    Me.Name = ThisNodeResult.Name
    If Me.Key Is Nothing Then Return
    If IsCopyChildren Then
      For Each ThisNode In TreeNode.ToChildren
        ThisTreeNodeForChildren = New TreeNodeList(Of T)(ThisNode, TraverseEvent, IsCopyChildren)
        If ThisTreeNodeForChildren.Key Is Nothing Then
          ThisTreeNodeForChildren.Key = ThisTreeNodeForChildren.Key
        Else
          Me.ToChildren.Add(ThisTreeNodeForChildren)
        End If
      Next
    End If
  End Sub

  Public Property Item As T Implements ITreeNode(Of T).Item
  Public Property Key As String Implements ITreeNode(Of T).Key
  Public Property Name As String Implements ITreeNode(Of T).Name
  Public Property ToChildren() As List(Of TreeNodeList(Of T)) Implements ITreeNode(Of T).ToChildren

  <Xml.Serialization.XmlIgnore>
  Public Property ToParent() As ITreeNode(Of T) Implements ITreeNode(Of T).ToParent
    Get
      Return MyParent
    End Get
    Set(value As ITreeNode(Of T))
      MyParent = value
      If MyParent IsNot Nothing Then
        MyParentFullPath = MyParent.FullPath
      Else
        MyParentFullPath = ""
      End If
      'find the root node
      MyRootNode = Me
      Do Until MyRootNode.ToParent Is Nothing
        MyRootNode = MyRootNode.ToParent
      Loop
    End Set
  End Property

  <Xml.Serialization.XmlIgnore>
  Public ReadOnly Property ToRootNode As ITreeNode(Of T) Implements ITreeNode(Of T).ToRootNode
    Get
      Return MyRootNode
    End Get
  End Property

  ''' <summary>
  ''' Use to update all parent object after deserialization
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Refresh(ByVal IsBuildDictionaryOfPath As Boolean)
    Me.Refresh()
    If IsBuildDictionaryOfPath Then
      MyDictionaryOfFullPath = New Dictionary(Of String, ITreeNode(Of T))
      Call BuildDictionaryOfFullPath(Me)
    End If
  End Sub

  ''' <summary>
  ''' Use to update all parent object after deserialization
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Refresh()
    MyDictionaryOfFullPath = Nothing
    MyDictionaryOfChildren.Clear()
    MyParentFullPath = ""
    Me.FullPath()
    For Each ThisNode In Me.ToChildren
      'If ThisNode.Name = "Test1" Then
      '  ThisNode.Name = "Test1"
      'End If
      ThisNode.ToParent = Me
      MyDictionaryOfChildren.Add(ThisNode.Key, ThisNode)
      ThisNode.Refresh()
    Next
    'If Me.Name = "Stock.rep" Then
    '  Me.Name = Me.Name
    'End If
  End Sub

  ''' <summary>
  ''' Copy the object element and it childrens
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Copy() As TreeNodeList(Of T) Implements ITreeNode(Of T).Copy
    Dim ThisTreeNodeList As New TreeNodeList(Of T)(Me)
    'remove the link to the parent of the first node
    ThisTreeNodeList.ToParent = Nothing
    Return ThisTreeNodeList
  End Function

  Public Function Copy(ByVal TraverseEvent As ITreeTraverse(Of T)) As TreeNodeList(Of T)
    Dim ThisTreeNodeList As New TreeNodeList(Of T)(Me, TraverseEvent)
    'remove the link to the parent of the first node
    ThisTreeNodeList.ToParent = Nothing
    Return ThisTreeNodeList
  End Function

  Private Function ITreeNode_IsNodeExist(Key As String) As Boolean Implements ITreeNode(Of T).IsNodeExist
    Return Me.IsNodeExist(Key)
  End Function

  Private Function ITreeNode_Search(Key As String) As ITreeNode(Of T) Implements ITreeNode(Of T).Search
    Return Me.Search(Key)
  End Function

  Private Function ITreeNode_SearchChildren(Key As String) As ITreeNode(Of T) Implements ITreeNode(Of T).SearchChildren
    Return Me.SearchChildren(Key)
  End Function

  Public Function AddChildren(item As T) As ITreeNode(Of T) Implements ITreeNode(Of T).AddChildren
    Dim ThisChildren = New TreeNodeList(Of T)(item, Me)
    Return Me.AddChildren(ThisChildren)
  End Function

  Public Function AddChildren(TreeNodeChildren As TreeNodeList(Of T)) As ITreeNode(Of T) Implements ITreeNode(Of T).AddChildren
    If MyDictionaryOfChildren.ContainsKey(TreeNodeChildren.Key) = False Then
      TreeNodeChildren.ToParent = Me
      MyDictionaryOfChildren.Add(TreeNodeChildren.Key, TreeNodeChildren)
      Me.ToChildren.Add(TreeNodeChildren)
      Return TreeNodeChildren
    Else
      Return Nothing
    End If
  End Function

  Public Function AddChildren(item As T, Name As String) As ITreeNode(Of T) Implements ITreeNode(Of T).AddChildren
    Dim ThisChildren = New TreeNodeList(Of T)(item, Me) With {.Name = Name, .Key = Name}
    Return Me.AddChildren(ThisChildren)
  End Function

  Public Function AddChildren(item As T, Name As String, Key As String) As ITreeNode(Of T) Implements ITreeNode(Of T).AddChildren
    Dim ThisChildren = New TreeNodeList(Of T)(item, Me) With {.Name = Name, .Key = Key}
    Return Me.AddChildren(ThisChildren)
  End Function

  Public Sub RemoveChildren(Key As String) Implements ITreeNode(Of T).RemoveChildren
    If MyDictionaryOfChildren.ContainsKey(Key) Then
      Dim ThisItem = DirectCast(MyDictionaryOfChildren(Key), TreeNodeList(Of T))
      Me.ToChildren.Remove(ThisItem)
      MyDictionaryOfChildren.Remove(Key)
    End If
  End Sub

  Public Function FullPath() As String Implements ITreeNode(Of T).FullPath
    Return Me.FullPath("")
  End Function

  Public Function FullPath(ByVal PathBase As String) As String Implements ITreeNode(Of T).FullPath
    Dim ThisParentFullPath As String = MyParentFullPath
    If ThisParentFullPath = "" Then
      If Me.ToParent IsNot Nothing Then
        MyParentFullPath = Me.ToParent.FullPath
        ThisParentFullPath = MyParentFullPath
      End If
    End If
    If PathBase.Length > 0 Then
      ThisParentFullPath = Strings.Replace(ThisParentFullPath, PathBase, "")
    End If
    If ThisParentFullPath.Length > 0 Then
      Return String.Format("{0}\{1}", ThisParentFullPath, Me.Name)
    Else
      Return Me.Name
    End If
  End Function


  ''' <summary>
  ''' Search all the parent and it's children for the key
  ''' </summary>
  ''' <param name="Key"></param>
  ''' <returns>return the object that contain the key </returns>
  ''' <remarks>the function return nothing if the object is not found</remarks>
  Public Function Search(ByVal Key As String) As ITreeNode(Of T)
    Dim ThisNodeResult As ITreeNode(Of T)
    If Me.Key = Key Then
      Return Me
    ElseIf MyDictionaryOfChildren.ContainsKey(Key) Then
      Return MyDictionaryOfChildren(Key)
    Else
      For Each ThisNode In Me.ToChildren
        ThisNodeResult = ThisNode.Search(Key)
        If ThisNodeResult IsNot Nothing Then
          Return ThisNodeResult
        End If
      Next
    End If
    Return Nothing
  End Function

  <Xml.Serialization.XmlIgnore>
  Public Property Tag As Object Implements ITreeNode(Of T).Tag

  ''' <summary>
  ''' Search the children of the current node
  ''' </summary>
  ''' <param name="Key"></param>
  ''' <returns>return the object that contain the key </returns>
  ''' <remarks>the function return nothing if the object is not found</remarks>
  Public Function SearchChildren(ByVal Key As String) As ITreeNode(Of T)
    If MyDictionaryOfChildren.ContainsKey(Key) Then
      Return MyDictionaryOfChildren(Key)
    Else
      Return Nothing
    End If
  End Function

  Public Function SearchPath(ByVal Path As String) As ITreeNode(Of T)
    If MyDictionaryOfFullPath Is Nothing Then
      MyDictionaryOfFullPath = New Dictionary(Of String, ITreeNode(Of T))
      Call BuildDictionaryOfFullPath(Me)
    End If
    If MyDictionaryOfFullPath.ContainsKey(Path) = True Then
      Return MyDictionaryOfFullPath(Path)
    Else
      Return Nothing
    End If
  End Function

  Private Sub BuildDictionaryOfFullPath(ByRef ThisNode As TreeNodeList(Of T))
    Dim ThisPath As String = ThisNode.FullPath
    If MyDictionaryOfFullPath.ContainsKey(ThisPath) = False Then
      MyDictionaryOfFullPath.Add(ThisPath, ThisNode)
      For Each ThisChildrenNode In ThisNode.ToChildren
        Call BuildDictionaryOfFullPath(ThisChildrenNode)
      Next
    End If
  End Sub

  Public Function IsNodeExist(ByVal Key As String) As Boolean
    Return Me.Search(Key) IsNot Nothing
  End Function

  ''' <summary>
  ''' Efficient recursive algorithm to build the tree search element
  ''' </summary>
  ''' <param name="TreeNodeBase"></param>
  ''' <remarks></remarks>
  Private Sub DictionaryBuild(ByRef Dictionary As Dictionary(Of String, ITreeNode(Of T)), ByRef TreeNodeBase As TreeNodeList(Of T))
    If Dictionary.ContainsKey(TreeNodeBase.Key) = False Then
      Dictionary.Add(TreeNodeBase.Key, TreeNodeBase)
      For Each ThisTreeNode In Me.ToChildren
        Call DictionaryBuild(Dictionary, ThisTreeNode)
      Next
    End If
  End Sub

  Public Overrides Function ToString() As String
    Return String.Format("{0}:{1}", Me.Key.ToString, Me.Name.ToString)
  End Function

#Region "IEquatable"
  Public Overloads Function Equals(other As ITreeNode(Of T)) As Boolean Implements IEquatable(Of ITreeNode(Of T)).Equals
    If other Is Nothing Then Return False
    'note that ToString also return the key
    If Me.Key = other.Key Then
      Return True
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is ITreeNode(Of T)) Then
      Return Me.Equals(DirectCast(obj, ITreeNode(Of T)))
    Else
      Return (False)
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.Key.GetHashCode()
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
  Public Function CompareTo(other As ITreeNode(Of T)) As Integer Implements IComparable(Of ITreeNode(Of T)).CompareTo
    Return Me.Name.CompareTo(other.Name)
  End Function

  'Public Function AreEqual(ByVal Value1 As T, ByVal Value2 As T) As Boolean
  '  Return EqualityComparer(Of T).[Default].Equals(Value1, Value2)
  'End Function
#End Region
End Class
#End Region
