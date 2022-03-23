Public Class CompareByName(Of T)
  Private _PropertyName As String

  Public Sub New()
    _PropertyName = ""
  End Sub

  Public Sub New(ByVal PropertyName As String)
    _PropertyName = PropertyName
  End Sub

  Public Function Compare(x As T, y As T) As Integer
    'the reflection is much faster than Callbyname VB function
    Dim ThisType As Type = x.GetType
    Dim ThisProperty As System.Reflection.PropertyInfo = ThisType.GetProperty(_PropertyName)
    If ThisProperty Is Nothing Then Return 0
    Dim ThisValueX = ThisProperty.GetValue(x, Nothing)
    Dim ThisValueY = ThisProperty.GetValue(y, Nothing)
    'this would be very slow compare to the code above
    'Dim ThisValueX = CallByName(x, MyTableName, CallType.Get, Nothing)
    'Dim ThisValueY = CallByName(y, MyTableName, CallType.Get, Nothing)
    If ThisValueX Is Nothing Then Return 0
    If ThisValueY Is Nothing Then Return 0
    If TypeOf ThisValueX Is Integer Then
      Return CType(ThisValueX, Integer).CompareTo(CType(ThisValueY, Integer))
    ElseIf TypeOf ThisValueX Is Single Then
      Return CType(ThisValueX, Single).CompareTo(CType(ThisValueY, Single))
    ElseIf TypeOf ThisValueX Is String Then
      Return CType(ThisValueX, String).CompareTo(CType(ThisValueY, String))
    ElseIf TypeOf ThisValueX Is Date Then
      Return CType(ThisValueX, Date).CompareTo(CType(ThisValueY, Date))
    ElseIf TypeOf ThisValueX Is Double Then
      Return CType(ThisValueX, Double).CompareTo(CType(ThisValueY, Double))
    ElseIf TypeOf ThisValueX Is Boolean Then
      Return CType(ThisValueX, Boolean).CompareTo(CType(ThisValueY, Boolean))
    ElseIf TypeOf ThisValueX Is [Enum] Then
      Return CType(ThisValueX, Integer).CompareTo(CType(ThisValueY, Integer))
    Else
      Return 0
    End If
  End Function

  Public Property PropertyName As String
    Get
      Return _PropertyName
    End Get
    Set(value As String)
      _PropertyName=value
    End Set
  End Property
End Class
