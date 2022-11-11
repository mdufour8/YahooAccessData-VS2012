Namespace ExtensionService
  Public Class VBArray(Of T)
    Private ThisLowBound As Integer
    Private ThisHighBound As Integer

    Private ThisArray() As T

    Public Sub New(ByVal LBound As Integer, ByVal UBound As Integer)
      Me.ReDim(LBound, UBound)
    End Sub

    Public ReadOnly Property LBound As Integer
      Get
        Return ThisLowBound
      End Get
    End Property

    Public ReadOnly Property UBound As Integer
      Get
        Return ThisHighBound
      End Get
    End Property

    Public Sub [ReDim](ByVal LBound As Integer, ByVal UBound As Integer)
      ThisLowBound = LBound
      ThisHighBound = UBound
      ReDim ThisArray(0 To (ThisHighBound - ThisLowBound))
    End Sub

    Public Function ToArray() As T()
      Return ThisArray
    End Function

    Default Public Property Item(ByVal Index As Integer) As T
      Get
        Return ThisArray(Index - ThisLowBound)
      End Get

      Set(ByVal Value As T)
        ThisArray(Index - ThisLowBound) = Value
      End Set
    End Property

    Public Overrides Function ToString() As String
      Return String.Format("{0}({1} to {2})", TypeName(Me), Me.LBound, Me.UBound)
    End Function
  End Class
End Namespace
