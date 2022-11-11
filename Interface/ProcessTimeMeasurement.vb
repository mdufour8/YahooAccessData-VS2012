Public Class ProcessTimeMeasurement
  Implements IProcessTimeMeasurement

  Private MyKey As String
  Private MyName As String
  Private MyElapsedTime As Long

  Public Sub New()

  End Sub

  Public Sub New(ByVal Key As String, ByVal Name As String)
    MyKey = Key
    MyName = Name
  End Sub

  Public Sub New(ByVal Key As String, ByVal Name As String, ByVal ElapsedTimeInMilliseconds As Long)
    MyKey = Key
    MyName = Name
    MyElapsedTime = ElapsedTimeInMilliseconds
  End Sub

  Public Property ElapsedMilliseconds As Long Implements IProcessTimeMeasurement.ElapsedMilliseconds
    Get
      Return MyElapsedTime
    End Get
    Set(value As Long)
      MyElapsedTime = value
    End Set
  End Property

  Public ReadOnly Property Key As String Implements IProcessTimeMeasurement.Key
    Get
      Return MyKey
    End Get
  End Property

  Public Property Name As String Implements IProcessTimeMeasurement.Name
End Class
