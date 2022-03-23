Public Class TimePeriodOverlap
  Private MyPeriodStart As Date
  Private MyPeriodStop As Date

  Public Sub New(ByVal PeriodStart As Date, ByVal PeriodStop As Date)
    MyPeriodStart = PeriodStart
    MyPeriodStop = PeriodStop
  End Sub

  Public Sub New(ByRef DateRange As IDateRange)
    MyPeriodStart = DateRange.DateStart
    MyPeriodStop = DateRange.DateStop
  End Sub

  Public Sub New(ByRef DateUpdate As IDateUpdate)
    MyPeriodStart = DateUpdate.DateStart
    MyPeriodStop = DateUpdate.DateStop
  End Sub

  Public Function IsOverlap(ByVal DateTime As Date) As Boolean
    If (DateTime >= MyPeriodStart) Then
      If (DateTime <= MyPeriodStop) Then
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Function IsOverlap(ByVal DateStart As Date, ByVal DateStop As Date) As Boolean
    If IsOverlap(DateStart) Then
      Return True
    Else
      Return IsOverlap(DateStop)
    End If
  End Function

  Public Function IsOverlap(ByRef DateRange As IDateRange) As Boolean
    Return IsOverlap(DateRange.DateStart, DateRange.DateStop)
  End Function

  Public Function IsOverlap(ByRef DateUpdate As IDateUpdate) As Boolean
    If TypeOf DateUpdate Is IRecordInfo Then
      If DirectCast(DateUpdate, IRecordInfo).CountTotal > 0 Then
        Return IsOverlap(DateUpdate.DateStart, DateUpdate.DateStop)
      Else
        Return False
      End If
    Else
      Return IsOverlap(DateUpdate.DateStart, DateUpdate.DateStop)
    End If
  End Function
End Class
