Public Class PriceVolShort
  Implements IPriceVolShort

  Public Sub New()
    Me.DateStamp = Now
    Me.Open = 0.0
    Me.Low = 0.0
    Me.High = 0.0
    Me.Last = 0.0
    Me.Vol = 0
  End Sub

  Public Sub New(ByRef PriceVol As PriceVol)
    Me.New(PriceVol.AsIPriceVol)
  End Sub

  Public Sub New(ByRef PriceVol As IPriceVol)
    With PriceVol
      Me.DateStamp = .DateDay
      Me.Open = .Open
      Me.Low = .Low
      Me.High = .High
      Me.Last = .Last
      Me.Vol = .Vol
    End With
  End Sub

  Private ReadOnly Property AsIPriceVolShort As IPriceVolShort Implements IPriceVolShort.AsIPriceVol
    Get
      Return Me
    End Get
  End Property

  Public Property DateStamp As Date Implements IPriceVolShort.DateStamp
  Public Property Open As Single Implements IPriceVolShort.Open
  Public Property High As Single Implements IPriceVolShort.High
  Public Property Low As Single Implements IPriceVolShort.Low
  Public Property Last As Single Implements IPriceVolShort.Last
  Public Property Vol As Integer Implements IPriceVolShort.Vol
End Class