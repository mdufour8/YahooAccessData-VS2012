<Serializable>
Public Class PriceVolData
  Implements IPriceVol

  Public Sub New()

  End Sub

  Public Sub New(ByVal PriceVol As IPriceVol)
    With PriceVol
      Me.DateDay = .DateDay
      Me.DateUpdate = .DateUpdate
      Me.High = .High
      Me.IsIntraDay = .IsIntraDay
      Me.Last = .Last
      Me.LastAdjusted = .LastAdjusted
      Me.LastPrevious = .LastPrevious
      Me.LastWeighted = .LastWeighted
      Me.Low = .Low
      Me.Open = .Open
      Me.OpenNext = .OpenNext
      Me.Range = .Range
      Me.Vol = .Vol
      Me.VolMinus = .VolMinus
      Me.VolPlus = .VolPlus
    End With
  End Sub

  Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
    Get
      Return Me
    End Get
  End Property

  Public Property DateDay As Date Implements IPriceVol.DateDay
  Public Property DateUpdate As Date Implements IPriceVol.DateUpdate
  Public Property High As Single Implements IPriceVol.High
  Public Property IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
  Public Property Last As Single Implements IPriceVol.Last
  Public Property LastAdjusted As Single Implements IPriceVol.LastAdjusted
  Public Property LastPrevious As Single Implements IPriceVol.LastPrevious
  Public Property LastWeighted As Single Implements IPriceVol.LastWeighted
  Public Property Low As Single Implements IPriceVol.Low
  Public Property Open As Single Implements IPriceVol.Open
  Public Property OpenNext As Single Implements IPriceVol.OpenNext
  Public Property Range As Single Implements IPriceVol.Range
  Public Property Vol As Integer Implements IPriceVol.Vol
  Public Property VolMinus As Integer Implements IPriceVol.VolMinus
  Public Property VolPlus As Integer Implements IPriceVol.VolPlus

  Public Property IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
  Public Property SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
End Class
