Public Class PriceVolGain
	Implements IPriceVolGain

	Private _PriceVolGain As IPriceVolGain
	Private _PriceVolData As IPriceVol

	Public Sub New(PriceVolGain As IPriceVolGain)
		_PriceVolGain = PriceVolGain
		_PriceVolData = _PriceVolGain.AsIPriceVol
	End Sub

	Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
		Get
			Return Me
		End Get
	End Property

	Public Property DateDay As Date Implements IPriceVol.DateDay
		Get
			Return _PriceVolGain.DateDay
		End Get
		Private Set(value As Date)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property DateUpdate As Date Implements IPriceVol.DateUpdate
		Get
			Return _PriceVolGain.DateUpdate
		End Get
		Private Set(value As Date)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property Open As Single Implements IPriceVol.Open
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property OpenNext As Single Implements IPriceVol.OpenNext
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property Last As Single Implements IPriceVol.Last
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property LastPrevious As Single Implements IPriceVol.LastPrevious
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property LastWeighted As Single Implements IPriceVol.LastWeighted
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property LastAdjusted As Single Implements IPriceVol.LastAdjusted
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property High As Single Implements IPriceVol.High
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property Low As Single Implements IPriceVol.Low
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property Vol As Integer Implements IPriceVol.Vol
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Integer)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property VolPlus As Integer Implements IPriceVol.VolPlus
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Integer)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property VolMinus As Integer Implements IPriceVol.VolMinus
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Integer)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Boolean)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property Range As Single Implements IPriceVol.Range
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Boolean)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Property SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As Single)
			Throw New NotImplementedException()
		End Set
	End Property

	Private _PriceOffset As Double?
	Public ReadOnly Property PriceOffset As Double? Implements IPriceVolGain.PriceOffset
		Get
			Return PriceOffset
		End Get
	End Property

	Private _PriceReference As Double?
	Public ReadOnly Property PriceReference As Double? Implements IPriceVolGain.PriceReference
		Get
			Return PriceReference
		End Get
	End Property

	Public Sub SetPriceReference(PriceRefValue As Double, PriceOffset As Double) Implements IPriceVolGain.SetPriceReference
		_PriceReference = PriceRefValue
		_PriceOffset = PriceOffset
	End Sub
End Class
