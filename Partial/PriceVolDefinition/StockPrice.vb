Public Class StockPrice
	Implements IStockPrice
	Implements IPriceVol

	Public Sub New()

	End Sub

	Public Sub New(PriceVol As IPriceVol)
		Me.DateDay = PriceVol.DateDay
		Me.Open = PriceVol.Open
		Me.OpenNext = PriceVol.OpenNext
		Me.Last = PriceVol.Last
		Me.LastPrevious = PriceVol.LastPrevious
		Me.High = PriceVol.High
		Me.Low = PriceVol.Low
		Me.Vol = PriceVol.Vol
	End Sub


	Public Property DateDay As Date Implements IStockPrice.DateDay
	Public Property Open As Double Implements IStockPrice.Open
	Public Property OpenNext As Double Implements IStockPrice.OpenNext
	Public Property Last As Double Implements IStockPrice.Last
	Public Property LastPrevious As Double Implements IStockPrice.LastPrevious
	Public Property High As Double Implements IStockPrice.High
	Public Property Low As Double Implements IStockPrice.Low
	Public Property Vol As Long Implements IStockPrice.Vol
	Public ReadOnly Property AsIStockPrice As IStockPrice Implements IStockPrice.AsIStockPrice
		Get
			Return Me
		End Get
	End Property

	Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
		Get
			Return Me
		End Get
	End Property

	Private Property IPriceVol_DateUpdate As Date Implements IPriceVol.DateUpdate
		Get
			Return Me.DateDay
		End Get
		Set(value As Date)
			Me.DateDay = value
		End Set
	End Property

	Private Property IPriceVol_LastWeighted As Single Implements IPriceVol.LastWeighted
		Get
			Return CSng(Me.Last)
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_LastAdjusted As Single Implements IPriceVol.LastAdjusted
		Get
			Return CSng(Last)
		End Get
		Set(value As Single)
			Last = value
		End Set
	End Property

	Private Property IPriceVol_VolPlus As Integer Implements IPriceVol.VolPlus
		Get
			If Vol > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Vol < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Vol)
			End If
		End Get
		Set(value As Integer)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_VolMinus As Integer Implements IPriceVol.VolMinus
		Get
			If Vol > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Vol < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Vol)
			End If
		End Get
		Set(value As Integer)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
		Get
			Return False ' Assuming this is not an intra-day price
		End Get
		Set(value As Boolean)
			'ignore the setter for now
		End Set
	End Property

	Private Property IPriceVol_Range As Single Implements IPriceVol.Range
		Get
			Return CSng(High - Low)
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
		Get
			Return False
		End Get
		Set(value As Boolean)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
		Get
			Return 0
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_DateDay As Date Implements IPriceVol.DateDay
		Get
			Return DateDay
		End Get
		Set(value As Date)
			DateDay = value
		End Set
	End Property

	Private Property IPriceVol_Open As Single Implements IPriceVol.Open
		Get
			Return CSng(Open)
		End Get
		Set(value As Single)
			Open = value
		End Set
	End Property

	Private Property IPriceVol_OpenNext As Single Implements IPriceVol.OpenNext
		Get
			Return CSng(OpenNext)
		End Get
		Set(value As Single)
			OpenNext = value
		End Set
	End Property

	Private Property IPriceVol_Last As Single Implements IPriceVol.Last
		Get
			Return CSng(Last)
		End Get
		Set(value As Single)
			Last = value
		End Set
	End Property

	Private Property IPriceVol_LastPrevious As Single Implements IPriceVol.LastPrevious
		Get
			Return CSng(LastPrevious)
		End Get
		Set(value As Single)
			LastPrevious = value
		End Set
	End Property

	Private Property IPriceVol_High As Single Implements IPriceVol.High
		Get
			Return CSng(High)
		End Get
		Set(value As Single)
			High = value
		End Set
	End Property

	Private Property IPriceVol_Low As Single Implements IPriceVol.Low
		Get
			Return CSng(Low)
		End Get
		Set(value As Single)
			Low = value
		End Set
	End Property

	Private Property IPriceVol_Vol As Integer Implements IPriceVol.Vol
		Get
			If Vol > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Vol < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Vol)
			End If
		End Get
		Set(value As Integer)
			Vol = value
		End Set
	End Property
End Class

Public Interface IStockPrice
	ReadOnly Property AsIStockPrice As IStockPrice
	Property DateDay As Date
	Property Open As Double
	Property OpenNext As Double
	Property Last As Double
	Property LastPrevious As Double
	Property High As Double
	Property Low As Double
	Property Vol As Long
End Interface
