Public Interface IPriceVol
	ReadOnly Property AsIPriceVol As IPriceVol
	Property DateDay As Date
	Property DateUpdate As Date
	Property Open As Single
	Property OpenNext As Single
	Property Last As Single
	Property LastPrevious As Single
	Property LastWeighted As Single
	Property LastAdjusted As Single
	Property High As Single
	Property Low As Single
	Property Vol As Integer
	Property VolPlus As Integer
	Property VolMinus As Integer
	Property IsIntraDay As Boolean
	Property Range As Single
	Property IsSpecialDividendPayout As Boolean
	Property SpecialDividendPayoutValue As Single
End Interface

Public Interface IPriceVol(Of T)
	ReadOnly Property AsIPriceVol As IPriceVol(Of T)
	Property DateDay As Date
	Property DateUpdate As Date
	Property Open As T
	Property OpenNext As T
	Property Last As T
	Property FilterLast As T
	Property LastPrevious As T
	Property LastWeighted As T
	Property LastAdjusted As T
	Property High As T
	Property Low As T
	Property Vol As Integer
	Property VolPlus As Integer
	Property VolMinus As Integer
	Property IsIntraDay As Boolean
	Property Range As T
End Interface
