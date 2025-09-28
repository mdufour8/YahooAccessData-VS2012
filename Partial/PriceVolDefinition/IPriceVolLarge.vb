Public Interface IPriceVolLarge
	ReadOnly Property AsIPriceVolLarge As IPriceVolLarge
	Property DateDay As Date
	Property DateUpdate As Date
	Property Open As Double
	Property OpenNext As Double
	Property Last As Double
	Property LastPrevious As Double
	Property LastWeighted As Double
	Property LastAdjusted As Double
	Property High As Double
	Property Low As Double
	Property Vol As Integer
	Property VolPlus As Integer
	Property VolMinus As Integer
	Property IsIntraDay As Boolean
	Property Range As Double
	Property FilterLast As Double
End Interface