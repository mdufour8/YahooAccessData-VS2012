#Region "IPriceVolShort"
Public Interface IPriceVolShort
	ReadOnly Property AsIPriceVol As IPriceVolShort
	Property DateStamp As Date
	Property High As Single
	Property Open As Single
	Property Low As Single
	Property Last As Single
	Property Vol As Integer
End Interface
#End Region