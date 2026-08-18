Public Interface IStockPriceVolExport
	Property DateDay As Date
	Property High As Double
	Property Open As Double
	Property Low As Double
	Property Last As Double
	Property Volume As Long
	ReadOnly Property LDV As Double
End Interface
