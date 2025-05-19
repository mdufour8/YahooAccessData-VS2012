Imports YahooAccessData.MathPlus.Filter

Public Interface IFilterRun
	Function FilterRun(Value As Double) As Double
	ReadOnly Property FilterLast As Double
	ReadOnly Property FilterLast(Index As Integer) As Double
	ReadOnly Property FilterTrendLast As Double
	ReadOnly Property FilterRate() As Double
	ReadOnly Property FilterDetails() As String
	Sub Reset()
	ReadOnly Property IsReset As Boolean
End Interface
