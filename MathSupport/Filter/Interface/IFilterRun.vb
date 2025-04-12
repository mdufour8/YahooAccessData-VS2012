Imports YahooAccessData.MathPlus.Filter

Public Interface IFilterRun
	Function FilterRun(Value As Double) As Double
	ReadOnly Property FilterLast As Double
	ReadOnly Property FilterRate() As Double
	ReadOnly Property FilterDetails() As String
	Sub Reset()
End Interface
