Imports YahooAccessData.MathPlus.Filter

Public Interface IFilterRun
	Function FilterRun(Value As Double) As Double
	Function FilterRun(Value As Double, FilterPLLDetector As IFilterPLLDetector) As Double
	ReadOnly Property FilterLast As Double
	ReadOnly Property FilterRate() As Double
	Sub Reset()
End Interface
