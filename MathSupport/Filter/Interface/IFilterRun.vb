Imports YahooAccessData.MathPlus.Filter

Public Interface IFilterRun
	Function FilterRun(Value As Double) As Double
	ReadOnly Property InputLast As Double

	ReadOnly Property FilterLast As Double
	ReadOnly Property FilterLast(Index As Integer) As Double
	ReadOnly Property FilterTrendLast As Double
	ReadOnly Property FilterRate() As Double
	ReadOnly Property Count() As Integer
	ReadOnly Property ToList() As IList(Of Double)
	ReadOnly Property FilterDetails() As String
	Sub Reset()
	Sub Reset(BufferCapacity As Integer)
	ReadOnly Property IsReset As Boolean
End Interface



Public Interface IFilterRun(Of T)
	Function FilterRun(Value As Double) As T
	ReadOnly Property InputLast As Double

	ReadOnly Property FilterLast As T
	ReadOnly Property FilterLast(Index As Integer) As T
	ReadOnly Property FilterTrendLast As T
	ReadOnly Property FilterRate() As Double
	ReadOnly Property Count() As Integer
	ReadOnly Property ToList() As IList(Of T)
	ReadOnly Property FilterDetails() As String
	Sub Reset()
	Sub Reset(BufferCapacity As Integer)
	ReadOnly Property IsReset As Boolean
End Interface




