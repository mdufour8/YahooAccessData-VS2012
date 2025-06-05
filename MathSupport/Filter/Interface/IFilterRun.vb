Imports YahooAccessData.MathPlus.Filter

Public Interface IFilterRun
	Function FilterRun(Value As Double) As Double
	ReadOnly Property InputLast As Double

	ReadOnly Property FilterLast As Double
	ReadOnly Property FilterLast(Index As Integer) As Double
	ReadOnly Property FilterTrendLast As Double
	ReadOnly Property FilterRate() As Double
	ReadOnly Property ToBufferList() As IList(Of Double)
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
	ReadOnly Property ToBufferList() As IList(Of T)
	ReadOnly Property FilterDetails() As String
	Sub Reset()
	Sub Reset(BufferCapacity As Integer)
	ReadOnly Property IsReset As Boolean
End Interface

Public Interface IFilterChain
	ReadOnly Property Filters As IReadOnlyList(Of IFilterRun)
End Interface

Public Interface IFilterNode
	Function Process(inputs As IReadOnlyList(Of Double)) As IReadOnlyList(Of Double)
	ReadOnly Property InputCount As Integer
	ReadOnly Property OutputCount As Integer
	ReadOnly Property Name As String
	Sub Reset()
End Interface

''' <summary>
'''	If indexing isn't intuitive enough, define named signal connections:
''' </summary>
Public Interface INamedFilterNode
	Function Process(inputs As IDictionary(Of String, Double)) As IDictionary(Of String, Double)
	ReadOnly Property InputNames As IReadOnlyList(Of String)
	ReadOnly Property OutputNames As IReadOnlyList(Of String)
	ReadOnly Property Name As String
	Sub Reset()
End Interface








