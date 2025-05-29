Imports YahooAccessData.MathPlus.Filter

Public Class FilterCDFProbability
	Implements IFilterRun

	Private MyFilter As IFilterRun
	Private MyVolatility As IFilterRun
	Private MyListOfProbability As List(Of Double)
	Private MyPivotOffset As Integer
	Private MyCircularBuffer As CircularBuffer(Of Double)

	''' <summary>
	''' Initializes a new instance of the FilterCDFProbability class.
	''' </summary>
	''' <param name="Filter">The filter providing past and current price samples.</param>
	''' <param name="Volatility">The filter providing the most recent volatility estimate.</param>
	''' <param name="PivotOffset">The number of samples ago to use as the pivot reference point.</param>
	''' <remarks>
	''' This constructor initializes the CDF probability calculation based 
	''' on the provided filters and pivot offset.
	''' </remarks>	
	Public Sub New(Filter As IFilterRun, Volatility As IFilterRun, PivotOffset As Integer)
		MyFilter = Filter
		MyVolatility = Volatility
		MyPivotOffset = PivotOffset
	End Sub

	Public ReadOnly Property InputLast As Double Implements IFilterRun.InputLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterLast(Index As Integer) As Double Implements IFilterRun.FilterLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun.FilterRate
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilterRun.Count
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of Double) Implements IFilterRun.ToList
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property IsReset As Boolean Implements IFilterRun.IsReset
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		Throw New NotImplementedException()
	End Sub

	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun.Reset
		Throw New NotImplementedException()
	End Sub

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		Throw New NotImplementedException()
	End Function
End Class
