Public Class FilterCascaded
	Implements IFilterRun, IFilterChain

	Private ReadOnly _filters As List(Of IFilterRun)

	Public Sub New(ParamArray filters As IFilterRun())
		If filters Is Nothing OrElse filters.Length = 0 Then
			Throw New ArgumentException("At least one filter must be provided.")
		End If
		_filters = filters.ToList()
	End Sub

	Public Function FilterRun(value As Double) As Double Implements IFilterRun.FilterRun
		Dim result As Double = value
		For Each Filter In _filters
			result = Filter.FilterRun(result)
		Next
		Return result
	End Function

	Public ReadOnly Property InputLast As Double Implements IFilterRun.InputLast
		Get
			Return _filters.First().InputLast
		End Get
	End Property

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return _filters.Last().FilterLast
		End Get
	End Property

	Public ReadOnly Property FilterLast(index As Integer) As Double Implements IFilterRun.FilterLast
		Get
			Return _filters.Last().FilterLast(index)
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Return _filters.Last().FilterTrendLast
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double() Implements IFilterRun.FilterRate
		Get
			Return _filters.Last().FilterRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilterRun.Count
		Get
			Return _filters.Last().Count
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of Double) Implements IFilterRun.ToList
		Get
			Return _filters.Last().ToList
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return String.Join(" → ", _filters.Select(Function(f) f.FilterDetails))
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		For Each Filter In _filters
			Filter.Reset()
		Next
	End Sub

	Public Sub Reset(bufferCapacity As Integer) Implements IFilterRun.Reset
		For Each Filter In _filters
			Filter.Reset(bufferCapacity)
		Next
	End Sub

	Public ReadOnly Property IsReset As Boolean Implements IFilterRun.IsReset
		Get
			Return _filters.All(Function(f) f.IsReset)
		End Get
	End Property

	Public ReadOnly Property Filters As IReadOnlyList(Of IFilterRun) Implements IFilterChain.Filters
		Get
			Return _filters
		End Get
	End Property
End Class

