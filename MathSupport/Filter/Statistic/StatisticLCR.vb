Imports YahooAccessData.MathPlus.Filter

''' <summary>
''' Count the number of time 
''' </summary>
Public Class StatisticLCR
	Implements IFilterRun

	Public Sub New(ByVal FilterRate As Integer)

	End Sub

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

	Public ReadOnly Property FilterRate As Double Implements IFilterRun.FilterRate
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

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property InputLast As Double Implements IFilterRun.InputLast
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
