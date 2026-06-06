Imports System.Threading.Tasks
Imports YahooAccessData.MathPlus.Filter


''' <summary>
''' Use to evaluate the probability that a stock will reach a certain price level based on historical data and volatility.
''' Hence the class calculates the probability that the future price exceed the peak range the previous observed data.The filter result 
''' is a value between 0 and 1, representing the probability that the previous peak are exceeded in the future based on the filter rate period
''' and the volatility of the stock.
''' </summary>
Public Class FilterProbabilityState

	Private _FilterRate As Double
	Private _FilterVolatilityRate As Double
	Private _FilterLast As IProbabilityState
	Private _ValueLast As IPriceVol
	Private _IsReset As Boolean

	Public Sub New(ByVal FilterRate As Double, ByVal FilterVolatilityRate As Double)
		_FilterRate = FilterRate
		_FilterVolatilityRate = FilterVolatilityRate
		_IsReset = True


	End Sub


	''' <summary>
	''' Calculate the probability that the future price exceed the peak range the previous observed data based on the filter rate period and the volatility of the stock.	
	''' </summary>
	''' <param name="Value">Stock price value to evaluate.</param>
	''' <returns></returns>
	Public Function FilterRun(Value As IPriceVol) As IProbabilityState
		If _IsReset Then
			'initialization
			_FilterLast = New ProbabilityState(0.5, 0.5)
			_ValueLast = Value
			_IsReset = False
		End If
		'calculation here




		_ValueLast = Value
		Return _FilterLast
	End Function


	Public ReadOnly Property InputLast As IPriceVol
		Get
			Return _ValueLast
		End Get
	End Property

	Public ReadOnly Property FilterLast As IProbabilityState
		Get
			Return _FilterLast
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double
		Get
			Return _FilterRate
		End Get
	End Property

	Public ReadOnly Property FilterVolatilityRate As Double
		Get
			Return _FilterVolatilityRate
		End Get
	End Property


	Public Overrides Function ToString() As String
		Return $"{Me.GetType().Name}({_FilterRate}): {_FilterLast.ProbabilityUp}, {_FilterLast.StateFlipProbability}"
	End Function

	Public ReadOnly Property IsReset As Boolean
		Get
			Return _IsReset
		End Get
	End Property

	Public Sub Reset()
		_IsReset = True
	End Sub

	Public Property Tag As String
End Class


