Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.MathPlus.Filter.FilterVolatility

''' <summary>
''' The FilterStatistical class calculates the mean, variance, and standard deviation of a series of values.
''' It uses a sliding window approach if a FilterRate is specified.
''' </summary>
<Serializable()>
Public Class FilterStatistical
	Implements IFilter(Of IStatistical)
	Implements IRegisterKey(Of String)


	Private MyRate As Integer
	Private FilterValueLast As IStatistical
	Private MyStartPoint As Integer
	Private IsRunReady As Boolean

	Private ValueLast As Double
	Private MyListOfValueStatistical As List(Of IStatistical)
	Private MyStatistic As IFilterRun(Of IStatistical)

	''' <summary>
	''' Calculate the statistical information from all value 
	''' </summary>
	''' <remarks></remarks>
	Public Sub New()
		Me.New(1)
	End Sub

	''' <summary>
	''' Calculate the statistical information based on a windows of size given by FilterRate
	''' </summary>
	''' <param name="FilterRate">
	''' Should be greater than one, otherwise the calculation is done on all the data
	''' </param>
	''' <remarks></remarks>
	Public Sub New(
		ByVal FilterRate As Integer,
		Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard,
		Optional BufferCapacity As Integer = 0)

		If FilterRate < 2 Then FilterRate = 2
		MyRate = CInt(FilterRate)

		Select Case StatisticType
			Case enuVolatilityStatisticType.Exponential
				MyStatistic = New StatisticExponential(FilterRate, BufferCapacity:=BufferCapacity)
			Case enuVolatilityStatisticType.Standard
				MyStatistic = New StatisticWindows(FilterRate, BufferCapacity:=BufferCapacity)
		End Select
		MyListOfValueStatistical = New List(Of IStatistical)
		FilterValueLast = New StatisticalData(0, 0, 0)
		ValueLast = 0
		MyStartPoint = 0
		IsRunReady = False
	End Sub

	''' <summary>
	''' Calculate the statistical information based on a windows of size given by FilterRate
	''' </summary>
	''' <param name="FilterRate"></param>
	''' Should be greater than one, otherwise the calculation is done on all the data
	''' <param name="StartPoint">
	''' THe startPoint wher teh calculation is started
	''' </param>
	''' The startpoin
	''' <remarks></remarks>
	Public Sub New(ByVal FilterRate As Integer, ByVal StartPoint As Integer)
		Me.New(FilterRate)
		MyStartPoint = StartPoint
		IsRunReady = True
	End Sub

	''' <summary>
	'''		The Filter function:
	''' - Initializes the filter with the first value and sets MyStartPoint to 0 if IsRunReady is False.
	''' - If IsRunReady is False and the new value is different from ValueLast, it sets MyStartPoint to the current count of MyListOfValueStatistical and marks IsRunReady as True.
	''' - Uses a sliding window approach if IsRunReady is True and the count of MyListOfValueStatistical is greater than or equal to MyStartPoint.
	''' - Removes the oldest value and adds the new value, updating the sum and mean accordingly if the count of MyListOfValue is greater than or equal to MyRate.
	''' - Calculates the variance using the sum of squared differences from the mean, applying Bessel's correction (MyRate - 1).
	''' - Sets the mean to the current value and variance to 0 if IsRunReady is False.
	''' - Handles edge cases, such as when the count of MyListOfValue is 0 or 1, correctly.
	''' 
	''' ''' Potential improvements:
	''' - Add additional comments explaining each step, especially the logic for handling the sliding window and variance calculation.
	''' - Ensure that the function handles edge cases, such as when the count of MyListOfValue is 0 or 1, correctly.
	''' - Optimize the function by using a circular buffer or deque for better performance.
	''' 
	''' References:
	''' http://en.wikipedia.org/wiki/Algorithms_for_calculating_variance
	''' https://en.wikipedia.org/wiki/Bessel%27s_correction
	''' 
	''' </summary>
	''' <param name="Value"></param>
	''' <returns></returns>
	Public Function Filter(Value As Double) As IStatistical Implements IFilter(Of IStatistical).Filter
		Dim ThisValueToRemove As Double
		Dim ThisValueSquareToRemove As Double
		Dim ThisM2 As Double
		Dim ThisMean As Double
		Dim ThisVariance As Double

		If MyListOfValueStatistical.Count = 0 Then
			'initialization
			FilterValueLast = New StatisticalData(Value, 0, 1)
			If IsRunReady = False Then
				MyStartPoint = 0
			End If
		Else
			If IsRunReady = False Then
				If Value <> ValueLast Then
					MyStartPoint = MyListOfValueStatistical.Count
					IsRunReady = True
				End If
			End If
		End If
		If IsRunReady Then
			If MyListOfValueStatistical.Count >= MyStartPoint Then
				FilterValueLast = MyStatistic.FilterRun(Value)
			Else
				FilterValueLast = New StatisticalData(Mean:=Value, Variance:=0, NumberPoint:=1, ValueLast:=Value)
			End If
		Else
			FilterValueLast = New StatisticalData(Mean:=Value, Variance:=0, NumberPoint:=1, ValueLast:=Value)
		End If
		MyListOfValueStatistical.Add(FilterValueLast)
		ValueLast = Value
		Return FilterValueLast
	End Function

	Public Function Filter(Value As IPriceVol) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.Filter(Value.Last)
	End Function

	Public Function Filter(ByRef Value() As Double) As IStatistical() Implements IFilter(Of IStatistical).Filter

		For Each ThisValue In Value
			Me.Filter(ThisValue)
		Next
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As IStatistical() Implements IFilter(Of IStatistical).Filter
		Throw New NotImplementedException(message:="Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) is not implemented in FilterStatistical class.")
	End Function

	Public Function FilterErrorLast() As IStatistical Implements IFilter(Of IStatistical).FilterErrorLast
		Throw New NotImplementedException(message:=" FilterErrorLast is not implemented in FilterStatistical class.")
	End Function

	Public Function FilterBackTo(ByRef Value As IStatistical) As Double Implements IFilter(Of IStatistical).FilterBackTo
		Throw New NotImplementedException(message:="FilterBackTo is not implemented in FilterStatistical class.")
	End Function

	Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).FilterLastToPriceVol
		Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast.Mean))
		With ThisPriceVol
			.LastPrevious = CSng(FilterValueLast.Mean)
			.Open = CSng(FilterValueLast.Mean)
			.Last = CSng(FilterValueLast.Mean)
			.High = CSng(FilterValueLast.Mean + FilterValueLast.StandardDeviation)
			.Low = CSng(FilterValueLast.Mean - FilterValueLast.StandardDeviation)
			.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
		End With
		Return ThisPriceVol
	End Function

	Private Function LastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).LastToPriceVol
		Throw New NotSupportedException(message:="LastToPriceVol is not implemented in FilterStatistical class.")
	End Function

	Public Function Filter(Value As Single) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.Filter(CDbl(Value))
	End Function

	Private Function FilterPredictionNext(ByVal Value As Double) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Throw New NotSupportedException(message:="FilterPredictionNext is not implemented in FilterStatistical class.")
	End Function

	Private Function FilterPredictionNext(ByVal Value As Single) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Return Me.FilterPredictionNext(CDbl(Value))
	End Function

	Public Function FilterLast() As IStatistical Implements IFilter(Of IStatistical).FilterLast
		Return FilterValueLast
	End Function

	Public Function Last() As Double Implements IFilter(Of IStatistical).Last
		Return ValueLast
	End Function

	Public ReadOnly Property Rate As Integer Implements IFilter(Of IStatistical).Rate
		Get
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilter(Of IStatistical).Count
		Get
			Return MyListOfValueStatistical.Count
		End Get
	End Property

	Private ReadOnly Property Max As Double Implements IFilter(Of IStatistical).Max
		Get
			Throw New NotSupportedException(message:="Max is not supported in FilterStatistical class.")
		End Get
	End Property

	Private ReadOnly Property Min As Double Implements IFilter(Of IStatistical).Min
		Get
			Throw New NotSupportedException(message:="Min is not supported in FilterStatistical class.")
		End Get
	End Property

	Public ReadOnly Property ToList() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToList
		Get
			Return MyListOfValueStatistical
		End Get
	End Property

	Private ReadOnly Property ToListOfError() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToListOfError
		Get
			Throw New NotSupportedException(message:="ToListOfError is not supported in FilterStatistical class.")
		End Get
	End Property

	Private ReadOnly Property ToListScaled() As ListScaled Implements IFilter(Of IStatistical).ToListScaled
		Get
			Throw New NotSupportedException(message:="ToListScaled is not supported in FilterStatistical class.")
		End Get
	End Property

	Public Function ToArray() As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Property Tag As String Implements IFilter(Of IStatistical).Tag

	Public Overrides Function ToString() As String Implements IFilter(Of IStatistical).ToString
		Return Me.FilterLast.ToString
	End Function

#Region "IRegisterKey"
	Public Function AsIRegisterKey() As IRegisterKey(Of String)
		Return Me
	End Function
	Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
	Dim MyKeyValue As String
	Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
		Get
			Return MyKeyValue
		End Get
		Set(value As String)
			MyKeyValue = value
		End Set
	End Property
#End Region
End Class
'End Namespace
