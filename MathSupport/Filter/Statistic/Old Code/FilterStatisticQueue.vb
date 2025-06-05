Imports System.Threading.Tasks
Imports YahooAccessData.MathPlus.Filter

''' <summary>
''' The FilterStatistical class calculates the mean, variance, and standard deviation of a series of values.
''' It uses a sliding window approach with the specified size given by by the FilterRate. 
''' 
''' </summary>
<Serializable()>
Public Class FilterStatisticalQueue
	Implements IFilterRun(Of IStatistical)
	Implements IFilter(Of IStatistical)
	Implements IRegisterKey(Of String)


	Private MyListOfValueStatistical As List(Of IStatistical)
	Private MyStatQueue As IFilterRun(Of IStatistical)

#Region "New"
	''' <summary>
	''' Calculate the statistical information based on a windows size given by FilterRate. 
	''' This is a basic implementation for the statistical measurement using the square windows only.
	''' It does not use any exponential filtering and follow the usual statistic standard method.
	''' </summary>
	''' <param name="FilterRate">
	''' Should be set to be greater than two
	''' </param>
	''' <remarks></remarks>
	Public Sub New(ByVal FilterRate As Integer, Optional BufferCapacity As Integer = 0)
		MyStatQueue = New StatisticWindows(FilterRate:=FilterRate, BufferCapacity:=BufferCapacity)
		MyListOfValueStatistical = New List(Of IStatistical)
	End Sub
#End Region

#Region "IFilter"
	''' <summary>
	''' The Filter function:
	''' - Initializes the filter with the first value and sets MyStartPoint to 0 if IsRunReady is False.
	''' - If IsRunReady is False and the new value is different from ValueLast, it sets MyStartPoint to the current count of MyListOfValueStatistical and marks IsRunReady as True.
	''' - Uses a sliding window approach if IsRunReady is True and the count of MyListOfValueStatistical is greater than or equal to MyStartPoint.
	''' - Removes the oldest value and adds the new value, updating the sum and mean accordingly if the count of MyQueueOfValue is greater than or equal to MyRate.
	''' - Calculates the variance using the sum of squared differences from the mean, applying Bessel's correction (MyRate - 1).
	''' - Sets the mean to the current value and variance to 0 if IsRunReady is False.
	''' - Handles edge cases, such as when the count of MyQueueOfValue is 0 or 1, correctly.
	''' 
	''' References:
	''' http://en.wikipedia.org/wiki/Algorithms_for_calculating_variance
	''' https://en.wikipedia.org/wiki/Bessel%27s_correction
	''' 
	''' </summary>
	''' <param name="Value"></param>
	''' <returns></returns>
	Public Function Filter(Value As Double) As IStatistical Implements IFilter(Of IStatistical).Filter

		If MyListOfValueStatistical.Count = 0 OrElse _IsReset Then
			'initialization
			MyListOfValueStatistical.Clear()
			MyCircularBuffer.Clear()
			ValueLast = Value
			IsRunReady = False
			MyStartPoint = 0
			FilterValueLast = New StatisticalData(Mean:=Value, Variance:=0, NumberPoint:=1)
			IsRunReady = False
		Else
			If IsRunReady = False Then
				'wait until the value change to start the volatility measurement filter
				'this is to avoid initial value of zero to be used in the calculation
				If Value <> ValueLast Then
					MyStartPoint = MyListOfValueStatistical.Count
					IsRunReady = True
				End If
			End If
		End If
		FilterValueLastK1 = FilterValueLast.Copy
		If IsRunReady Then
			ThisMean = MyFilterExpMean.FilterRun(Value)
			ThisM2 = (Value - ThisMean) ^ 2
			ThisVariance = MyVarianceCorrection * MyFilterExpSquare.FilterRun(ThisM2)
		Else
			ThisMean = Value
			ThisVariance = 0
		End If
		FilterValueLast = New StatisticalData(ThisMean, ThisVariance, MyListOfValueStatistical.Count + 1)
		MyListOfValueStatistical.Add(FilterValueLast)
		ValueLastK1 = ValueLast
		ValueLast = Value
		MyCircularBuffer.AddLast(FilterValueLast)
		Return FilterValueLast
	End Function

	Public Function Filter(Value As IPriceVol) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.FilterRun(Value:=Value.Last)
	End Function

	''' <summary>
	''' This function fix the statistical value at the beginning to reflect the result including the data pass the index
	''' up to the rate value
	''' </summary>
	''' <param name="Value"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function Filter(ByRef Value() As Double) As IStatistical() Implements IFilter(Of IStatistical).Filter
		Dim ThisListOfValueStatistical As List(Of IStatistical)
		Dim ThisValue As Double
		Dim ThisPosition As Integer
		Dim ThisPoint As Integer
		Dim I As Integer

		MyListOfValueStatistical.Clear()
		MyStatQueue.Reset()
		For Each ThisValue In Value
			Me.Filter(ThisValue)
		Next
		Return MyListOfValueStatistical.ToArray
	End Function

	''' <summary>
	''' Special filtering that can be used to remove the delay starting at a specific point
	''' </summary>
	''' <param name="Value">The value to be filtered</param>
	''' <param name="DelayRemovedToItem">The point where the delay start to be removed</param>
	''' <returns>The result</returns>
	''' <remarks></remarks>
	Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As IStatistical() Implements IFilter(Of IStatistical).Filter
		Throw New NotImplementedException
	End Function

	Public Function FilterErrorLast() As IStatistical Implements IFilter(Of IStatistical).FilterErrorLast
		Throw New NotImplementedException
	End Function

	Public Function FilterBackTo(ByRef Value As IStatistical) As Double Implements IFilter(Of IStatistical).FilterBackTo
		Throw New NotImplementedException
	End Function

	Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).FilterLastToPriceVol
		Throw New NotSupportedException(message:="FilterLastToPriceVol is not supported in FilterStatisticalQueue.")
	End Function

	Public Function LastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).LastToPriceVol
		Throw New NotSupportedException
	End Function

	Public Function Filter(Value As Single) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.FilterRun(CDbl(Value))
	End Function

	Public Function FilterPredictionNext(ByVal Value As Double) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Throw New NotSupportedException
	End Function

	Public Function FilterPredictionNext(ByVal Value As Single) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Return Me.FilterPredictionNext(CDbl(Value))
	End Function

	Public Function FilterLast() As IStatistical Implements IFilter(Of IStatistical).FilterLast
		Return MyStatQueue.FilterLast
	End Function

	Public Function Last() As Double Implements IFilter(Of IStatistical).Last
		Return MyStatQueue.InputLast
	End Function

	Public ReadOnly Property Rate As Integer Implements IFilter(Of IStatistical).Rate
		Get
			Return CInt(MyStatQueue.FilterRate)
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilter(Of IStatistical).Count
		Get
			Return MyListOfValueStatistical.Count
		End Get
	End Property

	Public ReadOnly Property Max As Double Implements IFilter(Of IStatistical).Max
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property Min As Double Implements IFilter(Of IStatistical).Min
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property ToList() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToList
		Get
			Return MyListOfValueStatistical
		End Get
	End Property

	Public ReadOnly Property ToListOfError() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToListOfError
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter(Of IStatistical).ToListScaled
		Get
			Throw New NotSupportedException
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
#End Region

#Region "IFilterRun"
	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		MyListOfValueStatistical.Add(MyStatQueue.FilterRun(Value:=Value))
		Return MyStatQueue.FilterLast
	End Function

	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return MyStatQueue.InputLast
		End Get
	End Property

	Public ReadOnly Property IFilterRun_FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyStatQueue.FilterLast
		End Get
	End Property

	Public ReadOnly Property IFilterRun_FilterLast(Index As Integer) As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyStatQueue.FilterLast(Index:=Index)
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As IStatistical Implements IFilterRun(Of IStatistical).FilterTrendLast
		Get
			Return New StatisticalData(0, 0, 0)
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun(Of IStatistical).FilterRate
		Get
			Return Me.Rate
		End Get
	End Property

	Public ReadOnly Property ToBufferList As IList(Of IStatistical) Implements IFilterRun(Of IStatistical).ToBufferList
		Get
			Return MyStatQueue.ToBufferList
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun(Of IStatistical).FilterDetails
		Get
			Return $"{Me.GetType().Name}({Me.Rate})"
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun(Of IStatistical).Reset
		MyStatQueue.Reset()
	End Sub

	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun(Of IStatistical).Reset
		MyStatQueue.Reset(BufferCapacity:=BufferCapacity)
	End Sub
	Public ReadOnly Property IsReset As Boolean Implements IFilterRun(Of IStatistical).IsReset
		Get
			Return MyStatQueue.IsReset
		End Get
	End Property
#End Region
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

