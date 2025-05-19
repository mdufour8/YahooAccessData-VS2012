Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter

''' <summary>
''' Summary of 3 dB Bandwidth for Single vs. Double Exponential Filters
'''
''' A single exponential filter has a frequency response with a -3 dB cutoff frequency
''' approximately proportional to the smoothing factor:
'''     wc1 ≈ alpha      (in radians/sample)
'''
''' A double exponential filter (i.e., applying the filter twice) has a sharper roll-off.
''' Its -3 dB bandwidth is narrower and approximately given by:
'''     wc2 ≈ wc1 / sqrt(2) ≈ alpha / sqrt(2)
'''
''' Therefore, the ratio of the -3 dB bandwidths is:
'''     wc2 / wc1 ≈ 1 / sqrt(2) ≈ 0.707
'''
''' This behavior is similar to a second-order low-pass filter like a Butterworth filter.
''' It helps better capture trend while attenuating high-frequency noise.
'''
''' Example adjustment:
''' If you estimated alpha from a single-filter system and now use a double exponential filter,
''' Summary of 3 dB Bandwidth for Single vs. Double Exponential Filters
''' </summary>
Public Class FilterDoubleExp
	Implements IFilterRun
	Implements IFilter
	Implements IFilterState

	Private ReadOnly FirstFilter As FilterExp
	Private ReadOnly SecondFilter As FilterExp

	''' <summary>
	''' Initializes a new instance of the FilterDoubleExp class with the specified filter rate.
	''' An adjustment of 1/sqrt(2) should be applied by the user if the bandwith of the filter 
	''' need to be the same than an equivalent single pole filter.
	''' This adjustment is based on the behavior of double exponential filters.
	''' </summary>
	''' <param name="FilterRate">The filter rate for both internal filters.</param>
	Public Sub New(ByVal FilterRate As Double, Optional BufferCapacity As Integer = 0)
		' Create two instances of the single exponential filter
		FirstFilter = New FilterExp(FilterRate)
		SecondFilter = New FilterExp(FilterRate, BufferCapacity)
	End Sub

	''' <summary>
	''' Runs the double exponential filter by passing the value through two sequential filters.
	''' </summary>
	''' <param name="Value">The input value to filter.</param>
	''' <returns>The filtered value after two passes.</returns>
	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		' Pass the value through the first filter, then the second filter
		Dim firstPass As Double = FirstFilter.FilterRun(Value)
		Return SecondFilter.FilterRun(firstPass)
	End Function

	''' <summary>
	''' Gets the last filtered value from the second filter.
	''' </summary>
	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return SecondFilter.FilterLast
		End Get
	End Property

	Public ReadOnly Property FilterLast(Item As Integer) As Double Implements IFilterRun.FilterLast
		Get
			Return SecondFilter.FilterLast(Item)
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	''' <summary>
	''' Gets the filter rate used by the filters.
	''' </summary>
	Public ReadOnly Property FilterRate As Double Implements IFilterRun.FilterRate
		Get
			Return FirstFilter.FilterRate
		End Get
	End Property

	''' <summary>
	''' Resets both internal filters.
	''' </summary>
	Public Sub Reset() Implements IFilterRun.Reset
		FirstFilter.Reset()
		SecondFilter.Reset()
	End Sub

	''' <summary>
	''' Gets the details of the filter, including the filter rate.
	''' </summary>
	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({FirstFilter.FilterRate})"
		End Get
	End Property

#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		FirstFilter.ASIFilterState.ReturnPrevious()
		SecondFilter.ASIFilterState.ReturnPrevious()
	End Sub

	Private Sub IFilterState_Save() Implements IFilterState.Save
		FirstFilter.ASIFilterState.Save()
		SecondFilter.ASIFilterState.Save()
	End Sub
#End Region

#Region "IFilter"
	Private ReadOnly Property IFilter_Rate As Integer Implements IFilter.Rate
		Get
			Return DirectCast(SecondFilter, IFilter).Rate
		End Get
	End Property

	Private ReadOnly Property IFilter_Count As Integer Implements IFilter.Count
		Get
			Return DirectCast(SecondFilter, IFilter).Count
		End Get
	End Property

	Private ReadOnly Property IFilter_Max As Double Implements IFilter.Max
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_Min As Double Implements IFilter.Min
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_ToList As IList(Of Double) Implements IFilter.ToList
		Get
			Return DirectCast(SecondFilter, IFilter).ToList
		End Get
	End Property

	Private ReadOnly Property IFilter_ToListOfError As IList(Of Double) Implements IFilter.ToListOfError
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private ReadOnly Property IFilter_ToListScaled As ListScaled Implements IFilter.ToListScaled
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Private Property IFilter_Tag As String Implements IFilter.Tag

	Public ReadOnly Property IsReset As Boolean Implements IFilterRun.IsReset
		Get
			Return IsReset
		End Get
	End Property

	Private Function IFilter_Filter(Value As Double) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Private Function IFilter_Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_Filter(Value As Single) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Private Function IFilter_Filter(Value As IPriceVol) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value.Last)
	End Function

	Private Function IFilter_FilterErrorLast() As Double Implements IFilter.FilterErrorLast
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterLast() As Double Implements IFilter.FilterLast
		Return Me.FilterLast
	End Function

	Private Function IFilter_Last() As Double Implements IFilter.Last
		Return DirectCast(FirstFilter, IFilter).Last
	End Function

	Private Function IFilter_ToArray() As Double() Implements IFilter.ToArray
		Return DirectCast(SecondFilter, IFilter).ToArray
	End Function

	Private Function IFilter_ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Public Overrides Function ToString() As String Implements IFilter.ToString
		Return $"{Me.GetType().Name}: FilterRate={IFilter_Rate},{Me.FilterLast}"
	End Function
#End Region
End Class

