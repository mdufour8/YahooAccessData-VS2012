Imports YahooAccessData.MathPlus.Filter
Imports System.Numerics

Public Class FilterExp
	Implements IFilterRun

	Implements IFilter
	Implements IFilterState

	Private MyRate As Integer
	Private MyFilterRate As Double
	Private A As Double
	Private B As Double
	Private FilterValueLastK1 As Double
	Private FilterValueLast As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private IsReset As Boolean
	Private MyCircularBuffer As CircularBuffer(Of Double)

	Public Sub New(ByVal FilterRate As Double, Optional BufferCapacity As Integer = 0)

		If FilterRate < 1 Then FilterRate = 1
		MyFilterRate = FilterRate
		MyRate = CInt(MyFilterRate)

		' The factor A is calculated to provide the same bandwidth as a moving average with a flat window of FilterRate points.
		' For more details, see the "Comparison with moving average" section in:
		' https://en.wikipedia.org/wiki/Exponential_smoothing.
		' 
		' This result is based on the fact that the delay for a square window moving average is given by (N+1)/2,
		' while for an exponential filter, it is given by 1/Alpha (A). 
		' Note that this is the Noise Equivalent Bandwidth (NEB), not the 3 dB bandwidth. 
		' The NEB represents the bandwidth of a brick-wall filter that allows the same amount of noise to pass through
		' as the exponential filter in question. For further explanation, refer to the document:
		' 'IIR_Filter_NEB_vs_3dB_Bandwidth_FINAL.pdf' included in this directory.

		' The following formula originates from the technical comparison between exponential smoothing and standard moving averages:
		' - Both exponential smoothing and moving averages introduce a lag relative to the input data.
		' - For symmetrical kernels (e.g., moving averages or Gaussian filters), this lag can be corrected by shifting the result
		'   by half the window length. However, it is unclear how appropriate this correction would be for exponential smoothing.
		' - Both methods have a similar distribution of forecast error when Alpha (A) = 2 / (k + 1), where k is the number of past
		'   data points considered in the moving average.
		' - The key difference is that exponential smoothing considers all past data, while moving averages only consider the last k points.
		' - Computationally, exponential smoothing is more efficient, as it only requires the most recent forecast value to be stored,
		'   whereas moving averages require storing the past k data points or the data point at lag k+1.

		' The term A is not directly related to the 3 dB bandwidth or the Noise Equivalent Bandwidth (NEB).
		' Instead, it corresponds to the equivalent number of points that produce a similar result statistically to a flat moving average
		' with a fixed number of points. The formula for A is:
		' A = 2 / (FilterRate + 1)
		' For example, when FilterRate = 10, A = 0.1818.

		' However, given Alpha (A), the 3 dB bandwidth can be approximately calculated as:
		' f_3dB = (Sampling Frequency * ln(1 - A)) / (2 * PI),
		' where A is the factor for the previous value.
		' The 3 dB bandwidth is the frequency at which the power of the output signal is half that of the input signal.
		'Note it's value can also be calculated exactly as:
		'ω_3dB = ArcCos((A^2+2A-2)/2(A-1))t


		' The equivalent Noise Bandwidth (NEB) is given by:
		' ω_neb = PI * (A / (2 - A)). for largefilterRate > this ratio ω_neb/ω_3dB is ~PI/2.

		' Another important aspect is the time or number of samples required to reach 63% of the final value for a unit step input.
		' This is known as the Time Constant and is calculated as:
		' Time Constant = Sampling Period / ln(1 - A).
		' To estimate the number of samples required to reach 95% of the final value, use the formula:
		' Number of Samples = 3 * Time Constant.

		' For additional context, see: https://en.wikipedia.org/wiki/Low-pass_filter.
		' Note that B = 1 - A is the factor applied to the previous value.
		A = 2 / (FilterRate + 1)
		B = 1 - A
		FilterValueLast = 0
		FilterValueLastK1 = 0
		ValueLast = 0
		ValueLastK1 = 0
		MyCircularBuffer = New CircularBuffer(Of Double)(capacity:=BufferCapacity, 0.0)
		IsReset = True
	End Sub

	''' <summary>
	''' Return the 3 dB bandwidth for a given filter rate in rad/sec.
	''' The 3 dB bandwidth is the frequency at which the power of the output signal is half that of the input signal.
	''' The formula used is:
	''' for α=2/(FilterRate+1)
	''' where α is the filetring factor for the most recent imput value to the filter and
	''' ωc3dB = log(N+1/N-1)
	''' </summary>
	''' <param name="FilterRate"></param>
	''' <returns>The 3dB cut-off frequency for a given number of samples. The unit is in Radian/Sample </returns>
	Shared Function GetFrequency3dB(FilterRate As Double) As Double
		' This function calculates the 3 dB bandwidth for a given filter rate.
		Return Math.Log((FilterRate + 1) / (FilterRate - 1))
	End Function

	''' <summary>
	''' Get the response time for a given filter rate.
	''' The response time is the number of samples for the filter to reach 95% of it final final when subject to a Unit Step signal input.
	''' The formula used is:
	''' α=2/(FilterRate+1)
	''' t = -ln(1 - 0.95) / ln(1 - α)
	''' where α is the fitering factor for the most recent imput value to the filter.
	''' t = -ln(1 - 0.95) / ln(1 - α)
	''' Note that another relation is given approximativly by fc=2/(Response Time at 95%</summary>
	''' <param name="FilterRate"></param>
	''' <returns>The response time for a given filter rate. The unit is in Sample </returns>
	Shared Function GetResponseTime(FilterRate As Double) As Double
		Dim _α As Double = 2 / (FilterRate + 1)
		Return Math.Log(1 - 0.95) / Math.Log(1 - _α)
	End Function

	''' <summary>
	''' Get the time constant for a given filter rate.
	''' The time constant is the time required for the filter to reach approximately 63% of its final value when subject to a unit step input.
	''' The formula is simply for teh first order filter 1/ωc3dB:
	''' Also for α=2/(FilterRate+1)
	''' t = -ln(1 - 0.63) / ln(1 - α)
	''' where α is the fitering factor for the most recent imput value to the filter.
	''' Note that another relation is given approximativly by fc=2/(Response Time at 95%</summary>
	''' <param name="FilterRate"></param>
	''' <returns>The time to reach the 63% of the response in number of points</returns>
	Shared Function GetTimeConstant(FilterRate As Double) As Double
		Return 1 / GetFrequency3dB(FilterRate)
	End Function

	''' <summary>
	''' Calculates the transfer function for the current filter
	''' </summary>
	''' <param name="DigitalFrequencyΩ">
	''' The digital frequency for the transfer function in radian between [0,π]. In some case you may need to apply 
	''' a factor of 2*PI to convert from cycle/sample to radian/sample.
	''' The digital frequency is related to the sampling frequency and the frequency of the input signal.
	''' The formula used is:
	''' DigitalFrequencyΩ = 2 * PI * Frequency / SamplingFrequency
	''' where Frequency is the frequency of the input signal and SamplingFrequency is the sampling frequency.
	''' </param>
	''' <returns>A named tuple containing the amplitude and phase (in radians).</returns>
	Function GetTransfertFunction(DigitalFrequencyΩ As Double) As Complex
		' Example calculation for amplitude and phase
		Dim alpha As Double = 2 / (FilterRate + 1)
		Dim amplitude As Double = 1 / Math.Sqrt(1 + alpha ^ 2) ' Example formula for amplitude
		Dim phase As Double = -Math.Atan(alpha) ' Example formula for phase (in radians)

		Throw New NotImplementedException("The function GetTransfertFunction is not implemented yet.")
		' Return the amplitude and phase as a named tuple
		'Return (Amplitude:=amplitude, Phase:=phase)
	End Function


	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		If IsReset Then
			'initialization
			FilterValueLast = Value
			MyCircularBuffer.Clear()
			IsReset = False
		End If
		FilterValueLastK1 = FilterValueLast
		FilterValueLast = A * Value + B * FilterValueLast
		ValueLastK1 = ValueLast
		ValueLast = Value
		MyCircularBuffer.AddLast(FilterValueLast)
		Return FilterValueLast
	End Function
	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return FilterValueLast
		End Get
	End Property

	''' <summary>
	''' Return the last value of the filter at the given index.
	''' The index here is the number of samples from the last value. 0 is the most recent value.
	''' The index is in the range [0, FilterRate-1].
	''' </summary>
	''' <remarks>
	''' The index is in the range [0, FilterRate-1].
	''' </remarks>
	''' <param name="Index">range [0, FilterRate-1]</param>
	''' <returns>The filter value ay the Index value </returns>
	Public ReadOnly Property FilterLast(Index As Integer) As Double Implements IFilterRun.FilterLast
		Get
			'For the CircularBuffer note 0 is the oldest value MyCircularBuffer.Count -1 is
			'the most recent value.
			'The index is in the range [0, FilterRate-1].
			Dim ThisBufferIndex As Integer = MyCircularBuffer.Count - 1 - Index
			Select Case ThisBufferIndex
				Case < 0
					'return the oldest value
					Return MyCircularBuffer.PeekFirst
				Case >= MyCircularBuffer.Count
					'return the last value (most recent value)
					Return MyCircularBuffer.PeekLast
				Case Else
					'return at a sppecific location in the buffer	
					Return MyCircularBuffer.Item(BufferIndex:=ThisBufferIndex)
			End Select
		End Get
	End Property

	Public ReadOnly Property FilterRate() As Double Implements IFilterRun.FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		IsReset = True
	End Sub

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyFilterRate})"
		End Get
	End Property

#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		Try
			If MyQueueForState.Count = 0 Then Return
			ValueLast = MyQueueForState.Dequeue
			ValueLastK1 = MyQueueForState.Dequeue
			FilterValueLast = MyQueueForState.Dequeue
			FilterValueLastK1 = MyQueueForState.Dequeue
		Catch ex As InvalidOperationException
			' Handle error, perhaps log it or rethrow with additional info
			Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
		End Try
	End Sub

	Private Sub IFilterState_Save() Implements IFilterState.Save
		MyQueueForState.Enqueue(ValueLast)
		MyQueueForState.Enqueue(ValueLastK1)
		MyQueueForState.Enqueue(FilterValueLast)
		MyQueueForState.Enqueue(FilterValueLastK1)
	End Sub
#End Region

#Region "IFilter"
	Private ReadOnly Property IFilter_Rate As Integer Implements IFilter.Rate
		Get
			Return CInt(MyFilterRate)
		End Get
	End Property

	Private ReadOnly Property IFilter_Count As Integer Implements IFilter.Count
		Get
			Return MyCircularBuffer.Count
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
			Return MyCircularBuffer.ToList()
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

	Private ReadOnly Property IFilterRun_IsReset As Boolean Implements IFilterRun.IsReset
		Get
			Return IsReset
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Throw New NotImplementedException()
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
		Return Me.ValueLast
	End Function

	Private Function IFilter_ToArray() As Double() Implements IFilter.ToArray
		Return MyCircularBuffer.ToArray
	End Function

	Private Function IFilter_ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Public Overrides Function ToString() As String Implements IFilter.ToString
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
	End Function
#End Region
End Class

