
Imports Newtonsoft.Json.Linq
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter

''' <summary>
''' The FilterPLL class implements a Phase-Locked Loop (PLL) filter that processes an input signal to generate a filtered output.
''' The filter uses a set of coefficients and error terms to adjust the output based on the input signal.
''' The class provides methods to run the filter and reset it, as well as properties to access the current state of the filter.
''' </summary>
Public Class FilterPLL
	Implements IFilterRun
	Implements IFilter
	Implements IFilterState

	Private C As Double
	Private C1 As Double
	Private C2 As Double

	Private MyVCOPeriod As Double

	Private MyErrork0 As Double
	Private MyErrork1 As Double
	Private MyVCOk0 As Double
	Private MyVCOk1 As Double
	Private MyVCOk2 As Double
	Private MyRefValue As Double
	Private MyFilterValueLast As Double
	Private MyFilterValuek0 As Double
	Private MyFilterValuek1 As Double
	Private MyFilterValuek2 As Double
	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MySignalDelta As Double
	Private MyFilterError As Double
	Private MyFilterRate As Double
	Private MyDampingFactor As Double
	Private IsReset As Boolean
	'note we cannot create another PLL filter without causing stack overflow
	Private MyFilterForError As FilterExp

	Public Sub New(ByVal FilterRate As Double, Optional DampingFactor As Double = 1.0)
		Dim ThisFilterRateMinimum As Double
		Dim FreqDigital As Double
		Dim SamplingPeriod As Double
		Dim FrequencyNaturel As Double

		MyFilterRate = FilterRate
		MyDampingFactor = DampingFactor
		ThisFilterRateMinimum = 5 * MyDampingFactor
		'If MyFilterRate < ThisFilterRateMinimum Then
		'	MyFilterRate = ThisFilterRateMinimum
		'End If
		MyVCOPeriod = 0   'by default in this application
		MyErrork1 = 0
		MyFilterValuek2 = 0

		'original code
		'Let explain this relation:
		'FreqDigital = 1 / (PI * MyFilterRate)
		'Please refer to the article 'DigitalFrequency_vs_AngularFrequency.pdf' for more information.
		'In this article the relation bwetween the 3dB bandwidth and the filter resonance frequncy is explained.
		'It is noted that a formal relation between the 3dB bandwidth and the filter resonance frequency is not given.	
		'However it can be approxmated with simple relation of 2.0.  
		'The user provide a filter rate parameters that is effectivly the desired 3dB bandwidth of the filter.
		'The filter rate is the 3dB bandwidth of the filter but the PLL filter is described by it resonnance frequency:
		'Fn=Fc/2.0 = 2*PI/2.0*FilterRate)=PI/FilterRate	

		'Note also that the digital frequency Ω is the angular frequency ω (rad/s) times the sampling period T (s/sample),
		'so the units digital frequency are rad/sample.

		'Ω=ω*T	

		'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
		'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
		'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
		'A = CDbl((2 / (MyFilterRate + 1)))

		'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
		'B = 1 - A


		FreqDigital = 1 / (Math.PI * MyFilterRate)
		C = MyVCOPeriod
		C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital

		'see DPLL_Detailed_Frequency_Response_Analysis_1.pdf for explanation
		FreqDigital = 1.0  '(Sampling rate Is 1.0 day by Default))
		SamplingPeriod = 1 / FreqDigital
		FrequencyNaturel = 1 / MyFilterRate
		'C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital
		C2 = 2 * MyDampingFactor * (2 * Math.PI * FrequencyNaturel) * SamplingPeriod / Math.PI


		C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
		'check stability
		If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
			'loop is unstable
			Throw New Exception("Low pass PLL filter is not stable...")
		End If
		MySignalDelta = 0
		MyErrork0 = 0
		MyFilterForError = New FilterExp(FilterRate:=MyFilterRate)
		IsReset = True
	End Sub

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		If IsReset Then
			'initialization
			'initialize the loop with the first time sample
			'this is to minimize the PLL tracking error
			MyFilterValuek0 = Value
			MyFilterValueLast = MyFilterValuek0
			'MyRefValue is the first reference sample value and never change annymore
			MyRefValue = Value
			MyVCOk0 = Value
			MyFilterForError.Reset()
			IsReset = False
		End If
		'just the standard PLL here
		MySignalDelta = Value - MyFilterValuek0
		'calculate the filter loop parameters
		MyErrork1 = MyErrork0
		MyErrork0 = (C1 * MySignalDelta) + MyErrork1
		MyFilterError = (C2 * MySignalDelta) + MyErrork1
		MyFilterForError.FilterRun(MyFilterError)
		'calculate the integrator parameters
		MyVCOk2 = MyVCOk1
		MyVCOk1 = MyVCOk0
		MyVCOk0 = C + MyFilterError + MyVCOk1
		'in this implementation MyFilterError is the instantaneous slope of the signal output
		MyFilterValuek2 = MyFilterValuek1
		MyFilterValuek1 = MyFilterValuek0

		'MyFilterValuek0 is the next signal predicted value for the next input sample
		MyFilterValuek0 = MyVCOk0
		MyFilterValueLast = MyFilterValuek0
		ValueLast = Value
		Return MyFilterValueLast
	End Function

	Public Function FilterRun(Value As Double, FilterPLLDetector As IFilterPLLDetector) As Double Implements IFilterRun.FilterRun

		If FilterPLLDetector Is Nothing Then
			Throw New InvalidConstraintException("FilterPLLDetector is not initialized...")
		End If
		If IsReset Then
			'initialization
			'initialize the loop with the first time sample
			'this is to minimize the PLL tracking error
			MyFilterValuek0 = FilterPLLDetector.ValueInit
			'MyRefValue is the first reference sample value and never change annymore
			MyRefValue = MyFilterValuek0
			MyFilterValueLast = FilterPLLDetector.ValueOutput(Value, MyFilterValuek0)
			MyVCOk0 = 0
			IsReset = False
		End If
		Dim ThisNumberLoop As Integer = 0
		Dim ThisValueStart As Double = FilterPLLDetector.ValueOutput(Value, MyFilterValuek0)
		Dim ThisValueStop As Double
		Do
			'If Me.Count >= 100 And Me.Tag = "PriceFilterLowPassPLL" Then
			'	If ThisFilterPredictionGainYearly = 0.0 Then
			'		ThisFilterPredictionGainYearly = ThisFilterPredictionGainYearly
			'	End If
			'End If

			ThisNumberLoop = ThisNumberLoop + 1
			MySignalDelta = FilterPLLDetector.RunErrorDetector(Value, MyFilterValuek0)
			'ignore the error and hold the output if the status is false
			If FilterPLLDetector.Status = False Then Exit Do
			'calculate the filter loop parameters
			MyErrork1 = MyErrork0
			MyErrork0 = (C1 * MySignalDelta) + MyErrork1
			MyFilterError = (C2 * MySignalDelta) + MyErrork1
			'calculate the integrator parameters
			MyVCOk2 = MyVCOk1
			MyVCOk1 = MyVCOk0
			MyVCOk0 = C + MyFilterError + MyVCOk1
			If FilterPLLDetector.IsMaximum Then
				MyVCOk0 = WaveForm.SignalLimit(MyVCOk0, FilterPLLDetector.Maximum)
				'If MyVCOk0 > FilterPLLDetector.Maximum Then
				'Exit Do
				'MyVCOk0 = FilterPLLDetector.Maximum

				'End If
			End If
			If FilterPLLDetector.IsMinimum Then
				'If MyVCOk0 < FilterPLLDetector.Minimum Then
				MyVCOk0 = WaveForm.SignalLimit(MyVCOk0, -FilterPLLDetector.Minimum)
				'End If
			End If
			'in this implementation MyFilterError is the instantaneous slope of the signal output
			MyFilterValuek2 = MyFilterValuek1
			MyFilterValuek1 = MyFilterValuek0
			'MyFilterValuek0 is the next signal predicted value for the next sample
			MyFilterValuek0 = MyRefValue + MyVCOk0
			If FilterPLLDetector.ToErrorLimit > 0 Then
				If Math.Abs(MySignalDelta) < FilterPLLDetector.ToErrorLimit Then
					'exit but only if we run for at least the filter rate
					If ThisNumberLoop > Me.Rate Then
						Exit Do
					End If
				End If
			End If
		Loop Until ThisNumberLoop >= FilterPLLDetector.ToCount
		ThisValueStop = FilterPLLDetector.ValueOutput(Value, MyFilterValuek0)
		FilterPLLDetector.RunConvergence(ThisNumberLoop, ThisValueStart, ThisValueStop)
		MyFilterValueLast = ThisValueStop
		Return MyFilterValueLast
	End Function

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return MyFilterValueLast
		End Get
	End Property

	Public ReadOnly Property Rate() As Double Implements IFilterRun.FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		IsReset = True
	End Sub

	Public ReadOnly Property FilterValuek0 As Double
		Get
			Return MyFilterValuek0
		End Get
	End Property

	Public ReadOnly Property FilterValuek1 As Double
		Get
			Return MyFilterValuek1
		End Get
	End Property

	Public ReadOnly Property Errork0 As Double
		Get
			Return MyErrork0
		End Get
	End Property

	Public ReadOnly Property Errork1 As Double
		Get
			Return MyErrork1
		End Get
	End Property

	Public ReadOnly Property FilterError As Double
		Get
			Return MyFilterForError.FilterLast
		End Get
	End Property

	Public Overrides Function ToString() As String
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
	End Function

#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		Try
			If MyQueueForState.Count = 0 Then Return
			MyErrork0 = MyQueueForState.Dequeue
			MyErrork1 = MyQueueForState.Dequeue
			MyVCOk0 = MyQueueForState.Dequeue
			MyVCOk1 = MyQueueForState.Dequeue
			MyVCOk2 = MyQueueForState.Dequeue

			MyRefValue = MyQueueForState.Dequeue
			MyFilterValueLast = MyQueueForState.Dequeue
			MyFilterValuek0 = MyQueueForState.Dequeue
			MyFilterValuek1 = MyQueueForState.Dequeue
			MyFilterValuek2 = MyQueueForState.Dequeue
			ValueLast = MyQueueForState.Dequeue
			ValueLastK1 = MyQueueForState.Dequeue
			MySignalDelta = MyQueueForState.Dequeue
			MyFilterError = MyQueueForState.Dequeue
		Catch ex As InvalidOperationException
			' Handle error, perhaps log it or rethrow with additional info
			Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
		End Try
	End Sub

	Private Sub IFilterState_Save() Implements IFilterState.Save
		MyQueueForState.Enqueue(MyErrork0)
		MyQueueForState.Enqueue(MyErrork1)

		MyQueueForState.Enqueue(MyVCOk0)
		MyQueueForState.Enqueue(MyVCOk1)
		MyQueueForState.Enqueue(MyVCOk2)

		MyQueueForState.Enqueue(MyRefValue)
		MyQueueForState.Enqueue(MyFilterValueLast)
		MyQueueForState.Enqueue(MyFilterValuek0)
		MyQueueForState.Enqueue(MyFilterValuek1)
		MyQueueForState.Enqueue(MyFilterValuek2)
		MyQueueForState.Enqueue(ValueLast)
		MyQueueForState.Enqueue(ValueLastK1)
		MyQueueForState.Enqueue(MySignalDelta)
		MyQueueForState.Enqueue(MyFilterError)
	End Sub
#End Region

#Region "IFilter"
	Private ReadOnly Property IFilter_Rate As Integer Implements IFilter.Rate
		Get
			Return CInt(MyFilterRate)
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilter.Count
		Get
			Return 0
		End Get
	End Property

	Public ReadOnly Property Max As Double Implements IFilter.Max
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property Min As Double Implements IFilter.Min
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of Double) Implements IFilter.ToList
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property ToListOfError As IList(Of Double) Implements IFilter.ToListOfError
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property ToListScaled As ListScaled Implements IFilter.ToListScaled
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public Property Tag As String Implements IFilter.Tag
		Get
			Throw New NotImplementedException()
		End Get
		Set(value As String)
			Throw New NotImplementedException()
		End Set
	End Property

	Public Function Filter(Value As Double) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Private Function IFilter_Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Private Function Filter_IFilter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
		Throw New NotImplementedException()
	End Function

	Public Function Filter(Value As Single) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value)
	End Function

	Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value.Last)
	End Function

	Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
		Return Me.FilterError
	End Function

	Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
		Throw New NotImplementedException()
	End Function

	Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
		Throw New NotImplementedException()
	End Function

	Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
		Throw New NotImplementedException()
	End Function

	Public Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Public Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_FilterLast() As Double Implements IFilter.FilterLast
		Return Me.FilterLast
	End Function

	Public Function Last() As Double Implements IFilter.Last
		Throw New NotImplementedException()
	End Function

	Public Function ToArray() As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToString() As String Implements IFilter.ToString
		Return Me.ToString()
	End Function

#End Region
End Class
