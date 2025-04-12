
Imports System.Math
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
	Private MyFilterForError As FilterExpPredict

	Public Sub New(ByVal FilterRate As Double, Optional DampingFactor As Double = 1.0)
		Dim ThisFilterRateMinimum As Double
		Dim FreqDigital As Double
		Dim SamplingPeriod As Double
		Dim FrequencyNaturel As Double
		'Here using special character to better understand the code that follow
		Dim This_Ts As Double = 1.0    'Sampling rate Is 1 day by Default in all that base code 
		Dim This_fₙ As Double
		Dim This_ωₙ As Double
		Dim This_Ω As Double

		'the ADPLL need the natural frequency of the filter. 

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
		FreqDigital = 1 / (Math.PI * MyFilterRate)  'Digital Frequency: Ω = ω × T
		C = MyVCOPeriod
		C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital

		'see DPLL_Detailed_Frequency_Response_Analysis_1.pdf for explanation
		'FreqDigital = 1.0  '(Sampling rate Is 1.0 Sample/day by Default))
		'SamplingPeriod = 1 / FreqDigital
		'FrequencyNaturel = 1 / MyFilterRate
		''C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital
		'C2 = 2 * MyDampingFactor * (2 * Math.PI * FrequencyNaturel) * SamplingPeriod / Math.PI

		C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
		'check stability
		If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
			'loop is unstable
			Throw New Exception("Low pass PLL filter is not stable...")
		End If
		MySignalDelta = 0
		MyErrork0 = 0
		MyFilterForError = New FilterExpPredict(FilterRate:=MyFilterRate)
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

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyFilterRate},{MyDampingFactor})"
		End Get
	End Property

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

	''' <summary>
	''' Simulates a unit step input to the FilterPLL and prints analysis to console.
	''' </summary>
	Public Shared Sub TestUnitStepResponse()
		Dim pll As New FilterPLL(FilterRate:=20, DampingFactor:=1.0)
		pll.Reset()

		Dim response As New List(Of Double)
		For i = 0 To 50
			response.Add(pll.FilterRun(1.0))
		Next

		StepResponseAnalyzer.AnalyzeStepResponse(response, samplingInterval:=1.0)
	End Sub

End Class

Module StepResponseAnalyzer

	''' <summary>
	''' Should not be use if the damping factor is too close to one 
	''' </summary>
	''' <param name="signal"></param>
	''' <param name="samplingInterval"></param>
	Public Sub AnalyzeStepResponse(signal As List(Of Double), samplingInterval As Double)
		If signal Is Nothing OrElse signal.Count < 5 Then
			Console.WriteLine("Signal too short for analysis.")
			Return
		End If

		Dim finalValue As Double = signal.Last()
		Dim peakValue As Double = signal.Max()
		Dim peakIndex As Integer = signal.IndexOf(peakValue)
		Dim timeToPeak As Double = peakIndex * samplingInterval

		Dim overshoot As Double = (peakValue - finalValue) / finalValue

		' Estimate damping factor (zeta)
		Dim zeta As Double = 0
		If overshoot > 0 AndAlso overshoot < 1 Then
			Dim lnMp As Double = Log(overshoot)
			zeta = -lnMp / Sqrt(Math.PI ^ 2 + lnMp ^ 2)
		End If

		' Estimate natural frequency (omega_n)
		Dim omega_n As Double = 0
		If zeta > 0 AndAlso zeta < 1 Then
			omega_n = Math.PI / (timeToPeak * Sqrt(1 - zeta ^ 2))
		End If

		' Print results
		Console.WriteLine($"Estimated Final Value: {finalValue:F4}")
		Console.WriteLine($"Peak Value: {peakValue:F4}")
		Console.WriteLine($"Time to Peak: {timeToPeak:F4} s")
		Console.WriteLine($"Overshoot: {overshoot:P2}")
		Console.WriteLine($"Estimated ζ (Damping): {zeta:F4}")
		Console.WriteLine($"Estimated ωₙ (rad/s): {omega_n:F4}")
		Console.WriteLine($"Estimated fₙ (Hz): {omega_n / (2 * PI):F4}")
	End Sub


	Public Sub TestPhaseShift()
		' Parameters
		Dim frequency As Double = 0.1  ' Hz
		Dim samplingRate As Double = 1.0  ' Hz
		Dim omega As Double = 2 * PI * frequency
		Dim dt As Double = 1.0 / samplingRate
		Dim totalTime As Double = 100
		Dim steps As Integer = CInt(totalTime / dt)

		' Filter under test
		Dim pll As New FilterPLL(FilterRate:=7, DampingFactor:=1.0)
		pll.Reset()

		' Track last sign for zero-crossing detection
		Dim lastInput As Double = 0
		Dim lastOutput As Double = 0
		Dim inputPhaseZeroTime As Double = -1
		Dim outputPhaseZeroTime As Double = -1
		Dim found As Boolean = False

		For i As Integer = 0 To steps
			Dim t As Double = i * dt
			Dim inputSignal As Double = Sin(omega * t)
			Dim outputSignal As Double = pll.FilterRun(inputSignal)

			' Detect input zero-crossing (rising)
			If lastInput < 0 AndAlso inputSignal >= 0 AndAlso inputPhaseZeroTime < 0 Then
				inputPhaseZeroTime = t
			End If

			' Detect output zero-crossing (rising)
			If lastOutput < 0 AndAlso outputSignal >= 0 AndAlso outputPhaseZeroTime < 0 Then
				outputPhaseZeroTime = t
				found = True
			End If

			lastInput = inputSignal
			lastOutput = outputSignal

			If found Then Exit For
		Next

		If found Then
			Dim deltaT As Double = outputPhaseZeroTime - inputPhaseZeroTime
			Dim phaseRadians As Double = 2 * PI * frequency * deltaT
			Dim phaseDegrees As Double = phaseRadians * 180 / PI

			Console.WriteLine($"Input zero at t = {inputPhaseZeroTime:F4} s")
			Console.WriteLine($"Output zero at t = {outputPhaseZeroTime:F4} s")
			Console.WriteLine($"Phase delay: {phaseDegrees:F2} degrees ({phaseRadians:F2} rad)")
		Else
			Console.WriteLine("Could not detect both zero crossings.")
		End If

		Console.ReadKey()
	End Sub
End Module

