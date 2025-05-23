﻿Imports System.Math
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


	Private MyError() As Double
	Private MyVCO() As Double

	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private MySignalDelta As Double
	Private MyFilterError As Double
	Private MyFilterRate As Double
	Private MyDampingFactor As Double
	Private _IsReset As Boolean

	'note we cannot create another PLL filter without causing stack overflow
	Private MyFilterDoubleExpForError As FilterDoubleExp
	Private MyFilterDoubleExpForBandPass As FilterDoubleExp
	Private MyFilterDoubleExpForBandPassAmplitude As Double
	Private MyFilterTrendLast As Double
	Private MyFilterBandPassLast As Double
	Private MyCircularBuffer As CircularBuffer(Of Double)

	Public Sub New(ByVal FilterRate As Double, Optional DampingFactor As Double = 1.0, Optional BufferCapacity As Integer = 0)
		Dim ThisFilterRateMinimum As Double
		Dim FreqDigital As Double
		Dim SamplingPeriod As Double
		Dim FrequencyNaturel As Double
		'Here using special character to better understand the code that follow
		Dim This_Ts As Double = 1.0    'Sampling rate Is 1 day by Default in all that base code 
		Dim This_fₙ As Double
		Dim This_ωₙ As Double
		'Dim This_ωc As Double
		Dim This_Ωn As Double
		Dim This_Ωc As Double

		'the ADPLL need the natural frequency of the filter. 

		MyFilterRate = FilterRate
		MyDampingFactor = DampingFactor
		ThisFilterRateMinimum = 5 * MyDampingFactor
		'If MyFilterRate < ThisFilterRateMinimum Then
		'	MyFilterRate = ThisFilterRateMinimum
		'End If
		'original code
		'Let explain this relation:
		'FreqDigital = 1 / (PI * MyFilterRate)
		'Please refer to the article 'DigitalFrequency_vs_AngularFrequency.pdf' for more information.
		'In this article the relation bwetween the 3dB bandwidth and the filter resonance frequncy is explained.
		'It is noted that a formal relation between the 3dB bandwidth and the filter resonance frequency is not given.	
		'However it can be approxmated with simple relation of 2.0.  
		'The user provide a filter rate parameters that is effect					ivly the desired 3dB bandwidth of the filter.
		'The filter rate is the 3dB bandwidth of the filter but the PLL filter is described by it resonnance frequency:
		'Fn=Fc/2.0 = 2*PI/2.0*FilterRate)=PI/FilterRate	

		'Note also that the digital frequency Ω is the angular frequency ω (rad/s) times the sampling period T (s/sample),
		'so the units digital frequency are rad/sample.

		'Ω=ω*T	

		'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
		'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
		'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
		'A = CDbl((2 / (MyFilterRate + 1)))

		Dim Ts As Double = 1.0   '1 Sample per day by default
		'Dim fs As Double = Ts / MyFilterRate
		'Dim f
		This_Ωn = Math.Log((MyFilterRate + 1) / (MyFilterRate - 1)) 'Digital Frequency: Ω = ω × Ts



		ReDim MyError(0 To 2)
		ReDim MyVCO(0 To 2)
		For I = 0 To MyError.Length - 1
			MyError(I) = 0.0
			MyVCO(I) = 0.0
		Next
		MyVCOPeriod = 0   'by default in this application


		'FreqDigital = 1 / (Math.PI * MyFilterRate)  'Digital Frequency: Ω = ω × T
		'FreqDigital = 1 / (Math.PI * MyFilterRate)  'Digital Frequency: Ω = ω × T
		C = MyVCOPeriod
		'C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital
		'MyFilterDoubleExpForBandPassAmplitude = (MyDampingFactor * This_Ωn)
		MyFilterDoubleExpForBandPassAmplitude = 1.0 'can be removed not in used
		C2 = 2 * MyDampingFactor * This_Ωn

		'Dim FreqDigital1 = Math.Log((MyFilterRate + 1) / (MyFilterRate - 1)) / (2 * (MyDampingFactor + (1 / (4 * MyDampingFactor))))
		'C2 = 2 * MyDampingFactor * FreqDigital1

		'see DPLL_Detailed_Frequency_Response_Analysis_1.pdf for explanation
		'FreqDigital = 1.0  '(Sampling rate Is 1.0 Sample/day by Default))

		'FrequencyNaturel = 1 / MyFilterRate
		'C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital
		'C2 = 2 * MyDampingFactor * (2 * Math.PI * FrequencyNaturel) * SamplingPeriod / Math.PI

		C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
		'check stability
		If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
			'loop is unstable
			Throw New Exception("Low pass PLL filter is not stable...")
		End If
		MySignalDelta = 0
		MyFilterDoubleExpForError = New FilterDoubleExp(FilterRate:=MyFilterRate)
		MyFilterDoubleExpForBandPass = New FilterDoubleExp(FilterRate:=MyFilterRate)
		MyCircularBuffer = New CircularBuffer(Of Double)(capacity:=BufferCapacity, 0.0)
		_IsReset = True
	End Sub

#Region "IFilterRun"
	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		If _IsReset Then
			'initialization
			'initialize the loop with the first time sample
			'this is to minimize the PLL tracking error
			For I = 0 To MyError.Length - 1
				MyError(I) = 0
				MyVCO(I) = Value
			Next
			MyCircularBuffer.Clear()
			MyFilterDoubleExpForError.Reset()
			MyFilterDoubleExpForBandPass.Reset()
			_IsReset = False
		End If
		'just the standard PLL here
		MySignalDelta = Value - MyVCO(0)
		'calculate the filter loop parameters
		MyError(1) = MyError(0)
		MyError(0) = (C1 * MySignalDelta) + MyError(1)

		MyFilterError = (C2 * MySignalDelta) + MyError(1)
		MyFilterDoubleExpForError.FilterRun(MyFilterError)

		'calculate the integrator parameters
		MyVCO(2) = MyVCO(1)
		MyVCO(1) = MyVCO(0)
		MyVCO(0) = C + MyFilterError + MyVCO(1)
		MyFilterDoubleExpForBandPass.FilterRun(Value)
		'MyFilterBandPassLast = MyVCO(0) - (MyFilterDoubleExpForBandPassAmplitude * MyFilterDoubleExpForBandPass.FilterLast)
		MyFilterBandPassLast = MyVCO(0) - MyFilterDoubleExpForBandPass.FilterLast
		'MyFilterTrendLast = MyFilterBandPassLast / 2
		ValueLast = Value
		MyCircularBuffer.AddLast(MyVCO(0))
		Return MyVCO(0)
	End Function


	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Return MyVCO(0)
		End Get
	End Property

	Public ReadOnly Property FilterRate() As Double Implements IFilterRun.FilterRate
		Get
			Return MyFilterRate
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		_IsReset = True
	End Sub
	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun.Reset
		Me.Reset()
		MyCircularBuffer = New CircularBuffer(Of Double)(capacity:=BufferCapacity, 0.0)
	End Sub

	Private ReadOnly Property IFilter_IsReset As Boolean Implements IFilterRun.IsReset
		Get
			Return _IsReset
		End Get
	End Property

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

	Public ReadOnly Property Count As Integer Implements IFilterRun.Count
		Get
			Return MyCircularBuffer.Count
		End Get
	End Property
	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Return $"{Me.GetType().Name}({MyFilterRate},{MyDampingFactor})"
		End Get
	End Property
	Public ReadOnly Property InputLast As Double Implements IFilterRun.InputLast
		Get
			Return ValueLast
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of Double) Implements IFilterRun.ToList
		Get
			Return MyCircularBuffer.ToList()
		End Get
	End Property

	Public ReadOnly Property FilterValuek0 As Double
		Get
			Return MyVCO(0)
		End Get
	End Property

	Public ReadOnly Property FilterValuek1 As Double
		Get
			Return MyVCO(1)
		End Get
	End Property

	''' <summary>
	''' The filtered Bandpass output for this filter centered at the resonnance frequency.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterBandPassLast As Double
		Get
			Return MyFilterBandPassLast
		End Get
	End Property

	Public ReadOnly Property FilterTrendLast As Double Implements IFilterRun.FilterTrendLast
		Get
			Return MyFilterDoubleExpForError.FilterLast
		End Get
	End Property

	Public ReadOnly Property Errork0 As Double
		Get
			Return MyError(0)
		End Get
	End Property

	Public ReadOnly Property Errork1 As Double
		Get
			Return MyError(1)
		End Get
	End Property

	Public ReadOnly Property FilterError As Double
		Get
			Return MyFilterDoubleExpForError.FilterLast
		End Get
	End Property

	Public Overrides Function ToString() As String
		Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
	End Function
#End Region
#Region "IFilterState"
	Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
		Return Me
	End Function

	Private MyQueueForState As New Queue(Of Double)
	Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
		Try
			If MyQueueForState.Count = 0 Then Return
			MyError(0) = MyQueueForState.Dequeue
			MyError(1) = MyQueueForState.Dequeue
			MyVCO(0) = MyQueueForState.Dequeue
			MyVCO(1) = MyQueueForState.Dequeue
			MyVCO(2) = MyQueueForState.Dequeue
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
		MyQueueForState.Enqueue(MyError(0))
		MyQueueForState.Enqueue(MyError(1))

		MyQueueForState.Enqueue(MyVCO(0))
		MyQueueForState.Enqueue(MyVCO(1))
		MyQueueForState.Enqueue(MyVCO(2))

		MyQueueForState.Enqueue(ValueLast)
		MyQueueForState.Enqueue(ValueLastK1)
		MyQueueForState.Enqueue(MySignalDelta)
		MyQueueForState.Enqueue(MyFilterError)
	End Sub
#End Region
#Region "IFilter"
	Private ReadOnly Property IFilter_Count As Integer Implements IFilter.Count
		Get
			Return Me.Count
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

	Public Property Tag As String Implements IFilter.Tag

	Public ReadOnly Property Rate As Integer Implements IFilter.Rate
		Get
			Return CInt(Me.FilterRate)
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
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_Filter(Value As IPriceVol) As Double Implements IFilter.Filter
		Return Me.FilterRun(Value.Last)
	End Function

	Private Function IFilter_FilterErrorLast() As Double Implements IFilter.FilterErrorLast
		Return Me.FilterError
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
		Return MyCircularBuffer.ToArray()
	End Function

	Private Function IFilter_ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
		Throw New NotImplementedException()
	End Function

	Private Function IFilter_ToString() As String Implements IFilter.ToString
		Return Me.ToString()
	End Function
#End Region
End Class


#Region "StepResponseAnalyzer"
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

	''' <summary>
	''' Simulates a unit step input to the FilterPLL and prints analysis to console.
	''' </summary>
	Public Sub TestUnitStepResponse()
		Dim pll As New FilterPLL(FilterRate:=20, DampingFactor:=1.0)
		pll.Reset()

		Dim response As New List(Of Double)
		For i = 0 To 50
			response.Add(pll.FilterRun(1.0))
		Next

		StepResponseAnalyzer.AnalyzeStepResponse(response, samplingInterval:=1.0)
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
#End Region


