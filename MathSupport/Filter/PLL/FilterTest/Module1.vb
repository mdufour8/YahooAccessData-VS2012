Imports FilterTest.FilterTest

Imports System.IO

Module FilterTest
	Sub Main()
		Dim MyPLLFilterUnit As FilterPLL
		Dim MyPLLFilterSinusoidal As FilterPLL
		Dim ThisResultUnit As Double
		Dim ThisResultSinus As Double
		Dim InputSignal As Double
		Dim FrequencyHz As Double = 1 / 20.0  ' 10-day cycle frequency
		Dim SamplingRate As Double = 1.0  ' 1 sample per day
		Dim Time As Double

		Dim I As Integer

		' Initialize PLL filter with a 7-day cutoff and damping factor of 1.0
		Dim ThisAppDir = My.Application.Info.DirectoryPath
		Dim ThisLogFile As String = My.Computer.FileSystem.CombinePath(ThisAppDir, "output.csv")
		MyPLLFilterUnit = New FilterPLL(FilterRate:=7, DampingFactor:=1.0)
		MyPLLFilterSinusoidal = New FilterPLL(FilterRate:=7, DampingFactor:=1.0)
		' Create a StreamWriter to write to a CSV file
		Try
			File.Delete(ThisLogFile)
		Catch ex As Exception
			MsgBox(ex.Message)
			Environment.Exit(0) ' Exit the application with a success code
		End Try
		Using writer As New StreamWriter("output.csv")
			MyPLLFilterUnit.Reset()
			MyPLLFilterUnit.FilterRun(0.0)  ' Initialize with zero
			MyPLLFilterSinusoidal.Reset()
			MyPLLFilterSinusoidal.FilterRun(0.0) 'set the filter variable at zero 
			Console.WriteLine($"Time,Unit Input,Unit Output,Sinus Input,Sinus Output")
			writer.WriteLine($"Time,Unit Input,Unit Output,Sinus Input,Sinus Output")
			For I = 0 To 50
				Time = I * SamplingRate
				InputSignal = Math.Sin(2 * Math.PI * FrequencyHz * Time)
				ThisResultUnit = MyPLLFilterUnit.FilterRun(1.0)
				ThisResultSinus = MyPLLFilterSinusoidal.FilterRun(InputSignal)
				Console.WriteLine($"{Time},{1.0},{ThisResultUnit},{InputSignal},{ThisResultSinus}")
				writer.WriteLine($"{Time},{1.0},{ThisResultUnit},{InputSignal},{ThisResultSinus}")
			Next
		End Using
		Console.WriteLine("Press any key to exit...")
		Console.ReadKey()
		Environment.Exit(0) ' Exit the application with a success code
	End Sub
End Module


''' <summary>
''' The FilterPLL class implements a Phase-Locked Loop (PLL) filter that processes an input signal to generate a filtered output.
''' The filter uses a set of coefficients and error terms to adjust the output based on the input signal.
''' The class provides methods to run the filter and reset it, as well as properties to access the current state of the filter.
''' </summary>
Public Class FilterPLL
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


			'Old method
			'FreqDigital = 1 / (Math.PI * MyFilterRate)
			'C = MyVCOPeriod
			'C2 = 2 * MyDampingFactor * (2 * Math.PI) * FreqDigital

			'new method based on chatGPT calculation with more explicit detailed relation 
			'between the different frequency scaling relation.
			'this new method give exactly the same result than before but is just more explicit
			'if we need to change the sampling rate currentlt set at 1 day.
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
			IsReset = True
		End Sub

		Public Function FilterRun(Value As Double) As Double
			If IsReset Then
				'initialization
				'initialize the loop with the first time sample
				'this is to minimize the PLL tracking error
				MyFilterValuek0 = Value
				MyFilterValueLast = MyFilterValuek0
				'MyRefValue is the first reference sample value and never change annymore
				MyRefValue = Value
				MyVCOk0 = 0
				IsReset = False
			End If
			'just the standard PLL here
			MySignalDelta = Value - MyFilterValuek0
			'calculate the filter loop parameters
			MyErrork1 = MyErrork0
			MyErrork0 = (C1 * MySignalDelta) + MyErrork1
			MyFilterError = (C2 * MySignalDelta) + MyErrork1
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

		Public ReadOnly Property FilterLast As Double
			Get
				Return MyFilterValueLast
			End Get
		End Property

		Public ReadOnly Property Rate() As Double
			Get
				Return MyFilterRate
			End Get
		End Property

		Public Sub Reset()
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
				Return MyFilterError
			End Get
		End Property
	End Class
