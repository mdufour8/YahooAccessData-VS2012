﻿#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region


Namespace MathPlus.Filter
  <Serializable()>
  Public Class FilterLowPassPLL
    Implements IFilter
    Implements IFilterPrediction
    Implements IFilterControl
    Implements IFilterControlRate
    Implements YahooAccessData.IRegisterKey(Of String)
    Implements IFilterCopy
    Implements IFilterState
    Implements IFilterEstimate
    Implements IFilterCreateNew

    Public Const DAMPING_FACTOR As Double = 1.0#

    Private C As Double
    Private C1 As Double
    Private C2 As Double

    Private MyErrork0 As Double
    Private MyErrork1 As Double

    Private MyVCOPeriod As Double

    Private MyVCOk0 As Double
    Private MyVCOk1 As Double
    Private MyVCOk2 As Double

    Private MyRefValue As Double
    Private MyFilterValueLast As Double
    Private MyFilterValuekInput() As Double
    Private MyFilterValuek() As Double
    Private MyFilterValuek0 As Double
    Private MyFilterValuek1 As Double
    Private MyFilterValuek2 As Double
    Private MyRate As Integer
    Private MyFilterRate As Double
    Private ValueLast As Double
    Private ValueLastK1 As Double
    Private MySignalDelta As Double
    Private MyFilterError As Double
    Private MyListOfValue As ListScaled
    Private MyFilterPredictionOutput As Integer
    Private MyListOfFilterErrorValue As ListScaled
    Private MyFilterPrediction As Filter.FilterLowPassExpPredict
    Private MyFilterPredictionDerivative As Filter.FilterLowPassPLLPredict
    'Private MyFilterPLLForError As Filter.IFilter
    Private MyDampingFactor As Double
    Private IsPredictionEnabledLocal As Boolean
    Private MyInputValue() As Double
    Private MyQueueForState As Queue(Of Double)
    Private IsRunFilterLocal As Boolean

    Public Sub New(
                  ByVal FilterRate As Double,
                  ByVal NumberOfPredictionOutput As Integer,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional ByVal IsPredictionEnabled As Boolean = False)

      Dim FreqDigital As Double
      Dim ThisFilterRateMinimum As Double

      MyListOfValue = New ListScaled
      MyListOfFilterErrorValue = New ListScaled

      IsPredictionEnabledLocal = IsPredictionEnabled

      'to ensure stability restrict the range of value (see paper)
      MyFilterRate = FilterRate
      MyDampingFactor = DampingFactor
      ThisFilterRateMinimum = 5 * MyDampingFactor
      If MyFilterRate < ThisFilterRateMinimum Then
        MyFilterRate = ThisFilterRateMinimum
      End If
      MyRate = CInt(MyFilterRate)
      'FreqDigital = 1 / (2 * Math.PI * FilterRate)
      FreqDigital = 1 / (Math.PI * MyFilterRate)
      'FreqDigital = 1 / (FilterRate)

      MyVCOPeriod = 0   'by default in this application
      MyErrork1 = 0
      MyFilterValuek2 = 0
      ValueLast = 0
      ValueLastK1 = 0
      ReDim MyInputValue(-1)
      ReDim MyFilterValuek(0 To 2)
      ReDim MyFilterValuekInput(0 To 2)
      MyFilterValuekInput(0) = 0
      MyFilterValuekInput(1) = 1
      MyFilterValuekInput(2) = 2
      C = MyVCOPeriod
      C2 = 2 * MyDampingFactor * TwoPI * FreqDigital
      C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
      'check stability
      If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
        'loop is unstable
        Throw New Exception("Low pass PLL filter is not stable...")
      End If
      MySignalDelta = 0
      MyErrork0 = 0
      IsRunFilterLocal = False
      MyFilterPredictionOutput = NumberOfPredictionOutput
      If IsPredictionEnabledLocal Then
        'this is base on a exponetial filter
        MyFilterPrediction = New Filter.FilterLowPassExpPredict(
          NumberToPredict:=MyFilterPredictionOutput,
          FilterHead:=New FilterLowPassPLL(FilterRate:=MyFilterRate, NumberOfPredictionOutput:=MyFilterPredictionOutput, DampingFactor:=MyDampingFactor))

        MyFilterPredictionDerivative = New Filter.FilterLowPassPLLPredict(
          NumberToPredict:=MyFilterPredictionOutput,
          FilterHead:=New FilterLowPassPLL(FilterRate:=MyFilterRate, NumberOfPredictionOutput:=MyFilterPredictionOutput, DampingFactor:=MyDampingFactor))
      Else
        MyFilterPrediction = Nothing
        MyFilterPredictionDerivative = Nothing
      End If
      MyQueueForState = New Queue(Of Double)
    End Sub


    Public Sub New(
                  ByVal FilterRate As Double,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional ByVal IsPredictionEnabled As Boolean = False)

      Me.New(FilterRate:=FilterRate, NumberOfPredictionOutput:=0, DampingFactor:=DampingFactor, IsPredictionEnabled:=IsPredictionEnabled)
    End Sub

    Public Sub New(
                  ByVal FilterRate As Integer,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional IsPredictionEnabled As Boolean = False)

      Me.New(FilterRate:=FilterRate, NumberOfPredictionOutput:=0, DampingFactor:=DampingFactor, IsPredictionEnabled:=IsPredictionEnabled)
    End Sub

    Friend Sub New(
                  ByVal FilterRate As Double,
                  ByRef InputValue() As Double,
                  ByVal IsRunFilter As Boolean,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional IsPredictionEnabled As Boolean = False)

      Me.New(FilterRate:=FilterRate, NumberOfPredictionOutput:=0, DampingFactor:=DampingFactor, IsPredictionEnabled:=IsPredictionEnabled)
      ReDim MyInputValue(0 To InputValue.Length - 1)
      InputValue.CopyTo(MyInputValue, 0)
      If IsRunFilter Then
        IsRunFilterLocal = True
        Me.Filter(MyInputValue, MyInputValue.Length - 1)
      End If
    End Sub

    Public Sub New(
                  ByVal FilterRate As Integer,
                  ByRef InputValue() As Double,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional IsPredictionEnabled As Boolean = False)

      Me.New(CDbl(FilterRate), InputValue, DampingFactor, IsPredictionEnabled)
    End Sub

    Public Sub New(
                  ByVal FilterRate As Double,
                  ByRef InputValue() As Double,
                  Optional ByVal DampingFactor As Double = DAMPING_FACTOR,
                  Optional IsPredictionEnabled As Boolean = False)

      Me.New(FilterRate:=FilterRate, NumberOfPredictionOutput:=0, DampingFactor:=DampingFactor, IsPredictionEnabled:=IsPredictionEnabled)
      ReDim MyInputValue(0 To InputValue.Length - 1)
      InputValue.CopyTo(MyInputValue, 0)
      Me.Filter(MyInputValue)
    End Sub

    Public Function AsIFilterCreateNew() As IFilterCreateNew Implements IFilterCreateNew.AsIFilterCreateNew
      Return Me
    End Function

    Private Function IFilterCreateNew_CreateNew() As IFilter Implements IFilterCreateNew.CreateNew
      Dim ThisFilter As IFilter

      If MyInputValue.Length > -1 Then
        ThisFilter = New FilterLowPassPLL(FilterRate:=MyFilterRate, InputValue:=MyInputValue, DampingFactor:=MyDampingFactor, IsRunFilter:=IsRunFilterLocal, IsPredictionEnabled:=IsPredictionEnabledLocal)
      Else
        ThisFilter = New FilterLowPassPLL(FilterRate:=MyFilterRate, DampingFactor:=MyDampingFactor, IsPredictionEnabled:=IsPredictionEnabledLocal)
      End If
      Return ThisFilter
    End Function


#Region "Filter function"
    ''' <summary>
    ''' Function accessible by other inherited class allowing backgroud data loading for special filtering 
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <param name="ValueFiltered"></param>
    ''' <remarks></remarks>
    Friend Sub Filter(ByVal Value As Double, ValueFiltered As Double)
      Dim ThisFilterPredictionGainYearly As Double

      If MyListOfValue.Count = 0 Then
        'initialization
        'initialize the loop with the first time sample
        'this is to minimize the PLL tracking error at the beginning
        MyFilterValuek0 = Value
        MyFilterValueLast = MyFilterValuek0
        'MyRefValue is the first reference sample value and never change annymore
        MyRefValue = Value
        MyVCOk0 = 0
      End If
      ''comparaison signal input difference
      MySignalDelta = Value - MyFilterValuek0
      ''calculate the filter loop parameters
      'MyErrork1 = MyErrork0
      'MyErrork0 = (C1 * MySignalDelta) + MyErrork1
      'MyFilterError = (C2 * MySignalDelta) + MyErrork1
      ''calculate the integrator parameters
      'MyVCOk2 = MyVCOk1
      'MyVCOk1 = MyVCOk0
      'MyVCOk0 = C + MyFilterError + MyVCOk1
      'in this implementation MyFilterError is the instantaneous slope of the signal output
      MyFilterValuek2 = MyFilterValuek1
      MyFilterValuek1 = MyFilterValuek0

      'MyFilterValuek0 is the next signal predicted value for the next sample
      MyFilterValuek0 = ValueFiltered
      If MyFilterPredictionOutput = 0 Then
        MyFilterValueLast = MyFilterValuek0
      Else
        MyFilterValueLast = Me.FilterPredictionNext(MyFilterPredictionOutput)
      End If
      'this equation is an approximation of the gain valid for value >>1
      'the computation fail for small positive and negative value but the degradation is predictable and the derivative exist.
      'the intend is not to have an exact gain measurement but a closed form that behave on a predictable value

      ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(MyFilterValuek0, MyFilterValuek1))
      ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
      MyListOfFilterErrorValue.Add(ThisFilterPredictionGainYearly)

      'MyListOfFilterErrorValue.Add(MyFilterError)
      'MyListOfFilterErrorValue.Add(MyFilterValuek0)

      'returning the previous value ensure that the PLL loop delay is zero
      'MyFilterValuek0 is the predicted value for the next sample
      ValueLastK1 = ValueLast
      ValueLast = Value
      MyListOfValue.Add(MyFilterValueLast)
      If MyFilterPrediction IsNot Nothing Then
        MyFilterPrediction.Filter(Value)
        MyFilterPredictionDerivative.Filter(Value)
      End If
    End Sub

    Public Overridable Function Filter(ByVal Value As Double, ByVal FilterPLLDetector As IFilterPLLDetector) As Double
      Dim ThisFilterPredictionGainYearly As Double

      If MyListOfValue.Count = 0 Then
        'initialization
        'initialize the loop with the first time sample
        'this is to minimize the PLL tracking error
        If FilterPLLDetector IsNot Nothing Then
          MyFilterValuek0 = FilterPLLDetector.ValueInit
          'MyRefValue is the first reference sample value and never change annymore
          MyRefValue = MyFilterValuek0
          MyFilterValueLast = FilterPLLDetector.ValueOutput(Value, MyFilterValuek0)
        Else
          MyFilterValuek0 = Value
          MyFilterValueLast = MyFilterValuek0
          'MyRefValue is the first reference sample value and never change annymore
          MyRefValue = Value
        End If
        MyVCOk0 = 0
      End If
      'comparaison signal input difference
      If FilterPLLDetector IsNot Nothing Then
        Dim ThisNumberLoop As Integer = 0
        Dim ThisValueStart As Double = FilterPLLDetector.ValueOutput(Value, MyFilterValuek0)
        Dim ThisValueStop As Double
        Do
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
        If MyFilterPredictionOutput = 0 Then
          MyFilterValueLast = ThisValueStop
        Else
          MyFilterValueLast = FilterPLLDetector.ValueOutput(Value, Me.FilterPredictionNext(MyFilterPredictionOutput))
        End If
        MyListOfValue.Add(MyFilterValueLast)
      Else
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

        'MyFilterValuek0 is the next signal predicted value for the next sample
        MyFilterValuek0 = MyRefValue + MyVCOk0
        If MyFilterPredictionOutput = 0 Then
          MyFilterValueLast = MyFilterValuek0
        Else
          MyFilterValueLast = Me.FilterPredictionNext(MyFilterPredictionOutput)
        End If
        MyListOfValue.Add(MyFilterValueLast)
      End If
      'this equation is an approximation of the gain valid for value >>1
      'the computation fail for small positive and negative value but the degradation is predictable and the derivative exist.
      'the intend is not to have an exact gain measurement but a closed form that behave on a predictable value
      'ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * Math.Log(((MyFilterValuek0 ^ 2 + 1) / (MyFilterValuek1 ^ 2 + 1)))) / 2
      'also limit exponentially the gain value between -1 and +1
      'ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
      'filter this error to reduce the instantaneous variation
      'MyFilterPLLForError.Filter(ThisFilterPredictionGainYearly)
      'MyFilterError is the unit per day slope
      'ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(MyFilterValuek0, MyFilterValuek1))
      'ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(MyVCOk0, MyVCOk1) / MyRate)
      ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * MathPlus.Measure.Measure.GainLog(MyVCOk0, MyVCOk1))

      ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
      MyListOfFilterErrorValue.Add(ThisFilterPredictionGainYearly)

      'MyListOfFilterErrorValue.Add(MyFilterError)
      'MyListOfFilterErrorValue.Add(MyFilterValuek0)

      'returning the previous value ensure that the PLL loop delay is zero
      'MyFilterValuek0 is the predicted value for the next sample
      ValueLastK1 = ValueLast
      ValueLast = Value
      If MyFilterPrediction IsNot Nothing Then
        MyFilterPrediction.Filter(Value)
        MyFilterPredictionDerivative.Filter(Value)
      End If
      Return MyFilterValuek0
    End Function


    Public Overridable Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
      Dim ThisFilterDetector As IFilterPLLDetector = Nothing
      Return Me.Filter(Value, ThisFilterDetector)
    End Function

    Public Overridable Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
      Dim ThisValue As Double
      For Each ThisValue In Value
        Me.Filter(ThisValue)
      Next
      Return Me.ToArray
    End Function

    Public Overridable Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
      Return Me.Filter(CDbl(Value))
    End Function

    Public Overridable Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
      Return Me.Filter(CDbl(Value.Last))
    End Function

    ''' <summary>
    ''' Special filtering that can be used to remove the delay starting at a specific point
    ''' </summary>
    ''' <param name="Value">The value to be filtered</param>
    ''' <param name="DelayRemovedToItem">The point where the delay start to be removed</param>
    ''' <returns>The result</returns>
    ''' <remarks></remarks>
    Public Overridable Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
      Dim ThisValues(0 To Value.Length - 1) As Double
      Dim I As Integer
      Dim J As Integer

      Dim ThisFilterLeft As FilterLowPassPLL = New FilterLowPassPLL(MyFilterRate)
      Dim ThisFilterRight As FilterLowPassPLL = New FilterLowPassPLL(MyFilterRate)
      Dim ThisFilterLeftItem As Double
      Dim ThisFilterRightItem As Double

      'filter from the left
      ThisFilterLeft.Filter(Value)
      'filter from the right the section with the reverse filtering
      For I = DelayRemovedToItem To 0 Step -1
        ThisFilterRight.Filter(Value(I))
      Next
      'the data in ThisFilterRightList is reversed
      'need to look at it in reverse order using J
      J = DelayRemovedToItem
      If MyFilterPrediction Is Nothing Then
        For I = 0 To Value.Length - 1
          MyFilterValuek2 = MyFilterValuek1
          MyFilterValuek1 = MyFilterValuek0
          ThisFilterLeftItem = ThisFilterLeft.ToList(I)
          If I > DelayRemovedToItem Then
            MyFilterValuek0 = ThisFilterLeftItem
          Else
            ThisFilterRightItem = ThisFilterRight.ToList(J)
            MyFilterValuek0 = (ThisFilterLeftItem + ThisFilterRightItem) / 2
          End If
          If MyFilterPredictionOutput = 0 Then
            MyFilterValueLast = MyFilterValuek0
          Else
            MyFilterValueLast = Me.FilterPredictionNext(MyFilterPredictionOutput)
          End If
          MyListOfValue.Add(MyFilterValueLast)
          ThisValues(I) = MyFilterValueLast
          J = J - 1
        Next
      Else
        For I = 0 To Value.Length - 1
          MyFilterValuek2 = MyFilterValuek1
          MyFilterValuek1 = MyFilterValuek0
          ThisFilterLeftItem = ThisFilterLeft.ToList(I)
          If I > DelayRemovedToItem Then
            MyFilterValuek0 = ThisFilterLeftItem
          Else
            ThisFilterRightItem = ThisFilterRight.ToList(J)
            MyFilterValuek0 = (ThisFilterLeftItem + ThisFilterRightItem) / 2
          End If
          If MyFilterPredictionOutput = 0 Then
            MyFilterValueLast = MyFilterValuek0
          Else
            MyFilterValueLast = Me.FilterPredictionNext(MyFilterPredictionOutput)
          End If
          MyListOfValue.Add(MyFilterValueLast)
          ThisValues(I) = MyFilterValueLast
          MyFilterPrediction.Filter(Value(I))
          MyFilterPredictionDerivative.Filter(Value(I))
          J = J - 1
        Next
      End If
      Return ThisValues
    End Function
#End Region

    Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
      'not implemented yet
      Throw New NotImplementedException
      Return 0.0
    End Function

    ''' <summary>
    ''' Calculate the output prediction for the next sample using the filter last output value as the input
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FilterPredictionNext() As Double
      Return Me.FilterPredictionNext(1)
    End Function


    ''' <summary>
    ''' Calculate the output prediction for the specified number of points using the specified input filter last output value as the input.
    ''' The filter state is not changed by this call
    ''' </summary>
    ''' <returns>the filtered value of the last points</returns>
    ''' <remarks></remarks>
    Public Function FilterPredictionNext(ByVal NumberPoints As Integer) As Double
      Dim I As Integer
      Dim ThisValue As Double

      Dim ThisSignalDelta As Double
      Dim ThisFilterError As Double = MyFilterError
      Dim ThisFilterValuek0 As Double = MyFilterValuek0
      Dim ThisFilterValuek1 As Double = MyFilterValuek1
      Dim ThisFilterValuek2 As Double = MyFilterValuek2
      Dim ThisVCOk0 As Double = MyVCOk0
      Dim ThisVCOk1 As Double = MyVCOk1
      Dim ThisErrork0 As Double = MyErrork0
      Dim ThisErrork1 As Double = MyErrork1
      Dim ThisValueLast As Double
      'Dim ThisCubicSpline As Interpolation.CubicSpline
      'calculate the slope at the edge point and then interpolate
      '
      'do not take the last value has the input 
      For I = 1 To NumberPoints
        'comparaison signal input difference
        'using the spline to obtain the slope is not necessary since the PLL already have a faily good estimete of the slope
        'with the ThisFilterError variable
        'note the k value are reversed here
        'MyFilterValuek(0) = ThisFilterValuek2
        'MyFilterValuek(1) = MyFilterValuek1
        'MyFilterValuek(2) = MyFilterValuek0
        'ThisCubicSpline = Interpolation.CubicSpline.InterpolateNaturalSorted(MyFilterValuekInput, MyFilterValuek)
        'ThisFilterError = ThisCubicSpline.Differentiate(2.0)
        'ThisFilterError is the slope per sample
        'it follows that the next sample is given by
        'note that ThisFilterError is the instantaneous slope of the signal output that track the imput
        'here it is taken as the best estimate of the slope of the signal input
        'take the filter output has the best estimate of the current input 
        'this significantly reduce the noise on the futur estimate of output value
        'and is a superior for signal trading
        ThisValueLast = ThisFilterValuek0
        ThisValue = ThisValueLast + ThisFilterError
        ThisSignalDelta = ThisValue - ThisFilterValuek0
        'calculate the filter loop parameters
        ThisErrork1 = ThisErrork0
        ThisErrork0 = (C1 * ThisSignalDelta) + ThisErrork1
        ThisFilterError = (C2 * ThisSignalDelta) + ThisErrork1
        'calculate the integrator parameters
        ThisVCOk1 = ThisVCOk0
        ThisVCOk0 = C + ThisFilterError + ThisVCOk1
        'in this implementation ThisFilterError is the instantaneous slope of the signal output
        ThisFilterValuek2 = ThisFilterValuek1
        ThisFilterValuek1 = ThisFilterValuek0
        'ThisFilterValuek0 is the next signal predicted value for the next sample
        ThisFilterValuek0 = MyRefValue + ThisVCOk0
      Next
      Return ThisFilterValuek0
    End Function

    ''' <summary>
    ''' This function calculate the filtered value for the provided input using  the current filter state.
    ''' The filter state is NOT changed by this function
    ''' </summary>
    ''' <param name="Value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
      Dim ThisValue As Double

      Dim ThisSignalDelta As Double
      Dim ThisFilterError As Double = MyFilterError
      Dim ThisFilterValuek0 As Double = MyFilterValuek0
      Dim ThisFilterValuek1 As Double = MyFilterValuek1
      Dim ThisFilterValuek2 As Double = MyFilterValuek2
      Dim ThisVCOk0 As Double = MyVCOk0
      Dim ThisVCOk1 As Double = MyVCOk1
      Dim ThisErrork0 As Double = MyErrork0
      Dim ThisErrork1 As Double = MyErrork1

      ThisValue = Value
      ThisSignalDelta = ThisValue - ThisFilterValuek0
      'calculate the filter loop parameters
      ThisErrork1 = ThisErrork0
      ThisErrork0 = (C1 * ThisSignalDelta) + ThisErrork1
      ThisFilterError = (C2 * ThisSignalDelta) + ThisErrork1
      'calculate the integrator parameters
      ThisVCOk1 = ThisVCOk0
      ThisVCOk0 = C + ThisFilterError + ThisVCOk1
      'in this implementation ThisFilterError is the instantaneous slope of the signal output
      ThisFilterValuek2 = ThisFilterValuek1
      ThisFilterValuek1 = ThisFilterValuek0
      'ThisFilterValuek0 is the next signal predicted value for the next sample
      ThisFilterValuek0 = MyRefValue + ThisVCOk0
      Return ThisFilterValuek0
    End Function

    Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
      Return Me.FilterPredictionNext(CDbl(Value))
    End Function

    Public Function FilterLast() As Double Implements IFilter.FilterLast
      Return MyFilterValueLast
    End Function

    Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
      'Return MyFilterError
      Return MyListOfFilterErrorValue.Last
    End Function

    Public Function Last() As Double Implements IFilter.Last
      Return ValueLast
    End Function

    Public Overridable ReadOnly Property Rate As Integer Implements IFilter.Rate
      Get
        Return MyRate
      End Get
    End Property

    Public ReadOnly Property Count As Integer Implements IFilter.Count
      Get
        Return MyListOfValue.Count
      End Get
    End Property

    Public ReadOnly Property Max As Double Implements IFilter.Max
      Get
        Return MyListOfValue.Max
      End Get
    End Property

    Public ReadOnly Property Min As Double Implements IFilter.Min
      Get
        Return MyListOfValue.Min
      End Get
    End Property

    Public Property Tag As String Implements IFilter.Tag

    Public Overrides Function ToString() As String Implements IFilter.ToString
      Return Me.FilterLast.ToString
    End Function

    Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
      Get
        Return MyListOfValue
      End Get
    End Property

    Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
      Get
        Return MyListOfValue
      End Get
    End Property

    Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
      Get
        Return MyListOfFilterErrorValue
      End Get
    End Property

    Public Function ToArray() As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray
    End Function

    Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
      Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
    End Function

    Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
      Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
      With ThisPriceVol
        .LastPrevious = CSng(MyFilterValuek1)
        If Me.Last > .Last Then
          .High = CSng(Me.Last)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        ElseIf Me.Last < .Last Then
          .Low = CSng(Me.Last)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        End If
      End With
      Return ThisPriceVol
    End Function

    Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
      Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
      With ThisPriceVol
        .LastPrevious = CSng(ValueLastK1)
        If Me.FilterLast > .Last Then
          .High = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        ElseIf Me.FilterLast < .Last Then
          .Low = CSng(Me.FilterLast)
          .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
        End If
      End With
      Return ThisPriceVol
    End Function

#Region "IFilterEstimate"
    Public Function AsIFilterEstimate() As IFilterEstimate Implements IFilterEstimate.AsIFilterEstimate
      Return Me
    End Function

    Public Function IFilterEstimate_Filter(Value As Double) As IFilterEstimateResult Implements IFilterEstimate.Filter
      Dim ThisValue As Double

      Dim ThisSignalDelta As Double
      Dim ThisFilterError As Double = MyFilterError
      Dim ThisFilterValuek0 As Double = MyFilterValuek0
      Dim ThisFilterValuek1 As Double = MyFilterValuek1
      Dim ThisFilterValuek2 As Double = MyFilterValuek2
      Dim ThisVCOk0 As Double = MyVCOk0
      Dim ThisVCOk1 As Double = MyVCOk1
      Dim ThisErrork0 As Double = MyErrork0
      Dim ThisErrork1 As Double = MyErrork1

      ThisValue = Value
      ThisSignalDelta = ThisValue - ThisFilterValuek0
      'calculate the filter loop parameters
      ThisErrork1 = ThisErrork0
      ThisErrork0 = (C1 * ThisSignalDelta) + ThisErrork1
      ThisFilterError = (C2 * ThisSignalDelta) + ThisErrork1
      'calculate the integrator parameters
      ThisVCOk1 = ThisVCOk0
      ThisVCOk0 = C + ThisFilterError + ThisVCOk1
      'in this implementation ThisFilterError is the instantaneous slope of the signal output
      ThisFilterValuek2 = ThisFilterValuek1
      ThisFilterValuek1 = ThisFilterValuek0
      'ThisFilterValuek0 is the next signal predicted value for the next sample
      ThisFilterValuek0 = MyRefValue + ThisVCOk0
      Return New FilterEstimateResult(ThisFilterValuek0)
    End Function

    Public Function IFilterEstimate_Filter(Value() As Double) As System.Collections.Generic.IList(Of IFilterEstimateResult) Implements IFilterEstimate.Filter
      Dim ThisList As IList(Of IFilterEstimateResult)
      Dim ThisValue As Double

      ThisList = New List(Of IFilterEstimateResult)

      Dim ThisSignalDelta As Double
      Dim ThisFilterError As Double = MyFilterError
      Dim ThisFilterValuek0 As Double = MyFilterValuek0
      Dim ThisFilterValuek1 As Double = MyFilterValuek1
      Dim ThisFilterValuek2 As Double = MyFilterValuek2
      Dim ThisVCOk0 As Double = MyVCOk0
      Dim ThisVCOk1 As Double = MyVCOk1
      Dim ThisErrork0 As Double = MyErrork0
      Dim ThisErrork1 As Double = MyErrork1

      For Each ThisValue In Value
        ThisSignalDelta = ThisValue - ThisFilterValuek0
        'calculate the filter loop parameters
        ThisErrork1 = ThisErrork0
        ThisErrork0 = (C1 * ThisSignalDelta) + ThisErrork1
        ThisFilterError = (C2 * ThisSignalDelta) + ThisErrork1
        'calculate the integrator parameters
        ThisVCOk1 = ThisVCOk0
        ThisVCOk0 = C + ThisFilterError + ThisVCOk1
        'in this implementation ThisFilterError is the instantaneous slope of the signal output
        ThisFilterValuek2 = ThisFilterValuek1
        ThisFilterValuek1 = ThisFilterValuek0
        'ThisFilterValuek0 is the next signal predicted value for the next sample
        ThisFilterValuek0 = MyRefValue + ThisVCOk0
        ThisList.Add(New FilterEstimateResult(ThisFilterValuek0))
      Next
      Return ThisList
    End Function
#End Region

#Region "IFilterPrediction"
    Public Overridable Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
      Return Me
    End Function

    Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
      If MyFilterPrediction Is Nothing Then
        Return Me.FilterLast
      Else
        Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(NumberOfPrediction)
      End If
    End Function

    Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
      If MyFilterPrediction Is Nothing Then
        Return Me.FilterLast
      Else
        Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(NumberOfPrediction, GainPerYear)
      End If
    End Function

    Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
      If MyFilterPrediction Is Nothing Then
        Return Me.FilterLast
      Else
        Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction)
      End If
    End Function

    Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
      If MyFilterPrediction Is Nothing Then
        Return Me.FilterLast
      Else
        Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction, GainPerYear)
      End If
    End Function

    Private ReadOnly Property IFilterPrediction_IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
      Get
        If MyFilterPrediction Is Nothing Then
          Return False
        Else
          Return True
        End If
      End Get
    End Property

    Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
      Get
        If MyFilterPrediction Is Nothing Then
          Return Nothing
        Else
          Return MyFilterPrediction.AsIFilterPrediction.ToListOfGainPerYear
        End If
      End Get
    End Property

    Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
      Get
        If MyFilterPredictionDerivative Is Nothing Then
          Return Nothing
        Else
          Return MyFilterPredictionDerivative.AsIFilterPrediction.ToListOfGainPerYear
        End If
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

#Region "IFilterControl"
    Public Function ASIFilterControl() As IFilterControl Implements IFilterControl.ASIFilterControl
      Return Me
    End Function

    Private Sub IFilterControl_Clear() Implements IFilterControl.Clear
      Static IsHere As Boolean

      'no re-entrency allowed
      If IsHere Then Exit Sub
      IsHere = True

      MyListOfValue.Clear()
      MyListOfFilterErrorValue.Clear()

      MySignalDelta = 0
      MyErrork0 = 0
      MyErrork1 = 0
      MyFilterValuek2 = 0
      ValueLast = 0
      ValueLastK1 = 0
      If IsPredictionEnabledLocal Then
        MyFilterPrediction.ASIFilterControl.Clear()
        MyFilterPredictionDerivative.ASIFilterControl.Clear()
      End If
      IsHere = False
    End Sub

    Private Sub IFilterControl_Refresh(FilterRate As Double) Implements IFilterControl.Refresh
      Dim FreqDigital As Double
      Dim ThisFilterRateMinimum As Double

      Static IsHere As Boolean

      'no re-entrency allowed
      If IsHere Then Exit Sub
      IsHere = True

      If Me.Count > 0 Then
        'Clear the filter before changing the rate
        IFilterControl_Clear()
      End If

      'to ensure stability restrict the range of value (see paper)
      ThisFilterRateMinimum = 5 * MyDampingFactor
      MyFilterRate = FilterRate
      If MyFilterRate < ThisFilterRateMinimum Then
        MyFilterRate = ThisFilterRateMinimum
      End If
      MyRate = CInt(MyFilterRate)
      'FreqDigital = 1 / (2 * Math.PI * FilterRate)
      FreqDigital = 1 / (Math.PI * MyFilterRate)
      'FreqDigital = 1 / (FilterRate)

      MyVCOPeriod = 0   'by default in this application
      C = MyVCOPeriod
      C2 = 2 * MyDampingFactor * TwoPI * FreqDigital
      C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
      'check stability
      If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
        'loop is unstable
        Throw New Exception("Low pass PLL filter is not stable...")
      End If
      MySignalDelta = 0
      MyErrork0 = 0
      MyErrork1 = 0
      MyFilterValuek2 = 0
      ValueLast = 0
      ValueLastK1 = 0
      If IsPredictionEnabledLocal Then
        MyFilterPrediction.ASIFilterControl.Refresh(MyFilterRate)
      End If
      'reload the filter if we have the input value
      If MyInputValue.Length > 0 Then
        Me.Filter(MyInputValue)
      End If
      IsHere = False
    End Sub

    Private Sub IFilterControl_Refresh(Rate As Integer) Implements IFilterControl.Refresh
      IFilterControl_Refresh(CDbl(Rate))
    End Sub

    Private ReadOnly Property IFilterControl_FilterRate As Double Implements IFilterControl.FilterRate
      Get
        Return MyFilterRate
      End Get
    End Property

    Private Function IFilterControl_InputValue() As Double() Implements IFilterControl.InputValue
      Return MyInputValue
    End Function

    Private ReadOnly Property IFilterControl_IsInputEnabled As Boolean Implements IFilterControl.IsInputEnabled
      Get
        Return MyInputValue.Length > 0
      End Get
    End Property
#End Region
#Region "IFilterControlRate"
    Public Function AsIFilterControlRate() As IFilterControlRate Implements IFilterControlRate.AsIFilterControlRate
      Return Me
    End Function

    Private Sub IFilterControlRate_UpdateRate(FilterRate As Double) Implements IFilterControlRate.UpdateRate
      Dim FreqDigital As Double
      Dim ThisFilterRateMinimum As Double

      'to ensure stability restrict the range of value (see paper)
      ThisFilterRateMinimum = 5 * MyDampingFactor
      MyFilterRate = FilterRate
      If MyFilterRate < ThisFilterRateMinimum Then
        MyFilterRate = ThisFilterRateMinimum
      End If
      MyRate = CInt(MyFilterRate)
      'FreqDigital = 1 / (2 * Math.PI * FilterRate)
      FreqDigital = 1 / (Math.PI * MyFilterRate)
      'FreqDigital = 1 / (FilterRate)

      MyVCOPeriod = 0   'by default in this application
      C = MyVCOPeriod
      C2 = 2 * MyDampingFactor * TwoPI * FreqDigital
      C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
      'check stability
      If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
        'loop is unstable
        Throw New Exception("Low pass PLL filter is not stable...")
      End If
    End Sub

    Private Sub IFilterControlRate_UpdateRate(Rate As Integer) Implements IFilterControlRate.UpdateRate
      IFilterControlRate_UpdateRate(CDbl(Rate))
    End Sub

    Private Property IFilterControlRate_Enabled As Boolean Implements IFilterControlRate.Enabled
      'always true here
      Get
        Return True
      End Get
      Set(value As Boolean)

      End Set
    End Property
#End Region
#Region "Private Function"
    Private Sub CalculatePLLParameters(ByVal FilterRate As Double, ByVal DampingFactor As Double)
      'note FreqDigital is the ratio of fn/fs where fn is the
      'natural frequency of the filter or the 3 dB passband and Fs the sampling rate
      'in this application fs is generally 1 day and the VCO Period is zero
      'it follow that
      'FreqDigital = 1 / (2 * Math.PI * FilterRate)
      'FreqDigital = 1 / (FilterRate)
      Dim FreqDigital = 1 / (Math.PI * FilterRate)

      If DampingFactor > 3 Then DampingFactor = 3
      MyDampingFactor = DampingFactor
      MyVCOPeriod = 0   'by default in this application
      C = MyVCOPeriod
      C2 = 2 * DampingFactor * TwoPI * FreqDigital
      C1 = (C2 ^ 2) / (4 * (MyDampingFactor ^ 2))
      'check stability
      If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
        'loop is unstable
        Throw New Exception("Low pass PLL filter is not stable...")
      End If
    End Sub
#End Region
#Region "IFilterCopy"
    Public Function AsIFilterCopy() As IFilterCopy Implements IFilterCopy.AsIFilterCopy
      Return Me
    End Function

    Private Function IFilterCopy_CopyFrom() As IFilter Implements IFilterCopy.CopyFrom
      Dim ThisFilter As FilterLowPassExp

      If MyInputValue.Length > 0 Then
        If Me.Count = 0 Then
          ThisFilter = New FilterLowPassExp(MyFilterRate, MyInputValue, IsRunFilter:=False, IsPredictionEnabled:=IsPredictionEnabledLocal)
        Else
          ThisFilter = New FilterLowPassExp(MyFilterRate, MyInputValue, IsRunFilter:=True, IsPredictionEnabled:=IsPredictionEnabledLocal)
        End If
      Else
        ThisFilter = New FilterLowPassExp(MyFilterRate, IsPredictionEnabled:=IsPredictionEnabledLocal)
      End If
      Return ThisFilter
    End Function
#End Region
#Region "IFilterState"
    Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
      Return Me
    End Function

    Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
      Dim ThisCount As Integer

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
      ThisCount = CInt(MyQueueForState.Dequeue)
      If MyListOfValue.Count > ThisCount Then
        Do
          MyListOfFilterErrorValue.RemoveAt(MyListOfFilterErrorValue.Count - 1)
          MyListOfValue.RemoveAt(MyListOfValue.Count - 1)
        Loop Until MyListOfValue.Count = ThisCount
      End If
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
      MyQueueForState.Enqueue(Me.Count)
    End Sub
#End Region
  End Class
End Namespace