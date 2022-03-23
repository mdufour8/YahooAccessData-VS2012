VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimePLL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const FREQ_DIGITAL_DEFAULT As Double = 0.01
Private Const DAMPING_FACTOR As Double = 2#
'Private MyOcc As clsTimeOscillator

Private C As Double
Private C1 As Double
Private C2 As Double
Private MyDateRefValue As Date
Private MyDateValuek0 As Date
Private MyDateValuek1 As Date
Private MyDateValuek2 As Date

Private MyTimeErrork0 As Double
Private MyTimeErrork1 As Double

Private MyVCOPeriod As Double
Private MyDateDelay As Double
Private IsFilterReset As Boolean

Private MyVCOk0 As Double
Private MyVCOk1 As Double
Private MyVCOk2 As Double
'

Public Function Run(ByVal DateValue As Date) As Date
Dim PDTimeError As Double
Dim FilterTimeError As Double

If IsFilterReset Then
  IsFilterReset = False
  'initialize the loop with the first time sample
  'this is to minimize the TPLL tracking error
  MyDateValuek0 = DateValue
  'MyDateRefValue is the first reference sample value and never change annymore
  MyDateRefValue = DateValue
  MyVCOk0 = 0
End If
'time comparaison is in the date format
MyDateDelay = (DateValue - MyDateValuek0)
'the loop calculation assume 1 sec = 2PI
'scale all the filter calculation in second
PDTimeError = NUMBER_SEC_DAY * MyDateDelay
'Debug.Print PDTimeError, NUMBER_SEC_DAY * (DateValue - ThisDateValueLast)
'ThisDateValueLast = DateValue
'calculate the filter loop parameters
MyTimeErrork1 = MyTimeErrork0
MyTimeErrork0 = (C1 * PDTimeError) + MyTimeErrork1
FilterTimeError = (C2 * PDTimeError) + MyTimeErrork1
'VCO time output are in second
'calculate the integrator parameters
'note all the MyVCOk* value are in sec from the MyDateRefValue
'This scale will not overflow and is easier to follow for debugging purpose
MyVCOk2 = MyVCOk1
MyVCOk1 = MyVCOk0
MyVCOk0 = C + FilterTimeError + MyVCOk1
'scale back the VCO sec scale to date scale for output
MyDateValuek2 = MyDateValuek1
MyDateValuek1 = MyDateValuek0
'MyDateValuek0 is the next date predicted value for the next sample
MyDateValuek0 = MyDateRefValue + (MyVCOk0 / NUMBER_SEC_DAY)
'returning the previous value ensure that the TPLL loop delay is zero
'This is necessary to track the time without an offset
'MyDateValuek0 is the predicted value for the next sample
Run = MyDateValuek1
End Function

Public Function ValuePrevious() As Date
ValuePrevious = MyDateValuek2
End Function

Public Function ValueActuel() As Date
ValueActuel = MyDateValuek1
End Function

Public Function ValueNext() As Date
ValueNext = MyDateValuek0
End Function

Public Sub Reset( _
  ByVal FreqDigital As Double, _
  ByVal DampingFactor As Double, _
  ByVal VCOPeriod As Double)

'note FreqDigital is the ratio of fn/fs where fn is the
'natural frequency of the filter and Fs the sampling rate
'
IsFilterReset = True
C = VCOPeriod
C2 = 2 * DampingFactor * TwoPI * FreqDigital
C1 = (C2 ^ 2) / (4 * (DampingFactor ^ 2))
'check stability

If Not (((2 * C2 - 4) < C1) And (C1 < C2) And (C1 > 0)) Then
  'loop is unstable
  MsgBox "GPS TPLL is not stable..."
End If

MyDateDelay = 0
MyTimeErrork0 = 0
MyTimeErrork1 = 0
MyDateValuek2 = 0
MyVCOPeriod = VCOPeriod
End Sub


