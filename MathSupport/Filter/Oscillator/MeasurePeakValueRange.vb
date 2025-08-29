Imports YahooAccessData.MathPlus.Filter

Public Class MeasurePeakValueRange
  Private Const FILTER_RATE_BAND As Integer = 5

  Private MyListOfWindowFrameHigh As ListWindowFrame
  Private MyListOfWindowFrameLow As ListWindowFrame
  Private MyListOfWindowFrameHighMinusOne As ListWindowFrame
  Private MyListOfWindowFrameLowMinusOne As ListWindowFrame
  Private MyFilterHigh As IFilter
  Private MyFilterLow As IFilter
  Private MyListOfPeak As List(Of IPeakValueRange)
  Private IsFilterEnabledLocal As Boolean

  Public Sub New(ByVal FilterRate As Integer)
    Me.New(FilterRate, IsFilterPeakEnabled:=False)
  End Sub

  Public Sub New(ByVal FilterRate As Integer, ByVal IsFilterPeakEnabled As Boolean)
    MyListOfWindowFrameHigh = New ListWindowFrame(FilterRate)
    MyListOfWindowFrameLow = New ListWindowFrame(FilterRate)
    MyListOfWindowFrameHighMinusOne = New ListWindowFrame(FilterRate - 1)
    MyListOfWindowFrameLowMinusOne = New ListWindowFrame(FilterRate - 1)
    IsFilterEnabledLocal = IsFilterPeakEnabled
		'go back to exponential filter because FilterPrediction next does not work and have a bug
		'MyFilterHigh = New FilterLowPassPLL(FilterRate:=FILTER_RATE_BAND, DampingFactor:=1.5, NumberOfPredictionOutput:=0)
		'MyFilterLow = New FilterLowPassPLL(FilterRate:=FILTER_RATE_BAND, DampingFactor:=1.5, NumberOfPredictionOutput:=0)
		MyFilterHigh = New FilterLowPassExp(FilterRate:=FILTER_RATE_BAND)
		MyFilterLow = New FilterLowPassExp(FilterRate:=FILTER_RATE_BAND)
		MyListOfPeak = New List(Of IPeakValueRange)
  End Sub

  Public Function Filter(ByVal ValueLow As Double, ByVal ValueHigh As Double) As IPeakValueRange
    Dim ThisValuePeakHigh As Double
    Dim ThisValuePeakLow As Double
    Dim ThisPeakValueRange As IPeakValueRange

    MyListOfWindowFrameHigh.Add(ValueHigh)
    MyListOfWindowFrameLow.Add(ValueLow)
    MyListOfWindowFrameHighMinusOne.Add(ValueHigh)
    MyListOfWindowFrameLowMinusOne.Add(ValueLow)


    If IsFilterEnabledLocal Then
			ThisValuePeakHigh = MyFilterHigh.Filter(MyListOfWindowFrameHigh.ItemHigh.Value)
			ThisValuePeakLow = MyFilterLow.Filter(MyListOfWindowFrameLow.ItemLow.Value)
      'Only for positive price value do not let the price go below zero
      If ThisValuePeakHigh < 0 Then ThisValuePeakHigh = 0.0
      If ThisValuePeakLow < 0 Then ThisValuePeakLow = 0.0
      ThisPeakValueRange = New PeakValueRange(ThisValuePeakLow, ThisValuePeakHigh)
    Else
      ThisPeakValueRange = New PeakValueRange(MyListOfWindowFrameLow.ItemLow.Value, MyListOfWindowFrameHigh.ItemHigh.Value)
    End If
    MyListOfPeak.Add(ThisPeakValueRange)
    Return ThisPeakValueRange
  End Function

	''' <summary>
	''' Calculate a new Peak estimate based on a new range without changing the actual class memory
	''' </summary>
	''' <param name="ValueLow"></param>
	''' <param name="ValueHigh"></param>
	''' <returns></returns>
	''' <remarks>does not affect any internal memory of the class</remarks>
	Public Function FilterPredictionEstimate(ByVal ValueLow As Double, ByVal ValueHigh As Double) As IPeakValueRange
    Dim ThisPeakPredictionHigh As Double = MyListOfWindowFrameHigh.ItemHigh.Value
    Dim ThisPeakPredictionLow As Double = MyListOfWindowFrameLow.ItemLow.Value
    Dim ThisPeakPredictionHighMinusOne As Double = MyListOfWindowFrameHighMinusOne.ItemHigh.Value
    Dim ThisPeakPredictionLowMinusOne As Double = MyListOfWindowFrameLowMinusOne.ItemLow.Value
    Dim ThisPeakPredictionHighPrediction As Double
    Dim ThisPeakPredictionLowPrediction As Double

    If MyListOfWindowFrameHigh.ItemHighIndex = 0 Then
      'the high value is ready to be drop 
      ThisPeakPredictionHighPrediction = ThisPeakPredictionHighMinusOne
    Else
      ThisPeakPredictionHighPrediction = ThisPeakPredictionHigh
    End If
    If MyListOfWindowFrameLow.ItemLowIndex = 0 Then
      'the low value is ready to be drop 
      ThisPeakPredictionLowPrediction = ThisPeakPredictionLowMinusOne
    Else
      ThisPeakPredictionLowPrediction = ThisPeakPredictionLow
    End If
    If IsFilterEnabledLocal Then
      If ValueHigh > ThisPeakPredictionHighPrediction Then
        ThisPeakPredictionHighPrediction = ValueHigh
        ThisPeakPredictionHighPrediction = MyFilterHigh.FilterPredictionNext(ThisPeakPredictionHighPrediction)
      End If
      If ValueLow < ThisPeakPredictionLowPrediction Then
        ThisPeakPredictionLowPrediction = ValueLow
        ThisPeakPredictionLowPrediction = MyFilterLow.FilterPredictionNext(ThisPeakPredictionLowPrediction)
      End If
    Else
      If ValueHigh > ThisPeakPredictionHighPrediction Then
        ThisPeakPredictionHighPrediction = ValueHigh
      End If
      If ValueLow < ThisPeakPredictionLowPrediction Then
        ThisPeakPredictionLowPrediction = ValueLow
      End If
    End If
    Return New PeakValueRange(ThisPeakPredictionLowPrediction, ThisPeakPredictionHighPrediction)
  End Function

  ReadOnly Property ToList() As IList(Of IPeakValueRange)
    Get
      Return MyListOfPeak
    End Get
  End Property

  Property IsFilterEnabled As Boolean
    Get
      Return IsFilterEnabledLocal
    End Get
    Set(value As Boolean)
      IsFilterEnabledLocal = value
    End Set
  End Property
End Class

Public Class PeakValueRange
  Implements IPeakValueRange

  Private MyValueHigh As Double
  Private MyValueLow As Double
  Private MyValueRange As Double

  Public Sub New(ByVal ValueLow As Double, ByVal ValueHigh As Double)
    MyValueHigh = ValueHigh
    MyValueLow = ValueLow
    MyValueRange = MyValueHigh - MyValueLow
  End Sub


  Public ReadOnly Property High As Double Implements IPeakValueRange.High
    Get
      Return MyValueHigh
    End Get
  End Property

  Public ReadOnly Property Low As Double Implements IPeakValueRange.Low
    Get
      Return MyValueLow
    End Get
  End Property

  Public ReadOnly Property Range As Double Implements IPeakValueRange.Range
    Get
      Return MyValueRange
    End Get
  End Property

  Public Overrides Function ToString() As String
    Return String.Format("{0}Low:{1}, High:{2},{3}", "{", Me.Low, Me.High, "}")
  End Function
End Class
