Public Class StatisticRangeExcess
  Private MyNumberPoint As Integer
  Private MyNumberPointDouble As Integer
  Private MyQueueOfDailyRangeSigmaDoubleExcess As Queue(Of Integer)
  Private MyQueueOfDailyRangeSigmaExcess As Queue(Of Integer)
  Private MyQueueOfDailyRangeSigmaLowExcess As Queue(Of Integer)
  Private MyQueueOfDailyRangeSigmaHighExcess As Queue(Of Integer)

  Private MySumOfDailyRangeSigmaDoubleExcess As Integer
  Private MySumOfDailyRangeSigmaExcess As Integer
  Private MySumOfDailyRangeSigmaLowExcess As Integer
  Private MySumOfDailyRangeSigmaHighExcess As Integer

  Private MyBandLimitExcess As Double
  Private MyBandLimitDoubleExcess As Double
  Private MyBandLimitLowExcess As Double
  Private MyBandLimitHighExcess As Double

  Private IsBandExceededLocal As Boolean
  Private IsBandExceededLowLocal As Boolean
  Private IsBandExceededHighLocal As Boolean
  Private IsBandExceededDoubleLocal As Boolean

  Public Sub New(ByVal NumberPoint As Integer)
    MyNumberPoint = NumberPoint
    MyNumberPointDouble = 2 * MyNumberPoint
    MyQueueOfDailyRangeSigmaExcess = New Queue(Of Integer)(capacity:=NumberPoint)
    MyQueueOfDailyRangeSigmaLowExcess = New Queue(Of Integer)(capacity:=NumberPoint)
    MyQueueOfDailyRangeSigmaHighExcess = New Queue(Of Integer)(capacity:=NumberPoint)
    MyQueueOfDailyRangeSigmaDoubleExcess = New Queue(Of Integer)(capacity:=MyNumberPointDouble)
  End Sub

  Public Sub Run(ByVal Value As IPriceVol, ByVal LimitLow As Double, ByVal LimitHigh As Double)
    Dim IsBandExceededLocalLast As Boolean
    Dim ThisNumberPoint As Integer

    If MyQueueOfDailyRangeSigmaExcess.Count = MyNumberPoint Then
      'remove the last value from the sum
      MySumOfDailyRangeSigmaExcess = MySumOfDailyRangeSigmaExcess - MyQueueOfDailyRangeSigmaExcess.Dequeue
      MySumOfDailyRangeSigmaLowExcess = MySumOfDailyRangeSigmaLowExcess - MyQueueOfDailyRangeSigmaLowExcess.Dequeue
      MySumOfDailyRangeSigmaHighExcess = MySumOfDailyRangeSigmaHighExcess - MyQueueOfDailyRangeSigmaHighExcess.Dequeue
    End If
    If MyQueueOfDailyRangeSigmaDoubleExcess.Count = MyNumberPointDouble Then
      MySumOfDailyRangeSigmaDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess - MyQueueOfDailyRangeSigmaDoubleExcess.Dequeue
    End If
    IsBandExceededLocalLast = IsBandExceededLocal
    IsBandExceededLocal = False
    IsBandExceededLowLocal = False
    IsBandExceededHighLocal = False
    IsBandExceededDoubleLocal = False
    'note that for now if the high and low limit is exceeded the sum increase by 2 rather than by one
    If (Value.High > LimitHigh) Then
      MySumOfDailyRangeSigmaHighExcess = MySumOfDailyRangeSigmaHighExcess + 1
      IsBandExceededLocal = True
      IsBandExceededHighLocal = True
      MyQueueOfDailyRangeSigmaHighExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaHighExcess.Enqueue(0)
    End If
    If Value.Low < LimitLow Then
      MySumOfDailyRangeSigmaLowExcess = MySumOfDailyRangeSigmaLowExcess + 1
      IsBandExceededLocal = True
      IsBandExceededLowLocal = True
      MyQueueOfDailyRangeSigmaLowExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaLowExcess.Enqueue(0)
    End If
    If IsBandExceededLocal Then
      MySumOfDailyRangeSigmaExcess = MySumOfDailyRangeSigmaExcess + 1
      MyQueueOfDailyRangeSigmaExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaExcess.Enqueue(0)
    End If
    If IsBandExceededLocalLast And IsBandExceededLocal Then
      'two consecutive limit reach
      MySumOfDailyRangeSigmaDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess + 1
      IsBandExceededDoubleLocal = True
      MyQueueOfDailyRangeSigmaDoubleExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaDoubleExcess.Enqueue(0)
    End If
    'calculate the probability
    ThisNumberPoint = MyQueueOfDailyRangeSigmaExcess.Count
    MyBandLimitExcess = MySumOfDailyRangeSigmaExcess / ThisNumberPoint
    MyBandLimitLowExcess = MySumOfDailyRangeSigmaLowExcess / ThisNumberPoint
    MyBandLimitHighExcess = MySumOfDailyRangeSigmaHighExcess / ThisNumberPoint
    ThisNumberPoint = MyQueueOfDailyRangeSigmaDoubleExcess.Count
    MyBandLimitDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess / ThisNumberPoint
  End Sub

  Public Sub Run(ByVal IsLimitLowExcess As Boolean, ByVal IsLimitHighExcess As Boolean)
    Dim IsBandExceededLocalLast As Boolean
    Dim ThisNumberPoint As Integer

    If MyQueueOfDailyRangeSigmaExcess.Count = MyNumberPoint Then
      'remove the last value from the sum
      MySumOfDailyRangeSigmaExcess = MySumOfDailyRangeSigmaExcess - MyQueueOfDailyRangeSigmaExcess.Dequeue
      MySumOfDailyRangeSigmaLowExcess = MySumOfDailyRangeSigmaLowExcess - MyQueueOfDailyRangeSigmaLowExcess.Dequeue
      MySumOfDailyRangeSigmaHighExcess = MySumOfDailyRangeSigmaHighExcess - MyQueueOfDailyRangeSigmaHighExcess.Dequeue
    End If
    If MyQueueOfDailyRangeSigmaDoubleExcess.Count = MyNumberPointDouble Then
      MySumOfDailyRangeSigmaDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess - MyQueueOfDailyRangeSigmaDoubleExcess.Dequeue
    End If
    IsBandExceededLocalLast = IsBandExceededLocal
    IsBandExceededLocal = False
    IsBandExceededLowLocal = False
    IsBandExceededHighLocal = False
    IsBandExceededDoubleLocal = False
    'note that for now if the high and low limit is exceeded the sum increase by 2 rather than by one
    If IsLimitHighExcess Then
      MySumOfDailyRangeSigmaHighExcess = MySumOfDailyRangeSigmaHighExcess + 1
      IsBandExceededLocal = True
      IsBandExceededHighLocal = True
      MyQueueOfDailyRangeSigmaHighExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaHighExcess.Enqueue(0)
    End If
    If IsLimitLowExcess Then
      MySumOfDailyRangeSigmaLowExcess = MySumOfDailyRangeSigmaLowExcess + 1
      IsBandExceededLocal = True
      IsBandExceededLowLocal = True
      MyQueueOfDailyRangeSigmaLowExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaLowExcess.Enqueue(0)
    End If
    If IsBandExceededLocal Then
      MySumOfDailyRangeSigmaExcess = MySumOfDailyRangeSigmaExcess + 1
      MyQueueOfDailyRangeSigmaExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaExcess.Enqueue(0)
    End If
    If IsBandExceededLocalLast And IsBandExceededLocal Then
      'two consecutive limit reach
      MySumOfDailyRangeSigmaDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess + 1
      IsBandExceededDoubleLocal = True
      MyQueueOfDailyRangeSigmaDoubleExcess.Enqueue(1)
    Else
      MyQueueOfDailyRangeSigmaDoubleExcess.Enqueue(0)
    End If
    'calculate the probability
    ThisNumberPoint = MyQueueOfDailyRangeSigmaExcess.Count
    MyBandLimitExcess = MySumOfDailyRangeSigmaExcess / ThisNumberPoint
    MyBandLimitLowExcess = MySumOfDailyRangeSigmaLowExcess / ThisNumberPoint
    MyBandLimitHighExcess = MySumOfDailyRangeSigmaHighExcess / ThisNumberPoint
    ThisNumberPoint = MyQueueOfDailyRangeSigmaDoubleExcess.Count
    MyBandLimitDoubleExcess = MySumOfDailyRangeSigmaDoubleExcess / ThisNumberPoint
  End Sub

  ReadOnly Property BandLimitExcess As Double
    Get
      Return MyBandLimitExcess
    End Get
  End Property

  ReadOnly Property BandLimitDoubleExcess As Double
    Get
      Return MyBandLimitDoubleExcess
    End Get
  End Property

  ReadOnly Property BandLimitLowExcess As Double
    Get
      Return MyBandLimitLowExcess
    End Get
  End Property

  ReadOnly Property BandLimitHighExcess As Double
    Get
      Return MyBandLimitHighExcess
    End Get
  End Property

  ReadOnly Property BandLimitHighLowBalance As Double
    Get
      Return MyBandLimitHighExcess - MyBandLimitLowExcess
    End Get
  End Property

  Public ReadOnly Property IsBandExceeded As Boolean
    Get
      Return IsBandExceededLocal
    End Get
  End Property

  Public ReadOnly Property IsBandExceededDouble As Boolean
    Get
      Return IsBandExceededDoubleLocal
    End Get
  End Property

  Public ReadOnly Property IsBandExceededHigh As Boolean
    Get
      Return IsBandExceededHighLocal
    End Get
  End Property

  Public ReadOnly Property IsBandExceededLow As Boolean
    Get
      Return IsBandExceededLowLocal
    End Get
  End Property
End Class
