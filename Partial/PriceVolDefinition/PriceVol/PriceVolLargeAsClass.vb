Public Class PriceVolLargeAsClass
  Implements IPriceVolLarge
#Region "New"
  Public Sub New()
  End Sub

  Public Sub New(ByVal PriceValue As Double)
    Me.Open = PriceValue
    Me.OpenNext = PriceValue
    Me.Low = PriceValue
    Me.High = PriceValue
    Me.Last = PriceValue
    Me.LastPrevious = PriceValue
    Me.LastWeighted = PriceValue
    Me.FilterLast = PriceValue
  End Sub

  Public Sub New(ByVal PriceValue As PriceVol)
    With Me
      .DateLastTrade = PriceValue.DateLastTrade
      .Open = PriceValue.Open
      .OpenNext = PriceValue.OpenNext
      .Last = PriceValue.Last
      .LastPrevious = PriceValue.LastPrevious
      .High = PriceValue.High
      .Low = PriceValue.Low
      .LastWeighted = PriceValue.LastWeighted
      .Vol = PriceValue.Vol
      .OneyrTargetPrice = PriceValue.OneyrTargetPrice
      .OneyrTargetEarning = PriceValue.OneyrTargetEarning
      .OneyrTargetEarningGrow = PriceValue.OneyrTargetEarningGrow
      .FiveyrTargetEarningGrow = PriceValue.FiveyrTargetEarningGrow
      .OneyrPEG = PriceValue.OneyrPEG
      .FiveyrPEG = PriceValue.FiveyrPEG
      .Range = PriceValue.Range
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = PriceValue.LastAdjusted
      .FilterLast = PriceValue.Last
      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
    End With
  End Sub

  Public Sub New(ByVal PriceValue As IPriceVol)
    Me.New(DirectCast(PriceValue, PriceVol))
  End Sub

  Public Sub New(ByVal PriceValue As PriceVolLarge)
    With Me
      .DateLastTrade = PriceValue.DateLastTrade
      .Open = PriceValue.Open
      .OpenNext = PriceValue.OpenNext
      .Last = PriceValue.Last
      .LastPrevious = PriceValue.LastPrevious
      .High = PriceValue.High
      .Low = PriceValue.Low
      .LastWeighted = PriceValue.LastWeighted
      .Vol = PriceValue.Vol
      .OneyrTargetPrice = PriceValue.OneyrTargetPrice
      .OneyrTargetEarning = PriceValue.OneyrTargetEarning
      .OneyrTargetEarningGrow = PriceValue.OneyrTargetEarningGrow
      .FiveyrTargetEarningGrow = PriceValue.FiveyrTargetEarningGrow
      .OneyrPEG = PriceValue.OneyrPEG
      .FiveyrPEG = PriceValue.FiveyrPEG
      .Range = PriceValue.Range
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = PriceValue.LastAdjusted
      .FilterLast = PriceValue.FilterLast
      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
    End With
  End Sub

	'Public Sub New(ByVal PriceValue As PriceVolAsClass)
	'  With Me
	'    .DateLastTrade = PriceValue.DateLastTrade
	'    .Open = PriceValue.Open
	'    .OpenNext = PriceValue.OpenNext
	'    .Last = PriceValue.Last
	'    .LastPrevious = PriceValue.LastPrevious
	'    .High = PriceValue.High
	'    .Low = PriceValue.Low
	'    .LastWeighted = PriceValue.LastWeighted
	'    .Vol = PriceValue.Vol
	'    .OneyrTargetPrice = PriceValue.OneyrTargetPrice
	'    .OneyrTargetEarning = PriceValue.OneyrTargetEarning
	'    .OneyrTargetEarningGrow = PriceValue.OneyrTargetEarningGrow
	'    .FiveyrTargetEarningGrow = PriceValue.FiveyrTargetEarningGrow
	'    .OneyrPEG = PriceValue.OneyrPEG
	'    .FiveyrPEG = PriceValue.FiveyrPEG
	'    .Range = PriceValue.Range
	'    .RecordQuoteValue = PriceValue.RecordQuoteValue
	'    .IsNull = PriceValue.IsNull
	'    .LastAdjusted = PriceValue.LastAdjusted
	'    .FilterLast = PriceValue.FilterLast
	'    .DividendShare = PriceValue.DividendShare
	'    .DividendYield = PriceValue.DividendYield
	'    .DividendPayDate = PriceValue.DividendPayDate
	'    .ExDividendDate = PriceValue.ExDividendDate
	'    .EarningsShare = PriceValue.EarningsShare
	'    .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
	'    .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
	'    .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
	'  End With
	'End Sub

	Public Sub New(ByVal PriceValue As PriceVolLargeAsClass)
    With Me
      .DateLastTrade = PriceValue.DateLastTrade
      .Open = PriceValue.Open
      .OpenNext = PriceValue.OpenNext
      .Last = PriceValue.Last
      .LastPrevious = PriceValue.LastPrevious
      .High = PriceValue.High
      .Low = PriceValue.Low
      .LastWeighted = PriceValue.LastWeighted
      .Vol = PriceValue.Vol
      .OneyrTargetPrice = PriceValue.OneyrTargetPrice
      .OneyrTargetEarning = PriceValue.OneyrTargetEarning
      .OneyrTargetEarningGrow = PriceValue.OneyrTargetEarningGrow
      .FiveyrTargetEarningGrow = PriceValue.FiveyrTargetEarningGrow
      .OneyrPEG = PriceValue.OneyrPEG
      .FiveyrPEG = PriceValue.FiveyrPEG
      .Range = PriceValue.Range
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = PriceValue.LastAdjusted
      .FilterLast = PriceValue.FilterLast
      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
    End With
  End Sub

  Public Sub New(ByVal PriceValue As IPriceVolLarge)
    Me.New(DirectCast(PriceValue, PriceVolLarge))
  End Sub
#End Region
#Region "Main properties"
  Public DateLastTrade As Date
  Public Open As Double
  Public OpenNext As Double
  Public Last As Double
  Public LastPrevious As Double
  Public High As Double
  Public Low As Double
  Public LastWeighted As Double
  Public Vol As Integer
  Public VolPlus As Integer
  Public VolMinus As Integer
  Public IsIntraDay As Boolean
  Public OneyrTargetPrice As Double
  Public OneyrTargetEarning As Double
  Public OneyrTargetEarningGrow As Double
  Public FiveyrTargetEarningGrow As Double
  Public OneyrPEG As Double
  Public FiveyrPEG As Double
  Public Range As Double
  Public RecordQuoteValue As IRecordQuoteValue
  Public IsNull As Boolean
  Public LastAdjusted As Double
  Public FilterLast As Double
  Public DividendShare As Double
  Public DividendYield As Double
  Public DividendPayDate As Date
  Public ExDividendDate As Date
  Public ExDividendDatePrevious As Date
  Public ExDividendDateEstimated As Date
  Public EarningsShare As Double
  Public EPSEstimateCurrentYear As Double
  Public EPSEstimateNextQuarter As Double
  Public EPSEstimateNextYear As Double

  Public Overrides Function ToString() As String
    Return String.Format("{0},Open:{1},High:{2},Low:{3},Last:{4},OpenNext:{5},Vol:{6},IsNull:{7}", TypeName(Me), Me.Open, Me.High, Me.Low, Me.Last, Me.OpenNext, Me.Vol, Me.IsNull)
  End Function

  Public Function CopyFrom() As PriceVolLarge
    Return New PriceVolLarge(Me)
  End Function

  Public Function CopyFromAsClass() As PriceVolLargeAsClass
    Return New PriceVolLargeAsClass(Me)
  End Function
#End Region
#Region "IPriceVolLarge"
  Public ReadOnly Property AsIPriceVolLarge As IPriceVolLarge Implements IPriceVolLarge.AsIPriceVolLarge
    Get
      Return Me
    End Get
  End Property

  Private Property IPriceVolLarge_DateDay As Date Implements IPriceVolLarge.DateDay
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)
      Me.DateLastTrade = value
    End Set
  End Property
  Private Property IPriceVolLarge_DateUpdate As Date Implements IPriceVolLarge.DateUpdate
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)
      Me.DateLastTrade = value
    End Set
  End Property
  Private Property IPriceVolLarge_High As Double Implements IPriceVolLarge.High
    Get
      Return Me.High
    End Get
    Set(value As Double)
      Me.High = value
    End Set
  End Property
  Private Property IPriceVolLarge_Last As Double Implements IPriceVolLarge.Last
    Get
      Return Me.Last
    End Get
    Set(value As Double)
      Me.Last = value
    End Set
  End Property
  Private Property IPriceVolLarge_LastWeighted As Double Implements IPriceVolLarge.LastWeighted
    Get
      Return Me.LastWeighted
    End Get
    Set(value As Double)
      Me.LastWeighted = value
    End Set
  End Property

  Private Property IPriceVolLarge_LastAdjusted As Double Implements IPriceVolLarge.LastAdjusted
    Get
      Return Me.LastAdjusted
    End Get
    Set(value As Double)
      Me.LastAdjusted = value
    End Set
  End Property
  Private Property IPriceVolLarge_Low As Double Implements IPriceVolLarge.Low
    Get
      Return Me.Low
    End Get
    Set(value As Double)
      Me.Low = value
    End Set
  End Property
  Private Property IPriceVolLarge_Open As Double Implements IPriceVolLarge.Open
    Get
      Return Me.Open
    End Get
    Set(value As Double)
      Me.Open = value
    End Set
  End Property

  Private Property IPriceVolLarge_OpenNext As Double Implements IPriceVolLarge.OpenNext
    Get
      Return Me.OpenNext
    End Get
    Set(value As Double)
      Me.OpenNext = value
    End Set
  End Property

  Private Property IPriceVolLarge_Vol As Integer Implements IPriceVolLarge.Vol
    Get
      Return Me.Vol
    End Get
    Set(value As Integer)
      Me.Vol = value
    End Set
  End Property

  Private Property IPriceVolLarge_Range As Double Implements IPriceVolLarge.Range
    Get
      Return Me.Range
    End Get
    Set(value As Double)
      Me.Range = value
    End Set
  End Property

  Private Property IPriceVolLarge_LastPrevious As Double Implements IPriceVolLarge.LastPrevious
    Get
      Return Me.LastPrevious
    End Get
    Set(value As Double)
      Me.LastPrevious = value
    End Set
  End Property

  Private Property IPriceVolLarge_FilterLast As Double Implements IPriceVolLarge.FilterLast
    Get
      Return Me.FilterLast
    End Get
    Set(value As Double)
      Me.FilterLast = value
    End Set
  End Property
  Public Property IPriceVolLarge_IsIntraDay As Boolean Implements IPriceVolLarge.IsIntraDay
    Get
      Return Me.IsIntraDay
    End Get
    Set(value As Boolean)
      Me.IsIntraDay = value
    End Set
  End Property

  Public Property IPriceVolLarge_VolMinus As Integer Implements IPriceVolLarge.VolMinus
    Get
      Return Me.VolMinus
    End Get
    Set(value As Integer)
      Me.VolMinus = value
    End Set
  End Property
  Public Property IPriceVolLarge_VolPlus As Integer Implements IPriceVolLarge.VolPlus
    Get
      Return Me.VolPlus
    End Get
    Set(value As Integer)
      Me.VolPlus = value
    End Set
  End Property
#End Region
End Class
