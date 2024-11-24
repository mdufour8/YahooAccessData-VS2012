#Region "PriceVol"
Public Structure PriceVol
  Implements IPriceVol
  Implements IPriceVolLarge
  Implements IPricePivotPoint
  Implements ISentimentIndicator

#Region "New"
  Public Sub New(ByVal PriceValue As Single)
    Me.Open = PriceValue
    Me.OpenNext = PriceValue
    Me.Low = PriceValue
    Me.High = PriceValue
    Me.Last = PriceValue
    Me.LastPrevious = PriceValue
    Me.LastWeighted = PriceValue
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
      .FilterLast = PriceValue.FilterLast

      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear

      .AsISentimentIndicator.Count = PriceValue.AsISentimentIndicator.Count
      .AsISentimentIndicator.Value = PriceValue.AsISentimentIndicator.Value
    End With
  End Sub


  ''' <summary>
  ''' Note that IPriceVol is a subset of PriceVol data and not all parameter are updated when using this 
  ''' interface
  ''' </summary>
  ''' <param name="PriceValue"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal PriceValue As IPriceVol)
    With Me
      .DateLastTrade = PriceValue.DateDay
      .Open = PriceValue.Open
      .OpenNext = PriceValue.OpenNext
      .Last = PriceValue.Last
      .LastPrevious = PriceValue.LastPrevious
      .High = PriceValue.High
      .Low = PriceValue.Low
      .LastWeighted = PriceValue.LastWeighted
      .Vol = PriceValue.Vol
      .Range = PriceValue.Range
      .LastAdjusted = PriceValue.LastAdjusted
      If TypeOf PriceValue Is ISentimentIndicator Then
        .AsISentimentIndicator.Count = DirectCast(PriceValue, ISentimentIndicator).Count
        .AsISentimentIndicator.Value = DirectCast(PriceValue, ISentimentIndicator).Value
      End If
    End With
  End Sub

  Public Sub New(ByVal PriceValue As PriceVolLarge)
    With Me
      .DateLastTrade = PriceValue.DateLastTrade
      .Open = CSng(PriceValue.Open)
      .OpenNext = CSng(PriceValue.OpenNext)
      .Last = CSng(PriceValue.Last)
      .LastPrevious = CSng(PriceValue.LastPrevious)
      .High = CSng(PriceValue.High)
      .Low = CSng(PriceValue.Low)
      .LastWeighted = CSng(PriceValue.LastWeighted)
      .Vol = PriceValue.Vol
      .OneyrTargetPrice = CSng(PriceValue.OneyrTargetPrice)
      .OneyrTargetEarning = CSng(PriceValue.OneyrTargetEarning)
      .OneyrTargetEarningGrow = CSng(PriceValue.OneyrTargetEarningGrow)
      .FiveyrTargetEarningGrow = CSng(PriceValue.FiveyrTargetEarningGrow)
      .OneyrPEG = CSng(PriceValue.OneyrPEG)
      .FiveyrPEG = CSng(PriceValue.FiveyrPEG)
      .Range = CSng(PriceValue.Range)
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = CSng(PriceValue.LastAdjusted)
      .FilterLast = CSng(PriceValue.FilterLast)

      .DividendShare = CSng(PriceValue.DividendShare)
      .DividendYield = CSng(PriceValue.DividendYield)
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .EarningsShare = CSng(PriceValue.EarningsShare)
      .EPSEstimateCurrentYear = CSng(PriceValue.EPSEstimateCurrentYear)
      .EPSEstimateNextQuarter = CSng(PriceValue.EPSEstimateNextQuarter)
      .EPSEstimateNextYear = CSng(PriceValue.EPSEstimateNextYear)
    End With
  End Sub

  Public Sub New(ByVal PriceValue As PriceVolAsClass)
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
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
    End With
    'IsPriceOnHoldLocal = False
  End Sub

  Public Sub New(ByVal PriceValue As PriceVolLargeAsClass)
    With Me
      .DateLastTrade = PriceValue.DateLastTrade
      .Open = CSng(PriceValue.Open)
      .OpenNext = CSng(PriceValue.OpenNext)
      .Last = CSng(PriceValue.Last)
      .LastPrevious = CSng(PriceValue.LastPrevious)
      .High = CSng(PriceValue.High)
      .Low = CSng(PriceValue.Low)
      .LastWeighted = CSng(PriceValue.LastWeighted)
      .Vol = PriceValue.Vol
      .OneyrTargetPrice = CSng(PriceValue.OneyrTargetPrice)
      .OneyrTargetEarning = CSng(PriceValue.OneyrTargetEarning)
      .OneyrTargetEarningGrow = CSng(PriceValue.OneyrTargetEarningGrow)
      .FiveyrTargetEarningGrow = CSng(PriceValue.FiveyrTargetEarningGrow)
      .OneyrPEG = CSng(PriceValue.OneyrPEG)
      .FiveyrPEG = CSng(PriceValue.FiveyrPEG)
      .Range = CSng(PriceValue.Range)
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = CSng(PriceValue.LastAdjusted)
      .FilterLast = CSng(PriceValue.FilterLast)

      .DividendShare = CSng(PriceValue.DividendShare)
      .DividendYield = CSng(PriceValue.DividendYield)
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .EarningsShare = CSng(PriceValue.EarningsShare)
      .EPSEstimateCurrentYear = CSng(PriceValue.EPSEstimateCurrentYear)
      .EPSEstimateNextQuarter = CSng(PriceValue.EPSEstimateNextQuarter)
      .EPSEstimateNextYear = CSng(PriceValue.EPSEstimateNextYear)
    End With
    'IsPriceOnHoldLocal = False
  End Sub
#End Region
#Region "Main properties"
  Public DateLastTrade As Date
  Public Open As Single
  Public OpenNext As Single
  Public Last As Single
  Public LastPrevious As Single
  Public High As Single
  Public Low As Single
  Public LastWeighted As Single
  Public Vol As Integer
  Public VolPlus As Integer
  Public VolMinus As Integer
  Public IsIntraDay As Boolean
  Public OneyrTargetPrice As Single
  Public OneyrTargetEarning As Single
  Public OneyrTargetEarningGrow As Single
  Public FiveyrTargetEarningGrow As Single
  Public OneyrPEG As Single
  Public FiveyrPEG As Single
  Public Range As Single
  Public RecordQuoteValue As IRecordQuoteValue
  Public IsNull As Boolean
  Public LastAdjusted As Single
  Public FilterLast As Single
  Public DividendShare As Single
  Public DividendYield As Single
  Public DividendPayDate As Date
  Public ExDividendDate As Date
  Public ExDividendDatePrevious As Date
  Public ExDividendDateEstimated As Date
  Public EarningsShare As Single
  Public EPSEstimateCurrentYear As Single
  Public EPSEstimateNextQuarter As Single
  Public EPSEstimateNextYear As Single
  Public Property IsSpecialDividendPayout As Boolean
  Public Property SpecialDividendPayoutValue As Single

  Public Overrides Function ToString() As String
    Return String.Format("{0},Open:{1},High:{2},Low:{3},Last:{4},OpenNext:{5},Vol:{6},IsNull:{7}", TypeName(Me), Me.Open, Me.High, Me.Low, Me.Last, Me.OpenNext, Me.Vol, Me.IsNull)
  End Function

  Public Function CopyFrom() As PriceVol
    Return New PriceVol(Me)
  End Function

  Public Function CopyFromAsRecord() As Record
    Dim ThisRecord As New Record
    Dim ThisQuoteValue = New QuoteValue
    With ThisRecord
      .DateLastTrade = Me.DateLastTrade
      .DateDay = .DateLastTrade.Date
      .DateUpdate = .DateLastTrade
      .High = Me.High
      .Last = Me.Last
      .Low = Me.Low
      .Open = Me.Open
      .Vol = Me.Vol
      .IsSpecialDividendPayout = Me.IsSpecialDividendPayout
      .SpecialDividendPayoutValue = Me.SpecialDividendPayoutValue
      .VolMinus = Me.VolMinus
      .VolPlus = Me.VolPlus
    End With
    With ThisQuoteValue
      .DateUpdate = ThisRecord.DateUpdate
      .OneyrTargetPrice = Me.OneyrTargetPrice
      .LastTradeDate = .DateUpdate
      .LastTradePriceOnly = ThisRecord.Last
      .PEGRatio = Me.OneyrPEG
      .EarningsShare = Me.EarningsShare
      .DividendShare = Me.DividendShare
      .DividendYield = Me.DividendYield
      .DividendPayDate = Me.DividendPayDate
      .ExDividendDate = Me.ExDividendDate
      .EPSEstimateCurrentYear = Me.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = Me.EPSEstimateNextQuarter
      .EPSEstimateNextYear = Me.EPSEstimateNextYear
    End With
    ThisRecord.QuoteValues.Add(ThisQuoteValue)
    Return ThisRecord
  End Function

  Public Function CopyFromAsClass() As PriceVolAsClass
    Return New PriceVolAsClass(Me)
  End Function
#End Region
#Region "Arithmetic Properties"
  Public Sub MultiPly(ByVal Ratio As Single)
    Dim ThisVolMultiply As Double
    With Me
      .Open = Ratio * .Open
      .OpenNext = Ratio * .OpenNext
      .Last = Ratio * .Last
      .LastPrevious = Ratio * .LastPrevious
      .High = Ratio * .High
      .Low = Ratio * .Low
      'treat the vol as a constant of P*V equivalent to a split
      If Ratio > 0 Then
        ThisVolMultiply = .Vol / Ratio
        If ThisVolMultiply > Integer.MaxValue Then
          ThisVolMultiply = Integer.MaxValue
        End If
        .Vol = CInt(ThisVolMultiply)
      Else
        'do nothing
      End If
      .OneyrTargetPrice = Ratio * .OneyrTargetPrice
      .OneyrTargetEarning = Ratio * .OneyrTargetEarning
      .OneyrTargetEarningGrow = Ratio * .OneyrTargetEarningGrow
      .FiveyrTargetEarningGrow = Ratio * .FiveyrTargetEarningGrow
      .Range = Ratio * .Range
      .DividendShare = Ratio * DividendShare
      'dividend yield do not change with a multiplication i.e
      '.DividendYield = .DividendYield
      .EarningsShare = Ratio * .EarningsShare
      .EPSEstimateCurrentYear = Ratio * .EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = Ratio * .EPSEstimateNextQuarter
      .EPSEstimateNextYear = Ratio * .EPSEstimateNextYear
    End With
  End Sub

  ''' <summary>
  ''' hold the value to the previous last price value
  ''' </summary>
  Public Sub PriceVolHold(ByVal ValueLast As Single)
    Me.PriceVolHold(ValueLast, ValueLast, ValueLast)
  End Sub

  Public Sub PriceVolHold(ByVal ValueLast As Single, ByVal ValueOfOpenNext As Single, ByVal ValueOfLastPrevious As Single)
    With Me
      .Open = ValueLast
      .OpenNext = ValueOfOpenNext
      .Last = ValueLast
      .LastPrevious = ValueOfLastPrevious
      .High = ValueLast
      .Low = ValueLast
      .Vol = 0
      .Range = 0.0
    End With
  End Sub



  ''' <summary>
  ''' Mostly used for the purpose of transforming daily data to weekly price.
  ''' This function change the price High, low, last, Range,Volume and date value of the current object.
  ''' </summary>
  ''' <param name="PriceVolIn">
  ''' PriceVolIn need to be by increasing date value otherwise it return false
  ''' </param>
  ''' <returns>
  ''' return True if the merging is succesful in a given week otherwise false and the data is left unchanged
  ''' </returns>
  Public Function AddToWeeklyMerge(ByRef PriceVolIn As IPriceVol) As Boolean
    Return PriceVol.AddToWeeklyMerge(PriceVolIn, Me)
  End Function

  ''' <summary>
  ''' Mostly used for the purpose of transforming daily data to weekly price.
  ''' This function only change the price High, low, last, Range,Volume and date value of the current object.
  ''' 'It does not affect the current object
  ''' </summary>
  ''' <param name="PriceVolIn">The PriceVol to merge in the PriceVolResult </param>
  ''' <param name="PriceVolResult">The result of teh merging</param>
  ''' <returns>
  ''' Return True if the merging is succesful in a given week otherwise false and the data is left unchanged
  ''' </returns>
  Public Shared Function AddToWeeklyMerge(ByRef PriceVolIn As IPriceVol, ByRef PriceVolResult As IPriceVol) As Boolean
    Dim ThisVolAdd As Long
    Dim ThisDateForFriday As Date = ReportDate.DateToFriday(PriceVolResult.DateUpdate).Date
    With PriceVolResult
      If PriceVolIn.DateUpdate < .DateUpdate Then Return False
      If PriceVolIn.DateUpdate.Date > ThisDateForFriday Then Return False
      If PriceVolIn.High > .High Then
        .High = PriceVolIn.High
      End If
      If PriceVolIn.Low < .Low Then
        .Low = PriceVolIn.Low
      End If
      .Last = PriceVolIn.Last
      .OpenNext = PriceVolIn.OpenNext
      ThisVolAdd = CLng(PriceVolIn.Vol) + CLng(.Vol)
      If ThisVolAdd > Integer.MaxValue Then
        ThisVolAdd = Integer.MaxValue
      End If
      .DateUpdate = PriceVolIn.DateUpdate
      .Vol = CInt(ThisVolAdd)
    End With
    PriceVolResult.Range = RecordPrices.CalculateTrueRange(PriceVolResult)
    Return True
  End Function

  ''' <summary>
  ''' Mostly used for the purpose of transforming daily data to weekly price
  ''' this function only change the price High, low, last, Range,Volume and date value.
  ''' </summary>
  ''' <param name="PriceVol">PriceVol need to be by increasing date value otherwise it return false</param>
  ''' <returns>return True if the merging is succesful in a given week otherwise false and the data is left unchanged</returns>
  Public Function AddToWeeklyMerge(ByRef PriceVol As PriceVol) As Boolean
    Dim ThisVolAdd As Long
    Dim ThisDateForFriday As Date = ReportDate.DateToFriday(Me.DateLastTrade).Date
    With Me
      If PriceVol.DateLastTrade < Me.DateLastTrade Then Return False
      If PriceVol.DateLastTrade.Date > ThisDateForFriday Then Return False
      If PriceVol.High > .High Then
        .High = PriceVol.High
      End If
      If PriceVol.Low > .Low Then
        .Low = PriceVol.Low
      End If
      .Last = PriceVol.Last
      ThisVolAdd = CLng(PriceVol.Vol) + CLng(.Vol)
      If ThisVolAdd > Integer.MaxValue Then
        ThisVolAdd = Integer.MaxValue
      End If
      .DateLastTrade = PriceVol.DateLastTrade
      .Vol = CInt(ThisVolAdd)
      Me.Range = RecordPrices.CalculateTrueRange(Me.AsIPriceVol)
      Return True
    End With
  End Function

  Public Sub Clear(ByRef PriceVol As PriceVol)
    With Me
      .Open = 0.0
      .OpenNext = 0.0
      .LastPrevious = 0.0
      .High = 0.0
      'it can be shown that 
      .DividendYield = 0.0
      .Last = 0.0
      .Low = 0.0
      .Vol = 0
      .OneyrTargetPrice = 0.0
      .OneyrTargetEarning = 0.0
      .OneyrTargetEarningGrow = 0.0
      .FiveyrTargetEarningGrow = 0.0
      .Range = 0.0
      .DividendShare = 0.0
      .EarningsShare = 0.0
      .EPSEstimateCurrentYear = 0.0
      .EPSEstimateNextQuarter = 0.0
      .EPSEstimateNextYear = 0.0
    End With
  End Sub
  Public Sub Add(ByRef PriceVol As PriceVol)
    Dim Temp As Single
    Dim ThisVolAdd As Long
    With Me
      .Open = PriceVol.Open + .Open
      .OpenNext = PriceVol.OpenNext + .OpenNext

      .LastPrevious = PriceVol.LastPrevious + .LastPrevious
      .High = PriceVol.High + .High
      'it can be shown that 
      'DYt=(P1*DY1+P2*DY2)/(P1+P2)
      Temp = Me.Last + PriceVol.Last
      If Temp > 0 Then
        .DividendYield = (Me.Last * Me.DividendYield + PriceVol.Last * PriceVol.DividendYield) / (Me.Last + PriceVol.Last)
      Else
        .DividendYield = 0.0
      End If
      .Last = Temp
      .Low = PriceVol.Low + .Low
      ThisVolAdd = CLng(PriceVol.Vol) + CLng(.Vol)
      If ThisVolAdd > Integer.MaxValue Then
        ThisVolAdd = Integer.MaxValue
      End If
      .Vol = CInt(ThisVolAdd)
      .OneyrTargetPrice = PriceVol.OneyrTargetPrice + .OneyrTargetPrice
      .OneyrTargetEarning = PriceVol.OneyrTargetEarning + .OneyrTargetEarning
      .OneyrTargetEarningGrow = PriceVol.OneyrTargetEarningGrow + .OneyrTargetEarningGrow
      .FiveyrTargetEarningGrow = PriceVol.FiveyrTargetEarningGrow + .FiveyrTargetEarningGrow
      .Range = PriceVol.Range + .Range
      .DividendShare = PriceVol.DividendShare + .DividendShare

      .EarningsShare = PriceVol.EarningsShare + .EarningsShare
      .EPSEstimateCurrentYear = PriceVol.EPSEstimateCurrentYear + .EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceVol.EPSEstimateNextQuarter + .EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceVol.EPSEstimateNextYear + .EPSEstimateNextYear
    End With
  End Sub
#End Region
#Region "IPriceVol"
  Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
    Get
      Return Me
    End Get
  End Property

  Private Property IPriceVol_DateDay As Date Implements IPriceVol.DateDay
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)
      Me.DateLastTrade = value
    End Set
  End Property
  Private Property IPriceVol_DateUpdate As Date Implements IPriceVol.DateUpdate
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)
      Me.DateLastTrade = value
    End Set
  End Property
  Private Property IPriceVol_High As Single Implements IPriceVol.High
    Get
      Return Me.High
    End Get
    Set(value As Single)
      Me.High = value
    End Set
  End Property
  Private Property IPriceVol_Last As Single Implements IPriceVol.Last
    Get
      Return Me.Last
    End Get
    Set(value As Single)
      Me.Last = value
    End Set
  End Property
  Private Property IPriceVol_LastWeighted As Single Implements IPriceVol.LastWeighted
    Get
      Return Me.LastWeighted
    End Get
    Set(value As Single)
      Me.LastWeighted = value
    End Set
  End Property

  Private Property IPriceVol_LastAdjusted As Single Implements IPriceVol.LastAdjusted
    Get
      Return Me.LastAdjusted
    End Get
    Set(value As Single)
      Me.LastAdjusted = value
    End Set
  End Property
  Private Property IPriceVol_Low As Single Implements IPriceVol.Low
    Get
      Return Me.Low
    End Get
    Set(value As Single)
      Me.Low = value
    End Set
  End Property
  Private Property IPriceVol_Open As Single Implements IPriceVol.Open
    Get
      Return Me.Open
    End Get
    Set(value As Single)
      Me.Open = value
    End Set
  End Property

  Private Property IPriceVol_OpenNext As Single Implements IPriceVol.OpenNext
    Get
      Return Me.OpenNext
    End Get
    Set(value As Single)
      Me.OpenNext = value
    End Set
  End Property

  Private Property IPriceVol_Vol As Integer Implements IPriceVol.Vol
    Get
      Return Me.Vol
    End Get
    Set(value As Integer)
      Me.Vol = value
    End Set
  End Property

  Private Property IPriceVol_Range As Single Implements IPriceVol.Range
    Get
      Return Me.Range
    End Get
    Set(value As Single)
      Me.Range = value
    End Set
  End Property

  Private Property IPriceVol_LastPrevious As Single Implements IPriceVol.LastPrevious
    Get
      Return Me.LastPrevious
    End Get
    Set(value As Single)
      Me.LastPrevious = value
    End Set
  End Property

  Public Property IPriceVol_IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
    Get
      Return Me.IsIntraDay
    End Get
    Set(value As Boolean)
      Me.IsIntraDay = value
    End Set
  End Property

  Public Property IPriceVol_VolMinus As Integer Implements IPriceVol.VolMinus
    Get
      Return Me.VolMinus
    End Get
    Set(value As Integer)
      Me.VolMinus = value
    End Set
  End Property

  Public Property IPriceVol_VolPlus As Integer Implements IPriceVol.VolPlus
    Get
      Return Me.VolPlus
    End Get
    Set(value As Integer)
      Me.VolPlus = value
    End Set
  End Property

  Private Property IPriceVol_IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
    Get
      Return Me.IsSpecialDividendPayout
    End Get
    Set(value As Boolean)
      Me.IsSpecialDividendPayout = value
    End Set
  End Property

  Private Property IPriceVol_SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
    Get
      Return Me.SpecialDividendPayoutValue
    End Get
    Set(value As Single)
      SpecialDividendPayoutValue = value
    End Set
  End Property
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
      Me.High = CSng(value)
    End Set
  End Property
  Private Property IPriceVolLarge_Last As Double Implements IPriceVolLarge.Last
    Get
      Return Me.Last
    End Get
    Set(value As Double)
      Me.Last = CSng(value)
    End Set
  End Property
  Private Property IPriceVolLarge_LastWeighted As Double Implements IPriceVolLarge.LastWeighted
    Get
      Return Me.LastWeighted
    End Get
    Set(value As Double)
      Me.LastWeighted = CSng(value)
    End Set
  End Property

  Private Property IPriceVolLarge_LastAdjusted As Double Implements IPriceVolLarge.LastAdjusted
    Get
      Return Me.LastAdjusted
    End Get
    Set(value As Double)
      Me.LastAdjusted = CSng(value)
    End Set
  End Property
  Private Property IPriceVolLarge_Low As Double Implements IPriceVolLarge.Low
    Get
      Return Me.Low
    End Get
    Set(value As Double)
      Me.Low = CSng(value)
    End Set
  End Property
  Private Property IPriceVolLarge_Open As Double Implements IPriceVolLarge.Open
    Get
      Return Me.Open
    End Get
    Set(value As Double)
      Me.Open = CSng(value)
    End Set
  End Property

  Private Property IPriceVolLarge_OpenNext As Double Implements IPriceVolLarge.OpenNext
    Get
      Return Me.OpenNext
    End Get
    Set(value As Double)
      Me.OpenNext = CSng(value)
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
      Me.Range = CSng(value)
    End Set
  End Property

  Private Property IPriceVolLarge_LastPrevious As Double Implements IPriceVolLarge.LastPrevious
    Get
      Return Me.LastPrevious
    End Get
    Set(value As Double)
      Me.LastPrevious = CSng(value)
    End Set
  End Property

  Private Property IPriceVolLarge_FilterLast As Double Implements IPriceVolLarge.FilterLast
    Get
      Return Me.FilterLast
    End Get
    Set(value As Double)
      Me.FilterLast = CSng(value)
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
#Region "IPricePivotPoint"
  ''' <summary>
  ''' see definition:https://www.fidelity.com/learning-center/trading-investing/technical-analysis/technical-indicator-guide/pivot-points-resistance-support
  ''' Calculation:
  ''' Note: The 'previous day' here only mean that the calculated value can be used to predict the next day trading pivot point and support level
  ''' Resistance Level 3 = Previous Day High + 2(Pivot – Previous Day Low)
  ''' Resistance Level 2 = Pivot + (Resistance Level 1 – Support Level 1)
  ''' Resistance Level 1 = (Pivot x 2) – Previous Day Low
  ''' Pivot = Previous Day (High + Low + Close) / 3
  ''' Support Level 1 = (Pivot x 2) – Previous Day High
  ''' Support Level 2 = Pivot – (Resistance Level 1 – Support Level 1)
  ''' Support Level 3 = Previous Day Low – 2(Previous Day High – Pivot)
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property AsIPricePivotPoint As IPricePivotPoint Implements IPricePivotPoint.AsIPricePivotPoint
    Get
      Return Me
    End Get
  End Property



  Private Function IPricePivotPoint_PriceVolPivot(Level As IPricePivotPoint.enuPivotLevel) As IPriceVol Implements IPricePivotPoint.PriceVolPivot
    Dim ThisPriceVol As New PriceVol(Me)

    ThisPriceVol.Last = IPricePivotPoint_PivotLast
    ThisPriceVol.Open = IPricePivotPoint_PivotOpen
    ThisPriceVol.High = IPricePivotPoint_Resistance(Level)
    ThisPriceVol.Low = IPricePivotPoint_Support(Level)
    ThisPriceVol.Range = RecordPrices.CalculateTrueRange(ThisPriceVol.AsIPriceVol)
    Return ThisPriceVol
  End Function

  Private ReadOnly Property IPricePivotPoint_Resistance(Level As IPricePivotPoint.enuPivotLevel) As Single Implements IPricePivotPoint.Resistance
    Get
      Dim ThisResult As Single = 0.0
      Select Case Level
        Case IPricePivotPoint.enuPivotLevel.Level1
          ThisResult = (2 * Me.IPricePivotPoint_PivotLast) - Me.Low
        Case IPricePivotPoint.enuPivotLevel.Level2
          ThisResult = Me.IPricePivotPoint_PivotLast + (Me.High - Me.Low)
        Case IPricePivotPoint.enuPivotLevel.Level3
          ThisResult = Me.High + 2 * (Me.IPricePivotPoint_PivotLast - Me.Low)
      End Select
      Return ThisResult
    End Get
  End Property

  Private ReadOnly Property IPricePivotPoint_Support(Level As IPricePivotPoint.enuPivotLevel) As Single Implements IPricePivotPoint.Support
    Get
      Dim ThisResult As Single
      Select Case Level
        Case IPricePivotPoint.enuPivotLevel.Level1
          ThisResult = 2 * IPricePivotPoint_PivotLast - Me.High
        Case IPricePivotPoint.enuPivotLevel.Level2
          ThisResult = IPricePivotPoint_PivotLast - (Me.High - Me.Low)
        Case IPricePivotPoint.enuPivotLevel.Level3
          ThisResult = Me.Low - 2 * (Me.High - IPricePivotPoint_PivotLast)
      End Select
      Return ThisResult
    End Get
  End Property

  Private ReadOnly Property IPricePivotPoint_PivotLast As Single Implements IPricePivotPoint.PivotLast
    Get
      Return (Me.High + Me.Low + Me.Last) / 3
    End Get
  End Property

  Private ReadOnly Property IPricePivotPoint_PivotOpen As Single Implements IPricePivotPoint.PivotOpen
    Get
      Return (Me.High + Me.Low + Me.Open) / 3
    End Get
  End Property
#End Region
#Region "ISentimentIndicator"
  Public Function AsISentimentIndicator() As ISentimentIndicator Implements ISentimentIndicator.AsISentimentIndicator
    Return Me
  End Function

  Private _ISentimentIndicator_Count As Integer
  Private Property ISentimentIndicator_Count As Integer Implements ISentimentIndicator.Count
    Get
      Return _ISentimentIndicator_Count
    End Get
    Set(value As Integer)
      _ISentimentIndicator_Count = value
    End Set
  End Property

  Private _ISentimentIndicator_Value As Double
  Private Property ISentimentIndicator_Value As Double Implements ISentimentIndicator.Value
    Get
      Return _ISentimentIndicator_Value
    End Get
    Set(value As Double)
      _ISentimentIndicator_Value = value
    End Set
  End Property
#End Region
End Structure
#End Region
#Region "PriceVolLarge"
Public Structure PriceVolLarge
  Implements IPriceVolLarge

#Region "New"
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

      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .OneyrPEG = PriceValue.OneyrPEG
      .FiveyrPEG = PriceValue.FiveyrPEG
      .Range = PriceValue.Range
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = PriceValue.LastAdjusted
      .FilterLast = PriceValue.FilterLast
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

      .DividendShare = PriceValue.DividendShare
      .DividendYield = PriceValue.DividendYield
      .DividendPayDate = PriceValue.DividendPayDate
      .ExDividendDate = PriceValue.ExDividendDate
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDatePrevious

      .OneyrPEG = PriceValue.OneyrPEG
      .FiveyrPEG = PriceValue.FiveyrPEG
      .Range = PriceValue.Range
      .RecordQuoteValue = PriceValue.RecordQuoteValue
      .IsNull = PriceValue.IsNull
      .LastAdjusted = PriceValue.LastAdjusted
      .FilterLast = PriceValue.FilterLast
    End With
  End Sub

  Public Sub New(ByVal PriceValue As PriceVolAsClass)
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
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDateEstimated

      .EarningsShare = PriceValue.EarningsShare
      .EPSEstimateCurrentYear = PriceValue.EPSEstimateCurrentYear
      .EPSEstimateNextQuarter = PriceValue.EPSEstimateNextQuarter
      .EPSEstimateNextYear = PriceValue.EPSEstimateNextYear
    End With
  End Sub

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
      .ExDividendDatePrevious = PriceValue.ExDividendDatePrevious
      .ExDividendDateEstimated = PriceValue.ExDividendDateEstimated

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
      Me.FilterLast = FilterLast
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
End Structure
#End Region

#Region "IPriceVol"
Public Interface IPriceVol
  ReadOnly Property AsIPriceVol As IPriceVol
  Property DateDay As Date
  Property DateUpdate As Date
  Property Open As Single
  Property OpenNext As Single
  Property Last As Single
  Property LastPrevious As Single
  Property LastWeighted As Single
  Property LastAdjusted As Single
  Property High As Single
  Property Low As Single
  Property Vol As Integer
  Property VolPlus As Integer
  Property VolMinus As Integer
  Property IsIntraDay As Boolean
  Property Range As Single
  Property IsSpecialDividendPayout As Boolean
  Property SpecialDividendPayoutValue As Single
End Interface
#End Region

#Region "IPriceVolShort"
Public Interface IPriceVolShort
  ReadOnly Property AsIPriceVol As IPriceVolShort
  Property DateStamp As Date
  Property High As Single
  Property Open As Single
  Property Low As Single
  Property Last As Single
  Property Vol As Integer
End Interface

Public Interface IPriceVolLarge
  ReadOnly Property AsIPriceVolLarge As IPriceVolLarge
  Property DateDay As Date
  Property DateUpdate As Date
  Property Open As Double
  Property OpenNext As Double
  Property Last As Double
  Property LastPrevious As Double
  Property LastWeighted As Double
  Property LastAdjusted As Double
  Property High As Double
  Property Low As Double
  Property Vol As Integer
  Property VolPlus As Integer
  Property VolMinus As Integer
  Property IsIntraDay As Boolean
  Property Range As Double
  Property FilterLast As Double
End Interface

Public Interface IPriceVol(Of T)
  ReadOnly Property AsIPriceVol As IPriceVol(Of T)
  Property DateDay As Date
  Property DateUpdate As Date
  Property Open As T
  Property OpenNext As T
  Property Last As T
  Property FilterLast As T
  Property LastPrevious As T
  Property LastWeighted As T
  Property LastAdjusted As T
  Property High As T
  Property Low As T
  Property Vol As Integer
  Property VolPlus As Integer
  Property VolMinus As Integer
  Property IsIntraDay As Boolean
  Property Range As T
End Interface
#End Region
#Region "IPricePivotPoint"
''' <summary>
''' see definition:https://www.fidelity.com/learning-center/trading-investing/technical-analysis/technical-indicator-guide/pivot-points-resistance-support
''' </summary>
''' <remarks></remarks>
Public Interface IPricePivotPoint
  Enum enuPivotLevel
    Level1
    Level2
    Level3
  End Enum

  ReadOnly Property AsIPricePivotPoint As IPricePivotPoint
  ReadOnly Property PivotOpen As Single
  ReadOnly Property PivotLast As Single
  ReadOnly Property Resistance(ByVal Level As enuPivotLevel) As Single
  ReadOnly Property Support(ByVal Level As enuPivotLevel) As Single
  Function PriceVolPivot(ByVal Level As enuPivotLevel) As IPriceVol
End Interface
#End Region


