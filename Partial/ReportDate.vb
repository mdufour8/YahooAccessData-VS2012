#Region "Imports"
Imports YahooAccessData.ExtensionService
Imports YahooAccessData.MathPlus.Measure
#End Region
#Region "IReportDateHoliday"
Public Interface IReportDateHoliday
    Property DateDay As Date
    Property Enabled As Boolean
    Property IsFullDay As Boolean
    Property MarketOpenTime As Date
    Property MarketCloseTime As Date
    Property Description As String
  End Interface
#End Region
#Region "ReportDateHoliday"
  <Serializable()>
  Public Class ReportDateHoliday
    Implements IReportDateHoliday

    Public Sub New()
      Me.DateDay = Now.Date
      Me.Enabled = False
      Me.IsFullDay = True
      Me.MarketOpenTime = Me.DateDay.AddHours(9).AddMinutes(30)
      Me.MarketCloseTime = Me.DateDay.AddHours(16)
      Me.Description = ""
    End Sub

    Public Property DateDay As Date Implements IReportDateHoliday.DateDay
    Public Property Description As String Implements IReportDateHoliday.Description
    Public Property IsFullDay As Boolean Implements IReportDateHoliday.IsFullDay
    Public Property MarketCloseTime As Date Implements IReportDateHoliday.MarketCloseTime
    Public Property MarketOpenTime As Date Implements IReportDateHoliday.MarketOpenTime
    Public Property Enabled As Boolean Implements IReportDateHoliday.Enabled
  End Class
#End Region
#Region "ReportDate"
''' <summary>
''' calculate the US market day trading taking into account the official holiday
''' see: http://www.rightline.net/calendar/market-holidays.html
''' </summary>
''' <remarks></remarks>
Public Class ReportDate
  Public Enum MarketHolidaySchedule
    US
    CDN
  End Enum

  Public Enum MonthsOfYear
    January = 1
    February = 2
    March = 3
    April = 4
    May = 5
    June = 6
    July = 7
    August = 8
    September = 9
    October = 10
    November = 11
    December = 12
  End Enum

  Public Const DATE_NULL_VALUE As Date = #1/1/1900#
  Public Const MARKET_OPEN_TIME_DEFAULT As Date = #9:30:00 AM#
  Public Const MARKET_OPEN_TIME_SEC_DEFAULT As Integer = 3600 * 9 + 1800
  Public Const MARKET_CLOSE_TIME_SEC_DEFAULT As Integer = 3600 * 16
  Public Const MARKET_CLOSE_TIME_DEFAULT As Date = #4:00:00 PM#
  Public Const MARKET_OPEN_TO_CLOSE_PERIOD_HOUR_DEFAULT As Double = 6.5
  Public Const MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT As Double = MARKET_OPEN_TO_CLOSE_PERIOD_HOUR_DEFAULT / 24
  Public Const MARKET_OPEN_TO_CLOSE_PERIOD_SEC_DEFAULT As Double = 3600 * MARKET_OPEN_TO_CLOSE_PERIOD_HOUR_DEFAULT
  Public Const MARKET_PREVIOUS_CLOSE_TO_OPEN_TIME_DAY_DEFAULT As Double = 1.0 - MARKET_OPEN_TO_CLOSE_PERIOD_DAY_DEFAULT



#Region "MarketOpenCloseTime"
  Private Structure MarketOpenCloseTime
    Public Sub New(
      ByVal DateValue As Date,
      ByVal MarketOpenTime As Date,
      ByVal MarketCloseTime As Date,
      ByVal MarketDelayMinutes As Integer,
      ByVal MarketHolidaySchedule As MarketHolidaySchedule)

      Dim ThisHolidaySpecialEvent As ReportDateHoliday
      Me.DateValue = DateValue.Date
      Me.IsSpecialHoliday = False
      If MyDictOfSpecialHoliday Is Nothing Then
        Call FileListLoad()
      End If
      If MyDictOfSpecialHoliday.ContainsKey(Me.DateValue) Then
        ThisHolidaySpecialEvent = MyDictOfSpecialHoliday(Me.DateValue)
        If ThisHolidaySpecialEvent.IsFullDay Then
          Me.IsSpecialHoliday = True
          Me.OpenTime = Me.DateValue.Add(MarketOpenTime.TimeOfDay)
          Me.CloseTime = Me.DateValue.Add(MarketCloseTime.TimeOfDay)
        Else
          'this is a trading day with restricted trading time only
          With ThisHolidaySpecialEvent
            Me.OpenTime = Me.DateValue.Add(.MarketOpenTime.TimeOfDay)
            Me.CloseTime = Me.DateValue.Add(.MarketCloseTime.TimeOfDay)
          End With
        End If
      Else
        Me.OpenTime = Me.DateValue.Add(MarketOpenTime.TimeOfDay)
        Me.CloseTime = Me.DateValue.Add(MarketCloseTime.TimeOfDay)
      End If
      If DateValue.Month = 7 Then
        If DateValue.Day = 3 Then
          'the market close early at 13:00 PM the day before independance day
          Dim ThisTimeSpan = New TimeSpan(hours:=13, minutes:=0, seconds:=0)
          If Me.CloseTime.TimeOfDay > ThisTimeSpan Then
            Me.CloseTime = Me.DateValue.Add(ThisTimeSpan)
          End If
        End If
      End If
      If DateValue.Month = 11 Then
        'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
        'Traditionally, it has been a time to give thanks to God, friends, and family.
        'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
        'and in the United States, it is celebrated on the fourth Thursday of November. 
        'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
        Dim ThisDateForThanksGivingNextDay = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday).AddDays(1)
        If Me.DateValue = ThisDateForThanksGivingNextDay Then
          'The NYSE, AMEX and NASDAQ will close trading early (at 1:00 PM ET) on Friday, the day after Thanksgiving.
          Dim ThisTimeSpan = New TimeSpan(hours:=13, minutes:=0, seconds:=0)
          If Me.CloseTime.TimeOfDay > ThisTimeSpan Then
            Me.CloseTime = Me.DateValue.Add(ThisTimeSpan)
          End If
        End If
      End If
      If DateValue.Month = 12 Then
        If DateValue.Day = 24 Then
          'the market close early at 13:00 PM the day before christmas
          Dim ThisTimeSpan = New TimeSpan(hours:=13, minutes:=0, seconds:=0)
          If Me.CloseTime.TimeOfDay > ThisTimeSpan Then
            Me.CloseTime = Me.DateValue.Add(ThisTimeSpan)
          End If
        End If
      End If
      Me.OpenTimeDelay = Me.OpenTime.AddMinutes(MarketDelayMinutes)
      Me.CloseTimeDelay = Me.CloseTime.AddMinutes(MarketDelayMinutes)
      Me.OpenTimeSpanDelay = Me.OpenTimeDelay.TimeOfDay
      Me.CloseTimeSpanDelay = Me.CloseTimeDelay.TimeOfDay
    End Sub

    Public Sub New(
      ByVal DateValue As Date,
      ByVal MarketOpenTime As Date,
      ByVal MarketCloseTime As Date,
      ByVal MarketDelayMinutes As Integer)

      'by default open the US market
      Me.New(DateValue, MarketOpenTime, MarketCloseTime, MarketDelayMinutes, MarketHolidaySchedule.US)
    End Sub

    Private Sub FileListLoad()
      Dim ThisFile = My.Application.Info.DirectoryPath & "\MarketHoliday\MarketHoliday.xml"
      Dim MyListOfSpecialHolidayKnown = New List(Of ReportDateHoliday)

      'load the default holiday use for testing and validation for Tuesday-Wedesday January 15-16 1980
      MyListOfSpecialHolidayKnown.Add(New ReportDateHoliday With {
        .DateDay = DateSerial(1980, 1, 15),
        .Description = "Testing opening late at 12:00",
        .Enabled = True,
        .IsFullDay = False,
        .MarketOpenTime = .DateDay.AddHours(12)})
      MyListOfSpecialHolidayKnown.Add(New ReportDateHoliday With {
        .DateDay = DateSerial(1980, 1, 16),
        .Description = "Testing closing early at 13:00",
        .Enabled = True,
        .IsFullDay = False,
        .MarketCloseTime = .DateDay.AddHours(13)})

      'load the default holiday that are already known so the user do not have to enter this list
      MyListOfSpecialHolidayKnown.Add(New ReportDateHoliday With {
        .DateDay = DateSerial(2012, 10, 29),
        .Description = "Sandy storm",
        .Enabled = True,
        .IsFullDay = True})
      MyListOfSpecialHolidayKnown.Add(New ReportDateHoliday With {
        .DateDay = DateSerial(2012, 10, 30),
        .Description = "Sandy storm",
        .Enabled = True,
        .IsFullDay = True})

      MyListOfSpecialHoliday = FileListRead(Of ReportDateHoliday)(ThisFile, MyListOfSpecialHolidayKnown, Nothing)
      MyDictOfSpecialHoliday = New Dictionary(Of Date, ReportDateHoliday)
      'for quick searching move all the enabled data in the list in the dictionary
      For Each ThisHoliday In MyListOfSpecialHoliday
        If ThisHoliday.Enabled Then
          MyDictOfSpecialHoliday.Add(ThisHoliday.DateDay, ThisHoliday)
        End If
      Next
    End Sub

    Dim DateValue As Date
    Dim OpenTime As Date
    Dim CloseTime As Date
    Dim OpenTimeDelay As Date
    Dim CloseTimeDelay As Date
    Dim OpenTimeSpanDelay As TimeSpan
    Dim CloseTimeSpanDelay As TimeSpan
    Dim IsSpecialHoliday As Boolean
  End Structure
#End Region
  Private Shared MyListOfSpecialHoliday As List(Of ReportDateHoliday)
  Private Shared MyDictOfSpecialHoliday As Dictionary(Of Date, ReportDateHoliday)

#Region "Date main function"
  ''' <summary>
  ''' This function return the date of market trading days assuming 5 days trading per week and the week optionally starting on monday . Here, DateStart is always 
  ''' synchronized back to the previous Monday. This equation is valid for any range of NumberTradingDays positive or negative relative to DateStart
  ''' </summary>
  ''' <param name="DateStart"></param>
  ''' <param name="NumberTradingDays">
  ''' The number of trading days starting at zero from the Date reference DateStart referenced back to the previous monday
  ''' </param>
  ''' <returns>The market trading date</returns>
  ''' <remarks>This function has been tested using excel</remarks>
  Public Shared Function MarketTradingDeltaDays(ByVal DateStart As Date, ByVal NumberTradingDays As Integer, Optional ByVal IsMondaySync As Boolean = True) As Date
    Dim ThisDeltaDays As Integer
    Dim ThisNumberOfDayFromMonday As Integer

    'make sure the date do not fall on a weekend
    'if it does bring it back to last friday
    'remove the time information
    DateStart = DateToWeekEndRemovePrevious(DateStart.Date)
    'this is equivalent than calling DateToMondayPrevious(DateStart)
    'the transformation is generally non-linear
    'remove the time information
    DateStart = DateStart.Date
    Select Case DateStart.DayOfWeek
      Case DayOfWeek.Monday    '1
        ThisNumberOfDayFromMonday = 0
        'Case DayOfWeek.Sunday   '0
        'this test is not necessary
        'ThisNumberOfDayFromMonday = 6
      Case Else
        'this portion is linear
        ThisNumberOfDayFromMonday = CInt(DateStart.DayOfWeek) - 1
    End Select
    If IsMondaySync = False Then
      NumberTradingDays = NumberTradingDays + ThisNumberOfDayFromMonday
    End If
    If NumberTradingDays < 0 Then
      'synchronize to friday before we apply the transformation 
      'and then put back the reference to monday by adding 4 before return
      ThisDeltaDays = NumberTradingDays - 4
      ThisDeltaDays = 7 * (ThisDeltaDays \ 5) + (ThisDeltaDays Mod 5) - ThisNumberOfDayFromMonday
      ThisDeltaDays = ThisDeltaDays + 4
    Else
      ThisDeltaDays = 7 * (NumberTradingDays \ 5) + (NumberTradingDays Mod 5) - ThisNumberOfDayFromMonday
    End If
    'the return date will never fall on a weekend
    Return DateStart.AddDays(ThisDeltaDays)
  End Function

  ''' <summary>
  ''' This function return the number of market trading days assuming 5 days trading per week and the week starting on monday. Here, DateStart is always 
  ''' synchronized back to the previous Monday. This equation is valid for any range of DateStop positive or negative relative to DateStart
  ''' </summary>
  ''' <param name="DateStart"></param>
  ''' <param name="DateStop"></param>
  ''' <returns>
  '''   A linear daily date scale number assuming 5 days a week and with the week always starting on Monday
  ''' </returns>
  ''' <remarks>This function has been tested using excel</remarks>
  Public Shared Function MarketTradingDeltaDays(ByVal DateStart As Date, ByVal DateStop As Date, Optional ByVal IsMondaySync As Boolean = True) As Integer
    Dim ThisNumberOfDayFromMonday As Integer
    'make sure the date do not fall on a weekend
    'if it does bring it back to last friday
    'remove the time information
    DateStart = DateToWeekEndRemovePrevious(DateStart.Date)
    DateStop = ReportDate.DateToWeekEndRemovePrevious(DateStop.Date)
    'this is equivalent than calling DateToMondayPrevious(DateStart)
    'the transformation is generally non-linear and we procede by condition
    Select Case DateStart.DayOfWeek
      Case DayOfWeek.Monday    '1
        ThisNumberOfDayFromMonday = 0
        'Case DayOfWeek.Sunday   '0
        'do not need to test for this case
        'ThisNumberOfDayFromMonday = 6
      Case Else
        'this portion is linear
        ThisNumberOfDayFromMonday = CInt(DateStart.DayOfWeek) - 1
    End Select
    'make sure the number of delta days include the number of day from Monday
    Dim ThisDeltaDays As Integer = DateStop.Subtract(DateStart).Days + ThisNumberOfDayFromMonday
    'this equation is valid only when the starting day is Monday which is an imposed constraint here
    'compress the linear date scale of ThisDeltaDays to eliminate the weekend
    If DateStop < DateStart Then
      'synchronize to friday before we apply the transformation 
      'and then put back the reference to monday by adding 4 before return
      ThisDeltaDays = ThisDeltaDays - 4
      ThisDeltaDays = (5 * (ThisDeltaDays \ 7)) + (ThisDeltaDays Mod 7)
      ThisDeltaDays = ThisDeltaDays + 4
    Else
      ThisDeltaDays = (5 * (ThisDeltaDays \ 7)) + (ThisDeltaDays Mod 7)
    End If
    If IsMondaySync = False Then
      ThisDeltaDays = ThisDeltaDays - ThisNumberOfDayFromMonday
    End If
    Return ThisDeltaDays
  End Function

  ''' <summary>
  ''' Synchronize a date to the previous Friday
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function DateToPreviousFriday(ByVal DateValue As Date) As Date
    Return DateToNextFriday(DateValue).AddDays(-7)
  End Function

  ''' <summary>
  ''' Synchronize a date to Friday. If the date is already on friday it is left unchanged
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function DateToFriday(ByVal DateValue As Date) As Date
    Select Case DateAndTime.Weekday(DateValue.Date)
      Case FirstDayOfWeek.Saturday
        DateValue = DateValue.AddDays(6)
      Case FirstDayOfWeek.Sunday
        DateValue = DateValue.AddDays(5)
      Case FirstDayOfWeek.Monday
        DateValue = DateValue.AddDays(4)
      Case FirstDayOfWeek.Tuesday
        DateValue = DateValue.AddDays(3)
      Case FirstDayOfWeek.Wednesday
        DateValue = DateValue.AddDays(2)
      Case FirstDayOfWeek.Thursday
        DateValue = DateValue.AddDays(1)
    End Select
    Return DateValue
  End Function

  ''' <summary>
  ''' Synchronize a date to Friday. If the date is on friday it is move to the next friday
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function DateToNextFriday(ByVal DateValue As Date) As Date
    Select Case DateAndTime.Weekday(DateValue.Date)
      Case FirstDayOfWeek.Saturday
        DateValue = DateValue.AddDays(6)
      Case FirstDayOfWeek.Sunday
        DateValue = DateValue.AddDays(5)
      Case FirstDayOfWeek.Monday
        DateValue = DateValue.AddDays(4)
      Case FirstDayOfWeek.Tuesday
        DateValue = DateValue.AddDays(3)
      Case FirstDayOfWeek.Wednesday
        DateValue = DateValue.AddDays(2)
      Case FirstDayOfWeek.Thursday
        DateValue = DateValue.AddDays(1)
      Case FirstDayOfWeek.Friday
        DateValue = DateValue.AddDays(7)
    End Select
    Return DateValue
  End Function

  ''' <summary>
  ''' Return teh week number from a date
  ''' </summary>
  ''' <param name="TargetDate"></param>
  ''' <returns></returns>
  Public Shared Function WeekNumber(ByVal TargetDate As Date) As Integer
    Return (TargetDate.Day - 1) \ 7 + 1
  End Function

  ''' <summary>
  ''' This funtion return true if the data are in the same work week from Monday to Friday included.
  ''' the function return false if any of the data fall on a weekend. The date can be of any order realtive to each other.. 
  ''' </summary>
  ''' <param name="DateValue1"></param>
  ''' <param name="DateValue2"></param>
  ''' <returns></returns>
  Public Shared Function IsSameTradingWeek(ByVal DateValue1 As Date, DateValue2 As Date) As Boolean
    Dim ThisDateRef As Date   'the smallest date is use for reference
    Dim ThisDateTest As Date
    Dim ThisNumberOfDeltaDay As Integer

    If DateValue1 < DateValue2 Then
      ThisDateRef = DateValue1
      ThisDateTest = DateValue2
    Else
      ThisDateRef = DateValue2
      ThisDateTest = DateValue1
    End If
    'ThisDateRef Is the smaller or previous date and time
    ThisNumberOfDeltaDay = (ThisDateTest.Date - ThisDateRef.Date).Days
    If ThisNumberOfDeltaDay > 4 Then Return False
    If (ThisDateRef.DayOfWeek = DayOfWeek.Saturday) Or ((ThisDateRef.DayOfWeek = DayOfWeek.Sunday)) Then Return False
    If (ThisDateTest.DayOfWeek = DayOfWeek.Saturday) Or ((ThisDateTest.DayOfWeek = DayOfWeek.Sunday)) Then Return False
    Select Case ThisDateRef.DayOfWeek
      Case DayOfWeek.Monday
        Return True
      Case DayOfWeek.Tuesday
        If ThisNumberOfDeltaDay > 3 Then Return False
      Case DayOfWeek.Wednesday
        If ThisNumberOfDeltaDay > 2 Then Return False
      Case DayOfWeek.Thursday
        If ThisNumberOfDeltaDay > 1 Then Return False
      Case DayOfWeek.Friday
        If ThisNumberOfDeltaDay > 0 Then Return False
    End Select
    Return True
  End Function

  ''' <summary>
  ''' Synchronize a date to previous Friday if it fall on a WeekEnd
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function DateToWeekEndRemovePrevious(ByVal DateValue As Date) As Date
    Select Case DateAndTime.Weekday(DateValue.Date)
      Case FirstDayOfWeek.Saturday
        DateValue = DateValue.AddDays(-1)
      Case FirstDayOfWeek.Sunday
        DateValue = DateValue.AddDays(-2)
    End Select
    Return DateValue
  End Function

  Public Shared Function DateToWeekEndRemoveNext(ByVal DateValue As Date) As Date
    Select Case DateAndTime.Weekday(DateValue.Date)
      Case FirstDayOfWeek.Saturday
        DateValue = DateValue.AddDays(2)
      Case FirstDayOfWeek.Sunday
        DateValue = DateValue.AddDays(1)
    End Select
    Return DateValue
  End Function

  ''' <summary>
  ''' Return the previous monday date. If the date fall on a Monday it return the current monday 
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks>
  ''' The value of the constants in the DayOfWeek enumeration ranges from DayOfWeek..::..Sunday to 
  ''' DayOfWeek..::..Saturday. If cast to an integer, 
  ''' its value ranges from zero (which indicates DayOfWeek..::..Sunday) to six 
  ''' (which indicates DayOfWeek..::..Saturday).
  ''' </remarks>
  Public Shared Function DateToMondayPrevious(ByVal DateValue As Date) As Date
    'the transformation is generally non-linear
    Select Case DateValue.DayOfWeek
      Case DayOfWeek.Sunday   '0
        Return DateValue.AddDays(-6)
      Case DayOfWeek.Monday    '1
        Return DateValue
      Case Else
        'this portion is linear
        Return DateValue.AddDays(1 - CInt(DateValue.DayOfWeek))
    End Select
  End Function
  ''' <summary>
  ''' Return the next monday date if the imput date is not already a Monday.
  ''' </summary>
  ''' <param name="DateValue"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function DateToMondayNext(ByVal DateValue As Date) As Date
    Select Case DateValue.DayOfWeek
      Case DayOfWeek.Sunday   '0
        Return DateValue.AddDays(1)
      Case DayOfWeek.Monday   '1
        Return DateValue
      Case Else
        Return DateValue.AddDays(8 - CInt(DateValue.DayOfWeek))
    End Select
  End Function

  Public Shared Function DateNullValue() As Date
    Return DateSerial(1900, 1, 1)
  End Function

  Public Shared Function IsPeriodOverlap(
    ByVal PeriodStart As Date,
    ByVal PeriodStop As Date,
    ByVal DateValue As Date) As Boolean

    If (DateValue >= PeriodStart) Then
      If (DateValue <= PeriodStop) Then
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Shared Function IsPeriodOverlap(
    ByVal PeriodStart As Date,
    ByVal PeriodStop As Date,
    ByVal DateStartValue As Date,
    ByVal DateStopValue As Date) As Boolean

    If ReportDate.IsPeriodOverlap(PeriodStart, PeriodStop, DateStartValue) Then
      Return True
    Else
      Return ReportDate.IsPeriodOverlap(PeriodStart, PeriodStop, DateStopValue)
    End If
  End Function

  ' ''' <summary>
  ' ''' Return the beginning of the trading week always the Monday
  ' ''' </summary>
  ' ''' <param name="DateValue"></param>
  ' ''' <returns></returns>
  ' ''' <remarks></remarks>
  'Public Shared Function WeekOfTrade(ByVal DateValue As Date) As Date
  '  Return DateValue.Date.AddDays(1 - DateValue.DayOfWeek)
  'End Function
  Public Shared Function DayOfTradeNext(ByVal DateValue As Date) As Date
    'Static DateValueInLast As Date
    'Static DateValueOutLast As Date
    Dim DateValueInLast As Date
    Dim DateValueOutLast As Date

    'If DateValueInLast = DateValue Then
    '  Return DateValueOutLast
    'End If
    DateValueInLast = DateValue
    DateValueOutLast = DayOfTradeNextLocal(DateValue, True)
    Return DateValueOutLast
  End Function

  Public Shared Function DayOfTradeNextToEndOfDay(ByVal DateValue As Date) As Date
    Return DayOfTradeNext(DateValue).Date.AddSeconds(MARKET_CLOSE_TIME_SEC_DEFAULT)
  End Function

  ''' <summary>
  ''' Calculate the Market time to End of day 
  ''' </summary>
  ''' <param name="DateValue">Current Date and time</param>
  ''' <returns>time left in unit of days before the close of days trading</returns>
  Public Shared Function MarketTimeToEndOfDay(ByVal DateValue As Date) As Double
    Dim ThisDateNextToMarketOpen As Date
    Dim ThisDateNextToMarketClose As Date
    Dim ThisNumberOfTradingDays As Double

    ThisDateNextToMarketOpen = YahooAccessData.ReportDate.DayOfTradeNext(DateValue, MarketDelayMinutes:=0)
    ThisDateNextToMarketClose = YahooAccessData.ReportDate.DayOfTradeNextToEndOfDay(ThisDateNextToMarketOpen)
    ThisNumberOfTradingDays = ThisDateNextToMarketClose.Subtract(DateValue).TotalHours / 24
    Return ThisNumberOfTradingDays
  End Function

  ''' <summary>
  ''' Calculate the Market time to the open of day 
  ''' </summary>
  ''' <param name="DateValue">Current Date and time</param>
  ''' <returns>time left in unit of days before the open of days trading</returns>
  Public Shared Function MarketTimeToOpenOfDay(ByVal DateValue As Date) As Double
    Dim ThisDateNextToMarketOpen As Date
    Dim ThisDateNextToMarketClose As Date
    Dim ThisNumberOfTradingDaysToOpen As Double

    ThisDateNextToMarketOpen = YahooAccessData.ReportDate.DayOfTradeNext(DateValue, MarketDelayMinutes:=0)
    ThisDateNextToMarketClose = YahooAccessData.ReportDate.DayOfTradeNextToEndOfDay(ThisDateNextToMarketOpen)
    ThisNumberOfTradingDaysToOpen = ThisDateNextToMarketOpen.Subtract(DateValue).TotalHours / 24
    If ThisNumberOfTradingDaysToOpen > Measure.TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY Then
      ThisNumberOfTradingDaysToOpen = Measure.TIME_TO_MARKET_PREVIOUS_CLOSE_TO_OPEN_IN_DAY
    End If
    Return ThisNumberOfTradingDaysToOpen
  End Function

  Public Shared Function DayOfTradePrevious(ByVal DateValue As Date) As Date
    Return DayOfTrade(DateValue.AddDays(-1))
  End Function

  Public Shared Function DayOfTradeNext(
    ByVal DateValue As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date
    'execute all the time here
    Return DayOfTradeNextLocal(
      DateValue,
      False,
      MARKET_OPEN_TIME_DEFAULT,
      MARKET_CLOSE_TIME_DEFAULT,
      MarketDelayMinutes)
  End Function

  Public Shared Function DayOfTradeNext(
    ByVal DateValue As Date,
    ByVal MarketOpenTime As Date,
    ByVal MarketCloseTime As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    'execute all the time here
    Return DayOfTradeNextLocal(
      DateValue,
      False,
      MarketOpenTime,
      MarketCloseTime,
      MarketDelayMinutes)
  End Function

  Public Shared Function DayOfTradeToEndOfDay(ByVal DateValue As Date) As Date
    Return DayOfTrade(DateValue).Date.AddSeconds(MARKET_CLOSE_TIME_SEC_DEFAULT)
  End Function

  Public Shared Function DayOfTrade(ByVal DateValue As Date) As Date
    Dim DateValueInLast As Date
    Dim DateValueOutLast As Date

    'If DateValueInLast = DateValue Then
    '  Return DateValueOutLast
    'End If
    DateValueInLast = DateValue
    DateValueOutLast = DayOfTradeLocal(DateValue, IsMarketTimeToDefault:=True)
    Return DateValueOutLast
  End Function

  Public Shared Function DayOfTrade(
    ByVal DateValue As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    Return DayOfTrade(DateValue, MARKET_OPEN_TIME_DEFAULT, MARKET_CLOSE_TIME_DEFAULT, MarketDelayMinutes)
  End Function

  Public Shared Function DayOfTrade(
    ByVal DateValue As Date,
    ByVal MarketOpenTime As Date,
    ByVal MarketCloseTime As Date,
    Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    'execute all the time here
    Return DayOfTradeLocal(
      DateValue,
      False,
      MarketOpenTime,
      MarketCloseTime,
      MarketDelayMinutes)
  End Function
#End Region
#Region "Date computation support function"
  Private Shared Function DayOfTradeNextLocal(
      ByVal DateValue As Date,
      ByVal IsMarketTimeToDefault As Boolean,
      Optional ByVal MarketOpenTime As Date = #9:30:00 AM#,
      Optional ByVal MarketCloseTime As Date = #4:00:00 PM#,
      Optional ByVal MarketDelayMinutes As Integer = 15) As Date

    Dim ThisDateValueIn As Date = DateValue
    Dim ThisMarketOpenCloseTime As MarketOpenCloseTime
    Dim ThisDateTemp As Date

    'add a day until we get to a date that is not a special holiday
    ThisMarketOpenCloseTime = New MarketOpenCloseTime(DateValue, MarketOpenTime, MarketCloseTime, MarketDelayMinutes)
    If ThisMarketOpenCloseTime.IsSpecialHoliday Then
      Return DayOfTradeNextLocal(DateValue.Date.AddDays(1).Add(ThisMarketOpenCloseTime.OpenTimeSpanDelay), IsMarketTimeToDefault)
    End If
    'adjust for normal daily trading close time
    If DateValue.TimeOfDay > ThisMarketOpenCloseTime.CloseTimeSpanDelay Then
      'the trading is for the next day
      DateValue = DateValue.AddDays(1)
    End If
    'remove the time component
    DateValue = DateValue.Date
    'correct for weekend
    Select Case DateValue.DayOfWeek
      Case System.DayOfWeek.Sunday
        DateValue = DateValue.AddDays(1)
      Case System.DayOfWeek.Saturday
        DateValue = DateValue.AddDays(2)
    End Select
    'correct for the holiday
    Select Case DateValue.Month
      Case 1
        'check for new year i.e.
        'New Years' Day (January 1) in 2011 falls on a Saturday. 
        'The rules of the applicable exchanges state that when a holiday falls on a Saturday, 
        'the preceding Friday is observed unless the Friday is the end of a monthly or yearly 
        'accounting period. In this case, Friday, December 31, 2010 is the end of both a monthly 
        'and yearly accounting period; therefore the exchanges will be open that day 
        'and the following Monday.
        'When any stock market holiday falls on a Saturday, 
        'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
        'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
        Select Case DateValue.Day
          Case 1
            'this is January first
            If DateValue.DayOfWeek = DayOfWeek.Friday Then
              DateValue = DateValue.AddDays(3)
            Else
              DateValue = DateValue.AddDays(1)
            End If
          Case 2
            'Jan 2 could be an holiday if the new year did fall on Sunday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              'new year was on Sunday and in this case the following rule apply:
              'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
              DateValue = DateValue.AddDays(1)
            End If
          Case Else
            'check for other Holiday happening on Monday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              'Martin Luther King is observed on the third Monday of January each year
              If DateValue = DateFromDayOfWeek(DateValue.Year, 1, 3, FirstDayOfWeek.Monday) Then
                'this is the Martin Luther King holiday
                DateValue = DateValue.AddDays(1)
              End If
            End If
        End Select
      Case 2
        'Washington's Birthday is a United States federal holiday 
        'celebrated on the third Monday of February in honor of George Washington, the first President of the United States.
        'check for Holliday
        'Martin Luther King is observed on the third Monday of January each year
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          If DateValue = DateFromDayOfWeek(DateValue.Year, 2, 3, FirstDayOfWeek.Monday) Then
            'this is Washington's Birthday holiday
            'the trading is for the next day
            DateValue = DateValue.AddDays(1)
          End If
        End If
      Case 3, 4
        'check for good Friday
        If DateValue.DayOfWeek = DayOfWeek.Friday Then
          'check if this is good Friday
          'the market is not open on good Friday in the US and Canada
          'EasterDate is Sunday
          If DateValue = EasterDate(DateValue).AddDays(-2) Then
            'this is good Friday
            'next Monday is not an holiday in the US but it is in Canada
            'for now assume the market is the US and the market is trading on Monday
            DateValue = DateValue.AddDays(3)
          End If
        End If
      Case 5
        'Memorial Day is a United States federal holiday observed on the last Monday of May.
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          'check if this is good Friday
          'the market is not open on good Friday in the US and Canada
          'EasterDate is Sunday
          ThisDateTemp = DateValue.AddDays(7)
          If ThisDateTemp.Month = 6 Then
            'this is the Memorial Day holiday
            'move to the next day
            DateValue = DateValue.AddDays(1)
          End If

          'the code below does not work for 24/05/2021
          'it was replaced by the above one on 24/05/2021
          'If DateValue = DateFromDayOfWeek(
          '  ThisYear:=DateValue.Year,
          '  ThisMonth:=5,
          '  NumberOfFirstDayOfWeek:=-1,
          '  ThisFirstDayOfWeek:=FirstDayOfWeek.Monday) Then

          '  'this is the Memorial Day holiday
          '  DateValue = DateValue.AddDays(1)
          'End If
        End If
      Case 7
        'check for Independence day always on July 4
        'When any stock market holiday falls on a Saturday, 
        'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
        'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
        If DateValue.Day = 4 Then
          If DateValue.DayOfWeek = DayOfWeek.Friday Then
            DateValue = DateValue.AddDays(3)
          Else
            DateValue = DateValue.AddDays(1)
          End If
        Else
          If DateValue.Day = 3 Then
            If DateValue.DayOfWeek = DayOfWeek.Friday Then
              'Friday July 3 is also and holiday for Independence day on Saturday
              DateValue = DateValue.AddDays(3)
            End If
          ElseIf DateValue.Day = 5 Then
            'Monday July 5 is also and holiday for Independence day on Sunday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              DateValue = DateValue.AddDays(1)
            End If
          End If
        End If
      Case 9
        'Labor Day:
        'The first Monday in September, observed as a holiday in the United States and Canada
        'in honor of working people.
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          If DateValue = DateFromDayOfWeek(DateValue.Year, 9, 1, FirstDayOfWeek.Monday) Then
            'this is labor day
            DateValue = DateValue.AddDays(1)
          End If
        End If
      Case 11
        'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
        'Traditionally, it has been a time to give thanks to God, friends, and family.
        'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
        'and in the United States, it is celebrated on the fourth Thursday of November. 
        'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
        If DateValue.DayOfWeek = DayOfWeek.Thursday Then
          If DateValue = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday) Then
            'this is the Thanksgiving Day holiday
            DateValue = DateValue.AddDays(1)
          End If
        End If
      Case 12
        'Christmas
        Select Case DateValue.Day
          Case 25
            'this is Christmas
            If DateValue.DayOfWeek = DayOfWeek.Friday Then
              DateValue = DateValue.AddDays(3)
            Else
              DateValue = DateValue.AddDays(1)
            End If
          Case 26
            'Dec 26 could be an holiday if Christmas did fall on Sunday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              'Christmas was on Sunday and in this case the following rule apply:
              'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
              DateValue = DateValue.AddDays(1)
            End If
        End Select
    End Select
    If DateValue <> ThisMarketOpenCloseTime.DateValue Then
      'this is a new day and we need the market time open and close
      'add a day until we get to a date that is not a special holiday
      ThisMarketOpenCloseTime = New MarketOpenCloseTime(DateValue, MarketOpenTime, MarketCloseTime, MarketDelayMinutes)
      If ThisMarketOpenCloseTime.IsSpecialHoliday Then
        Return DayOfTradeNextLocal(DateValue.Date.AddDays(1).Add(ThisMarketOpenCloseTime.OpenTimeSpanDelay), IsMarketTimeToDefault)
      End If
    End If
    If DateValue.Date > ThisDateValueIn.Date Then
      DateValue = DateValue.Add(ThisMarketOpenCloseTime.OpenTimeSpanDelay)
    Else
      If ThisDateValueIn.TimeOfDay < ThisMarketOpenCloseTime.OpenTimeSpanDelay Then
        DateValue = DateValue.Add(ThisMarketOpenCloseTime.OpenTimeSpanDelay)
      Else
        DateValue = ThisDateValueIn
      End If
    End If
    Return DateValue
  End Function


  Private Shared Function DayOfTradeLocal(
    ByVal DateValue As Date,
    ByVal IsMarketTimeToDefault As Boolean,
    Optional ByVal MarketOpenTime As Date = #9:30:00 AM#,
    Optional ByVal MarketCloseTime As Date = #4:00:00 PM#,
    Optional ByVal MarketDelayMinutes As Integer = 15,
    Optional ByVal MarketHolidaySchedule As MarketHolidaySchedule = MarketHolidaySchedule.US) As Date

    Dim ThisDateValueIn As Date = DateValue
    Dim ThisMarketOpenCloseTime As MarketOpenCloseTime

    'substract a day until we get to a date that is not a special holiday
    ThisMarketOpenCloseTime = New MarketOpenCloseTime(DateValue, MarketOpenTime, MarketCloseTime, MarketDelayMinutes)
    If ThisMarketOpenCloseTime.IsSpecialHoliday Then
      Return DayOfTradeLocal(DateValue.Date.AddDays(-1).Add(ThisMarketOpenCloseTime.CloseTimeSpanDelay), IsMarketTimeToDefault)
    End If
    'adjust for normal daily trading time
    'If DateValue.TimeOfDay < MarketOpenTime.TimeOfDay.Add(New TimeSpan(hours:=0, minutes:=MarketDelayMinutes, seconds:=0)) Then
    If DateValue.TimeOfDay < ThisMarketOpenCloseTime.OpenTimeSpanDelay Then
      'the trading is for the previous day
      DateValue = DateValue.AddDays(-1)
    End If
    'remove the time component
    DateValue = DateValue.Date
    'correct for weekend
    Select Case DateValue.DayOfWeek
      Case System.DayOfWeek.Sunday
        DateValue = DateValue.AddDays(-2)
      Case System.DayOfWeek.Saturday
        DateValue = DateValue.AddDays(-1)
    End Select
    'correct for the holiday
    Select Case DateValue.Month
      Case 1
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US, MarketHolidaySchedule.CDN
            'check for new year i.e.
            'New Years' Day (January 1) in 2011 falls on a Saturday. 
            'The rules of the applicable exchanges state that when a holiday falls on a Saturday, 
            'the preceding Friday is observed unless the Friday is the end of a monthly or yearly 
            'accounting period. In this case, Friday, December 31, 2010 is the end of both a monthly 
            'and yearly accounting period; therefore the exchanges will be open that day 
            'and the following Monday.
            'The NYSE, NYSE AMEX and NASDAQ will close trading early (at 1:00 PM ET) on 
            'Friday, November 25, 2011 (the day after Thanksgiving). 

            'Although the day after Thanksgiving (Friday) is not an official holiday, 
            'the market has a tradition of closing at 1:00 p.m. ET. 
            'When any stock market holiday falls on a Saturday, 
            'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
            'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
            Select Case DateValue.Day
              Case 1
                'this is January first
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  DateValue = DateValue.AddDays(-3)
                Else
                  DateValue = DateValue.AddDays(-1)
                End If
              Case 2
                'Jan 2 could be an holiday if the new year did fall on Sunday
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  'new year was on Sunday and in this case the following rule apply:
                  'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
                  DateValue = DateValue.AddDays(-3)
                End If
              Case Else
                'check for other Holiday happening on Monday in January
                Select Case MarketHolidaySchedule
                  Case MarketHolidaySchedule.US
                    If DateValue.DayOfWeek = DayOfWeek.Monday Then
                      'Martin Luther King is observed on the third Monday of January each year
                      If DateValue = DateFromDayOfWeek(DateValue.Year, 1, 3, FirstDayOfWeek.Monday) Then
                        'this is the Martin Luther King holiday
                        DateValue = DateValue.AddDays(-3)
                      End If
                    End If
                  Case MarketHolidaySchedule.CDN
                    'no other holiday for the canadien market in january
                End Select
            End Select
        End Select
      Case 2
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US
            'Washington's Birthday is a United States federal holiday 
            'celebrated on the third Monday of February in honor of George Washington, the first President of the United States.
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              If DateValue = DateFromDayOfWeek(DateValue.Year, 2, 3, FirstDayOfWeek.Monday) Then
                'this is the Washington's Birthday holiday
                DateValue = DateValue.AddDays(-3)
              End If
            End If
          Case MarketHolidaySchedule.CDN
            'In most provinces of Canada, the third Monday in February is observed as a regional
            'statutory holiday, typically known in general as Family Day (French: Jour de la famille)
            'Both the Toronto Stock Exchange and the TSX Venture Exchange are close for the family day
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              If DateValue = DateFromDayOfWeek(DateValue.Year, 2, 3, FirstDayOfWeek.Monday) Then
                'this is the Washington's Birthday holiday
                DateValue = DateValue.AddDays(-3)
              End If
            End If
        End Select
      Case 3, 4
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US, MarketHolidaySchedule.CDN
            'check for good Friday
            If DateValue.DayOfWeek = DayOfWeek.Friday Then
              If DateValue = EasterDate(DateValue).AddDays(-2) Then
                'this is good Friday
                'next Monday is not an holiday in the US but it is in Canada
                DateValue = DateValue.AddDays(-1)
              End If
            End If
        End Select
      Case 5
        'Memorial Day is a United States federal holiday observed on the last Monday of May.
        'only in the US
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              Dim ThisDateTemp = DateValue.AddDays(7)
              If ThisDateTemp.Month = 6 Then
                'this is the Memorial Day holiday
                'move to the next day
                DateValue = DateValue.AddDays(-3)
              End If
            End If
        End Select
      Case 7
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US
            'complicated rule here:
            'check for Independence day always on July 4
            'When any stock market holiday falls on a Saturday, 
            'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
            'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
            If DateValue.Day = 4 Then
              If DateValue.DayOfWeek = DayOfWeek.Monday Then
                DateValue = DateValue.AddDays(-3)
              Else
                DateValue = DateValue.AddDays(-1)
              End If
            Else
              'Friday July 3 is also an holiday for Independence day on Saturday
              If DateValue.Day = 3 Then
                If DateValue.DayOfWeek = DayOfWeek.Friday Then
                  DateValue = DateValue.AddDays(-1)
                End If
              ElseIf DateValue.Day = 5 Then
                'Monday July 5 is also and holiday for Independence day on Sunday
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  DateValue = DateValue.AddDays(-3)
                End If
              End If
            End If
          Case MarketHolidaySchedule.CDN
            'applied the same complicated rule here than in the US for July 4:
            'check for canada day always on July 1
            'When any stock market holiday falls on a Saturday, 
            'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
            'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
            If DateValue.Day = 1 Then
              If DateValue.DayOfWeek = DayOfWeek.Monday Then
                DateValue = DateValue.AddDays(-3)
              Else
                DateValue = DateValue.AddDays(-1)
              End If
            Else
              'July 2 is also an holiday for canada day if Canada day fall on a sunday
              If DateValue.Day = 2 Then
                If DateValue.DayOfWeek = DayOfWeek.Sunday Then
                  DateValue = DateValue.AddDays(-3)
                End If
              ElseIf DateValue.Day = 3 Then
                'July 3 is also an holiday for canada day if Canada day fall on a saturday
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  DateValue = DateValue.AddDays(-3)
                End If
              End If
            End If
        End Select
      Case 8
        'The Civic Holiday is celebrated on the first Monday of August 
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.CDN
            If DateValue = DateFromDayOfWeek(DateValue.Year, 8, 1, FirstDayOfWeek.Monday) Then
              'subtract 3 days to get to the previous Friday
              DateValue = DateValue.AddDays(-3)
            End If
        End Select
      Case 9
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US, MarketHolidaySchedule.CDN
            'Labor Day:
            'The first Monday in September, observed as a holiday in the United States and Canada
            'in honor of working people.
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              If DateValue = DateFromDayOfWeek(DateValue.Year, 9, 1, FirstDayOfWeek.Monday) Then
                'this is labor day
                'subtract 3 days to get to the previous Friday
                DateValue = DateValue.AddDays(-3)
              End If
            End If
        End Select
      Case 10
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.CDN
            'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
            'Traditionally, it has been a time to give thanks to God, friends, and family.
            'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
            'and in the United States, it is celebrated on the fourth Thursday of November. 
            'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
            Dim ThisDateForThanksGiving = DateFromDayOfWeek(DateValue.Year, 11, 2, FirstDayOfWeek.Monday)
            If DateValue.DayOfWeek = DayOfWeek.Thursday Then
              If DateValue = ThisDateForThanksGiving Then
                'this is the Thanksgiving Day holiday
                'subtract 3 days to get to the previous day
                DateValue = DateValue.AddDays(-3)
              End If
            End If
        End Select
      Case 11
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US
            'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
            'Traditionally, it has been a time to give thanks to God, friends, and family.
            'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
            'and in the United States, it is celebrated on the fourth Thursday of November. 
            'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
            Dim ThisDateForThanksGiving = DateFromDayOfWeek(DateValue.Year, 11, 4, FirstDayOfWeek.Thursday)
            If DateValue.DayOfWeek = DayOfWeek.Thursday Then
              If DateValue = ThisDateForThanksGiving Then
                'this is the Thanksgiving Day holiday
                'subtract 1 days to get to the previous day
                DateValue = DateValue.AddDays(-1)
              End If
            End If
        End Select
      Case 12
        Select Case MarketHolidaySchedule
          Case MarketHolidaySchedule.US, MarketHolidaySchedule.CDN
            'Christmas
            Select Case DateValue.Day
              Case 25
                'this is Christmas
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  DateValue = DateValue.AddDays(-3)
                Else
                  DateValue = DateValue.AddDays(-1)
                End If
              Case 26
                'Dec 26 could be an holiday if Christmas did fall on Sunday
                If DateValue.DayOfWeek = DayOfWeek.Monday Then
                  'Christmas was on Sunday and in this case the following rule apply:
                  'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
                  DateValue = DateValue.AddDays(-3)
                End If
            End Select
        End Select
    End Select
    If DateValue <> ThisMarketOpenCloseTime.DateValue Then
      'this is a new day and we need the market time open and close
      'substract a day until we get to a date that is not a special holiday
      ThisMarketOpenCloseTime = New MarketOpenCloseTime(DateValue, MarketOpenTime, MarketCloseTime, MarketDelayMinutes)
      If ThisMarketOpenCloseTime.IsSpecialHoliday Then
        Return DayOfTradeLocal(DateValue.Date.AddDays(-1).Add(ThisMarketOpenCloseTime.CloseTimeSpanDelay), IsMarketTimeToDefault)
      End If
    End If
    'note that DateValue does not contain the time component
    If DateValue.Date < ThisDateValueIn.Date Then
      DateValue = DateValue.Add(ThisMarketOpenCloseTime.CloseTimeSpanDelay)
    Else
      If ThisDateValueIn.TimeOfDay > ThisMarketOpenCloseTime.CloseTimeSpanDelay Then
        'DateValue does not contain the time value
        DateValue = DateValue.Add(ThisMarketOpenCloseTime.CloseTimeSpanDelay)
      Else
        DateValue = ThisDateValueIn
      End If
    End If
    Return DateValue
  End Function

  Public Shared Function DateFromDayOfWeek(
    ByVal ThisYear As Integer,
    ByVal ThisMonth As Integer,
    ByVal NumberOfFirstDayOfWeek As Integer,
    ByVal ThisFirstDayOfWeek As Microsoft.VisualBasic.FirstDayOfWeek) As Date

    Dim ThisNumberDeltaDay As Integer
    Dim ThisDate As Date

    'note: The value returned by the Weekday function corresponds to the values 
    'of the FirstDayOfWeek enumeration; that is, 1 indicates Sunday and 7 indicates Saturday
    Select Case NumberOfFirstDayOfWeek
      Case 0
        ThisDate = DateSerial(ThisYear, ThisMonth, 1)
      Case Is > 0
        Dim ThisFirstDateOfMonth As Date
        ThisFirstDateOfMonth = DateSerial(ThisYear, ThisMonth, 1)
        'Number of day to add before we reach the ThisFirstDayOfWeek for the given the year and month
        ThisNumberDeltaDay = (7 - ((Weekday(ThisFirstDateOfMonth, ThisFirstDayOfWeek) - 1))) Mod 7
        ThisDate = ThisFirstDateOfMonth.AddDays(ThisNumberDeltaDay).AddDays(7 * (NumberOfFirstDayOfWeek - 1))
      Case Is < 0
        Throw New NotImplementedException("Function 'DateFromDayOfWeek' is not supported with negative weeks")
        'there is some issue here
        'measure from the last day of month
        'Dim ThisLastDateOfMonth As Date
        'ThisLastDateOfMonth = DateSerial(ThisYear, ThisMonth, 1).AddMonths(1)
        'ThisLastDateOfMonth = DateFromDayOfWeek(ThisLastDateOfMonth.Year, ThisLastDateOfMonth.Month, 1,)
        'negatif number do not work
        Debugger.Break()
        Dim ThisLastDateOfMonth As Date
        ThisLastDateOfMonth = DateSerial(ThisYear, ThisMonth, 1).AddMonths(1).AddDays(-1)
        'Number of day to add before we reach the ThisFirstDayOfWeek for the given the next month
        ThisNumberDeltaDay = (7 - ((Weekday(ThisLastDateOfMonth, ThisFirstDayOfWeek) - 1))) Mod 7
        ThisDate = ThisLastDateOfMonth.AddDays(ThisNumberDeltaDay - 7).AddDays(7 * (-NumberOfFirstDayOfWeek - 1))
    End Select
    Return ThisDate
  End Function


  ''' <summary>
  ''' For example to find the day for 2nd Friday, February, 2016
  ''' call FindDay(2016, 2, DayOfWeek.Friday, 2)
  ''' </summary>
  ''' <param name="year"></param>
  ''' <param name="month"></param>
  ''' <param name="Day"></param>
  ''' <param name="occurance"></param>
  ''' <returns></returns>
  Public Shared Function FindDay(ByVal year As Integer, ByVal month As Integer, ByVal Day As DayOfWeek, ByVal occurance As Integer) As Integer
    If ((occurance <= 0) OrElse (occurance > 5)) Then
      Throw New Exception("Occurance is invalid")
    End If

    'Note:
    'The DayOfWeek enumeration represents the day of the week in calendars that have seven days per week. 
    'The value Of the constants In this enumeration ranges from Sunday To Saturday. 
    'If cast Then To an Integer, its value ranges from zero (which indicates Sunday) To six (which indicates Saturday).

    Dim firstDayOfMonth As DateTime = New DateTime(year, month, 1)
    'Substract first day of the month with the required day of the week 
    Dim daysneeded = (CType(Day, Integer) - CType(firstDayOfMonth.DayOfWeek, Integer))
    'if it is less than zero we need to get the next week day (add 7 days)
    If (daysneeded < 0) Then
      daysneeded = (daysneeded + 7)
    End If

    'DayOfWeek is zero index based; multiply by the Occurance to get the day
    Dim resultedDay = ((daysneeded + 1) + (7 _
                    * (occurance - 1)))
    If (resultedDay > (firstDayOfMonth.AddMonths(1) - firstDayOfMonth).Days) Then
      Throw New Exception(String.Format("No {0} occurance(s) of {1} in the required month", occurance, Day.ToString))
    End If

    Return resultedDay
  End Function

  ' ===================================================================
  ' Easter Calculator 0.1
  ' -------------------------------------------------------------------
  'Easter Date Calculator in Visual Basic.NET
  'This class calculates the date of Easter using an algorithm that 
  'was first published in Butcher's Ecclesiastical Calendar in 1876.  
  'It is valid for all years in the Gregorian Calendar (1583+).  
  'The algorithm used in this class was adapted from pseudo-code 
  'in Practical Astronomy for the Calculator, 3rd Edition by 
  'Peter Duffett-Smith. 
  ' Copyright (c) 2007 David Pinch.
  ' Download from http://www.thoughtproject.com/Snippets/Easter/
  ' ===================================================================
  Public Shared Function EasterDate(ByVal DateValue As Date) As Date
    ' ===============================================================
    ' Easter
    ' ---------------------------------------------------------------
    ' Calculates the date of Easter using an algorithm that was first
    ' published in Butcher's Ecclesiastical Calendar (1876).  It is
    ' valid for all years in the Gregorian calendar (1583+).  The
    ' code is based on an implementation by Peter Duffett-Smith in
    ' Practical Astronomy with your Calculator (3rd Edition).
    ' ===============================================================

    Dim Year As Integer
    Static EasterDateLast As Date
    Static YearLast As Integer

    Year = DateValue.Year
    If Year = YearLast Then
      Return EasterDateLast
    End If
    YearLast = Year
    If Year < 1583 Then
      'this calculation is not valid but e do not want to retunr an error
      'Err.Raise(5)
      EasterDateLast = DateSerial(Year:=1583, Month:=3, Day:=22)
    Else
      Dim a As Integer
      Dim b As Integer
      Dim c As Integer
      Dim d As Integer
      Dim e As Integer
      Dim f As Integer
      Dim g As Integer
      Dim h As Integer
      Dim i As Integer
      Dim k As Integer
      Dim l As Integer
      Dim m As Integer
      Dim n As Integer
      Dim p As Integer
      ' Step 1: Divide the year by 19 and store the
      ' remainder in variable A.  Example: If the year
      ' is 2000, then A is initialized to 5.

      a = Year Mod 19

      ' Step 2: Divide the year by 100.  Store the integer
      ' result in B and the remainder in C.

      b = Year \ 100
      c = Year Mod 100

      ' Step 3: Divide B (calculated above).  Store the
      ' integer result in D and the remainder in E.

      d = b \ 4
      e = b Mod 4

      ' Step 4: Divide (b+8)/25 and store the integer
      ' portion of the result in F.

      f = (b + 8) \ 25

      ' Step 5: Divide (b-f+1)/3 and store the integer
      ' portion of the result in G.

      g = (b - f + 1) \ 3

      ' Step 6: Divide (19a+b-d-g+15)/30 and store the
      ' remainder of the result in H.

      h = (19 * a + b - d - g + 15) Mod 30

      ' Step 7: Divide C by 4.  Store the integer result
      ' in I and the remainder in K.

      i = c \ 4
      k = c Mod 4

      ' Step 8: Divide (32+2e+2i-h-k) by 7.  Store the
      ' remainder of the result in L.

      l = (32 + 2 * e + 2 * i - h - k) Mod 7

      ' Step 9: Divide (a + 11h + 22l) by 451 and
      ' store the integer portion of the result in M.

      m = (a + 11 * h + 22 * l) \ 451

      ' Step 10: Divide (h + l - 7m + 114) by 31.  Store
      ' the integer portion of the result in N and the
      ' remainder in P.

      n = (h + l - 7 * m + 114) \ 31
      p = (h + l - 7 * m + 114) Mod 31

      ' At this point p+1 is the day on which Easter falls.
      ' n is 3 for March or 4 for April.

      EasterDateLast = DateSerial(Year, n, p + 1)
    End If
    Return EasterDateLast
  End Function
#End Region      'Date computation support function
End Class
#End Region
