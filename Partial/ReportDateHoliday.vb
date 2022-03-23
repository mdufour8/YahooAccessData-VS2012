Public Class MarketHolidayUS
  Implements IHolidayType(Of EnumMarketHolidayUS)

  Public Enum EnumMarketHolidayUS
    None
    NewYear
    MartinLutherKing
    PresidentDay
    GoodFriday
    MemorialDay
    IndependenceDay
    LabourDay
    ThanksGivingDay
    Christmas
  End Enum


  Private MyDateValue As Date
  Private MyDateOfPreviousTradingDay As Date
  Private MyDateOfNextTradingDay As Date
  Private MyMarketHoliday As EnumMarketHolidayUS
  Private IsWeekEndDayLocal As Boolean


  Public Sub New(ByVal DateValue As Date)
    MyDateValue = DateValue
    MyDateOfPreviousTradingDay = MyDateValue.Date
    MyDateOfNextTradingDay = MyDateValue.Date
    MyMarketHoliday = EnumMarketHolidayUS.None
    IsWeekEndDayLocal = False
    'correct for weekend
    Select Case MyDateOfPreviousTradingDay.DayOfWeek
      Case System.DayOfWeek.Sunday
        IsWeekEndDayLocal = True
        MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-2)
      Case System.DayOfWeek.Saturday
        IsWeekEndDayLocal = True
        MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
    End Select
    Select Case MyDateOfPreviousTradingDay.Month
      Case 1
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
        Select Case MyDateOfPreviousTradingDay.Day
          Case 1
            'this is January first
            MyMarketHoliday = EnumMarketHolidayUS.NewYear
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
            Else
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
            End If
          Case 2
            'Jan 2 could be an holiday if the new year did fall on Sunday
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              'new year was on Sunday and in this case the following rule apply:
              'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
              MyMarketHoliday = EnumMarketHolidayUS.NewYear
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
            End If
          Case Else
            'check for other Holiday happening on Monday in January
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              'Martin Luther King is observed on the third Monday of January each year
              If MyDateOfPreviousTradingDay = ReportDate.DateFromDayOfWeek(MyDateOfPreviousTradingDay.Year, 1, 3, FirstDayOfWeek.Monday) Then
                'this is the Martin Luther King holiday
                MyMarketHoliday = EnumMarketHolidayUS.MartinLutherKing
                MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
              End If
            End If
        End Select
      Case 2
        'Washington's Birthday is a United States federal holiday 
        'celebrated on the third Monday of February in honor of George Washington, the first President of the United States.
        If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
          If MyDateOfPreviousTradingDay = ReportDate.DateFromDayOfWeek(MyDateOfPreviousTradingDay.Year, 2, 3, FirstDayOfWeek.Monday) Then
            'this is the Washington's Birthday holiday
            MyMarketHoliday = EnumMarketHolidayUS.PresidentDay
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
          End If
        End If
      Case 3, 4
        'check for good Friday
        If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Friday Then
          If MyDateOfPreviousTradingDay = ReportDate.EasterDate(MyDateOfPreviousTradingDay).AddDays(-2) Then
            'this is good Friday holiday
            MyMarketHoliday = EnumMarketHolidayUS.GoodFriday
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
          End If
        End If
      Case 5
        If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
          Dim ThisDateTemp = MyDateOfPreviousTradingDay.AddDays(7)
          If ThisDateTemp.Month = 6 Then
            'this is the Memorial Day holiday
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
          End If
        End If
      Case 7
        'complicated rule here:
        'check for Independence day always on July 4
        'When any stock market holiday falls on a Saturday, 
        'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
        'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
        If MyDateOfPreviousTradingDay.Day = 4 Then
          MyMarketHoliday = EnumMarketHolidayUS.IndependenceDay
          If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
          Else
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
          End If
        Else
          'Friday July 3 is also an holiday for Independence day on Saturday
          If MyDateOfPreviousTradingDay.Day = 3 Then
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Friday Then
              MyMarketHoliday = EnumMarketHolidayUS.IndependenceDay
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
            End If
          ElseIf MyDateOfPreviousTradingDay.Day = 5 Then
            'Monday July 5 is also and holiday for Independence day on Sunday
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              MyMarketHoliday = EnumMarketHolidayUS.IndependenceDay
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
            End If
          End If
        End If
      Case 9
        'Labor Day:
        'The first Monday in September, observed as a holiday in the United States and Canada
        'in honor of working people.
        If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
          If MyDateOfPreviousTradingDay = ReportDate.DateFromDayOfWeek(MyDateOfPreviousTradingDay.Year, 9, 1, FirstDayOfWeek.Monday) Then
            'this is labor day
            'subtract 3 days to get to the previous Friday
            MyMarketHoliday = EnumMarketHolidayUS.LabourDay
            MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
          End If
        End If
      Case 11
        'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
        'Traditionally, it has been a time to give thanks to God, friends, and family.
        'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
        'and in the United States, it is celebrated on the fourth Thursday of November. 
        'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
        Dim ThisDateForThanksGiving = ReportDate.DateFromDayOfWeek(MyDateOfPreviousTradingDay.Year, 11, 4, FirstDayOfWeek.Thursday)
        If MyDateOfPreviousTradingDay = ThisDateForThanksGiving Then
          'this is the Thanksgiving day holiday
          'subtract 1 days to get to the previous day
          MyMarketHoliday = EnumMarketHolidayUS.ThanksGivingDay
          MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
        End If
      Case 12
        'Christmas
        Select Case MyDateOfPreviousTradingDay.Day
          Case 25
            'this is Christmas
            MyMarketHoliday = EnumMarketHolidayUS.Christmas
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
            Else
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
            End If
          Case 26
            'Dec 26 could be an holiday if Christmas did fall on Sunday
            If MyDateOfPreviousTradingDay.DayOfWeek = DayOfWeek.Monday Then
              'Christmas was on Sunday and in this case the following rule apply:
              'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
              MyMarketHoliday = EnumMarketHolidayUS.Christmas
              MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-3)
            End If
        End Select
    End Select
  End Sub

  Public ReadOnly Property DateValue As Date Implements IHolidayType(Of EnumMarketHolidayUS).DateValue
    Get
      Return MyDateValue
    End Get
  End Property

  Public ReadOnly Property IsHoliday As Boolean Implements IHolidayType(Of EnumMarketHolidayUS).IsHoliday
    Get
      If MyMarketHoliday = EnumMarketHolidayUS.None Then
        Return False
      Else
        Return True
      End If
    End Get
  End Property

  Public ReadOnly Property Holiday As EnumMarketHolidayUS Implements IHolidayType(Of EnumMarketHolidayUS).Holiday
    Get
      Return MyMarketHoliday
    End Get
  End Property

  Public ReadOnly Property DateOfLastTradingDay As Date Implements IHolidayType(Of EnumMarketHolidayUS).DateOfLastTradingDay
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Public ReadOnly Property IsWeekEndDay As Boolean Implements IHolidayType(Of EnumMarketHolidayUS).IsWeekEndDay
    Get
      Throw New NotImplementedException()
    End Get
  End Property

  Public ReadOnly Property DateOfNexTradingDay As Date Implements IHolidayType(Of EnumMarketHolidayUS).DateOfNexTradingDay
    Get
      Throw New NotImplementedException()
    End Get
  End Property
End Class

Public Class MarketHolidayCDN
  Implements IHolidayType(Of EnumMarketHolidayCDN)

  Public Enum EnumMarketHolidayCDN
    None
    NewYear
    FamilyDay
    GoodFriday
    VictoriaDay
    CanadaDay
    CivicHoliday
    LabourDay
    ThanksGivingDay
    Christmas
    BoxingDay
  End Enum

  Private MyDateValue As Date
  Private MyDateOfPreviousTradingDay As Date
  Private MyDateOfNextTradingDay As Date
  Private MyMarketHoliday As EnumMarketHolidayCDN
  Private IsWeekEndDayLocal As Boolean

  Public Sub New(ByVal DateValue As Date)
    MyDateValue = DateValue
    MyDateOfPreviousTradingDay = MyDateValue.Date
    MyDateOfNextTradingDay = MyDateValue.Date
    MyMarketHoliday = EnumMarketHolidayCDN.None
    IsWeekEndDayLocal = False    'correct for weekend
    Select Case MyDateOfPreviousTradingDay.DayOfWeek
      Case System.DayOfWeek.Sunday
        IsWeekEndDayLocal = True
        MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-2)
        MyDateOfNextTradingDay = MyDateOfNextTradingDay.AddDays(1)
      Case System.DayOfWeek.Saturday
        IsWeekEndDayLocal = True
        MyDateOfPreviousTradingDay = MyDateOfPreviousTradingDay.AddDays(-1)
        MyDateOfNextTradingDay = MyDateOfPreviousTradingDay.AddDays(1)
    End Select
    'for a weekend there is two date to process by callback
    If IsWeekEndDayLocal Then
      Dim ThisMarketHolidayForPreviousTradingDay = New MarketHolidayCDN(MyDateOfPreviousTradingDay)
      Dim ThisMarketHolidayForNextTradingDay = New MarketHolidayCDN(MyDateOfNextTradingDay)
      'both should not be in an holiday
      'MyDateOfPreviousTradingDay = ThisMarketHolidayForPreviousTradingDay
      'Private MyDateOfNextTradingDay As Date
      'Private MyMarketHoliday As EnumMarketHolidayCDN


    End If


    Dim ThisDateofPreviousTrading = DateValue
    Dim ThisDateofNextTrading = DateValue
    Select Case DateValue.Month
      Case 1
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
            MyMarketHoliday = EnumMarketHolidayCDN.NewYear
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
              MyMarketHoliday = EnumMarketHolidayCDN.NewYear
              DateValue = DateValue.AddDays(-3)
            End If
        End Select
      Case 2
        'In most provinces of Canada, the third Monday in February is observed as a regional
        'statutory holiday, typically known in general as Family Day (French: Jour de la famille)
        'Both the Toronto Stock Exchange and the TSX Venture Exchange are close for the family day
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          If DateValue = ReportDate.DateFromDayOfWeek(DateValue.Year, ThisMonth:=2, NumberOfFirstDayOfWeek:=3, FirstDayOfWeek.Monday) Then
            'this is the Family day holiday
            MyMarketHoliday = EnumMarketHolidayCDN.FamilyDay
            DateValue = DateValue.AddDays(-3)
          End If
        End If
      Case 3, 4
        'check for good Friday
        If DateValue.DayOfWeek = DayOfWeek.Friday Then
          If DateValue = ReportDate.EasterDate(DateValue).AddDays(-2) Then
            'this is good Friday
            'next Monday is not an holiday in the US but it is in Canada
            DateValue = DateValue.AddDays(-1)
          End If
        End If
      Case 5
        'Victoria Day(French: Fête de la Reine, lit. 'Celebration of the Queen') is a federal Canadian
        'public holiday celebrated on the last Monday preceding May 25.
        'Initially in honour of Queen Victoria's birthday, it has since been
        'celebrated as the official birthday of Canada's sovereign.
        'It is informally considered to be the beginning of the summer season in Canada.
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          Dim ThisDateMaxForVictoriaDay = New DateTime(year:=DateValue.Year, month:=5, day:=25)
          Dim ThisDateForVictoriaDay = ReportDate.DateFromDayOfWeek(DateValue.Year, ThisMonth:=5, NumberOfFirstDayOfWeek:=5, FirstDayOfWeek.Monday)
          Do
            ThisDateForVictoriaDay = ThisDateForVictoriaDay.AddDays(-7)
          Loop Until ThisDateForVictoriaDay < ThisDateMaxForVictoriaDay
          If DateValue = ThisDateForVictoriaDay Then
            'this is the victoria day holiday
            MyMarketHoliday = EnumMarketHolidayCDN.VictoriaDay
            DateValue = DateValue.AddDays(-3)
          End If
        End If
      Case 7
        'applied the same complicated rule here than in the US for July 4:
        'check for canada day always on July 1
        'When any stock market holiday falls on a Saturday, 
        'the market will be closed on the previous day (Friday) unless the Friday is the end of a monthly or yearly accounting period. 
        'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
        If DateValue.Day = 1 Then
          MyMarketHoliday = EnumMarketHolidayCDN.CanadaDay
          If DateValue.DayOfWeek = DayOfWeek.Monday Then
            DateValue = DateValue.AddDays(-3)
          Else
            DateValue = DateValue.AddDays(-1)
          End If
        Else
          'July 2 is also an holiday for canada day if Canada day fall on a sunday
          If DateValue.Day = 2 Then
            If DateValue.DayOfWeek = DayOfWeek.Sunday Then
              MyMarketHoliday = EnumMarketHolidayCDN.CanadaDay
              DateValue = DateValue.AddDays(-3)
            End If
          ElseIf DateValue.Day = 3 Then
            'July 3 is also an holiday for canada day if Canada day fall on a saturday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              MyMarketHoliday = EnumMarketHolidayCDN.CanadaDay
              DateValue = DateValue.AddDays(-3)
            End If
          End If
        End If
      Case 8
        'The Civic Holiday is celebrated on the first Monday of August 
        If DateValue = ReportDate.DateFromDayOfWeek(DateValue.Year, 8, 1, FirstDayOfWeek.Monday) Then
          MyMarketHoliday = EnumMarketHolidayCDN.CivicHoliday
          'subtract 3 days to get to the previous Friday
          DateValue = DateValue.AddDays(-3)
        End If
      Case 9
        'Labor Day:
        'The first Monday in September, observed as a holiday in the United States and Canada
        'in honor of working people.
        If DateValue.DayOfWeek = DayOfWeek.Monday Then
          If DateValue = ReportDate.DateFromDayOfWeek(DateValue.Year, 9, 1, FirstDayOfWeek.Monday) Then
            MyMarketHoliday = EnumMarketHolidayCDN.LabourDay
            'this is labor day
            'subtract 3 days to get to the previous Friday
            DateValue = DateValue.AddDays(-3)
          End If
        End If
      Case 10
        'Thanksgiving Day is a holiday celebrated primarily in the United States and Canada. 
        'Traditionally, it has been a time to give thanks to God, friends, and family.
        'Currently, in Canada, Thanksgiving is celebrated on the second Monday of October 
        'and in the United States, it is celebrated on the fourth Thursday of November. 
        'Thanksgiving in Canada falls on the same day as Columbus Day in the United States.
        Dim ThisDateForThanksGiving = ReportDate.DateFromDayOfWeek(DateValue.Year, 10, 2, FirstDayOfWeek.Monday)
        If DateValue = ThisDateForThanksGiving Then
          'this is the Thanksgiving Day holiday
          'subtract 3 days to get to the previous day
          MyMarketHoliday = EnumMarketHolidayCDN.ThanksGivingDay
          DateValue = DateValue.AddDays(-3)
        End If
      Case 12
        'Christmas
        Select Case DateValue.Day
          Case 25
            'this is Christmas
            MyMarketHoliday = EnumMarketHolidayCDN.Christmas
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              DateValue = DateValue.AddDays(-3)
            Else
              DateValue = DateValue.AddDays(-1)
            End If
          Case 26
            'Dec 26 could be an holiday if Christmas did fall on Sunday
            If DateValue.DayOfWeek = DayOfWeek.Monday Then
              MyMarketHoliday = EnumMarketHolidayCDN.Christmas
              'Christmas was on Sunday and in this case the following rule apply:
              'When any stock market holiday falls on a Sunday, the market will be closed the next day (Monday). 
              DateValue = DateValue.AddDays(-3)
            End If
        End Select
    End Select
  End Sub

  Public ReadOnly Property DateValue As Date Implements IHolidayType(Of EnumMarketHolidayCDN).DateValue
    Get
      Return MyDateValue
    End Get
  End Property

  Public ReadOnly Property IsHoliday As Boolean Implements IHolidayType(Of EnumMarketHolidayCDN).IsHoliday
    Get
      If MyMarketHoliday = EnumMarketHolidayCDN.None Then
        Return False
      Else
        Return True
      End If
    End Get
  End Property

  Public ReadOnly Property Holiday As EnumMarketHolidayCDN Implements IHolidayType(Of EnumMarketHolidayCDN).Holiday
    Get
      Return MyMarketHoliday
    End Get
  End Property

  Public ReadOnly Property DateOfLastTradingDay As Date Implements IHolidayType(Of EnumMarketHolidayCDN).DateOfLastTradingDay
    Get
      Return MyDateOfPreviousTradingDay
    End Get
  End Property

  Public ReadOnly Property IsWeekEndDay As Boolean Implements IHolidayType(Of EnumMarketHolidayCDN).IsWeekEndDay
    Get
      Return IsWeekEndDay
    End Get
  End Property

  Public ReadOnly Property DateOfNexTradingDay As Date Implements IHolidayType(Of EnumMarketHolidayCDN).DateOfNexTradingDay
    Get
      Throw New NotImplementedException()
    End Get
  End Property
End Class


Public Interface IHolidayType(Of T)
  ReadOnly Property DateValue As Date

  ReadOnly Property DateOfLastTradingDay As Date
  ReadOnly Property DateOfNexTradingDay As Date

  ReadOnly Property IsWeekEndDay As Boolean
  ReadOnly Property IsHoliday As Boolean
  ReadOnly Property Holiday As T
End Interface
