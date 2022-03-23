Imports YahooAccessData.OptionValuation

Namespace MathPlus.Filter
  Public Class FilterVolatilityMultiple
    Implements IFilter

    Private MyFilterVolatilityMerged As IList(Of Double)
    Private MyListOfVolatilityRegulated As IList(Of Double)
    Private MyLocalFilterVolatilityForList As FilterVolatilityForList
    Private MyMeasurementMethod As IStockOption.enuVolatilityMeasurementMethod

    Private MyFilterValueLast As Double

    ''' <summary>
    ''' Calculate the standard yearly volatility using a monthly windows
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()
      Me.New(
        MeasurementMethod:=IStockOption.enuVolatilityMeasurementMethod.Standard,
        VolatilityType:=IStockOption.enuVolatilityStandardYearlyType.Monthly,
        NumberOfDayToExpiration:=0,
        ListOfVolatilityRegulated:=Nothing)
    End Sub

    Public Sub New(
      ByVal MeasurementMethod As IStockOption.enuVolatilityMeasurementMethod,
      ByVal VolatilityType As IStockOption.enuVolatilityStandardYearlyType)

      Me.New(
        MeasurementMethod:=MeasurementMethod,
        VolatilityType:=VolatilityType,
        NumberOfDayToExpiration:=0,
        ListOfVolatilityRegulated:=Nothing)
    End Sub

    Public Sub New(
      ByVal MeasurementMethod As IStockOption.enuVolatilityMeasurementMethod,
      ByVal VolatilityType As IStockOption.enuVolatilityStandardYearlyType,
      ByVal NumberOfDayToExpiration As Integer)

      Me.New(
        MeasurementMethod,
        VolatilityType,
        NumberOfDayToExpiration,
        ListOfVolatilityRegulated:=Nothing)
    End Sub
    Public Sub New(
      ByVal MeasurementMethod As IStockOption.enuVolatilityMeasurementMethod,
      ByVal VolatilityType As IStockOption.enuVolatilityStandardYearlyType,
      ByVal NumberOfDayToExpiration As Integer,
      ByVal ListOfVolatilityRegulated As IList(Of Double))

      Dim ThisListOfFilter As IList(Of IFilter) = New List(Of IFilter)
      MyFilterVolatilityMerged = New List(Of Double)
      MyMeasurementMethod = MeasurementMethod
      MyListOfVolatilityRegulated = ListOfVolatilityRegulated

      If MyMeasurementMethod = IStockOption.enuVolatilityMeasurementMethod.YangZhangExpRegulated Then
        'special volatility provide by the user
        If MyListOfVolatilityRegulated Is Nothing Then
          Throw New InvalidOperationException("The YangZhangExpRegulated volatility provided input is not valid...")
        End If
        'ignore the rest here we already have the volatility calculated
        Return
      End If
      Select Case VolatilityType
        Case IStockOption.enuVolatilityStandardYearlyType.Daily10
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.Daily15
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.Monthly
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.BiMonthly
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiMonthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.Quaterly
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Quaterly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.BiAnnual
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.BiAnnual), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.Yearly
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.YearlyDaily10
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(10, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.YearlyDaily15
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(15, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.YearlyMonthly
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatility(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly), FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Yearly), FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
        Case IStockOption.enuVolatilityStandardYearlyType.ToExpiration
          If NumberOfDayToExpiration <= 0 Then
            Throw New System.ArgumentOutOfRangeException("NumberOfDayToExpiration must be greater than zero...")
          End If
          Select Case MeasurementMethod
            Case IStockOption.enuVolatilityMeasurementMethod.Standard
              ThisListOfFilter.Add(New FilterVolatility(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardExp
              ThisListOfFilter.Add(New FilterVolatility(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhang
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.YangZhangExp
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Exponential))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMerged
              ThisListOfFilter.Add(New FilterVolatility(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Standard))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Standard))
            Case IStockOption.enuVolatilityMeasurementMethod.StandardYangZhangMergedExp
              ThisListOfFilter.Add(New FilterVolatility(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Exponential))
              ThisListOfFilter.Add(New FilterVolatilityYangZhang(NumberOfDayToExpiration, FilterVolatility.enuVolatilityStatisticType.Exponential))
          End Select
      End Select
      MyLocalFilterVolatilityForList = New FilterVolatilityForList(ThisListOfFilter)
    End Sub

    Public ReadOnly Property Count As Integer Implements IFilter.Count
      Get
        Return MyFilterVolatilityMerged.Count
      End Get
    End Property

    Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
      If MyMeasurementMethod = IStockOption.enuVolatilityMeasurementMethod.YangZhangExpRegulated Then
        'transfert the data
        MyFilterValueLast = MyListOfVolatilityRegulated(MyFilterVolatilityMerged.Count)
      Else
        MyFilterValueLast = MyLocalFilterVolatilityForList.Filter(Value)
      End If
      MyFilterVolatilityMerged.Add(MyFilterValueLast)
      Return MyFilterValueLast
    End Function

    Public Function FilterLast() As Double Implements IFilter.FilterLast
      Return MyFilterValueLast
    End Function

    Public Property Tag As String Implements IFilter.Tag

    Public Function ToArray() As Double() Implements IFilter.ToArray
      Return MyFilterVolatilityMerged.ToArray
    End Function

    Public ReadOnly Property ToList As IList(Of Double) Implements IFilter.ToList
      Get
        Return MyFilterVolatilityMerged
      End Get
    End Property

    Public Overrides Function ToString() As String Implements IFilter.ToString
      Return Me.FilterLast.ToString
    End Function

#Region "Private function"
    Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo

    End Function

    Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast

    End Function

    Private ReadOnly Property ToListOfError As IList(Of Double) Implements IFilter.ToListOfError
      Get

      End Get
    End Property

    Private ReadOnly Property ToListScaled As ListScaled Implements IFilter.ToListScaled
      Get

      End Get
    End Property


    Private Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter

    End Function

    Private Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter

    End Function

    Private Function Filter(Value As Double) As Double Implements IFilter.Filter

    End Function

    Private Function Filter(Value As Single) As Double Implements IFilter.Filter

    End Function

    Private Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray

    End Function

    Private Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray

    End Function

    Private Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol

    End Function

    Private Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext

    End Function

    Private Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext

    End Function

    Private Function Last() As Double Implements IFilter.Last

    End Function

    Private Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol

    End Function

    Private ReadOnly Property Max As Double Implements IFilter.Max
      Get

      End Get
    End Property

    Private ReadOnly Property Min As Double Implements IFilter.Min
      Get

      End Get
    End Property

    Private ReadOnly Property Rate As Integer Implements IFilter.Rate
      Get

      End Get
    End Property

#End Region
#Region "Private Class"
    Private Class FilterVolatilityForList
      Private MyFilterVolatility As IList(Of IFilter)

      Public Sub New(ByVal FilterVolatility As IList(Of IFilter))
        MyFilterVolatility = FilterVolatility
      End Sub

      Public Function Filter(Value As IPriceVol) As Double
        Dim ThisFilterMax As Double = Double.MinValue
        For Each ThisFilter In MyFilterVolatility
          If ThisFilter.Filter(Value) > ThisFilterMax Then
            ThisFilterMax = ThisFilter.FilterLast
          End If
        Next
        Return ThisFilterMax
      End Function
    End Class
#End Region
  End Class
End Namespace
