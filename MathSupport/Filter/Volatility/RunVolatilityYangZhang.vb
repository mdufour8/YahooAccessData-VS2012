Namespace MathPlus.Filter
  Friend Class RunVolatilityYangZhang

    Private MyStatisticalForOpen As IFilter(Of IStatistical)
    Private MyStatisticalForClose As IFilter(Of IStatistical)
    Private MyStatisticalForVRSHighAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSLowAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSHigh As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSLow As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyStatisticalForVRSTotal As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
    Private MyValueForK As Double
    Private MyStatisticType As FilterVolatility.enuVolatilityStatisticType


    Private IsUseLastSampleHighLowTrailLocal As Boolean

    Public Sub New(ByVal IsUseLastSampleHighLowTrail As Boolean)
      IsUseLastSampleHighLowTrailLocal = IsUseLastSampleHighLowTrail
    End Sub

    ''' <summary>
    ''' Calculate the key parameters for the volatility
    ''' </summary>
    ''' <param name="Value">
    '''   the price data
    ''' </param>
    ''' <param name="IsVolatityHoldToLast">
    '''   Can be use to hold the previous value in case of invalid or abnorval price value
    '''  </param>
    ''' <returns></returns>
    Public Function Run(ByRef Value As YahooAccessData.IPriceVol, ByRef ValueLast As YahooAccessData.IPriceVol, ByVal IsVolatityHoldToLast As Boolean) As Double
      Dim ThisVRSTotalOpenToHighLow As Double
      Dim ThisVRSPartialOpenToHigh As Double
      Dim ThisVRSPartialOpenToLow As Double
      Dim ThisVRSForHighFromLast As Double
      Dim ThisVRSForLowFromLast As Double
      Dim ThisVRSTotalMeanForOpenToHighLow_Vrs As Double
      Dim ThisVarianceForPreviousCloseToOpen_Vo As Double
      Dim ThisVarianceForOpenToClose_Vc As Double
      Dim ThisVariancePreviousCloseToOpenHighLowClose_V As Double
      Dim ThisVarianceOpenToHighLowClose_KValue As Double
      Dim ThisVarianceCloseToHigh As Double
      Dim ThisVarianceOpenToVRSLowClose As Double
      Dim ThisVarianceOpenToVRSHighClose As Double
      Dim ThisVarianceCloseToLow As Double
      Dim ThisValueLow As Single
      Dim ThisValueHigh As Single
      Dim ThisReturnLogForPreviousCloseToOpen As Double
      Dim ThisReturnLogForOpenToLow As Double
      Dim ThisReturnLogForOpenToHigh As Double
      Dim ThisReturnLogForOpenToClose As Double


      ThisValueLow = Value.Low
      ThisValueHigh = Value.High

      If IsUseLastSampleHighLowTrailLocal Then
        If ValueLast.Low < ThisValueLow Then
          ThisValueLow = ValueLast.Low
        End If
        If ValueLast.High > ThisValueHigh Then
          ThisValueHigh = ValueLast.High
        End If
      End If
      'filter for value less than zero
      'If DirectCast(Value, PriceVol).IsNull = False Then
      '        MyCountOfVolNotNull = MyCountOfVolNotNull + 1
      'End If


      ThisReturnLogForPreviousCloseToOpen = LogPriceReturn(Value.Open, ValueLast.Last)
      ThisReturnLogForOpenToLow = LogPriceReturn(ThisValueLow, Value.Open)
      ThisReturnLogForOpenToHigh = LogPriceReturn(ThisValueHigh, Value.Open)
      ThisReturnLogForOpenToClose = LogPriceReturn(Value.Last, Value.Open)
      ThisVRSPartialOpenToHigh = ThisReturnLogForOpenToHigh * (ThisReturnLogForOpenToHigh - ThisReturnLogForOpenToClose)
      ThisVRSPartialOpenToLow = ThisReturnLogForOpenToLow * (ThisReturnLogForOpenToLow - ThisReturnLogForOpenToClose)
      ThisVRSTotalOpenToHighLow = ThisVRSPartialOpenToHigh + ThisVRSPartialOpenToLow

      'calculate the variance for the open and close
      'This is Vo in the ref. paper
      ThisVarianceForPreviousCloseToOpen_Vo = MyStatisticalForOpen.Filter(ThisReturnLogForPreviousCloseToOpen).Variance
      'This is Vc in the ref. paper
      ThisVarianceForOpenToClose_Vc = MyStatisticalForClose.Filter(ThisReturnLogForOpenToClose).Variance
      'calculate le mean for the total high and low variation to close
      'this is Vrs in the ref. paper
      ThisVRSTotalMeanForOpenToHighLow_Vrs = MyStatisticalForVRSTotal.Filter(ThisVRSTotalOpenToHighLow).Mean
      ThisVarianceOpenToHighLowClose_KValue = MyValueForK * ThisVarianceForOpenToClose_Vc + (1 - MyValueForK) * ThisVRSTotalMeanForOpenToHighLow_Vrs
      'thsi si the final volatility
      ThisVariancePreviousCloseToOpenHighLowClose_V = ThisVarianceForPreviousCloseToOpen_Vo + ThisVarianceOpenToHighLowClose_KValue


      'correct the value for the yearly variation
      MyFilterValueLastK1 = MyFilterValueLast
      MyFilterValueLast = ToYearCorrected(ThisVariancePreviousCloseToOpenHighLowClose_V)
      MyListOfPreviousCloseToOpenHighLowClose.Add(MyFilterValueLast)
      MyListOfPreviousCloseToOpen.Add(ToYearCorrected(ThisVarianceForPreviousCloseToOpen_Vo))
      MyListOfOpenToClose.Add(ToYearCorrected(ThisVarianceOpenToHighLowClose_KValue))

      '~~~~~~~~~~~~~~~
      MyReturnLogForOpenToHighFromLast = LogPriceReturn(ThisValueHigh, ValueLast.Last)
      MyReturnLogForOpenToLowFromLast = LogPriceReturn(ThisValueLow, ValueLast.Last)
      MyReturnLogForOpenToCloseFromLast = LogPriceReturn(Value.Last, ValueLast.Last)
      ThisVRSForLowFromLast = MyReturnLogForOpenToLowFromLast * (MyReturnLogForOpenToLowFromLast - MyReturnLogForOpenToCloseFromLast)
      ThisVRSForHighFromLast = MyReturnLogForOpenToHighFromLast * (MyReturnLogForOpenToHighFromLast - MyReturnLogForOpenToCloseFromLast)
      'get the VRS mean
      ThisVRSForLowFromLast = MyStatisticalForVRSLow.Filter(ThisVRSPartialOpenToLow).Mean
      ThisVRSForHighFromLast = MyStatisticalForVRSHigh.Filter(ThisVRSPartialOpenToHigh).Mean

      'Get the variance
      ThisVarianceCloseToLow = MyStatisticalForVRSLowAsClose.Filter(MyReturnLogForOpenToLowFromLast).Variance
      ThisVarianceCloseToHigh = MyStatisticalForVRSHighAsClose.Filter(MyReturnLogForOpenToHighFromLast).Variance
      ThisVarianceOpenToVRSLowClose = MyValueForK * ThisVarianceForOpenToClose_Vc + (1 - MyValueForK) * ThisVRSForLowFromLast
      ThisVarianceOpenToVRSHighClose = MyValueForK * ThisVarianceForOpenToClose_Vc + (1 - MyValueForK) * ThisVRSForHighFromLast

      MyListOfOpenHighAsClose.Add(ToYearCorrected(ThisVarianceOpenToVRSHighClose))
      'MyListOfOpenLowAsClose.Add(ToYearCorrected(ThisVarianceCloseToLow))
      MyListOfOpenLowAsClose.Add(ToYearCorrected(ThisVarianceOpenToVRSLowClose))

      '~~~~~~~~~~~~~~~

      MyValueLastK1 = ValueLast
      ValueLast = Value
      Return MyFilterValueLast


    End Function
  End Class
End Namespace