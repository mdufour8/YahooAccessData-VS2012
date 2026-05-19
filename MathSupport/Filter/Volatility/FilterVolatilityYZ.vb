#Region "Imports"
Imports YahooAccessData.ExtensionService.Extensions
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.MathPlus.Filter.FilterVolatilityYangZhang
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.MathPlus.Probability
Imports YahooAccessData.OptionValuation


#End Region

Namespace MathPlus.Filter
	''' <summary>
	''' FilterVolatilityYZ implements the volatility filter proposed by Yang and Zhang (2000) in a simplified form. 
	''' It uses the last price and volatility to compute a new volatility estimate based on the current price and 
	''' volatility. This version does not keep any list of past values, but only the last value, which makes it more 
	''' efficient and separate the logic of the filter from the logic of the list management. The filter is designed to be 
	''' used in a streaming context, where new price and volatility data are continuously fed into the filter, 
	''' and it updates its estimate accordingly.
	''' </summary>
	Public Class FilterVolatilityYZ
		Implements IVolatilityResult


		Public Const FILTER_RATE_DEFAULT As Integer = MathPlus.NUMBER_TRADINGDAY_PER_YEAR \ 12
		Private MyFilterDirection As FilterRSI.SlopeDirection
		Private MyRate As Integer
		Private MyFilterValueLastK1 As Double
		Private MyFilterValueLast As Double
		Private MyValueLast As YahooAccessData.IPriceVol
		Private MyValueLastK1 As YahooAccessData.IPriceVol
		Private MyFilterVolatilityYearlyCorrection As Double
		Private MyFilterOfOpenToHighToLowAsCloseRatio As Filter.FilterLowPassPLL
		Private MyStatisticOfOpenToHighToLowAsCloseRatio As FilterStatistical
		Private _IsReset As Boolean

		'Private MyCountOfVolNotNull As Integer
		Private MyProbOfOpenToHighToLowAsCloseRatioToSDRatio As Double
		Private MyReturnLogForOpenToPreviousClose As Double
		Private MyReturnLogForCloseToOpen As Double
		Private MyReturnLogForHighToOpen As Double
		Private MyReturnLogForLowToOpen As Double
		Private MyReturnLogForHighToPreviousClose As Double
		Private MyReturnLogForLowToPreviousClose As Double

		Private MyVarianceOpenToHighLowClose_KValue_Yearly As Double
		Private MyVarianceForPreviousCloseToOpen_Vo_Yearly As Double
		Private MyVariancePositifSum_Yearly As Double
		Private MyVarianceNegatifSum_Yearly As Double

		Private IsUseLastSampleHighLowTrailLocal As Boolean

		Private MyStatisticalForOpen As IFilter(Of IStatistical)
		Private MyStatisticalForClose As IFilter(Of IStatistical)
		Private MyStatisticalForOpenToHigh As IFilter(Of IStatistical)
		Private MyStatisticalForOpenToLow As IFilter(Of IStatistical)
		Private MyStatisticalForPreviousCloseToHigh As IFilter(Of IStatistical)
		Private MyStatisticalForPreviousCloseToLow As IFilter(Of IStatistical)
		Private MyFilterExpForPositiveVariance As Filter.FilterLowPassExp
		Private MyFilterExpForNegativeVariance As Filter.FilterLowPassExp

		'Rogers-Satchell is an estimator for measuring the volatility of securities
		'with an average return not equal to zero. Unlike Parkinson and Garman-Klass estimators,
		'Rogers-Satchell incorporates a drift term (mean return not equal to zero).2022
		Private MyStatisticalForVRSHighAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyStatisticalForVRSLowAsClose As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyStatisticalForVRSHigh As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyStatisticalForVRSLow As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyStatisticalForVRSTotal As IFilter(Of IStatistical)    'for Rogers and Satchell statistic
		Private MyValueForK As Double
		Private MyStatisticType As FilterVolatility.enuVolatilityStatisticType
		Private MyVariancePositifSumLast As Double
		Private MyVarianceNegatifSumLast As Double
		Private MyFilterOfVolatilityPositif As Filter.FilterLowPassPLL
		Private MyFilterOfVolatilityNegatif As Filter.FilterLowPassPLL
		Private MyPriceNextDailyHighPreviousCloseToOpenSigma2 As Double

		Public Sub New()
			Me.New(FILTER_RATE_DEFAULT)
		End Sub

		''' <summary>
		''' 
		''' </summary>
		''' <param name="FilterRate">The number of sample use for the volatility measurement</param>
		''' <param name="VolatilityScale">The scale correction factor for the volatility measurement if needed. The default is Volatility per year</param>
		''' <param name="StatisticType">The type of statistical method to use for the volatility measurement </param>
		''' <param name="IsUseLastSampleHighLowTrail">
		'''		Indicates whether to use the last sample's high and low as a trail for the next sample. 
		'''		This is useful to calculate the worst case volatility assuming that the last point high or low was also part of daily steam.
		'''		Interesting to note that this method often match the CBOE Vix volatility index that is based on the SP500 index.
		'''		This measurement can also be used to evaluate the degree of correlation between the previous close and the next day opening.
		'''		A low correlation could indicate outside market hour activity that could be considered abnormal.
		'''	</param>
		Public Sub New(
			FilterRate As Integer,
			Optional StatisticType As FilterVolatility.enuVolatilityStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard,
			Optional VolatilityScale As Double = VOLATILITY_DAILY_TO_YEARLY_RATIO,
			Optional IsUseLastSampleHighLowTrail As Boolean = False)

			IsUseLastSampleHighLowTrailLocal = IsUseLastSampleHighLowTrail
			MyFilterVolatilityYearlyCorrection = VolatilityScale
			MyStatisticType = StatisticType
			MyFilterOfOpenToHighToLowAsCloseRatio = New Filter.FilterLowPassPLL(FilterRate:=FilterRate)
			MyStatisticOfOpenToHighToLowAsCloseRatio = New FilterStatistical(FilterRate:=FilterRate)

			If FilterRate < 2 Then FilterRate = 2
			MyRate = CInt(FilterRate)
			'MyStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard
			Select Case MyStatisticType
				Case FilterVolatility.enuVolatilityStatisticType.Exponential
					MyStatisticalForOpen = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForOpenToLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForOpenToHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForPreviousCloseToHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForPreviousCloseToLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSTotal = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSHigh = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSLow = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSHighAsClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
					MyStatisticalForVRSLowAsClose = New FilterStatistical(FilterRate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				Case Else
					'the default is a windows Square of data statistic
					MyStatisticalForOpen = New FilterStatistical(FilterRate)
					MyStatisticalForClose = New FilterStatistical(FilterRate)
					MyStatisticalForOpenToLow = New FilterStatistical(FilterRate)
					MyStatisticalForOpenToHigh = New FilterStatistical(FilterRate)
					MyStatisticalForPreviousCloseToHigh = New FilterStatistical(FilterRate)
					MyStatisticalForPreviousCloseToLow = New FilterStatistical(FilterRate)
					MyStatisticalForVRSTotal = New FilterStatistical(FilterRate)
					MyStatisticalForVRSHigh = New FilterStatistical(FilterRate)
					MyStatisticalForVRSLow = New FilterStatistical(FilterRate)
					MyStatisticalForVRSHighAsClose = New FilterStatistical(FilterRate)
					MyStatisticalForVRSLowAsClose = New FilterStatistical(FilterRate)
			End Select
			MyFilterExpForPositiveVariance = New Filter.FilterLowPassExp(FilterRate)
			MyFilterExpForNegativeVariance = New Filter.FilterLowPassExp(FilterRate)

			'see paper
			MyValueForK = 0.34 / (1.34 + ((MyRate + 1) / (MyRate - 1)))

			MyFilterValueLast = 0
			MyFilterValueLastK1 = 0
			MyValueLast = New YahooAccessData.PriceVol(0)
			MyValueLastK1 = MyValueLast
			MyFilterDirection = FilterRSI.SlopeDirection.Zero
			_IsReset = True
		End Sub

		Public Function Filter(Value As YahooAccessData.IPriceVol, IsVolatityHoldToLast As Boolean) As Double
			Dim ThisResult As Double

			If IsVolatityHoldToLast Then
				ThisResult = Me.CalculateFilterLocal(MyValueLast)
				'restore the correct value for the last parameters
				MyValueLast = Value
				MyValueLastK1 = MyValueLast
			Else
				ThisResult = Me.CalculateFilterLocal(Value)
			End If
			Return ThisResult
		End Function
		Private Function CalculateFilterLocal(ByRef Value As YahooAccessData.IPriceVol) As Double
			Dim ThisOpenToHighToLowAsCloseRatio As Double
			Dim ThisVRSTotalOpenToHighLow As Double
			Dim ThisVRSPartialOpenToHigh As Double
			Dim ThisVRSPartialOpenToLow As Double
			Dim ThisVRSTotalMeanForOpenToHighLow_Vrs As Double
			Dim ThisVarianceForPreviousCloseToOpen_Vo As Double
			Dim ThisVarianceForOpenToClose_Vc As Double
			Dim ThisVariancePreviousCloseToOpenHighLowClose_V As Double
			Dim ThisVarianceOpenToHighLowClose_KValue As Double
			Dim ThisValueLow As Single
			Dim ThisValueHigh As Single
			Dim ThisVRSMeanUp As Double
			Dim ThisVRSMeanDown As Double
			Dim ThisVarianceForOpenToHigh_Vh As Double
			Dim ThisVarianceForOpenToLow_Vl As Double
			Dim ThisVariancePositifSum As Double
			Dim ThisVarianceNegatifSum As Double
			Dim ThisVRSUp2 As Double
			Dim ThisVRSDown2 As Double

			'this copy is needed to make sure the IsUseLastSampleHighLowTrailLocal is not modifying the original value
			'and is only used for the local calculation of the volatility
			ThisValueLow = Value.Low
			ThisValueHigh = Value.High
			If _IsReset Then
				'measure volatility from initial value only at start
				MyFilterValueLast = 0
				MyValueLast = New PriceVol(Value.Open)
				MyReturnLogForOpenToPreviousClose = LogPriceReturn(Value.Open, MyValueLast.Last)
				MyReturnLogForLowToOpen = LogPriceReturn(ThisValueLow, Value.Open)
				'MyReturnLogForLowToOpen = LogPriceReturn(Value.Open, ThisValueLow)
				MyReturnLogForHighToOpen = LogPriceReturn(ThisValueHigh, Value.Open)
				MyReturnLogForCloseToOpen = LogPriceReturn(Value.Last, Value.Open)
			End If
			If IsUseLastSampleHighLowTrailLocal Then
				'just some test code and it does not seem to be a good idea to use the last sample high low as trail for the next sample,
				'but it was worth testing
				'Select Case Value.Open '(MyValueLast.Low + MyValueLast.High) / 2
				'	Case = MyValueLast.Last
				'		If MyValueLast.Low < ThisValueLow Then
				'			ThisValueLow = MyValueLast.Low
				'		End If
				'		If MyValueLast.High > ThisValueHigh Then
				'			ThisValueHigh = MyValueLast.High
				'		End If
				'	Case > MyValueLast.Last
				'		If MyValueLast.Low < ThisValueLow Then
				'			ThisValueLow = MyValueLast.Low
				'		End If
				'	Case < MyValueLast.Last
				'		If MyValueLast.High > ThisValueHigh Then
				'			If MyValueLast.High > ThisValueHigh Then
				'				ThisValueHigh = MyValueLast.High
				'			End If
				'		End If
				'End Select
				'first let see if we are in an extreme volatility period
				'where the current high is below the previouss day low 
				'replace the high with le previous day low to calculate the volatility in the worst case scenario
				'similarly if the current low is above the previous day high
				'replace the low with the previous day high to calculate the volatility in the worst case scenario
				'
				If ThisValueHigh < MyValueLast.Low Then
					ThisValueHigh = MyValueLast.Low
				ElseIf ThisValueLow > MyValueLast.High Then
					ThisValueLow = MyValueLast.High
				Else
					'where the open is above the previous close and	
					'the low is above the previous low or the open is below the previous close and
					'the high is below the previous high, in that case we can consider that
					'the last sample high low was also part of the daily steam and we can use it as a trail for the next sample
					If MyValueLast.Low < ThisValueLow Then
						ThisValueLow = MyValueLast.Low
					End If
					If MyValueLast.High > ThisValueHigh Then
						ThisValueHigh = MyValueLast.High
					End If
				End If
			End If
			'filter for value less than zero
			'If DirectCast(Value, PriceVol).IsNull = False Then
			'        MyCountOfVolNotNull = MyCountOfVolNotNull + 1
			'End If

			'ln(Open1/Close0)
			MyReturnLogForOpenToPreviousClose = LogPriceReturn(Value.Open, MyValueLast.Last)
			'ln(Low1/Open1)
			MyReturnLogForLowToOpen = LogPriceReturn(ThisValueLow, Value.Open)
			'ln(High1/Open1)
			MyReturnLogForHighToOpen = LogPriceReturn(ThisValueHigh, Value.Open)
			'ln(Close1/Open1)
			MyReturnLogForCloseToOpen = LogPriceReturn(Value.Last, Value.Open)

			Dim ThisReturnLogForPreviousHighToOpen = LogPriceReturn(Value.Open, MyValueLast.High)
			Dim ThisReturnLogForPreviousLowToOpen = LogPriceReturn(Value.Open, MyValueLast.Low)

			'VRS calculation
			'ThisVRSUp2 = (ln(High/Open)^ 2) + ((ln(Last/Low) ^ 2))


			'this is a local calcul to measure and compare the volatility Up and down of the stock
			'it is an intraday calculation that is very predictive and related to the comportment of the stock in the future
			'but it is not related and is not needed for the VolatilityYangZhang calculation
			'Note that some observation show the opening to be a bit more predictive than the closing
			'that is the reason we divide the close by 2
			'however other trader seem to indicate that the close is more predictive
			'an investigation is needed to clear this aspect
			'to do: calculate both for testing in the future
			'ThisVRSUp2 = (MyReturnLogForHighToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.Low) ^ 2) / 2)
			'ThisVRSDown2 = (MyReturnLogForLowToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.High) ^ 2) / 2)
			'modified mars 2024 taking into account the previous close
			ThisVRSUp2 = ((ThisReturnLogForPreviousLowToOpen ^ 2) / 2) + (MyReturnLogForHighToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.Low) ^ 2) / 2)
			ThisVRSDown2 = ((ThisReturnLogForPreviousHighToOpen ^ 2) / 2) + (MyReturnLogForLowToOpen ^ 2) + ((LogPriceReturn(Value.Last, Value.High) ^ 2) / 2)

			ThisVRSPartialOpenToHigh = MyReturnLogForHighToOpen * (MyReturnLogForHighToOpen - MyReturnLogForCloseToOpen)
			ThisVRSPartialOpenToLow = MyReturnLogForLowToOpen * (MyReturnLogForLowToOpen - MyReturnLogForCloseToOpen)

			MyReturnLogForHighToPreviousClose = LogPriceReturn(ThisValueHigh, MyValueLast.Last)
			MyReturnLogForLowToPreviousClose = LogPriceReturn(ThisValueLow, MyValueLast.Last)

			'If Me.Count = 1576 Then
			'ThisVarianceNegatifSum = ThisVarianceNegatifSum
			'End If

			'ThisVRSPartialOpenToHigh = MyReturnLogForHighToOpen * (MyReturnLogForHighToOpen - 0)
			'ThisVRSPartialOpenToLow = MyReturnLogForLowToOpen * (MyReturnLogForLowToOpen - 0)
			'this is the Vrs(i) value in the ref. paper, but we need to average it to get it current value
			ThisVRSTotalOpenToHighLow = ThisVRSPartialOpenToHigh + ThisVRSPartialOpenToLow

			'calculate the variance for the open and close
			'This is Vo in the ref. paper
			ThisVarianceForPreviousCloseToOpen_Vo = MyStatisticalForOpen.Filter(MyReturnLogForOpenToPreviousClose).Variance
			'This is Vc in the ref. paper
			ThisVarianceForOpenToClose_Vc = MyStatisticalForClose.Filter(MyReturnLogForCloseToOpen).Variance
			'those two variable are nor use in the calculation
			'may be of interesting to evaluate
			ThisVarianceForOpenToHigh_Vh = MyStatisticalForOpenToHigh.Filter(MyReturnLogForHighToOpen).Variance
			ThisVarianceForOpenToLow_Vl = MyStatisticalForOpenToLow.Filter(MyReturnLogForLowToOpen).Variance

			'not used right now
			'ThisVarianceForPreviousCloseToHigh_Vh = MyStatisticalForOpenToLow.Filter(MyReturnLogForHighToPreviousClose).Variance
			'ThisVarianceForPreviousCloseToLow_Vl = MyStatisticalForOpenToLow.Filter(MyReturnLogForLowToPreviousClose).Variance

			'calculate le mean for the total high and low variation to close
			'this is Vrs in the ref. paper
			ThisVRSTotalMeanForOpenToHighLow_Vrs = MyStatisticalForVRSTotal.Filter(ThisVRSTotalOpenToHighLow).Mean

			ThisVarianceOpenToHighLowClose_KValue = MyValueForK * ThisVarianceForOpenToClose_Vc + (1 - MyValueForK) * ThisVRSTotalMeanForOpenToHighLow_Vrs
			'this is the final volatility
			ThisVariancePreviousCloseToOpenHighLowClose_V = ThisVarianceForPreviousCloseToOpen_Vo + ThisVarianceOpenToHighLowClose_KValue

			MyVarianceOpenToHighLowClose_KValue_Yearly = ToVolatilityYearly(ThisVarianceOpenToHighLowClose_KValue)
			MyVarianceForPreviousCloseToOpen_Vo_Yearly = ToVolatilityYearly(ThisVarianceForPreviousCloseToOpen_Vo)

			'separate the variance in positif and negatif value for calculation of the PriceVolatilityPositif measurement
			'ThisVariancePositifSum = (1 - MyValueForK) * (ThisVRSMeanUp / ThisVariancePreviousCloseToOpenHighLowClose_V)
			'ThisVarianceNegatifSum = (1 - MyValueForK) * (ThisVRSMeanDown / ThisVariancePreviousCloseToOpenHighLowClose_V)
			ThisVRSMeanUp = MyStatisticalForVRSHigh.Filter(ThisVRSUp2).Mean
			ThisVRSMeanDown = MyStatisticalForVRSLow.Filter(ThisVRSDown2).Mean
			ThisVariancePositifSum = (1 - MyValueForK) * ThisVRSMeanUp
			ThisVarianceNegatifSum = (1 - MyValueForK) * ThisVRSMeanDown
			'correct the value for the yearly variation

			MyFilterValueLastK1 = MyFilterValueLast
			MyFilterValueLast = ToVolatilityYearly(ThisVariancePreviousCloseToOpenHighLowClose_V)

			'~~~~~~~~~~~~~~~
			'the filter is not used right now
			MyFilterExpForPositiveVariance.Filter(ThisVariancePositifSum)
			MyFilterExpForNegativeVariance.Filter(ThisVarianceNegatifSum)
			MyVariancePositifSum_Yearly = ToVolatilityYearly(ThisVariancePositifSum)
			MyVarianceNegatifSum_Yearly = ToVolatilityYearly(ThisVarianceNegatifSum)

			Dim ThisSumOfVolatilityPositifNegatif = ThisVariancePositifSum + ThisVarianceNegatifSum
			'the balance value of this ratio is 0.5
			If ThisSumOfVolatilityPositifNegatif > 0 Then
				ThisOpenToHighToLowAsCloseRatio = ThisVariancePositifSum / ThisSumOfVolatilityPositifNegatif
			Else
				ThisOpenToHighToLowAsCloseRatio = 0.5
			End If
			If _IsReset = False Then
				'this is been shown to be much better and predictive even if more testing may be needed to confirm this aspect,
				'but it is interesting to note that the balance between the volatility up and down of the stock is a good predictor
				'of the future volatility of the stock
				Select Case ThisOpenToHighToLowAsCloseRatio
					Case > 0.5
						MyFilterDirection = FilterRSI.SlopeDirection.Positive
					Case < 0.5
						MyFilterDirection = FilterRSI.SlopeDirection.Negative
					Case Else
						MyFilterDirection = FilterRSI.SlopeDirection.Zero
				End Select
			Else
				'Zero at reset to avoid any bias at the start of the filter
				MyFilterDirection = FilterRSI.SlopeDirection.Zero
			End If
			'If Me.Count = 500 Then
			'	MyFilterDirection = MyFilterDirection
			'End If
			Dim ThisSD = MyStatisticOfOpenToHighToLowAsCloseRatio.Filter(ThisOpenToHighToLowAsCloseRatio).StandardDeviation
			'ThisOpenToHighToLowAsCloseRatio is contain between [0,1] assume 25% of the maximum for the standard deviation
			ThisSD = If(ThisSD > 0, ThisSD, 0.25)
			Dim ThisOpenToHighToLowAsCloseRatioToSDRatio = (ThisOpenToHighToLowAsCloseRatio - 0.5) / ThisSD
			'Assuming a Gaussian probability what is the equivalent probability 
			MyProbOfOpenToHighToLowAsCloseRatioToSDRatio = ProbabilityMapping.GaussianCDF(ThisOpenToHighToLowAsCloseRatioToSDRatio)
			''now put that probability on a Gaussian scale for display
			''bit a compression around 0.5 and expansion at higher probability
			'ThisProbOfOpenToHighToLowAsCloseRatioToSDRatio = ProbabilityMapping.ProbabilityToGaussianScale(ThisProbOfOpenToHighToLowAsCloseRatioToSDRatio, ScaleOfX:=2.0)
			MyFilterOfOpenToHighToLowAsCloseRatio.Filter(MyProbOfOpenToHighToLowAsCloseRatioToSDRatio)
			'~~~~~~~~~~~~~~~

			MyValueLastK1 = MyValueLast
			MyValueLast = Value
			_IsReset = False
			Return MyFilterValueLast
		End Function

		'''' <summary>
		'''' Compute the volatility using the standard Logarithmic or Continuously Compounded Return method. 
		'''' This method expect positive value of asset price and use the YangZhang method

		'''' </summary>
		'''' <param name="Value">The current positive value of the asset</param>
		'''' <remarks>
		'''' The function assume by default a daily data input and scale the result to a yearly period. The scale factor may need to be adjusted if the data
		'''' is not at the daily sample rate or another type of volatility is needed.
		'''' </remarks>
		'''' 


		''' <summary>
		''' 
		''' </summary>
		''' <param name="Value"></param>
		''' <returns></returns>
		Public Function Filter(ByVal Value As YahooAccessData.IPriceVol) As Double
			Return Me.Filter(Value, IsVolatityHoldToLast:=False)
		End Function


		Public Function FilterLast() As Double
			Return MyFilterValueLast
		End Function

		Public Function FilterLast(ByVal Type As enuVolatilityDailyPeriodType) As Double
			Select Case Type
				Case enuVolatilityDailyPeriodType.FullDay
					Return MyFilterValueLast
				Case enuVolatilityDailyPeriodType.OpenToClose
					Return MyVarianceOpenToHighLowClose_KValue_Yearly
				Case enuVolatilityDailyPeriodType.PreviousCloseToOpen
					Return MyVarianceForPreviousCloseToOpen_Vo_Yearly
				Case enuVolatilityDailyPeriodType.OpenToHighClose
					Return MyVariancePositifSum_Yearly
				Case enuVolatilityDailyPeriodType.OpenToLowClose
					Return MyVarianceNegatifSum_Yearly
				Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio
					Return MyProbOfOpenToHighToLowAsCloseRatioToSDRatio
				Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatioFiltered
					Return MyFilterOfOpenToHighToLowAsCloseRatio.FilterLast
				Case Else
					Return MyFilterValueLast
			End Select
		End Function

		Public Function FilterDirection() As FilterRSI.SlopeDirection
			Return MyFilterDirection
		End Function

		Public Function Last() As IPriceVol
			Return MyValueLast
		End Function

		Public ReadOnly Property Rate As Integer
			Get
				Return MyRate
			End Get
		End Property

		Public Property Tag As String

		Public Overrides Function ToString() As String
			Return Me.FilterLast.ToString
		End Function

		Public Function Filter(Value As Double) As Double
			Return Me.Filter(New PriceVol(CSng(Value)))
		End Function

		Public Function Filter(Value As Single) As Double
			Return Me.Filter(New PriceVol(Value))
		End Function

		Public Sub Reset()
			_IsReset = True
		End Sub

		Private Function ToVolatilityYearly(ByVal VariancePerDay As Double) As Double
			Return MyFilterVolatilityYearlyCorrection * Math.Sqrt(VariancePerDay)
		End Function

#Region "IVolatilityResult"
		Public ReadOnly Property FullDay As Double Implements IVolatilityResult.FullDay
			Get
				Return MyFilterValueLast
			End Get
		End Property

		Public ReadOnly Property PreviousCloseToOpen As Double Implements IVolatilityResult.PreviousCloseToOpen
			Get
				Return MyVarianceForPreviousCloseToOpen_Vo_Yearly
			End Get
		End Property

		Public ReadOnly Property OpenToClose As Double Implements IVolatilityResult.OpenToClose
			Get
				Return MyVarianceOpenToHighLowClose_KValue_Yearly
			End Get
		End Property

		Public ReadOnly Property OpenToHighClose As Double Implements IVolatilityResult.OpenToHighClose
			Get
				Return MyVariancePositifSum_Yearly
			End Get
		End Property

		Public ReadOnly Property OpenToLowClose As Double Implements IVolatilityResult.OpenToLowClose
			Get
				Return MyVarianceNegatifSum_Yearly
			End Get
		End Property

		Public ReadOnly Property OpenToHighToLowCloseRatio As Double Implements IVolatilityResult.OpenToHighToLowCloseRatio
			Get
				Return MyProbOfOpenToHighToLowAsCloseRatioToSDRatio
			End Get
		End Property

		Public ReadOnly Property OpenToHighToLowCloseRatioFiltered As Double Implements IVolatilityResult.OpenToHighToLowCloseRatioFiltered
			Get
				Return MyFilterOfOpenToHighToLowAsCloseRatio.FilterLast
			End Get
		End Property

		Public ReadOnly Property Rogers_Satchell_Yoon_Vrs As Double Implements IVolatilityResult.Rogers_Satchell_Yoon_Vrs
			Get
				Return MyFilterValueLast
			End Get
		End Property

		Public ReadOnly Property Parkison_Vp As Double Implements IVolatilityResult.Parkison_Vp
			Get
				Return MyFilterValueLast
			End Get
		End Property
#End Region
	End Class
End Namespace