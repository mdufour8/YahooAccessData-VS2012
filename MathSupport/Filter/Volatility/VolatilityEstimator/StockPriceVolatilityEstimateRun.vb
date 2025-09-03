
#Const PARALLEL_EXECUTION = False

Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter
Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Namespace OptionValuation
	''' <summary>
	''' This class is a container for the stock price prediction run. 
	''' It contains the stock price prediction run data and the stock price prediction run results.	
	''' It os a simplified version of the StockPriceVolatilityPredictionBand class specialized for 
	''' the specific need to adjust teh stock price prediction level in function of the volatility.
	''' </summary>
	Public Class StockPriceVolatilityEstimateRun

		Private Const FILTER_RATE_FOR_GAIN As Integer = 5
		Private Const FILTER_RATE_FOR_VOLATILITY_CORRECTION As Integer = 20

		Private _stockPricePredictionData As StockPriceVolatilityEstimateData
		Private _StockPriceLast As IPriceVol

		Private MyNumberTradingDays As Double
		Private MyVolatilityMeasurementPeriod As Integer
		Private MyProbabilityOfInterval As Double
		Private MyVolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		Private MyQueue As Queue(Of StockPriceVolatilityEstimateData)
		Private MySumOfGainLog As Double
		Private MySumOfThresholdExcess As Integer
		Private MySumOfThresholdExcessHigh As Integer
		Private MySumOfThresholdExcessLow As Integer
		Private MyProbabilityOfThresholdExcess As Double
		Private MyProbabilityOfThresholdExcessHigh As Double
		Private MyProbabilityOfThresholdExcessLow As Double
		Private MyProbabilityHigh As Double
		Private MyProbabilityLow As Double
		Private MyProbabilityRatioToSigma As Double

		Private MyFilterForProbability As FilterExp
		Private MyStatisticalOfProbOfExcesDeltaHighLow As FilterStatistical
		Private MyFilterBrownExpPredict As FilterExpPredict
		Private MyFilterBrownExpPredictDerivative As FilterExpPredict
		Private MyFilterVolatilityYZ As FilterVolatilityYangZhang
		Private MyStatisticalOfStockPriceGain As FilterStatistical
		Private MyVolatilityDelta As Double
		Private MyVolatilityDeltaHigh As Double
		Private MyVolatilityDeltaLow As Double
		Private MyVolatilitySumOfError As Double
		Private MyFilterPLLHigh As FilterPLL
		Private MyFilterPLLLow As FilterPLL
		Private MyVolatilityEstimate As Double


		''' <summary>
		''' 
		''' </summary>
		''' <param name="NumberTradingDays">The prediction period in days use for calculating the 1 sigma threshold given the volatility</param>
		''' <param name="VolatilityMeasurementPeriodInDays">Filter Volatility period in days </param>
		''' <param name="ProbabilityOfInterval"></param>
		''' <param name="VolatilityPredictionBandType"></param>
		Public Sub New(
				ByVal NumberTradingDays As Double,
				ByVal VolatilityMeasurementPeriodInDays As Integer,
				ByVal ProbabilityOfInterval As Double,
				ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)

			MyNumberTradingDays = NumberTradingDays
			MyProbabilityOfInterval = ProbabilityOfInterval
			MyProbabilityOfThresholdExcess = MyProbabilityOfInterval
			MyProbabilityOfThresholdExcessHigh = MyProbabilityOfInterval / 2
			MyProbabilityOfThresholdExcessLow = MyProbabilityOfInterval / 2
			MyProbabilityHigh = 0.5 + MyProbabilityOfThresholdExcessHigh
			MyProbabilityLow = 0.5 - MyProbabilityOfThresholdExcessLow
			MyProbabilityRatioToSigma = MyProbabilityHigh / Measure.GAUSSIAN_PROBABILITY_MEAN_PLUS_SIGMA1

			MyVolatilityPredictionBandType = VolatilityPredictionBandType
			MyVolatilityMeasurementPeriod = VolatilityMeasurementPeriodInDays
			MyFilterForProbability = New FilterExp(FilterRate:=VolatilityMeasurementPeriodInDays)
			MyQueue = New Queue(Of StockPriceVolatilityEstimateData)(capacity:=VolatilityMeasurementPeriodInDays)
			MyStatisticalOfProbOfExcesDeltaHighLow = New FilterStatistical(
				FilterRate:=VolatilityMeasurementPeriodInDays,
				StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)

			MyStatisticalOfStockPriceGain = New FilterStatistical(
				FilterRate:=VolatilityMeasurementPeriodInDays,
				StatisticType:=FilterVolatility.enuVolatilityStatisticType.Standard)

			'this Is base on a	Brown exponential filter to calculate the gain And the derivative of the stock price
			'this Is use when the user does not provide a value for the Gain And its derivative
			MyFilterBrownExpPredict = New FilterExpPredict(FilterRate:=FILTER_RATE_FOR_GAIN) With {.Tag = "StockPriceVolatilityEstimateRun_BrownExpPredict"}

			MyFilterBrownExpPredictDerivative = New FilterExpPredict(FilterRate:=FILTER_RATE_FOR_GAIN) With {.Tag = "StockPriceVolatilityEstimateRun_BrownExpPredictDerivarive"}

			MyFilterVolatilityYZ = New FilterVolatilityYangZhang(FilterRate:=VolatilityMeasurementPeriodInDays, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Standard)
			MyFilterPLLHigh = New FilterPLL(FILTER_RATE_FOR_VOLATILITY_CORRECTION)
			MyFilterPLLLow = New FilterPLL(FILTER_RATE_FOR_VOLATILITY_CORRECTION)
			_StockPriceLast = Nothing
			MyVolatilityDelta = 0
			MyVolatilityDeltaHigh = 0
			MyVolatilityDeltaLow = 0
		End Sub

		''' <summary>
		''' Add an element in the Queue use for volatility estimation. 
		''' This method will calculate the gain and its derivative with the volatility using a Brown exponential filter and the powerful Yang-Zhang volatility measurements method. It is recommended to use
		''' this method when the user does not want to provide a value for the voaltility, Gain and Gain derivative . The method provide a very good estimate 
		''' of the gain and its derivative but also of the volatility. 
		''' </summary>
		''' <param name="StockPrice"></param>
		''' <returns></returns>
		Public Function Add(
			ByVal StockPrice As IPriceVol) As Double

			Dim ThisVolatility As Double = MyFilterVolatilityYZ.Filter(StockPrice)
			Return Me.Add(StockPrice, ThisVolatility)
		End Function

		''' <summary>
		''' Add an element in the Queue use for volatility estimation. 
		''' This method will calculate the gain and its derivative using a Brown exponential filter. It is recommended to use
		''' this method when the user does not provide a value for the Gain and its derivative. The method provide a very good estimate 
		''' of the gain and its derivative. 
		''' </summary>
		''' <param name="StockPrice">The surrent stock price</param>
		''' <param name="Volatility">The current volatility</param>
		''' <returns></returns>
		Public Function Add(
			ByVal StockPrice As IPriceVol,
			ByVal Volatility As Double) As Double

			Dim ThisGainDailyLog As Double
			Dim ThisGainYearlyLog As Double
			Dim ThisGainDailyLogDerivative As Double
			Dim ThisGainYearlyLogDerivative As Double

			With MyFilterBrownExpPredict
				.FilterRun(StockPrice.Last)
				ThisGainDailyLog = .GainLog
				ThisGainDailyLogDerivative = .GainLogDerivative
				ThisGainYearlyLog = MathPlus.General.NUMBER_TRADINGDAY_PER_YEAR * ThisGainDailyLog
				ThisGainYearlyLogDerivative = MathPlus.General.NUMBER_TRADINGDAY_PER_YEAR * ThisGainDailyLogDerivative
			End With
			'With MyFilterBrownExpPredictDerivative
			'	.FilterRun(ThisGainYearlyLog)
			'	ThisGainDailyLogDerivative = .GainLogDerivative
			'	ThisGainYearlyLogDerivative = MathPlus.General.NUMBER_TRADINGDAY_PER_YEAR * ThisGainDailyLogDerivative
			'End With
			'for testing
			'ThisGainYearlyLog = 0.0
			'ThisGainYearlyLogDerivative = 0.0
			Return Me.Add(
				StockPrice:=StockPrice,
				Gain:=ThisGainYearlyLog,
				GainDerivative:=ThisGainYearlyLogDerivative,
				Volatility:=Volatility)

		End Function

		''' <summary>
		''' Add an element in the Queue use for volatility estimation
		''' </summary>
		''' <param name="StockPrice"></param>
		''' <param name="Gain"></param>
		''' <param name="GainDerivative"></param>
		''' <param name="Volatility"></param>
		''' <returns></returns>
		Public Function Add(
			ByVal StockPrice As IPriceVol,
			ByVal Gain As Double,
			ByVal GainDerivative As Double,
			ByVal Volatility As Double) As Double

			Dim ThisQueueDataLastDate As Date = Nothing
			Dim ThisQueueDataRemovedDate As Date = Nothing
			Dim ThisQueueDataLast As StockPriceVolatilityEstimateData
			Dim ThisQueueDataRemoved As StockPriceVolatilityEstimateData
			Dim ThisStockPricePredictionData As StockPriceVolatilityEstimateData
			Dim ThisEstimateOfError As Double
			Dim ThisLogNormalMu As Double
			Dim ThisLogNormalSigma As Double
			Dim ThisVolatilityEstimate As Double
			Dim ThisResult As (ProbabilityOfThresholdExcess As Double, ProbabilityOfThresholdExcessHigh As Double, MyProbabilityOfThresholdExcessLow As Double)

			'Gain = 0
			'GainDerivative = 0
			If (StockPrice.Vol = 0) OrElse (Volatility = 0) Then
				Return MyVolatilityEstimate
			End If
			If _StockPriceLast Is Nothing Then
				_StockPriceLast = StockPrice
			End If
			'it is important to make sure the previous last value is updated so that the log normal GainLog property parameters
			'in StockPriceVolatilityEstimateData is calculated correctly before the object is created
			If StockPrice.LastPrevious = 0 Then
				StockPrice.LastPrevious = _StockPriceLast.Last
			End If
			' Create a new StockPricePredictionData instance with the provided parameters
			ThisStockPricePredictionData = New StockPriceVolatilityEstimateData(
				NumberTradingDays:=MyNumberTradingDays,
				StockPrice:=StockPrice,
				Gain:=Gain,
				GainDerivative:=GainDerivative,
				Volatility:=Volatility,
				MyProbabilityOfInterval,
				MyVolatilityPredictionBandType)

			' Note: The next sample is not yet available. In this case, the class will use the intra-daily band instead of the next day band.
			' This should still provide a good indication of the stock price's real volatility estimate.
			ThisStockPricePredictionData.Refresh(MyVolatilitySumOfError)

			' If there are items in the queue, update the last item with the current stock price
			If MyQueue.Count > 0 Then
				With MyQueue.Last
					ThisQueueDataLastDate = .StockPrice.DateDay
					.StockPriceNext = StockPrice
					.Refresh()
				End With
			End If
			'Update the queue by removing the oldest item if the queue has reached its maximum size
			If MyQueue.Count = MyVolatilityMeasurementPeriod Then
				ThisQueueDataRemoved = MyQueue.Dequeue
				With ThisQueueDataRemoved
					ThisQueueDataRemovedDate = .StockPrice.DateDay
				End With
			End If
			' Add the new data to the queue
			MyStatisticalOfStockPriceGain.Filter(ThisStockPricePredictionData.GainLog)
			MyQueue.Enqueue(ThisStockPricePredictionData)
			'Estimate the Mu and Sigma of the stock price gain
			'With MyStatisticalOfStockPriceGain.FilterLast
			'	ThisLogNormalMu = .LogNormalMu
			'	ThisLogNormalSigma = .LogNormalSigma
			'End With
			Dim ThisValueHigh As Double
			Dim ThisValueLow As Double
			Dim ThisValueExpHigh As Double
			Dim ThisValueExpLow As Double

			Dim ThisVolatilityDeltaHigh As Double
			Dim ThisVolatilityDeltaLow As Double
			' Estimate the error in the probability of threshold excess
			'leave the error to zero until ther is enough data to estimate it accuratly
			If MyQueue.Count = MyVolatilityMeasurementPeriod Then
				'try a linear estimator for threshold estimated volatility
				Dim ThisProbToBeLessThanExcessHigh As Double = 1 - MyProbabilityOfThresholdExcessHigh
				Dim ThisProbToBeLessThanExcessLow As Double = MyProbabilityOfThresholdExcessLow
				With MyStatisticalOfStockPriceGain.FilterLast
					'the inverselognormal is use here to estimate the stock price gain at the high and low probability of excess
					'and compare with the measured stock price volatility base on the daily log gain
					ThisValueHigh = Measure.InverseLogNormal(ThisProbToBeLessThanExcessHigh, .Mean, .StandardDeviation)
					ThisValueLow = Measure.InverseLogNormal(ThisProbToBeLessThanExcessLow, .Mean, .StandardDeviation)
					'Note the use of the exponential here rather than the inverse InverseLogNormal
					'as the exponential is 7 time faster and it give the same result as the InverseLogNormal above
					'this is possible because we already have the exact measure of the Mu and Sigma for the current data windows buffer
					'ThisValueExpHigh = Measure.InverseLogNormal(MyProbabilityHigh, .Mean, .StandardDeviation)
					'ThisValueExpLow = Measure.InverseLogNormal(MyProbabilityLow, .Mean, .StandardDeviation)
					ThisValueExpHigh = Math.Exp(.Mean + (MyProbabilityRatioToSigma * .StandardDeviation))
					ThisValueExpLow = Math.Exp(.Mean - (MyProbabilityRatioToSigma * .StandardDeviation))
				End With
				'take the average of the error of the high and low probabilty of excess and adjust the volatility of the lognormal to reflect
				'the measured excess probability
				'this give us a better estimate of the volatility and the risk assciated with the stock price estimate in the day.
				ThisVolatilityDeltaHigh = MathPlus.General.STATISTICAL_SIGMA_DAILY_TO_YEARLY_RATIO * (ThisValueExpHigh - ThisValueHigh)
				ThisVolatilityDeltaLow = MathPlus.General.STATISTICAL_SIGMA_DAILY_TO_YEARLY_RATIO * (ThisValueLow - ThisValueExpLow)

				'second order filter to quickly bring the volatility to the right level	as far a excess stock price probability is concerned
				'this fast tracking filter help the smooth the variation in the measured probability of excess 
				'MyVolatilityDelta = MyFilterPLL.FilterRun((ThisVolatilityDeltaHigh + ThisVolatilityDeltaLow) / 2)

				MyVolatilityDeltaHigh = MyFilterPLLHigh.FilterRun(ThisVolatilityDeltaHigh)
				MyVolatilityDeltaLow = MyFilterPLLLow.FilterRun(ThisVolatilityDeltaLow)
				MyVolatilityDelta = (MyVolatilityDeltaHigh + MyVolatilityDeltaLow) / 2
				MyVolatilitySumOfError = MyVolatilitySumOfError + MyVolatilityDelta
				With CalculateThresholdExcess(VolatilityDeltaLow:=MyVolatilityDeltaLow, VolatilityDeltaHigh:=MyVolatilityDeltaHigh)
					MyProbabilityOfThresholdExcess = .ProbabilityOfThresholdExcess
					MyProbabilityOfThresholdExcessHigh = .ProbabilityOfThresholdExcessHigh
					MyProbabilityOfThresholdExcessLow = .MyProbabilityOfThresholdExcessLow
				End With
				MyVolatilityEstimate = (MathPlus.General.STATISTICAL_SIGMA_DAILY_TO_YEARLY_RATIO * MyStatisticalOfStockPriceGain.FilterLast.StandardDeviation) + MyVolatilitySumOfError
				'Trace.WriteLine($"MyProbabilityOfThresholdExcess: {MyProbabilityOfThresholdExcess}, VolatilityDeltaLow: {MyVolatilityDeltaLow},VolatilityDeltaHigh: {MyVolatilityDeltaHigh}")
				'Test code for speed evaluation
				'Dim ThisStopWatch As New Stopwatch
				'ThisStopWatch.Start()
				'For I As Integer = 1 To 100
				'	CalculateThresholdExcess(0.0)
				'Next
				'ThisStopWatch.Stop()
				'Trace.WriteLine($"Time to calculate the threshold excess: {ThisStopWatch.ElapsedMilliseconds} ms")
				'Note MyProbabilityOfThresholdExcessHigh and MyProbabilityOfThresholdExcessLow are always >=1 and never zero
				ThisEstimateOfError = (MyProbabilityOfThresholdExcessHigh - MyProbabilityOfThresholdExcessLow) / (MyProbabilityOfThresholdExcessHigh + MyProbabilityOfThresholdExcessLow)
				MyStatisticalOfProbOfExcesDeltaHighLow.Filter(ThisEstimateOfError)
			Else
				MyVolatilitySumOfError = 0.0
				If Volatility > StockPriceVolatilityEstimateData.VOLATILITY_TOTAL_MINIMUM Then
					ThisResult = CalculateThresholdExcess(0.0, 0.0)
					With ThisResult
						MyProbabilityOfThresholdExcess = .ProbabilityOfThresholdExcess
						MyProbabilityOfThresholdExcessHigh = .ProbabilityOfThresholdExcessHigh
						MyProbabilityOfThresholdExcessLow = .MyProbabilityOfThresholdExcessLow
					End With
					MyVolatilityDeltaHigh = MyFilterPLLHigh.FilterRun(0.0)
					MyVolatilityDeltaLow = MyFilterPLLLow.FilterRun(0.0)
					MyVolatilityDelta = (MyVolatilityDeltaHigh + MyVolatilityDeltaLow) / 2
				Else
					MyProbabilityOfThresholdExcess = MyProbabilityOfInterval
					MyProbabilityOfThresholdExcessHigh = MyProbabilityOfInterval / 2
					MyProbabilityOfThresholdExcessLow = MyProbabilityOfInterval / 2
				End If
				MyStatisticalOfProbOfExcesDeltaHighLow.Filter(0.0)
				_StockPriceLast = StockPrice
				MyVolatilityEstimate = MathPlus.General.STATISTICAL_SIGMA_DAILY_TO_YEARLY_RATIO * MyStatisticalOfStockPriceGain.FilterLast.StandardDeviation
				'do not reduce the volatility below the measured volatility
			End If
			Return MyVolatilityEstimate
		End Function

		Public ReadOnly Property VolatilityEstimate As Double
			Get
				Return MyVolatilityEstimate
			End Get
		End Property

		Private Function CalculateThresholdExcess(VolatilityDeltaLow As Double, VolatilityDeltaHigh As Double) As (ProbabilityOfThresholdExcess As Double, ProbabilityOfThresholdExcessHigh As Double, MyProbabilityOfThresholdExcessLow As Double)
			Dim ThisProbabilityOfThresholdExcess As Double
			Dim ThisProbabilityOfThresholdExcessHigh As Double
			Dim ThisProbabilityOfThresholdExcessLow As Double
			Dim ThisSumOfThresholdExcess As Integer = 0
			Dim ThisSumOfThresholdExcessHigh As Integer = 0
			Dim ThisSumOfThresholdExcessLow As Integer = 0


			' Use a thread-safe collection to store the results
			Dim thresholdExcessResults As New ConcurrentBag(Of (IsBandExceeded As Boolean, IsBandExceededHigh As Boolean, IsBandExceededLow As Boolean))


			'Note the parallel execution is slower for typical buffer size and should not be used unless the buffer size is very large >>100
#If PARALLEL_EXECUTION Then
    ' Parallelize the loop using Parallel.ForEach
    Parallel.ForEach(MyQueue, Sub(ThisStockPricePredictionData)
                                  ThisStockPricePredictionData.Refresh(VolatilityDelta)
                                  thresholdExcessResults.Add((ThisStockPricePredictionData.IsBandExceeded, ThisStockPricePredictionData.IsBandExceededHigh, ThisStockPricePredictionData.IsBandExceededLow))
                              End Sub)
#Else
			' Synchronous execution
			For Each ThisStockPricePredictionData In MyQueue
				ThisStockPricePredictionData.Refresh(VolatilityDeltaLow:=VolatilityDeltaLow, VolatilityDeltaHigh:=VolatilityDeltaHigh)
				thresholdExcessResults.Add((ThisStockPricePredictionData.IsBandExceeded, ThisStockPricePredictionData.IsBandExceededHigh, ThisStockPricePredictionData.IsBandExceededLow))
			Next
#End If

			' Aggregate the results
			For Each result In thresholdExcessResults
				If result.IsBandExceeded Then
					ThisSumOfThresholdExcess += 1
					If result.IsBandExceededHigh Then
						ThisSumOfThresholdExcessHigh += 1
					End If
					If result.IsBandExceededLow Then
						ThisSumOfThresholdExcessLow += 1
					End If
				End If
			Next

			' Ensure non-zero values to avoid division by zero
			If ThisSumOfThresholdExcess = 0 Then ThisSumOfThresholdExcess = 1
			If ThisSumOfThresholdExcessHigh = 0 Then ThisSumOfThresholdExcessHigh = 1
			If ThisSumOfThresholdExcessLow = 0 Then ThisSumOfThresholdExcessLow = 1

			ThisProbabilityOfThresholdExcess = ThisSumOfThresholdExcess / MyVolatilityMeasurementPeriod
			ThisProbabilityOfThresholdExcessHigh = ThisSumOfThresholdExcessHigh / MyVolatilityMeasurementPeriod
			ThisProbabilityOfThresholdExcessLow = ThisSumOfThresholdExcessLow / MyVolatilityMeasurementPeriod

			Return (ThisProbabilityOfThresholdExcess, ThisProbabilityOfThresholdExcessHigh, ThisProbabilityOfThresholdExcessLow)
		End Function




		Public ReadOnly Property ProbabilityOfThresholdExcess As Double
			Get
				Return MyProbabilityOfThresholdExcess
			End Get
		End Property

		Public ReadOnly Property ProbabilityOfThresholdExcessHigh As Double
			Get
				Return MyProbabilityOfThresholdExcessHigh
			End Get
		End Property

		Public ReadOnly Property ProbabilityOfThresholdExcessLow As Double
			Get
				Return MyProbabilityOfThresholdExcessLow
			End Get
		End Property

		Public ReadOnly Property LogNormalEstimateOfMu As Double
			Get
				Return MyStatisticalOfStockPriceGain.FilterLast.Mean
			End Get
		End Property

		Public ReadOnly Property LogNormalEstimateOfSigma As Double
			Get
				Return MyStatisticalOfStockPriceGain.FilterLast.StandardDeviation
			End Get
		End Property

		Public ReadOnly Property StatisticEstimateOfError As IStatistical
			Get
				Return MyStatisticalOfProbOfExcesDeltaHighLow.FilterLast
			End Get
		End Property
	End Class
End Namespace
