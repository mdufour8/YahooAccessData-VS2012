
Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter


Namespace OptionValuation
	''' <summary>
	''' This class is a container for the stock price prediction run. 
	''' It contains the stock price prediction run data and the stock price prediction run results.	
	''' It os a simplified version of the StockPriceVolatilityPredictionBand class specialized for 
	''' the specific need to adjust teh stock price prediction level in function of the volatility.
	''' </summary>
	Public Class StockPriceVolatilityEstimateRun

		Private Const FILTER_RATE_FOR_GAIN As Integer = 5

		Private _stockPricePredictionData As StockPriceVolatilityEstimateData

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
		Private MyFilterForProbability As FilterExp
		Private MyStatisticalOfProbOfExcesDeltaHighLow As FilterStatisticalExp
		Private MyFilterBrownExpPredict As FilterExpPredict
		Private MyFilterBrownExpPredictDerivative As FilterExpPredict
		Private MyFilterVolatilityYZ As FilterVolatilityYangZhang

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
			MyVolatilityPredictionBandType = VolatilityPredictionBandType
			MyVolatilityMeasurementPeriod = VolatilityMeasurementPeriodInDays
			MyFilterForProbability = New FilterExp(FilterRate:=VolatilityMeasurementPeriodInDays)
			MyQueue = New Queue(Of StockPriceVolatilityEstimateData)(capacity:=VolatilityMeasurementPeriodInDays)
			MyStatisticalOfProbOfExcesDeltaHighLow = New FilterStatisticalExp(FilterRate:=VolatilityMeasurementPeriodInDays)
			'this Is base on a	Brown exponential filter to calculate the gain And the derivative of the stock price
			'this Is use when the user does not provide a value for the Gain And its derivative
			MyFilterBrownExpPredict = New FilterExpPredict(FilterRate:=FILTER_RATE_FOR_GAIN) With {.Tag = "StockPriceVolatilityEstimateRun_BrownExpPredict"}
			MyFilterBrownExpPredictDerivative = New FilterExpPredict(FilterRate:=FILTER_RATE_FOR_GAIN) With {.Tag = "StockPriceVolatilityEstimateRun_BrownExpPredictDerivarive"}
			MyFilterVolatilityYZ = New FilterVolatilityYangZhang(FilterRate:=VolatilityMeasurementPeriodInDays, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Standard)
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
			MyFilterVolatilityYZ.Filter(StockPrice)
			Return Me.Add(StockPrice, MyFilterVolatilityYZ.FilterLast)
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
				ThisGainYearlyLog = MathPlus.General.NUMBER_TRADINGDAY_PER_YEAR * ThisGainDailyLog
			End With
			With MyFilterBrownExpPredictDerivative
				.FilterRun(ThisGainYearlyLog)
				ThisGainDailyLogDerivative = .GainLog
				ThisGainYearlyLogDerivative = MathPlus.General.NUMBER_TRADINGDAY_PER_YEAR * ThisGainDailyLogDerivative
			End With
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

			If (StockPrice.Vol = 0) OrElse (Volatility = 0) Then
				Return MyProbabilityOfThresholdExcess
			End If
			ThisQueueDataLast = MyQueue.Last

			If MyQueue.Count > 0 Then
				ThisQueueDataLast = MyQueue.Last
				'it is important to make sure the previous last value is updated so that the log normal GainLog property parameters
				'in StockPriceVolatilityEstimateData is calculated correctly
				StockPrice.LastPrevious = ThisQueueDataLast.StockPrice.Last
			Else
				StockPrice.LastPrevious = StockPrice.Last
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
			ThisStockPricePredictionData.Refresh()

			' If there are items in the queue, update the last item with the current stock price
			If MyQueue.Count > 0 Then
				ThisQueueDataLastDate = ThisQueueDataLast.StockPrice.DateDay
				' The last data in the queue does not have the current data yet, so we need to update it.
				' For the last data in the queue, this new data is updated via the StockPriceNext property.
				With ThisQueueDataLast
					'remove the previous calculation since it was an intraday value 
					'and the result may change over one days
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess - 1
						If .IsBandExceededHigh Then
							MySumOfThresholdExcessHigh = MySumOfThresholdExcessHigh - 1
						End If
						If .IsBandExceededLow Then
							MySumOfThresholdExcessLow = MySumOfThresholdExcessLow - 1
						End If
					End If
					.StockPriceNext = StockPrice
					.Refresh()
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess + 1
						If .IsBandExceededHigh Then
							MySumOfThresholdExcessHigh = MySumOfThresholdExcessHigh + 1
						End If
						If .IsBandExceededLow Then
							MySumOfThresholdExcessLow = MySumOfThresholdExcessLow + 1
						End If
					End If
				End With
			End If
			'Update the queue by removing the oldest item if the queue has reached its maximum size
			If MyQueue.Count = MyVolatilityMeasurementPeriod Then
				ThisQueueDataRemoved = MyQueue.Dequeue
				With ThisQueueDataRemoved
					'adjust the sum for the removed item
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess - 1
						If .IsBandExceededHigh Then
							MySumOfThresholdExcessHigh = MySumOfThresholdExcessHigh - 1
						End If
						If .IsBandExceededLow Then
							MySumOfThresholdExcessLow = MySumOfThresholdExcessLow - 1
						End If
					End If
					ThisQueueDataRemovedDate = .StockPrice.DateDay
				End With
			End If
			' Add the new data to the queue
			MyQueue.Enqueue(ThisStockPricePredictionData)
			With ThisStockPricePredictionData
				If .IsBandExceeded Then
					MySumOfThresholdExcess = MySumOfThresholdExcess + 1
					MyFilterForProbability.FilterRun(1.0)
					If .IsBandExceededHigh Then
						MySumOfThresholdExcessHigh = MySumOfThresholdExcessHigh + 1
					End If
					If .IsBandExceededLow Then
						MySumOfThresholdExcessLow = MySumOfThresholdExcessLow + 1
					End If
				Else
					MyFilterForProbability.FilterRun(0.0)
				End If
			End With
			MyProbabilityOfThresholdExcess = MySumOfThresholdExcess / MyVolatilityMeasurementPeriod
			MyProbabilityOfThresholdExcessHigh = MySumOfThresholdExcessHigh / MyVolatilityMeasurementPeriod
			MyProbabilityOfThresholdExcessLow = MySumOfThresholdExcessLow / MyVolatilityMeasurementPeriod
			'Measure.InverseLogNormal()

			' Estimate the error in the probability of threshold excess
			'leave the error to zero until ther is enough data to estimate it accuratly
			If MyQueue.Count = MyVolatilityMeasurementPeriod Then
				ThisEstimateOfError = (MyProbabilityOfThresholdExcessHigh - MyProbabilityOfThresholdExcessLow) / (MyProbabilityOfThresholdExcessHigh + MyProbabilityOfThresholdExcessLow)
				MyStatisticalOfProbOfExcesDeltaHighLow.Filter(ThisEstimateOfError)
			Else
				MyStatisticalOfProbOfExcesDeltaHighLow.Filter(0.0)
			End If
			Return MyProbabilityOfThresholdExcess
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

		Public ReadOnly Property StatisticEstimateOfError As IStatistical
			Get
				Return MyStatisticalOfProbOfExcesDeltaHighLow.FilterLast
			End Get
		End Property
	End Class
End Namespace
