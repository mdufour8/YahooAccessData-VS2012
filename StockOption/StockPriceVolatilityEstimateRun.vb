
Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics


Namespace OptionValuation
	''' <summary>
	''' This class is a container for the stock price prediction run. 
	''' It contains the stock price prediction run data and the stock price prediction run results.	
	''' It os a simplified version of the StockPriceVolatilityPredictionBand class specialized for 
	''' the specific need to adjust teh stock price prediction level in function of the volatility.
	''' </summary>
	Public Class StockPriceVolatilityEstimateRun

		Private _stockPricePredictionData As StockPriceVolatilityEstimateData

		Private MyNumberTradingDays As Double
		Private MyVolatilityMeasurementPeriod As Integer
		Private MyProbabilityOfInterval As Double
		Private MyVolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
		Private MyQueue As Queue(Of StockPriceVolatilityEstimateData)
		Private MySumOfThresholdExcess As Integer
		Private MyProbabilityOfThresholdExcess As Double
		Private MyFilterForProbability As FilterExp
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
			MyVolatilityPredictionBandType = VolatilityPredictionBandType
			MyVolatilityMeasurementPeriod = VolatilityMeasurementPeriodInDays
			MyFilterForProbability = New FilterExp(FilterRate:=VolatilityMeasurementPeriodInDays)
			MyQueue = New Queue(Of StockPriceVolatilityEstimateData)(capacity:=VolatilityMeasurementPeriodInDays)
		End Sub

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

			If (StockPrice.Vol = 0) OrElse (Volatility = 0) Then
				Return MyProbabilityOfThresholdExcess
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
				ThisQueueDataLast = MyQueue.Last
				ThisQueueDataLastDate = ThisQueueDataLast.StockPrice.DateDay
				' The last data in the queue does not have the current data yet, so we need to update it.
				' For the last data in the queue, this new data is updated via the StockPriceNext property.
				With ThisQueueDataLast
					'remove the previous calculation since it was an intraday value 
					'and the result may change over one days
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess - 1
					End If
					.StockPriceNext = StockPrice
					.Refresh()
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess + 1
					End If
				End With
			End If
			' Update the queue by removing the oldest item if the queue has reached its maximum size
			If MyQueue.Count = MyVolatilityMeasurementPeriod Then
				ThisQueueDataRemoved = MyQueue.Dequeue
				With ThisQueueDataRemoved
					'adjust the sum for the removed item
					If .IsBandExceeded Then
						MySumOfThresholdExcess = MySumOfThresholdExcess - 1
					End If
					ThisQueueDataRemovedDate = .StockPrice.DateDay
				End With
			End If
			' Add the new data to the queue
			MyQueue.Enqueue(ThisStockPricePredictionData)
			If ThisStockPricePredictionData.IsBandExceeded Then
				MySumOfThresholdExcess = MySumOfThresholdExcess + 1
				MyFilterForProbability.FilterRun(1.0)
			Else
				MyFilterForProbability.FilterRun(0.0)
			End If
			MyProbabilityOfThresholdExcess = MySumOfThresholdExcess / MyVolatilityMeasurementPeriod
			'MyProbabilityOfThresholdExcess = MyFilterForProbability.FilterLast
			Return MyProbabilityOfThresholdExcess
		End Function
	End Class
End Namespace
