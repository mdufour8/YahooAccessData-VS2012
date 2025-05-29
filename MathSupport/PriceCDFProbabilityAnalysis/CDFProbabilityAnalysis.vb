Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Filter

Namespace MathPlus.Filter
	''' <summary>
	''' Module providing functions to compute the probability that a stock's price
	''' movement from a given pivot point in the past to the present is consistent
	''' with the current trend and volatility, based on a cumulative distribution
	''' function (CDF) model.
	'''
	''' This implementation allows evaluation from multiple pivot offsets, enabling
	''' analysis of how different historical points influence perceived trend strength
	''' or reversal likelihood. This is useful for statistical trading models,
	''' signal filtering, and adaptive analysis strategies.
	'''
	''' The CDF probability reflects the likelihood that the stock price will be below
	''' its current value after a given interval, assuming a normal distribution of returns.
	''' </summary>
	Public Module CDFProbability

		''' <summary>
		''' Calculates the probability (using a cumulative distribution function) that the observed
		''' price movement between a past pivot point and the current price is statistically significant,
		''' under the assumption of zero expected gain (flat trend).
		''' 
		''' This function is used as a trend detection tool. It estimates whether the recent price
		''' movement deviates meaningfully from a flat distribution, given the current volatility.
		''' 
		''' The returned probability approaches 0.5 when there is no significant trend. Values near
		''' 0 or 1 indicate strong deviations from flat behavior, suggesting a possible trend.
		''' 
		''' Important: The estimated trend is not used in this calculation to avoid circular reasoning.
		''' </summary>
		''' <param name="Filter">Filter providing past and current price samples.</param>
		''' <param name="Volatility">Filter providing the most recent volatility estimate.</param>
		''' <param name="PivotOffset">Number of samples ago to use as the pivot reference point.</param>
		''' <returns>A probability value between 0 and 1 indicating how likely the observed move is under a flat-trend model.</returns>
		Public Function CalculateSinglePivotCDFProbability(Filter As IFilterRun, Volatility As IFilterRun, PivotOffset As Integer) As Double
			If Volatility.FilterLast = 0 Then Return 0.5 ' Cannot infer trend if volatility is zero

			' Get pivot and current price values
			Dim pastPrice = Filter.FilterLast(Index:=PivotOffset)
			Dim currentPrice = Filter.FilterLast

			' Trend is not included in this probability calculation.
			' A zero-gain model assumes no expected upward or downward movement.
			Dim probability = StockOption.StockPricePredictionInverse(
				NumberTradingDays:=PivotOffset,
				StockPriceStart:=pastPrice,
				Gain:=0.0, ' Flat expectation (null hypothesis)
				GainDerivative:=0.0,
				Volatility:=Volatility.FilterLast,
				StockPriceEnd:=currentPrice)

			Return probability
		End Function


		''' <summary>
		''' Returns a dictionary mapping each pivot offset to its individual CDF probability.
		''' Useful for comparative analysis or heatmap visualization of pivot effectiveness.
		''' </summary>
		Public Function GetAllPivotProbabilities(Filter As IFilterRun, Volatility As IFilter, PivotOffsets As IEnumerable(Of Integer)) As Dictionary(Of Integer, Double)
			Dim result As New Dictionary(Of Integer, Double)
			For Each offset In PivotOffsets
				Try
					Dim p = CalculateSinglePivotCDFProbability(Filter, Volatility, offset)
					result(offset) = p
				Catch ex As ArgumentOutOfRangeException
					' Skip if the offset is beyond available data
				End Try
			Next
			Return result
		End Function

		Public Function CalculateAveragePriceCDFProbability(Filter As IFilterRun, Volatility As IFilter, PivotOffsets As IEnumerable(Of Integer)) As Double
			Dim result As Double

			If PivotOffsets.Count = 0 Then
				Return 0.5 ' No offsets provided, return neutral probability
			End If
			For Each offset In PivotOffsets
				Try
					Dim p = CalculateSinglePivotCDFProbability(Filter, Volatility, offset)
					result = result + p
				Catch ex As ArgumentOutOfRangeException
					' Skip if the offset is beyond available data
				End Try
			Next
			Return result / PivotOffsets.Count
		End Function
	End Module
End Namespace
