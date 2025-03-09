
Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics


Namespace OptionValuation
	''' <summary>
	''' This class is a container for the stock price prediction run. 
	''' It contains the stock price prediction run data and the stock price prediction run results.	
	''' It os a simplified version of the StockPriceVolatilityPredictionBand class specialized for 
	''' the specific need to adjust teh stock price prediction level in function of the volatility.
	''' </summary>
	Public Class StockPricePredictionRun

		'Private _stockPricePredictionRunData As StockPricePredictionRunData

		Private MyProbabilityHigh As Double
		Private MyProbabilityLow As Double
		Private MyVolatilityDelta As Double
		Private MyVolatilityTotal As Double
		Private MyStockPriceHighValue As Double
		Private MyStockPriceLowValue As Double
		Private MyData As IStockPricePredictionData

		Public Sub New(Data As IStockPricePredictionData)
			MyData = Data
			MyProbabilityHigh = 0.5 + MyData.ProbabilityOfInterval / 2
			MyProbabilityLow = 0.5 - MyData.ProbabilityOfInterval / 2
			MyVolatilityDelta = 0.0
			MyVolatilityTotal = MyData.Volatility
			_IsBandExceeded = False
			_IsBandExceededHigh = False
			_IsBandExceededLow = False
		End Sub

		Public Function Refresh() As Boolean
			Return Me.Refresh(VolatilityDelta:=0.0)
		End Function

		Public Function Refresh(VolatilityDelta As Double) As Boolean

		End Function
	End Class
End Namespace
