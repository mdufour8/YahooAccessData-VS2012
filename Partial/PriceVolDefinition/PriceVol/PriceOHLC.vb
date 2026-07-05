
''' <summary>
''' Class representing the Open, High, Low, and Close prices of a financial instrument for a specific time period.
''' Can be use to represnt a candlestick or OHLC bar in financial charts.
''' </summary>
Public Class PriceOHLC

	Public Sub New()

	End Sub

	''' <summary>
	''' Initializes a new instance of the PriceOHLC class using an object that implements the IStockPriceVol interface.
	''' </summary>
	''' <param name="PriceVol"></param>
	Public Sub New(PriceVol As IStockPriceVol)
		Me.Open = PriceVol.Open
		Me.High = PriceVol.High
		Me.Low = PriceVol.Low
		Me.Close = PriceVol.Last
	End Sub
	Public Sub New(open As Double, high As Double, low As Double, close As Double)
		Me.Open = open
		Me.High = high
		Me.Low = low
		Me.Close = close
	End Sub

	Property Open As Double
	Property High As Double
	Property Low As Double
	Property Close As Double
End Class
