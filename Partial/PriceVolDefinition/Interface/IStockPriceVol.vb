Public Interface IStockPriceVol
	ReadOnly Property AsIStockPrice As IStockPriceVol
	Property DateDay As Date
	Property Open As Double
	Property OpenNext As Double
	Property Last As Double
	Property LastPrevious As Double
	Property High As Double
	Property Low As Double
	Property Volume As Long

	''' <summary>
	''' Just a placeholder for the volume of the previous trading day.
	''' 'this volume can be nul if there is no trading on the previous day before the current one
	''' </summary>
	Property VolumePrevious As Long

	''' <summary>
	''' VolumePreviousTrading is the volume of the previous trading day.
	''' The value may still be zero if the stock did not trade yet but as soon as trading occur it will never be zero
	''' This value can safely be used to calculate the volume logarithmic change compared to the previous trading day
	''' this Value need to be set externally while the list of StockPriceVol is processed, it is not automatically calculated by the class itself 
	''' because it can be used in different ways and the logic to set this value can be different based on the best use case
	''' </summary>
	Property VolumePreviousTrading As Long


	'''' <summary>
	'''' Just a placeholder for the 20-day average volume.
	'''' </summary>
	'Property VolumeAverage20 As Long

	''' <summary>
	''' Just a placeholder for the 60-day average volume.
	''' </summary>
	Property VolumeAverage60 As Double

	'''' <summary>
	'''' Just a placeholder for the 100-day average volume.
	'''' </summary>
	'Property VolumeAverage100 As Long

	'''' <summary>
	'''' Just a placeholder for the 60-day relative volume.
	'''' (Volume/AvgVol60)
	'''' </summary>
	'Property VolumeRelatif60 As Double

	'''' <summary>
	'''' Just a placeholder for the 60-day relative volume in logarithmic scale.
	'''' V=ln(Volume/AvgVol60)
	'''' </summary>
	'Property VolumeRelatifLn60 As Double

	''' <summary>
	''' Liquidity index or Dollar Volume(DV) Is the total actual monetary value Of a security traded during a specific period. 
	''' Calculated As Shares Traded × Share Price, it reveals how much capital Is flowing through an asset, 
	''' helping measure liquidity without being misled by share price.
	''' DVolume = Volume × Current Price
	''' What Institutions Often Use
	''' Instead of raw volume
	''' DV=Price×Volume
	''' Or ln(DV)
	''' This measures : 
	''' Amount of capital exchanged.
	'''	This Is much more meaningful.
	''' </summary>
	Function DV() As Double

	Property DVAverage60 As Double

	' --------------------------------------------------------------------
	' Liquidity Deviation Volume Index or LDV
	'
	' LDV(i) = ln( DV(i) / AvgDV60(i) )
	' or:
	' LDV(i)    = ln( DV(i) / AvgDV60(i) )
	' AvgLDV    = Sum( LDV(i) ) / N
	' Liquidity = Exp( AvgLDV )

	' where:
	'   DV(i)      = Price(i) * Volume(i)
	'   AvgDV60(i)= 60-day average Dollar Volume

	' Interpretation:
	'It mirror what is done for the price return:
	' Return(i) = ln( Price(i) / Price(i-1) )
	' AvgReturn = Sum( Return(i) ) / N
	' Growth    = Exp( AvgReturn )
	'
	'
	' Interpretation:
	'   LDV equal to 0  -> Normal liquidity
	'   LDV greater than 0  -> Above-average participation
	'   LDV less than 0  -> Below-average participation
	'
	' Example:
	'   DV = 2 * AvgDV60
	'   LDV = ln(2) = 0.693
	' --------------------------------------------------------------------
	Function LDV() As Double
End Interface

