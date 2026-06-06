#Region "IPricePivotPoint"
''' <summary>
''' see definition:https://www.fidelity.com/learning-center/trading-investing/technical-analysis/technical-indicator-guide/pivot-points-resistance-support
''' </summary>
''' <remarks></remarks>
Public Interface IPricePivotPoint
	Enum enuPivotLevel
		Level1
		Level2
		Level3
	End Enum

	ReadOnly Property AsIPricePivotPoint As IPricePivotPoint
	ReadOnly Property PivotOpen As Single
	ReadOnly Property PivotLast As Single
	ReadOnly Property Resistance(ByVal Level As enuPivotLevel) As Single
	ReadOnly Property Support(ByVal Level As enuPivotLevel) As Single
	Function PriceVolPivot(ByVal Level As enuPivotLevel) As IPriceVol
End Interface
#End Region