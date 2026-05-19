Public Interface IVolatilityResult
	ReadOnly Property FullDay As Double
	ReadOnly Property PreviousCloseToOpen As Double
	ReadOnly Property OpenToClose As Double
	ReadOnly Property OpenToHighClose As Double
	ReadOnly Property OpenToLowClose As Double
	ReadOnly Property OpenToHighToLowCloseRatio As Double
	ReadOnly Property OpenToHighToLowCloseRatioFiltered As Double
	ReadOnly Property Rogers_Satchell_Yoon_Vrs As Double
	ReadOnly Property Parkison_Vp As Double
End Interface
