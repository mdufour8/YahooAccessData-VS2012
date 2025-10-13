Public Interface IPriceVolGain
	Inherits IPriceVol

	ReadOnly Property PriceReference As Double?

	ReadOnly Property PriceOffset As Double?

	Sub SetPriceReference(PriceReference As Double, PriceOffset As Double)
End Interface
