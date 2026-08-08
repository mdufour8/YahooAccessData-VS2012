Imports WebEODData
Imports YahooAccessData.ExtensionService.Extensions
Public Class StockPriceVol
	Implements IEquatable(Of StockPriceVol)
	Implements IStockPriceVol
	Implements IStockPriceAdjusted
	Implements IPriceVol

	Public Enum StockPriceDataType
		RawPrice = 0
		CumulativeLogReturn = 1
		CumulativeLogYearlyReturn = 2
		LogReturn = 3
	End Enum

	Private _dataType As StockPriceDataType = StockPriceDataType.RawPrice

	Public Sub New()
		Me.DateDay = Now.Date
	End Sub

	Public Sub New(scalarValue As Double)
		Me.New()
		Me.Open = scalarValue
		Me.High = scalarValue
		Me.Low = scalarValue
		Me.Last = scalarValue
		Me.OpenNext = scalarValue
		Me.LastPrevious = scalarValue
		Me.Volume = 0
		_Ratio = 1.0
		_PriceDelta = 0.0
		_IsPriceAdjustedEnabled = False
	End Sub

	Public Sub New(PriceVol As PriceVol)
		If PriceVol Is Nothing Then
			Throw New ArgumentNullException(NameOf(PriceVol))
		End If
		Me.DataType = PriceVol.DataType
		Me.DateDay = PriceVol.DateLastTrade
		Me.Open = PriceVol.Open
		Me.OpenNext = PriceVol.OpenNext
		Me.Last = PriceVol.Last
		Me.LastPrevious = PriceVol.LastPrevious
		Me.High = PriceVol.High
		Me.Low = PriceVol.Low
		Me.Volume = PriceVol.Volume
		Volume = Volume
		VolumePrevious = Volume
		VolumePreviousTrading = Volume
		Me.AsStockPriceAdjusted.SetPriceAdjusted(PriceVol.AsStockPriceAdjusted)
	End Sub

	''' <summary>
	''' Copy constructor: creates a new StockPriceVol with the same values as the given one.
	''' </summary>
	''' <param name="StockPriceVol"></param>
	Public Sub New(StockPriceVol As StockPriceVol)
		If StockPriceVol Is Nothing Then
			Throw New ArgumentNullException(NameOf(StockPriceVol))
		End If
		Me.DataType = StockPriceVol.DataType
		Me.DateDay = StockPriceVol.DateDay
		Me.Open = StockPriceVol.Open
		Me.OpenNext = StockPriceVol.OpenNext
		Me.Last = StockPriceVol.Last
		Me.LastPrevious = StockPriceVol.LastPrevious
		Me.High = StockPriceVol.High
		Me.Low = StockPriceVol.Low
		Me.Volume = StockPriceVol.Volume
		VolumePrevious = StockPriceVol.VolumePrevious
		VolumePreviousTrading = StockPriceVol.VolumePreviousTrading
		IStockPriceAdjusted_SetPriceAdjusted(StockPriceVol.AsStockPriceAdjusted)
	End Sub

	Public Property DataType As StockPriceDataType
		Get
			Return _dataType
		End Get
		Set(value As StockPriceDataType)
			_dataType = value
		End Set
	End Property

	Private _priceFirst As Double

	''' <summary>
	''' This data can be usuful to return from a CumulativeLog value to the original price value, 
	''' for example when the data type is set to CumulativeLogReturn or CumulativeLogYearlyReturn
	''' </summary>
	''' <returns></returns>
	Public Property PriceFirst As Double
		Get
			Return _priceFirst
		End Get
		Set(value As Double)
			_priceFirst = value
		End Set
	End Property

	Public Property DateDay As Date Implements IStockPriceVol.DateDay
	Public Property Open As Double Implements IStockPriceVol.Open
	Public Property OpenNext As Double Implements IStockPriceVol.OpenNext
	Public Property Last As Double Implements IStockPriceVol.Last
	Public Property LastPrevious As Double Implements IStockPriceVol.LastPrevious
	Public Property High As Double Implements IStockPriceVol.High
	Public Property Low As Double Implements IStockPriceVol.Low
	Public Property Volume As Long Implements IStockPriceVol.Volume

	Public Overrides Function ToString() As String
		Return $"{TypeName(Me)},Date:{Me.DateDay},LastPrevious:{Me.LastPrevious:F3},Open:{Me.Open:F3},High:{Me.High:F3},Low:{Me.Low:F3},Last:{Me.Last:F3},OpenNext:{Me.OpenNext:F3},Volume:{Me.Volume}"
	End Function

#Region "Equality Test"
	' Custom method with configurable precision

	Public Overloads Function Equals(other As StockPriceVol) As Boolean Implements IEquatable(Of StockPriceVol).Equals
		If other Is Nothing Then Return False
		Return Me.DateDay = other.DateDay AndAlso
					 Me.Open = other.Open AndAlso
					 Me.High = other.High AndAlso
					 Me.Low = other.Low AndAlso
					 Me.Last = other.Last AndAlso
					 Me.OpenNext = other.OpenNext AndAlso
					 Me.LastPrevious = other.LastPrevious AndAlso
					 Me.Volume = other.Volume
	End Function

	Public Overloads Function Equals(other As StockPriceVol, decimalPlaces As Integer) As Boolean
		If other Is Nothing Then Return False
		Return Me.DateDay = other.DateDay AndAlso
						 Math.Round(Me.Open, decimalPlaces) = Math.Round(other.Open, decimalPlaces) AndAlso
						 Math.Round(Me.High, decimalPlaces) = Math.Round(other.High, decimalPlaces) AndAlso
						 Math.Round(Me.Low, decimalPlaces) = Math.Round(other.Low, decimalPlaces) AndAlso
						 Math.Round(Me.Last, decimalPlaces) = Math.Round(other.Last, decimalPlaces) AndAlso
						 Math.Round(Me.OpenNext, decimalPlaces) = Math.Round(other.OpenNext, decimalPlaces) AndAlso
						 Math.Round(Me.LastPrevious, decimalPlaces) = Math.Round(other.LastPrevious, decimalPlaces) AndAlso
						 Me.Volume = other.Volume
	End Function

	Public Overrides Function Equals(obj As Object) As Boolean
		If obj Is Nothing OrElse Not Me.GetType() Is obj.GetType() Then
			Return False
		End If
		Dim other As StockPriceVol = DirectCast(obj, StockPriceVol)
		Return Me.DateDay = other.DateDay AndAlso
					 Me.Open = other.Open AndAlso
					 Me.High = other.High AndAlso
					 Me.Low = other.Low AndAlso
					 Me.Last = other.Last AndAlso
					 Me.OpenNext = other.OpenNext AndAlso
					 Me.LastPrevious = other.LastPrevious AndAlso
					 Me.Volume = other.Volume
		' 👉 You can add more fields here if you need deeper checks
	End Function

	Public Overrides Function GetHashCode() As Integer
		Dim hash As Integer = 17
		hash = hash * 31 + DateDay.GetHashCode()
		hash = hash * 31 + Open.GetHashCode()
		hash = hash * 31 + High.GetHashCode()
		hash = hash * 31 + Low.GetHashCode()
		hash = hash * 31 + Last.GetHashCode()
		hash = hash * 31 + OpenNext.GetHashCode()
		hash = hash * 31 + LastPrevious.GetHashCode()
		hash = hash * 31 + Volume.GetHashCode()
		Return hash
	End Function

	' Quick helper: check if two references point to the *same* object
	Public Overloads Shared Function ReferenceEquals(pv1 As StockPriceVol, pv2 As StockPriceVol) As Boolean
		Return Object.ReferenceEquals(pv1, pv2)
	End Function
#End Region

	Public ReadOnly Property AsIStockPrice As IStockPriceVol Implements IStockPriceVol.AsIStockPrice
		Get
			Return Me
		End Get
	End Property


#Region "IPriceVol Implementation"
	Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
		Get
			Return Me
		End Get
	End Property

	Private Property IPriceVol_DateUpdate As Date Implements IPriceVol.DateUpdate
		Get
			Return Me.DateDay
		End Get
		Set(value As Date)
			Me.DateDay = value
		End Set
	End Property

	Private Property IPriceVol_LastWeighted As Single Implements IPriceVol.LastWeighted
		Get
			Return CSng(Me.Last)
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_LastAdjusted As Single Implements IPriceVol.LastAdjusted
		Get
			Return CSng(Last)
		End Get
		Set(value As Single)
			Last = value
		End Set
	End Property

	Private Property IPriceVol_VolPlus As Integer Implements IPriceVol.VolPlus
		Get
			If Me.Volume > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Me.Volume < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Me.Volume)
			End If
		End Get
		Set(value As Integer)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_VolMinus As Integer Implements IPriceVol.VolMinus
		Get
			If Me.Volume > Integer.MaxValue Then
				Return Integer.MaxValue
			ElseIf Me.Volume < Integer.MinValue Then
				Return Integer.MinValue
			Else
				Return CInt(Me.Volume)
			End If
		End Get
		Set(value As Integer)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
		Get
			Return False ' Assuming this is not an intra-day price
		End Get
		Set(value As Boolean)
			'ignore the setter for now
		End Set
	End Property

	Private Property IPriceVol_Range As Single Implements IPriceVol.Range
		Get
			Return CSng(High - Low)
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
		Get
			Return False
		End Get
		Set(value As Boolean)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue
		Get
			Return 0
		End Get
		Set(value As Single)
			'ignore the setter
		End Set
	End Property

	Private Property IPriceVol_DateDay As Date Implements IPriceVol.DateDay
		Get
			Return DateDay
		End Get
		Set(value As Date)
			DateDay = value
		End Set
	End Property

	Private Property IPriceVol_Open As Single Implements IPriceVol.Open
		Get
			Return CSng(Open)
		End Get
		Set(value As Single)
			Open = value
		End Set
	End Property

	Private Property IPriceVol_OpenNext As Single Implements IPriceVol.OpenNext
		Get
			Return CSng(OpenNext)
		End Get
		Set(value As Single)
			OpenNext = value
		End Set
	End Property

	Private Property IPriceVol_Last As Single Implements IPriceVol.Last
		Get
			Return CSng(Last)
		End Get
		Set(value As Single)
			Last = value
		End Set
	End Property

	Private Property IPriceVol_LastPrevious As Single Implements IPriceVol.LastPrevious
		Get
			Return CSng(LastPrevious)
		End Get
		Set(value As Single)
			LastPrevious = value
		End Set
	End Property

	Private Property IPriceVol_High As Single Implements IPriceVol.High
		Get
			Return CSng(High)
		End Get
		Set(value As Single)
			High = value
		End Set
	End Property

	Private Property IPriceVol_Low As Single Implements IPriceVol.Low
		Get
			Return CSng(Low)
		End Get
		Set(value As Single)
			Low = value
		End Set
	End Property

	Private Property IPriceVol_Vol As Integer Implements IPriceVol.Vol
		Get
			Return Me.Volume.ToIntegerSafe
		End Get
		Set(value As Integer)
			Me.Volume = value
		End Set
	End Property
#End Region
#Region "Factory Methods"
	''' <summary>
	''' Creates a new StockPriceVol from the local instance.
	''' </summary>
	''' <returns>Return a copy of the current class instance.</returns>
	Public Function CopyFrom() As StockPriceVol
		Return New StockPriceVol(Me)
	End Function

	Private _IsIntraDay As Boolean
	''' <summary>
	''' Indicate that the price vol data is an intraday data update and the day has not yet finished trading
	''' Mainly use at lower level to indicate that this is an intraday price quotation
	''' </summary>
	''' <returns></returns>
	Public Property IsIntraDay As Boolean
		Get
			Return _IsIntraDay
		End Get
		Set(value As Boolean)
			_IsIntraDay = value
		End Set
	End Property

	Private _VolumePrevious As Long
	Public Property VolumePrevious As Long Implements IStockPriceVol.VolumePrevious
		Get
			Return _VolumePrevious
		End Get
		Set(value As Long)
			_VolumePrevious = value
		End Set
	End Property

	Private _VolumePreviousTrading As Long

	''' <summary>
	''' VolumePreviousTrading is the volume of the previous trading day.
	''' The value is never zero and if the stock did not trade yet its value is set back to the volume of the first trading day.
	''' This value can safely be used to calculate the volume logarithmic change compared to the previous trading day
	''' this Value need to be set extrenally while the list of StockPriceVol is processed, it is not automatically calculated by the class itself 
	''' because it can be used in different ways and the logic to set this value can be different based on the best use case
	''' </summary>
	Public Property VolumePreviousTrading As Long Implements IStockPriceVol.VolumePreviousTrading
		Get
			Return _VolumePreviousTrading
		End Get
		Set(value As Long)
			_VolumePreviousTrading = value
		End Set
	End Property


	''' <summary>
	''' Liquidity index or Dollar Volume(DV) Is the total actual monetary value Of a security traded during a specific period. 
	''' Calculated As the naturel log of the Volume × Share Price, it reveals how much capital Is flowing through an asset, 
	''' helping measure liquidity without being misled by share price.
	''' DV=(Volume × Current Price) 
	''' It is reasonable to represent the DV as a natural log because the DV can vary 
	''' over several orders of magnitude and the nature of the price process is theoritically 
	''' represented a a lognormal distribution. The log of the DV therefore more likely 
	''' to be normally distributed and can be used in statistical analysis which is 
	''' What iInstitutions Often Use instead of raw volume
	''' DV=Price×Volume
	''' Or ln(DV) see LDV for more details
	''' This measures : 
	''' Amount of capital exchanged on a logarithmic .
	'''	This Is much more meaningful in term of significance than teh DV itself.
	''' </summary>
	Private Function IStockPriceVol_DV() As Double Implements IStockPriceVol.DV
		Return Me.Last * Me.Volume
	End Function

	''' <summary>
	''' --------------------------------------------------------------------
	''' Liquidity Deviation Volume Index or LDV
	'''
	''' LDV(i) = ln( DV(i) / AvgDV60(i) )
	''' or:
	''' LDV(i)    = ln( DV(i) / AvgDV60(i) )
	''' AvgLDV    = Sum( LDV(i) ) / N
	''' Liquidity = Exp( AvgLDV )
	''' 
	''' Interpretation:
	'''	It mirror what is done for the price return:
	''' Return(i) = ln( Price(i) / Price(i-1) )
	''' AvgReturn = Sum( Return(i) ) / N
	''' Growth    = Exp( AvgReturn )
	''' Her it return the LDV based on the DVReference passed as parameter
	''' LDV(i)    = ln(DV(i) / DVReference)
	''' Generally the DVReference is the average DV over a certain period, for example 60 days, 
	''' but it can be any other reference value including the DV of the previous day or the DV of the previous trading day, etc.
	''' 
	''' where:
	'''   DV(i)      = Price(i) * Volume(i)
	'''   AvgDV60(i)= 60-day average Dollar Volume
	''' Interpretation:
	'''It mirror what is done for the price return:
	''' Return(i) = ln( Price(i) / Price(i-1) )
	''' AvgReturn = Sum( Return(i) ) / N
	''' Growth    = Exp( AvgReturn )
	'''
	''' Interpretation:
	'''   LDV equal to 0  -> Normal liquidity
	'''   LDV greater than 0  -> Above-average participation
	'''   LDV less than 0  -> Below-average participation
	'''
	''' Example:
	'''   DV = 2 * AvgDV60
	'''   LDV = ln(2) = 0.693
	''' --------------------------------------------------------------------
	''' </summary>
	Public Function LDV() As Double Implements IStockPriceVol.LDV
		Dim ThisDV As Double = Me.Last * Me.Volume
		Return If(ThisDV > 0, Math.Log(ThisDV), 0)
	End Function

	Public Function LDV(DVReference As Double) As Double Implements IStockPriceVol.LDV
		Dim ThisDV As Double = Me.Last * Me.Volume
		Return If(ThisDV > 0 AndAlso DVReference > 0, Math.Log(ThisDV / DVReference), 0)
	End Function
#End Region
#Region "IStockPriceAdjusted"
	Function AsStockPriceAdjusted() As IStockPriceAdjusted Implements IStockPriceAdjusted.AsStockPriceAdjusted
		Return Me
	End Function

	Private _Ratio As Double
	Private ReadOnly Property IStockPriceAdjusted_Ratio As Double Implements IStockPriceAdjusted.Ratio
		Get
			Return _Ratio
		End Get
	End Property

	Private _PriceDelta As Double
	Private ReadOnly Property IStockPriceAdjusted_PriceDelta As Double Implements IStockPriceAdjusted.PriceDelta
		Get
			Return _PriceDelta
		End Get
	End Property

	Public Function IStockPriceAdjusted_PriceDeltaPerCent() As Double Implements IStockPriceAdjusted.PriceDeltaPerCent
		Return If(Me.Last > 0, 100 * (_PriceDelta / Me.Last), 0)
	End Function

	Public Sub IStockPriceAdjusted_SetPriceAdjusted(value As IStockPriceAdjusted) Implements IStockPriceAdjusted.SetPriceAdjusted
		_PriceDelta = value.PriceDelta
		_Ratio = value.Ratio
		_IsPriceAdjustedEnabled = value.IsEnabled
	End Sub

	Private _IsPriceAdjustedEnabled As Boolean
	Public ReadOnly Property IsEnabled As Boolean Implements IStockPriceAdjusted.IsEnabled
		Get
			Return _IsPriceAdjustedEnabled
		End Get
	End Property

	Public Sub SetPriceAdjusted(Enable As Boolean) Implements IStockPriceAdjusted.SetPriceAdjusted
		Throw New NotImplementedException
		'_IsPriceAdjustedEnabled = Enable
	End Sub
#End Region  '"IStockPriceAdjusted"
End Class

