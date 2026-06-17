Imports YahooAccessData.ExtensionService.Extensions
Public Class StockPriceVol
	Implements IEquatable(Of StockPriceVol)
	Implements IStockPriceVol

	Implements IPriceVol

	Public Enum StockPriceDataType
		RawPrice = 0
		CumulativeLogReturn = 1
		CumulativeLogYearlyReturn = 2
	End Enum

	Private _dataType As StockPriceDataType = StockPriceDataType.RawPrice

	Public Sub New()
		Me.DateDay = Now.Date
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
		With PriceVol.AsIStockPriceVol
			Me.VolumeAverage60 = .VolumeAverage60
			Me.VolumePrevious = .VolumePrevious
			Me.VolumePreviousTrading = .VolumePreviousTrading
			Me.DVAverage60 = .DVAverage60
		End With
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
		With StockPriceVol
			Me.VolumeAverage60 = .VolumeAverage60
			Me.VolumePrevious = .VolumePrevious
			Me.VolumePreviousTrading = .VolumePreviousTrading
			Me.DVAverage60 = .DVAverage60
		End With
	End Sub

	Public Property DataType As StockPriceDataType
		Get
			Return _dataType
		End Get
		Set(value As StockPriceDataType)
			_dataType = value
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


	''' <summary>
	''' Calculates the price * volume so that the volume become a weight for the price, 
	''' this can be used to calculate a weighted average price over a period of time or for other purposes where the price need to be weighted by the volume 
	''' for example to calculate a logaritmic weighted average price over a period of time or other pupose
	''' by summing the PriceVolume and dividing by the total volume of the period
	''' </summary>
	Public Function DV() As Double Implements IStockPriceVol.DV
		'try to return a positive value for the DV even if the volume is zero, this can be useful
		'to avoid division by zero error when calculating the logarithmic change of the DV
		'compared to its average over a period of time	
		If Volume > 0 Then
			Return Last * Volume
		Else
			Return Last * VolumePreviousTrading
		End If
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



	Private _VolumeAverage60 As Double

	''' <summary>
	''' Just a placeholder for the 60-day average volume often use for other calculation.
	''' The value need to be set extrenally while the list of StockPriceVol is processed
	''' </summary>
	Public Property VolumeAverage60 As Double Implements IStockPriceVol.VolumeAverage60
		Get
			Return _VolumeAverage60
		End Get
		Set(value As Double)
			_VolumeAverage60 = value
		End Set
	End Property


	Private _DVAverage60 As Double

	''' <summary>
	''' The dollar volume average over 60 days, this is the average of the price * volume over the last 60 days. 
	''' This can be used to calculate the liquidity deviation volume index (LDV) which is a measure of how much 
	''' the current dollar volume deviates from its average over the last 60 days.
	''' </summary>
	''' <returns></returns>
	Private Property DVAverage60 As Double Implements IStockPriceVol.DVAverage60
		Get
			Return _DVAverage60
		End Get
		Set(value As Double)
			_DVAverage60 = value
		End Set
	End Property

	Public Function LDV() As Double Implements IStockPriceVol.LDV
		Dim thisDV = DV()
		If thisDV > 0 AndAlso _DVAverage60 > 0 Then
			Return Math.Log(thisDV / _DVAverage60)
		Else
			Return 0
		End If
	End Function
#End Region
End Class

