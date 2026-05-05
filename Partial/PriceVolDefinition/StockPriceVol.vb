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


	Public Sub New(PriceVol As IPriceVol)
		Me.New(DirectCast(PriceVol, IStockPriceVol), StockPriceDataType.RawPrice)
	End Sub

	Public Sub New(PriceVol As IStockPriceVol, Optional DataType As StockPriceDataType = StockPriceDataType.RawPrice)
		_dataType = DataType

		Me.DateDay = PriceVol.DateDay
		Me.Open = PriceVol.Open
		Me.OpenNext = PriceVol.OpenNext
		Me.Last = PriceVol.Last
		Me.LastPrevious = PriceVol.LastPrevious
		Me.High = PriceVol.High
		Me.Low = PriceVol.Low
		Me.Volume = PriceVol.Volume
	End Sub

	Public ReadOnly Property DataType As StockPriceDataType
		Get
			Return _dataType
		End Get
	End Property

	Public Sub SetDataType(DataType As StockPriceDataType)
		_dataType = DataType
	End Sub


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
End Class

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
End Interface
