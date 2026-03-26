Imports YahooAccessData.MathPlus.Measure
Imports YahooAccessData.ExtensionService
Imports System.IO
Imports WebEODData

Public Class RecordPrices
#Const IS_SPLIT_LOCAL_ENABLED = False
#Const IS_SPECIAL_DIVIDEND_ENABLED = False
#Region "Definition"
	Private Const SPLIT_LOG_HIGH As Double = 0.095310179804324865
	Private Const SPLIT_LOG_LOW As Double = -0.1053605156578263
	Private Const SPLIT_LOG_1_5 As Double = 0.40546510810816438
	Private Const SPLIT_LOG_2 As Double = 0.69314718055994529
	Private Const SPLIT_LOG_3 As Double = 1.0986122886681098
	Private Const SPLIT_LOG_4 As Double = 1.3862943611198906
	Private Const SPLIT_LOG_5 As Double = 1.6094379124341003
	Private Const SPLIT_LOG_7 As Double = 1.9459101490553132
	Private Const SPLIT_LOG_10 As Double = 2.3025850929940459
	Private Const SPLIT_LOG_0_1 As Double = -2.3025850929940459
	Private Const SPLIT_LOG_0_25 As Double = -1.3862943611198906
	Private Const SPLIT_LOG_1_FOR_3 As Double = -1.0986122886681098
	Private Const SPLIT_LOG_LIMIT_FOR_CHECK_HIGH As Double = 0.26236426446749106 'limit of 30%
	Private Const SPLIT_LOG_LIMIT_FOR_CHECK_LOW As Double = -0.35667494393873239 'limit of 30%
	Private Const PRICE_SPLIT_CHECK_TO_FUTURE_DAY_POSITION As Integer = 200
	Private Const FILTER_HOLD_FROM_ZERO As Integer = 5

	Private MyPriceVolLast As PriceVol
	Private MyPriceVols() As PriceVol
	Private MyPriceVolsAsLog() As PriceVol
	Private MyPriceVolsIntraDay()() As PriceVol
	Private IsIntraDayLocalEnabled As Boolean
	Private MyFilterForEarningsShare As FilterHoldFromZero
	Private MyFilterForEPSEstimateCurrentYear As FilterHoldFromZero
	Private MyFilterForEPSEstimateNextQuarter As FilterHoldFromZero
	Private MyFilterForEPSEstimateNextYear As FilterHoldFromZero
	Private MyExDividendDateLast As Date
	Private MyDictionaryOfSpecialSplit As Dictionary(Of String, List(Of SplitFactor))
	Private MyDictionaryOfStockDividendSinglePayout As Dictionary(Of String, List(Of StockDividendSinglePayout))
	Private MyDictionaryOfStockPriceDataError As Dictionary(Of String, IList(Of IPriceVol))
	Private MyListOfIPriceVol As List(Of IPriceVol)
	Private MyListOfPriceVol As List(Of PriceVol)

#End Region
#Region "New"

	''' <summary>
	''' the stock data is extracted from the record quote value
	''' The date start is set to the first date in the record quote value
	''' The date stop is set to the user DateStopValue
	''' </summary>
	''' <param name="colData"></param>
	''' <param name="DateStartValue"></param>
	''' <param name="DateStopValue"></param>
	Public Sub New(
		ByRef colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
		ByVal DateStartValue As Date,
		ByVal DateStopValue As Date)

		MyFilterForEarningsShare = New FilterHoldFromZero(FILTER_HOLD_FROM_ZERO)
		MyFilterForEPSEstimateCurrentYear = New FilterHoldFromZero(FILTER_HOLD_FROM_ZERO)
		MyFilterForEPSEstimateNextQuarter = New FilterHoldFromZero(FILTER_HOLD_FROM_ZERO)
		MyFilterForEPSEstimateNextYear = New FilterHoldFromZero(FILTER_HOLD_FROM_ZERO)
		MyExDividendDateLast = YahooAccessData.ReportDate.DateNullValue

		MyDictionaryOfSpecialSplit = New Dictionary(Of String, List(Of SplitFactor))
		MyDictionaryOfStockDividendSinglePayout = New Dictionary(Of String, List(Of StockDividendSinglePayout))
		MyDictionaryOfStockPriceDataError = New Dictionary(Of String, IList(Of IPriceVol))

		If colData.Count = 0 Then
			Throw New InvalidDataException("The record collection is empty.")
		End If
		Me.Stock = colData.Last.Record.Stock
		If Me.Stock Is Nothing Then
			Throw New InvalidDataException("The stock information is missing in the record collection.")
		End If
		Me.Symbol = Me.Stock.Symbol
		'check the limit on the data
		If DateStartValue < Me.Stock.DateStart Then
			'assign the stock StartDate if not specified
			DateStartValue = Me.Stock.DateStart
		End If
		If DateStartValue > DateStopValue Then
			DateStartValue = DateStopValue
		End If

		'add special split factor here
		'eventually need to be in a file
		'Dim ThisSpecialSplit As SplitFactor
		Dim ThisListOfSplitFactor As List(Of SplitFactor)
		ThisListOfSplitFactor = New List(Of SplitFactor)

#If IS_SPLIT_LOCAL_ENABLED Then
		If Me.Stock IsNot Nothing Then
			If Me.Stock.SplitFactors.Count > 0 Then
				For Each ThisSplitFactor In Me.Stock.SplitFactors
					ThisListOfSplitFactor.Add(ThisSplitFactor)
				Next
			End If
			MyDictionaryOfSpecialSplit.Add(Me.Stock.Symbol, ThisListOfSplitFactor)
		Else
			'use the default split factor
			'for example AA here AA divided from the parent company with a share adjustement of 0.81%
			'on Yahoo it is considered a new company but it kept the same symbol with a name change
			'here it is interpreted as a split
			'AA 2016/11/1, 1/0.81
			ThisSpecialSplit = New SplitFactor(New Date(2016, 11, 1))
			With ThisSpecialSplit
				.Ratio = 1 / 0.81
			End With
			ThisListOfSplitFactor.Add(ThisSpecialSplit)
			MyDictionaryOfSpecialSplit.Add("AA", ThisListOfSplitFactor)

			ThisSpecialSplit = New SplitFactor(New Date(year:=2022, month:=2, day:=3))
			ThisListOfSplitFactor = New List(Of SplitFactor)
			With ThisSpecialSplit
				.Ratio = 1.0
			End With
			'ThisListOfSplitFactor.Add(ThisSpecialSplit)
			'FB has a drop a 20% add this so that that the 20% is not interpreted as a split 
			'MyDictionaryOfSpecialSplit.Add("FB", ThisListOfSplitFactor)
			'note for ebay special split on 20 Juillet 2015 of 2376/1000
			ThisSpecialSplit = New SplitFactor(New Date(2015, 7, 20))
			ThisListOfSplitFactor = New List(Of SplitFactor)
			With ThisSpecialSplit
				.Ratio = 2376 / 1000
			End With
			ThisListOfSplitFactor.Add(ThisSpecialSplit)
			MyDictionaryOfSpecialSplit.Add("EBAY", ThisListOfSplitFactor)
		End If
#End If

#If IS_SPECIAL_DIVIDEND_ENABLED Then
		'add StockDividendSinglePayout here
		'eventually need to be in a file
		'Dim ThisStockDividendSinglePayout As StockDividendSinglePayout
		'Dim ThisListOfStockDividendSinglePayout As List(Of StockDividendSinglePayout)

		'for example 
		'Costco Wholesale Corporation Declares Special Cash Dividend of $10 Per Share
		'https://www.globenewswire.com/news-release/2020/11/16/2127799/0/en/Costco-Wholesale-Corporation-Declares-Special-Cash-Dividend-of-10-Per-Share.html#:~:Text = 16%2C%202020%20(GLOBE%20NEWSWIRE),of%20business%20on%20December%202%2C
		'ISSAQUAH, Wash., Nov. 16, 2020 (GLOBE NEWSWIRE) -- 
		'Costco Wholesale Corporation (“Costco” Or the “Company”) (Nasdaq: COST) announced today 
		'that its Board Of Directors has declared a special cash dividend On Costco common stock
		'Of $10 per share, payable December 11, 2020, to shareholders of record as of the close of business on December 2, 2020.
		'The Aggregate payment will be approximately $4.4 billion. The special dividend will be funded through existing cash.

		'note that the announce was fully confirmed only le Lundi 30 now 2020 after the close 2 days before payout ex-dividend 
		'thsi take into account the artificial jump in COST price
		'ThisStockDividendSinglePayout = New StockDividendSinglePayout(
		'  DateReference:=New Date(year:=2020, month:=12, day:=1),
		'  DateOfExDividend:=New Date(year:=2020, month:=12, day:=2),
		'  PricePayoutValue:=10.0)

		'ThisListOfStockDividendSinglePayout = New List(Of StockDividendSinglePayout)
		'ThisListOfStockDividendSinglePayout.Add(ThisStockDividendSinglePayout)
		'MyDictionaryOfStockDividendSinglePayout.Add("COST", ThisListOfStockDividendSinglePayout)

		'Dim ThisListOfStockPriceDataError As IList(Of IPriceVol) = New List(Of IPriceVol)
		'Dim ThisPriceVol As IPriceVol = New PriceVol()
		'With ThisPriceVol
		'  .DateDay = New Date(year:=2018, month:=11, day:=5)
		'  .DateUpdate = .DateDay
		'  .Open = 2.745
		'  .High = 2.8235
		'  .Low = 2.745
		'  .Last = 2.756
		'  .Vol = 163413
		'End With
		'ThisListOfStockPriceDataError.Add(ThisPriceVol)
		'the code to replace the pricevol is not yet completed
		'not use anymore
		'MyDictionaryOfStockPriceDataError.Add("CPR.IDX", ThisListOfStockPriceDataError)
#End If
		'the current object expect the data to be layout as a daily weekly datat set.
		'The data is then processed to create the daily intra day data set
		Dim ThisNumberPoint = ReportDate.MarketTradingDeltaDays(Me.DateStart, Me.DateStop) + 1

		'Filter colData to exclude records that fall on Saturday or Sunday
		'The approach Of checking If two records match Using record Is lastRecord In the Where clause
		'Is relatively fast because it performs a reference equality check. This means
		'it checks whether the two objects refer To the same memory location, which Is an O(1)
		'operation. However, the overall performance of the filtering operation depends
		'on the size of the dataset And the LINQ implementation.
		'The approach Of checking If two records match Using record Is lastRecord In the Where clause
		'Is relatively fast because it performs a reference equality check. This means it
		'checks whether the two objects refer To the same memory location, which Is an O(1) operation.
		'However, the overall performance of the filtering operation depends on the size of the dataset
		'and the LINQ implementation.

		'The behavior of the Where clause in LINQ can be a bit implicit if you're not familiar with how it works.
		'The Where method filters a collection by evaluating the predicate (the function you provide) for each element.
		'If the predicate returns True, the element is included in the result; otherwise, it is excluded.
		'This behavior Is Not explicitly stated In the code itself, which can make it less obvious to someone reading it. 

		'flag indicating that there is more than one item per day
		'always false
		IsIntraDayLocalEnabled = False
		'splitIndex is not use anymore
		ReDim Me.SplitIndex(0 To 0)
		Me.IsVol = False
		If Me.Stock IsNot Nothing Then
			If Me.Stock.IsInternational Then
				colData = FilterAndAdjustWeekendData(colData)
				If colData.Last.DateDay > DateStopValue Then
					DateStopValue = colData.Last.DateDay
				End If
			End If
			Call ProcessDataDailyIntraDay(colData, DateStartValue, DateStopValue)
		End If
	End Sub

	''' <summary>
	''' the stock data is extracted from the record quote value
	''' The date start is set to the first date in the record quote value
	''' The date stop is set to the user DateStopValue
	''' </summary>
	''' <param name="colData"></param>
	''' <param name="DateStopValue"></param>
	Public Sub New(
		ByRef colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
		ByVal DateStopValue As Date)

		Me.New(colData, Date.MinValue, DateStopValue)
	End Sub

	'''' <summary>
	'''' 
	'''' </summary>
	'''' <param name="Symbol"></param>
	'''' <param name="colData"></param>
	'Public Sub New(ByVal Symbol As String, ByRef colData As IEnumerable(Of YahooAccessData.IPriceVol))
	'	Dim I As Integer = 0

	'	Dim ThisPriceVolLast = colData.Last
	'	If ThisPriceVolLast Is Stock Then
	'		Me.Stock = DirectCast(ThisPriceVolLast, Stock)
	'	End If
	'	Me.Symbol = Symbol
	'	If colData.Count = 0 Then
	'		Me.IsError = True
	'		Me.ErrorDescription = "Index Out Of Range Exception."
	'		Throw New IndexOutOfRangeException
	'	End If

	'	ReDim MyPriceVols(0 To colData.Count - 1)
	'	Me.PriceMin = Single.MaxValue
	'	Me.PriceMin = Single.MinValue
	'	Me.IsPriceTarget = False
	'	Me.IsVol = False
	'	Me.VolMin = 0
	'	Me.VolMax = 0
	'	Me.NumberNullPoint = 0
	'	Me.NumberNullPointToEnd = 0
	'	Me.IsError = False
	'	Me.ErrorDescription = ""
	'	IsIntraDayLocalEnabled = False
	'	For Each ThisPriceVol As PriceVol In colData
	'		MyPriceVols(I) = ThisPriceVol
	'		With ThisPriceVol
	'			If .IsIntraDay Then
	'				IsIntraDayLocalEnabled = True
	'			End If
	'			If .Vol > 0 Then
	'				If .Vol < Me.VolMin Then
	'					Me.VolMin = .Vol
	'				End If
	'				If .Vol > Me.VolMax Then
	'					Me.VolMax = .Vol
	'				End If
	'			End If
	'			If .High > Me.PriceMax Then
	'				Me.PriceMax = .High
	'			End If
	'			If .Low > Me.PriceMin Then
	'				Me.PriceMin = .Low
	'			End If
	'			If .IsNull Then
	'				Me.NumberNullPoint += 1
	'			End If
	'		End With
	'		I += 1
	'	Next
	'	If Me.VolMin > 0 Then
	'		Me.IsVol = True
	'	End If
	'	Me.IsSplit = False
	'	Me.DateStart = MyPriceVols(0).DateLastTrade
	'	Me.DateStop = MyPriceVols(MyPriceVols.Count - 1).DateLastTrade
	'	Me.NumberPoint = MyPriceVols.Count
	'	Me.StartPoint = 0
	'	Me.StopPoint = MyPriceVols.Count - 1
	'	Me.PriceVolLast = MyPriceVols(MyPriceVols.Count - 1)
	'End Sub
#End Region
#Region "Public Shared Function"
	''' <summary>
	''' Create a new PriceVol array with time inversed data.
	''' To mark the time inversion the original date of the data is not affected. Only
	''' the open and last value in the structure sequence order is swapped to reflect the time reversal.
	''' </summary>
	''' <param name="PriceVol"></param>
	''' <returns></returns>
	''' <remarks>This can be usuful for special purpose function such as genetic optimization for symetry</remarks>
	Public Shared Function TransformTimeInverse(ByRef PriceVol() As PriceVol) As PriceVol()
		Dim ThisPriceVol(0 To PriceVol.Length - 1) As PriceVol
		Dim I As Integer
		Dim J As Integer

		J = PriceVol.Length - 1
		For I = 0 To PriceVol.Length - 1
			ThisPriceVol(J) = PriceVol(I)
			With ThisPriceVol(J)
				.Open = .Last
				.Last = PriceVol(I).Open
				.OpenNext = .LastPrevious
				.LastPrevious = PriceVol(I).OpenNext
				.LastWeighted = CalculateLastWeighted(ThisPriceVol(J).AsIPriceVol)
				.Range = CalculateTrueRange(ThisPriceVol(J).AsIPriceVol)
			End With
			J = J - 1
		Next
		Return ThisPriceVol
	End Function

	''' <summary>
	''' Create a new array with time inversed data and scaled data.
	''' </summary>
	''' <param name="Value"></param>
	''' <param name="ScaleGain">
	''' The multiplication factor
	''' </param>
	''' <param name="ScaleOffset">
	''' The offset factor</param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function TransformTimeInverse(ByRef Value() As Double, ByVal ScaleGain As Double, ByVal ScaleOffset As Double) As Double()
		Dim ThisValue(0 To Value.Length - 1) As Double
		Dim I As Integer
		Dim J As Integer

		J = Value.Length - 1
		For I = 0 To Value.Length - 1
			ThisValue(J) = ScaleGain * Value(I) + ScaleOffset
			J = J - 1
		Next
		Return ThisValue
	End Function

	''' <summary>
	''' Create a new array with time inversed data with no scaling
	''' </summary>
	''' <param name="Value"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function TransformTimeInverse(ByRef Value() As Double) As Double()
		Return RecordPrices.TransformTimeInverse(Value, 1, 0)
	End Function

	''' <summary>
	''' Move the price last value while keeping all the other price value in the same range
	''' </summary>
	''' <param name="PriceVol">The PriceVol structure to be updated</param>
	''' <param name="PriceLast">The new price last to update</param>
	''' <remarks>
	''' This function return a new PriceVol structure and does not affect the original one. 
	''' It can be use for filtering while preserving the data price range
	''' </remarks>
	Public Shared Function TransformMovePriceVolLast(ByVal PriceVol As PriceVol, ByVal PriceLast As Single) As PriceVol
		With PriceVol
			.Open = PriceLast + (.Open - .Last)
			.High = PriceLast + (.High - .Last)
			.Low = PriceLast + (.Low - .Last)
			.Last = PriceLast
			.FilterLast = PriceLast
		End With
		PriceVol.LastWeighted = RecordPrices.CalculateLastWeighted(PriceVol.AsIPriceVol)
		Return PriceVol
	End Function

	Public Shared Function TransformShiftPriceVolLast(ByVal PriceVol As PriceVol, ByVal PriceLast As Double) As PriceVol
		Dim ThisPriceLast As Single = CSng(PriceLast)
		With PriceVol
			.Open = ThisPriceLast + (.Open - .Last)
			.High = ThisPriceLast + (.High - .Last)
			.Low = ThisPriceLast + (.Low - .Last)
			.Last = ThisPriceLast
			.FilterLast = ThisPriceLast
		End With
		PriceVol.LastWeighted = RecordPrices.CalculateLastWeighted(PriceVol.AsIPriceVol)
		Return PriceVol
	End Function

	Public Shared Function ToListOfPriceVolLast(ByRef PriceVolData() As PriceVol) As List(Of Double)
		If PriceVolData Is Nothing Then Return New List(Of Double)
		Return PriceVolData.Select(Function(pv) CDbl(pv.Last)).ToList()
	End Function


	''' <summary>
	''' calculate a weighted price based on the open, high, low and last of the day
	''' </summary>
	''' <param name="PriceVol"></param>
	''' <returns>return the weighted price</returns>
	''' <remarks></remarks>
	Public Shared Function CalculateLastWeighted(ByRef PriceVol As IPriceVol) As Single
		With PriceVol
			Return (.Open + 2 * (.High + .Low) + 5 * .Last) / 10
		End With
	End Function
#End Region

	''' <summary>
	''' calculate the true range based on the previous last, high, low and current last value of the price
	''' </summary>
	''' <param name="PriceVol"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Function CalculateTrueRange(ByRef PriceVol As IPriceVol) As Single
		Dim ThisRange As Single
		With PriceVol
			If .LastPrevious > .High Then
				ThisRange = .LastPrevious - .Low
			ElseIf .LastPrevious < .Low Then
				ThisRange = .High - .LastPrevious
			Else
				ThisRange = .High - .Low
			End If
		End With
		Return ThisRange
	End Function

	Public Shared Function CalculateLastWeighted(ByRef PriceVol As IPriceVolLarge) As Double
		With PriceVol
			Return (.Open + 2 * (.High + .Low) + 5 * .Last) / 10
		End With
	End Function

	Public Shared Function CalculateTrueRange(ByRef PriceVol As IPriceVolLarge) As Double
		With PriceVol
			If .LastPrevious > .High Then
				Return .LastPrevious - .Low
			ElseIf .LastPrevious < .Low Then
				Return .High - .LastPrevious
			Else
				Return .High - .Low
			End If
		End With
	End Function

#Region "Private Update Function"
	Public Function FilterAndAdjustWeekendData(colData As IEnumerable(Of YahooAccessData.RecordQuoteValue)) As IEnumerable(Of YahooAccessData.RecordQuoteValue)
		Dim lastRecord = colData.LastOrDefault()
		If lastRecord Is Nothing Then
			Throw New InvalidOperationException("The record collection is empty.")
		End If

		'keep week-end tradin data but move it to next monday
		Select Case lastRecord.DateDay.DayOfWeek
			Case DayOfWeek.Saturday
				lastRecord.Record.DateDay = lastRecord.DateDay.AddDays(2)
			Case DayOfWeek.Sunday
				lastRecord.Record.DateDay = lastRecord.DateDay.AddDays(1)
		End Select
		Return colData

		'Return colData.Where(
		'	Function(record)
		'		Return record Is lastRecord OrElse
		'			(record.DateDay.DayOfWeek <> DayOfWeek.Saturday AndAlso
		'			record.DateDay.DayOfWeek <> DayOfWeek.Sunday)
		'	End Function)
	End Function


	''' <summary>
	''' Estimate the end-of-day (EOD) trading volume based on the current volume and time of day.
	''' 
	''' This uses a flat-clock extrapolation, assuming trading volume continues at the same average pace
	''' observed so far in the session. Optionally, the estimate is blended with yesterday's total EOD volume,
	''' where the influence of yesterday decreases gradually as the trading day progresses.
	''' 
	''' - Early in the day, the estimate relies more on yesterday's volume (if provided).
	''' - Later in the day, it relies more on today's realized pace (flat-clock extrapolation).
	''' - If no VolumeEODLast is provided, only the flat estimate is returned.
	''' </summary>
	''' <param name="VolumeNow">The cumulative traded volume observed so far today.</param>
	''' <param name="TimeOfDayNow">The current local time (exchange time).</param>
	''' <param name="TimeOfOpen">The market session opening time.</param>
	''' <param name="TimeOfClose">The market session closing time.</param>
	''' <param name="VolumeEODLast">
	''' The total volume observed at the end of the previous trading day (optional).
	''' If provided, it is blended with the flat estimate with decreasing weight over time.
	''' </param>
	''' <param name="p">
	''' Controls how fast the weight shifts from yesterday's EOD volume to today's flat estimate:
	'''  p = 1.0 : linear fade (default)
	'''  p less than 1.0 : slower fade (yesterday influences longer)
	'''  p > 1.0 : faster fade (today dominates earlier)
	''' </param>
	''' <returns>
	''' The estimated total end-of-day volume based on current trading activity and optional historical reference.
	''' </returns>
	Public Function EstimateEODVolumeSimple(
		VolumeNow As Double,
		TimeOfDayNow As DateTime,
		TimeOfOpen As DateTime,
		TimeOfClose As DateTime,
		Optional VolumeEODLast As Double? = Nothing,
		Optional Fading As Double = 1.0) As Long

		If TimeOfClose <= TimeOfOpen Then Throw New ArgumentException("TimeOfClose must be after TimeOfOpen.")
		Dim totalSecs = (TimeOfClose - TimeOfOpen).TotalSeconds

		'Dim elapsed = Math.Max(0.0, Math.Min(totalSecs, (TimeOfDayNow - TimeOfOpen).TotalSeconds))

		' Compute elapsed seconds since market open
		Dim elapsed As Double = (TimeOfDayNow - TimeOfOpen).TotalSeconds

		' Clamp elapsed time between 0 and totalSecs
		If elapsed < 0.0 Then
			elapsed = 0.0
		ElseIf elapsed > totalSecs Then
			elapsed = totalSecs
		End If


		' Session progress τ in [0,1]
		Dim tau = elapsed / totalSecs

		' Flat-clock extrapolation: estimate EOD volume assuming constant trading pace.
		' Use a small floor (eps) to avoid division by zero near the open.
		Dim eps As Double = 0.01
		'Dim flatFrac As double= Math.Max(eps, tau)
		Dim flatFrac As Double
		If tau < eps Then
			flatFrac = eps
		Else
			flatFrac = tau
		End If
		Dim Vflat As Double

		' If no previous day volume is provided, return the flat estimate directly.
		If VolumeEODLast.HasValue Then
			' Linear or power fade from yesterday's EOD volume to today's flat estimate.
			' At the open (tau=0), weight of yesterday = 1.0
			' At the close (tau=1), weight of yesterday = 0.0
			Dim wFlat As Double = Math.Pow(tau, Math.Max(0.1, Fading)) 'tau^0.1 or tau^Fading
			Dim wYest As Double = 1.0 - wFlat

			' Final blended estimate
			Vflat = wFlat * (VolumeNow / flatFrac) + wYest * VolumeEODLast.Value
		Else
			Vflat = VolumeNow / flatFrac
		End If

		If Vflat > Long.MaxValue Then
			Return Long.MaxValue
		ElseIf Vflat < 0 Then
			Return 0
		Else
			Return CLng(Vflat)
		End If
	End Function


	Private Sub ProcessDataDailyIntraDay(
		ByRef colData As IEnumerable(Of YahooAccessData.RecordQuoteValue),
		ByVal DateStartValue As Date,
		ByVal DateStopValue As Date)

		Dim ThisDateCurrent As Date
		'Dim colDataDaily As IEnumerable(Of YahooAccessData.RecordQuoteValue)


		Dim ThisListOfSplitIndex As New List(Of Integer)
		Dim ThisListOfSSpecialDividendPayout As New List(Of Integer)
		Dim ThisRecordQuoteValue As YahooAccessData.RecordQuoteValue
		'Dim ThisRecordQuoteValuePrevious As YahooAccessData.RecordQuoteValue
		Dim ThisListOfPriceVols As New List(Of PriceVol)
		Dim IsLiveUpdate As Boolean
		'Dim ThisDateStopEndOfDay As Date = DateStopValue.Date.AddHours(24).AddSeconds(-1)
		'set the default value
		'adjust the date for the constraint of always starting on Monday and ignore the weekend 
		Me.DateStart = ReportDate.DateToMondayPrevious(DateStartValue.Date)
		Me.DateStop = DateStopValue.Date
		If Me.DateStop < Me.DateStart Then
			Me.DateStop = Me.DateStart
		End If
		'If Date fall on a weekend go to the next monday even if it is a future date
		Select Case Me.DateStop.DayOfWeek
			Case DayOfWeek.Saturday
				Me.DateStop = Me.DateStop.AddDays(2)
			Case DayOfWeek.Sunday
				Me.DateStop = Me.DateStop.AddDays(1)
		End Select
		'not used anymore
		'because MarketTradingDeltaDays take into account the current date and the number of tradindays may not be correct
		'for a 24h tradin period particularly in the weekend trading session which is now supported
		'to solve this it is better to do teh process in reverse in use a list isntead of an array
		'at the end to maintain teh compatibillity with the previous code the list is converted to an array
		'not use anymore
		'Me.NumberPoint = ReportDate.MarketTradingDeltaDays(Me.DateStart, Me.DateStop) + 1
		'ReDim MyPriceVols(0 To Me.NumberPoint - 1)
		'array declaration

		'same story with the intraday data
		'MyPriceVolsIntraDay = New PriceVol(0 To Me.NumberPoint - 1)() {}
		'extract the intraday data from the full stream time data for all the stream
		'only the daily data is used for the calculation	
		'do I really need to read from the intraday? No
		'this was writtent long time ago to deal with the yahoo dat
		'let try something else
		'Dim ThisListOfPriceVolsIntraDay = New List(Of IEnumerable(Of YahooAccessData.RecordQuoteValue))(colData.ToDailyIntraDay(Me.DateStart, Me.DateStop.AddHours(24).AddSeconds(-1)))
		'colData As IEnumerable(Of YahooAccessData.RecordQuoteValue)

		'Dim ThisListOfRecordQuoteValue = colData.SkipWhile(Function(RecordQuote) RecordQuote.DateDay.DayOfWeek = DayOfWeek.Saturday OrElse RecordQuote.DateDay.DayOfWeek = DayOfWeek.Sunday).ToList

		'this call group the data by day it include all date including weekend between DateStart and DateStop
		'colDataDailyIntraDay = colData.ToDailyIntraDay(Me.DateStart, Me.DateStop.AddHours(24).AddSeconds(-1))
		'patch the data to a standard weekly stream data from Monday to Friday
		'data that fall on a weekend are ignored unles it is the last data point which is used to update the last price
		'and possibly reporting the weekend data to teh future stream opening on monday

		'MyPriceVolsIntraDay = New PriceVol(0 To Me.NumberPoint - 1)() {}
		'Start the processing of the data with a nul price
		'just in case make sure teh data in teh record also start on a weekly stream
		'we doing that in case teh datae stsrt have been improperly selected
		'to start on a date that is not the usual working day
		'essentialy a safeguard againt mistaken supported date. More relevant for 24h trading stream
		Dim ThisListOfRecordQuoteValue = New List(Of YahooAccessData.RecordQuoteValue)
		For Each ThisRecordQuoteValueWithIndex In colData.WithIndex
			'important to remove the time information in this test
			'the last record may contain the time information if it is a live type of record
			'set the value of the valid data in the data stream
			If _
					ThisRecordQuoteValueWithIndex.Item.DateDay.DayOfWeek = DayOfWeek.Saturday OrElse
					ThisRecordQuoteValueWithIndex.Item.DateDay.DayOfWeek = DayOfWeek.Sunday Then

				'ignore the week end date
				'the current program expect to have the data presented as a weekly stream
				Continue For
			End If
			ThisListOfRecordQuoteValue = New List(Of YahooAccessData.RecordQuoteValue)(colData.Skip(ThisRecordQuoteValueWithIndex.Index))
			Exit For
		Next
		MyListOfIPriceVol = New List(Of IPriceVol)
		'set the default value for that condition
		Me.StartPoint = -1
		Me.StopPoint = -1
		If ThisListOfRecordQuoteValue.Count = 0 Then
			Me.IsNull = True
			'in that case initialize the pricevol with null v`alues
			Dim ThisDateStop As Date = Me.DateStart
			Do
				MyListOfIPriceVol.Add(New PriceVol(0) With {.IsNull = True, .DateLastTrade = ThisDateStop})
				ThisDateStop = ThisDateStop.AddDays(1)
				'sunday is not possible here because DateStart is always a monday and we increase by one day at a time
				If Me.DateStop.DayOfWeek = DayOfWeek.Saturday Then
					ThisDateStop = ThisDateStop.AddDays(2)
				End If
			Loop Until ThisDateStop > Me.DateStop
			Me.NumberPoint = MyListOfIPriceVol.Count
			Me.NumberNullPoint = Me.NumberPoint
			ReDim MyPriceVols(0 To Me.NumberPoint - 1)

			For Each ItemIndexed In MyListOfIPriceVol.WithIndex
				MyPriceVols(ItemIndexed.Index) = DirectCast(ItemIndexed.Item, PriceVol)
			Next
			Me.PriceMax = 0
			Me.PriceMin = 0
			Me.PriceMinTarget = 0
			Me.PriceMaxTarget = 0
			Me.VolMax = 0
			Me.VolMin = 0
			Me.IsPriceTarget = False



			Return
		Else
			'initialize the range variable before we start the data processing
			Me.PriceMax = 0
			Me.PriceMin = Single.MaxValue
			Me.PriceMinTarget = Me.PriceMin
			Me.PriceMaxTarget = Me.PriceMax
			Me.VolMax = 0
			Me.VolMin = Integer.MaxValue
			Me.IsPriceTarget = False
		End If
		'adjust the data
		ThisListOfPriceVols.Clear()
		Me.StartPoint = -1

		'synchronize the start with the actual data
		Dim ThisPriceVol As PriceVol
		ThisDateCurrent = Me.DateStart
		MyListOfIPriceVol.Clear()
		'collect to be on a valid weekly working day
		'for the 24h trading stream we will keep only the last item and change the trading day
		'to be on the next monday
		Dim ThisRecordQuoteValueFirst = ThisListOfRecordQuoteValue.First
		'scan the list and collect the priceVol data 
		Do Until ThisDateCurrent >= ThisRecordQuoteValueFirst.DateDay.Date
			ThisPriceVol = PriceVolUpdateToNull(ThisRecordQuoteValueFirst, ThisDateCurrent)
			Me.NumberNullPoint = Me.NumberNullPoint + 1
			If MyListOfIPriceVol.Count > 0 Then
				MyListOfIPriceVol.Last.OpenNext = ThisPriceVol.Open
			End If
			MyListOfIPriceVol.Add(ThisPriceVol)
			ThisDateCurrent = ThisDateCurrent.AddDays(1)
			ThisDateCurrent = If(ThisDateCurrent.DayOfWeek = DayOfWeek.Saturday, ThisDateCurrent.AddDays(2), ThisDateCurrent)
		Loop
		Me.StartPoint = MyListOfIPriceVol.Count - 1
		Dim ThisRecordQuoteValueLast = ThisListOfRecordQuoteValue.Last
		Dim ThisRecordQuoteValuePrevious = ThisRecordQuoteValueFirst
		For Each ThisRecordQuoteValue In ThisListOfRecordQuoteValue
			'important to remove the time information in this test
			'the last record may contain the time information if it is a live type of record
			'set the value of the valid data in the data stream
			If ThisRecordQuoteValue Is ThisRecordQuoteValueLast Then
				ThisRecordQuoteValue = ThisRecordQuoteValue
			End If
			If _
				ThisRecordQuoteValue.DateDay.DayOfWeek = DayOfWeek.Saturday OrElse
				ThisRecordQuoteValue.DateDay.DayOfWeek = DayOfWeek.Sunday Then
				'ignore the week end date
				'the current program expect to have the data presented as a weekly stream
				Continue For
			End If
			Do Until ThisDateCurrent >= ThisRecordQuoteValue.DateDay.Date
				ThisPriceVol = PriceVolUpdateToNull(ThisRecordQuoteValuePrevious, ThisDateCurrent)
				Me.NumberNullPoint = Me.NumberNullPoint + 1
				If MyListOfIPriceVol.Count > 0 Then
					MyListOfIPriceVol.Last.OpenNext = ThisPriceVol.Open
				End If
				MyListOfIPriceVol.Add(ThisPriceVol)
				ThisDateCurrent = ThisDateCurrent.AddDays(1)
				ThisDateCurrent = If(ThisDateCurrent.DayOfWeek = DayOfWeek.Saturday, ThisDateCurrent.AddDays(2), ThisDateCurrent)
			Loop
			ThisPriceVol = PriceVolUpdate(ThisRecordQuoteValue, ThisDateCurrent)
			If MyListOfIPriceVol.Count > 0 Then
				MyListOfIPriceVol.Last.OpenNext = ThisPriceVol.Open
			End If
			MyListOfIPriceVol.Add(ThisPriceVol)
			If ThisPriceVol.Volume > 0 Then
				Me.IsVol = True
			End If
			ThisDateCurrent = ThisDateCurrent.AddDays(1)
			ThisDateCurrent = If(ThisDateCurrent.DayOfWeek = DayOfWeek.Saturday, ThisDateCurrent.AddDays(2), ThisDateCurrent)
			ThisRecordQuoteValuePrevious = ThisRecordQuoteValue
		Next
		Me.StopPoint = MyListOfIPriceVol.Count - 1
		Me.NumberPoint = MyListOfIPriceVol.Count
		If Me.NumberNullPoint = Me.NumberPoint Then
			Me.IsNull = True
		Else
			Me.IsNull = False
		End If
		'last we need to adjust for to see if the last record is liveupdate or not and update PriceVol
		If colData.Last.Record.AsIRecordType.RecordType = IRecordType.enuRecordType.LiveUpdate Then
			IsLiveUpdate = True
		Else
			IsLiveUpdate = False
		End If
		'propagate the liveupdate information form the record to the PriceVol object	
		'note that there is a name change here 
		ReDim MyPriceVols(0 To Me.NumberPoint - 1)
		For Each ThisItemIndexed In MyListOfIPriceVol.WithIndex
			MyPriceVols(ThisItemIndexed.Index) = DirectCast(ThisItemIndexed.Item, PriceVol)
		Next
		MyPriceVols(Me.NumberPoint - 1).IsIntraDay = IsLiveUpdate
	End Sub

	Private Sub ProcessSplitAdjustForIntraDay(ByRef PriceVolIntraDay() As PriceVol, ByRef PriceVol As PriceVol)
		Dim ThisPriceVol As PriceVol
		Dim ThisVol As Long
		Dim ThisVolPlus As Long
		Dim ThisVolMinus As Long
		Dim ThisVolLast As Long = 0
		Dim ThisLastPrevious As Single = PriceVol.Last

		For Each ThisPriceVol In PriceVolIntraDay
			With ThisPriceVol
				If PriceVol.LastAdjusted <> 1.0 Then
					.Open = .Open / PriceVol.LastAdjusted
					.OpenNext = PriceVol.OpenNext
					.High = .High / PriceVol.LastAdjusted
					.Low = .Low / PriceVol.LastAdjusted
					.Last = .Last / PriceVol.LastAdjusted
					.LastWeighted = .LastWeighted / PriceVol.LastAdjusted
					ThisVol = CLng(.Volume * PriceVol.LastAdjusted)
					.Vol = ThisVol.ToIntegerSafe
					.Volume = ThisVol
					.LastAdjusted = PriceVol.LastAdjusted
					.OneyrTargetPrice = PriceVol.OneyrTargetPrice
					.LastPrevious = PriceVol.LastPrevious
					.Range = CalculateTrueRange(ThisPriceVol.AsIPriceVol)
				End If
				If .Last >= ThisLastPrevious Then
					ThisVolPlus = ThisVolPlus + (.Volume - ThisVolLast)
				Else
					ThisVolMinus = ThisVolMinus + (.Volume - ThisVolLast)
				End If
				ThisLastPrevious = .Last
				ThisVolLast = .Volume
			End With
		Next
		If ThisVolPlus > Long.MaxValue Then
			PriceVol.VolPlus = Long.MaxValue
		Else
			PriceVol.VolPlus = ThisVolPlus
		End If
		If ThisVolMinus > Long.MaxValue Then
			PriceVol.VolMinus = Long.MaxValue
		Else
			PriceVol.VolMinus = ThisVolMinus
		End If
		PriceVol.IsIntraDayMultipleValueDetected = IsIntraDayLocalEnabled
	End Sub

	Private Function MeasureStockSplit(ByVal PriceLast As Single, ByVal PriceLastPrevious As Single) As Single
		Dim ThisPriceLogRatio As Double

		If Me.Stock.IsSplitEnabled = False Then Return 1.0
		If PriceLastPrevious <> 0 Then
			ThisPriceLogRatio = Math.Log(PriceLastPrevious / PriceLast)
			If (ThisPriceLogRatio > SPLIT_LOG_LIMIT_FOR_CHECK_HIGH) Or (ThisPriceLogRatio < SPLIT_LOG_LIMIT_FOR_CHECK_LOW) Then
				'The most common stock splits are, 2-for-1, 3-for-2 and 3-for-1
				'It is also possible to have a reverse stock split: i.e. 1-for-10,  1-for-3
				'pick the closest one to these ratio in 10% and add more later when needed
				Select Case ThisPriceLogRatio
					Case (SPLIT_LOG_1_5 + SPLIT_LOG_LOW) To (SPLIT_LOG_1_5 + SPLIT_LOG_HIGH)
						Return 1.5
					Case (SPLIT_LOG_2 + SPLIT_LOG_LOW) To (SPLIT_LOG_2 + SPLIT_LOG_HIGH)
						Return 2.0
					Case (SPLIT_LOG_3 + SPLIT_LOG_LOW) To (SPLIT_LOG_3 + SPLIT_LOG_HIGH)
						Return 3.0
					Case (SPLIT_LOG_4 + SPLIT_LOG_LOW) To (SPLIT_LOG_4 + SPLIT_LOG_HIGH)
						Return 4.0
					Case (SPLIT_LOG_5 + SPLIT_LOG_LOW) To (SPLIT_LOG_5 + SPLIT_LOG_HIGH)
						Return 5.0
					Case (SPLIT_LOG_7 + SPLIT_LOG_LOW) To (SPLIT_LOG_7 + SPLIT_LOG_HIGH)
						Return 7.0
					Case (SPLIT_LOG_10 + SPLIT_LOG_LOW) To (SPLIT_LOG_10 + SPLIT_LOG_HIGH)
						Return 10.0
					Case (SPLIT_LOG_0_1 + SPLIT_LOG_LOW) To (SPLIT_LOG_0_1 + SPLIT_LOG_HIGH)
						Return 0.1
					Case (SPLIT_LOG_1_FOR_3 + SPLIT_LOG_LOW) To (SPLIT_LOG_1_FOR_3 + SPLIT_LOG_HIGH)
						Return 1 / 3
					Case (SPLIT_LOG_0_25 + SPLIT_LOG_LOW) To (SPLIT_LOG_0_25 + SPLIT_LOG_HIGH)
						Return 0.25
					Case Else
						Return 1.0
				End Select
			Else
				Return 1.0
			End If
		Else
			Return 1.0
		End If
	End Function

	Private Function PriceVolUpdateToNull(
		ByRef Record As YahooAccessData.RecordQuoteValue,
		ByRef DateValue As Date) As PriceVol

		If MyPriceVolLast Is Nothing Then
			MyPriceVolLast = New PriceVol(Record.Open)
		End If
		Dim ThisPriceVol As New PriceVol
		With ThisPriceVol
			.DateLastTrade = DateValue
			.LastPrevious = MyPriceVolLast.Last
			If DateValue < Record.DateDay Then
				.Open = .LastPrevious
				.DividendYield = MyPriceVolLast.DividendYield
				.DividendShare = MyPriceVolLast.DividendShare
				.DividendPayDate = MyPriceVolLast.DividendPayDate
				.ExDividendDate = MyPriceVolLast.ExDividendDate
				.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
				.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
				.EarningsShare = MyPriceVolLast.EarningsShare
				.EPSEstimateCurrentYear = MyPriceVolLast.EPSEstimateCurrentYear
				.EPSEstimateNextQuarter = MyPriceVolLast.EPSEstimateNextQuarter
				.EPSEstimateNextYear = MyPriceVolLast.EPSEstimateNextYear
				.OneyrPEG = MyPriceVolLast.OneyrPEG
				.FiveyrPEG = MyPriceVolLast.FiveyrPEG
			Else
				If (Record.Vol = 0 And Record.Last = 0) Then
					.Open = MyPriceVolLast.Last
					.DividendYield = MyPriceVolLast.DividendYield
					.DividendShare = MyPriceVolLast.DividendShare
					.DividendPayDate = MyPriceVolLast.DividendPayDate
					.ExDividendDate = MyPriceVolLast.ExDividendDate
					.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
					.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
					.EarningsShare = MyPriceVolLast.EarningsShare
					.EPSEstimateCurrentYear = MyPriceVolLast.EPSEstimateCurrentYear
					.EPSEstimateNextQuarter = MyPriceVolLast.EPSEstimateNextQuarter
					.EPSEstimateNextYear = MyPriceVolLast.EPSEstimateNextYear
					.OneyrPEG = MyPriceVolLast.OneyrPEG
				Else
					.Open = Record.Last
					.DividendYield = Record.DividendYield
					.DividendShare = Record.DividendShare
					.DividendPayDate = Record.DividendPayDate
					.ExDividendDate = Record.ExDividendDate
					.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
					.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
					.EarningsShare = Record.EarningsShare
					.EPSEstimateCurrentYear = Record.EPSEstimateCurrentYear
					.EPSEstimateNextQuarter = Record.EPSEstimateNextQuarter
					.EPSEstimateNextYear = Record.EPSEstimateNextYear
					.OneyrPEG = Record.PEGRatio
				End If
				.FiveyrPEG = .OneyrPEG
			End If
			'If .OneyrPEG < 0 Then .OneyrPEG = 0
			If .OneyrPEG = 0 Then
				.OneyrPEG = MyPriceVolLast.OneyrPEG
			End If
			.FiveyrPEG = .OneyrPEG
			.High = .Open
			.Low = .Open
			.Last = .Open
			.OpenNext = .Last 'by default
			.LastWeighted = .Open
			.Vol = 0
			.Volume = 0
			.Range = 0
			.IsNull = True
			If DateValue < Record.DateDay Then
				.OneyrTargetPrice = MyPriceVolLast.OneyrTargetPrice
				.OneyrTargetEarning = MyPriceVolLast.OneyrTargetEarning
				.OneyrTargetEarningGrow = MyPriceVolLast.OneyrTargetEarningGrow
				.FiveyrTargetEarningGrow = MyPriceVolLast.FiveyrTargetEarningGrow
			Else
				If Record.OneyrTargetPrice = 0 Then
					.OneyrTargetPrice = MyPriceVolLast.OneyrTargetPrice
				Else
					.OneyrTargetPrice = Record.OneyrTargetPrice
				End If
				If Record.EPSEstimateNextYear = 0 Then
					.OneyrTargetEarning = MyPriceVolLast.OneyrTargetEarning
				End If
				If .EarningsShare > 0 Then
					.OneyrTargetEarningGrow = CSng(Math.Log(.OneyrTargetEarning / .EarningsShare))
				Else
					.OneyrTargetEarningGrow = 0
				End If
				.FiveyrTargetEarningGrow = 0   'by defauly
				If (.EPSEstimateCurrentYear > 0) And (.OneyrPEG > 0) Then
					.FiveyrTargetEarningGrow = (.Last / .EPSEstimateCurrentYear) / .OneyrPEG
				End If
			End If
			If .Last > 0 Then
				'check the range
				If .High > Me.PriceMax Then
					Me.PriceMax = .High
				End If
				If .Low > 0 Then
					If .Low < Me.PriceMin Then
						Me.PriceMin = .Low
					End If
				End If
				If .OneyrTargetPrice > 0 Then
					Me.IsPriceTarget = True
					If .OneyrTargetPrice > Me.PriceMaxTarget Then
						Me.PriceMaxTarget = .OneyrTargetPrice
					End If
					If .OneyrTargetPrice < Me.PriceMinTarget Then
						Me.PriceMinTarget = .OneyrTargetPrice
					End If
				End If
				If .Vol < Me.VolMin Then
					Me.VolMin = .Vol
				End If
				If .Vol > Me.VolMax Then
					Me.VolMax = .Vol
				End If
				.DividendYield = 100 * .DividendShare / .Last
			End If
			If .EarningsShare = 0 Then
				.EarningsShare = MyPriceVolLast.EarningsShare
			End If
			MyPriceVolLast = ThisPriceVol
			.RecordQuoteValue = Record
		End With
		Return ThisPriceVol
	End Function

	Private Function PriceVolUpdate(
		ByRef RecordQuote As YahooAccessData.RecordQuoteValue,
		ByVal DateValue As Date) As PriceVol

		If MyPriceVolLast Is Nothing Then
			MyPriceVolLast = New PriceVol(RecordQuote.Open)
		End If
		Dim ThisPriceVol As New PriceVol
		With ThisPriceVol
			With .AsISentimentIndicator
				.Count = RecordQuote.AsISentimentIndicator.Count
				.Value = RecordQuote.AsISentimentIndicator.Value
			End With
			If RecordQuote.Record.AsIRecordType.RecordType = IRecordType.enuRecordType.LiveUpdate Then
				'use the live update time in this specia; case when the day is not finished
				.DateLastTrade = RecordQuote.Record.DateUpdate
			Else
				.DateLastTrade = DateValue
			End If
			.LastPrevious = MyPriceVolLast.Last
			If .LastPrevious = 0 Then
				.LastPrevious = RecordQuote.Open
			End If
			.Open = RecordQuote.Open
			.High = RecordQuote.High
			.Low = RecordQuote.Low
			.Last = RecordQuote.Last
			If RecordQuote.OneyrTargetPrice = 0 Then
				.OneyrTargetPrice = MyPriceVolLast.OneyrTargetPrice
				.EarningsShare = MyPriceVolLast.EarningsShare
				.EPSEstimateCurrentYear = MyPriceVolLast.EPSEstimateCurrentYear
				.EPSEstimateNextQuarter = MyPriceVolLast.EPSEstimateNextQuarter
				.EPSEstimateNextYear = MyPriceVolLast.EPSEstimateNextYear
				.OneyrTargetEarning = MyPriceVolLast.OneyrTargetEarning
				.OneyrTargetEarningGrow = MyPriceVolLast.OneyrTargetEarningGrow
				.OneyrPEG = MyPriceVolLast.OneyrPEG
			Else
				.OneyrTargetPrice = RecordQuote.OneyrTargetPrice
				.EarningsShare = RecordQuote.EarningsShare
				.EPSEstimateCurrentYear = RecordQuote.EPSEstimateCurrentYear
				.EPSEstimateNextQuarter = RecordQuote.EPSEstimateNextQuarter
				.EPSEstimateNextYear = RecordQuote.EPSEstimateNextYear

				'filter for zero yahoo anomalies
				.EarningsShare = MyFilterForEarningsShare.Filter(.EarningsShare)
				.EPSEstimateCurrentYear = MyFilterForEPSEstimateCurrentYear.Filter(.EPSEstimateCurrentYear)
				.EPSEstimateNextQuarter = MyFilterForEPSEstimateNextQuarter.Filter(.EPSEstimateNextQuarter)
				.EPSEstimateNextYear = MyFilterForEPSEstimateNextYear.Filter(.EPSEstimateNextYear)

				.OneyrTargetEarning = (RecordQuote.EPSEstimateNextYear + .EPSEstimateCurrentYear) / 2
				.OneyrPEG = RecordQuote.PEGRatio
				If .EarningsShare > 0 Then
					.OneyrTargetEarningGrow = CSng(Math.Log(.OneyrTargetEarning / .EarningsShare))
				Else
					.OneyrTargetEarningGrow = 0
				End If
				If .OneyrPEG = 0 Then
					.OneyrPEG = MyPriceVolLast.OneyrPEG
				End If
				.FiveyrPEG = .OneyrPEG
				.DividendYield = RecordQuote.DividendYield
				.DividendShare = RecordQuote.DividendShare
				.DividendPayDate = RecordQuote.DividendPayDate
				.ExDividendDate = RecordQuote.ExDividendDate
				If .ExDividendDate > ReportDate.DateNullValue Then
					If MyPriceVolLast.ExDividendDate > ReportDate.DateNullValue Then
						If .ExDividendDate <> MyPriceVolLast.ExDividendDate Then
							If .ExDividendDate > MyPriceVolLast.ExDividendDate Then
								'this is a new ex dividend date
								.ExDividendDatePrevious = MyPriceVolLast.ExDividendDate
								.ExDividendDateEstimated = .ExDividendDate.AddDays(.ExDividendDate.Subtract(.ExDividendDatePrevious).Days)
							Else
								'the ex-dividend can go back down in time
								.ExDividendDate = MyPriceVolLast.ExDividendDate
								.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
								.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
							End If
						Else
							.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
							.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
						End If
					Else
						.ExDividendDatePrevious = .ExDividendDate
						.ExDividendDateEstimated = .ExDividendDate
					End If
				Else
					If MyPriceVolLast.ExDividendDate >= ReportDate.DateNullValue Then
						'this is an invalid update due to yahoo poor management of ex-dividend
						'in this case all the data on the dividend is false
						'a correction is possible using the the previous data
						.DividendYield = MyPriceVolLast.DividendYield
						.DividendShare = MyPriceVolLast.DividendShare
						.DividendPayDate = MyPriceVolLast.DividendPayDate
						.ExDividendDate = MyPriceVolLast.ExDividendDate
						.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
						.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
					Else
						.ExDividendDate = ReportDate.DateNullValue
						.ExDividendDatePrevious = ReportDate.DateNullValue
						.ExDividendDateEstimated = ReportDate.DateNullValue
					End If
				End If
			End If
			'note from Wikipedia:
			'The growth rate is expressed as a percentage above 100%, and should use real growth only, 
			'to correct for inflation. E.g. if a company is growing at 30% a year, and has a P/E of 30, it would have a PEG of 1.
			'A lower ratio is "better" (cheaper) and a higher ratio is "worse" (expensive).
			'The P/E ratio used in the calculation may be projected or trailing, and the annual growth rate may be the expected growth rate for the next year or the next five years.
			'Examples:
			'Yahoo! Finance uses 5-year expected growth rate and a P/E based on the EPS estimate for the current fiscal year for calculating PEG (PEG for IBM is 1.26 on Aug 9, 2008 [1]).
			'The NASDAQ web-site uses the forecast growth rate (based on the consensus of professional analysts) and forecast earnings over the next 12 months. (PEG for IBM is 1.148 on Aug 9, 2008 [2]).
			.FiveyrTargetEarningGrow = 0   'by default
			If (.EPSEstimateCurrentYear > 0) And (.OneyrPEG > 0) Then
				.FiveyrTargetEarningGrow = (.Last / .EPSEstimateCurrentYear) / .OneyrPEG
			End If

			.FiveyrPEG = RecordQuote.PEGRatio
			.Vol = RecordQuote.Vol
			.Volume = RecordQuote.Volume
			If .DividendShare = 0 Then
				.DividendShare = MyPriceVolLast.DividendShare
			End If
			If (.Volume = 0 And .Last = 0) Then
				.Open = MyPriceVolLast.Last
				.High = .Open
				.Low = .Open
				.Last = .Open
				.IsNull = True
				Me.NumberNullPoint = Me.NumberNullPoint + 1
			Else
				.IsNull = False
			End If
			'check the range
			If .High > Me.PriceMax Then
				Me.PriceMax = .High
			End If
			If .Low > 0 Then
				If .Low < Me.PriceMin Then
					Me.PriceMin = .Low
				End If
			End If
			If .OneyrTargetPrice > 0 Then
				Me.IsPriceTarget = True
				If .OneyrTargetPrice > Me.PriceMaxTarget Then
					Me.PriceMaxTarget = .OneyrTargetPrice
				End If
				If .OneyrTargetPrice < Me.PriceMinTarget Then
					Me.PriceMinTarget = .OneyrTargetPrice
				End If
			End If
			If .Volume < Me.VolMin Then
				Me.VolMin = .Volume
			End If
			If .Volume > Me.VolMax Then
				Me.VolMax = .Volume
			End If
			.LastWeighted = RecordPrices.CalculateLastWeighted(DirectCast(ThisPriceVol, IPriceVol))
			.Range = RecordPrices.CalculateTrueRange(DirectCast(ThisPriceVol, IPriceVol))
			MyPriceVolLast = ThisPriceVol
			If .Last > 0 Then
				.DividendYield = 100 * .DividendShare / .Last
			End If
			.RecordQuoteValue = RecordQuote
		End With
		Return ThisPriceVol
	End Function

	Private Function PriceVolUpdateIntraDay(
		ByRef Record As YahooAccessData.RecordQuoteValue,
		ByRef DateValue As Date) As PriceVol

		Dim ThisPriceVol As New PriceVol
		With ThisPriceVol
			.DateLastTrade = Record.DateLastTrade
			With .AsIPriceVol
				.DateDay = DateValue
				.DateUpdate = Record.DateUpdate
			End With
			.LastPrevious = MyPriceVolLast.Last
			If .LastPrevious = 0 Then
				.LastPrevious = Record.Open
			End If
			.Open = Record.Open
			.High = Record.High
			.Low = Record.Low
			.Last = Record.Last
			If Record.OneyrTargetPrice = 0 Then
				.OneyrTargetPrice = MyPriceVolLast.OneyrTargetPrice
				.EarningsShare = MyPriceVolLast.EarningsShare
				.EPSEstimateCurrentYear = MyPriceVolLast.EPSEstimateCurrentYear
				.EPSEstimateNextQuarter = MyPriceVolLast.EPSEstimateNextQuarter
				.EPSEstimateNextYear = MyPriceVolLast.EPSEstimateNextYear
				.OneyrTargetEarning = MyPriceVolLast.OneyrTargetEarning
				.OneyrTargetEarningGrow = MyPriceVolLast.OneyrTargetEarningGrow
				.OneyrPEG = MyPriceVolLast.OneyrPEG
			Else
				.OneyrTargetPrice = Record.OneyrTargetPrice
				.EarningsShare = Record.EarningsShare
				.EPSEstimateCurrentYear = Record.EPSEstimateCurrentYear
				.EPSEstimateNextQuarter = Record.EPSEstimateNextQuarter
				.EPSEstimateNextYear = Record.EPSEstimateNextYear
				.OneyrTargetEarning = (Record.EPSEstimateNextYear + .EPSEstimateCurrentYear) / 2
				.OneyrPEG = Record.PEGRatio
				If .EarningsShare > 0 Then
					.OneyrTargetEarningGrow = CSng(Math.Log(.OneyrTargetEarning / .EarningsShare))
				Else
					.OneyrTargetEarningGrow = 0
				End If
			End If
			If .OneyrPEG = 0 Then
				.OneyrPEG = MyPriceVolLast.OneyrPEG
			End If
			.FiveyrPEG = .OneyrPEG
			.DividendYield = Record.DividendYield
			.DividendShare = Record.DividendShare
			.DividendPayDate = Record.DividendPayDate
			.ExDividendDate = Record.ExDividendDate
			If .ExDividendDate > ReportDate.DateNullValue Then
				If MyPriceVolLast.ExDividendDate > ReportDate.DateNullValue Then
					If .ExDividendDate <> MyPriceVolLast.ExDividendDate Then
						If .ExDividendDate > MyPriceVolLast.ExDividendDate Then
							'this is a new ex dividend date
							.ExDividendDatePrevious = MyPriceVolLast.ExDividendDate
							.ExDividendDateEstimated = .ExDividendDate.AddDays(.ExDividendDate.Subtract(.ExDividendDatePrevious).Days)
						Else
							'the ex-dividend can go back down in time due to yahoo
							.ExDividendDate = MyPriceVolLast.ExDividendDate
							.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
							.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
						End If
					Else
						.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
						.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
					End If
				Else
					.ExDividendDatePrevious = .ExDividendDate
					.ExDividendDateEstimated = .ExDividendDate
				End If
			Else
				If MyPriceVolLast.ExDividendDate >= ReportDate.DateNullValue Then
					'this is an invalid update due to yahoo poor management of ex-dividend
					'in this case all the data on the dividend is false
					'a correction is possible using the the previous data
					.DividendYield = MyPriceVolLast.DividendYield
					.DividendShare = MyPriceVolLast.DividendShare
					.DividendPayDate = MyPriceVolLast.DividendPayDate
					.ExDividendDate = MyPriceVolLast.ExDividendDate
					.ExDividendDatePrevious = MyPriceVolLast.ExDividendDatePrevious
					.ExDividendDateEstimated = MyPriceVolLast.ExDividendDateEstimated
				Else
					.ExDividendDate = ReportDate.DateNullValue
					.ExDividendDatePrevious = ReportDate.DateNullValue
					.ExDividendDateEstimated = ReportDate.DateNullValue
				End If
			End If
			.Vol = Record.Vol
			.Volume = Record.Volume
			If (.Volume = 0 And .Last = 0) Then
				.Open = MyPriceVolLast.Last
				.High = .Open
				.Low = .Open
				.Last = .Open
				.IsNull = True
			Else
				.IsNull = False
			End If
			.LastWeighted = RecordPrices.CalculateLastWeighted(DirectCast(ThisPriceVol, IPriceVol))
			.Range = RecordPrices.CalculateTrueRange(DirectCast(ThisPriceVol, IPriceVol))
			.RecordQuoteValue = Record
		End With
		Return ThisPriceVol
	End Function
#End Region
#Region "Data Properties and functions"
	Public Function ToListOfStockQuote() As List(Of WebEODData.StockQuote)
		Dim ThisList = New List(Of WebEODData.StockQuote)
		For I = 0 To Me.NumberPoint - 1
			With MyPriceVols(I)
				Dim ThisStockQuote As StockQuote = New StockQuote(
					Symbol:=Symbol,
					DateTime:= .DateLastTrade,
					Open:= .Open,
					High:= .High,
					Low:= .Low,
					Close:= .Last,
					Volume:= .Volume)
				ThisStockQuote.AsIStockSentiment.Count = .AsISentimentIndicator.Count
				ThisStockQuote.AsIStockSentiment.Value = .AsISentimentIndicator.Value
				ThisList.Add(ThisStockQuote)
			End With
		Next
		Return ThisList
	End Function

	Public Function ToListOfStockPriceVol() As List(Of IStockPriceVol)
		Dim ThisList = New List(Of IStockPriceVol)
		For I = 0 To Me.NumberPoint - 1
			ThisList.Add(New StockPriceVol(MyPriceVols(I)))
		Next
		Return ThisList
	End Function

	Public Function ToListOfStockPriceVolGain() As List(Of IStockPriceVol)
		Dim ThisList = New List(Of IStockPriceVol)
		For I = 0 To Me.NumberPoint - 1
			ThisList.Add(New StockPriceVol(MyPriceVols(I)))
		Next
		Return ThisList
	End Function

	Public Function ToListOfPriceGainLog() As List(Of IPriceVol)
		Dim ThisList = New List(Of IPriceVol)

		'take the reference as the first price
		Dim ThisPriceRef As Double = Me.PriceVols(0).Last
		Dim ThisPriceRefOffset As Double = 100.0
		'prevent the problem that will occur if the price is zero
		'take a reference of 100.0. In a way this is a similar aproach to simulating a bond issued a a value of 100.0
		'with the price that can vary mostly on the interest rate	or the incertainty and financial stability of the issuer.
		'In this case since it is a stock the influence to the interest rate is not direct but the incertainty and financial
		'stability and the revenue growth are a major concern

		'Dim ThisPriceVolGain As IPriceVolGain
		For Each PriceVol In MyListOfIPriceVol



		Next


		For I = 0 To Me.NumberPoint - 1

			ThisList.Add(New StockPriceVol(MyPriceVols(I)))
		Next
		Return ThisList
	End Function

	Public Function ToListOfPriceVol() As List(Of IPriceVol)
		Return MyListOfIPriceVol
	End Function

	Public Function ToListOfPriceVolIndexed() As IEnumerable(Of (Index As Integer, Item As IPriceVol))
		Return MyListOfIPriceVol.WithIndex
	End Function

	Public Function ToWeeklyIndex(ByVal DateValue As Date) As Integer
		Dim ThisDateRefToPreviousMonday = ReportDate.DateToMondayPrevious(Me.DateStart)

		Dim ThisDeltaDays = ReportDate.MarketTradingDeltaDays(ThisDateRefToPreviousMonday, DateValue)
		Return ThisDeltaDays \ 5
	End Function

	Public Function ToIndex(ByVal DateValue As Date) As Integer
		Return ReportDate.MarketTradingDeltaDays(Me.DateStart, DateValue)
	End Function

	Public Shared Function ToDate(ByVal DateStart As Date, ByVal Index As Integer) As Date
		Dim ThisNumberOfTradingFullWeek As Integer = Index \ 5
		Dim ThisNumberOfTradingDayLeft As Integer = Index Mod 5
		Return DateStart.AddDays(7 * ThisNumberOfTradingFullWeek).AddDays(ThisNumberOfTradingDayLeft)
	End Function

	Public Function ToDate(ByVal Index As Integer) As Date
		Return RecordPrices.ToDate(Me.DateStart, Index)
	End Function

	Public Property PriceVolLast As PriceVol
		Get
			Return DirectCast(MyListOfIPriceVol.Last, PriceVol)
		End Get
		Set(value As PriceVol)
			Throw New NotSupportedException
		End Set
	End Property

	Public Function PriceVols(ByVal Index As Integer) As PriceVol
		If MyListOfIPriceVol.Count = 0 Then Return Nothing
		If Index >= 0 Then
			If Index < MyListOfIPriceVol.Count Then
				Return DirectCast(MyListOfIPriceVol.Item(Index), PriceVol)
			Else
				Return DirectCast(MyListOfIPriceVol.Last, PriceVol)
			End If
		Else
			Return DirectCast(MyListOfIPriceVol.Item(0), PriceVol)
		End If
	End Function

	Public Function GetPriceVolInterface(ByVal Index As Integer) As IPriceVol
		If MyListOfIPriceVol.Count = 0 Then Return Nothing
		If Index >= 0 Then
			If Index < MyListOfIPriceVol.Count Then
				Return MyListOfIPriceVol.Item(Index)
			Else
				Return MyListOfIPriceVol.Last
			End If
		Else
			Return MyListOfIPriceVol.Item(0)
		End If
	End Function

	''' <summary>
	''' Iteration on IPriceVol
	''' </summary>
	''' <returns></returns>
	Public Iterator Function GetPriceVolItems() As IEnumerable(Of IPriceVol)
		For i As Integer = 0 To MyPriceVols.Length - 1
			Yield GetPriceVolInterface(i)
		Next
	End Function

	''' <summary>
	''' iteration on IPriceVol with an index
	''' </summary>
	''' <returns></returns>
	Public Iterator Function GetPriceVolItemsWithIndex() As IEnumerable(Of (Index As Integer, Item As IPriceVol))
		For i As Integer = 0 To MyPriceVols.Length - 1
			Yield (i, GetPriceVolInterface(i))
		Next
	End Function

	Public Sub PriceVolsMultiPly(ByVal Ratio As Single)
		Me.PriceMin = Single.MaxValue
		Me.PriceMax = Single.MinValue
		Me.VolMax = Integer.MinValue
		Me.VolMin = Integer.MaxValue
		For I = Me.StartPoint To Me.StopPoint
			With MyPriceVols(I)
				.MultiPly(Ratio)
				If .Low < Me.PriceMin Then
					Me.PriceMin = .Low
				End If
				If .High > Me.PriceMax Then
					Me.PriceMax = .High
				End If
				If .Volume < Me.VolMin Then
					Me.VolMin = .Volume
				End If
				If .Volume > Me.VolMax Then
					Me.VolMax = .Volume
				End If
			End With
		Next
		MyPriceVolLast = MyPriceVols(Me.StopPoint)
	End Sub

	Public Sub PriceVolsAdd(ByVal RecordPrices As YahooAccessData.RecordPrices)
		Dim I As Integer

		'the number of point need to be the same to be valid

		If RecordPrices.NumberPoint <> Me.NumberPoint Then
			Throw New InvalidDataException(String.Format("Invalid number of point for {0} price addition...", RecordPrices.Symbol))
		End If
		Me.PriceMin = Single.MaxValue
		Me.PriceMax = Single.MinValue
		Me.VolMax = Integer.MinValue
		Me.VolMin = Integer.MaxValue
		For I = Me.StartPoint To Me.StopPoint
			With MyPriceVols(I)
				.Add(RecordPrices.PriceVols(I))
				If .Low < Me.PriceMin Then
					Me.PriceMin = .Low
				End If
				If .High > Me.PriceMax Then
					Me.PriceMax = .High
				End If
				If .Volume < Me.VolMin Then
					Me.VolMin = .Volume
				End If
				If .Volume > Me.VolMax Then
					Me.VolMax = .Volume
				End If
			End With
		Next
		MyPriceVolLast = MyPriceVols(Me.StopPoint)
	End Sub

	Private Sub PriceVolsMultiPly(ByVal Index As Integer, ByVal Ratio As Single)
		MyPriceVols(Index).MultiPly(Ratio)
	End Sub

	Private Sub PriceVolsAdd(ByVal Index As Integer, ByRef PriceVol As PriceVol)
		MyPriceVols(Index).Add(PriceVol)
	End Sub

	Private Function PriceVolsIntraDay(ByVal Index As Integer) As PriceVol()
		Return MyPriceVolsIntraDay(Index)
	End Function

	Public Function PriceVolsData() As PriceVol()
		Return MyPriceVols
	End Function

	Private Function PriceVolsDataIntraDay() As PriceVol()()
		Return MyPriceVolsIntraDay
	End Function


	Private ReadOnly Property IsIntraDayEnabled As Boolean
		Get
			Return IsIntraDayLocalEnabled
		End Get
	End Property

	Public SplitIndex() As Integer
	Private StockDividendSinglePayoutIndex() As Integer 'note not public for now
	Private IsStockDividendSinglePayout As Boolean    'note not public for now
	Public IsSplit As Boolean
	Public PriceMin As Single
	Public PriceMax As Single
	Public PriceMinTarget As Single
	Public PriceMaxTarget As Single
	Public IsPriceTarget As Boolean
	Public PriceToEarningTarget As Double
	Public IsVol As Boolean

	Private _VolMin As Long
	Public Property VolMin As Long
		Get
			Return _VolMin
		End Get
		Set(value As Long)
			_VolMin = value
		End Set
	End Property

	Private _VolMax As Long
	Public Property VolMax As Long
		Get
			'note do not retrun a value egual to zero if not the graph may have a problem
			'in some zooming scenario
			If _VolMax = 0 Then Return 1
			Return _VolMax
		End Get
		Set(value As Long)
			_VolMax = value
		End Set
	End Property

	Public DateStart As Date
	Public DateStop As Date
	Public NumberPoint As Integer
	Public NumberNullPoint As Integer
	Public NumberNullPointToEnd As Integer
	Public StartPoint As Integer
	Public StopPoint As Integer
	Public IsError As Boolean
	Public ErrorDescription As String
	Public Symbol As String
	Public IsNull As Boolean
	Public Tag As String
	Public Property Stock As Stock
#End Region
End Class
