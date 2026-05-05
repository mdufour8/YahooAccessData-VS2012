Imports System.Runtime.CompilerServices
Imports YahooAccessData.StockPriceVol
Imports YahooAccessData.MathPlus.Measure

Public Module StockPriceLogGainExtensions

	''' <summary>
	''' Converts a raw price series to cumulative log-gain space.
	''' Optionally rebases the series so that the bar at anchorIndex has value = startValue.
	'''
	''' Notes:
	''' - If the first element is already LogCumulative, returns the series as-is (idempotent).
	''' - The first bar defines the anchor (gain = 0, value = startValue), so no artificial jump exists.
	''' 
	''' LOG / LEVEL DOMAIN CONVENTION
	'''
	''' Internal computations are done in LOG space:
	'''   r(t) = ln( P(t) / P(t-1) )        ' log return
	'''   C(t) = Σ r(t)                     ' cumulative log return (no offset)
	'''
	''' C(t) may be negative and is used for:
	'''   - volatility
	'''   - correlation
	'''   - averaging and weighting across assets
	'''   - log-normal probability modeling
	'''
	''' Conversion back to a price-like level is done ONLY at the boundary:
	'''   Index(t) = BaseIndex * Exp( C(t) - C(anchor) )
	'''
	''' BaseIndex (e.g. 100) is a presentation choice, NOT part of log-space math.
	''' Never apply offsets or Exp() inside log-space calculations.
	''' ---------------------------------------------------------------------------
	''' </summary>
	<Extension>
	Public Function ToCumulativeLogGain(
			source As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource = source.ToList()
		If DataSource.Count = 0 Then Return New List(Of StockPriceVol)
		' Idempotent guard: already in log space
		If DataSource.First.DataType <> StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Rebase as LogCumulative require raw price data source.")
		End If

		Dim DataResult As New List(Of StockPriceVol)(DataSource.Count)

		' Log-domain OHLC are cumulative coordinates:
		' C_open  = C_prevClose + ln(Open / PrevClose)
		' C_high  = C_prevClose + ln(High / PrevClose)
		' C_low   = C_prevClose + ln(Low  / PrevClose)
		' C_close = C_prevClose + ln(Close/ PrevClose)
		Dim ThisSourceItemLast As StockPriceVol = DataSource.First
		Dim ThisDataResultLast As StockPriceVol = New StockPriceVol
		For Each SourceItem In DataSource
			Dim ThisDataResultItem = New StockPriceVol(SourceItem, DataType:=StockPriceDataType.CumulativeLogReturn)
			'The data input is expected to be clean but just in case, we try to compute log gain,
			'but if any error occurs (e.g. division by zero, log of negative), return 0.0 for that bar and continue
			'the value of zero should noy affect the rest of calculatyion excepy for that point, and it is better than throwing
			'an exception and breaking the whole series. This may happen for example if there are bad data points
			'with zero or negative price, which would cause log gain to be undefined.
			'By returning 0.0 for that point, we can still compute the rest of the series without interruption and hopefully the
			'bad data has not too much impact on the overall analysis.
			'This is a pragmatic approach to handle data quality issues while still providing useful results.
			With ThisDataResultItem
				.Open = Measure.GainLog(SourceItem.Open, ThisSourceItemLast.Last, OnErrorReturn:=0.0) + ThisDataResultLast.Last
				.High = Measure.GainLog(SourceItem.High, ThisSourceItemLast.Last, OnErrorReturn:=0.0) + ThisDataResultLast.Last
				.Low = Measure.GainLog(SourceItem.Low, ThisSourceItemLast.Last, OnErrorReturn:=0.0) + ThisDataResultLast.Last
				.Last = Measure.GainLog(SourceItem.Last, ThisSourceItemLast.Last, OnErrorReturn:=0.0) + ThisDataResultLast.Last
				'Open next is not yet known but by default we can set it to be the same as Last
				'until we see the next bar and can update it, this way we avoid having a gap
				'in the data and we will have a valid value for OpenNext for all bars,
				'even if it is not the exact value until we update it with the next bar's Open
				.OpenNext = .Last
				If DataResult.Count = 0 Then
					'first bar, we can set LastPrevious to be the same as Last since there is no previous bar
					.LastPrevious = .Open
				Else
					.LastPrevious = ThisDataResultLast.Last
					'update the previous bar's OpenNext to be the current bar's Open, which is
					'the correct value for the next bar's OpenNext
					ThisDataResultLast.OpenNext = .Open
				End If
				'add the actual log gain item to the list after we have updated the previous bar's OpenNext, so that the list always has the correct values for all bars
				DataResult.Add(ThisDataResultItem)
			End With
			're-assign the "last" variables for the next iteration
			ThisSourceItemLast = SourceItem
			ThisDataResultLast = ThisDataResultItem
		Next
		Return DataResult
	End Function


	''' <summary>
	''' Similar to the ToCumulativeLogGain but assumes the input is already in log space 
	''' and just re-computes the cumulative coordinates to the last year period. Can be useful for calculating the 
	''' sharpe ratio mesurement or other yearly statistic. If needed the data will be transform in the correct format
	''' in the function, so the user can call this function directly on raw price data without having to 
	''' call ToCumulativeLogGain first, but if the data is already in log space, it will just re-use it and avoid 
	''' unnecessary conversion. 
	''' The function will return a new list of StockPriceVol with the same number of items as the input, 
	''' but with the Open, High, Low, Last, OpenNext and LastPrevious fields adjusted to represent the 
	''' cumulative log gain relative to the last year period. The function will handle edge cases such as when 
	''' there are less than one year of data by using the earliest available bar as the anchor for the gain calculation.
	''' </summary>
	''' <param name="source"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainYearly(
				source As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)

		' Defensive: handle null
		If source Is Nothing Then
			Return New List(Of StockPriceVol)
		End If

		' Materialize once (IMPORTANT: avoids multiple enumerations)
		Dim SourceList As List(Of StockPriceVol) = source.ToList()

		If SourceList.Count = 0 Then
			Return New List(Of StockPriceVol)
		End If

		Dim DataSource As List(Of StockPriceVol)

		' Idempotent guard: ensure data is in cumulative log space
		Select Case SourceList(0).DataType
			Case StockPriceVol.StockPriceDataType.RawPrice
				' Convert to cumulative log return
				DataSource = StockPriceLogGainExtensions.ToCumulativeLogGain(SourceList)
			Case StockPriceVol.StockPriceDataType.CumulativeLogReturn
				' Already correct format
				DataSource = SourceList
			Case StockPriceVol.StockPriceDataType.CumulativeLogYearlyReturn
				' Already computed → return copy
				Return SourceList.ToList()

			Case Else
				Throw New NotSupportedException("Unsupported StockPriceDataType.")

		End Select

		' Prepare result container
		Dim DataResult As New List(Of StockPriceVol)(DataSource.Count)
		For i As Integer = 0 To DataSource.Count - 1
			' Position one trading year back
			Dim positionLastYear As Integer = Math.Max(0, i - MathPlus.NUMBER_TRADINGDAY_PER_YEAR)

			' Anchor: yearly log gain relative to prior year's CLOSE (Last)
			Dim gainLastYear As Double = DataSource(positionLastYear).Last
			Dim sourceItem = DataSource(i)
			Dim resultItem As New StockPriceVol(
						sourceItem,
						DataType:=StockPriceVol.StockPriceDataType.CumulativeLogYearlyReturn)

			With resultItem
				.Open -= gainLastYear
				.High -= gainLastYear
				.Low -= gainLastYear
				.Last -= gainLastYear
				.OpenNext -= gainLastYear
				.LastPrevious -= gainLastYear
			End With
			DataResult.Add(resultItem)
		Next
		Return DataResult
	End Function

	''' <summary>
	''' 
	''' ' ---------------------------------------------------------------------------
	''' LOG / LEVEL DOMAIN CONVENTION
	'''
	''' Internal computations are done in LOG space:
	'''   r(t) = ln( P(t) / P(t-1) )        ' log return
	'''   C(t) = Σ r(t)                     ' cumulative log return (no offset)
	'''
	''' C(t) may be negative and is used for:
	'''   - volatility
	'''   - correlation
	'''   - averaging and weighting across assets
	'''   - log-normal probability modeling
	'''
	''' Conversion back to a price-like level is done ONLY at the boundary:
	'''   Index(t) = BaseIndex * Exp( C(t) - C(anchor) )
	'''
	''' PriceBase (e.g. 100) is a presentation choice, NOT part of log-space math.
	''' Never apply offsets or Exp() inside log-space calculations.
	''' ---------------------------------------------------------------------------
	''' </summary>
	<Extension>
	Public Function ToCumulativeLogGainInverse(
			source As IEnumerable(Of StockPriceVol),
			PriceBase As Double,
			Optional PriceBaseIndex As Integer = 0) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource = source.ToList()
		If DataSource.Count = 0 Then Return New List(Of StockPriceVol)
		' Idempotent guard: already in log space
		If DataSource.First.DataType <> StockPriceDataType.CumulativeLogReturn Then
			Throw New InvalidOperationException(
					"CumulativeLogGainInverse requires a CumulativeLogReturn data source. " &
					"Use ToCumulativeLogGain() first to convert raw prices to log space.")
		End If
		'create the list to hold the result, we will populate it with the converted price data
		Dim DataResult As New List(Of StockPriceVol)(DataSource.Count)

		'Validate Price (must be positive for stock prices)
		If PriceBase <= 0 Then
			Throw New ArgumentException(
				"Price must be greater than 0 for valid stock prices.", paramName:=NameOf(PriceBase))
		End If
		' Log-domain OHLC are cumulative coordinates:
		' C_open  = C_prevClose + ln(Open / PrevClose)
		' C_high  = C_prevClose + ln(High / PrevClose)
		' C_low   = C_prevClose + ln(Low  / PrevClose)
		' C_close = C_prevClose + ln(Close/ PrevClose)
		'Conversion back to a price-like level is done as follow:
		'PriceResult = Price * Exp(C(i) - C(Index))
		Dim ThisPriceLogAtIndex As Double = DataSource(PriceBaseIndex).Last
		Dim ThisSourceItemLast As StockPriceVol = DataSource.First
		Dim ThisDataResultLast As StockPriceVol = Nothing
		For Each SourceItem In DataSource
			Dim ThisDataResultItem = New StockPriceVol(SourceItem, DataType:=StockPriceDataType.RawPrice)
			With ThisDataResultItem
				.Open = PriceBase * Math.Exp(SourceItem.Open - ThisPriceLogAtIndex)
				.High = PriceBase * Math.Exp(SourceItem.High - ThisPriceLogAtIndex)
				.Low = PriceBase * Math.Exp(SourceItem.Low - ThisPriceLogAtIndex)
				.Last = PriceBase * Math.Exp(SourceItem.Last - ThisPriceLogAtIndex)
				'Open next is not yet known but by default we can set it to be the same as Last
				'until we see the next bar and can update it, this way we avoid having a gap
				'in the data and we will have a valid value for OpenNext for all bars,
				'even if it is not the exact value until we update it with the next bar's Open
				.OpenNext = .Last
				If DataResult.Count = 0 Then
					'first bar, we can set LastPrevious to be the same as Last since there is no previous bar availaible
					.LastPrevious = .Open
				Else
					'Subsequent bars: LastPrevious is previous bar's Last
					.LastPrevious = ThisDataResultLast.Last
					'update the previous bar's OpenNext to be the current bar's Open, which is
					'the correct value for the next bar's OpenNext
					ThisDataResultLast.OpenNext = .Open
				End If
				DataResult.Add(ThisDataResultItem)
			End With
			're-assign the "last" object to be ready for the next iteration
			ThisDataResultLast = ThisDataResultItem
		Next
		Return DataResult
	End Function
End Module
