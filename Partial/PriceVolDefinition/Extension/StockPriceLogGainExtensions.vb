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
			Dim ThisDataResultItem = New StockPriceVol(SourceItem)
			'The data input is expected to be clean but just in case, we try to compute log gain,
			'but if any error occurs (e.g. division by zero, log of negative), return 0.0 for that bar and continue
			'the value of zero should noy affect the rest of calculatyion excepy for that point, and it is better than throwing
			'an exception and breaking the whole series. This may happen for example if there are bad data points
			'with zero or negative price, which would cause log gain to be undefined.
			'By returning 0.0 for that point, we can still compute the rest of the series without interruption and hopefully the
			'bad data has not too much impact on the overall analysis.
			'This is a pragmatic approach to handle data quality issues while still providing useful results.
			With ThisDataResultItem
				.DataType = StockPriceDataType.CumulativeLogReturn
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
	''' sharpe ratio measurement or other yearly statistic. If needed the data will be transformed in the correct format
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
			Dim resultItem As New StockPriceVol(DataSource(i))
			With resultItem
				.DataType = StockPriceDataType.CumulativeLogYearlyReturn
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
			PriceNormalized As Double,
			Optional Index As Integer = 0) As List(Of StockPriceVol)


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
		If PriceNormalized <= 0 Then
			Throw New ArgumentException(
				"Price must be greater than 0 for valid stock prices.", paramName:=NameOf(PriceNormalized))
		End If
		' Log-domain OHLC are cumulative coordinates:
		' C_open  = C_prevClose + ln(Open / PrevClose)
		' C_high  = C_prevClose + ln(High / PrevClose)
		' C_low   = C_prevClose + ln(Low  / PrevClose)
		' C_close = C_prevClose + ln(Close/ PrevClose)
		'Conversion back to a price-like level is done as follow:
		'PriceResult = Price * Exp(C(i) - C(Index))
		Dim ThisPriceLogAtIndex As Double = DataSource(Index).Last
		Dim ThisSourceItemLast As StockPriceVol = DataSource.First
		Dim ThisDataResultLast As StockPriceVol = Nothing
		For Each SourceItem In DataSource
			Dim ThisDataResultItem = New StockPriceVol(SourceItem)
			With ThisDataResultItem
				.DataType = StockPriceDataType.RawPrice
				.Open = PriceNormalized * Math.Exp(SourceItem.Open - ThisPriceLogAtIndex)
				.High = PriceNormalized * Math.Exp(SourceItem.High - ThisPriceLogAtIndex)
				.Low = PriceNormalized * Math.Exp(SourceItem.Low - ThisPriceLogAtIndex)
				.Last = PriceNormalized * Math.Exp(SourceItem.Last - ThisPriceLogAtIndex)
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



	''' <summary>
	''' Adds two cumulative log gain series together, effectively compounding their returns.
	''' Both series must be in cumulative log gain space and have the same number of bars.
	''' The resulting series will also be in cumulative log gain space, representing the combined effect of both inputs.
	''' This can be used to combine the log returns of two assets or to apply a log return adjustment to an existing series.
	''' </summary>
	''' <param name="source1"></param>
	''' <param name="source2"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainAdd(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		Dim DataSource2 = source2.ToList()
		'check that both data sources have the same number of bars, otherwise we cannot add them together
		If DataSource1.Count <> DataSource2.Count Then
			Throw New ArgumentException("Both data sources must have the same number of bars to be added together.")
		End If
		'also check that both data sources are not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice OrElse
			DataSource2.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItems In DataSource1.Zip(DataSource2, Function(s1, s2) (s1, s2))
			'we could check for teh time stamp but the operational requirement for teh array guarantee that teh data data date are aligned corrcetly.
			'for that reason the test is superfluous but we can keep it just in case to catch any potential data issue, and it is not too costly
			'since we are already iterating through the data.
			'check that the timestamps of both data sources match for the current bar, otherwise we cannot add them together
			If SourceItems.s1.DateDay <> SourceItems.s2.DateDay Then
				Throw New ArgumentException("Timestamps of both data sources must match for each bar to be added together.")
			End If

			Dim ThisDataResultItem = New StockPriceVol(SourceItems.s1)

			With ThisDataResultItem
				.Open = .Open + SourceItems.s2.Open
				.High = .High + SourceItems.s2.High
				.Low = .Low + SourceItems.s2.Low
				.Last = .Last + SourceItems.s2.Last
				.OpenNext = .OpenNext + SourceItems.s2.OpenNext
				.LastPrevious = .LastPrevious + SourceItems.s2.LastPrevious
				.Volume = .Volume + SourceItems.s2.Volume
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	<Extension>
	Public Function ToCumulativeLogGainAdd(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As Double) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		'check that the data source is not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItem In DataSource1
			Dim ThisDataResultItem = New StockPriceVol(SourceItem)

			With ThisDataResultItem
				.Open = .Open + source2
				.High = .High + source2
				.Low = .Low + source2
				.Last = .Last + source2
				.OpenNext = .OpenNext + source2
				.LastPrevious = .LastPrevious + source2
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	<Extension>
	Public Function ToCumulativeLogGainAdd(
			source1 As Double,
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)

		Return ToCumulativeLogGainAdd(source2, source1)
	End Function

	''' <summary>
	''' Subtract two cumulative log gain series, effectively reversing the compounding effect of their returns.
	''' Both series must be in cumulative log gain space and have the same number of bars.
	''' The resulting series will also be in cumulative log gain space, representing the combined effect of both inputs.
	''' This can be used to reverse the log returns of two assets or to apply a log return adjustment to an existing series.
	''' </summary>
	''' <param name="source1"></param>
	''' <param name="source2"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainSubtract(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		'for maintenance it might be easier to just call the add function with the negative value,
		'but we prefer here to keep the code to avoid having to create a new list with the negative value of source2,
		'which would be an unnecessary copy and could be costly if the data is large.
		'By keeping the code here we can directly subtract the values without having to create a new list,
		'which is more efficient and avoids unnecessary memory allocation.

		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		Dim DataSource2 = source2.ToList()
		'check that both data sources have the same number of bars, otherwise we cannot add them together
		If DataSource1.Count <> DataSource2.Count Then
			Throw New ArgumentException("Both data sources must have the same number of bars to be added together.")
		End If
		'also check that both data sources are not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot divide them
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice OrElse
			DataSource2.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItems In DataSource1.Zip(DataSource2, Function(s1, s2) (s1, s2))
			'we could check for teh time stamp but the operational requirement for teh array garantee thta teh data data date are aligned corrcetly.
			'for that reason the test is superfluous but we can keep it just in case to catch any potential data issue, and it is not too costly
			'since we are already iterating through the data.
			'check that the timestamps of both data sources match for the current bar, otherwise we cannot add them together
			If SourceItems.s1.DateDay <> SourceItems.s2.DateDay Then
				Throw New ArgumentException("Timestamps of both data sources must match for each bar to be added together.")
			End If

			Dim ThisDataResultItem = New StockPriceVol(SourceItems.s1)

			With ThisDataResultItem
				.Open = .Open - SourceItems.s2.Open
				.High = .High - SourceItems.s2.High
				.Low = .Low - SourceItems.s2.Low
				.Last = .Last - SourceItems.s2.Last
				.OpenNext = .OpenNext - SourceItems.s2.OpenNext
				.LastPrevious = .LastPrevious - SourceItems.s2.LastPrevious
				'ìt make no sense to subtract volume and since negative volume for stock even in log gain is not reasonably acceptable
				'because it would have no meaning
				'in this case we keep adding volume after all th volume represent more the number of contract invole here
				.Volume = .Volume + SourceItems.s2.Volume
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	<Extension>
	Public Function ToCumulativeLogGainSubtract(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As Double) As List(Of StockPriceVol)

		Return ToCumulativeLogGainAdd(source1, -source2)
	End Function

	<Extension>
	Public Function ToCumulativeLogGainSubstract(
			source1 As Double,
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource2 = source2.ToList()
		'check that the data source is not empty, otherwise we cannot add them together
		If DataSource2.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource2.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Subtracting is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource2.Count)

		For Each SourceItem In DataSource2
			Dim ThisDataResultItem = New StockPriceVol(SourceItem)

			With ThisDataResultItem
				.Open = source1 - .Open
				.High = source1 - .High
				.Low = source1 - .Low
				.Last = source1 - .Last
				.OpenNext = source1 - .OpenNext
				.LastPrevious = source1 - .LastPrevious
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	''' <summary>
	''' Multiply two cumulative log gain series, effectively combining the compounding effect of their returns.
	''' Both series must be in cumulative log gain space and have the same number of bars.
	''' The resulting series will also be in cumulative log gain space, representing the combined effect of both inputs.
	''' This can be used to apply a log return adjustment to an existing series or to combine the log returns of two assets.
	''' </summary>
	''' <param name="source1"></param>
	''' <param name="source2"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainMultiply(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		Dim DataSource2 = source2.ToList()
		'check that both data sources have the same number of bars, otherwise we cannot add them together
		If DataSource1.Count <> DataSource2.Count Then
			Throw New ArgumentException("Both data sources must have the same number of bars to be added together.")
		End If
		'also check that both data sources are not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice OrElse
			DataSource2.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItems In DataSource1.Zip(DataSource2, Function(s1, s2) (s1, s2))
			'we could check for teh time stamp but the operational requirement for teh array garantee thta teh data data date are aligned corrcetly.
			'for that reason the test is superfluous but we can keep it just in case to catch any potential data issue, and it is not too costly
			'since we are already iterating through the data.
			'check that the timestamps of both data sources match for the current bar, otherwise we cannot add them together
			If SourceItems.s1.DateDay <> SourceItems.s2.DateDay Then
				Throw New ArgumentException("Timestamps of both data sources must match for each bar to be added together.")
			End If

			Dim ThisDataResultItem = New StockPriceVol(SourceItems.s1)

			With ThisDataResultItem
				.Open = .Open * SourceItems.s2.Open
				.High = .High * SourceItems.s2.High
				.Low = .Low * SourceItems.s2.Low
				.Last = .Last * SourceItems.s2.Last
				.OpenNext = .OpenNext * SourceItems.s2.OpenNext
				.LastPrevious = .LastPrevious * SourceItems.s2.LastPrevious
				'keep adding volume after all the volume represent more the number of contract involved in the mathematical process
				.Volume = .Volume + SourceItems.s2.Volume
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	<Extension>
	Public Function ToCumulativeLogGainMultiply(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As Double) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		'check that the data source is not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItem In DataSource1
			Dim ThisDataResultItem = New StockPriceVol(SourceItem)

			With ThisDataResultItem
				.Open = .Open * source2
				.High = .High * source2
				.Low = .Low * source2
				.Last = .Last * source2
				.OpenNext = .OpenNext * source2
				.LastPrevious = .LastPrevious * source2
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	<Extension>
	Public Function ToCumulativeLogGainMultiply(
			source1 As Double,
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)
		Return ToCumulativeLogGainMultiply(source2, source1)
	End Function

	''' <summary>
	''' Divide two cumulative log gain series, effectively combining the compounding effect of their returns.
	''' Both series must be in cumulative log gain space and have the same number of bars.
	''' The resulting series will also be in cumulative log gain space, representing the combined effect of both inputs.
	''' This can be used to apply a log return adjustment to an existing series or to combine the log returns of two assets.
	''' </summary>
	''' <param name="source1"></param>
	''' <param name="source2"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainDivide(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As IEnumerable(Of StockPriceVol)) As List(Of StockPriceVol)


		' IMPORTANT:
		' We intentionally materialize the sequence into a List here.
		' This guarantees:
		'   - single, deterministic enumeration
		'   - stable indexing for anchor/position validation
		'   - safe behavior if the source is a LINQ iterator or generator
		'
		' Do NOT "optimize" this away unless profiling proves it is a bottleneck.
		' The copy cost is negligible compared to the safety it provides.
		Dim DataSource1 = source1.ToList()
		Dim DataSource2 = source2.ToList()
		'check that both data sources have the same number of bars, otherwise we cannot add them together
		If DataSource1.Count <> DataSource2.Count Then
			Throw New ArgumentException("Both data sources must have the same number of bars to be added together.")
		End If
		'also check that both data sources are not empty, otherwise we cannot add them together
		If DataSource1.Count = 0 Then
			Throw New ArgumentException("Data sources cannot be empty.")
		End If
		' Idempotent guard: already in log space
		'both data sources must be in cumulative log return space, otherwise we cannot add them together
		If _
			DataSource1.First.DataType = StockPriceDataType.RawPrice OrElse
			DataSource2.First.DataType = StockPriceDataType.RawPrice Then
			'already in log space
			Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
		End If
		Dim DataResult As New List(Of StockPriceVol)(DataSource1.Count)

		For Each SourceItems In DataSource1.Zip(DataSource2, Function(s1, s2) (s1, s2))
			'we could check for teh time stamp but the operational requirement for teh array garantee thta teh data data date are aligned corrcetly.
			'for that reason the test is superflous but we can keep it just in case to catch any potential data issue, and it is not too costly
			'since we are already iterating through the data.
			'check that the timestamps of both data sources match for the current bar, otherwise we cannot add them together
			If SourceItems.s1.DateDay <> SourceItems.s2.DateDay Then
				Throw New ArgumentException("Timestamps of both data sources must match for each bar to be added together.")
			End If

			Dim ThisDataResultItem = New StockPriceVol(SourceItems.s1)

			With ThisDataResultItem
				.Open = .Open / SourceItems.s2.Open
				.High = .High / SourceItems.s2.High
				.Low = .Low / SourceItems.s2.Low
				.Last = .Last / SourceItems.s2.Last
				.OpenNext = .OpenNext / SourceItems.s2.OpenNext
				.LastPrevious = .LastPrevious / SourceItems.s2.LastPrevious
				'keep adding volume after all the volume represent more the number of contract involved in the mathematical process
				.Volume = .Volume + SourceItems.s2.Volume
				DataResult.Add(ThisDataResultItem)
			End With
		Next
		Return DataResult
	End Function

	''' <summary>
	''' Divide a cumulative log gain series by a scalar value, effectively adjusting the series by the inverse of the scalar.
	''' The series must be in cumulative log gain space.
	''' The resulting series will also be in cumulative log gain space, representing the adjusted effect of the input series.
	''' This can be used to apply a log return adjustment to an existing series.
	''' </summary>
	''' <param name="source1"></param>
	''' <param name="source2"></param>
	''' <returns></returns>
	<Extension>
	Public Function ToCumulativeLogGainDivide(
			source1 As IEnumerable(Of StockPriceVol),
			source2 As Double) As List(Of StockPriceVol)

		'it is easier here to just multiply by the inverse of the source2 value rather than doing a division
		If source2 = 0 Then
			Throw New ArgumentException("Division by zero is not allowed.")
		End If
		Dim inverseSource2 = 1.0 / source2
		Return ToCumulativeLogGainMultiply(source1, inverseSource2)
	End Function
End Module
