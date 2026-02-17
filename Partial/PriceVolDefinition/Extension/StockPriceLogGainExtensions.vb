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
		If DataSource.First.DataType = StockPriceDataType.CumulativeLogReturn Then
			'already in log space
			Return DataSource
		End If

		'Dim DataResult As New List(Of IStockPriceVol)(DataSource.Count)
		Dim DataResult As New List(Of StockPriceVol)

		' Log-domain OHLC are cumulative coordinates:
		' C_open  = C_prevClose + ln(Open / PrevClose)
		' C_high  = C_prevClose + ln(High / PrevClose)
		' C_low   = C_prevClose + ln(Low  / PrevClose)
		' C_close = C_prevClose + ln(Close/ PrevClose)
		Dim ThisSourceItemLast As StockPriceVol = DataSource.First
		Dim ThisDataResultLast As StockPriceVol = New StockPriceVol
		For Each SourceItem In DataSource
			Dim ThisDataResultItem = New StockPriceVol(SourceItem, DataType:=StockPriceDataType.CumulativeLogReturn)
			With ThisDataResultItem
				.Open = Math.Log(SourceItem.Open / ThisSourceItemLast.Last) + ThisDataResultLast.Last
				.High = Math.Log(SourceItem.High / ThisSourceItemLast.Last) + ThisDataResultLast.Last
				.Low = Math.Log(SourceItem.Low / ThisSourceItemLast.Last) + ThisDataResultLast.Last
				.Last = Math.Log(SourceItem.Last / ThisSourceItemLast.Last) + ThisDataResultLast.Last
				.LastPrevious = Math.Log(SourceItem.LastPrevious / ThisSourceItemLast.LastPrevious) + ThisDataResultLast.Last
				DataResult.Add(ThisDataResultItem)
				ThisDataResultLast = ThisDataResultItem
			End With
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
	''' BaseIndex (e.g. 100) is a presentation choice, NOT part of log-space math.
	''' Never apply offsets or Exp() inside log-space calculations.
	''' ---------------------------------------------------------------------------
	''' </summary>
	<Extension>
	Public Function CumulativeLogGainInverse(
			source As IEnumerable(Of StockPriceVol),
			Price As Double,
			Index As Integer) As List(Of StockPriceVol)


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
			Throw New InvalidOperationException("Rebase at price Index requires a LogCumulative data source.")
		End If
		Dim DataResult As New List(Of StockPriceVol)
		' Log-domain OHLC are cumulative coordinates:
		' C_open  = C_prevClose + ln(Open / PrevClose)
		' C_high  = C_prevClose + ln(High / PrevClose)
		' C_low   = C_prevClose + ln(Low  / PrevClose)
		' C_close = C_prevClose + ln(Close/ PrevClose)
		'Conversion back to a price-like level is done as follow:
		'PriceResult = Price * Exp(C(i) - C(Index))
		Dim ThisPriceAtIndex As Double = DataSource(Index).Last
		Dim ThisSourceItemLast As StockPriceVol = DataSource.First
		Dim ThisDataResultLast As StockPriceVol = New StockPriceVol

		For Each SourceItem In DataSource
			Dim ThisDataResultItem = New StockPriceVol(SourceItem, DataType:=StockPriceDataType.RawPrice)
			With ThisDataResultItem
				.Open = Math.Exp(SourceItem.Open - ThisPriceAtIndex)
				.High = Math.Exp(SourceItem.High - ThisPriceAtIndex)
				.Low = Math.Exp(SourceItem.Low - ThisPriceAtIndex)
				.Last = Math.Exp(SourceItem.Last - ThisPriceAtIndex)
				.LastPrevious = Math.Exp(SourceItem.LastPrevious - ThisPriceAtIndex)
				DataResult.Add(ThisDataResultItem)
				ThisDataResultLast = ThisDataResultItem
			End With
		Next
		Return DataResult
	End Function
End Module
