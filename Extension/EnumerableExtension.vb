Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Namespace ExtensionService
	Public Module EnumerableExtension
		''' <summary>
		''' Enumerates a sequence, yielding each item along with its zero-based index.
		''' Equivalent to Python's enumerate().
		''' </summary>
		<Extension>
		Public Iterator Function WithIndex(Of T)(source As IEnumerable(Of T)) As IEnumerable(Of (Index As Integer, Item As T))
			Dim i As Integer = 0
			For Each SourceItem In source
				Yield (Index:=i, Item:=SourceItem)
				i += 1
			Next
		End Function

		''' <summary>
		''' Enumerates a sequence, yielding each item along with its index starting from a custom offset.
		''' usage is rare, but can be useful in some scenarios.
		''' Equivalent to Python's enumerate() with a custom start index.
		''' </summary>
		<Extension>
		Public Iterator Function WithIndexFrom(Of T)(source As IEnumerable(Of T), startIndex As Integer) As IEnumerable(Of (Index As Integer, Item As T))
			Dim i As Integer = startIndex
			For Each SourceItem In source
				Yield (Index:=i, Item:=SourceItem)
				i += 1
			Next
		End Function

		''' <summary>
		''' Min–max scales a sequence of Double to [newMin, newMax].
		''' Empty sequence → empty list. If all inputs are identical (min=max),
		''' maps everything to the midpoint of the target range.
		''' If newMin=newMax, returns a constant list of newMin.
		''' </summary>
		<Extension>
		Public Function ScaleToRange(values As IEnumerable(Of Double),
																 newMin As Double,
																 newMax As Double) As List(Of Double)
			If values Is Nothing Then Throw New ArgumentNullException(NameOf(values))

			' Materialize once to avoid multiple enumerations
			Dim arr = TryCast(values, IList(Of Double))
			If arr Is Nothing Then arr = values.ToList()

			If arr.Count = 0 Then Return New List(Of Double)()

			' Target is a single point → constant result
			If newMin = newMax Then
				Return Enumerable.Repeat(newMin, arr.Count).ToList()
			End If

			' Compute min/max in one pass
			Dim minVal = arr(0)
			Dim maxVal = arr(0)
			For i = 1 To arr.Count - 1
				Dim v = arr(i)
				If v < minVal Then minVal = v
				If v > maxVal Then maxVal = v
			Next

			' All inputs identical → map to midpoint of target range
			If maxVal = minVal Then
				Dim mid = (newMin + newMax) / 2.0
				Return Enumerable.Repeat(mid, arr.Count).ToList()
			End If

			Dim scale = (newMax - newMin) / (maxVal - minVal)
			Dim result = New List(Of Double)(arr.Count)
			For Each v In arr
				result.Add(newMin + (v - minVal) * scale)
			Next
			Return result
		End Function

		' Optional generic overload if you sometimes scale a property of objects:
		''' <summary>
		''' Scales a numeric projection of a sequence to [newMin, newMax].
		'''  Example usage:
		''' 2) Map values to pixel space for drawing (e.g., y-coordinates)
		'''
		''' Dim chartHeight As Integer = 400
		''' Dim yMin As Double = 0
		''' Dim yMax As Double = chartHeight - 1
		'''
		''' Suppose bars is a List(Of PriceVol) where Last represents the price
		''' Dim ys As List(Of Double) =
		'''     bars.ScaleToRange(Function(b) b.Last, yMin, yMax)
		'''
		''' ' → ys now contains Y positions in pixel coordinates aligned with your chart height.
		''' '   Useful for mapping price values directly to drawing coordinates on a chart.
		''' 
		''' Or also: 
		''' 
		''' Dim ys = bars.ScaleToRange(Function(b) b.Last, 0, chartHeight - 1)
		''' </summary>
		<Extension>
		<EditorBrowsable(EditorBrowsableState.Advanced)>
		Public Function ScaleToRange(Of T)(source As IEnumerable(Of T),
																			 selector As Func(Of T, Double),
																			 newMin As Double,
																			 newMax As Double) As List(Of Double)
			If source Is Nothing Then Throw New ArgumentNullException(NameOf(source))
			If selector Is Nothing Then Throw New ArgumentNullException(NameOf(selector))
			Return source.Select(selector).ScaleToRange(newMin, newMax)
		End Function
	End Module
End Namespace
