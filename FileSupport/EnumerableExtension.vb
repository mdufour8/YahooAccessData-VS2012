Imports System.Runtime.CompilerServices
Imports System.Collections.Generic

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
		public Iterator Function WithIndexFrom(Of T)(source As IEnumerable(Of T), startIndex As Integer) As IEnumerable(Of (Index As Integer, Item As T))
			Dim i As Integer = startIndex
			For Each SourceItem In source
				Yield (Index:=i, Item:=SourceItem)
				i += 1
			Next
		End Function
	End Module
End Namespace
