Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports YahooAccessData.MathPlus.Filter


' NOTE: Migrated from struct → class.
' Key things to re-check if debug traces or charts look wrong:
' 1. ItemDecimate: make sure we are creating a *new T()*
'    (so it behaves like a struct copy and doesn’t mutate list items).
' 2. High/Low indices: verify MyItemHighIndex / MyItemLowIndex
'    track the correct elements after Add/Remove.
' 3. Reference vs. copy: with class types, assignment shares references,
'    so any missed New() may cause shared mutations.
' 4. Use WindowFrameDebug traces (Add, State, Decimate, Summary)
'    to quickly catch mismatches against graph outputs.

Public Class ListWindowFrameAsClass(Of T As {Class, New, IPriceVol, IPricePivotPoint})
	Implements IList(Of T)
	'Implements IListWindowsFrame(Of T)
	Implements IListWindowsFrame1(Of T)

	Private MyListOfIPriceVol As List(Of T)
	Private MyListWindowsFrameForPivotOpen As ListWindowFrame
	Private MyFilterHullForDecimateOpen As IFilter
	Private MyVolumeSum As Long
	Private MyPriceSum As Double
	Private MyWindowSize As Integer
	Private MyItemHighIndex As Integer
	Private MyItemLowIndex As Integer
	Private MyItemRemoved As T
	Private IsItemRemoved As Boolean

#Region "New"
	Public Sub New(ByVal WindowSize As Integer)
		MyListOfIPriceVol = New List(Of T)(WindowSize)
		MyListWindowsFrameForPivotOpen = New ListWindowFrame(2 * WindowSize)
		MyFilterHullForDecimateOpen = New FilterLowPassExpHull(WindowSize \ 2)
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
		MyWindowSize = WindowSize
	End Sub
#End Region
#Region "ICollection"
	Public Sub Add(item As T) Implements ICollection(Of T).Add
		Dim I As Integer

		With MyListOfIPriceVol
			If .Count = MyWindowSize Then
				MyItemRemoved = .First
				MyVolumeSum = MyVolumeSum - CLng(.First.Vol)
				MyPriceSum = MyPriceSum - CDbl(.First.Last)
				.RemoveAt(0)
				'adjust the index position due to the item being removed
				MyItemHighIndex = MyItemHighIndex - 1
				MyItemLowIndex = MyItemLowIndex - 1
			Else
				MyItemRemoved = Nothing
			End If
			If .Count > 0 Then
				'if the element removed was a min or a max we need to find another one
				'the min and the max are not necessary located at the same index
				If MyItemRemoved IsNot Nothing Then
					If MyItemHighIndex < 0 Then
						'MyItemRemoved Is MyItemHigh 
						If MyItemLowIndex < 0 Then
							'MyItemRemoved is also MyItemLow 
							'need to search for a maximum and a minimum at the same time
							'should be a rare occurence
							MyItemHighIndex = 0
							MyItemLowIndex = 0
							For I = 1 To MyListOfIPriceVol.Count - 1
								If MyListOfIPriceVol(I).High > MyListOfIPriceVol(MyItemHighIndex).High Then
									MyItemHighIndex = I
								End If
								If MyListOfIPriceVol(I).Low < MyListOfIPriceVol(MyItemLowIndex).Low Then
									MyItemLowIndex = I
								End If
							Next
						Else
							'MyItemRemoved Is MyItemHigh 
							'MyItemLow is not changed
							'search only for a maximum
							MyItemHighIndex = 0
							For I = 1 To MyListOfIPriceVol.Count - 1
								If MyListOfIPriceVol(I).High > MyListOfIPriceVol(MyItemHighIndex).High Then
									MyItemHighIndex = I
								End If
							Next
						End If
					Else
						If MyItemLowIndex < 0 Then
							'MyItemRemoved is MyItemLow 
							'need to search for a minimum 
							MyItemLowIndex = 0
							For I = 1 To MyListOfIPriceVol.Count - 1
								If MyListOfIPriceVol(I).Low < MyListOfIPriceVol(MyItemLowIndex).Low Then
									MyItemLowIndex = I
								End If
							Next
						End If
					End If
				End If
				'update the max and min with the latest data
				If item.High > MyListOfIPriceVol(MyItemHighIndex).High Then
					MyItemHighIndex = MyListOfIPriceVol.Count
				End If
				If item.Low < MyListOfIPriceVol(MyItemLowIndex).Low Then
					MyItemLowIndex = MyListOfIPriceVol.Count
				End If
			Else
				MyItemHighIndex = MyListOfIPriceVol.Count
				MyItemLowIndex = MyListOfIPriceVol.Count
			End If
			MyVolumeSum = MyVolumeSum + CLng(item.Vol)
			MyPriceSum = MyPriceSum + CDbl(item.Last)
			.Add(item)
			'MyFilterHullForDecimateOpen.Filter(.Item(0).Last)
			MyFilterHullForDecimateOpen.Filter(.Item(0).AsIPricePivotPoint.PivotLast)
			'Debug traces
			WindowFrameDebug.TraceAdd(item, .Count - 1)
			WindowFrameDebug.TraceState(MyListOfIPriceVol, MyItemHighIndex, MyItemLowIndex)
		End With
		'MyListWindowsFrameForPivotOpen.Add(item.PivotOpen)
		'MyListWindowsFrameForPivotOpen.Add(item.Open)
	End Sub

	Public Sub Clear() Implements ICollection(Of T).Clear
		MyListOfIPriceVol.Clear()
		MyVolumeSum = 0
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
	End Sub

	Public Function Contains(item As T) As Boolean Implements ICollection(Of T).Contains
		Return MyListOfIPriceVol.Contains(item)
	End Function

	Public Sub CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
		MyListOfIPriceVol.CopyTo(array, arrayIndex)
	End Sub

	Public ReadOnly Property Count As Integer Implements ICollection(Of T).Count
		Get
			Return MyListOfIPriceVol.Count
		End Get
	End Property

	Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
		Get
			Return False
		End Get
	End Property

	Public Function Remove(item As T) As Boolean Implements ICollection(Of T).Remove
		Throw New NotImplementedException
	End Function
#End Region
#Region "IEnumerable"
	Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
		Return MyListOfIPriceVol.GetEnumerator
	End Function

	''' <summary>
	''' non generic implementation does not need to be public
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
	Private Function IList_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
		Return Me.GetEnumerator()
	End Function
#End Region
#Region "IList"
	Public Function IndexOf(item As T) As Integer Implements IList(Of T).IndexOf
		Return MyListOfIPriceVol.IndexOf(item)
	End Function

	Public Sub Insert(index As Integer, item As T) Implements IList(Of T).Insert
		Throw New NotImplementedException
	End Sub

	Default Public Property Item(index As Integer) As T Implements IList(Of T).Item
		Get
			Return MyListOfIPriceVol.Item(index)
		End Get
		Set(value As T)
			Throw New NotImplementedException
		End Set
	End Property

	Public Sub RemoveAt(index As Integer) Implements IList(Of T).RemoveAt
		Throw New NotImplementedException
	End Sub
#End Region
#Region "IListWindowsFrame1"
	Public ReadOnly Property AsIListWindowsFrame1 As IListWindowsFrame1(Of T) Implements IListWindowsFrame1(Of T).AsIListWindowsFrame1
		Get
			Return Me
		End Get
	End Property

	Public ReadOnly Property ItemLowIndex As Integer Implements IListWindowsFrame1(Of T).ItemLowIndex
		Get
			Return MyItemLowIndex
		End Get
	End Property

	Public ReadOnly Property ItemHighIndex As Integer Implements IListWindowsFrame1(Of T).ItemHighIndex
		Get
			Return MyItemHighIndex
		End Get
	End Property

	Public ReadOnly Property WindowSize As Integer Implements IListWindowsFrame1(Of T).WindowSize
		Get
			Return MyWindowSize
		End Get
	End Property

	Public Function ItemLow() As T Implements IListWindowsFrame1(Of T).ItemLow
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyItemLowIndex)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemHigh() As T Implements IListWindowsFrame1(Of T).ItemHigh
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyItemHighIndex)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemFirst() As T Implements IListWindowsFrame1(Of T).ItemFirst
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(0)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemLast() As T Implements IListWindowsFrame1(Of T).ItemLast
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyListOfIPriceVol.Count - 1)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemDecimate() As T Implements IListWindowsFrame1(Of T).ItemDecimate
		If MyListOfIPriceVol.Count > 0 Then
			Dim ThisData As New T
			With ThisData
				.Open = MyListOfIPriceVol(0).Open
				.OpenNext = .Open
				.High = MyListOfIPriceVol(MyItemHighIndex).High
				.Low = MyListOfIPriceVol(MyItemLowIndex).Low
				If MyVolumeSum > Integer.MaxValue Then
					.Vol = Integer.MaxValue
				ElseIf MyVolumeSum < 0 Then
					.Vol = 0
				Else
					.Vol = CInt(MyVolumeSum)
				End If
				.LastPrevious = .Last
				.Range = RecordPrices.CalculateTrueRange(DirectCast(ThisData, IPriceVol))
			End With
			Return ThisData
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemRemoved() As T Implements IListWindowsFrame1(Of T).ItemRemoved
		Return MyItemRemoved
	End Function
#End Region
End Class
