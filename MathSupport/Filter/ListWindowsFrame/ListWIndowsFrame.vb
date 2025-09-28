Imports YahooAccessData.MathPlus.Filter

Public Class ListWindowFrame
	Implements IList(Of Double)
	Implements IListWindowsFrame(Of Double)
	Implements IListWindowsFrame1(Of Double)

	Private MyListOfIPriceVol As List(Of Double)
	Private MyWindowSize As Integer
	Private MyItemsSum As Double
	Private MyItemHighIndex As Integer
	Private MyItemLowIndex As Integer
	Private MyItemRemoved As Nullable(Of Double)
	Private IsItemRemoved As Boolean

#Region "New"
	Public Sub New(ByVal WindowSize As Integer)
		MyListOfIPriceVol = New List(Of Double)(WindowSize)
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
		MyWindowSize = WindowSize
	End Sub
#End Region
#Region "IListWindowsFrame"
	Private ReadOnly Property AsIListWindowsFrame As IListWindowsFrame(Of Double) Implements IListWindowsFrame(Of Double).AsIListWindowsFrame
		Get
			Return Me
		End Get
	End Property

	Public Function ItemDecimate() As Double? Implements IListWindowsFrame(Of Double).ItemDecimate
		Return MyItemsSum
	End Function

	Public Function ItemFirst() As Double? Implements IListWindowsFrame(Of Double).ItemFirst
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(0)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemHigh() As Double? Implements IListWindowsFrame(Of Double).ItemHigh
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyItemHighIndex)
		Else
			Return Nothing
		End If
	End Function

	Public ReadOnly Property ItemHighIndex As Integer Implements IListWindowsFrame(Of Double).ItemHighIndex
		Get
			Return MyItemHighIndex
		End Get
	End Property

	Public Function ItemLast() As Double? Implements IListWindowsFrame(Of Double).ItemLast
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyListOfIPriceVol.Count - 1)
		Else
			Return Nothing
		End If
	End Function

	Public Function ItemLow() As Double? Implements IListWindowsFrame(Of Double).ItemLow
		If MyListOfIPriceVol.Count > 0 Then
			Return MyListOfIPriceVol(MyItemLowIndex)
		Else
			Return Nothing
		End If
	End Function

	Public ReadOnly Property ItemLowIndex As Integer Implements IListWindowsFrame(Of Double).ItemLowIndex
		Get
			Return MyItemLowIndex
		End Get
	End Property

	Public Function ItemRemoved() As Double? Implements IListWindowsFrame(Of Double).ItemRemoved
		Return MyItemRemoved
	End Function

	Public ReadOnly Property WindowSize As Integer Implements IListWindowsFrame(Of Double).WindowSize
		Get
			Return MyWindowSize
		End Get
	End Property
#End Region
#Region "IListWindowsFrame1"
	Public ReadOnly Property AsIListWindowsFrame1 As IListWindowsFrame1(Of Double) Implements IListWindowsFrame1(Of Double).AsIListWindowsFrame1
		Get
			Return Me
		End Get
	End Property

	Private ReadOnly Property IListWindowsFrame1_ItemLowIndex As Integer Implements IListWindowsFrame1(Of Double).ItemLowIndex
		Get
			Return ItemLowIndex
		End Get
	End Property

	Private ReadOnly Property IListWindowsFrame1_ItemHighIndex As Integer Implements IListWindowsFrame1(Of Double).ItemHighIndex
		Get
			Return ItemHighIndex
		End Get
	End Property

	Private ReadOnly Property IListWindowsFrame1_WindowSize As Integer Implements IListWindowsFrame1(Of Double).WindowSize
		Get
			Return WindowSize
		End Get
	End Property
	Private Function IListWindowsFrame1_ItemLow() As Double Implements IListWindowsFrame1(Of Double).ItemLow
		If Me.ItemLow.HasValue Then
			Return Me.ItemLow.Value
		Else
			Return Double.NaN
		End If
	End Function

	Private Function IListWindowsFrame1_ItemHigh() As Double Implements IListWindowsFrame1(Of Double).ItemHigh
		If Me.ItemHigh.HasValue Then
			Return Me.ItemHigh.Value
		Else
			Return Double.NaN
		End If
	End Function

	Private Function IListWindowsFrame1_ItemFirst() As Double Implements IListWindowsFrame1(Of Double).ItemFirst
		If Me.ItemFirst.HasValue Then
			Return Me.ItemFirst.Value
		Else
			Return Double.NaN
		End If
	End Function

	Private Function IListWindowsFrame1_ItemLast() As Double Implements IListWindowsFrame1(Of Double).ItemLast
		If Me.ItemLast.HasValue Then
			Return Me.ItemFirst.Value
		Else
			Return Double.NaN
		End If
	End Function

	Private Function IListWindowsFrame1_ItemDecimate() As Double Implements IListWindowsFrame1(Of Double).ItemDecimate
		If Me.ItemDecimate.HasValue Then
			Return Me.ItemDecimate.Value
		Else
			Return Double.NaN
		End If
	End Function

	Private Function IListWindowsFrame1_ItemRemoved() As Double Implements IListWindowsFrame1(Of Double).ItemRemoved
		If Me.ItemRemoved.HasValue Then
			Return Me.ItemRemoved.Value
		Else
			Return Double.NaN
		End If
	End Function
#End Region
#Region "ICollection"
	Public Sub Add(item As Double) Implements ICollection(Of Double).Add
		Dim I As Integer

		With MyListOfIPriceVol
			If .Count = MyWindowSize Then
				MyItemRemoved = .First
				MyItemsSum = MyItemsSum - .First
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
								If MyListOfIPriceVol(I) > MyListOfIPriceVol(MyItemHighIndex) Then
									MyItemHighIndex = I
								End If
								If MyListOfIPriceVol(I) < MyListOfIPriceVol(MyItemLowIndex) Then
									MyItemLowIndex = I
								End If
							Next
						Else
							'MyItemRemoved Is MyItemHigh 
							'MyItemLow is not changed
							'search only for a maximum
							MyItemHighIndex = 0
							For I = 1 To MyListOfIPriceVol.Count - 1
								If MyListOfIPriceVol(I) > MyListOfIPriceVol(MyItemHighIndex) Then
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
								If MyListOfIPriceVol(I) < MyListOfIPriceVol(MyItemLowIndex) Then
									MyItemLowIndex = I
								End If
							Next
						End If
					End If
				End If
				'update the max and min with the latest data
				If item > MyListOfIPriceVol(MyItemHighIndex) Then
					MyItemHighIndex = MyListOfIPriceVol.Count
				End If
				If item < MyListOfIPriceVol(MyItemLowIndex) Then
					MyItemLowIndex = MyListOfIPriceVol.Count
				End If
			Else
				MyItemHighIndex = MyListOfIPriceVol.Count
				MyItemLowIndex = MyListOfIPriceVol.Count
			End If
			MyItemsSum = MyItemsSum + item
			.Add(item)
		End With
	End Sub

	Public Sub Clear() Implements ICollection(Of Double).Clear
		MyListOfIPriceVol.Clear()
		MyItemsSum = 0
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
	End Sub

	Public Function Contains(item As Double) As Boolean Implements ICollection(Of Double).Contains
		Return MyListOfIPriceVol.Contains(item)
	End Function

	Public Sub CopyTo(array() As Double, arrayIndex As Integer) Implements ICollection(Of Double).CopyTo
		MyListOfIPriceVol.CopyTo(array, arrayIndex)
	End Sub

	Public ReadOnly Property Count As Integer Implements ICollection(Of Double).Count
		Get
			Return MyListOfIPriceVol.Count
		End Get
	End Property

	Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of Double).IsReadOnly
		Get
			Return False
		End Get
	End Property

	Public Function Remove(item As Double) As Boolean Implements ICollection(Of Double).Remove
		Throw New NotImplementedException
	End Function
#End Region
#Region "IEnumerable"
	Public Function GetEnumerator() As IEnumerator(Of Double) Implements IEnumerable(Of Double).GetEnumerator
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
	Public Function IndexOf(item As Double) As Integer Implements IList(Of Double).IndexOf
		Return MyListOfIPriceVol.IndexOf(item)
	End Function

	Public Sub Insert(index As Integer, item As Double) Implements IList(Of Double).Insert
		Throw New NotImplementedException
	End Sub

	Default Public Property Item(index As Integer) As Double Implements IList(Of Double).Item
		Get
			Return MyListOfIPriceVol.Item(index)
		End Get
		Set(value As Double)
			Throw New NotImplementedException
		End Set
	End Property

	Public Sub RemoveAt(index As Integer) Implements IList(Of Double).RemoveAt
		Throw New NotImplementedException
	End Sub
#End Region
End Class


