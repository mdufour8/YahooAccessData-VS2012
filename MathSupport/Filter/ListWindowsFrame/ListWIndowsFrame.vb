Imports YahooAccessData.MathPlus.Filter

Public Class ListWindowFrame
	Implements IUndoLastState
	Implements IListWindowsFrame(Of Double)
	Implements IListWindowsFrame1(Of Double)

	Private _buf As CircularBuffer(Of Double)
	Private _dirtyMinMax As Boolean

	Private MyWindowSize As Integer
	Private MyItemsSum As Double
	Private MyItemHighIndex As Integer
	Private MyItemLowIndex As Integer
	Private MyItemRemoved As Double?

	Private _hasUndo As Boolean
	Private Structure UndoFrame
		Public PrevSum As Double
		Public PrevRemoved As Double?
		Public PrevDirty As Boolean
		Public PrevHighIndex As Integer
		Public PrevLowIndex As Integer
	End Structure

	Private _undo As UndoFrame


#Region "New"
	Public Sub New(ByVal WindowSize As Integer)
		_buf = New CircularBuffer(Of Double)(WindowSize, defaultValue:=0.0R)
		MyWindowSize = _buf.Capacity   ' ← clamp to actual buffer size
		MyItemsSum = 0.0
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
		_dirtyMinMax = True
		_hasUndo = False
	End Sub
#End Region

	Public Sub Add(item As Double)
		' Save undo for the FRAME state (buffer handles its own undo)
		_undo = New UndoFrame With {
				.PrevSum = MyItemsSum,
				.PrevRemoved = MyItemRemoved,
				.PrevDirty = _dirtyMinMax,
				.PrevHighIndex = MyItemHighIndex,
				.PrevLowIndex = MyItemLowIndex
		}
		_hasUndo = True

		Dim evicted As Double
		Dim hadEviction As Boolean = _buf.AddLast(item, evicted)

		If hadEviction Then
			MyItemRemoved = evicted
			MyItemsSum -= evicted
		Else
			MyItemRemoved = Nothing
		End If
		MyItemsSum += item
		_dirtyMinMax = True
	End Sub

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
		If _buf.Count = 0 Then Return Nothing
		Return _buf.PeekFirst()
	End Function

	Public Function ItemHigh() As Double? Implements IListWindowsFrame(Of Double).ItemHigh
		If _buf.Count = 0 Then Return Nothing
		EnsureMinMax()
		Return _buf(MyItemHighIndex)
	End Function

	Public ReadOnly Property ItemHighIndex As Integer Implements IListWindowsFrame(Of Double).ItemHighIndex
		Get
			If _buf.Count = 0 Then Return -1
			EnsureMinMax()
			Return MyItemHighIndex
		End Get
	End Property

	Public Function ItemLast() As Double? Implements IListWindowsFrame(Of Double).ItemLast
		If _buf.Count = 0 Then Return Nothing
		Return _buf.PeekLast()
	End Function

	Public Function ItemLow() As Double? Implements IListWindowsFrame(Of Double).ItemLow
		If _buf.Count = 0 Then Return Nothing
		EnsureMinMax()
		Return _buf(MyItemLowIndex)
	End Function

	Public ReadOnly Property ItemLowIndex As Integer Implements IListWindowsFrame(Of Double).ItemLowIndex
		Get
			If _buf.Count = 0 Then Return -1
			EnsureMinMax()
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
			Return Me.ItemLast.Value   ' <-- not ItemFirst
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

	Private Sub EnsureMinMax()
		If Not _dirtyMinMax Then Return

		If _buf.Count = 0 Then
			MyItemHighIndex = -1
			MyItemLowIndex = -1
			_dirtyMinMax = False
			Return
		End If

		Dim hi As Integer = 0
		Dim lo As Integer = 0
		Dim vHi As Double = _buf(0)
		Dim vLo As Double = vHi

		'starting at 1, since we already used 0 to initialize the min/max
		For i = 1 To _buf.Count - 1
			Dim v = _buf(i)
			If v > vHi Then
				vHi = v
				hi = i
			End If
			If v < vLo Then
				vLo = v
				lo = i
			End If
		Next
		MyItemHighIndex = hi
		MyItemLowIndex = lo
		_dirtyMinMax = False
	End Sub

	Public Sub Clear()
		_buf.Clear()
		MyItemsSum = 0
		MyItemHighIndex = -1
		MyItemLowIndex = -1
		MyItemRemoved = Nothing
		_dirtyMinMax = True
		_hasUndo = False
	End Sub

	Public ReadOnly Property Count As Integer
		Get
			Return _buf.Count
		End Get
	End Property

	Public Function RestoreLastState() As Boolean Implements IUndoLastState.RestoreLastState
		If Not _hasUndo Then Return False
		If Not _buf.RestoreLastState() Then Return False

		MyItemsSum = _undo.PrevSum
		MyItemRemoved = _undo.PrevRemoved
		_dirtyMinMax = _undo.PrevDirty
		MyItemHighIndex = _undo.PrevHighIndex
		MyItemLowIndex = _undo.PrevLowIndex

		_hasUndo = False
		Return True
	End Function
End Class


