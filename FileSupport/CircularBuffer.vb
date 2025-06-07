''' <summary>
''' Implement a circular buffer with a fixed capacity.
''' This class is useful for scenarios where you need to maintain a sliding window of data,
''' such as in signal processing or real-time data analysis.	
''' 
''' The buffer is implemented as a Queue using a list, allowing efficient addition 
''' and removal of elements but also fast peek at different index element. The buffer can be 
''' used to store a fixed number of elements, and when it reaches its capacity
''' it overwrites remove the oldest elements with new ones.
''' 
''' The current implementation is not to be use in a a concurrent multi threaded environment.
''' <example>
''' Dim buffer As New CircularBuffer(5)
''' buffer.AddFirst(1.0)
''' buffer.AddLast(2.0)
''' Dim value As Double = buffer(0) ' Gets the oldest value (1.0)
''' Dim count As Integer = buffer.Count ' Gets the number of elements in the buffer
''' Dim removedValue As Double = buffer.RemoveFirst() ' Removes and returns the oldest value (1.0)
''' </example>
''' </summary>
''' <typeparam name="T">The type of elements stored in the buffer.</typeparam>
Public Class CircularBuffer(Of T)
	Private ReadOnly buffer As List(Of T)
	Private head As Integer
	Private MyBufferCount As Integer
	Public ReadOnly Capacity As Integer

	Public Sub New(capacity As Integer, Optional defaultValue As T = Nothing)
		If capacity < 1 Then
			'set the capacity at the minimum of 1 to avoid division by zero and other issues.
			'no need to raise an error, just set it to 1.		
			capacity = 1
			'Throw New ArgumentOutOfRangeException(NameOf(capacity))
		End If
		Me.Capacity = capacity
		buffer = Enumerable.Repeat(defaultValue, capacity).ToList()
		head = 0
		MyBufferCount = 0
	End Sub

	''' <summary>
	''' Adds a value at the head (front). Overwrites the oldest value if full.
	''' </summary>
	Public Sub AddFirst(value As T)
		head = (head - 1 + Capacity) Mod Capacity
		buffer(head) = value
		If MyBufferCount < Capacity Then
			MyBufferCount += 1
		End If
	End Sub

	''' <summary>
	''' Adds a value at the tail (end). Overwrites the oldest value if full.
	''' </summary>
	Public Sub AddLast(value As T)
		Dim tail = (head + MyBufferCount) Mod Capacity
		buffer(tail) = value
		If MyBufferCount < Capacity Then
			MyBufferCount += 1
		Else
			head = (head + 1) Mod Capacity
		End If
	End Sub

	''' <summary>
	''' Gets the item at logical index (0 = oldest, Count - 1 = newest).
	''' </summary>
	Default Public ReadOnly Property Item(BufferIndex As Integer) As T
		Get
			If BufferIndex < 0 OrElse BufferIndex >= MyBufferCount Then Throw New ArgumentOutOfRangeException(NameOf(BufferIndex))
			Return buffer((head + BufferIndex) Mod Capacity)
		End Get
	End Property

	''' <summary>
	''' Number of elements in the buffer.
	''' </summary>
	Public ReadOnly Property Count As Integer
		Get
			Return MyBufferCount
		End Get
	End Property

	''' <summary>
	''' Returns the first (oldest) value without removing it.
	''' </summary>
	Public Function PeekFirst() As T
		If MyBufferCount = 0 Then Throw New InvalidOperationException("Buffer is empty.")
		Return buffer(head)
	End Function

	''' <summary>
	''' Returns the last (newest) value without removing it.
	''' </summary>
	Public Function PeekLast() As T
		If MyBufferCount = 0 Then Throw New InvalidOperationException("Buffer is empty.")
		Dim tail = (head + MyBufferCount - 1 + Capacity) Mod Capacity
		Return buffer(tail)
	End Function

	''' <summary>
	''' Removes and returns the oldest value.
	''' </summary>
	Public Function RemoveFirst() As T
		If MyBufferCount = 0 Then Throw New InvalidOperationException("Buffer is empty.")
		Dim value = buffer(head)
		head = (head + 1) Mod Capacity
		MyBufferCount -= 1
		Return value
	End Function

	''' <summary>
	''' Removes and returns the newest value.
	''' </summary>
	Public Function RemoveLast() As T
		If MyBufferCount = 0 Then Throw New InvalidOperationException("Buffer is empty.")
		Dim tail = (head + MyBufferCount - 1 + Capacity) Mod Capacity
		Dim value = buffer(tail)
		MyBufferCount -= 1
		Return value
	End Function

	''' <summary>
	''' Returns the buffer contents in order (oldest to newest).
	''' </summary>
	Public Function ToArray() As T()
		Dim arr(MyBufferCount - 1) As T
		For i = 0 To MyBufferCount - 1
			arr(i) = Me(i)
		Next
		Return arr
	End Function

	''' <summary>
	''' Returns the buffer contents in order (oldest to newest) in a list.
	''' </summary>
	Public Function ToList() As IList(Of T)
		Dim ThisList = New List(Of T)(MyBufferCount)
		For i = 0 To MyBufferCount - 1
			ThisList.Add(Me.Item(i))
		Next
		Return ThisList
	End Function

	''' <summary>
	''' Clears the buffer.
	''' </summary>
	Public Sub Clear()
		head = 0
		MyBufferCount = 0
	End Sub
End Class
