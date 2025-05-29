''' <summary>
''' Once you’ve modeled nodes with multiple I/O, you’ll want to connect them. 
''' This can be done by a graph-like manager:
''' </summary>

Public Class FilterNetwork
	Private _nodes As New List(Of IFilterNode)
	Private _connections As New List(Of (FromNode As IFilterNode, FromOutput As Integer, ToNode As IFilterNode, ToInput As Integer))

	Public Sub AddNode(node As IFilterNode)
		_nodes.Add(node)
	End Sub

	Public Sub Connect(fromNode As IFilterNode, fromOutput As Integer, toNode As IFilterNode, toInput As Integer)
		_connections.Add((fromNode, fromOutput, toNode, toInput))
	End Sub

	Public Sub ProcessAll()
		' Add scheduling or iteration logic to ensure correct update order
	End Sub
End Class

''' <summary>
'''  example of a simple filter node that smooths input and calculates its slope
'''  Smoothing and Trend Filter Node
'''  This node processes a single input to produce a smoothed value and its slope (derivative).
''' </summary>
Public Class SmoothingAndTrendFilter
	Implements IFilterNode

	Public Function Process(inputs As IReadOnlyList(Of Double)) As IReadOnlyList(Of Double) Implements IFilterNode.Process
		Dim x = inputs(0)
		'Dim smooth = ... ' smooth value
		'Dim slope = ... ' derivative
		'Return New Double() {smooth, slope}
		Return Nothing
	End Function

	Public ReadOnly Property InputCount As Integer Implements IFilterNode.InputCount
		Get
			Return 1
		End Get
	End Property

	Public ReadOnly Property OutputCount As Integer Implements IFilterNode.OutputCount
		Get
			Return 2
		End Get
	End Property

	Public ReadOnly Property Name As String Implements IFilterNode.Name
		Get
			Return "Smooth+Slope"
		End Get
	End Property

	Public Sub Reset() Implements IFilterNode.Reset
		' clear internal buffer
	End Sub
End Class