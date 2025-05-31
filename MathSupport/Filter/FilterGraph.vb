Public Class FilterGraph
	Private ReadOnly _nodes As New List(Of IFilterNode)
	Private ReadOnly _connections As New List(Of (FromNode As IFilterNode, FromOutput As Integer, ToNode As IFilterNode, ToInput As Integer))

	Public Sub AddNode(node As IFilterNode)
		_nodes.Add(node)
	End Sub

	Public Sub Connect(fromNode As IFilterNode, fromOutput As Integer, toNode As IFilterNode, toInput As Integer)
		_connections.Add((fromNode, fromOutput, toNode, toInput))
	End Sub

	Private ReadOnly _signalBus As New Dictionary(Of (Node As IFilterNode, OutputIndex As Integer), Double)
	Public ReadOnly Property SignalBus As IReadOnlyDictionary(Of (IFilterNode, Integer), Double)
		Get
			Return _signalBus
		End Get
	End Property

	Public Sub ProcessAll()
		_signalBus.Clear()

		' Map each node to its input values and run it
		For Each node In _nodes
			Dim inputValues As New List(Of Double)(New Double(node.InputCount - 1) {})

			' Fill inputValues from _signalBus based on _connections
			For Each conn In _connections
				If conn.ToNode Is node Then
					Dim sourceKey = (conn.FromNode, conn.FromOutput)
					If _signalBus.ContainsKey(sourceKey) Then
						inputValues(conn.ToInput) = _signalBus(sourceKey)
					Else
						Throw New InvalidOperationException($"Missing signal for {conn.FromNode.Name} output {conn.FromOutput}")
					End If
				End If
			Next

			' Process the node
			Dim outputs = node.Process(inputValues)

			' Store outputs in signal bus
			For i = 0 To outputs.Count - 1
				_signalBus((node, i)) = outputs(i)
			Next
		Next
	End Sub
End Class
