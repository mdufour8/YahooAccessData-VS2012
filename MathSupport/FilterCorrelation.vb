Imports YahooAccessData.ExtensionService.Extensions
Imports YahooAccessData.MathPlus.Filter

Namespace MathPlus.Filter
	''' <summary>
	''' Calculate the Pearson correlation coefficient r on a list of data gouped in a size n 
	''' https://en.wikipedia.org/wiki/Pearson_correlation_coefficient
	''' </summary>
	''' <remarks></remarks>
	Public Class FilterPearsonCorrelation
		Implements IFilterDuplex

		Private MyQueueOfX As Queue(Of Double)
		Private MyQueueOfY As Queue(Of Double)
		Private MyQueueOfXY As Queue(Of Double)
		Private MyQueueOfX2 As Queue(Of Double)
		Private MyQueueOfY2 As Queue(Of Double)
		Private MySumOfX As Double
		Private MySumOfY As Double
		Private MySumOfXY As Double
		Private MySumOfX2 As Double
		Private MySumOfY2 As Double
		Private MySumOfXMean As Double
		Private MySumOfYMean As Double
		Private MySumOfXYMean As Double
		Private MySumOfX2Mean As Double
		Private MySumOfY2Mean As Double
		Private MyFilterLast As Double
		Private _lastInputs As (X As Double, Y As Double)

		Private MyListOfCorrelation As List(Of Double)
		Private MyRate As Integer

		Public Sub New(ByVal FilterRate As Integer)
			If FilterRate < 1 Then FilterRate = 1
			MyRate = FilterRate

			MyQueueOfX = New Queue(Of Double)(MyRate)
			MyQueueOfY = New Queue(Of Double)(MyRate)
			MyQueueOfXY = New Queue(Of Double)(MyRate)
			MyQueueOfX2 = New Queue(Of Double)(MyRate)
			MyQueueOfY2 = New Queue(Of Double)(MyRate)
			MyListOfCorrelation = New List(Of Double)
			_lastInputs.X = 0.0
			_lastInputs.Y = 0.0
		End Sub

		Public ReadOnly Property Count As Integer Implements IFilterDuplex.Count
			Get
				Return MyListOfCorrelation.Count
			End Get
		End Property

		Public Function Filter(X As Double, Y As Double) As Double Implements IFilterDuplex.Filter
			Dim ThisTemp As Double
			Dim ThisCount As Integer
			Dim ThisCovarianceOfXY As Double
			Dim ThisVarianceOfX As Double
			Dim ThisVarianceOfY As Double
			Dim ThisStandardDeviationOfXY As Double


			_lastInputs.X = X
			_lastInputs.Y = Y
			'data capture 
			If MyListOfCorrelation.Count = 0 Then
				'initialization
				MyQueueOfX.Enqueue(X)
				MySumOfX = X

				MyQueueOfY.Enqueue(Y)
				MySumOfY = Y

				ThisTemp = X * Y
				MyQueueOfXY.Enqueue(ThisTemp)
				MySumOfXY = ThisTemp

				ThisTemp = X ^ 2
				MyQueueOfX2.Enqueue(ThisTemp)
				MySumOfX2 = ThisTemp

				ThisTemp = Y ^ 2
				MyQueueOfY2.Enqueue(ThisTemp)
				MySumOfY2 = ThisTemp
			Else
				If MyQueueOfX.Count < Rate Then
					MyQueueOfX.Enqueue(X)
					MySumOfX = MySumOfX + X

					MyQueueOfY.Enqueue(Y)
					MySumOfY = MySumOfY + Y

					ThisTemp = X * Y
					MyQueueOfXY.Enqueue(ThisTemp)
					MySumOfXY = MySumOfXY + ThisTemp

					ThisTemp = X ^ 2
					MyQueueOfX2.Enqueue(ThisTemp)
					MySumOfX2 = MySumOfX2 + ThisTemp

					ThisTemp = Y ^ 2
					MyQueueOfY2.Enqueue(ThisTemp)
					MySumOfY2 = MySumOfY2 + ThisTemp
				Else
					'add and remove from the queue
					MySumOfX = MySumOfX + X - MyQueueOfX.Dequeue
					MyQueueOfX.Enqueue(X)

					MySumOfY = MySumOfY + Y - MyQueueOfY.Dequeue
					MyQueueOfY.Enqueue(Y)

					ThisTemp = X * Y
					MySumOfXY = MySumOfXY + ThisTemp - MyQueueOfXY.Dequeue
					MyQueueOfXY.Enqueue(ThisTemp)


					ThisTemp = X ^ 2
					MySumOfX2 = MySumOfX2 + ThisTemp - MyQueueOfX2.Dequeue
					MyQueueOfX2.Enqueue(ThisTemp)

					ThisTemp = Y ^ 2
					MySumOfY2 = MySumOfY2 + ThisTemp - MyQueueOfY2.Dequeue
					MyQueueOfY2.Enqueue(ThisTemp)
				End If
			End If
			'we now have all the data ready for the calculation
			ThisCount = MyQueueOfX.Count
			MySumOfXMean = MySumOfX / ThisCount
			MySumOfYMean = MySumOfY / ThisCount
			MySumOfXYMean = MySumOfXY / ThisCount
			MySumOfX2Mean = MySumOfX2 / ThisCount
			MySumOfY2Mean = MySumOfY2 / ThisCount
			ThisCovarianceOfXY = MySumOfXYMean - (MySumOfXMean * MySumOfYMean)
			ThisVarianceOfX = MySumOfX2Mean - (MySumOfXMean * MySumOfXMean)
			ThisVarianceOfY = MySumOfY2Mean - (MySumOfYMean * MySumOfYMean)
			ThisStandardDeviationOfXY = Math.Sqrt(ThisVarianceOfX) * Math.Sqrt(ThisVarianceOfY)

			If ThisStandardDeviationOfXY > 0 Then
				MyFilterLast = ThisCovarianceOfXY / ThisStandardDeviationOfXY
			Else
				'in this case assume the last value is the correct one
			End If
			MyListOfCorrelation.Add(MyFilterLast)
			'MyListOfCorrelation.Add((X - Y) ^ 2 / (X + Y) ^ 2)
			Return MyFilterLast
		End Function

		Public Function FilterLast() As Double Implements IFilterDuplex.FilterLast
			Return MyFilterLast
		End Function

		Public Function Last() As (X As Double, Y As Double) Implements IFilterDuplex.Last
			Return _lastInputs
		End Function

		Public ReadOnly Property Rate As Integer Implements IFilterDuplex.Rate
			Get
				Return MyRate
			End Get
		End Property

		Public Property Tag As String Implements IFilterDuplex.Tag

		Public ReadOnly Property ToListOfCorrelation As List(Of Double)
			Get
				Return MyListOfCorrelation
			End Get
		End Property

		Public ReadOnly Property ToList As IList(Of Double) Implements IFilterDuplex.ToList
			Get
				Return MyListOfCorrelation
			End Get
		End Property

		Public Overrides Function ToString() As String Implements IFilterDuplex.ToString
			Return ""
		End Function
	End Class


	Public Interface IFilterDuplex
		''' <summary>
		''' Calculate continiouly the correlation factor for the last n samples (given by rate)
		''' </summary>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <returns>the correlation factor for the last n samples</returns>
		''' <remarks></remarks>
		Function Filter(ByVal X As Double, ByVal Y As Double) As Double
		Function FilterLast() As Double
		Function Last() As (X As Double, Y As Double)
		ReadOnly Property Rate As Integer
		ReadOnly Property Count As Integer
		ReadOnly Property ToList() As IList(Of Double)
		Property Tag As String
		Function ToString() As String
	End Interface
End Namespace