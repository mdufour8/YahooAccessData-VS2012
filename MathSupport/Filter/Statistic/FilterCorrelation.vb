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
		Private MyFilterLast As ICorrelationData
		Private MyCovarianceOfXY As Double
		Private MyVarianceOfX As Double
		Private MyVarianceOfY As Double
		Private MyStandardDeviationOfX As Double
		Private MyStandardDeviationOfY As Double
		Private MyStandardDeviationOfXY As Double
		Private MyJoinProbabilityOfXY As Double

		Private _lastInputs As (X As Double, Y As Double)

		Private MyListOfCorrelationProbability As List(Of ICorrelationData)
		Private MyRate As Integer

		Public Sub New(ByVal FilterRate As Integer)
			If FilterRate < 1 Then FilterRate = 1
			MyRate = FilterRate

			MyQueueOfX = New Queue(Of Double)(MyRate)
			MyQueueOfY = New Queue(Of Double)(MyRate)
			MyQueueOfXY = New Queue(Of Double)(MyRate)
			MyQueueOfX2 = New Queue(Of Double)(MyRate)
			MyQueueOfY2 = New Queue(Of Double)(MyRate)
			MyListOfCorrelationProbability = New List(Of ICorrelationData)
			_lastInputs.X = 0.0
			_lastInputs.Y = 0.0
		End Sub

		Public ReadOnly Property Count As Integer Implements IFilterDuplex.Count
			Get
				Return MyListOfCorrelationProbability.Count
			End Get
		End Property

		Public Function Filter(X As Double, Y As Double) As ICorrelationData Implements IFilterDuplex.Filter
			Dim ThisTemp As Double
			Dim ThisCount As Integer

			_lastInputs.X = X
			_lastInputs.Y = Y
			'data capture 
			If MyListOfCorrelationProbability.Count = 0 Then
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
			MyCovarianceOfXY = MySumOfXYMean - (MySumOfXMean * MySumOfYMean)
			MyVarianceOfX = MySumOfX2Mean - (MySumOfXMean * MySumOfXMean)
			If MyVarianceOfX < 0.0 Then MyVarianceOfX = 0.0
			MyVarianceOfY = MySumOfY2Mean - (MySumOfYMean * MySumOfYMean)
			If MyVarianceOfY < 0.0 Then MyVarianceOfY = 0.0
			MyStandardDeviationOfX = Math.Sqrt(MyVarianceOfX)
			MyStandardDeviationOfY = Math.Sqrt(MyVarianceOfY)
			MyStandardDeviationOfXY = MyStandardDeviationOfX * MyStandardDeviationOfY
			If MyStandardDeviationOfXY > 0 Then
				Dim ThisCorrelation = MyCovarianceOfXY / MyStandardDeviationOfXY
				'calculate the Probability that both process are moving together or are positive at the same time i.e.	P(X And Y)
				Dim ThisJoinProbabilityOfXY = Measure.Measure.JointProbabilityApproximate(ThisCorrelation)
				'In simple linear regression, the coefficient represents the slope of the line between the two variales 
				'and Is calculated as the covariance of the two variables divided by the variance of the independent variable.
				'In this case, we are treating X as the independent variable And Y as the dependent variable, so the regression coefficient Is calculated as the covariance of X And Y
				'divided by the variance of X.
				Dim ThisRegression = If(MyVarianceOfX > 0.0, MyCovarianceOfXY / MyVarianceOfX, 0.0)
				MyFilterLast = New CorrelationData(ThisCorrelation, ThisJoinProbabilityOfXY, ThisRegression, MySumOfYMean, MyStandardDeviationOfY)
			Else
				MyFilterLast = New CorrelationData(0.0, 0.5, 0.0, MySumOfYMean, MyStandardDeviationOfY)
			End If
			MyListOfCorrelationProbability.Add(MyFilterLast)
			Return MyFilterLast
		End Function

		Public Function FilterLast() As ICorrelationData Implements IFilterDuplex.FilterLast
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

		Public ReadOnly Property ToList As IList(Of ICorrelationData) Implements IFilterDuplex.ToList
			Get
				Return MyListOfCorrelationProbability
			End Get
		End Property

		Public Overrides Function ToString() As String Implements IFilterDuplex.ToString
			Return FilterLast.ToString
		End Function

		Public ReadOnly Property SumOfX As Double
			Get
				Return MySumOfX
			End Get
		End Property
		Public ReadOnly Property SumOfY As Double
			Get
				Return MySumOfY
			End Get
		End Property
		Public ReadOnly Property SumOfXY As Double
			Get
				Return MySumOfXY
			End Get
		End Property
		Public ReadOnly Property SumOfX2 As Double
			Get
				Return MySumOfX2
			End Get
		End Property
		Public ReadOnly Property SumOfY2 As Double
			Get
				Return MySumOfY2
			End Get
		End Property
		Public ReadOnly Property SumOfXMean As Double
			Get
				Return MySumOfXMean
			End Get
		End Property
		Public ReadOnly Property SumOfYMean As Double
			Get
				Return MySumOfYMean
			End Get
		End Property
		Public ReadOnly Property SumOfXYMean As Double
			Get
				Return MySumOfXYMean
			End Get
		End Property
		Public ReadOnly Property SumOfX2Mean As Double
			Get
				Return MySumOfX2Mean
			End Get
		End Property
		Public ReadOnly Property SumOfY2Mean As Double
			Get
				Return MySumOfY2Mean
			End Get
		End Property
		Public ReadOnly Property CovarianceOfXY As Double
			Get
				Return MyCovarianceOfXY
			End Get
		End Property
		Public ReadOnly Property VarianceOfX As Double
			Get
				Return MyVarianceOfX
			End Get
		End Property
		Public ReadOnly Property VarianceOfY As Double
			Get
				Return MyVarianceOfY
			End Get
		End Property
		Public ReadOnly Property StandardDeviationOfXY As Double
			Get
				Return MyStandardDeviationOfXY
			End Get
		End Property

		Public ReadOnly Property StandardDeviationOfX As Double
			Get
				Return MyStandardDeviationOfX
			End Get
		End Property

		Public ReadOnly Property StandardDeviationOfY As Double
			Get
				Return MyStandardDeviationOfY
			End Get
		End Property
	End Class

	Public Interface IFilterDuplex
		''' <summary>
		''' Calculate continiouly the correlation factor for the last n samples (given by rate)
		''' </summary>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <returns>the correlation factor for the last n samples</returns>
		''' <remarks></remarks>
		Function Filter(ByVal X As Double, ByVal Y As Double) As ICorrelationData
		Function FilterLast() As ICorrelationData
		Function Last() As (X As Double, Y As Double)
		ReadOnly Property Rate As Integer
		ReadOnly Property Count As Integer
		ReadOnly Property ToList() As IList(Of ICorrelationData)
		Property Tag As String
		Function ToString() As String
	End Interface
End Namespace