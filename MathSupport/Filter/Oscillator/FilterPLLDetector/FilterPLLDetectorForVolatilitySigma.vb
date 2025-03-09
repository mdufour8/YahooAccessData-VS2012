Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions

Namespace MathPlus.Filter
	Public Class FilterPLLDetectorForVolatilitySigma
		Implements IFilterPLLDetector

		Public Const RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE As Double = 1.0
		Public Const BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED As Integer = 20

		Private MyRate As Integer
		Private MyCount As Integer
		Private MyErrorLast As Double
		Private MyToCountLimit As Integer
		Private MyToCountLimitSelected As Integer
		Private MyToErrorLimit As Double
		Private MyValueInit As Double
		Private MyValueOutput As Double
		Private MyQueueForSigmaStatisticDaily As Queue(Of IStockPriceVolatilityPredictionBand)
		Private MyQueueForSigmaStatisticDailyVolPlus As Queue(Of IStockPriceVolatilityPredictionBand)
		Private MyQueueForSigmaStatisticDailyVolMinus As Queue(Of IStockPriceVolatilityPredictionBand)
		Private MyFilterExpForQueuePLE As FilterExp
		Private MyFilterExpForQueuePLEVolPlus As FilterExp
		Private MyFilterExpForQueuePLEVolMinus As FilterExp

		Private MyQueueSumOfExcess As Double
		Private MyQueueSumOfExcessVolPlus As Double
		Private MyQueueSumOfExcessVolMinus As Double
		Private MySumForSigmaStatisticDaily As Double
		Private MyCountOfPLLRun As Integer
		Private MyStatus As Boolean
		Private MyVolatilityAverage As Double
		Private MyFilterPLL As FilterPLL
		Private MyDetectorBandExcessBalanceSum As Double
		Private MyDetectorBalanceLast As Double
		Private MyMaximum As Double
		Private MyListOfConvergence As IList(Of Double)
		Private MyProbabilityOfExcessMeasuredLast As Double
		Private MyListOfProbabilityOfExcess As IList(Of Double)
		Private MyListOfProbabilityOfExcessBalance As IList(Of Double)
		Private MyFastAttackCountForBandExceededLow As Double
		Private MyFastAttackCountForBandExceededHigh As Double
		Private MyListOfValueOutput As IList(Of Double)

		Public Sub New(ByVal Rate As Integer, Optional ByVal ToCountLimit As Integer = 1, Optional ToErrorLimit As Double = 0.001)
			MyRate = Rate
			If MyRate < BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED Then
				MyRate = BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
			End If
			Me.IsUseFeedbackRegulatedVolatilityFastAttackEvent = False    'by default
			MyToCountLimit = ToCountLimit
			MyToCountLimitSelected = ToCountLimit
			MyToErrorLimit = ToErrorLimit
			MyCount = 0
			MyValueInit = 0
			MyCountOfPLLRun = 0
			MyStatus = False
			MyQueueForSigmaStatisticDaily = New Queue(Of IStockPriceVolatilityPredictionBand)(capacity:=CInt(MyRate))
			MyQueueForSigmaStatisticDailyVolPlus = New Queue(Of IStockPriceVolatilityPredictionBand)(capacity:=CInt(MyRate))
			MyQueueForSigmaStatisticDailyVolMinus = New Queue(Of IStockPriceVolatilityPredictionBand)(capacity:=CInt(MyRate))
			MyFilterExpForQueuePLE = New FilterExp(FilterRate:=MyRate)
			MyFilterExpForQueuePLEVolPlus = New FilterExp(FilterRate:=MyRate)
			MyFilterExpForQueuePLEVolMinus = New FilterExp(FilterRate:=MyRate)
			MyFilterPLL = New FilterPLL(FilterRate:=7, DampingFactor:=1.0)
			MyListOfConvergence = New List(Of Double)
			MyListOfProbabilityOfExcess = New List(Of Double)
			MyListOfProbabilityOfExcessBalance = New List(Of Double)
			MyListOfValueOutput = New List(Of Double)
			Me.Tag = TypeName(Me)
		End Sub

		Private Function GetQueueProbabilityOfExcess() As Double
			Dim ThisProbOfBandExceedExpected As Double
			Dim ThisCountForBandExcess As Integer = 0
			Dim I As Integer

			Throw New NotImplementedException
			For Each ThisItem In MyQueueForSigmaStatisticDaily
				ThisItem.Refresh()
				If ThisItem.IsBandExceeded Then
					ThisCountForBandExcess += 1
				End If
			Next

			'MySumForSigmaStatisticDaily = 0
			'	MyDetectorBandExcessBalanceSum = 0
			'	ThisProbOfBandExceedExpected = 1 - MyQueueForSigmaStatisticDaily(0).ProbabilityOfInterval
			'	For I = 0 To MyQueueForSigmaStatisticDaily.Count - 1

			'		If .IsBandExceededHigh Then
			'			MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum + 1
			'		End If
			'		If .IsBandExceededLow Then
			'			MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum - 1
			'		End If
			'		Else
			'		IsBandExceededLast = False
			'		End If
			'		End If
			'		End With
			'	Next
			'	If ThisCount > 1 Then
			'		MyDetectorBalanceLast = MyDetectorBandExcessBalanceSum / ThisCount
			'		ThisGradientMean = ThisGradientSum / ThisCount
			'		If MySumForSigmaStatisticDaily > ThisCount Then
			'			MySumForSigmaStatisticDaily = ThisCount
			'		End If
			'		MyProbabilityOfExcessMeasuredLast = MySumForSigmaStatisticDaily / ThisCount
			'		MyErrorLast = (MyProbabilityOfExcessMeasuredLast - ThisProbOfBandExceedExpected)
			'		MyErrorLast = -MyErrorLast / ThisGradientMean
			'		MyStatus = True
			'		Exit Do
			'	Else
			'		MyDetectorBalanceLast = 0
			'		ThisGradientMean = 0
			'		MyProbabilityOfExcessMeasuredLast = 0
			'		MyErrorLast = 0
			'		MyStatus = False
			'		Exit Do
			'	End If
			'Loop
		End Function

		Private Function IFilterPLLDetector_RunErrorDetector(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.RunErrorDetector
			Dim ThisStockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand
			Dim IsBandExceededLast As Boolean
			Dim ThisProbOfBandExceedExpected As Double
			Dim ThisGradientSum As Double
			Dim ThisGradientMean As Double
			Dim ThisValueOuput As Double = Me.IFilterPLLDetector_ValueOutput(Input, InputFeedback)
			Dim ThisVolatilityChangePerCent As Double
			Dim ThisCount As Integer
			Dim ThisWeight As Double
			Dim ThisWeightStep As Double
			Dim I As Integer

			If Me.IsUseFeedbackRegulatedVolatilityFastAttackEvent Then

			Else

			End If

			MyStatus = False
			MyErrorLast = 0   'by default
			'this code execute every time because the count is a multiple of 1
			If MyCount Mod 1 = 0 Then
				MyCountOfPLLRun = MyCountOfPLLRun + 1
				Do
					MySumForSigmaStatisticDaily = 0
					MyDetectorBandExcessBalanceSum = 0
					ThisGradientSum = 0
					ThisCount = 0
					If Input <= 0 Then Exit Do
					If Me.Count = 500 Then
						I = I
					End If
					ThisVolatilityChangePerCent = InputFeedback / Input

					ThisProbOfBandExceedExpected = 1 - MyQueueForSigmaStatisticDaily(0).ProbabilityOfInterval

					For I = 0 To MyQueueForSigmaStatisticDaily.Count - 1
						ThisStockPriceVolatilityPredictionBand = MyQueueForSigmaStatisticDaily(I)
						ThisWeight = ThisWeight + ThisWeightStep
						With ThisStockPriceVolatilityPredictionBand
							If .Volatility = 0 Then
								Exit Do
							Else
								.Refresh(VolatilityDelta:=ThisVolatilityChangePerCent * .Volatility)
								ThisCount = ThisCount + 1
								ThisGradientSum = ThisGradientSum + .RatioOfΔProbabilityToΔVolatility
								If .IsBandExceeded Then
									If IsBandExceededLast Then
									Else
										IsBandExceededLast = True
									End If
									MySumForSigmaStatisticDaily = MySumForSigmaStatisticDaily + 1
									If .IsBandExceededHigh Then
										MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum + 1
									End If
									If .IsBandExceededLow Then
										MyDetectorBandExcessBalanceSum = MyDetectorBandExcessBalanceSum - 1
									End If
								Else
									IsBandExceededLast = False
								End If
							End If
						End With
					Next
					If ThisCount > 1 Then
						MyDetectorBalanceLast = MyDetectorBandExcessBalanceSum / ThisCount
						ThisGradientMean = ThisGradientSum / ThisCount
						If MySumForSigmaStatisticDaily > ThisCount Then
							MySumForSigmaStatisticDaily = ThisCount
						End If
						MyProbabilityOfExcessMeasuredLast = MySumForSigmaStatisticDaily / ThisCount
						MyErrorLast = (MyProbabilityOfExcessMeasuredLast - ThisProbOfBandExceedExpected)
						MyErrorLast = -MyErrorLast / ThisGradientMean
						MyStatus = True
						Exit Do
					Else
						MyDetectorBalanceLast = 0
						ThisGradientMean = 0
						MyProbabilityOfExcessMeasuredLast = 0
						MyErrorLast = 0
						MyStatus = False
						Exit Do
					End If
				Loop
			End If
			Return MyErrorLast
		End Function

		Public Sub UpdateData(ByVal StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand)
			Dim ThisQueueDataLast As IStockPriceVolatilityPredictionBand = Nothing
			Dim ThisQueueDataRemoved As IStockPriceVolatilityPredictionBand = Nothing
			Dim ThisVolatilityDelta As Double
			Dim ThisQueueDataLastDate As Date
			Dim ThisQueueDataRemovedDate As Date

			MyCount = MyCount + 1
			If MyQueueForSigmaStatisticDaily.Count > 0 Then
				ThisQueueDataLast = MyQueueForSigmaStatisticDaily.Last
				ThisQueueDataLastDate = ThisQueueDataLast.StockPrice.DateDay
				'so  update the last PriceVolatility with the actual price 
				With ThisQueueDataLast
					.Refresh(StockPriceVolatilityPredictionBand.StockPrice)
					'.Refresh(ThisQueueDataLast.VolatilityDelta, StockPriceVolatilityPredictionBand.StockPrice)
				End With
			End If
			If MyQueueForSigmaStatisticDaily.Count = MyRate Then
				ThisQueueDataRemoved = MyQueueForSigmaStatisticDaily.Dequeue
				ThisQueueDataRemovedDate = ThisQueueDataRemoved.StockPrice.DateDay
			End If
			MyToCountLimitSelected = MyToCountLimit
			MyQueueForSigmaStatisticDaily.Enqueue(StockPriceVolatilityPredictionBand)

			ThisVolatilityDelta = MyFilterPLL.FilterRun(StockPriceVolatilityPredictionBand.Volatility, Me)
			MyListOfValueOutput.Add(ThisVolatilityDelta)
			Me.IFilterPLLDetector_RunErrorDetector(StockPriceVolatilityPredictionBand.Volatility, ThisVolatilityDelta)
			StockPriceVolatilityPredictionBand.Refresh(VolatilityDelta:=ThisVolatilityDelta)

			MyListOfProbabilityOfExcess.Add(MyProbabilityOfExcessMeasuredLast)
			MyListOfProbabilityOfExcessBalance.Add(MyDetectorBalanceLast)
			MyToCountLimitSelected = MyToCountLimit
		End Sub

		Public Sub Update(ByVal StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand)
			Dim ThisQueueDataLast As IStockPriceVolatilityPredictionBand = Nothing
			Dim ThisQueueDataRemoved As IStockPriceVolatilityPredictionBand = Nothing
			Dim ThisVolatilityDelta As Double
			Dim ThisQueueDataLastDate As Date
			Dim ThisQueueDataRemovedDate As Date
			Dim ThisProbabilityOfExcess As Double

			MyCount = MyCount + 1

			QueueUpdate(StockPriceVolatilityPredictionBand, QueueSource:=MyQueueForSigmaStatisticDaily, FilterForPLE:=MyFilterExpForQueuePLE)





			MyToCountLimitSelected = MyToCountLimit

			'ThisVolatilityDelta = MyFilterPLL.FilterRun(StockPriceVolatilityPredictionBand.Volatility, Me)
			'MyListOfValueOutput.Add(ThisVolatilityDelta)

			'Me.IFilterPLLDetector_RunErrorDetector(StockPriceVolatilityPredictionBand.Volatility, ThisVolatilityDelta)

			'StockPriceVolatilityPredictionBand.Refresh(VolatilityDelta:=ThisVolatilityDelta)

			MyListOfProbabilityOfExcess.Add(MyProbabilityOfExcessMeasuredLast)
			MyListOfProbabilityOfExcessBalance.Add(MyDetectorBalanceLast)
			MyToCountLimitSelected = MyToCountLimit
		End Sub

		Private Sub QueueUpdate(
			StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand,
			QueueSource As Queue(Of IStockPriceVolatilityPredictionBand),
			FilterForPLE As IFilterRun)

			Dim ThisQueueDataLastDate As Date = Nothing
			Dim ThisQueueDataRemovedDate As Date = Nothing
			Dim ThisQueueDataRemoved As IStockPriceVolatilityPredictionBand = Nothing
			Dim ThisQueueDataLast As IStockPriceVolatilityPredictionBand = Nothing

			If QueueSource.Count > 0 Then
				ThisQueueDataLast = QueueSource.Last
				ThisQueueDataLastDate = ThisQueueDataLast.StockPrice.DateDay
				With ThisQueueDataLast
					.Refresh(StockPriceVolatilityPredictionBand.StockPrice)
					If .IsBandExceeded Then
						FilterForPLE.FilterRun(1.0)
					Else
						FilterForPLE.FilterRun(0.0)
					End If
				End With
			Else
				FilterForPLE.FilterRun(0.5)
			End If
			'so  update the last PriceVolatility with the actual price
			If QueueSource.Count = MyRate Then
				ThisQueueDataRemoved = MyQueueForSigmaStatisticDaily.Dequeue
				ThisQueueDataRemovedDate = ThisQueueDataRemoved.StockPrice.DateDay
			End If
			QueueSource.Enqueue(StockPriceVolatilityPredictionBand)
		End Sub

		Private IsUseFeedbackRegulatedVolatilityFastAttackEventLocal As Boolean

		''' <summary>
		''' not in used
		''' </summary>
		''' <returns></returns>
		Public Property IsUseFeedbackRegulatedVolatilityFastAttackEvent As Boolean
			Get
				Return IsUseFeedbackRegulatedVolatilityFastAttackEventLocal
			End Get
			Set(value As Boolean)
				IsUseFeedbackRegulatedVolatilityFastAttackEventLocal = value
				MyFastAttackCountForBandExceededLow = 1
				MyFastAttackCountForBandExceededHigh = 1
			End Set
		End Property

		Public ReadOnly Property Count As Integer
			Get
				Return MyCount
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToErrorLimit As Double Implements IFilterPLLDetector.ToErrorLimit
			Get
				Return MyToErrorLimit
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToCount As Integer Implements IFilterPLLDetector.ToCount
			Get
				Return MyToCountLimitSelected
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_Count As Double Implements IFilterPLLDetector.Count
			Get
				Return MyCount
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ErrorLast As Double Implements IFilterPLLDetector.ErrorLast
			Get
				Return MyErrorLast
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ValueInit As Double Implements IFilterPLLDetector.ValueInit
			Get
				Return MyValueInit
			End Get
		End Property

		Private Function IFilterPLLDetector_ValueOutput(Input As Double, InputFeedback As Double) As Double Implements IFilterPLLDetector.ValueOutput
			Return Input + InputFeedback
		End Function

		Private ReadOnly Property IFilterPLLDetector_Status As Boolean Implements IFilterPLLDetector.Status
			Get
				Return MyStatus
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_IsMaximum As Boolean Implements IFilterPLLDetector.IsMaximum
			Get
				Return False
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_IsMinimum As Boolean Implements IFilterPLLDetector.IsMinimum
			Get
				Return False
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_Maximum As Double Implements IFilterPLLDetector.Maximum
			Get
				Return MyMaximum
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_Minimum As Double Implements IFilterPLLDetector.Minimum
			Get
				Return 0
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_DetectorBalance As Double Implements IFilterPLLDetector.DetectorBalance
			Get
				Return MyDetectorBalanceLast
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToListOfProbabilityOfExcess As IList(Of Double)
			Get
				Return MyListOfProbabilityOfExcess
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToListOfProbabilityOfExcessBalance As IList(Of Double)
			Get
				Return MyListOfProbabilityOfExcessBalance
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToList As IList(Of Double) Implements IFilterPLLDetector.ToList
			Get
				Return MyListOfValueOutput
			End Get
		End Property

		Public Shared Function VolatilityRate() As Double
			Return RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
		End Function

		Public Shared Function VolatilityRate(ByVal FilterRate As Double) As Double
			Return RATIO_OF_VOLATILITY_RATE_TO_BUFFER_SIZE * BUFFER_SIZE_FOR_VOLATILITY_FEEDBACK_STABILIZED
		End Function

		Public Property Tag As String Implements IFilterPLLDetector.Tag

		Private ReadOnly Property IFilterPLLDetector_ToListOfConvergence As IList(Of Double) Implements IFilterPLLDetector.ToListOfConvergence
			Get
				Return MyListOfConvergence
			End Get
		End Property

		Private Sub IFilterPLLDetector_RunConvergence(NumberOfIteration As Integer, ValueBegin As Double, ValueEnd As Double) Implements IFilterPLLDetector.RunConvergence
			If ValueBegin <> 0 Then
				MyListOfConvergence.Add(Math.Abs((ValueEnd - ValueBegin)) / ValueBegin)
			Else
				MyListOfConvergence.Add(0)
			End If
		End Sub

		Private ReadOnly Property IFilterPLLDetector_ToListOfVolatility As IList(Of Double) Implements IFilterPLLDetector.ToListOfVolatility
			Get
				Throw New NotImplementedException
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToListOfPriceMedianNextDayHigh As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayHigh
			Get
				Throw New NotImplementedException
			End Get
		End Property

		Private ReadOnly Property IFilterPLLDetector_ToListOfPriceMedianNextDayLow As IList(Of Double) Implements IFilterPLLDetector.ToListOfPriceMedianNextDayLow
			Get
				Throw New NotImplementedException
			End Get
		End Property
	End Class
End Namespace
