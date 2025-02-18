#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region


Namespace MathPlus.Filter
#Region "FilterLowPassExpPredict"
  ''' <summary>
  ''' This is an implementation of a double exponential filtering for simple prediction using the 
  ''' Brown's linear exponential smoothing (LES) method
  ''' See wikipedia: https://en.wikipedia.org/wiki/Exponential_smoothing#Double_exponential_smoothing
  ''' </summary>
  ''' <remarks></remarks>
  <Serializable()>
  Public Class FilterLowPassExpPredict
    Implements IFilter
    Implements IFilterControl
    Implements IFilterControlRate
    Implements IFilterPrediction
    Implements IRegisterKey(Of String)
    Implements IFilterState
    Implements IFilterEstimate


		Private MyRate As Integer
		Private MyFilterRate As Double
    Private AFilterLast As Double
    Private BFilterLast As Double
    Private ABRatio As Double
		Private FilterValueLastK1 As Double

		Private FilterValuePredictH1 As Double     'future 1 point
    Private FilterValuePredictH1Last As Double
    Private MyFilterPredictionGainYearlyLast As Double
    Private MyGainStandardDeviationLast As Double
    Private FilterValueLast As Double
    Private FilterValueLastY As Double
    Private FilterValueSlopeLastK1 As Double
    Private FilterValueSlopeLast As Double
    Private ValueLast As Double
    Private ValueLastK1 As Double
		Private MyListOfValue As ListScaled
		Private MyListOfPredictionGainPerYear As ListScaled
		Private MyListOfStatisticalVarianceError As ListScaled
		Private MyListOfAFilter As List(Of Double)
    Private MyListOfBFilter As List(Of Double)
		Private MyStatisticalForPredictionError As FilterStatistical
		Private MyStatisticalForGain As FilterStatistical
		Private MyFilter As IFilter
		Private MyFilterY As IFilter
		Private MyNumberToPredict As Double
		Private MyInputValue() As Double
		Private MyQueueForState As Queue(Of Double)


		''' <summary>
		''' Protected means that the member is accessible within the class in which it is defined and by derived classes 
		''' (i.e., subclasses). However, it is not accessible from outside the class or any other non-derived classes.
		''' </summary>
		''' <param name="NumberToPredict"></param>
		''' <param name="FilterHead"></param>
		''' <param name="FilterBase"></param>
		Protected Sub New(ByVal NumberToPredict As Double, ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
			Dim A As Double
			Dim B As Double
			MyQueueForState = New Queue(Of Double)

			'check parameter validity before proceeding
			If (FilterHead Is Nothing) Then
				'filter cannot be nothing a reference is needed for the filter rate
				Throw New ArgumentException("Invalid Filter type!")
			End If
			If TypeOf FilterHead Is IFilterControl Then
				MyFilterRate = DirectCast(FilterHead, IFilterControl).FilterRate
			Else
				MyFilterRate = FilterHead.Rate
			End If
			If MyFilterRate < 2 Then
				Throw New ArgumentException("Invalid filter rate for FilterLowPassExpPredict!")
			End If
			If (FilterBase Is Nothing) Then
				'create the default exp base filter use for the prediction
				FilterBase = New FilterExp(MyFilterRate)
			End If
			MyFilter = FilterHead
			MyFilterY = FilterBase

			MyNumberToPredict = NumberToPredict
			MyListOfValue = New YahooAccessData.ListScaled
			MyListOfPredictionGainPerYear = New YahooAccessData.ListScaled

			Dim ThisFilterRateForStatistical As Double = 5 * MyFilterRate
			If ThisFilterRateForStatistical < YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_MONTH Then
				ThisFilterRateForStatistical = YahooAccessData.MathPlus.NUMBER_TRADINGDAY_PER_MONTH
			End If

			MyStatisticalForPredictionError = New FilterStatistical(CInt(ThisFilterRateForStatistical))
			MyStatisticalForGain = New FilterStatistical(CInt(ThisFilterRateForStatistical))
			'MyListOfStatisticalExp = New FilterStatisticalExp(MyRate)
			MyListOfStatisticalVarianceError = New YahooAccessData.ListScaled
			MyListOfAFilter = New List(Of Double)
			MyListOfBFilter = New List(Of Double)

			FilterValueLast = 0
			FilterValueLastK1 = 0
			ValueLast = 0
			ValueLastK1 = 0

			A = CDbl((2 / (MyFilterRate + 1)))
			B = 1 - A
			ABRatio = A / B

			ReDim MyInputValue(-1)
		End Sub

		Public Sub New(ByVal FilterRate As Double, ByVal NumberToPredict As Double)
			Me.New(NumberToPredict:=NumberToPredict, FilterHead:=New FilterExp(FilterRate), FilterBase:=Nothing)
		End Sub

		Public Sub New(ByVal NumberToPredict As Double, ByVal FilterHead As IFilter)
			Me.New(NumberToPredict:=NumberToPredict, FilterHead:=FilterHead, FilterBase:=Nothing)
		End Sub

		''' <summary>
		''' This function can be used to change the default filter for gain calculation
		''' </summary>
		''' <param name="FilterHead"></param>
		''' <param name="FilterBase"></param>
		''' <remarks>See FilterLowPassPLLPredict for use in second order filter derivative</remarks>
		Friend Sub SetFilter(ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
			Debugger.Break()
			If FilterHead IsNot Nothing Then
				MyFilter = FilterHead
			End If
			If FilterBase IsNot Nothing Then
				MyFilterY = FilterBase
			End If
		End Sub

		''' <summary>
		''' return the filtered prediction value
		''' </summary>
		''' <param name="Value"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
			Dim Ap As Double
			Dim Bp As Double
			Dim Result As Double
			Dim ResultY As Double
			Dim ThisFilterPredictionGainYearly As Double
			Dim ThisGainStandardDeviation As Double


			FilterValueLastK1 = FilterValueLast
			Result = MyFilter.Filter(Value)
			ResultY = MyFilterY.Filter(Result)

			Ap = (2 * Result) - ResultY
			Bp = ABRatio * (Result - ResultY)
			AFilterLast = Ap
			BFilterLast = Bp
			MyListOfAFilter.Add(Ap)
			MyListOfBFilter.Add(Bp)
			FilterValuePredictH1 = Ap + Bp
			If MyListOfPredictionGainPerYear.Count = 0 Then
				FilterValuePredictH1Last = Value
			End If
			'note that B is the average trend
			FilterValueLast = Ap + Bp * MyNumberToPredict
			MyListOfValue.Add(FilterValueLast)

			MyStatisticalForGain.Filter(GainLog(Value:=FilterValuePredictH1, ValueRef:=Ap))
			ThisFilterPredictionGainYearly = MyStatisticalForGain.FilterLast.ToGaussianScale(ScaleToSignedUnit:=True)
			'ThisFilterPredictionGainYearly = Measure.Measure.ProbabilityToGaussianScale((Me.ToGaussianScale + 1, GaussianPropabilityRangeOfX)

			MyStatisticalForPredictionError.Filter(GainLog(Value:=(Value - FilterValuePredictH1Last), ValueRef:=Ap))
			ThisGainStandardDeviation = MyStatisticalForPredictionError.FilterLast.ToGaussianScale(ScaleToSignedUnit:=True)

			MyListOfPredictionGainPerYear.Add(ThisFilterPredictionGainYearly)
			MyListOfStatisticalVarianceError.Add(ThisGainStandardDeviation)

			MyFilterPredictionGainYearlyLast = ThisFilterPredictionGainYearly
			MyGainStandardDeviationLast = ThisGainStandardDeviation
			FilterValuePredictH1Last = FilterValuePredictH1
			ValueLastK1 = ValueLast
			ValueLast = Value
			Return FilterValueLast
		End Function

		Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
			Return Me.Filter(CDbl(Value.Last))
		End Function

		Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
			Dim ThisValue As Double
			For Each ThisValue In Value
				Me.Filter(ThisValue)
			Next
			Return Me.ToArray
		End Function

		Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
			Throw New NotImplementedException
		End Function

		Public Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
			Return Me
		End Function

		Private ReadOnly Property IFilterPrediction_IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
			Get
				Return True
			End Get
		End Property

		''' <summary>
		''' Calculate the future output signal from the value of the input signal at the Index given the specified gain per year.
		''' This function generally apply only for price value type signal
		''' </summary>
		''' <param name="Index"></param>
		''' <param name="NumberOfPrediction"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function IFilterPrediction_FilterPrediction(ByVal Index As Integer, ByVal NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
			Return MyListOfAFilter(Index) + MyListOfBFilter(Index) * NumberOfPrediction
		End Function

		''' <summary>
		''' Calculate the future output signal from the value of the input signal at the Index given the specified gain per year.
		''' This function generally apply only for price value type signal
		''' </summary>
		''' <param name="Index"></param>
		''' <param name="NumberOfPrediction"></param>
		''' <param name="GainPerYear"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function IFilterPrediction_FilterPrediction(ByVal Index As Integer, ByVal NumberOfPrediction As Integer, ByVal GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
			Throw New NotSupportedException
			Return MyListOfAFilter(Index) * (1 + NumberOfPrediction * (Math.Exp(GainPerYear / MathPlus.General.NUMBER_WORKDAY_PER_YEAR) - 1))
		End Function

		''' <summary>
		''' Calculate the future output signal from the last input signal and gain per year.
		''' This function generally apply only for price value type signal
		''' </summary>
		''' <param name="NumberOfPrediction"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function IFilterPrediction_FilterPrediction(ByVal NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
			Return AFilterLast + BFilterLast * NumberOfPrediction
		End Function

		''' <summary>
		''' Calculate the future output signal from the last input signal and specified gain per year.
		''' This function generally apply only for price value type  signal
		''' </summary>
		''' <param name="NumberOfPrediction"></param>
		''' <param name="GainPerYear"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function IFilterPrediction_FilterPrediction(ByVal NumberOfPrediction As Integer, ByVal GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
			Throw New NotSupportedException
			Return AFilterLast * (1 + NumberOfPrediction * (Math.Exp(GainPerYear / MathPlus.General.NUMBER_WORKDAY_PER_YEAR) - 1))
		End Function

		''' <summary>
		''' Return the logarithmic gain per year. This function is valid as long as the filter input signal is always positive >1
		''' and generally represent a price vs time function
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear() As IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
			Get
				Return MyListOfPredictionGainPerYear
			End Get
		End Property

		'not supported
		Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
			Get
				Return Nothing
			End Get
		End Property


		Public Function FilterLastToPriceVol() As YahooAccessData.IPriceVol Implements IFilter.FilterLastToPriceVol
			Dim ThisPriceVol As YahooAccessData.IPriceVol = New YahooAccessData.PriceVol(CSng(Me.FilterLast))
			With ThisPriceVol
				.LastPrevious = CSng(FilterValueLastK1)
				If Me.Last > .Last Then
					.High = CSng(Me.Last)
					.Range = YahooAccessData.RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.Last < .Last Then
					.Low = CSng(Me.Last)
					.Range = YahooAccessData.RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		Public Function LastToPriceVol() As YahooAccessData.IPriceVol Implements IFilter.LastToPriceVol
			Dim ThisPriceVol As YahooAccessData.IPriceVol = New YahooAccessData.PriceVol(CSng(Me.Last))
			With ThisPriceVol
				.LastPrevious = CSng(ValueLastK1)
				If Me.FilterLast > .Last Then
					.High = CSng(Me.FilterLast)
					.Range = YahooAccessData.RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.FilterLast < .Last Then
					.Low = CSng(Me.FilterLast)
					.Range = YahooAccessData.RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
			Return Me.Filter(CDbl(Value))
		End Function

		Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
			Throw New NotImplementedException

			Dim Ap As Double
			Dim Bp As Double
			Dim Result As Double
			Dim ResultY As Double
			Dim ThisFilterPredictionGainYearly As Double
			Dim ThisGainStandardDeviation As Double
			Dim ThisError As Double


			FilterValueLastK1 = FilterValueLast
			Result = MyFilter.Filter(Value)
			ResultY = MyFilterY.Filter(Result)

			Ap = (2 * Result) - ResultY
			Bp = ABRatio * (Result - ResultY)
			AFilterLast = Ap
			BFilterLast = Bp
			MyListOfAFilter.Add(Ap)
			MyListOfBFilter.Add(Bp)
			FilterValuePredictH1 = Ap + Bp
			If MyListOfPredictionGainPerYear.Count = 0 Then
				FilterValuePredictH1Last = Value
			End If
			FilterValueLast = Ap + Bp * MyNumberToPredict
			MyListOfValue.Add(FilterValueLast)

			ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * GainLog(FilterValuePredictH1, Ap))
			ThisError = Value - FilterValuePredictH1Last
			MyStatisticalForPredictionError.Filter(ThisError)

			'also limit exponentially the gain value between -1 and +1
			ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
			If Double.IsNaN(ThisFilterPredictionGainYearly) Then
				ThisFilterPredictionGainYearly = ThisFilterPredictionGainYearly
			End If
			MyListOfStatisticalVarianceError.Add(ThisGainStandardDeviation)
			MyListOfPredictionGainPerYear.Add(ThisFilterPredictionGainYearly)
			MyFilterPredictionGainYearlyLast = ThisFilterPredictionGainYearly
			MyGainStandardDeviationLast = ThisGainStandardDeviation
			FilterValuePredictH1Last = FilterValuePredictH1
			ValueLastK1 = ValueLast
			ValueLast = Value
			Return FilterValueLast
		End Function

		Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
			Throw New NotImplementedException
		End Function

		Public Function FilterLast() As Double Implements IFilter.FilterLast
			Return FilterValueLast
		End Function

		Public Function Last() As Double Implements IFilter.Last
			Return ValueLast
		End Function

		Public ReadOnly Property Rate As Integer Implements IFilter.Rate
			Get
				Return MyRate
			End Get
		End Property

		Public ReadOnly Property Count As Integer Implements IFilter.Count
			Get
				Return MyListOfValue.Count
			End Get
		End Property

		Public ReadOnly Property Max As Double Implements IFilter.Max
			'Performance Considerations : 
			'Every Call to Max() Is O(n) Each time you call Max(), the list Is traversed completely to find
			'the maximum value, so it performs a full scan of the list.
			'No Automatic Updates: The Max() Function does Not keep track Of changes To the list.
			'If you add, remove, Or modify elements In the list, the Max() method does Not automatically update the maximum value. It only calculates it When you Call the method.
			Get
				Return MyListOfValue.Max
			End Get
		End Property

		Public ReadOnly Property Min As Double Implements IFilter.Min
			Get
				Return MyListOfValue.Min
			End Get
		End Property

		Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
			Get
				Return MyListOfValue
			End Get
		End Property

		Public ReadOnly Property ToListScaled() As YahooAccessData.ListScaled Implements IFilter.ToListScaled
			Get
				Return MyListOfValue
			End Get
		End Property

		Public Function ToArray() As Double() Implements IFilter.ToArray
			Return MyListOfValue.ToArray
		End Function

		Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
		End Function

		Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
		End Function

		Public Property Tag As String Implements IFilter.Tag

		Public Overrides Function ToString() As String Implements IFilter.ToString
			Return $"{Me.GetType().Name}: FilterRate={MyFilterRate}"
		End Function

		Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
			Throw New NotSupportedException
			'Return ExtensionService.ShiftTo(Of Double)(MyListOfValue, MyNumberToPredict).ToArray
		End Function

		Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
			Return 0
		End Function

		''' <summary>
		''' Calculate the variance between the real data and the predicted gain value
		''' </summary>
		''' <value></value>
		''' <returns>Return the variance of the error</returns>
		''' <remarks></remarks>
		Public ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of Double) Implements IFilter.ToListOfError
			Get
				Return MyListOfStatisticalVarianceError
			End Get
		End Property
#Region "IRegisterKey"
		Public Function AsIRegisterKey() As IRegisterKey(Of String)
			Return Me
		End Function
		Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
		Dim MyKeyValue As String
		Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
			Get
				Return MyKeyValue
			End Get
			Set(value As String)
				MyKeyValue = value
			End Set
		End Property
#End Region
#Region "IFilterControl"
		Public Function ASIFilterControl() As IFilterControl Implements IFilterControl.AsIFilterControl
			Return Me
		End Function

		Private Sub IFilterControl_Clear() Implements IFilterControl.Clear
			Static IsHere As Boolean

			'no re-entrency allowed
			If IsHere Then Exit Sub
			IsHere = True

			MyListOfValue.Clear()
			MyListOfPredictionGainPerYear.Clear()
			MyStatisticalForPredictionError = New FilterStatistical(YahooAccessData.MathPlus.NUMBER_WORKDAY_PER_YEAR)
			MyListOfStatisticalVarianceError.Clear()
      MyListOfAFilter.Clear()
      MyListOfBFilter.Clear()

      FilterValueLast = 0
      FilterValueLastK1 = 0
      ValueLast = 0
      ValueLastK1 = 0
      If TypeOf MyFilter Is IFilterControl Then
        DirectCast(MyFilter, IFilterControl).Clear()
      Else
        Throw New NotSupportedException("Interface IFilterControl not availaible for this filter!")
      End If
      If TypeOf MyFilterY Is IFilterControl Then
        DirectCast(MyFilterY, IFilterControl).Clear()
      Else
        Throw New NotSupportedException("Interface IFilterControl not availaible for this filter!")
      End If
      IsHere = False
    End Sub

    Private Sub IFilterControl_Refresh(FilterRate As Double) Implements IFilterControl.Refresh
      Static IsHere As Boolean

      'no re-entrency allowed
      If IsHere Then Exit Sub
      IsHere = True

      Dim A As Double
      Dim B As Double

      If Me.Count > 0 Then
        'Clear the filter before changing the rate
        IFilterControl_Clear()
      End If
      'calculate the filter parameters
      If FilterRate < 2 Then FilterRate = 2
      MyFilterRate = FilterRate
      MyRate = CInt(FilterRate)

      A = CDbl((2 / (MyFilterRate + 1)))
      B = 1 - A
      ABRatio = A / B

      If TypeOf MyFilter Is IFilterControl Then
        DirectCast(MyFilter, IFilterControl).Refresh(MyFilterRate)
      Else
        Throw New NotSupportedException("Interface IFilterControl not availaible for this filter!")
      End If
      If TypeOf MyFilterY Is IFilterControl Then
        DirectCast(MyFilterY, IFilterControl).Refresh(MyFilterRate)
      Else
        Throw New NotSupportedException("Interface IFilterControl not availaible for this filter!")
      End If
      'reload the filter if we have the input value
      If MyInputValue.Length > 0 Then
        Me.Filter(MyInputValue)
      End If
      IsHere = False
    End Sub

    Private Sub IFilterControl_Refresh(Rate As Integer) Implements IFilterControl.Refresh
      IFilterControl_Refresh(CDbl(Rate))
    End Sub

    Public ReadOnly Property FilterRate As Double Implements IFilterControl.FilterRate
      Get
        Return MyFilterRate
      End Get
    End Property

    Private Function IFilterControl_InputValue() As Double() Implements IFilterControl.InputValue
      Return MyInputValue
    End Function

    Private ReadOnly Property IFilterControl_IsInputEnabled As Boolean Implements IFilterControl.IsInputEnabled
      Get
        Return MyInputValue.Length > 0
      End Get
    End Property
#End Region
#Region "IFilterControlRate"
    Public Function AsIFilterControlRate() As IFilterControlRate Implements IFilterControlRate.AsIFilterControlRate
      Return Me
    End Function

    Private Sub IFilterControlRate_UpdateRate(Rate As Double) Implements IFilterControlRate.UpdateRate
      Dim A As Double
      Dim B As Double

      'calculate the filter parameters
      If Rate < 2 Then Rate = 2
      MyFilterRate = FilterRate
      MyRate = CInt(FilterRate)

      A = CDbl((2 / (MyFilterRate + 1)))
      B = 1 - A
      ABRatio = A / B

      If TypeOf MyFilter Is IFilterControlRate Then
        DirectCast(MyFilter, IFilterControlRate).UpdateRate(MyFilterRate)
      Else
        Throw New NotSupportedException("Interface IFilterControlRate not availaible for this filter!")
      End If
      If TypeOf MyFilterY Is IFilterControlRate Then
        DirectCast(MyFilterY, IFilterControlRate).UpdateRate(MyFilterRate)
      Else
        Throw New NotSupportedException("Interface IFilterControlRate not availaible for this filter!")
      End If
    End Sub

    Private Sub IFilterControlRate_UpdateRate(Rate As Integer) Implements IFilterControlRate.UpdateRate
      IFilterControlRate_UpdateRate(CDbl(Rate))
    End Sub

    Private Property IFilterControlRate_Enabled As Boolean Implements IFilterControlRate.Enabled
      'always true here
      Get
        Return True
      End Get
      Set(value As Boolean)

      End Set
    End Property
#End Region

#Region "IFilterState"
    Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
      Return Me
    End Function

    Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
      Dim ThisCount As Integer

      If MyQueueForState.Count = 0 Then Return

      AFilterLast = MyQueueForState.Dequeue
      BFilterLast = MyQueueForState.Dequeue
      ABRatio = MyQueueForState.Dequeue
      FilterValueLastK1 = MyQueueForState.Dequeue
      FilterValuePredictH1 = MyQueueForState.Dequeue
      FilterValuePredictH1Last = MyQueueForState.Dequeue
      MyFilterPredictionGainYearlyLast = MyQueueForState.Dequeue
      MyGainStandardDeviationLast = MyQueueForState.Dequeue
      FilterValueLast = MyQueueForState.Dequeue
      FilterValueLastY = MyQueueForState.Dequeue
      FilterValueLastY = MyQueueForState.Dequeue
      FilterValueSlopeLastK1 = MyQueueForState.Dequeue
      FilterValueSlopeLast = MyQueueForState.Dequeue
      ValueLast = MyQueueForState.Dequeue
      ValueLastK1 = MyQueueForState.Dequeue
      ThisCount = CInt(MyQueueForState.Dequeue)
      If MyListOfValue.Count > ThisCount Then
        Do
          MyListOfPredictionGainPerYear.RemoveAt(MyListOfPredictionGainPerYear.Count - 1)
          MyListOfStatisticalVarianceError.RemoveAt(MyListOfStatisticalVarianceError.Count - 1)
          MyListOfAFilter.RemoveAt(MyListOfAFilter.Count - 1)
          MyListOfBFilter.RemoveAt(MyListOfBFilter.Count - 1)
          MyListOfValue.RemoveAt(MyListOfValue.Count - 1)
        Loop Until MyListOfValue.Count = ThisCount
      End If
    End Sub

    Private Sub IFilterState_Save() Implements IFilterState.Save
      MyQueueForState.Enqueue(AFilterLast)
      MyQueueForState.Enqueue(BFilterLast)

      MyQueueForState.Enqueue(ABRatio)
      MyQueueForState.Enqueue(FilterValueLastK1)
      MyQueueForState.Enqueue(FilterValuePredictH1)

      MyQueueForState.Enqueue(FilterValuePredictH1Last)
      MyQueueForState.Enqueue(MyFilterPredictionGainYearlyLast)
      MyQueueForState.Enqueue(MyGainStandardDeviationLast)
      MyQueueForState.Enqueue(FilterValueLast)
      MyQueueForState.Enqueue(FilterValueLastY)
      MyQueueForState.Enqueue(FilterValueLastY)
      MyQueueForState.Enqueue(FilterValueSlopeLastK1)
      MyQueueForState.Enqueue(FilterValueSlopeLast)
      MyQueueForState.Enqueue(ValueLast)
      MyQueueForState.Enqueue(ValueLastK1)
      MyQueueForState.Enqueue(Me.Count)
    End Sub
#End Region
#Region "IFilterEstimate"
    Public Function AsIFilterEstimate() As IFilterEstimate Implements IFilterEstimate.AsIFilterEstimate
      Return Me
    End Function

    Public Function IFilterEstimate_Filter(Value As Double) As IFilterEstimateResult Implements IFilterEstimate.Filter
      Dim Ap As Double
      Dim Bp As Double
      Dim ThisResult As IFilterEstimateResult = Nothing
      Dim ThisResultY As IFilterEstimateResult = Nothing
      Dim ThisFilterPredictionGainYearly As Double
      Dim ThisFilterEstimate As IFilterEstimate
      Dim ThisFilterEstimateY As IFilterEstimate
      Dim ThisFilterResultLast As Double
      Dim ThisFilterValuePredictH1 As Double

			Throw New NotSupportedException
			Debugger.Break()
			If (TypeOf MyFilter Is IFilterEstimate = False) Or (TypeOf MyFilterY Is IFilterEstimate = False) Then
				Throw New NotSupportedException("The Filter does not support the required interface!")
			Else
				ThisFilterEstimate = DirectCast(MyFilter, IFilterEstimate)
        ThisFilterEstimateY = DirectCast(MyFilterY, IFilterEstimate)
      End If
      Select Case (MyFilter.Count - MyFilterY.Count)
        Case 0
          ThisResult = ThisFilterEstimate.Filter(Value)
          ThisResultY = ThisFilterEstimateY.Filter(ThisResult.Value)
        Case 1
          ThisResult = New FilterEstimateResult(MyFilter.FilterLast)
          ThisResultY = ThisFilterEstimateY.Filter(MyFilter.FilterLast)
        Case Is > 1, Is < 0
          Throw New ArgumentOutOfRangeException("Invalid Filter count!")
      End Select
      Ap = (2 * ThisResult.Value) - ThisResultY.Value
      Bp = ABRatio * (ThisResult.Value - ThisResultY.Value)
      ThisFilterValuePredictH1 = Ap + Bp
      ThisFilterResultLast = Ap + Bp * MyNumberToPredict
      ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * GainLog(ThisFilterValuePredictH1, Ap))
      'also limit exponentially the gain value between -1 and +1
      ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
      Return New FilterEstimateResult(Value:=ThisFilterResultLast, Gain:=ThisFilterPredictionGainYearly, GainDerivative:=0.0)
    End Function

    Public Function IFilterEstimate_Filter(Value() As Double) As System.Collections.Generic.IList(Of IFilterEstimateResult) Implements IFilterEstimate.Filter
      Dim Ap As Double
      Dim Bp As Double
      Dim ThisResult As IFilterEstimateResult = Nothing
      Dim ThisResultY As IFilterEstimateResult = Nothing
      Dim ThisFilterPredictionGainYearly As Double
      Dim ThisFilterEstimate As IFilterEstimate
      Dim ThisFilterEstimateY As IFilterEstimate
      Dim ThisFilterResultLast As Double
      Dim ThisList As New List(Of IFilterEstimateResult)
      Throw New NotImplementedException
      If (TypeOf MyFilter Is IFilterEstimate = False) Or (TypeOf MyFilterY Is IFilterEstimate = False) Then
        Throw New NotSupportedException("The Filter does not support the required interface!")
      Else
        ThisFilterEstimate = DirectCast(MyFilter, IFilterEstimate)
        ThisFilterEstimateY = DirectCast(MyFilterY, IFilterEstimate)
      End If
      Dim ThisListOfFilterEstimateResult = ThisFilterEstimate.Filter(Value)
      Dim ThisListOfFilterEstimateResultY = ThisFilterEstimate.Filter(Value)
      Dim I As Integer
      For I = 0 To ThisListOfFilterEstimateResult.Count - 1
        ThisResult = ThisListOfFilterEstimateResult(I)
        ThisResultY = ThisListOfFilterEstimateResultY(I)
        Ap = (2 * ThisResult.Value) - ThisResultY.Value
        Bp = ABRatio * (ThisResult.Value - ThisResultY.Value)
        ThisFilterResultLast = Ap + Bp * MyNumberToPredict
        ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * GainLog(FilterValuePredictH1, Ap))
        'ThisFilterPredictionGainYearly = (MathPlus.General.NUMBER_WORKDAY_PER_YEAR * Math.Log(((FilterValuePredictH1 ^ 2 + 1) / (Ap ^ 2 + 1)))) / 2
        'also limit exponentially the gain value between -1 and +1
        ThisFilterPredictionGainYearly = MathPlus.WaveForm.SignalLimit(ThisFilterPredictionGainYearly, 1)
      Next
      'Return New FilterEstimateResult(Value:=ThisFilterResultLast, Gain:=ThisFilterPredictionGainYearly, GainDerivative:=0.0)
    End Function
#End Region
  End Class
#End Region
End Namespace
