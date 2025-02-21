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
  <Serializable()>
  Public Class FilterLowPassExp
    Implements IFilter
    Implements IFilterPrediction
    Implements IFilterControl
    Implements IFilterControlRate
    Implements IRegisterKey(Of String)
    Implements IFilterCopy
		Implements IFilterCreateNew
		Implements IFilterState

		Private MyRate As Integer
		Private MyFilterRate As Double
		Private IsPredictionEnabledLocal As Boolean
		Private A As Double
		Private B As Double
		Private FilterValueLastK1 As Double
		Private FilterValueLast As Double
		Private ValueLast As Double
		Private ValueLastK1 As Double
		'Private MyValueSumForInit As Double
		Private IsRunFilterLocal As Boolean
		Private IsValueInitial As Boolean
		Private MyListOfValue As ListScaled
		Private MyFilterPrediction As Filter.FilterLowPassExpPredict
		Private MyFilterPredictionDerivative As Filter.FilterLowPassPLLPredict
		Private MyInputValue() As Double

		Public Sub New(ByVal FilterRate As Double, Optional IsPredictionEnabled As Boolean = False)
			MyListOfValue = New ListScaled

			If FilterRate < 1 Then FilterRate = 1
			MyFilterRate = FilterRate
			MyRate = CInt(MyFilterRate)
			IsPredictionEnabledLocal = IsPredictionEnabled

			'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
			'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
			'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
			A = CDbl((2 / (MyFilterRate + 1)))

			'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
			B = 1 - A
			IsValueInitial = False
			FilterValueLast = 0
			FilterValueLastK1 = 0
			ValueLast = 0
			ValueLastK1 = 0

			If IsPredictionEnabledLocal Then
				'this is base on a exponential filter
				MyFilterPrediction = New Filter.FilterLowPassExpPredict(
					NumberToPredict:=0.0,
					FilterHead:=New FilterExp(FilterRate:=MyFilterRate))

				MyFilterPredictionDerivative = New Filter.FilterLowPassPLLPredict(
					NumberToPredict:=0.0,
					FilterHead:=New FilterExp(FilterRate:=MyFilterRate),
					FilterBase:=New FilterPLL(FilterRate:=MyFilterRate))
			Else
				MyFilterPrediction = Nothing
				MyFilterPredictionDerivative = Nothing
			End If
			IsRunFilterLocal = False
			ReDim MyInputValue(-1)
		End Sub

		Public Sub New(ByVal FilterRate As Integer, Optional IsPredictionEnabled As Boolean = False)
			Me.New(CDbl(FilterRate), IsPredictionEnabled)
		End Sub

		Public Sub New(ByVal FilterRate As Double, ByRef InputValue() As Double, Optional IsPredictionEnabled As Boolean = False)
			Me.New(FilterRate, IsPredictionEnabled)
			ReDim MyInputValue(0 To InputValue.Length - 1)
			InputValue.CopyTo(MyInputValue, 0)
			Me.Filter(MyInputValue)
		End Sub

		Friend Sub New(ByVal FilterRate As Double, ByRef InputValue() As Double, ByVal IsRunFilter As Boolean, Optional IsPredictionEnabled As Boolean = False)
			Me.New(FilterRate, IsPredictionEnabled)
			ReDim MyInputValue(0 To InputValue.Length - 1)
			InputValue.CopyTo(MyInputValue, 0)
			If IsRunFilter Then
				IsRunFilterLocal = True
				Me.Filter(MyInputValue, MyInputValue.Length - 1)
			End If
		End Sub

		Public Sub New(ByVal FilterRate As Integer, ByRef InputValue() As Double, Optional IsPredictionEnabled As Boolean = False)
			Me.New(CDbl(FilterRate), InputValue, IsPredictionEnabled)
		End Sub

		Public Function AsIFilterCreateNew() As IFilterCreateNew Implements IFilterCreateNew.AsIFilterCreateNew
			Return Me
		End Function

		Private Function IFilterCreateNew_CreateNew() As IFilter Implements IFilterCreateNew.CreateNew
			Dim ThisFilter As IFilter

			If MyInputValue.Length > -1 Then
				ThisFilter = New FilterLowPassExp(FilterRate:=MyFilterRate, InputValue:=MyInputValue, IsRunFilter:=IsRunFilterLocal, IsPredictionEnabled:=IsPredictionEnabledLocal)
			Else
				ThisFilter = New FilterLowPassExp(FilterRate:=MyFilterRate, IsPredictionEnabled:=IsPredictionEnabledLocal)
			End If
			Return ThisFilter
		End Function


		'Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Double, Optional IsPredictionEnabled As Boolean = False)
		'  Me.New(FilterRate, IsPredictionEnabled)
		'  FilterValueLast = ValueInitial
		'  FilterValueLastK1 = FilterValueLast
		'  ValueLast = ValueInitial
		'  ValueLastK1 = ValueLast
		'  IsValueInitial = True
		'End Sub

		'Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Single, Optional IsPredictionEnabled As Boolean = False)
		'  Me.New(FilterRate, CDbl(ValueInitial), IsPredictionEnabled)
		'End Sub

		''' <summary>
		''' Function accessible by other inherited class allowing backgroud data loading for special filtering 
		''' </summary>
		''' <param name="Value"></param>
		''' <param name="ValueFiltered"></param>
		''' <remarks></remarks>
		Friend Sub Filter(ByVal Value As Double, ValueFiltered As Double)
			If MyListOfValue.Count = 0 Then
				'initialization
				FilterValueLast = Value
			End If
			FilterValueLastK1 = FilterValueLast
			FilterValueLast = ValueFiltered
			MyListOfValue.Add(FilterValueLast)
			ValueLastK1 = ValueLast
			ValueLast = Value
			If MyFilterPrediction IsNot Nothing Then
				MyFilterPrediction.Filter(Value)
				MyFilterPredictionDerivative.Filter(Value)
			End If
		End Sub

		Public Overridable Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual <> ThisResultPrediction Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
			If MyListOfValue.Count = 0 Then
				'initialization
				FilterValueLast = Value
			End If
			FilterValueLastK1 = FilterValueLast
			FilterValueLast = A * Value + B * FilterValueLast
			MyListOfValue.Add(FilterValueLast)
			ValueLastK1 = ValueLast
			ValueLast = Value
			If MyFilterPrediction IsNot Nothing Then
				MyFilterPrediction.Filter(Value)
				MyFilterPredictionDerivative.Filter(Value)
			End If
			Return FilterValueLast
		End Function

		Public Overridable Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
			Return CSng(Me.Filter(CDbl(Value)))
		End Function

		Public Overridable Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
			Return Me.Filter(CDbl(Value.Last))
		End Function

		Public Overridable Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
			Dim ThisValue As Double
			For Each ThisValue In Value
				Me.Filter(ThisValue)
			Next
			Return Me.ToArray
		End Function

		''' <summary>
		''' Special filtering that can be used to remove the delay starting at a specific point
		''' </summary>
		''' <param name="Value">The value to be filtered</param>
		''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
		''' <returns>The result</returns>
		''' <remarks></remarks>
		Public Overridable Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
			Dim ThisValues(0 To Value.Length - 1) As Double
			Dim I As Integer
			Dim J As Integer

			Dim ThisFilterLeft As New FilterLowPassExp(Me.Rate)
			Dim ThisFilterRight As New FilterLowPassExp(Me.Rate)
			Dim ThisFilterLeftItem As Double
			Dim ThisFilterRightItem As Double

			'filter from the left
			ThisFilterLeft.Filter(Value)
			'filter from the right the section with the reverse filtering
			For I = DelayRemovedToItem To 0 Step -1
				ThisFilterRight.Filter(Value(I))
			Next
			'the data in ThisFilterRightList is reversed
			'need to look at it in reverse order using J
			J = DelayRemovedToItem
			If MyFilterPrediction Is Nothing Then
				For I = 0 To Value.Length - 1
					FilterValueLastK1 = FilterValueLast
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						FilterValueLast = ThisFilterLeftItem
					Else
						ThisFilterRightItem = ThisFilterRight.ToList(J)
						FilterValueLast = (ThisFilterLeftItem + ThisFilterRightItem) / 2
					End If
					MyListOfValue.Add(FilterValueLast)
					ThisValues(I) = FilterValueLast
					J = J - 1
				Next
			Else
				For I = 0 To Value.Length - 1
					FilterValueLastK1 = FilterValueLast
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						FilterValueLast = ThisFilterLeftItem
					Else
						ThisFilterRightItem = ThisFilterRight.ToList(J)
						FilterValueLast = (ThisFilterLeftItem + ThisFilterRightItem) / 2
					End If
					MyListOfValue.Add(FilterValueLast)
					ThisValues(I) = FilterValueLast
					MyFilterPrediction.Filter(Value(I))
					MyFilterPredictionDerivative.Filter(Value(I))
					J = J - 1
				Next
			End If
			Return ThisValues
		End Function

		Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
			Return 0.0
		End Function

		Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
			If A > 0 Then
				Return (Value - B * FilterValueLastK1) / A
			Else
				Return Value
			End If
		End Function

		Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
			Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
			With ThisPriceVol
				.LastPrevious = CSng(FilterValueLastK1)
				If Me.Last > .Last Then
					.High = CSng(Me.Last)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.Last < .Last Then
					.Low = CSng(Me.Last)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
			Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
			With ThisPriceVol
				.LastPrevious = CSng(ValueLastK1)
				If Me.FilterLast > .Last Then
					.High = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				ElseIf Me.FilterLast < .Last Then
					.Low = CSng(Me.FilterLast)
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				End If
			End With
			Return ThisPriceVol
		End Function

		''' <summary>
		''' Calculate the filter next value input for a given output value. The filter memory is not affected. 
		''' The function is useful to evaluate the input value required for a given filter output value. 
		''' </summary>
		''' <param name="Value">The filter output value</param>
		''' <returns>The input value</returns>
		Public Function FilterPredictionInverseNext(ByVal Value As Double) As Double
			Dim ThisFilterValueLast As Double = FilterValueLast
			If MyListOfValue.Count = 0 Then
				'initialization
				ThisFilterValueLast = Value
				'note by definition: B = 1 - A
			End If
			ThisFilterValueLast = (Value - B * ThisFilterValueLast) / A
			Return ThisFilterValueLast
		End Function

		''' <summary>
		''' Calculate the filter next value for a given input without changing the current filter memory
		''' </summary>
		''' <param name="Value"></param>
		''' <returns></returns>
		Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
			Dim ThisFilterValueLast As Double = FilterValueLast
			If MyListOfValue.Count = 0 Then
				'initialization
				ThisFilterValueLast = Value
			End If
			ThisFilterValueLast = A * Value + B * ThisFilterValueLast
			Return ThisFilterValueLast
		End Function

		Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
			Return Me.FilterPredictionNext(CDbl(Value))
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

		Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
			Get
				Throw New NotSupportedException
			End Get
		End Property

		Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
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
			Return $"{Me.GetType().Name}: FilterRate={MyFilterRate},{Me.FilterLast}"
		End Function

#Region "IFilterPrediction"
		Public Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
			Return Me
		End Function

		Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
			If MyFilterPrediction Is Nothing Then
				Return Me.FilterLast
			Else
				Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(NumberOfPrediction)
			End If
		End Function

		Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
			If MyFilterPrediction Is Nothing Then
				Return Me.FilterLast
			Else
				Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(NumberOfPrediction, GainPerYear)
			End If
		End Function

		Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
			If MyFilterPrediction Is Nothing Then
				Return Me.FilterLast
			Else
				Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction)
			End If
		End Function

		Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
			If MyFilterPrediction Is Nothing Then
				Return Me.FilterLast
			Else
				Return MyFilterPrediction.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction, GainPerYear)
			End If
		End Function

		Private ReadOnly Property IFilterPrediction_IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
			Get
				If MyFilterPrediction Is Nothing Then
					Return False
				Else
					Return True
				End If
			End Get
		End Property

		Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
			Get
				If MyFilterPrediction Is Nothing Then
					Return Nothing
				Else
					Return MyFilterPrediction.AsIFilterPrediction.ToListOfGainPerYear
				End If
			End Get
		End Property

		Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
			Get
				Return MyFilterPredictionDerivative.AsIFilterPrediction.ToListOfGainPerYear
			End Get
		End Property
#End Region
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
			FilterValueLast = 0
			FilterValueLastK1 = 0
			ValueLast = 0
			ValueLastK1 = 0
			IsValueInitial = False
			If IsPredictionEnabledLocal Then
				MyFilterPrediction.ASIFilterControl.Clear()
				MyFilterPredictionDerivative.ASIFilterControl.Clear()
			End If
			IsHere = False
		End Sub

		Private Sub IFilterControl_Refresh(FilterRate As Double) Implements IFilterControl.Refresh
			Static IsHere As Boolean

			'no re-entrency allowed
			If IsHere Then Exit Sub
			IsHere = True

			If Me.Count > 0 Then
				'Clear the filter before changing the rate
				IFilterControl_Clear()
			End If
			'calculate the new filter parameters
			If FilterRate < 1 Then FilterRate = 1
			MyFilterRate = FilterRate
			MyRate = CInt(MyFilterRate)

			'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
			'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
			'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
			A = CDbl((2 / (MyFilterRate + 1)))

			'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
			B = 1 - A
			If IsPredictionEnabledLocal Then
				MyFilterPrediction.ASIFilterControl.Refresh(MyFilterRate)
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

		Private ReadOnly Property IFilterControl_FilterRate As Double Implements IFilterControl.FilterRate
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
			'calculate the new filter parameters
			If Rate < 1 Then Rate = 1
			MyFilterRate = Rate
			MyRate = CInt(MyFilterRate)

			'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
			'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
			'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
			A = CDbl((2 / (MyFilterRate + 1)))

			'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
			B = 1 - A
			If IsPredictionEnabledLocal Then
				MyFilterPrediction.ASIFilterControl.Refresh(MyFilterRate)
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
#Region "IFilterCopy"
		Public Function AsIFilterCopy() As IFilterCopy Implements IFilterCopy.AsIFilterCopy
			Return Me
		End Function

		Private Function IFilterCopy_CopyFrom() As IFilter Implements IFilterCopy.CopyFrom
			Dim ThisFilter As FilterLowPassExp

			Throw New NotSupportedException
			If MyInputValue.Length > 0 Then
				If Me.Count = 0 Then
					ThisFilter = New FilterLowPassExp(MyFilterRate, MyInputValue, IsRunFilter:=False, IsPredictionEnabled:=IsPredictionEnabledLocal)
				Else
					ThisFilter = New FilterLowPassExp(MyFilterRate, MyInputValue, IsPredictionEnabledLocal)
				End If
			Else
				ThisFilter = New FilterLowPassExp(MyFilterRate, IsPredictionEnabledLocal)
			End If
			Return ThisFilter
		End Function
#End Region
		Protected Overrides Sub Finalize()
			MyBase.Finalize()
		End Sub

#Region "IFilterState"
		Public Function ASIFilterState() As IFilterState Implements IFilterState.ASIFilterState
			Return Me
		End Function

		Private MyQueueForState As New Queue(Of Double)
		Private Sub IFilterState_ReturnPrevious() Implements IFilterState.ReturnPrevious
			Dim ThisCount As Integer

			Throw New NotImplementedException($"Function IFilterState_ReturnPrevious in {Me.GetType().Name} is not supported...")
			Try
				If MyQueueForState.Count = 0 Then Return
				ThisCount = CInt(MyQueueForState.Dequeue)
				If MyListOfValue.Count > ThisCount Then
					Do
						MyListOfValue.RemoveAt(MyListOfValue.Count - 1)
					Loop Until MyListOfValue.Count = ThisCount
				End If
			Catch ex As InvalidOperationException
				' Handle error, perhaps log it or rethrow with additional info
				Throw New Exception($"Failed to restore state from queue in {Me.GetType().Name}. Queue may be empty or corrupted.", ex)
			End Try
		End Sub

		Private Sub IFilterState_Save() Implements IFilterState.Save
			Throw New NotImplementedException($"Function IFilterState_Save in {Me.GetType().Name} is not supported...")
			MyQueueForState.Enqueue(FilterValueLast)
			MyQueueForState.Enqueue(Me.Count)
		End Sub
#End Region
	End Class
End Namespace