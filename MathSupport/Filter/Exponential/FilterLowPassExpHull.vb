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
  Public Class FilterLowPassExpHull
    Implements IFilter
    Implements IFilterPrediction
    Implements IFilterControlRate
    Implements IRegisterKey(Of String)
    Implements IFilterCreateNew

    Private MyRate As Integer
    Private A As Double
    Private B As Double
    Private FilterValueLastK1 As Double
    Private FilterValueLast As Double
    Private ValueLast As Double
    Private ValueLastK1 As Double
    'Private MyValueSumForInit As Double
    Private MyListOfValue As ListScaled
    Private MyFilter1 As FilterLowPassExp
    Private MyFilter2 As FilterLowPassExp
    Private MyFilter3 As FilterLowPassExp
    Private MyFilterPrediction As Filter.FilterLowPassExpPredict
    Private MyFilterPredictionDerivative As Filter.FilterLowPassPLLPredict
    Private MyFilterRate As Double
    Private IsPredictionEnabledLocal As Boolean


    Public Sub New(ByVal FilterRate As Double, Optional IsPredictionEnabled As Boolean = False)
      MyListOfValue = New ListScaled
      If FilterRate < 1 Then FilterRate = 1
      MyRate = CInt(FilterRate)
      MyFilterRate = FilterRate
      FilterValueLast = 0
      FilterValueLastK1 = 0
      ValueLast = 0
      ValueLastK1 = 0
      'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
      'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
      'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
      A = CDbl((2 / (FilterRate + 1)))

      'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
      B = 1 - A
      MyFilter1 = New FilterLowPassExp(FilterRate / 2)
      MyFilter2 = New FilterLowPassExp(FilterRate)
      MyFilter3 = New FilterLowPassExp(Math.Sqrt(FilterRate))
      If IsPredictionEnabled Then
        MyFilterPrediction = New Filter.FilterLowPassExpPredict(
          NumberToPredict:=0.0,
          FilterHead:=New FilterLowPassExpHull(FilterRate:=MyFilterRate))

        MyFilterPredictionDerivative = New Filter.FilterLowPassPLLPredict(
          NumberToPredict:=0.0,
          FilterHead:=New FilterLowPassExpHull(FilterRate:=MyFilterRate))
      Else
        MyFilterPrediction = Nothing
        MyFilterPredictionDerivative = Nothing
      End If
      IsPredictionEnabledLocal = IsPredictionEnabled
    End Sub

    Public Sub New(ByVal FilterRate As Integer, Optional IsPredictionEnabled As Boolean = False)
      Me.New(CDbl(FilterRate), IsPredictionEnabled)
    End Sub

    Public Function AsIFilterCreateNew() As IFilterCreateNew Implements IFilterCreateNew.AsIFilterCreateNew
      Return Me
    End Function

    Private Function IFilterCreateNew_CreateNew() As IFilter Implements IFilterCreateNew.CreateNew
      Dim ThisFilter As IFilter

      ThisFilter = New FilterLowPassExp(FilterRate:=MyFilterRate, IsPredictionEnabled:=IsPredictionEnabledLocal)
      Return ThisFilter
    End Function


    Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
      If MyListOfValue.Count = 0 Then
        FilterValueLast = Value
      End If
      FilterValueLastK1 = FilterValueLast
      MyFilter1.Filter(Value)
      MyFilter2.Filter(Value)
      FilterValueLast = MyFilter3.Filter(2 * MyFilter1.FilterLast - MyFilter2.FilterLast)
      MyListOfValue.Add(FilterValueLast)
      ValueLastK1 = ValueLast
      ValueLast = Value
      If MyFilterPrediction IsNot Nothing Then
        MyFilterPrediction.Filter(Value)
        MyFilterPredictionDerivative.Filter(Value)
      End If
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

    ''' <summary>
    ''' Special filtering that can be used to remove the delay starting at a specific point
    ''' </summary>
    ''' <param name="Value">The value to be filtered</param>
    ''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
    ''' <returns>The result</returns>
    ''' <remarks></remarks>
    Public Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
      Throw New NotImplementedException
    End Function

    Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
      Return 0.0
    End Function

    Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
      Throw New NotImplementedException
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

    Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
      Return CSng(Me.Filter(CDbl(Value)))
    End Function

    Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
      Throw New NotImplementedException
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
      Return Me.FilterLast.ToString
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

    Public ReadOnly Property ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
      Get
        If MyFilterPrediction Is Nothing Then
          Return Nothing
        Else
          Return MyFilterPredictionDerivative.AsIFilterPrediction.ToListOfGainPerYear
        End If
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
#Region "IFilterControlRate"
    Public Function AsIFilterControlRate() As IFilterControlRate Implements IFilterControlRate.AsIFilterControlRate
      Return Me
    End Function

    Private Sub IFilterControlRate_UpdateRate(FilterRate As Double) Implements IFilterControlRate.UpdateRate
      'calculate the new filter parameters
      If FilterRate < 1 Then FilterRate = 1
      MyRate = CInt(FilterRate)
      MyFilterRate = FilterRate

      'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
      'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
      'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
      A = CDbl((2 / (FilterRate + 1)))

      'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
      B = 1 - A
      MyFilter1.AsIFilterControlRate.UpdateRate(FilterRate / 2)
      MyFilter2.AsIFilterControlRate.UpdateRate(FilterRate)
      MyFilter3.AsIFilterControlRate.UpdateRate(Math.Sqrt(FilterRate))
      If MyFilterPrediction IsNot Nothing Then
        MyFilterPrediction.AsIFilterControlRate.UpdateRate(FilterRate)
        MyFilterPredictionDerivative.AsIFilterControlRate.UpdateRate(FilterRate)
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
  End Class
End Namespace

