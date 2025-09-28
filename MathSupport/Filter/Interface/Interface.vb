#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region

Namespace MathPlus
  Namespace Filter
    Public Interface IStatisticalDistribution
      ''' <summary>
      ''' ArrayStandard is a local standard starting with the element zero as the number of point 
      ''' following by all the data from the minimum to the maximum bins
      ''' </summary>
      ''' <remarks></remarks>
      Enum enuRefreshType
        PerCent
        NumberOfPoint
        ArrayStandard
      End Enum

      ReadOnly Property AsIStatisticalDistribution As IStatisticalDistribution
      ReadOnly Property Mean As Double
      ReadOnly Property MeanOfSquare As Double
      ReadOnly Property StandardDeviation As Double
      ReadOnly Property Variance As Double
      ReadOnly Property NumberPoint As Integer
      ReadOnly Property ToListOfPDF As IList(Of Double)
      ReadOnly Property ToListOfLCR As IList(Of Double)
      ReadOnly Property ToListOfCDF As IList(Of Double)
      ReadOnly Property ToListOfFD As IList(Of Double)
      ReadOnly Property BucketLimitLow As Integer
      ReadOnly Property BucketLimitHigh As Integer
      ReadOnly Property BucketHigh As Integer
      ReadOnly Property BucketLow As Integer
      ReadOnly Property NumberOfBucket As Integer
      Property Tag As String
      Sub BucketFill(ByVal Value As Double)
      Sub Refresh(Optional Type As enuRefreshType = enuRefreshType.PerCent)
    End Interface

    Public Interface IStatisticalDistributionFunction
      Function ToBucket(Value As Double) As Integer
      Function FromBucket(Index As Integer) As Double
      ReadOnly Property BucketLimitLow As Integer
      ReadOnly Property BucketLimitHigh As Integer
      ReadOnly Property BucketLow As Integer
      ReadOnly Property BucketHigh As Integer
      ReadOnly Property NumberBucket As Integer
    End Interface

		Public Interface IListWindowsFrame(Of T As {Structure})
			ReadOnly Property AsIListWindowsFrame As IListWindowsFrame(Of T)
			ReadOnly Property ItemLowIndex() As Integer
			ReadOnly Property ItemHighIndex() As Integer
			Function ItemLow() As Nullable(Of T)
			Function ItemHigh() As Nullable(Of T)
			Function ItemFirst() As Nullable(Of T)
			Function ItemLast() As Nullable(Of T)
			Function ItemDecimate() As Nullable(Of T)
			ReadOnly Property WindowSize As Integer
			Function ItemRemoved() As Nullable(Of T)
		End Interface

		Public Interface IListWindowsFrame1(Of T)
			ReadOnly Property AsIListWindowsFrame1 As IListWindowsFrame1(Of T)
			ReadOnly Property ItemLowIndex() As Integer
			ReadOnly Property ItemHighIndex() As Integer
			Function ItemLow() As T
			Function ItemHigh() As T
			Function ItemFirst() As T
			Function ItemLast() As T
			Function ItemDecimate() As T
			ReadOnly Property WindowSize As Integer
			Function ItemRemoved() As T
		End Interface



		Friend Interface IFilterData
			Property FilterInput As Double
			Property FilterLast As Double
			Property Range As Double
		End Interface

		Public Interface IFilterRateInterpolated
      Enum enuOutputType
        Standard
        GainPerYear
        GainPerYearDerivative
        NotDefined
      End Enum

      ReadOnly Property OutputType As enuOutputType
      ReadOnly Property Rate As Double
      ReadOnly Property RateMinimum As Double
      ReadOnly Property RateMaximum As Double
      Sub Refresh(ByVal Rate As Double)
      ReadOnly Property ToList() As IList(Of Double)
      ReadOnly Property ToList(ByVal Rate As Double) As IList(Of Double)
      Function CopyFrom() As IFilterRateInterpolated
      ''' <summary>
      ''' Calculate the filter rate based on a value in percent from 0 and 1
      ''' </summary>
      ''' <param name="Value">The value in Percent </param>
      ''' <returns>The filter rate</returns>
      ''' <remarks></remarks>
      Function ToFilterRate(ByVal Value As Double) As Double
      ''' <summary>
      ''' Calculate the filter Rate in percent base on an input filter rate value
      ''' </summary>
      ''' <param name="Value"></param>
      ''' <returns>The value in Percent between 0.0 and 1.0</returns>
      ''' <remarks>Useful for some scaling</remarks>
      Function ToFilterRateInverse(ByVal Value As Double) As Double
    End Interface
    Public Interface ITransaction
      Enum enuTransactionType
        StockBuy
        StockSell
        StockOptionCall
        StockOptionPut
      End Enum

      Function Filter(ByRef Value As IPriceVol) As Double
      Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As Double) As Double
      Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As IPriceVol) As Double
      ReadOnly Property Type As enuTransactionType
      ReadOnly Property PriceTransactionStart As Double
      ReadOnly Property PriceTransactionStop As Double
      ReadOnly Property TransactionCost As Double
      ReadOnly Property TransactionCount As Integer
      ReadOnly Property IsStop As Boolean
      Property IsStopEnabled As Boolean
      Property IsStopReverse As Boolean
      ReadOnly Property PriceStop As Double
      ReadOnly Property PriceStopValue As IPriceVol
      Function FilterLast() As Double
      Function Last() As IPriceVol
      ReadOnly Property Count As Integer
      ReadOnly Property CountStop As Integer
      ReadOnly Property Max As Double
      ReadOnly Property Min As Double
      ReadOnly Property ToList() As IList(Of Double)
      Property Tag As String
    End Interface

    Public Interface IFilterPivotPointList
      ReadOnly Property AsIFilterPivotPointList As IFilterPivotPointList
      ReadOnly Property ToListOfHighIndex As IList(Of Integer)
      ReadOnly Property ToListOfLowIndex As IList(Of Integer)
      ReadOnly Property ToListOfLowDistance As IList(Of Double)
      ReadOnly Property ToListOfHighDistance As IList(Of Double)
      ReadOnly Property ToListOfCompositeDistance As IList(Of Double)
      ReadOnly Property ToListOfFilteredCompositeDistance As IList(Of Double)
      ReadOnly Property ToListOfPriceGain As IList(Of Double)
      ReadOnly Property ToListOfPriceGainDerivative As IList(Of Double)
      ReadOnly Property ToListOfPriceGainDifference As IList(Of Double)
      ReadOnly Property ToListOfPriceGainDerivativeDifference As IList(Of Double)
      ReadOnly Property HighIndexForeCast As Nullable(Of Integer)
      ReadOnly Property LowIndexForeCast As Nullable(Of Integer)
      ReadOnly Property HighIndexForeCastDistanceToPivot As Integer
      ReadOnly Property LowIndexForeCastDistanceToPivot As Integer
    End Interface


    Public Interface IFilterType
      Enum enuFilterType
        LowPassExp
        HighPassExp
        LowPassRMS
        LowPassPLL
        HighPassPLL
        AttackDecayExp
        LowPassExpBrownPrediction
        LowPassExpHull
        PriceDecimation
        VolatilityStandard
        VolatilityYangZhang
        Stochastic
        Bollinger
        RSI
        OBV
      End Enum

      Function AsIFilterType() As IFilterType
      ReadOnly Property FilterType As IFilterType.enuFilterType
    End Interface


    'provide direct control on the rate parameter without affecting any of the previous filtered data
    Public Interface IFilterControlRate
      Function AsIFilterControlRate() As IFilterControlRate

      Sub UpdateRate(ByVal FilterRate As Double)
      Sub UpdateRate(ByVal FilterRate As Integer)
      Property Enabled As Boolean
    End Interface

    ''' <summary>
    ''' This interface can be used to create a new filter using the same parameters use by the current one
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IFilterCreateNew
      Function AsIFilterCreateNew() As IFilterCreateNew
      Function CreateNew() As IFilter
    End Interface

    ''' <summary>
    ''' The Interface allow changing the rate of filtering on the fly and if needed clearing the memory for a fresh start
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IFilterControl
      Function AsIFilterControl() As IFilterControl
      Sub Clear()
      Sub Refresh(ByVal FilterRate As Double)
      Sub Refresh(ByVal FilterRate As Integer)
      ReadOnly Property FilterRate As Double
      ReadOnly Property IsInputEnabled As Boolean
      Function InputValue() As Double()
    End Interface

    Public Interface IFilterState
      Function ASIFilterState() As IFilterState
      Sub Save()
      Sub ReturnPrevious()
    End Interface

    Public Interface IFilterCopy
      Function AsIFilterCopy() As IFilterCopy
      Function CopyFrom() As IFilter
    End Interface

    Public Interface IFilterRunAsync
      Function FilterAsync(ByVal ReportPrices As YahooAccessData.RecordPrices) As Task(Of Boolean)
      Function FilterAsync(ByVal ReportPrices As YahooAccessData.RecordPrices, ByVal IsUseParallelBlock As Boolean) As Task(Of Boolean)
    End Interface

    Public Interface IFilter
      Function Filter(ByVal Value As Double) As Double
      Function Filter(ByRef Value() As Double) As Double()
      Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double()
      Function Filter(ByVal Value As Single) As Double
      Function Filter(ByVal Value As IPriceVol) As Double
      Function FilterErrorLast() As Double
      Function FilterBackTo(ByRef Value As Double) As Double
      Function FilterLastToPriceVol() As IPriceVol
      Function LastToPriceVol() As IPriceVol
      Function FilterPredictionNext(ByVal Value As Double) As Double
      Function FilterPredictionNext(ByVal Value As Single) As Double
      Function FilterLast() As Double
      Function Last() As Double
      ReadOnly Property Rate As Integer
      ReadOnly Property Count As Integer
      ReadOnly Property Max As Double
      ReadOnly Property Min As Double
      ReadOnly Property ToList() As IList(Of Double)
      ReadOnly Property ToListOfError() As IList(Of Double)
      ReadOnly Property ToListScaled() As ListScaled
      Function ToArray() As Double()
      Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
      Property Tag As String
			Function ToString() As String
		End Interface

    Public Interface IFilter(Of T)
      Function Filter(ByVal Value As Double) As T
      Function Filter(ByRef Value() As Double) As T()
      Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As T()
      Function Filter(ByVal Value As Single) As T
      Function Filter(ByVal Value As IPriceVol) As T
      Function FilterErrorLast() As T
      Function FilterBackTo(ByRef Value As T) As Double
      Function FilterLastToPriceVol() As IPriceVol
      Function LastToPriceVol() As IPriceVol
      Function FilterPredictionNext(ByVal Value As Double) As T
      Function FilterPredictionNext(ByVal Value As Single) As T
      Function FilterLast() As T
      Function Last() As Double
    ReadOnly Property Rate As Integer
      ReadOnly Property Count As Integer
      ReadOnly Property Max As Double
      ReadOnly Property Min As Double
      ReadOnly Property ToList() As IList(Of T)
      ReadOnly Property ToListOfError() As IList(Of T)
      ReadOnly Property ToListScaled() As ListScaled
      Function ToArray() As T()
      Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As T()
      Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As T()
      Property Tag As String
    Function ToString() As String
    End Interface
#Region "IFilterEstimate"
    ''' <summary>
    ''' This interface purpose is to estimate the next sample(s) for a given input(s) including gain and gain variation but without changing 
    ''' the state of the current filter. It is a new version of IPrediction that was created to be more consistent with the different filter
    ''' and the support of gain and gain variation measurement
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IFilterEstimate
      Function AsIFilterEstimate() As IFilterEstimate
      Function Filter(ByVal Value As Double) As IFilterEstimateResult
      Function Filter(ByVal Value() As Double) As IList(Of IFilterEstimateResult)
    End Interface

    Public Interface IFilterEstimateResult
      ReadOnly Property Value As Double
      ReadOnly Property Gain As Double
      ReadOnly Property GainDerivative As Double
      ReadOnly Property IsGainEnabled As Boolean
    End Interface

    Public Class FilterEstimateResult
      Implements IFilterEstimateResult

      Private MyValue As Double
      Private MyGain As Double
      Private MyGainDerivative As Double
      Private IsGainEnabledLocal As Boolean

      Public Sub New(ByVal Value As Double, ByVal Gain As Double, ByVal GainDerivative As Double)
        MyValue = Value
        MyGain = Gain
        MyGainDerivative = GainDerivative
        IsGainEnabledLocal = True
      End Sub

      Public Sub New(ByVal Value As Double)
        MyValue = Value
        MyGain = 0.0
        MyGainDerivative = 0.0
        IsGainEnabledLocal = False
      End Sub

      Public ReadOnly Property Gain As Double Implements IFilterEstimateResult.Gain
        Get
          Return MyGain
        End Get
      End Property

      Public ReadOnly Property GainDerivative As Double Implements IFilterEstimateResult.GainDerivative
        Get
          Return MyGainDerivative
        End Get
      End Property

      Public ReadOnly Property IsGainEnabled As Boolean Implements IFilterEstimateResult.IsGainEnabled
        Get
          Return IsGainEnabledLocal
        End Get
      End Property

      Public ReadOnly Property Value As Double Implements IFilterEstimateResult.Value
        Get
          Return MyValue
        End Get
      End Property
    End Class
#End Region

    Public Interface IFilterPrediction
      Function AsIFilterPrediction() As IFilterPrediction

      ReadOnly Property IsEnabled As Boolean

      ''' <summary>
      ''' Calculate the future output signal from the value of the input signal at the Index given the specified gain per year.
      ''' This function generally apply only for price value type signal
      ''' </summary>
      ''' <param name="Index"></param>
      ''' <param name="NumberOfPrediction"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Function FilterPrediction(ByVal Index As Integer, ByVal NumberOfPrediction As Integer) As Double
      ''' <summary>
      ''' Calculate the future output signal from the value of the input signal at the Index given the specified gain per year.
      ''' This function generally apply only for price value type signal
      ''' </summary>
      ''' <param name="Index"></param>
      ''' <param name="NumberOfPrediction"></param>
      ''' <param name="GainPerYear"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Function FilterPrediction(ByVal Index As Integer, ByVal NumberOfPrediction As Integer, ByVal GainPerYear As Double) As Double
      ''' <summary>
      ''' Calculate the future output signal from the last input signal and gain per year.
      ''' This function generally apply only for price value type signal
      ''' </summary>
      ''' <param name="NumberOfPrediction"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Function FilterPrediction(ByVal NumberOfPrediction As Integer) As Double
      ''' <summary>
      ''' Calculate the future output signal from the last input signal and specified gain per year.
      ''' This function generally apply only for price value type  signal
      ''' </summary>
      ''' <param name="NumberOfPrediction"></param>
      ''' <param name="GainPerYear"></param>
      ''' <returns></returns>
      ''' <remarks></remarks>
      Function FilterPrediction(ByVal NumberOfPrediction As Integer, ByVal GainPerYear As Double) As Double
      ReadOnly Property ToListOfGainPerYear() As IList(Of Double)
      ReadOnly Property ToListOfGainPerYearDerivative() As IList(Of Double)
    End Interface

		Public Interface IPeakValueRange
      ReadOnly Property High As Double
      ReadOnly Property Low As Double
      ReadOnly Property Range As Double
    End Interface
  End Namespace
End Namespace