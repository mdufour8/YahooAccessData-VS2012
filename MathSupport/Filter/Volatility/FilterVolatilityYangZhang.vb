#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.MathPlus.Probability
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region

Namespace MathPlus.Filter
#Region "FilterVolatilityYangZhang"
  ''' <summary>
  ''' This class implement the Yang-Zhang volatility measurements
  ''' see definition:
  ''' http://en.wikipedia.org/wiki/Rate_of_return#Logarithmic_or_continuously_compounded_return
  ''' http://en.wikipedia.org/wiki/Volatility_%28finance%29
  ''' https://www.youtube.com/watch?v=eiTCTibH010
  ''' http://en.wikipedia.org/wiki/Volatility_(finance)
  ''' https://en.wikipedia.org/wiki/Stochastic_volatility
  ''' https://en.wikipedia.org/wiki/Geometric_Brownian_motion
  ''' </summary>
  ''' <remarks>
  ''' In 2000 Yang-Zhang created the most powerful volatility
  ''' measure that handles both opening jumps and drift. It is the sum of the overnight
  ''' volatility (close to open volatility) and a weighted average of the Rogers-Satchell
  ''' volatility and the open to close volatility. The assumption of continuous prices does
  ''' mean the measure tends to slightly underestimate the volatility. The class use 
  ''' the normalized logarithm compounded return to measure the volatility and
  ''' is only valid for positive value of signal
  ''' </remarks>
  <Serializable()>
  Public Class FilterVolatilityYangZhang
    Implements IFilter
    Implements IRegisterKey(Of String)

    Public Enum enuVolatilityDailyPeriodType
      FullDay
      PreviousCloseToOpen
      OpenToClose
      OpenToHighClose
      OpenToLowClose
      OpenToHighToLowCloseRatio
			OpenToHighToLowCloseRatioFiltered
			Rogers_Satchell_Yoon_Vrs
			Parkison_Vp
		End Enum

		Private MyListOfPreviousCloseToOpenHighLowClose As List(Of Double)
		Private MyListOfPreviousCloseToOpen As List(Of Double)
    Private MyListOfOpenToClose As List(Of Double)
    Private MyListOfOpenHighAsClose As List(Of Double)
    Private MyListOfOpenLowAsClose As List(Of Double)
		Private MyListOfOpenToHighToLowAsCloseRatio As List(Of Double)
		Private MyListOfFilteredOpenToHighToLowAsCloseRatio As List(Of Double)

		Private MyFilterVolatilityYZ As FilterVolatilityYZ

		Public Sub New()
			Me.New(FilterVolatilityYZ.FILTER_RATE_DEFAULT)
		End Sub

		''' <summary>
		''' 
		''' </summary>
		''' <param name="FilterRate">The number of sample use for the volatility measurement</param>
		''' <param name="VolatilityScale">The scale correction factor for the volatility measurement if needed. The default is Volatility per year</param>
		''' <param name="StatisticType">The type of statistical method to use for the volatility measurement </param>
		''' <param name="IsUseLastSampleHighLowTrail">
		'''		Indicates whether to use the last sample's high and low as a trail for the next sample. 
		'''		This is useful to calculate the worst case volatility assuming that the last point high or low was also part of daily steam.
		'''		Interesting to note that this method often match the CBOE Vix volatility index that is based on the SP500 index.
		'''		This measurement can also be used to evaluate the degree of correlation between the previous close and the next day opening.
		'''		A low correlation could indicate outside market hour activity that could be considered abnormal.
		'''	</param>
		Public Sub New(
			FilterRate As Integer,
			Optional StatisticType As FilterVolatility.enuVolatilityStatisticType = FilterVolatility.enuVolatilityStatisticType.Standard,
			Optional VolatilityScale As Double = VOLATILITY_DAILY_TO_YEARLY_RATIO,
			Optional IsUseLastSampleHighLowTrail As Boolean = False)

			MyFilterVolatilityYZ = New FilterVolatilityYZ(
				FilterRate:=FilterRate,
				StatisticType:=StatisticType,
				VolatilityScale:=VolatilityScale,
				IsUseLastSampleHighLowTrail:=IsUseLastSampleHighLowTrail)

			MyListOfPreviousCloseToOpenHighLowClose = New List(Of Double)
			MyListOfPreviousCloseToOpen = New List(Of Double)
			MyListOfOpenToClose = New List(Of Double)
			MyListOfOpenHighAsClose = New List(Of Double)
			MyListOfOpenLowAsClose = New List(Of Double)
			MyListOfOpenToHighToLowAsCloseRatio = New List(Of Double)
			MyListOfFilteredOpenToHighToLowAsCloseRatio = New List(Of Double)
		End Sub

		Public Function Filter(Value As YahooAccessData.IPriceVol, IsVolatityHoldToLast As Boolean) As Double

			With MyFilterVolatilityYZ
				.Filter(Value, IsVolatityHoldToLast)
				MyListOfPreviousCloseToOpenHighLowClose.Add(.FilterLast(enuVolatilityDailyPeriodType.FullDay))
				MyListOfOpenToClose.Add(.FilterLast(enuVolatilityDailyPeriodType.OpenToClose))
				MyListOfPreviousCloseToOpen.Add(.FilterLast(enuVolatilityDailyPeriodType.PreviousCloseToOpen))
				MyListOfOpenHighAsClose.Add(.FilterLast(enuVolatilityDailyPeriodType.OpenToHighClose))
				MyListOfOpenLowAsClose.Add(.FilterLast(enuVolatilityDailyPeriodType.OpenToLowClose))
				MyListOfOpenToHighToLowAsCloseRatio.Add(.FilterLast(enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio))
				MyListOfFilteredOpenToHighToLowAsCloseRatio.Add(.FilterLast(enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio))
				Return .FilterLast
			End With
		End Function

		''' <summary>
		''' 
		''' </summary>
		''' <param name="Value"></param>
		''' <returns></returns>
		Public Function Filter(ByVal Value As YahooAccessData.IPriceVol) As Double Implements IFilter.Filter
			Return Me.Filter(Value, IsVolatityHoldToLast:=False)
		End Function

		Public Function FilterLast() As Double Implements IFilter.FilterLast
			Return MyFilterVolatilityYZ.FilterLast
		End Function

		Public Function FilterLast(ByVal Type As enuVolatilityDailyPeriodType) As Double
			Return MyFilterVolatilityYZ.FilterLast(Type)
		End Function

		Public Function FilterDirection() As FilterRSI.SlopeDirection
			Return MyFilterVolatilityYZ.FilterDirection
		End Function

		Public Function Last() As Double Implements IFilter.Last
			Return MyFilterVolatilityYZ.Last.Last
		End Function

		Public ReadOnly Property Rate As Integer Implements IFilter.Rate
			Get
				Return MyFilterVolatilityYZ.Rate
			End Get
		End Property

		Public ReadOnly Property Count As Integer Implements IFilter.Count
			Get
				Return MyListOfPreviousCloseToOpenHighLowClose.Count
			End Get
		End Property

		Public ReadOnly Property Max As Double Implements IFilter.Max
			Get
				Return MyListOfPreviousCloseToOpenHighLowClose.Max
			End Get
		End Property

		Public ReadOnly Property Min As Double Implements IFilter.Min
			Get
				Return MyListOfPreviousCloseToOpenHighLowClose.Min
			End Get
		End Property

		Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
			Get
				Return MyListOfPreviousCloseToOpenHighLowClose
			End Get
		End Property

		Public ReadOnly Property ToList(ByVal Type As enuVolatilityDailyPeriodType) As IList(Of Double)
			Get
				Select Case Type
					Case enuVolatilityDailyPeriodType.FullDay
						Return MyListOfPreviousCloseToOpenHighLowClose
					Case enuVolatilityDailyPeriodType.OpenToClose
						Return MyListOfOpenToClose
					Case enuVolatilityDailyPeriodType.PreviousCloseToOpen
						Return MyListOfPreviousCloseToOpen
					Case enuVolatilityDailyPeriodType.OpenToHighClose
						Return MyListOfOpenHighAsClose
					Case enuVolatilityDailyPeriodType.OpenToLowClose
						Return MyListOfOpenLowAsClose
					Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatio
						Return MyListOfOpenToHighToLowAsCloseRatio
					Case enuVolatilityDailyPeriodType.OpenToHighToLowCloseRatioFiltered
						Return MyListOfFilteredOpenToHighToLowAsCloseRatio
					Case Else
						Return MyListOfPreviousCloseToOpenHighLowClose
				End Select
			End Get
		End Property

		Private ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
			Get
				Throw New NotImplementedException
			End Get
		End Property

		Public Function ToArray() As Double() Implements IFilter.ToArray
			Return MyListOfPreviousCloseToOpenHighLowClose.ToArray
		End Function

		Private Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Throw New NotImplementedException
		End Function

		Private Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
			Throw New NotImplementedException
		End Function

		Public Property Tag As String Implements IFilter.Tag

		Public Overrides Function ToString() As String Implements IFilter.ToString
			Return Me.FilterLast.ToString
		End Function

		Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
			Dim ThisValue As Double
			For Each ThisValue In Value
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(ThisValue))
				Me.Filter(ThisPriceVol, IsVolatityHoldToLast:=False)
			Next
			Return Me.ToArray
		End Function

		Private Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
			Throw New NotImplementedException
		End Function

		Function Filter(Value As Double) As Double Implements IFilter.Filter
			Return Me.Filter(New PriceVol(CSng(Value)), IsVolatityHoldToLast:=False)
		End Function

		Public Function Filter(Value As Single) As Double Implements IFilter.Filter
			Return Me.Filter(New PriceVol(Value), IsVolatityHoldToLast:=False)
		End Function

		Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
			Throw New NotImplementedException
		End Function

		Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
			Throw New NotImplementedException
		End Function

		Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
			Throw New NotImplementedException
		End Function

		Private Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
			Throw New NotImplementedException
		End Function

		Private Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
			Throw New NotImplementedException
		End Function

		Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
			Return MyFilterVolatilityYZ.Last
		End Function

		Private ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of Double) Implements IFilter.ToListOfError
			Get
				Throw New NotImplementedException
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
  End Class
#End Region
End Namespace

