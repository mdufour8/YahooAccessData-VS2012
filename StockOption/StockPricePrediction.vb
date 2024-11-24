Imports YahooAccessData.MathPlus.Measure
Imports MathNet.Numerics


Namespace OptionValuation
  Public Class StockPriceVolatilityPredictionBand
    Implements IStockPriceVolatilityPredictionBand

    Private VOLATILITY_TOTAL_MINIMUM As Double = 0.01

    Private MyVolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType
    Private MyVolatilityDelta As Double
    Private MyStockPrice As IPriceVol
    Private MyStockPriceFutur As IPriceVol
    Private MyStockPriceLowValue As Double
    Private MyStockPriceHighValue As Double
    Private MyStockPriceLowValueStandard As Double
    Private MyStockPriceHighValueStandard As Double
    Private MyStockPriceLowValueReal As Double
    Private MyStockPriceHighValueReal As Double
    Private MyProbabilityOfPriceExcessVolatilityRatio As Double
    Private IsStockPriceValueRealEnabledLocal As Boolean
    Private MyProbabilityLow As Double
    Private MyProbabilityHigh As Double
    Private MyNumberTradingDays As Double
    Private MyStockPriceStartValue As Double
    Private MyGain As Double
    Private MyGainDerivative As Double
    Private MyVolatility As Double
    Private MyVolatilityTotal As Double
    Private MyProbabilityOfInterval As Double
    Private IsBandExceededLocal As Boolean
    Private IsBandExceededHighLocal As Boolean
    Private IsBandExceededLowLocal As Boolean
    Private MyRatioOfDeltaProbabilityToVolatility As Double

    Public Sub New(
                  ByVal NumberTradingDays As Double,
                  ByRef StockPrice As IPriceVol,
                  ByVal StockPriceStartValue As Double,
                  ByVal Gain As Double,
                  ByVal GainDerivative As Double,
                  ByVal Volatility As Double,
                  ByVal ProbabilityOfInterval As Double)

      Me.New(
        NumberTradingDays,
        StockPrice,
        StockPriceStartValue,
        Gain,
        GainDerivative,
        Volatility,
        ProbabilityOfInterval,
        VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose)
    End Sub


    Public Sub New(
                  ByVal NumberTradingDays As Double,
                  ByRef StockPrice As IPriceVol,
                  ByVal StockPriceStartValue As Double,
                  ByVal Gain As Double,
                  ByVal GainDerivative As Double,
                  ByVal Volatility As Double,
                  ByVal ProbabilityOfInterval As Double,
                  ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)


      MyNumberTradingDays = NumberTradingDays
      MyStockPrice = StockPrice
      MyStockPriceStartValue = StockPriceStartValue
      MyGain = Gain
      MyGainDerivative = GainDerivative
      MyVolatility = Volatility
      MyProbabilityOfInterval = ProbabilityOfInterval
      MyProbabilityHigh = 0.5 + ProbabilityOfInterval / 2
      MyProbabilityLow = 0.5 - ProbabilityOfInterval / 2
      MyVolatilityDelta = 0.0
      Me.IsVolatilityMaximumEnabled = False
      IsStockPriceValueRealEnabledLocal = False
      MyStockPriceLowValueReal = 0.0
      MyStockPriceHighValueReal = 0.0
      'measure just a the high probabilty 
      'the low probability is similar but negatif 
      'since the interest is in the variation between the high and low probability
      'the current variation need to be ly by 2
      MyRatioOfDeltaProbabilityToVolatility = 2 * StockOption.StockPricePredictionPartialDerivateOfProbabilityToVolatility(MyNumberTradingDays, MyVolatility, MyProbabilityHigh)
      MyVolatilityPredictionBandType = VolatilityPredictionBandType
    End Sub

    Public Sub New(
                  ByVal NumberTradingDays As Double,
                  ByVal StockPriceStartValue As Double,
                  ByVal Gain As Double,
                  ByVal GainDerivative As Double,
                  ByVal Volatility As Double,
                  ByVal ProbabilityOfInterval As Double)

      Me.New(
        NumberTradingDays,
        New PriceVol(CSng(StockPriceStartValue)),
        StockPriceStartValue,
        Gain,
        GainDerivative,
        Volatility,
        ProbabilityOfInterval,
        VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose)
    End Sub

    Public Sub New(
                  ByVal NumberTradingDays As Double,
                  ByVal StockPriceStartValue As Double,
                  ByVal Gain As Double,
                  ByVal GainDerivative As Double,
                  ByVal Volatility As Double,
                  ByVal ProbabilityOfInterval As Double,
                  ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)

      Me.New(
        NumberTradingDays,
        New PriceVol(CSng(StockPriceStartValue)),
        StockPriceStartValue,
        Gain,
        GainDerivative,
        Volatility,
        ProbabilityOfInterval,
        VolatilityPredictionBandType)
    End Sub


    ''' <summary>
    ''' Create a copy of the original input object
    ''' </summary>
    ''' <param name="StockPriceVolatilityPredictionBand">
    ''' The object that will be copied from 
    ''' </param>
    ''' <remarks></remarks>
    Public Sub New(ByVal StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand, ByVal VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType)
      Me.New(
        StockPriceVolatilityPredictionBand.NumberTradingDays,
        StockPriceVolatilityPredictionBand.StockPrice,
        StockPriceVolatilityPredictionBand.StockPriceStartValue,
        StockPriceVolatilityPredictionBand.Gain,
        StockPriceVolatilityPredictionBand.GainDerivative,
        StockPriceVolatilityPredictionBand.Volatility,
        StockPriceVolatilityPredictionBand.ProbabilityOfInterval,
        VolatilityPredictionBandType)

      'copy the comparaison factor to ensure the object perform the same way that the provided one
      Me.Refresh(StockPriceVolatilityPredictionBand.StockPriceFutur)
    End Sub

    Public Sub New(ByVal StockPriceVolatilityPredictionBand As IStockPriceVolatilityPredictionBand)
      Me.New(
        StockPriceVolatilityPredictionBand.NumberTradingDays,
        StockPriceVolatilityPredictionBand.StockPrice,
        StockPriceVolatilityPredictionBand.StockPriceStartValue,
        StockPriceVolatilityPredictionBand.Gain,
        StockPriceVolatilityPredictionBand.GainDerivative,
        StockPriceVolatilityPredictionBand.Volatility,
        StockPriceVolatilityPredictionBand.ProbabilityOfInterval,
        VolatilityPredictionBandType:=IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose)
    End Sub

    Private Function IStockPriceVolatilityPredictionBand_StockPriceFutur() As IPriceVol Implements IStockPriceVolatilityPredictionBand.StockPriceFutur
      Return MyStockPriceFutur
    End Function

    Public Function Refresh(VolatilityDelta As Double, ByRef StockPriceFuture As IPriceVol) As Boolean Implements IStockPriceVolatilityPredictionBand.Refresh
      Me.Refresh(StockPriceFuture)
      Return Me.Refresh(VolatilityDelta)
    End Function

    Public Sub Refresh(ByRef StockPriceFuture As IPriceVol) Implements IStockPriceVolatilityPredictionBand.Refresh
      MyStockPriceFutur = StockPriceFuture
      If StockPriceFuture Is Nothing Then
        MyStockPriceLowValueReal = 0.0
        MyStockPriceHighValueReal = 0.0
        IsStockPriceValueRealEnabledLocal = False
      Else
				Select Case MyVolatilityPredictionBandType
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToClose
						MyStockPriceLowValueReal = StockPriceFuture.Low
						MyStockPriceHighValueReal = StockPriceFuture.High
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromCloseToOpen
						MyStockPriceLowValueReal = StockPriceFuture.Open
						MyStockPriceHighValueReal = StockPriceFuture.Open
					Case IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType.FromOpenToClose
				End Select
				IsStockPriceValueRealEnabledLocal = True
      End If
    End Sub

    ''' <summary>
    ''' Calculate the new price threshold using the additional Volatility given by VolatilityDelta
    ''' </summary>
    ''' <param name="VolatilityDelta">
    ''' this additional volatility is added to the base volatility before updating the corresponding parameters
    ''' </param>
    ''' <returns>Return a boolean indicationg if the current price range exceeded the new threshold</returns>
    ''' <remarks></remarks>
    Public Function Refresh(VolatilityDelta As Double) As Boolean Implements IStockPriceVolatilityPredictionBand.Refresh
      Dim ThisProbabilityOfPriceHighExcessVolatilityRatio As Double
      Dim ThisProbabilityOfPriceLowExcessVolatilityRatio As Double

      'Dim ThisVolatilityChangePerCent = VolatilityDelta / MyVolatility
      'ThisVolatilityChangePerCent = MathPlus.WaveForm.SignalLimit(ThisVolatilityChangePerCent, MinScale:=-1.0, MaxScale:=1, Offset:=0)
      'VolatilityDelta = ThisVolatilityChangePerCent * MyVolatility


      MyVolatilityDelta = VolatilityDelta
      MyVolatilityTotal = MyVolatility + MyVolatilityDelta
      'If Me.IsVolatilityMaximumEnabled Then
      '  If MyVolatilityTotal > Me.VolatilityMaximum Then
      '    MyVolatilityTotal = Me.VolatilityMaximum
      '    MyVolatilityDelta = MyVolatilityTotal - MyVolatility
      '  End If
      'End If
      'If MyVolatilityTotal > 100 Then
      '  Debugger.Break()
      'End If
      'pass the volatility to a limiter to eliminate the possibility of negative input

      Dim ThisStockPriceSample As Double

      If MyVolatilityTotal < 0 Then
        MyVolatilityTotal = MyVolatilityTotal
      End If
      If MyVolatilityTotal < VOLATILITY_TOTAL_MINIMUM Then
        'add some noise to shake the data with 1% additional noise
        ThisStockPriceSample = StockOption.StockPricePredictionSample(
        MyNumberTradingDays,
        MyStockPriceStartValue,
        MyGain,
        MyGainDerivative,
        VOLATILITY_TOTAL_MINIMUM)
      Else
        ThisStockPriceSample = MyStockPriceStartValue
      End If
      ThisStockPriceSample = MyStockPriceStartValue



      MyStockPriceHighValue = StockOption.StockPricePrediction(
        MyNumberTradingDays,
        ThisStockPriceSample,
        MyGain,
        MyGainDerivative,
        MyVolatilityTotal,
        MyProbabilityHigh)

      MyStockPriceLowValue = StockOption.StockPricePrediction(
        MyNumberTradingDays,
        ThisStockPriceSample,
        MyGain,
        MyGainDerivative,
        MyVolatilityTotal,
        MyProbabilityLow)

      'calculate the standard high and low
      If MyVolatility = MyVolatilityTotal Then
        MyStockPriceHighValueStandard = MyStockPriceHighValue
        MyStockPriceLowValueStandard = MyStockPriceLowValue
      Else
        MyStockPriceHighValueStandard = StockOption.StockPricePrediction(
          MyNumberTradingDays,
          ThisStockPriceSample,
          MyGain,
          MyGainDerivative,
          MyVolatility,
          MyProbabilityHigh)

        MyStockPriceLowValueStandard = StockOption.StockPricePrediction(
          MyNumberTradingDays,
          ThisStockPriceSample,
          MyGain,
          MyGainDerivative,
          MyVolatility,
          MyProbabilityLow)
      End If
      IsBandExceededLocal = False
      IsBandExceededHighLocal = False
      IsBandExceededLowLocal = False
      MyProbabilityOfPriceExcessVolatilityRatio = 1.0
      If IsStockPriceValueRealEnabledLocal Then
        If MyStockPriceHighValueReal >= MyStockPriceHighValue Then
          IsBandExceededLocal = True
          IsBandExceededHighLocal = True

          ThisProbabilityOfPriceHighExcessVolatilityRatio = StockOption.StockPricePredictionInverseToVolatilityRatio(
            MyNumberTradingDays,
            ThisStockPriceSample,
            MyGain,
            MyGainDerivative,
            MyVolatility,
            MyStockPriceHighValueReal)
        End If
        If MyStockPriceLowValueReal <= MyStockPriceLowValue Then
          IsBandExceededLocal = True
          IsBandExceededLowLocal = True

          ThisProbabilityOfPriceLowExcessVolatilityRatio = StockOption.StockPricePredictionInverseToVolatilityRatio(
            MyNumberTradingDays,
            ThisStockPriceSample,
            MyGain,
            MyGainDerivative,
            MyVolatility,
            MyStockPriceLowValueReal)
        End If
      Else
        'compare with the current value since the future is not yet availaible
        If MyStockPrice.High >= MyStockPriceHighValue Then
          IsBandExceededLocal = True
          IsBandExceededHighLocal = True

          ThisProbabilityOfPriceHighExcessVolatilityRatio = StockOption.StockPricePredictionInverseToVolatilityRatio(
            MyNumberTradingDays,
            ThisStockPriceSample,
            MyGain,
            MyGainDerivative,
            MyVolatility,
            MyStockPrice.High)
        End If
        If MyStockPrice.Low <= MyStockPriceLowValue Then
          IsBandExceededLocal = True
          IsBandExceededLowLocal = True

          ThisProbabilityOfPriceLowExcessVolatilityRatio = StockOption.StockPricePredictionInverseToVolatilityRatio(
            MyNumberTradingDays,
            ThisStockPriceSample,
            MyGain,
            MyGainDerivative,
            MyVolatility,
            MyStockPrice.Low)
        End If
      End If
      If ThisProbabilityOfPriceHighExcessVolatilityRatio > MyProbabilityOfPriceExcessVolatilityRatio Then
        MyProbabilityOfPriceExcessVolatilityRatio = ThisProbabilityOfPriceHighExcessVolatilityRatio
      End If
      If ThisProbabilityOfPriceLowExcessVolatilityRatio > MyProbabilityOfPriceExcessVolatilityRatio Then
        MyProbabilityOfPriceExcessVolatilityRatio = ThisProbabilityOfPriceLowExcessVolatilityRatio
      End If
      Return IsBandExceededLocal
    End Function

    Public ReadOnly Property VolatilityDelta As Double Implements IStockPriceVolatilityPredictionBand.VolatilityDelta
      Get
        Return MyVolatilityDelta
      End Get
    End Property

    Public Function StockPrice() As IPriceVol Implements IStockPriceVolatilityPredictionBand.StockPrice
      Return MyStockPrice
    End Function

    Public ReadOnly Property ProbabilityHigh As Double Implements IStockPriceVolatilityPredictionBand.ProbabilityHigh
      Get
        Return MyProbabilityHigh
      End Get
    End Property

    Public ReadOnly Property ProbabilityLow As Double Implements IStockPriceVolatilityPredictionBand.ProbabilityLow
      Get
        Return MyProbabilityLow
      End Get
    End Property

    Public ReadOnly Property StockPriceHighValue As Double Implements IStockPriceVolatilityPredictionBand.StockPriceHighValue
      Get
        Return MyStockPriceHighValue
      End Get
    End Property

    Public ReadOnly Property StockPriceLowValue As Double Implements IStockPriceVolatilityPredictionBand.StockPriceLowValue
      Get
        Return MyStockPriceLowValue
      End Get
    End Property

    Public ReadOnly Property Gain As Double Implements IStockPriceVolatilityPredictionBand.Gain
      Get
        Return MyGain
      End Get
    End Property

    Public ReadOnly Property GainDerivative As Double Implements IStockPriceVolatilityPredictionBand.GainDerivative
      Get
        Return MyGainDerivative
      End Get
    End Property

    Public ReadOnly Property IsBandExceeded As Boolean Implements IStockPriceVolatilityPredictionBand.IsBandExceeded
      Get
        Return IsBandExceededLocal
      End Get
    End Property

    Public ReadOnly Property NumberTradingDays As Double Implements IStockPriceVolatilityPredictionBand.NumberTradingDays
      Get
        Return MyNumberTradingDays
      End Get
    End Property

    Public ReadOnly Property ProbabilityOfInterval As Double Implements IStockPriceVolatilityPredictionBand.ProbabilityOfInterval
      Get
        Return MyProbabilityOfInterval
      End Get
    End Property

    Public ReadOnly Property StockPriceStartValue As Double Implements IStockPriceVolatilityPredictionBand.StockPriceStartValue
      Get
        Return MyStockPriceStartValue
      End Get
    End Property

    Public ReadOnly Property Volatility As Double Implements IStockPriceVolatilityPredictionBand.Volatility
      Get
        Return MyVolatility
      End Get
    End Property

    Public ReadOnly Property IsBandExceededHigh As Boolean Implements IStockPriceVolatilityPredictionBand.IsBandExceededHigh
      Get
        Return IsBandExceededHighLocal
      End Get
    End Property

    Public ReadOnly Property IsBandExceededLow As Boolean Implements IStockPriceVolatilityPredictionBand.IsBandExceededLow
      Get
        Return IsBandExceededLowLocal
      End Get
    End Property

    Public Overrides Function ToString() As String
      Return String.Format("Object:{0}:{1},Volatility:{2:n3}, IsBandExceeded:{3},{4}", TypeName(Me), MyStockPrice.DateDay, Me.Volatility, Me.IsBandExceeded, MyStockPrice.ToString)
    End Function

    ''' <summary>
    ''' Return the ΔProbability/ΔVolatility of being exceeded. 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>It can be used when the variational probability slope is needed i.e. 
    ''' in a feedback system to better estimate and correcte for the real observed probability of the system or stock under analysis.
    ''' </remarks>
    Public ReadOnly Property RatioOfΔProbabilityToΔVolatility As Double Implements IStockPriceVolatilityPredictionBand.RatioOfΔProbabilityToΔVolatility
      Get
        Return MyRatioOfDeltaProbabilityToVolatility
      End Get
    End Property

    ''' <summary>
    ''' The total volatility
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property VolatilityTotal As Double Implements IStockPriceVolatilityPredictionBand.VolatilityTotal
      Get
        Return MyVolatilityTotal
      End Get
    End Property

    Public Property IsVolatilityMaximumEnabled As Boolean Implements IStockPriceVolatilityPredictionBand.IsVolatilityMaximumEnabled
    Public Property VolatilityMaximum As Double Implements IStockPriceVolatilityPredictionBand.VolatilityMaximum

    Public ReadOnly Property IsStockPriceValueRealEnabled As Boolean Implements IStockPriceVolatilityPredictionBand.IsStockPriceValueRealEnabled
      Get
        Return IsStockPriceValueRealEnabledLocal
      End Get
    End Property

    Public ReadOnly Property StockPriceHighValueStandard As Double Implements IStockPriceVolatilityPredictionBand.StockPriceHighValueStandard
      Get
        Return MyStockPriceHighValueStandard
      End Get
    End Property

    Public ReadOnly Property StockPriceLowValueStandard As Double Implements IStockPriceVolatilityPredictionBand.StockPriceLowValueStandard
      Get
        Return MyStockPriceLowValueStandard
      End Get
    End Property

    Public ReadOnly Property VolatilityExcessRatio As Double Implements IStockPriceVolatilityPredictionBand.VolatilityExcessRatio
      Get
        Return MyProbabilityOfPriceExcessVolatilityRatio
      End Get
    End Property

    ''' <summary>
    ''' Calculate the expected high future value for a time in day
    ''' </summary>
    ''' <param name="Index">the time in day</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StockPriceHighPrediction(Index As Double) As Double Implements IStockPriceVolatilityPredictionBand.StockPriceHighPrediction
      Return StockOption.StockPricePrediction(
        Index,
        MyStockPriceStartValue,
        MyGain,
        MyGainDerivative,
        MyVolatilityTotal,
        MyProbabilityHigh)
    End Function

    ''' <summary>
    ''' Calculate the expected low future value for a time in day
    ''' </summary>
    ''' <param name="Index">the time in day</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function StockPriceLowPrediction(Index As Double) As Double Implements IStockPriceVolatilityPredictionBand.StockPriceLowPrediction
      Return StockOption.StockPricePrediction(
        Index,
        MyStockPriceStartValue,
        MyGain,
        MyGainDerivative,
        MyVolatilityTotal,
        MyProbabilityLow)
    End Function

    ''' <summary>
    ''' Return the price level given the Index in day and the probability
    ''' </summary>
    ''' <param name="Index">time period in day</param>
    ''' <param name="Probability"></param>
    ''' <returns></returns>
    Public Function StockPricePrediction(Index As Double, Probability As Double) As Double Implements IStockPriceVolatilityPredictionBand.StockPricePrediction
      Return StockOption.StockPricePrediction(
        Index,
        MyStockPriceStartValue,
        MyGain,
        MyGainDerivative,
        MyVolatilityTotal,
        Probability)
    End Function


    Public ReadOnly Property VolatilityPredictionBandType As IStockPriceVolatilityPredictionBand.EnuVolatilityPredictionBandType Implements IStockPriceVolatilityPredictionBand.VolatilityPredictionBandType
      Get
        Return MyVolatilityPredictionBandType
      End Get
    End Property
  End Class

  Public Interface IStockPriceVolatilityPredictionBand
    Enum EnuVolatilityPredictionBandType
      FromCloseToClose
      FromCloseToOpen
      FromOpenToClose
    End Enum

    ReadOnly Property VolatilityPredictionBandType As EnuVolatilityPredictionBandType
    ReadOnly Property StockPriceHighValue As Double
    ReadOnly Property StockPriceLowValue As Double
    Function StockPriceHighPrediction(ByVal Index As Double) As Double
    Function StockPriceLowPrediction(ByVal Index As Double) As Double
    Function StockPricePrediction(ByVal Index As Double, ByVal Probability As Double) As Double
    ReadOnly Property StockPriceHighValueStandard As Double
    ReadOnly Property StockPriceLowValueStandard As Double
    ReadOnly Property IsStockPriceValueRealEnabled As Boolean
    ReadOnly Property ProbabilityHigh As Double
    ReadOnly Property ProbabilityLow As Double

    ReadOnly Property NumberTradingDays As Double
    ReadOnly Property StockPriceStartValue As Double

    ReadOnly Property Gain As Double
    ReadOnly Property GainDerivative As Double
    ReadOnly Property Volatility As Double
    Property VolatilityMaximum As Double
    Property IsVolatilityMaximumEnabled As Boolean
    ReadOnly Property ProbabilityOfInterval As Double
    ReadOnly Property VolatilityExcessRatio As Double
    ReadOnly Property VolatilityDelta As Double
    ReadOnly Property VolatilityTotal As Double
    ReadOnly Property RatioOfΔProbabilityToΔVolatility As Double
    ReadOnly Property IsBandExceeded As Boolean
    ReadOnly Property IsBandExceededHigh As Boolean
    ReadOnly Property IsBandExceededLow As Boolean
    Function StockPrice() As IPriceVol
    Function StockPriceFutur() As IPriceVol
    Function Refresh(ByVal VolatilityDelta As Double, ByRef StockPriceFuture As IPriceVol) As Boolean
    Sub Refresh(ByRef StockPriceFuture As IPriceVol)
    Function Refresh(ByVal VolatilityDelta As Double) As Boolean
  End Interface
End Namespace