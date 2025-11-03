#Region "FilterOBV"
Imports Newtonsoft.Json.Linq
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter

<Serializable()>
Public Class FilterOBV
	Private Const FILTER_RATE_FOR_AVERAGE_VOLUME As Integer = 20
	Private MyRate As Double
	Private MyRatePreFilter As Integer
	Private MyPriceLast As Double
	Private MyVolumeLast As Long
	Private MyOBVLast As Double
	Private MyPriceFilteredLast As Double
	Private MyVolumeLogFilteredLast As Double
	Private MyFilterVolatilityForPositifNegatif As Filter.FilterVolatilityYangZhang
	Private MyFilterForVolumeAverageExp As IFilterRun
	Private MyFilterLowPassForOBVOut As IFilter

	Public Sub New(ByVal FilterRate As Double, Optional FilterRateForAverageVolume As Double = FILTER_RATE_FOR_AVERAGE_VOLUME)
		If FilterRate < 1 Then FilterRate = 1
		MyRate = FilterRate
		MyFilterForVolumeAverageExp = New FilterExp(FilterRateForAverageVolume)
		MyFilterLowPassForOBVOut = New FilterLowPassExp(FilterRate)
		'use the volatility filter to later determine if the price is going up or down an establish if yes or no these volume are buy or sell acelerating volume
		MyFilterVolatilityForPositifNegatif = New FilterVolatilityYangZhang(CInt(MyRate), FilterVolatility.enuVolatilityStatisticType.Exponential, IsUseLastSampleHighLowTrail:=False)
	End Sub

	''' <summary>
	''' Use this call to measure the accelerating A-OBV base on the price direction that establish the positivity or not of the incoming volume
	''' </summary>
	''' <param name="Volume"></param>
	''' <param name="Direction">
	''' The direction of the price movement associated with this volume can be of 4 type: as defined here:
	'''     NotSpecified : the direction is not specified, the function will try to determine it based on the price movement
	'''     In that case the function behave more or less like teh standard OBV definition
	'''     Positive : the price is moving up
	'''     Negative : the price is moving down
	'''     Sideways : the price is moving sideways or the same as the previous sample
	''' </param>
	''' <returns></returns>
	Public Function Filter(ByVal Volume As Long, ByVal Direction As FilterRSI.SlopeDirection) As Double
		Dim ThisVolumeLogFiltered As Double
		Dim ThisVolumeLongTermAverageFiltered As Double

		If MyFilterLowPassForOBVOut.Count = 0 Then
			'OBV initialisation 
			MyOBVLast = 0
			MyVolumeLogFilteredLast = 0
			MyVolumeLast = 0
		End If
		If Volume < 0 Then Volume = 0
		If Volume = 0 Then
			'use the last values
			ThisVolumeLongTermAverageFiltered = MyFilterForVolumeAverageExp.FilterLast
			ThisVolumeLogFiltered = MyVolumeLogFilteredLast
		Else
			ThisVolumeLongTermAverageFiltered = MyFilterForVolumeAverageExp.FilterRun(Math.Log(Volume + 1))
			'applied a logaritmic fucntion to the volume
			'this is equivalent to a ratio of volume over average volume	
			ThisVolumeLogFiltered = Math.Log(Volume + 1) - ThisVolumeLongTermAverageFiltered
		End If
		If ThisVolumeLogFiltered > 0 Then
			'volume increasing relative to an average is interpreted as significant
			'this is only when we start to change the OBV.
			Select Case Direction
				Case FilterRSI.SlopeDirection.Positive
					MyOBVLast = MyOBVLast + ThisVolumeLogFiltered
				Case FilterRSI.SlopeDirection.Negative
					MyOBVLast = MyOBVLast - ThisVolumeLogFiltered
			End Select
		ElseIf ThisVolumeLogFiltered < 0 Then
			'decreasing volume in any direction is difficult ot interpret
			'so do nothign when the volume are decreasing for now
			''volume decrease
			'Select Case Direction
			'	Case FilterRSI.SlopeDirection.Positive
			'		MyOBVLast = MyOBVLast + 0.1 * ThisVolumeLogFiltered
			'	Case FilterRSI.SlopeDirection.Negative
			'		MyOBVLast = MyOBVLast - 0.1 * ThisVolumeLogFiltered
			'End Select
		End If
		MyVolumeLogFilteredLast = ThisVolumeLogFiltered
		MyVolumeLast = Volume
		Return MyFilterLowPassForOBVOut.Filter(MyOBVLast)
	End Function


	''' <summary>
	''' This function is old and should be updated 
	''' </summary>
	''' <param name="Price"></param>
	''' <param name="Volume"></param>
	''' <returns></returns>
	Public Function Filter(ByVal Price As Double, ByVal Volume As Long) As Double
		'It is important to update the price last before filtering
		'It is partially use to determine the direction of the price movement
		Dim ThisPriceVol As IPriceVol = New PriceVol(PriceValue:=CSng(Price), Volume:=Volume) With {
			.LastPrevious = CSng(MyPriceLast)}
		Return Me.Filter(ThisPriceVol)
	End Function

	Public Function Filter(ByRef Value() As YahooAccessData.IPriceVol) As Double()
		Dim ThisValue As YahooAccessData.IPriceVol
		For Each ThisValue In Value
			Me.Filter(ThisValue)
		Next
		Return Me.ToArray
	End Function

	Private Function FilterPredictionNext(ByVal Price As Double, ByVal Volume As Integer) As Double
		Dim ThisPriceVariation As Double
		Dim ThisPriceValueFiltered As Double
		Dim ThisVolumeValueFiltered As Double
		Dim ThisVolumePeakFiltered As Double

		Dim ThisPriceLast As Double = Price
		Dim ThisVolumeLast As Integer = Volume
		Dim ThisOBVLast = MyOBVLast
		Dim ThisPriceFilteredLast = MyPriceFilteredLast

		'do not filter volume if zero
		Throw New NotImplementedException

		'If Volume = 0 Then
		'	ThisVolumePeakFiltered = 0
		'Else
		'	ThisVolumePeakFiltered = MyFilterForVolumeAverageExp.FilterPredictionNext(CDbl(Volume))
		'	'limit the volume
		'	ThisVolumePeakFiltered = MathPlus.WaveForm.SignalLimit(CDbl(Volume), MyLimitFactorForVolume * ThisVolumePeakFiltered)
		'End If
		ThisPriceValueFiltered = Price
		ThisVolumeValueFiltered = ThisVolumePeakFiltered
		If MyFilterLowPassForOBVOut.Count = 0 Then
			'OBV initialisation 
			ThisOBVLast = 0
			ThisPriceFilteredLast = ThisPriceValueFiltered
		End If
		If ThisVolumeValueFiltered > 0 Then
			ThisPriceVariation = ThisPriceValueFiltered - ThisPriceFilteredLast
			If ThisPriceVariation < 0 Then
				ThisOBVLast = ThisOBVLast - ThisVolumeValueFiltered
			ElseIf ThisPriceVariation > 0 Then
				ThisOBVLast = ThisOBVLast + ThisVolumeValueFiltered
			End If
		End If
		ThisPriceFilteredLast = ThisPriceValueFiltered
		Return MyFilterLowPassForOBVOut.FilterPredictionNext(ThisOBVLast)
	End Function

	Public Function Filter(ByVal Price As Single, ByVal Volume As Integer) As Double
		Return Me.Filter(CDbl(Price), Volume)
	End Function

	Private Function FilterPredictionNext(ByVal Price As Single, ByVal Volume As Integer) As Double
		Return Me.FilterPredictionNext(CDbl(Price), Volume)
	End Function

	Public Function Filter(ByRef PriceVol As IPriceVol) As Double
		MyFilterVolatilityForPositifNegatif.Filter(PriceVol, IsVolatityHoldToLast:=False)
		MyPriceLast = PriceVol.Last
		Return Me.Filter(DirectCast(PriceVol, PriceVol).Volume, MyFilterVolatilityForPositifNegatif.FilterDirection)
	End Function

	Private Function FilterPredictionNext(ByRef PriceVol As IPriceVol) As Double
		With PriceVol
			Return Me.FilterPredictionNext(CDbl(.LastWeighted), .Vol)
		End With
	End Function


	Public Function FilterLast() As Double
		Return MyFilterLowPassForOBVOut.FilterLast
	End Function

	Public Function PriceLast() As Double
		Return MyPriceLast
	End Function

	Public Function VolumeLast() As Long
		Return MyVolumeLast
	End Function

	Public ReadOnly Property Rate As Double
		Get
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer
		Get
			Return MyFilterLowPassForOBVOut.Count
		End Get
	End Property

	Public ReadOnly Property Max As Double
		Get
			Return MyFilterLowPassForOBVOut.Max
		End Get
	End Property

	Public ReadOnly Property Min As Double
		Get
			Return MyFilterLowPassForOBVOut.Min
		End Get
	End Property

	Public ReadOnly Property ToList() As IList(Of Double)
		Get
			Return MyFilterLowPassForOBVOut.ToList
		End Get
	End Property

	Public Function ToArray() As Double()
		Return MyFilterLowPassForOBVOut.ToArray
	End Function

	Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
		Return Me.ToArray(Me.Min, Me.Max, ScaleToMinValue, ScaleToMaxValue)
	End Function

	Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
		Return MyFilterLowPassForOBVOut.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
	End Function

	Public Property Tag As String

	Public Overrides Function ToString() As String
		Return Me.FilterLast.ToString
	End Function
End Class
#End Region
