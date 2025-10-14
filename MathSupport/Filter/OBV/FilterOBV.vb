#Region "FilterOBV"
Imports YahooAccessData.MathPlus.Filter

<Serializable()>
Public Class FilterOBV
	Private Const FILTER_RATE_FOR_AVERAGE_VOLUME As Integer = 20
	Private MyRate As Integer
	Private MyRatePreFilter As Integer
	Private MyPriceLast As Double
	Private MyVolumeLast As Long
	Private MyOBVLast As Double
	Private MyPriceFilteredLast As Double
	Private MyVolumeLogFilteredLast As Double
	Private MyFilterForVolumeAverageExp As IFilterRun
	Private MyFilterLowPassForOBVOut As IFilter

	Public Sub New(ByVal FilterRate As Integer, Optional FilterRateForAverageVolume As Integer = FILTER_RATE_FOR_AVERAGE_VOLUME)
		If FilterRate < 1 Then FilterRate = 1
		MyRate = FilterRate
		MyFilterForVolumeAverageExp = New FilterExp(FilterRateForAverageVolume)
		MyFilterLowPassForOBVOut = New FilterLowPassExp(FilterRate)
	End Sub

	Public Function Filter(ByVal Volume As Long, ByVal Direction As FilterRSI.SlopeDirection) As Double
		Dim ThisVolumeLogFiltered As Double
		Dim ThisVolumeAverageFiltered As Double

		If MyFilterLowPassForOBVOut.Count = 0 Then
			'OBV initialisation 
			MyOBVLast = 0
			MyVolumeLogFilteredLast = 0
			MyVolumeLast = 0
		End If
		If Volume < 0 Then Volume = 0
		If Volume = 0 Then
			ThisVolumeAverageFiltered = MyFilterForVolumeAverageExp.FilterLast
			ThisVolumeLogFiltered = MyVolumeLogFilteredLast
		Else
			ThisVolumeAverageFiltered = MyFilterForVolumeAverageExp.FilterRun(CDbl(Volume))
			'limit the volume 
			'Dim ThisVolLimited = MathPlus.WaveForm.SignalLimit(CDbl(Volume), MyLimitFactorForVolume * ThisVolumeLogFiltered)
			'ThisVolumeLogFiltered = Math.Log((ThisVolLimited + 1) / ThisVolumeLogFiltered)
			'applied a logaritmic fucntion to the volume
			ThisVolumeLogFiltered = Math.Log(Volume / ThisVolumeAverageFiltered)
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

	Public Function Filter(ByVal Price As Double, ByVal Volume As Long, Optional ByVal Direction As FilterRSI.SlopeDirection = FilterRSI.SlopeDirection.NotSpecified) As Double
		Dim ThisPriceVariation As Double
		Dim ThisPriceValueFiltered As Double
		Dim ThisVolumeValueFiltered As Double
		Dim ThisVolumeFiltered As Double
		Dim ThisVolumeLogFiltered As Double

		MyPriceLast = Price
		MyVolumeLast = Volume

		If Volume = 0 Then
			ThisVolumeFiltered = 0
		Else
			ThisVolumeFiltered = MyFilterForVolumeAverageExp.FilterRun(CDbl(Volume))
			'limit the volume 
			'Dim ThisVolLimited = MathPlus.WaveForm.SignalLimit(CDbl(Volume), MyLimitFactorForVolume * ThisVolumeLogFiltered)
			'ThisVolumeLogFiltered = Math.Log((ThisVolLimited + 1) / ThisVolumeLogFiltered)
			If Volume = 0 Then
				'do not allow Volume to go to zero on a Log function
				Volume = 1
			End If
			ThisVolumeLogFiltered = Math.Log((Volume) / ThisVolumeFiltered)
		End If
		'do not filter
		ThisPriceValueFiltered = Price
		ThisVolumeValueFiltered = ThisVolumeLogFiltered
		If MyFilterLowPassForOBVOut.Count = 0 Then
			'OBV initialisation 
			MyOBVLast = 0
			MyPriceFilteredLast = ThisPriceValueFiltered
		End If
		If ThisVolumeValueFiltered > 0 Then
			Select Case Direction
				Case FilterRSI.SlopeDirection.NotSpecified
					ThisPriceVariation = ThisPriceValueFiltered - MyPriceFilteredLast
					If ThisPriceVariation < 0 Then
						MyOBVLast = MyOBVLast - ThisVolumeValueFiltered
					ElseIf ThisPriceVariation > 0 Then
						MyOBVLast = MyOBVLast + ThisVolumeValueFiltered
					End If
				Case FilterRSI.SlopeDirection.Positive
					MyOBVLast = MyOBVLast + ThisVolumeValueFiltered
				Case FilterRSI.SlopeDirection.Negative
					MyOBVLast = MyOBVLast - ThisVolumeValueFiltered
			End Select
		End If
		MyPriceFilteredLast = ThisPriceValueFiltered
		Return MyFilterLowPassForOBVOut.Filter(MyOBVLast)
	End Function

	Public Function Filter(ByRef Value() As YahooAccessData.IPriceVol) As Double()
		Dim ThisValue As YahooAccessData.IPriceVol
		For Each ThisValue In Value
			Me.Filter(ThisValue)
		Next
		Return Me.ToArray
	End Function

	Public Function FilterPredictionNext(ByVal Price As Double, ByVal Volume As Integer) As Double
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

	Public Function FilterPredictionNext(ByVal Price As Single, ByVal Volume As Integer) As Double
		Return Me.FilterPredictionNext(CDbl(Price), Volume)
	End Function

	Public Function Filter(ByRef PriceVol As IPriceVol) As Double
		Return Me.Filter(CDbl(PriceVol.Last), DirectCast(PriceVol, PriceVol).Volume)
	End Function

	Public Function FilterPredictionNext(ByRef PriceVol As IPriceVol) As Double
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

	Public ReadOnly Property Rate As Integer
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
