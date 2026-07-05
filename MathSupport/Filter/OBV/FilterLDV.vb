Imports Newtonsoft.Json.Linq
Imports YahooAccessData.MathPlus
Imports YahooAccessData.MathPlus.Filter

Public Class FilterLDV
	Private Const FILTER_RATE_FOR_AVERAGE_VOLUME As Integer = NUMBER_TRADINGDAY_PER_YEAR \ 12
	Private MyRate As Double
	Private MyLDVLast As Double
	Private MyOBLDVLast As Double
	Private MyLDVLogFilteredLast As Double
	Private MyFilterForLDVAverageExp As IFilterRun
	Private MyFilterLowPassForOBLDVOut As IFilter

	Public Sub New(ByVal FilterRate As Double, Optional FilterRateForAverageVolume As Double = FILTER_RATE_FOR_AVERAGE_VOLUME)
		If FilterRate < 1 Then FilterRate = 1
		MyRate = FilterRate
		MyFilterForLDVAverageExp = New FilterExp(FilterRateForAverageVolume)
		MyFilterLowPassForOBLDVOut = New FilterLowPassExp(FilterRate)
	End Sub

	''' <summary>
	''' Use this call to measure the accelerating A-OBV base on the price direction that establish the positivity or not of the incoming volume
	''' </summary>
	''' <param name="LDV"> The logarithmic dollar value of the trading volume</param>
	''' <param name="Direction">
	''' The direction of the price movement associated with this volume can be of 4 type: as defined here:
	'''     NotSpecified : the direction is not specified, the function will try to determine it based on the price movement
	'''     In that case the function behave more or less like teh standard OBV definition
	'''     Positive : the price is moving up
	'''     Negative : the price is moving down
	'''     Sideways : the price is moving sideways or the same as the previous sample
	''' </param>
	''' <returns></returns>
	Public Function Filter(ByVal LDV As Double, ByVal Direction As FilterRSI.SlopeDirection) As Double
		Dim ThisLDVLogFiltered As Double
		Dim ThisLDVLongTermAverageFiltered As Double

		If MyFilterLowPassForOBLDVOut.Count = 0 Then
			'OBV initialisation 
			MyOBLDVLast = 0
			MyLDVLogFilteredLast = 0
			MyLDVLast = 0
		End If
		If LDV = 0 Then
			'use the last values
			ThisLDVLongTermAverageFiltered = MyFilterForLDVAverageExp.FilterLast
			ThisLDVLogFiltered = MyLDVLogFilteredLast
		Else
			ThisLDVLongTermAverageFiltered = MyFilterForLDVAverageExp.FilterRun(LDV)
			'applied a logaritmic fucntion to the volume
			'this is equivalent to a ratio of volume over average volume	
			'with likely a small bias because with are averaging the Log value and not the value itself
			ThisLDVLogFiltered = LDV - ThisLDVLongTermAverageFiltered
		End If
		If ThisLDVLogFiltered > 0 Then
			'volume increasing relative to an average is interpreted as significant
			'this is only when we start to change the OBV.
			Select Case Direction
				Case FilterRSI.SlopeDirection.Positive
					MyOBLDVLast = MyOBLDVLast + ThisLDVLogFiltered
				Case FilterRSI.SlopeDirection.Negative
					MyOBLDVLast = MyOBLDVLast - ThisLDVLogFiltered
			End Select
		ElseIf ThisLDVLogFiltered < 0 Then
			'decreasing volume in any direction is difficult ot interpret
			'so do nothign when the volume are decreasing for now
			''volume decrease
			'Select Case Direction
			'	Case FilterRSI.SlopeDirection.Positive
			'		MyOBLDVLast = MyOBLDVLast + 0.1 * ThisLDVLogFiltered
			'	Case FilterRSI.SlopeDirection.Negative
			'		MyOBLDVLast = MyOBLDVLast - 0.1 * ThisLDVLogFiltered
			'End Select
		End If
		MyLDVLogFilteredLast = ThisLDVLogFiltered
		MyLDVLast = LDV
		'note OBLDV = On Balance Logarithmic Dollar Value
		Return MyFilterLowPassForOBLDVOut.Filter(MyOBLDVLast)
	End Function

	Public Function FilterLast() As Double
		Return MyFilterLowPassForOBLDVOut.FilterLast
	End Function

	Public Function LDVLast() As Double
		Return MyLDVLast
	End Function

	Public ReadOnly Property Rate As Double
		Get
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer
		Get
			Return MyFilterLowPassForOBLDVOut.Count
		End Get
	End Property

	Public ReadOnly Property Max As Double
		Get
			Return MyFilterLowPassForOBLDVOut.Max
		End Get
	End Property

	Public ReadOnly Property Min As Double
		Get
			Return MyFilterLowPassForOBLDVOut.Min
		End Get
	End Property

	Public ReadOnly Property ToList() As IList(Of Double)
		Get
			Return MyFilterLowPassForOBLDVOut.ToList
		End Get
	End Property

	Public Function ToArray() As Double()
		Return MyFilterLowPassForOBLDVOut.ToArray
	End Function

	Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
		Return Me.ToArray(Me.Min, Me.Max, ScaleToMinValue, ScaleToMaxValue)
	End Function

	Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
		Return MyFilterLowPassForOBLDVOut.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
	End Function

	Public Property Tag As String

	Public Overrides Function ToString() As String
		Return Me.FilterLast.ToString
	End Function
End Class
