﻿Imports YahooAccessData.MathPlus.Filter

Public Class FilterLCR
	Implements IFilter(Of IStatistical)

	Private MyRate As Integer
	Private FilterValueLastK1 As IStatistical
	Private FilterValueLast As IStatistical
	Private MyStartPoint As Integer
	Private MyVarianceCorrection As Double
	Private IsRunReady As Boolean

	Private ValueLast As Double
	Private ValueLastK1 As Double
	Private IsValueInitial As Boolean
	'Private MyListOfValue As List(Of Double)
	'Private MyListOfValueSquare As List(Of Double)
	'Private MySumOfValue As Double
	'Private MySumOfValueSquare As Double
	Private MyFilterExpMean As FilterExp
	Private MyFilterExpSquare As FilterExp

	Private MyListOfValueStatistical As List(Of IStatistical)

	''' <summary>
	''' Calculate the statistical information from all value 
	''' </summary>
	''' <remarks></remarks>
	Public Sub New()
		Me.New(3)
	End Sub

	''' <summary>
	''' Calculate the statistical information based on a windows of size given by FilterRate
	''' </summary>
	''' <param name="FilterRate">
	''' Should be greater than one, otherwise the calculation is done on all the data
	''' </param>
	''' <remarks></remarks>
	Public Sub New(ByVal FilterRate As Integer)
		'MyListOfValue = New List(Of Double)
		'MyListOfValueSquare = New List(Of Double)
		MyListOfValueStatistical = New List(Of IStatistical)
		'MyListOfValue = New List(Of Double)
		'MyListOfValueSquare = New List(Of Double)
		If FilterRate < 3 Then FilterRate = 3
		MyRate = CInt(FilterRate)
		'the -1 is for the finite number of samples correction from an infinite number of samples
		MyVarianceCorrection = MyRate / (MyRate - 1)
		MyFilterExpMean = New FilterExp(FilterRate)
		MyFilterExpSquare = New FilterExp(FilterRate)

		FilterValueLast = New StatisticalData(0, 0, 0)
		FilterValueLastK1 = New StatisticalData(0, 0, 0)
		ValueLast = 0
		ValueLastK1 = 0
		MyStartPoint = 0
		IsRunReady = False
	End Sub

	''' <summary>
	''' Calculate the statistical information based on a windows of size given by FilterRate
	''' </summary>
	''' <param name="FilterRate"></param>
	''' Should be greater than one, otherwise the calculation is done on all the data
	''' <param name="StartPoint">
	''' The startPoint where the calculation is started
	''' </param>
	''' The startpoin
	''' <remarks></remarks>
	Public Sub New(ByVal FilterRate As Integer, ByVal StartPoint As Integer)
		Me.New(FilterRate)
		MyStartPoint = StartPoint
		IsRunReady = True
	End Sub

	Public Function Filter(Value As Double) As IStatistical Implements IFilter(Of IStatistical).Filter
		Dim ThisM2 As Double
		Dim ThisMean As Double
		Dim ThisVariance As Double

		If MyListOfValueStatistical.Count = 0 Then
			'initialization
			FilterValueLast = New StatisticalData(Mean:=Value, Variance:=0, NumberPoint:=1)
			If IsRunReady = False Then
				MyStartPoint = 0
			End If
		Else
			If IsRunReady = False Then
				If Value <> ValueLast Then
					MyStartPoint = MyListOfValueStatistical.Count
					IsRunReady = True
				End If
			End If
		End If
		'	For J = StartPoint To StopPoint
		'    NewPowerLevel = FFTdB(J)
		'    If NewPowerLevel > NoiseThresholddB Then
		'      CountHit = CountHit + 1
		'    End If

		'	'bound the input value to a valid range
		'	If NewPowerLevel < LCRLBound Then
		'      NewPowerLevel = LCRLBound
		'    ElseIf NewPowerLevel > LCRUBound Then
		'      NewPowerLevel = LCRUBound
		'    End If
		'	'check the input range value
		'	If NewPowerLevel > PowerMax Then
		'	'clear the LCR array
		'	For K = PowerMax + 1 To NewPowerLevel
		'        LCR(K) = 0
		'        #If PDFNoiseMeasurement Then
		'          PDF(K) = 0
		'          CDF(K) = 0
		'        #End If
		'      Next
		'      PowerMax = NewPowerLevel
		'    ElseIf NewPowerLevel < PowerMin Then
		'	'clear the LCR array
		'	For K = NewPowerLevel To PowerMin - 1
		'        LCR(K) = 0
		'        #If PDFNoiseMeasurement Then
		'          PDF(K) = 0
		'          CDF(K) = 0
		'        #End If
		'      Next
		'      PowerMin = NewPowerLevel
		'    End If
		'#If PDFNoiseMeasurement Then
		'      PDF(NewPowerLevel) = PDF(NewPowerLevel) + 1
		'#End If
		'	'calculate the positive and negative LCR
		'	If NewPowerLevel > LastPowerLevel Then
		'	For K = LastPowerLevel To NewPowerLevel - 1
		'        LCR(K) = LCR(K) + 1
		'      Next
		'	ElseIf NewPowerLevel < LastPowerLevel Then
		'	For K = NewPowerLevel + 1 To LastPowerLevel
		'        LCR(K) = LCR(K) + 1
		'      Next
		'	End If
		'    LastPowerLevel = NewPowerLevel
		'  Next


		If IsRunReady Then
			ThisMean = MyFilterExpMean.FilterRun(Value)
			ThisM2 = (Value - ThisMean) ^ 2
			ThisVariance = MyVarianceCorrection * MyFilterExpSquare.FilterRun(ThisM2)
		Else
			ThisMean = Value
			ThisVariance = 0
		End If
		FilterValueLast = New StatisticalData(ThisMean, ThisVariance, MyListOfValueStatistical.Count + 1)
		MyListOfValueStatistical.Add(FilterValueLast)
		ValueLastK1 = ValueLast
		ValueLast = Value
		Return FilterValueLast
	End Function

	Public Function Filter(Value As IPriceVol) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.Filter(Value.Last)
	End Function

	''' <summary>
	''' This function fix the statistical value at the beginning to reflect the result including the data pass the index
	''' up to the rate value
	''' </summary>
	''' <param name="Value"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Function Filter(ByRef Value() As Double) As IStatistical() Implements IFilter(Of IStatistical).Filter
		Dim ThisListOfValueStatistical As List(Of IStatistical)
		Dim ThisValue As Double
		Dim ThisPosition As Integer
		Dim ThisPoint As Integer
		Dim I As Integer

		For Each ThisValue In Value
			Me.Filter(ThisValue)
		Next
		ThisListOfValueStatistical = New List(Of IStatistical)(MyListOfValueStatistical)
		MyListOfValueStatistical.Clear()
		ThisPoint = MyStartPoint + MyRate
		If ThisPoint < ThisListOfValueStatistical.Count Then
			ThisPosition = ThisPoint - 1
		Else
			ThisPosition = ThisListOfValueStatistical.Count - 1
		End If
		Dim ThisStatistical = ThisListOfValueStatistical(ThisPosition)
		For I = 0 To ThisPosition
			MyListOfValueStatistical.Add(ThisStatistical.Copy)
		Next
		For I = I To (ThisListOfValueStatistical.Count - 1)
			MyListOfValueStatistical.Add(ThisListOfValueStatistical(I))
		Next
		Return MyListOfValueStatistical.ToArray
	End Function

	''' <summary>
	''' Special filtering that can be used to remove the delay starting at a specific point
	''' </summary>
	''' <param name="Value">The value to be filtered</param>
	''' <param name="DelayRemovedToItem">The point where the delay start to be removed</param>
	''' <returns>The result</returns>
	''' <remarks></remarks>
	Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As IStatistical() Implements IFilter(Of IStatistical).Filter
		Dim ThisValues(0 To Value.Length - 1) As IStatistical
		Dim I As Integer
		Dim J As Integer

		Dim ThisFilterLeft As New FilterStatistical(Me.Rate)
		Dim ThisFilterRight As New FilterStatistical(Me.Rate)
		Dim ThisFilterLeftItem As IStatistical
		Dim ThisStatisticalValue As IStatistical
		MyListOfValueStatistical.Clear()
		'filter from the left
		ThisFilterLeft.Filter(Value)
		'filter from the right the section with the reverse filtering
		For I = DelayRemovedToItem To 0 Step -1
			ThisFilterRight.Filter(Value(I))
		Next
		'the data in ThisFilterRightList is reversed
		'need to look at it in reverse order using J
		J = DelayRemovedToItem
		For I = 0 To Value.Length - 1
			ThisFilterLeftItem = ThisFilterLeft.ToList(I)
			ThisStatisticalValue = ThisFilterLeftItem.Copy
			If I <= DelayRemovedToItem Then
				ThisStatisticalValue.Add(ThisFilterRight.ToList(J))
			End If
			MyListOfValueStatistical.Add(ThisStatisticalValue)
			ThisValues(I) = ThisStatisticalValue
			J = J - 1
		Next
		Return ThisValues
	End Function

	Public Function FilterErrorLast() As IStatistical Implements IFilter(Of IStatistical).FilterErrorLast
		Throw New NotImplementedException
	End Function

	Public Function FilterBackTo(ByRef Value As IStatistical) As Double Implements IFilter(Of IStatistical).FilterBackTo
		Throw New NotImplementedException
	End Function

	Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).FilterLastToPriceVol
		Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast.Mean))
		With ThisPriceVol
			.LastPrevious = CSng(FilterValueLastK1.Mean)
			If Me.FilterLast.High > .High Then
				.High = CSng(Me.FilterLast.High)
				.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
			End If
			If Me.FilterLast.Low > .Low Then
				.Low = CSng(Me.FilterLast.Low)
			End If
			.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
		End With
		Return ThisPriceVol
	End Function

	Public Function LastToPriceVol() As IPriceVol Implements IFilter(Of IStatistical).LastToPriceVol
		Throw New NotSupportedException
	End Function

	Public Function Filter(Value As Single) As IStatistical Implements IFilter(Of IStatistical).Filter
		Return Me.Filter(CDbl(Value))
	End Function

	Public Function FilterPredictionNext(ByVal Value As Double) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Throw New NotSupportedException
	End Function

	Public Function FilterPredictionNext(ByVal Value As Single) As IStatistical Implements IFilter(Of IStatistical).FilterPredictionNext
		Return Me.FilterPredictionNext(CDbl(Value))
	End Function

	Public Function FilterLast() As IStatistical Implements IFilter(Of IStatistical).FilterLast
		Return FilterValueLast
	End Function

	''' <summary>
	''' </summary>
	''' <returns>Return the last input value</returns>
	Public Function Last() As Double Implements IFilter(Of IStatistical).Last
		Return ValueLast
	End Function

	Public ReadOnly Property Rate As Integer Implements IFilter(Of IStatistical).Rate
		Get
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property Count As Integer Implements IFilter(Of IStatistical).Count
		Get
			Return MyListOfValueStatistical.Count
		End Get
	End Property

	Public ReadOnly Property Max As Double Implements IFilter(Of IStatistical).Max
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property Min As Double Implements IFilter(Of IStatistical).Min
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property ToList() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToList
		Get
			Return MyListOfValueStatistical
		End Get
	End Property

	Public ReadOnly Property ToListOfError() As IList(Of IStatistical) Implements IFilter(Of IStatistical).ToListOfError
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter(Of IStatistical).ToListScaled
		Get
			Throw New NotSupportedException
		End Get
	End Property

	Public Function ToArray() As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As IStatistical() Implements IFilter(Of IStatistical).ToArray
		Return MyListOfValueStatistical.ToArray
	End Function

	Public Property Tag As String Implements IFilter(Of IStatistical).Tag

	Public Overrides Function ToString() As String Implements IFilter(Of IStatistical).ToString
		Return Me.FilterLast.ToString
	End Function
End Class
