﻿Imports YahooAccessData.MathPlus.Filter

''' <summary>
''' Count the number of time 
''' </summary>
Public Class StatisticLCR
	Implements IFilterRun

	Public Sub New(ByVal FilterRate As Integer)

	End Sub

	Public ReadOnly Property FilterLast As Double Implements IFilterRun.FilterLast
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun.FilterRate
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun.FilterDetails
		Get
			Throw New NotImplementedException()
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun.Reset
		Throw New NotImplementedException()
	End Sub

	Public Function FilterRun(Value As Double) As Double Implements IFilterRun.FilterRun
		Throw New NotImplementedException()
	End Function
End Class
