Imports YahooAccessData.MathPlus.Filter.FilterVolatility

Public Class FilterRunStatistic
	Implements IFilterRun(Of IStatistical)

	Private MyRate As Integer
	Private MyListOfRecord As IList(Of IStatistical)
	Private MyFilterRunStatistic As IFilterRun(Of IStatistical)

#Region "New"
	''' <summary>
	''' Calculate the statistical information based on a windows size given by FilterRate. 
	''' This is a basic implementation for the statistical measurement using the square windows only.
	''' It does not use any exponential filtering and follow the usual statistic standard method.
	''' </summary>
	''' <param name="FilterRate">
	''' Should be set to be greater than two
	''' </param>
	''' <remarks></remarks>
	Public Sub New(
		ByVal FilterRate As Integer,
		Optional StatisticType As enuVolatilityStatisticType = enuVolatilityStatisticType.Standard,
		Optional IsListRecordingEnabled As Boolean = False,
		Optional BufferCapacity As Integer = 0)

		_IsListRecordingEnabled = IsListRecordingEnabled
		MyListOfRecord = New List(Of IStatistical)
		If FilterRate < 2 Then FilterRate = 2
		MyRate = CInt(FilterRate)
		Select Case StatisticType
			Case enuVolatilityStatisticType.Exponential
				MyFilterRunStatistic = New StatisticExponential(FilterRate, BufferCapacity:=BufferCapacity)
			Case enuVolatilityStatisticType.Standard
				MyFilterRunStatistic = New StatisticWindows(FilterRate, BufferCapacity:=BufferCapacity)
		End Select
	End Sub
#End Region

	Private _IsListRecordingEnabled As Boolean
	Public ReadOnly IsListRecordingEnabled As Boolean

	Public ReadOnly Property InputLast As Double Implements IFilterRun(Of IStatistical).InputLast
		Get
			Return MyFilterRunStatistic.InputLast
		End Get
	End Property

	''' <summary>
	''' Returns the last value of the filter run. Index 0 is the most recent value added to the filter.
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property FilterLast(Index As Integer) As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyFilterRunStatistic.FilterLast(Index)
		End Get
	End Property

	Private ReadOnly Property FilterTrendLast As IStatistical Implements IFilterRun(Of IStatistical).FilterTrendLast
		Get
			Throw New NotImplementedException("property 'FilterTrendLast' is not implemented in this context...")
		End Get
	End Property

	Public ReadOnly Property FilterRate As Double Implements IFilterRun(Of IStatistical).FilterRate
		Get
			Return MyRate
		End Get
	End Property

	Public ReadOnly Property ToList As IList(Of IStatistical)
		Get
			Return MyListOfRecord
		End Get
	End Property

	Public ReadOnly Property ToBufferList As IList(Of IStatistical) Implements IFilterRun(Of IStatistical).ToBufferList
		Get
			Return MyFilterRunStatistic.ToBufferList()
		End Get
	End Property

	Public ReadOnly Property FilterDetails As String Implements IFilterRun(Of IStatistical).FilterDetails
		Get
			Return $"{Me.GetType().Name}({Me.FilterRate})"
		End Get
	End Property

	Public Sub Reset() Implements IFilterRun(Of IStatistical).Reset
		MyFilterRunStatistic.Reset()
	End Sub

	Public Sub Reset(BufferCapacity As Integer) Implements IFilterRun(Of IStatistical).Reset
		MyFilterRunStatistic.Reset(BufferCapacity)
	End Sub
	Public ReadOnly Property IsReset As Boolean Implements IFilterRun(Of IStatistical).IsReset
		Get
			Return MyFilterRunStatistic.IsReset
		End Get
	End Property

	Public ReadOnly Property FilterLast As IStatistical Implements IFilterRun(Of IStatistical).FilterLast
		Get
			Return MyFilterRunStatistic.FilterLast
		End Get
	End Property

	Public Function FilterRun(Value As Double) As IStatistical Implements IFilterRun(Of IStatistical).FilterRun
		Dim ThisResult = MyFilterRunStatistic.FilterRun(Value)
		If Me.IsListRecordingEnabled Then
			MyListOfRecord.Add(ThisResult)
		End If
		Return ThisResult
	End Function
End Class


