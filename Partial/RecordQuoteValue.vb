Imports YahooAccessData.ExtensionService

Public Class RecordQuoteValue
  Implements IRecordQuoteValue
  Implements IFormatData
  Implements IEquatable(Of RecordQuoteValue)
  Implements IRegisterKey(Of Date)
  Implements IComparable(Of RecordQuoteValue)
  Implements IDataPosition
  Implements IDateUpdate
  Implements IDateTrade
  Implements IDisposable
  Implements ISentimentIndicator

  Private MyRecord As Record
  Private MyRecordQuoteValue As IRecordQuoteValue
  Private MySentimentIndicator As ISentimentIndicator
  'Private Shared MyCompareByName As CompareByName(Of RecordQuoteValue)

  Public Sub New()
    Me.New(New Record)
  End Sub
  Public Sub New(ByVal RecordQuoteValue As Record)
    MyRecord = RecordQuoteValue
    MyRecordQuoteValue = MyRecord
    MySentimentIndicator = MyRecord
    'If MyCompareByName Is Nothing Then
    '  MyCompareByName = New CompareByName(Of RecordQuoteValue)
    'End If
  End Sub

  Public ReadOnly Property Record As Record
    Get
      Return MyRecord
    End Get
  End Property

	Public ReadOnly Property ToDataPosition As IDataPosition
		Get
			Return Me
		End Get
	End Property

	Public Property Volume As Long
		Get
			Return Record.Volume
		End Get
		Set(value As Long)
			Record.Volume = value
		End Set
	End Property

#Region "IRecordQuoteValue"
	Private Sub IRecordQuoteValue_PriceChange(Open As Single, Low As Single, High As Single, Last As Single) Implements IRecordQuoteValue.PriceChange
    MyRecordQuoteValue.PriceChange(Open, Low, High, Last)
  End Sub

  Public ReadOnly Property AfterHoursChangeRealtime As Single Implements IRecordQuoteValue.AfterHoursChangeRealtime
    Get
      Return MyRecordQuoteValue.AfterHoursChangeRealtime
    End Get
  End Property

  Public ReadOnly Property AnnualizedGain As Single Implements IRecordQuoteValue.AnnualizedGain
    Get
      Return MyRecordQuoteValue.AnnualizedGain
    End Get
  End Property

  Public ReadOnly Property Ask As Single Implements IRecordQuoteValue.Ask
    Get
      Return MyRecordQuoteValue.Ask
    End Get
  End Property

  Public ReadOnly Property AskRealTime As Single Implements IRecordQuoteValue.AskRealTime
    Get
      Return MyRecordQuoteValue.AskRealTime
    End Get
  End Property

  Public ReadOnly Property AskSize As Integer Implements IRecordQuoteValue.AskSize
    Get
      Return MyRecordQuoteValue.AskSize
    End Get
  End Property

  Public ReadOnly Property AverageDailyVolume As Integer Implements IRecordQuoteValue.AverageDailyVolume
    Get
      Return MyRecordQuoteValue.AverageDailyVolume
    End Get
  End Property

  Public ReadOnly Property Bid As Single Implements IRecordQuoteValue.Bid
    Get
      Return MyRecordQuoteValue.Bid
    End Get
  End Property

  Public ReadOnly Property BidRealtime As Single Implements IRecordQuoteValue.BidRealtime
    Get
      Return MyRecordQuoteValue.BidRealtime
    End Get
  End Property

  Public ReadOnly Property BidSize As Integer Implements IRecordQuoteValue.BidSize
    Get
      Return MyRecordQuoteValue.BidSize
    End Get
  End Property

  Public ReadOnly Property BookValue As Single Implements IRecordQuoteValue.BookValue
    Get
      Return MyRecordQuoteValue.BookValue
    End Get
  End Property

  Public ReadOnly Property DateDay As Date Implements IRecordQuoteValue.DateDay
    Get
      Return MyRecordQuoteValue.DateDay
    End Get
  End Property

  Public ReadOnly Property DateLastTrade As Date Implements IRecordQuoteValue.DateLastTrade
    Get
      Return MyRecordQuoteValue.DateLastTrade
    End Get
  End Property

  Public ReadOnly Property DateUpdate As Date Implements IRecordQuoteValue.DateUpdate
    Get
      Return MyRecordQuoteValue.DateUpdate
    End Get
  End Property

  Public ReadOnly Property DividendPayDate As Date Implements IRecordQuoteValue.DividendPayDate
    Get
      Return MyRecordQuoteValue.DividendPayDate
    End Get
  End Property

  Public ReadOnly Property DividendShare As Single Implements IRecordQuoteValue.DividendShare
    Get
      Return MyRecordQuoteValue.DividendShare
    End Get
  End Property

  Public ReadOnly Property DividendYield As Single Implements IRecordQuoteValue.DividendYield
    Get
      Return MyRecordQuoteValue.DividendYield
    End Get
  End Property

  Public ReadOnly Property EarningsShare As Single Implements IRecordQuoteValue.EarningsShare
    Get
      Return MyRecordQuoteValue.EarningsShare
    End Get
  End Property

  Public ReadOnly Property EBITDA As Single Implements IRecordQuoteValue.EBITDA
    Get
      Return MyRecordQuoteValue.EBITDA
    End Get
  End Property

  Public ReadOnly Property EPSEstimateCurrentYear As Single Implements IRecordQuoteValue.EPSEstimateCurrentYear
    Get
      Return MyRecordQuoteValue.EPSEstimateCurrentYear
    End Get
  End Property

  Public ReadOnly Property EPSEstimateNextQuarter As Single Implements IRecordQuoteValue.EPSEstimateNextQuarter
    Get
      Return MyRecordQuoteValue.EPSEstimateNextQuarter
    End Get
  End Property

  Public ReadOnly Property EPSEstimateNextYear As Single Implements IRecordQuoteValue.EPSEstimateNextYear
    Get
      Return MyRecordQuoteValue.EPSEstimateNextYear
    End Get
  End Property

  Public ReadOnly Property ExDividendDate As Date Implements IRecordQuoteValue.ExDividendDate
    Get
      Return MyRecordQuoteValue.ExDividendDate
    End Get
  End Property

  Public ReadOnly Property FloatShares As Integer Implements IRecordQuoteValue.FloatShares
    Get
      Return MyRecordQuoteValue.FloatShares
    End Get
  End Property

  Public ReadOnly Property High As Single Implements IRecordQuoteValue.High
    Get
      Return MyRecordQuoteValue.High
    End Get
  End Property

  Public ReadOnly Property HoldingsValue As Single Implements IRecordQuoteValue.HoldingsValue
    Get
      Return MyRecordQuoteValue.HoldingsValue
    End Get
  End Property

  Public ReadOnly Property HoldingsValueRealtime As Single Implements IRecordQuoteValue.HoldingsValueRealtime
    Get
      Return MyRecordQuoteValue.HoldingsValueRealtime
    End Get
  End Property

  Public ReadOnly Property ID As Integer Implements IRecordQuoteValue.ID
    Get
      Return MyRecordQuoteValue.ID
    End Get
  End Property

  Public ReadOnly Property Last As Single Implements IRecordQuoteValue.Last
    Get
      Return MyRecordQuoteValue.Last
    End Get
  End Property

  Public ReadOnly Property LastTradeDate As Date Implements IRecordQuoteValue.LastTradeDate
    Get
      Return MyRecordQuoteValue.LastTradeDate
    End Get
  End Property

  Public ReadOnly Property LastTradePriceOnly As Single Implements IRecordQuoteValue.LastTradePriceOnly
    Get
      Return MyRecordQuoteValue.LastTradePriceOnly
    End Get
  End Property

  Public ReadOnly Property LastTradeRealtimeWithTime As Date Implements IRecordQuoteValue.LastTradeRealtimeWithTime
    Get
      Return MyRecordQuoteValue.LastTradeRealtimeWithTime
    End Get
  End Property

  Public ReadOnly Property LastTradeSize As Integer Implements IRecordQuoteValue.LastTradeSize
    Get
      Return MyRecordQuoteValue.LastTradeSize
    End Get
  End Property

  Public ReadOnly Property Low As Single Implements IRecordQuoteValue.Low
    Get
      Return MyRecordQuoteValue.Low
    End Get
  End Property

  Public ReadOnly Property MarketCapitalization As Single Implements IRecordQuoteValue.MarketCapitalization
    Get
      Return MyRecordQuoteValue.MarketCapitalization
    End Get
  End Property

  Public ReadOnly Property MarketCapRealtime As Single Implements IRecordQuoteValue.MarketCapRealtime
    Get
      Return MyRecordQuoteValue.MarketCapRealtime
    End Get
  End Property

  Public ReadOnly Property OneyrTargetPrice As Single Implements IRecordQuoteValue.OneyrTargetPrice
    Get
      Return MyRecordQuoteValue.OneyrTargetPrice
    End Get
  End Property

  Public ReadOnly Property Open As Single Implements IRecordQuoteValue.Open
    Get
      Return MyRecordQuoteValue.Open
    End Get
  End Property

  Public ReadOnly Property OrderBookRealtime As Single Implements IRecordQuoteValue.OrderBookRealtime
    Get
      Return MyRecordQuoteValue.OrderBookRealtime
    End Get
  End Property

  Public ReadOnly Property PEGRatio As Single Implements IRecordQuoteValue.PEGRatio
    Get
      Return MyRecordQuoteValue.PEGRatio
    End Get
  End Property

  Public ReadOnly Property PERatio As Single Implements IRecordQuoteValue.PERatio
    Get
      Return MyRecordQuoteValue.PERatio
    End Get
  End Property

  Public ReadOnly Property PERatioRealtime As Single Implements IRecordQuoteValue.PERatioRealtime
    Get
      Return MyRecordQuoteValue.PERatioRealtime
    End Get
  End Property

  Public ReadOnly Property PriceBook As Single Implements IRecordQuoteValue.PriceBook
    Get
      Return MyRecordQuoteValue.PriceBook
    End Get
  End Property

  Public ReadOnly Property PriceEPSEstimateCurrentYear As Single Implements IRecordQuoteValue.PriceEPSEstimateCurrentYear
    Get
      Return MyRecordQuoteValue.PriceEPSEstimateCurrentYear
    End Get
  End Property

  Public ReadOnly Property PriceEPSEstimateNextYear As Single Implements IRecordQuoteValue.PriceEPSEstimateNextYear
    Get
      Return MyRecordQuoteValue.PriceEPSEstimateNextYear
    End Get
  End Property

  Public ReadOnly Property PriceSales As Single Implements IRecordQuoteValue.PriceSales
    Get
      Return MyRecordQuoteValue.PriceSales
    End Get
  End Property

  Public ReadOnly Property SharesOwned As Integer Implements IRecordQuoteValue.SharesOwned
    Get
      Return MyRecordQuoteValue.SharesOwned
    End Get
  End Property

  Public ReadOnly Property ShortRatio As Single Implements IRecordQuoteValue.ShortRatio
    Get
      Return MyRecordQuoteValue.ShortRatio
    End Get
  End Property

  Public ReadOnly Property StockExchange As String Implements IRecordQuoteValue.StockExchange
    Get
      Return MyRecordQuoteValue.StockExchange
    End Get
  End Property

  Public ReadOnly Property TickerTrend As Single Implements IRecordQuoteValue.TickerTrend
    Get
      Return MyRecordQuoteValue.TickerTrend
    End Get
  End Property

  Public ReadOnly Property Vol As Integer Implements IRecordQuoteValue.Vol
		Get
			Return MyRecordQuoteValue.Vol
		End Get
	End Property

  Public Sub RemoveBookEarningGlitch(RecordQuoteValueLast As IRecordQuoteValue) Implements IRecordQuoteValue.RemoveBookEarningGlitch
    MyRecordQuoteValue.RemoveBookEarningGlitch(RecordQuoteValueLast)
  End Sub
#End Region
#Region "IFormatData"
  Public Function ToListOfHeader() As System.Collections.Generic.List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Dim ThisHeader As New List(Of HeaderInfo)(MyRecord.ToListOfHeader)
    Dim ThisQuoteValue As QuoteValue

    If MyRecord.QuoteValues.Count = 0 Then
      ThisQuoteValue = New QuoteValue
    Else
      ThisQuoteValue = MyRecord.QuoteValues(0)
    End If
    For Each ThisQuoteHeader In ThisQuoteValue.ToListOfHeader
      If ThisQuoteHeader.Visible Then
        ThisHeader.Add(ThisQuoteHeader)
      End If
    Next
    Return ThisHeader
  End Function

  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Dim ThisQuoteValue As QuoteValue

    If MyRecord.QuoteValues.Count = 0 Then
      ThisQuoteValue = New QuoteValue
    Else
      ThisQuoteValue = MyRecord.QuoteValues(0)
    End If
    Return Extensions.ToStingOfData(Of Record)(MyRecord).Concat(Extensions.ToStingOfData(Of QuoteValue)(ThisQuoteValue)).ToArray
  End Function
#End Region
#Region "IEquatable"
  Private Function IEquatable_Equals(other As RecordQuoteValue) As Boolean Implements System.IEquatable(Of RecordQuoteValue).Equals
    Return MyRecord.Equals(other.Record)
  End Function
#End Region
#Region "IRegisterKey"
  Private Property KeyID As Integer Implements IRegisterKey(Of Date).KeyID
    Get
      Return MyRecord.KeyID
    End Get
    Set(value As Integer)
      'this object is read only
    End Set
  End Property

  Private Property KeyValue As Date Implements IRegisterKey(Of Date).KeyValue
    Get
      Return MyRecord.KeyValue
    End Get
    Set(value As Date)

    End Set
  End Property
#End Region
#Region "IComparable"
  Private Function CompareTo(other As RecordQuoteValue) As Integer Implements System.IComparable(Of RecordQuoteValue).CompareTo
    Return MyRecord.CompareTo(other.Record)
  End Function
#End Region
#Region "IDataPosition"
  Private Property IDataPosition_Current As Long Implements IDataPosition.Current
    Get
      Return MyRecord.ToDataPosition.Current
    End Get
    Set(value As Long)
      MyRecord.ToDataPosition.Current = value
    End Set
  End Property
  Private Property IDataPosition_ToNext As Long Implements IDataPosition.ToNext
    Get
      Return MyRecord.ToDataPosition.ToNext
    End Get
    Set(value As Long)
      MyRecord.ToDataPosition.ToNext = value
    End Set
  End Property
  Private Property IDataPosition_ToPrevious As Long Implements IDataPosition.ToPrevious
    Get
      Return MyRecord.ToDataPosition.ToPrevious
    End Get
    Set(value As Long)
      MyRecord.ToDataPosition.ToPrevious = value
    End Set
  End Property
#End Region 'IDataPosition
#Region "IDateUpdate"
  Private Property IDateUpdate_DateStart As Date Implements IDateUpdate.DateStart
    Get
      Return Me.DateUpdate
    End Get
    Set(value As Date)

    End Set
  End Property
  Private Property IDateUpdate_DateStop As Date Implements IDateUpdate.DateStop
    Get
      Return Me.DateUpdate
    End Get
    Set(value As Date)

    End Set
  End Property
  Private ReadOnly Property IDateUpdate_DateUpdate As Date Implements IDateUpdate.DateUpdate
    Get
      Return Me.DateUpdate
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateLastTrade As Date Implements IDateUpdate.DateLastTrade
    Get
      Return Me.DateLastTrade
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateDay
    End Get
  End Property
#End Region
#Region "IDateTrade"
  Private Property IDateTrade_DateStart As Date Implements IDateTrade.DateStart
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)

    End Set
  End Property
  Private Property IDateTrade_DateStop As Date Implements IDateTrade.DateStop
    Get
      Return Me.DateLastTrade
    End Get
    Set(value As Date)

    End Set
  End Property

  Public Property SentimentCount As Integer Implements ISentimentIndicator.Count
    Get
      Return MySentimentIndicator.Count
    End Get
    Set(value As Integer)
      Throw New NotImplementedException()
    End Set
  End Property

  Public Property SentimentValue As Double Implements ISentimentIndicator.Value
    Get
      Return MySentimentIndicator.Value
    End Get
    Set(value As Double)
      Throw New NotImplementedException()
    End Set
  End Property
#End Region
#Region "IDisposable Support"
  Private disposedValue As Boolean ' To detect redundant calls

  ' IDisposable
  Protected Overridable Sub Dispose(disposing As Boolean)
    If Not Me.disposedValue Then
      If disposing Then
        ' dispose managed state (managed objects).
      End If

      ' free unmanaged resources (unmanaged objects) and override Finalize() below.
      ' set large fields to null.
    End If
    Me.disposedValue = True
  End Sub

  'override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
  'Protected Overrides Sub Finalize()
  '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
  '    Dispose(False)
  '    MyBase.Finalize()
  'End Sub

  ' This code added by Visual Basic to correctly implement the disposable pattern.
  Public Sub Dispose() Implements IDisposable.Dispose
    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    Dispose(True)
    GC.SuppressFinalize(Me)
  End Sub

  Public Function AsISentimentIndicator() As ISentimentIndicator Implements ISentimentIndicator.AsISentimentIndicator
    Return Me
  End Function
#End Region
End Class

#Region "IRecordQuoteValue Interface"
Public Interface IRecordQuoteValue
  Sub RemoveBookEarningGlitch(ByVal RecordQuoteValueLast As IRecordQuoteValue)
  Sub PriceChange(ByVal Open As Single, ByVal Low As Single, ByVal High As Single, ByVal Last As Single)
  ReadOnly Property ID As Integer
  ReadOnly Property DateDay As Date
  ReadOnly Property DateUpdate As Date
  ReadOnly Property DateLastTrade As Date
  ReadOnly Property Open As Single
  ReadOnly Property Last As Single
  ReadOnly Property High As Single
  ReadOnly Property Low As Single
  ReadOnly Property Vol As Integer
  ReadOnly Property AfterHoursChangeRealtime As Single
  ReadOnly Property AnnualizedGain As Single
  ReadOnly Property Ask As Single
  ReadOnly Property AskRealTime As Single
  ReadOnly Property AskSize As Integer
  ReadOnly Property AverageDailyVolume As Integer
  ReadOnly Property Bid As Single
  ReadOnly Property BidRealtime As Single
  ReadOnly Property BidSize As Integer
  ReadOnly Property BookValue As Single
  ReadOnly Property DividendPayDate As Date
  ReadOnly Property DividendShare As Single
  ReadOnly Property DividendYield As Single
  ReadOnly Property EarningsShare As Single
  ReadOnly Property EBITDA As Single
  ReadOnly Property EPSEstimateCurrentYear As Single
  ReadOnly Property EPSEstimateNextQuarter As Single
  ReadOnly Property EPSEstimateNextYear As Single
  ReadOnly Property ExDividendDate As Date
  ReadOnly Property FloatShares As Integer
  ReadOnly Property HoldingsValue As Single
  ReadOnly Property HoldingsValueRealtime As Single
  ReadOnly Property LastTradeDate As Date
  ReadOnly Property LastTradePriceOnly As Single
  ReadOnly Property LastTradeRealtimeWithTime As Date
  ReadOnly Property LastTradeSize As Integer
  ReadOnly Property MarketCapitalization As Single
  ReadOnly Property MarketCapRealtime As Single
  ReadOnly Property OneyrTargetPrice As Single
  ReadOnly Property OrderBookRealtime As Single
  ReadOnly Property PEGRatio As Single
  ReadOnly Property PERatio As Single
  ReadOnly Property PERatioRealtime As Single
  ReadOnly Property PriceBook As Single
  ReadOnly Property PriceEPSEstimateCurrentYear As Single
  ReadOnly Property PriceEPSEstimateNextYear As Single
  'ReadOnly Property PriceEstimateLastYear As Single
  'ReadOnly Property PriceEstimateNextQuarter As Single
  'ReadOnly Property PriceEstimateCurrentYear As Single
  'ReadOnly Property PriceEstimateNextYear As Single
  'ReadOnly Property PriceEstimateFromPEG As Single
  'ReadOnly Property FiscalYearEnd As Date
  ReadOnly Property PriceSales As Single
  ReadOnly Property SharesOwned As Integer
  ReadOnly Property ShortRatio As Single
  ReadOnly Property TickerTrend As Single
  ReadOnly Property StockExchange As String
End Interface
#End Region
