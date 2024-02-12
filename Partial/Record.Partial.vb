#Region "Imports"
Imports YahooAccessData.ExtensionService
Imports System.IO
#End Region

#Region "Record"
<Serializable()>
Partial Public Class Record
  Implements IRecordQuoteValue
  Implements IEquatable(Of Record)
  Implements IRegisterKey(Of Date)
  Implements IComparable(Of Record)
  Implements IMemoryStream
  Implements IFormatData
  Implements IDataPosition
  Implements IDateUpdate
  Implements IDateTrade
  Implements IPriceVol
  Implements IRecordType

#Region "Main"
  Private _RecordType As IRecordType.enuRecordType
  Private MyPriceRange As Single
  Private MyPriceLastPrevious As Single
  Private MyException As Exception
  Private Shared MyListHeaderInfo As List(Of HeaderInfo)
#End Region
#Region "New"
  Public Sub New(ByVal DateUpdate As Date)
    With Me
      With .ToDataPosition
        .Current = -1
        .ToNext = -1
        .ToPrevious = -1
      End With
      '.DateUpdate = Now  'replaced sept 2017
      .DateUpdate = DateUpdate
      .DateLastTrade = .DateUpdate
      .DateDay = .DateUpdate.Date
      .FinancialHighlights = New LinkedHashSet(Of FinancialHighlight, Date)
      .MarketQuoteDatas = New LinkedHashSet(Of MarketQuoteData, Date)
      .QuoteValues = New LinkedHashSet(Of QuoteValue, Date)
      .TradeInfoes = New LinkedHashSet(Of TradeInfo, Date)
      .ValuationMeasures = New LinkedHashSet(Of ValuationMeasure, Date)
    End With
    If MyListHeaderInfo Is Nothing Then
      Dim ThisFile = My.Application.Info.DirectoryPath & "\HeaderInfo\" & TypeName(Me) & ".HeaderInfo.json"
      MyListHeaderInfo = FileHeaderRead(ThisFile, ListOfHeader, Me.Exception)
    End If
    _RecordType = IRecordType.enuRecordType.EndOfDay  'by default
    'the record is nor considered changed on initialization
    _IsRecordChanged = False
  End Sub

  Public Sub New()
    Me.New(Now)
  End Sub

  Public Sub New(ByRef Parent As Stock, ByRef Stream As Stream)
    Me.New()
    With Me
      .SerializeLoadFrom(Stream)
      .Stock = Parent
      If .Stock IsNot Nothing Then
        .StockID = .Stock.ID
        .Stock.Records.Add(Me)
      End If
    End With
  End Sub
#End Region
#Region "ToData"
  Public ReadOnly Property ToDataPosition As IDataPosition
    Get
      Return Me
    End Get
  End Property

  Friend Function CopyDeep(ByRef Parent As Stock, Optional ByVal IsIgnoreID As Boolean = False) As YahooAccessData.Record
    Dim ThisRecord = New Record
    'Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    With ThisRecord
      If IsIgnoreID = False Then .ID = Me.ID
      .DateDay = Me.DateDay
      .DateUpdate = Me.DateUpdate
      .DateLastTrade = Me.DateLastTrade
      .Open = Me.Open
      .Last = Me.Last
      .High = Me.High
      .Low = Me.Low
      .Vol = Me.Vol
      .Stock = Parent
      If .Stock IsNot Nothing Then
        .StockID = .Stock.ID
        .Stock.Records.Add(ThisRecord)
      End If
    End With
    'ThisStopWatch.Restart()
    Me.QuoteValues.CopyDeep(ThisRecord, IsIgnoreID)
    'ThisStopWatch.Stop()
    Me.MarketQuoteDatas.CopyDeep(ThisRecord, IsIgnoreID)
    Me.FinancialHighlights.CopyDeep(ThisRecord, IsIgnoreID)
    Me.TradeInfoes.CopyDeep(ThisRecord, IsIgnoreID)
    Me.ValuationMeasures.CopyDeep(ThisRecord, IsIgnoreID)
    Return ThisRecord
  End Function

  ''' <summary>
  ''' This function check that the transaction high low are consistent with the open and last
  ''' </summary>
  ''' <returns>Return True if the record is consistent</returns>
  ''' <remarks>This function can also optionally fix the issues </remarks>
  Public Shared Function CheckPrice(
                                   ByRef ThisPriceVol As IPriceVol,
                                   Optional ByVal IsApplyCorrection As Boolean = False,
                                   Optional ByVal IsCorrectionOnVolume As Boolean = False) As Boolean
    With ThisPriceVol
      If IsCorrectionOnVolume Then
        If .Vol = 0 Then
          If .Low <> .Last Then
            If IsApplyCorrection Then
              .Low = .Last
            Else
              Return False
            End If
          End If
          If .Open <> .Last Then
            If IsApplyCorrection Then
              .Open = .Last
            Else
              Return False
            End If
          End If
          If .High <> .Last Then
            If IsApplyCorrection Then
              .High = .Last
            Else
              Return False
            End If
          End If
          Return True
        ElseIf .Last = 0 Then
          'we know already that the volume is > 0
          'we cannot have a volume > 0 with the last transaction value = 0
          'make sure everything is zero including the volume
          If IsApplyCorrection Then
            .Vol = 0
            .Open = 0
            .High = 0
            .Low = 0
          Else
            Return False
          End If
          Return True
        ElseIf .Open = 0 Then
          'we know already that last is not zero and that Volume is greater than zero
          'this is an error but estimate that the open is the same than the last
          If IsApplyCorrection Then
            .Open = .Last
            'we need to fix here the case where .Low=0
            If .Low = 0 Then
              If .Open < .Last Then
                .Low = .Open
              Else
                .Low = .Last
              End If
            End If
            'this will fix the case where .High is zero
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        ElseIf .Low = 0 Then
          'we know already that last and open is not zero and that Volume is greater than zero
          If IsApplyCorrection Then
            If .Open < .Last Then
              .Low = .Open
            Else
              .Low = .Last
            End If
            'this will fix the high=0 error condition
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        ElseIf .High = 0 Then
          'we know already that last and open and low are not zero and that Volume is greater than zero
          'this will fix the high=0 error condition
          If IsApplyCorrection Then
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        End If
      Else
        If .Last = 0 Then
          'we know already that the volume is > 0
          'we cannot have a volume > 0 with the last transaction value = 0
          'make sure everything is zero including the volume
          If IsApplyCorrection Then
            .Vol = 0
            .Open = 0
            .High = 0
            .Low = 0
          Else
            Return False
          End If
          Return True
        ElseIf .Open = 0 Then
          'we know already that last is not zero and that Volume is greater than zero
          'this is an error but estimate that the open is the same than the last
          If IsApplyCorrection Then
            .Open = .Last
            'we need to fix here the case where .Low=0
            If .Low = 0 Then
              If .Open < .Last Then
                .Low = .Open
              Else
                .Low = .Last
              End If
            End If
            'this will fix the case where .High is zero
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        ElseIf .Low = 0 Then
          'we know already that last and open is not zero and that Volume is greater than zero
          If IsApplyCorrection Then
            If .Open < .Last Then
              .Low = .Open
            Else
              .Low = .Last
            End If
            'this will fix the high=0 error condition
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        ElseIf .High = 0 Then
          'we know already that last and open and low are not zero and that Volume is greater than zero
          'this will fix the high=0 error condition
          If IsApplyCorrection Then
            CheckHighLow(ThisPriceVol, True)
          Else
            Return False
          End If
        End If
      End If
    End With
    Return True
  End Function

  ''' <summary>
  ''' Function use locally just to correct the high low anomalies based on open and last
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>The function does not check the validity of open and last</remarks>
  Private Shared Function CheckHighLow(ByRef ThisPriceVol As IPriceVol, Optional ByVal IsApplyCorrection As Boolean = False) As Boolean
    With ThisPriceVol
      If .Open > .Last Then
        If .High < .Open Then
          If IsApplyCorrection Then
            .High = .Open
          Else
            Return False
          End If
        End If
        If .Low > .Last Then
          If IsApplyCorrection Then
            .Low = .Last
          Else
            Return False
          End If
        End If
      Else
        If .High < .Last Then
          If IsApplyCorrection Then
            .High = .Last
          Else
            Return False
          End If
        End If
        If .Low > .Open Then
          If IsApplyCorrection Then
            .Low = .Open
          Else
            Return False
          End If
        End If
      End If
    End With
    Return True
  End Function

  ''' <summary>
  ''' This function check that the transaction high low are consistent with the open and last
  ''' </summary>
  ''' <returns>Return True if the record is consistent</returns>
  ''' <remarks>This function can also optionally fix the issues </remarks>
  Public Function CheckPrice(Optional ByVal IsApplyCorrection As Boolean = False, Optional ByVal IsCorrectionOnVolume As Boolean = True) As Boolean
    Return CheckPrice(Me, IsApplyCorrection, IsCorrectionOnVolume)
  End Function

  ''' <summary>
  ''' If an exception occurs, the exception object will be stored here. If no exception occurs, this property is null/Nothing.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Exception() As Exception
    Get
      Return MyException
    End Get
    Set(Exception As Exception)
      MyException = Exception
    End Set
  End Property

  Public Overrides Function ToString() As String
    Return String.Format("{0},ID:{1},Key:{2},Record:{3} of {4}", TypeName(Me), Me.KeyID, Me.KeyValue.ToString)
  End Function
#End Region 'Main
#Region "IComparable"
  ''' <summary>
  ''' </summary>
  ''' <param name="other"></param>
  ''' <returns>
  ''' Less than zero: This object is less than the other parameter. 
  ''' Zero : This object is equal to other. 
  ''' Greater than zero : This object is greater than other. 
  ''' </returns>
  ''' <remarks></remarks>
  Public Function CompareTo(other As Record) As Integer Implements System.IComparable(Of Record).CompareTo
    Return Me.KeyValue.CompareTo(other.KeyValue)
  End Function
#End Region
#Region "Register Key"
  Public Property KeyID As Integer Implements IRegisterKey(Of Date).KeyID
    Get
      Return Me.ID
    End Get
    Set(value As Integer)
      Me.ID = value
    End Set
  End Property

  Public Property KeyValue As Date Implements IRegisterKey(Of Date).KeyValue
    Get
      Return Me.DateUpdate
    End Get
    Set(value As Date)

    End Set
  End Property
#End Region
#Region "IMemoryStream"
  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeSaveTo
    Me.SerializeSaveTo(Stream, IMemoryStream.enuFileType.Standard)
  End Sub

  Public Sub SerializeSaveTo(ByRef Stream As System.IO.Stream, FileType As IMemoryStream.enuFileType) Implements IMemoryStream.SerializeSaveTo
    Dim ThisBinaryWriter As New BinaryWriter(Stream)
    Dim ThisListException As List(Of Exception)
    Dim ThisException As Exception
    With ThisBinaryWriter
      Me.ToDataPosition.Current = .BaseStream.Position
      .Write(VERSION_MEMORY_STREAM)
      '.Write(Me.ToDataPosition.ToPrevious)
      .Write(Me.ID)
      ThisListException = Me.Exception.ToList
      ThisListException.Reverse()
      .Write(ThisListException.Count)
      For Each ThisException In ThisListException
        .Write(ThisException.Message)
      Next
      .Write(Me.StockID)
      .Write(Me.DateDay.ToBinary)
      .Write(Me.DateUpdate.ToBinary)
      .Write(Me.DateLastTrade.ToBinary)
      .Write(Me.Open)
      .Write(Me.Last)
      .Write(Me.High)
      .Write(Me.Low)
      .Write(Me.Vol)
      .Write(Me.FinancialHighlights.Count)
      For Each ThisFinancialHighlight In Me.FinancialHighlights
        ThisFinancialHighlight.SerializeSaveTo(Stream)
      Next
      .Write(Me.MarketQuoteDatas.Count)
      For Each ThisMarketQuoteData In Me.MarketQuoteDatas
        ThisMarketQuoteData.SerializeSaveTo(Stream)
      Next
      .Write(Me.QuoteValues.Count)
      For Each ThisQuoteValue In Me.QuoteValues
        ThisQuoteValue.SerializeSaveTo(Stream)
      Next
      .Write(Me.TradeInfoes.Count)
      For Each ThisTradeInfoe In Me.TradeInfoes
        ThisTradeInfoe.SerializeSaveTo(Stream)
      Next
      .Write(Me.ValuationMeasures.Count)
      For Each ThisValuationMeasure In Me.ValuationMeasures
        ThisValuationMeasure.SerializeSaveTo(Stream)
      Next
    End With
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream) Implements IMemoryStream.SerializeLoadFrom
    Me.SerializeLoadFrom(Stream, False)
  End Sub

  Public Sub SerializeLoadFrom(ByRef Stream As System.IO.Stream, IsRecordVirtual As Boolean) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisBinaryReader As New BinaryReader(Stream, New System.Text.UTF8Encoding(), leaveOpen:=True)
    Dim ThisVersion As Single

    With ThisBinaryReader
      Me.ToDataPosition.Current = Stream.Position
      ThisVersion = .ReadSingle
      'Me.ToDataPosition.ToPrevious = .ReadInt64
      Me.ID = .ReadInt32
      Me.Exception = Nothing
      Dim I As Integer
      For I = 1 To .ReadInt32
        MyException = New Exception(.ReadString, MyException)
      Next
      Me.StockID = .ReadInt32
      Me.DateDay = DateTime.FromBinary(.ReadInt64)
      Me.DateUpdate = DateTime.FromBinary(.ReadInt64)
      Me.DateLastTrade = DateTime.FromBinary(.ReadInt64)
      Me.Open = .ReadSingle
      Me.Last = .ReadSingle
      Me.High = .ReadSingle
      Me.Low = .ReadSingle
      Me.Vol = .ReadInt32
      For I = 1 To .ReadInt32
        Dim ThisFinancialHighlight As New FinancialHighlight(Me, Stream)
      Next
      For I = 1 To .ReadInt32
        Dim ThisMarketQuoteData As New MarketQuoteData(Me, Stream)
      Next
      For I = 1 To .ReadInt32
        Dim ThisQuoteValue As New QuoteValue(Me, Stream)
      Next
      For I = 1 To .ReadInt32
        Dim ThisTradeInfo As New TradeInfo(Me, Stream)
      Next
      For I = 1 To .ReadInt32
        Dim ThisValuationMeasure As New ValuationMeasure(Me, Stream)
      Next
    End With
    ThisBinaryReader.Dispose()
  End Sub

  Public Sub SerializeLoadFrom(ByRef Data() As Byte) Implements IMemoryStream.SerializeLoadFrom
    Dim ThisStream As Stream = New System.IO.MemoryStream(Data, writable:=True)
    Me.SerializeLoadFrom(ThisStream)
    ThisStream.Dispose()
  End Sub

  Public Function SerializeSaveTo() As Byte() Implements IMemoryStream.SerializeSaveTo
    Dim ThisStream As Stream = New System.IO.MemoryStream
    Dim ThisBinaryReader As New BinaryReader(ThisStream)
    Dim ThisData As Byte()

    Me.SerializeSaveTo(ThisStream)
    ThisData = ThisBinaryReader.ReadBytes(CInt(ThisStream.Length))
    ThisBinaryReader.Dispose()
    ThisStream.Dispose()
    Return ThisData
  End Function
#End Region
#Region "IFormatData"
  Public Function ToStingOfData() As String() Implements IFormatData.ToStingOfData
    Return Extensions.ToStingOfData(Of Record)(Me)
  End Function

  Public Function ToListOfHeader() As List(Of HeaderInfo) Implements IFormatData.ToListOfHeader
    Return MyListHeaderInfo
  End Function

  Public Shared Function ListOfHeader() As List(Of HeaderInfo)
    Dim ThisListHeaderInfo As New List(Of HeaderInfo)
    With ThisListHeaderInfo
      .Add(New HeaderInfo With {.Name = "ID", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateDay", .Title = .Name, .Format = "{0:d}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateUpdate", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "DateLastTrade", .Title = .Name, .Format = "{0:dd/MM/yyyy HH:mm:ss}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Open", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "High", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Low", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Last", .Title = .Name, .Format = "{0:0.00}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
      .Add(New HeaderInfo With {.Name = "Vol", .Title = .Name, .Format = "{0}", .Visible = True, .IsSortable = True, .IsSortEnabled = True})
    End With
    Return ThisListHeaderInfo
  End Function
#End Region
#Region "IEquatable"
  Public Function EqualsDeep(other As Record, Optional ByVal IsIgnoreID As Boolean = False) As Boolean
    With other
      If IsIgnoreID = False Then
        If .ID <> Me.ID Then Return False
        If .StockID <> Me.StockID Then Return False
      End If
      If .DateDay <> Me.DateDay Then Return False
      If .DateUpdate <> Me.DateUpdate Then Return False
      If .DateLastTrade <> Me.DateLastTrade Then Return False
      If .Open <> Me.Open Then Return False
      If .Last <> Me.Last Then Return False
      If .High <> Me.High Then Return False
      If .Low <> Me.Low Then Return False
      If .Vol <> Me.Vol Then Return False
      If Me.FinancialHighlights.EqualsDeep(.FinancialHighlights, IsIgnoreID) = False Then Return False
      If Me.MarketQuoteDatas.EqualsDeep(.MarketQuoteDatas, IsIgnoreID) = False Then Return False
      If Me.QuoteValues.EqualsDeep(.QuoteValues, IsIgnoreID) = False Then Return False
      If Me.TradeInfoes.EqualsDeep(.TradeInfoes, IsIgnoreID) = False Then Return False
      If Me.ValuationMeasures.EqualsDeep(.ValuationMeasures, IsIgnoreID) = False Then Return False
    End With
    Return True
  End Function

  Public Overloads Function Equals(other As Record) As Boolean Implements IEquatable(Of Record).Equals
    If other Is Nothing Then Return False
    If Me.DateUpdate = other.DateUpdate Then
      If Me.Stock IsNot Nothing Then
        Return Me.Stock.Equals(other.Stock)
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  Public Overrides Function Equals(obj As Object) As Boolean
    If (TypeOf obj Is Record) Then
      Return Me.Equals(DirectCast(obj, Record))
    Else
      Return (False)
    End If
  End Function

  Public Overrides Function GetHashCode() As Integer
    Return Me.DateUpdate.GetHashCode()
  End Function
#End Region
#Region "IRecordQuoteValue"
  Private Sub IRecordQuoteValue_PriceChange(Open As Single, Low As Single, High As Single, Last As Single) Implements IRecordQuoteValue.PriceChange
    Me.Open = Open
    Me.Low = Low
    Me.High = High
    Me.Last = Last
  End Sub

  Private ReadOnly Property IRecordQuoteValue_AfterHoursChangeRealtime As Single Implements IRecordQuoteValue.AfterHoursChangeRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).AfterHoursChangeRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_AnnualizedGain As Single Implements IRecordQuoteValue.AnnualizedGain
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).AnnualizedGain
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Ask As Single Implements IRecordQuoteValue.Ask
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).Ask
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_AskRealTime As Single Implements IRecordQuoteValue.AskRealTime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).AskRealTime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_AskSize As Integer Implements IRecordQuoteValue.AskSize
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).AskSize
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_AverageDailyVolume As Integer Implements IRecordQuoteValue.AverageDailyVolume
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).AverageDailyVolume
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Bid As Single Implements IRecordQuoteValue.Bid
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).Bid
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_BidRealtime As Single Implements IRecordQuoteValue.BidRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).BidRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_BidSize As Integer Implements IRecordQuoteValue.BidSize
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).BidSize
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_BookValue As Single Implements IRecordQuoteValue.BookValue
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).BookValue
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DateDay As Date Implements IRecordQuoteValue.DateDay
    Get
      Return Me.DateDay
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DateLastTrade As Date Implements IRecordQuoteValue.DateLastTrade
    Get
      Return Me.DateLastTrade
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DateUpdate As Date Implements IRecordQuoteValue.DateUpdate
    Get
      Return Me.DateUpdate
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DividendPayDate As Date Implements IRecordQuoteValue.DividendPayDate
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).DividendPayDate
      Else
        Return ReportDate.DateNullValue
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DividendShare As Single Implements IRecordQuoteValue.DividendShare
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).DividendShare
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_DividendYield As Single Implements IRecordQuoteValue.DividendYield
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).DividendYield
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_EarningsShare As Single Implements IRecordQuoteValue.EarningsShare
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).EarningsShare
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_EBITDA As Single Implements IRecordQuoteValue.EBITDA
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).EBITDA
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_EPSEstimateCurrentYear As Single Implements IRecordQuoteValue.EPSEstimateCurrentYear
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).EPSEstimateCurrentYear
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_EPSEstimateNextQuarter As Single Implements IRecordQuoteValue.EPSEstimateNextQuarter
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).EPSEstimateNextQuarter
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_EPSEstimateNextYear As Single Implements IRecordQuoteValue.EPSEstimateNextYear
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).EPSEstimateNextYear
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_ExDividendDate As Date Implements IRecordQuoteValue.ExDividendDate
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).ExDividendDate
      Else
        Return ReportDate.DateNullValue
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_FloatShares As Integer Implements IRecordQuoteValue.FloatShares
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).FloatShares
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_High As Single Implements IRecordQuoteValue.High
    Get
      Return Me.High
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_HoldingsValue As Single Implements IRecordQuoteValue.HoldingsValue
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).HoldingsValue
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_HoldingsValueRealtime As Single Implements IRecordQuoteValue.HoldingsValueRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).HoldingsValueRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_ID As Integer Implements IRecordQuoteValue.ID
    Get
      Return Me.ID
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Last As Single Implements IRecordQuoteValue.Last
    Get
      Return Me.Last
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_LastTradeDate As Date Implements IRecordQuoteValue.LastTradeDate
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).LastTradeDate
      Else
        Return ReportDate.DateNullValue
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_LastTradePriceOnly As Single Implements IRecordQuoteValue.LastTradePriceOnly
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).LastTradePriceOnly
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_LastTradeRealtimeWithTime As Date Implements IRecordQuoteValue.LastTradeRealtimeWithTime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).LastTradeRealtimeWithTime
      Else
        Return ReportDate.DateNullValue
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_LastTradeSize As Integer Implements IRecordQuoteValue.LastTradeSize
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).LastTradeSize
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Low As Single Implements IRecordQuoteValue.Low
    Get
      Return Me.Low
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_MarketCapitalization As Single Implements IRecordQuoteValue.MarketCapitalization
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).MarketCapitalization
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_MarketCapRealtime As Single Implements IRecordQuoteValue.MarketCapRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).MarketCapRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_OneyrTargetPrice As Single Implements IRecordQuoteValue.OneyrTargetPrice
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).OneyrTargetPrice
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Open As Single Implements IRecordQuoteValue.Open
    Get
      Return Me.Open
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_OrderBookRealtime As Single Implements IRecordQuoteValue.OrderBookRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).OrderBookRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PEGRatio As Single Implements IRecordQuoteValue.PEGRatio
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PEGRatio
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PERatio As Single Implements IRecordQuoteValue.PERatio
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PERatio
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PERatioRealtime As Single Implements IRecordQuoteValue.PERatioRealtime
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PERatioRealtime
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PriceBook As Single Implements IRecordQuoteValue.PriceBook
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PriceBook
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PriceEPSEstimateCurrentYear As Single Implements IRecordQuoteValue.PriceEPSEstimateCurrentYear
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PriceEPSEstimateCurrentYear
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PriceEPSEstimateNextYear As Single Implements IRecordQuoteValue.PriceEPSEstimateNextYear
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PriceEPSEstimateNextYear
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_PriceSales As Single Implements IRecordQuoteValue.PriceSales
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).PriceSales
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_SharesOwned As Integer Implements IRecordQuoteValue.SharesOwned
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).SharesOwned
      Else
        Return 0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_ShortRatio As Single Implements IRecordQuoteValue.ShortRatio
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).ShortRatio
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_StockExchange As String Implements IRecordQuoteValue.StockExchange
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).StockExchange
      Else
        Return ""
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_TickerTrend As Single Implements IRecordQuoteValue.TickerTrend
    Get
      If Me.QuoteValues.Count > 0 Then
        Return Me.QuoteValues(0).TickerTrend
      Else
        Return 0.0
      End If
    End Get
  End Property

  Private ReadOnly Property IRecordQuoteValue_Vol As Integer Implements IRecordQuoteValue.Vol
    Get
      Return Me.Vol
    End Get
  End Property

  Public Sub RemoveBookEarningGlitch(RecordQuoteValueLast As IRecordQuoteValue) Implements IRecordQuoteValue.RemoveBookEarningGlitch
    If Me.QuoteValues.Count > 0 Then
      With Me.QuoteValues(0)
        .BookValue = RecordQuoteValueLast.BookValue
        .EarningsShare = RecordQuoteValueLast.EarningsShare
        .EBITDA = RecordQuoteValueLast.EBITDA
        .DividendPayDate = RecordQuoteValueLast.DividendPayDate
        .ExDividendDate = RecordQuoteValueLast.ExDividendDate
        .DividendShare = RecordQuoteValueLast.DividendShare
        .DividendYield = RecordQuoteValueLast.DividendYield
        .EPSEstimateCurrentYear = RecordQuoteValueLast.EPSEstimateCurrentYear
        .EPSEstimateNextQuarter = RecordQuoteValueLast.EPSEstimateNextQuarter
        .EPSEstimateNextYear = RecordQuoteValueLast.EPSEstimateNextYear
        .FloatShares = RecordQuoteValueLast.FloatShares
        .HoldingsValue = RecordQuoteValueLast.HoldingsValue
        .HoldingsValueRealtime = RecordQuoteValueLast.HoldingsValueRealtime
        .MarketCapitalization = RecordQuoteValueLast.MarketCapitalization
        .MarketCapRealtime = RecordQuoteValueLast.MarketCapRealtime
        .OneyrTargetPrice = RecordQuoteValueLast.OneyrTargetPrice
        .OrderBookRealtime = RecordQuoteValueLast.OrderBookRealtime
        .PEGRatio = RecordQuoteValueLast.PEGRatio
        .PERatio = RecordQuoteValueLast.PERatio
        .PERatioRealtime = RecordQuoteValueLast.PERatioRealtime
        .PriceBook = RecordQuoteValueLast.PriceBook
        .PriceEPSEstimateCurrentYear = RecordQuoteValueLast.PriceEPSEstimateCurrentYear
        .PriceEPSEstimateNextYear = RecordQuoteValueLast.PriceEPSEstimateNextYear
        .PriceSales = RecordQuoteValueLast.PriceSales
        .SharesOwned = RecordQuoteValueLast.SharesOwned
        .ShortRatio = RecordQuoteValueLast.ShortRatio
      End With
    End If
  End Sub

  Public Function AsIRecordType() As IRecordType Implements IRecordType.AsIRecordType
    Return Me
  End Function

  'Private ReadOnly Property IRecordQuoteValue_FiscalYearEnd As Date Implements IRecordQuoteValue.FiscalYearEnd
  '  Get

  '  End Get
  'End Property

  'Private ReadOnly Property IRecordQuoteValue_PriceEstimateCurrentYear As Single Implements IRecordQuoteValue.PriceEstimateCurrentYear
  '  Get

  '  End Get
  'End Property

  'Private ReadOnly Property IRecordQuoteValue_PriceEstimateFromPEG As Single Implements IRecordQuoteValue.PriceEstimateFromPEG
  '  Get

  '  End Get
  'End Property

  'Private ReadOnly Property IRecordQuoteValue_PriceEstimateLastYear As Single Implements IRecordQuoteValue.PriceEstimateLastYear
  '  Get

  '  End Get
  'End Property

  'Private ReadOnly Property IRecordQuoteValue_PriceEstimateNextQuarter As Single Implements IRecordQuoteValue.PriceEstimateNextQuarter
  '  Get

  '  End Get
  'End Property

  'Private ReadOnly Property IRecordQuoteValue_PriceEstimateNextYear As Single Implements IRecordQuoteValue.PriceEstimateNextYear
  '  Get

  '  End Get
  'End Property
#End Region
#Region "IDataPosition"
  Private Property IDataPosition_Current As Long Implements IDataPosition.Current
  Private Property IDataPosition_ToNext As Long Implements IDataPosition.ToNext
  Private Property IDataPosition_ToPrevious As Long Implements IDataPosition.ToPrevious
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
      Return Me.DateUpdate
    End Get
  End Property

  Public ReadOnly Property IDateUpdate_DateDay As Date Implements IDateUpdate.DateDay
    Get
      Return Me.DateDay
    End Get
  End Property
#End Region   'IDateUpdate
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
#End Region    'IDateTrade
#Region "IPriceVol"
  Public ReadOnly Property AsIPriceVol As IPriceVol Implements IPriceVol.AsIPriceVol
    Get
      Return Me
    End Get
  End Property

  Private Property IPriceVol_High As Single Implements IPriceVol.High
    Get
      Return Me.High
    End Get
    Set(value As Single)
      Me.High = value
    End Set
  End Property

  Private Property IPriceVol_Last As Single Implements IPriceVol.Last
    Get
      Return Me.Last
    End Get
    Set(value As Single)
      Me.Last = value
    End Set
  End Property

  Public Property IPriceVol_LastWeighted As Single Implements IPriceVol.LastWeighted
    Get
      Return RecordPrices.CalculateLastWeighted(Me)
    End Get
    Set(value As Single)

    End Set
  End Property

  Private Property IPriceVol_Low As Single Implements IPriceVol.Low
    Get
      Return Me.Low
    End Get
    Set(value As Single)
      Me.Low = value
    End Set
  End Property

  Private Property IPriceVol_Open As Single Implements IPriceVol.Open
    Get
      Return Me.Open
    End Get
    Set(value As Single)
      Me.Open = value
    End Set
  End Property

  Private Property IPriceVol_OpenNext As Single Implements IPriceVol.OpenNext
    Get
      Return Me.Last
    End Get
    Set(value As Single)

    End Set
  End Property

  Private Property IPriceVol_Vol As Integer Implements IPriceVol.Vol
    Get
      Return Me.Vol
    End Get
    Set(value As Integer)
      Me.Vol = value
    End Set
  End Property

  Private Property IPriceVol_DateDay As Date Implements IPriceVol.DateDay
    Get
      Return Me.DateDay
    End Get
    Set(value As Date)
      Me.DateDay = value
    End Set
  End Property

  Private Property IPriceVol_DateUpdate As Date Implements IPriceVol.DateUpdate
    Get
      Return Me.DateUpdate
    End Get
    Set(value As Date)
      Me.DateUpdate = value
    End Set
  End Property

  ''' <summary>
  ''' It is fixed to 1 at this time for the record
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Property IPriceVol_LastAdjusted As Single Implements IPriceVol.LastAdjusted
    Get

      Return 1.0
    End Get
    Set(value As Single)

    End Set
  End Property

  Private Property IPriceVol_Range As Single Implements IPriceVol.Range
    Get
      Return MyPriceRange
    End Get
    Set(value As Single)
      MyPriceRange = value
    End Set
  End Property

  Private Property IPriceVol_LastPrevious As Single Implements IPriceVol.LastPrevious
    Get
      Return MyPriceLastPrevious
    End Get
    Set(value As Single)
      MyPriceLastPrevious = value
    End Set
  End Property

  Public Property IsIntraDay As Boolean Implements IPriceVol.IsIntraDay
  Public Property VolMinus As Integer Implements IPriceVol.VolMinus
  Public Property VolPlus As Integer Implements IPriceVol.VolPlus
  Public Property IsSpecialDividendPayout As Boolean Implements IPriceVol.IsSpecialDividendPayout
  Public Property SpecialDividendPayoutValue As Single Implements IPriceVol.SpecialDividendPayoutValue

  Public Property RecordType As IRecordType.enuRecordType Implements IRecordType.RecordType
    Get
      Return _RecordType
    End Get
    Set(value As IRecordType.enuRecordType)
      _RecordType = value
    End Set
  End Property

  Private _IsRecordChanged As Boolean
  Public Property IsRecordChanged As Boolean Implements IRecordType.IsRecordChanged
    Get
      Return _IsRecordChanged
    End Get
    Set(value As Boolean)
      _IsRecordChanged = value
    End Set
  End Property
#End Region
End Class    'Record
#End Region

#Region "RecordQuoteValue"
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

  Private MyRecord As Record
  Private MyRecordQuoteValue As IRecordQuoteValue
  'Private Shared MyCompareByName As CompareByName(Of RecordQuoteValue)

  Public Sub New()
    Me.New(New Record)
  End Sub
  Public Sub New(ByVal RecordQuoteValue As Record)
    MyRecord = RecordQuoteValue
    MyRecordQuoteValue = MyRecord
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

  'Public ReadOnly Property FiscalYearEnd As Date Implements IRecordQuoteValue.FiscalYearEnd
  '  Get

  '  End Get
  'End Property

  'Public ReadOnly Property PriceEstimateCurrentYear As Single Implements IRecordQuoteValue.PriceEstimateCurrentYear
  '  Get

  '  End Get
  'End Property

  'Public ReadOnly Property PriceEstimateFromPEG As Single Implements IRecordQuoteValue.PriceEstimateFromPEG
  '  Get

  '  End Get
  'End Property

  'Public ReadOnly Property PriceEstimateLastYear As Single Implements IRecordQuoteValue.PriceEstimateLastYear
  '  Get

  '  End Get
  'End Property

  'Public ReadOnly Property PriceEstimateNextQuarter As Single Implements IRecordQuoteValue.PriceEstimateNextQuarter
  '  Get

  '  End Get
  'End Property

  'Public ReadOnly Property PriceEstimateNextYear As Single Implements IRecordQuoteValue.PriceEstimateNextYear
  '  Get

  '  End Get
  'End Property
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
#End Region
End Class
#End Region
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
#Region "Class EqualityComparerOfRecord"
<Serializable()>
Friend Class EqualityComparerOfRecord
  Implements IEqualityComparer(Of Record)

  Public Overloads Function Equals(x As Record, y As Record) As Boolean Implements IEqualityComparer(Of Record).Equals
    If (x Is Nothing) And (y Is Nothing) Then
      Return True
    ElseIf (x Is Nothing) Xor (y Is Nothing) Then
      Return False
    Else
      If x.DateUpdate = y.DateUpdate Then
        Return x.Stock.Equals(y.Stock)
      Else
        Return False
      End If
    End If
  End Function

  Public Overloads Function GetHashCode(obj As Record) As Integer Implements IEqualityComparer(Of Record).GetHashCode
    If obj IsNot Nothing Then
      Return obj.DateUpdate.GetHashCode
    Else
      Return obj.GetHashCode
    End If
  End Function
End Class

Friend Class EqualityComparerOfRecordQuoteValue
  Implements IEqualityComparer(Of RecordQuoteValue)

  Public Overloads Function Equals(x As RecordQuoteValue, y As RecordQuoteValue) As Boolean Implements IEqualityComparer(Of RecordQuoteValue).Equals
    If (x Is Nothing) And (y Is Nothing) Then
      Return True
    ElseIf (x Is Nothing) Xor (y Is Nothing) Then
      Return False
    Else
      If x.DateUpdate = y.DateUpdate Then
        Return x.Record.Stock.Equals(y.Record.Stock)
      Else
        Return False
      End If
    End If
  End Function

  Public Overloads Function GetHashCode(obj As RecordQuoteValue) As Integer Implements IEqualityComparer(Of RecordQuoteValue).GetHashCode
    If obj IsNot Nothing Then
      Return obj.DateUpdate.GetHashCode
    Else
      Return obj.GetHashCode
    End If
  End Function
End Class
#End Region

Friend Class FilterHoldFromZero
  Private MyFilterValueLast As Single
  Private MyFilterCount As Integer
  Private MyFilterRate As Integer

  Public Sub New(ByVal FilterRate As Integer)
    MyFilterRate = FilterRate
  End Sub

  Public Function Filter(ByVal Value As Single) As Single
    If Value = 0 Then
      MyFilterCount = MyFilterCount + 1
      If MyFilterCount <= MyFilterRate Then
        Value = MyFilterValueLast
      End If
    Else
      MyFilterValueLast = Value
      MyFilterCount = 0
    End If
    Return Value
  End Function
End Class

#Region "template"
'------------------------------------------------------------------------------
' <auto-generated>
'    This code was generated from a template.
'
'    Manual changes to this file may cause unexpected behavior in your application.
'    Manual changes to this file will be overwritten if the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------
Partial Public Class Record
  Public Property ID As Integer
  Public Property StockID As Integer
  Public Property DateDay As Date
  Public Property DateUpdate As Date
  Public Property DateLastTrade As Date
  Public Property Open As Single
  Public Property Last As Single
  Public Property High As Single
  Public Property Low As Single
  Public Property Vol As Integer

  Public Overridable Property FinancialHighlights As ICollection(Of FinancialHighlight) = New HashSet(Of FinancialHighlight)
  Public Overridable Property MarketQuoteDatas As ICollection(Of MarketQuoteData) = New HashSet(Of MarketQuoteData)
  Public Overridable Property QuoteValues As ICollection(Of QuoteValue) = New HashSet(Of QuoteValue)
  Public Overridable Property TradeInfoes As ICollection(Of TradeInfo) = New HashSet(Of TradeInfo)
  Public Overridable Property ValuationMeasures As ICollection(Of ValuationMeasure) = New HashSet(Of ValuationMeasure)
  Public Overridable Property Stock As Stock
End Class
#End Region

#Region "IRecordType"
'provide some information on the type of Record 
Public Interface IRecordType
  Enum enuRecordType
    EndOfDay
    LiveUpdate
  End Enum

  Function AsIRecordType() As IRecordType
  Property RecordType As enuRecordType
  Property IsRecordChanged As Boolean
End Interface
#End Region