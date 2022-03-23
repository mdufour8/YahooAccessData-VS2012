Imports YahooAccessData.ExtensionService

Public Class ProcessIndice
  Implements IProcessIndice

  Private MyStock As YahooAccessData.Stock
  Private MyPriceVol() As YahooAccessData.PriceVol
  Private MyCount As Integer
  Private MyWeightSum As Double
  Private MyStockIndiceLast As IStockIndice

  Public Sub New(ByVal Stock As YahooAccessData.Stock)
    MyStock = Stock
    MyStock.DateStart = Date.MaxValue
    MyStock.DateStop = Date.MinValue
    Me.ScaleRatio = 1.0
    Me.ScaleOffset = 0
    ReDim MyPriceVol(-1)
  End Sub

  ReadOnly Property Stock As YahooAccessData.Stock Implements IProcessIndice.Stock
    Get
      Return MyStock
    End Get
  End Property

  Public Sub Add(ByVal StockIndice As YahooAccessData.IStockIndice, ByVal RecordPrices As YahooAccessData.RecordPrices) Implements IProcessIndice.Add
    Dim ThisWeight As Single
    Dim ThisPriceVol As YahooAccessData.PriceVol
    Dim I As Integer

    If MyStock.Records.Count > 0 Then
      Throw New InvalidOperationException
    End If

    If MyPriceVol.Length = 0 Then
      ReDim MyPriceVol(0 To RecordPrices.NumberPoint - 1)
    End If
    If MyPriceVol.Length <> RecordPrices.NumberPoint Then
      Throw New IndexOutOfRangeException
    End If

    'adjust the date if needed
    If RecordPrices.DateStart < MyStock.DateStart Then
      MyStock.DateStart = RecordPrices.DateStart
    End If
    If RecordPrices.DateStop > MyStock.DateStop Then
      MyStock.DateStop = RecordPrices.DateStop
    End If
    ThisWeight = CSng(StockIndice.IndiceWeight)
    If MyCount = 0 Then
      'this is the first set of data
      For I = 0 To RecordPrices.NumberPoint - 1
        MyPriceVol(I) = RecordPrices.PriceVols(I)
        MyPriceVol(I).MultiPly(ThisWeight)
      Next
      MyWeightSum = ThisWeight
    Else
      For I = 0 To RecordPrices.NumberPoint - 1
        ThisPriceVol = RecordPrices.PriceVols(I)
        ThisPriceVol.MultiPly(ThisWeight)
        MyPriceVol(I).Add(ThisPriceVol)
        'If MyPriceVol(I).DividendYield > 50 Then
        '  'Debugger.Break()
        '  I = I
        'End If
      Next
      MyWeightSum = MyWeightSum + ThisWeight
    End If
    MyStockIndiceLast = StockIndice
    MyCount = MyCount + 1
  End Sub

  ''' <summary>
  ''' Finalize the process by adding the result to the stock record
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Refresh() Implements IProcessIndice.Refresh
    Dim I As Integer

    For I = 0 To MyPriceVol.Length - 1
      MyPriceVol(I).MultiPly(CSng(Me.ScaleRatio))
      MyPriceVol(I).Add(New PriceVol(CSng(Me.ScaleOffset)))
    Next
    If MyStock.Records.Count = 0 Then
      For I = 0 To MyPriceVol.Length - 1
        MyStock.Records.Add(MyPriceVol(I).CopyFromAsRecord)
      Next
    End If
  End Sub

  Public ReadOnly Property Count As Integer Implements IProcessIndice.Count
    Get
      Return MyCount
    End Get
  End Property

  Public Property ScaleOffset As Double Implements IProcessIndice.ScaleOffset

  Public Property ScaleRatio As Double Implements IProcessIndice.ScaleRatio
End Class

Public Interface IProcessIndice
  Sub Refresh()
  Sub Add(ByVal StockIndice As YahooAccessData.IStockIndice, ByVal RecordPrices As YahooAccessData.RecordPrices)
  ReadOnly Property Count As Integer
  ReadOnly Property Stock As YahooAccessData.Stock
  Property ScaleRatio As Double
  Property ScaleOffset As Double
End Interface


<Serializable>
Public Class StockIndice
  Implements IStockIndice

  Private MySymbol As String
  Private MyStock As Stock
  Private MyRecordPrices As RecordPrices


  Public Sub New(ByVal Symbol As String)
    Me.New(Symbol, 0)
  End Sub

  Public Sub New(ByVal Symbol As String, ByVal IndiceWeight As Double)
    MySymbol = Symbol
    Me.IndiceWeight = IndiceWeight
    MyStock = Nothing
    MyRecordPrices = Nothing
    Me.StartDate = Date.MinValue
    Me.StopDate = Date.MaxValue
  End Sub

  Public Property IndiceWeight As Double Implements IStockIndice.IndiceWeight
  Public Property Name As String Implements IStockIndice.Name

  Public ReadOnly Property Symbol As String Implements IStockIndice.Symbol
    Get
      Return MySymbol
    End Get
  End Property

  Public ReadOnly Property AsIStockIndice As IStockIndice Implements IStockIndice.AsIStockIndice
    Get
      Return Me
    End Get
  End Property

  Public Overrides Function ToString() As String Implements IStockIndice.ToString
    Return String.Format("{0}: {1:n4}", Me.Symbol, Me.IndiceWeight)
  End Function
  Public Property IsScalingEnabled As Boolean Implements IStockIndice.IsScalingEnabled
  Public Property IndiceAutoScaleWeight As Double Implements IStockIndice.IndiceAutoScaleWeight
  Public Property Process As IProcessIndice Implements IStockIndice.Process

  Public Property StartDate As Date Implements IStockIndice.StartDate

  Public Property StopDate As Date Implements IStockIndice.StopDate

  Public Function IsDateRangeValid() As Boolean Implements IStockIndice.IsDateRangeValid
    If Me.StopDate < Me.StartDate Then
      Return False
    End If
    If Me.StartDate > Me.RecordPrices.DateStop Then
      Return False
    End If
    Return True
  End Function

  Public Sub SetStock(Stock As Stock) Implements IStockIndice.SetStock
    MyStock = Stock
    If MyStock IsNot Nothing Then
      With MyStock
        MyRecordPrices = .RecordQuoteValues(YahooAccessData.Report.enuTimeFormat.Daily).ToDailyRecordPrices(.Report.DateStart, .Report.DateStop)
      End With
    Else
      MyRecordPrices = Nothing
    End If
  End Sub

  Public ReadOnly Property Stock As Stock Implements IStockIndice.Stock
    Get
      Return MyStock
    End Get
  End Property

  Public ReadOnly Property RecordPrices As RecordPrices Implements IStockIndice.RecordPrices
    Get
      Return MyRecordPrices
    End Get
  End Property
End Class

Public Interface IStockIndice
  ReadOnly Property AsIStockIndice As IStockIndice
  Property Name As String
  Property StartDate As Date
  Property StopDate As Date
  Function IsDateRangeValid() As Boolean
  Property IndiceWeight As Double
  Property IndiceAutoScaleWeight As Double
  Property IsScalingEnabled As Boolean
  ReadOnly Property Symbol As String

  Sub SetStock(ByVal Stock As Stock)
  ReadOnly Property Stock As YahooAccessData.Stock
  ReadOnly Property RecordPrices As YahooAccessData.RecordPrices
  Property Process As IProcessIndice
  Function ToString() As String
End Interface