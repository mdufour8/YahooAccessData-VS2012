Imports System
Imports System.Linq
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices
Imports YahooAccessData.ExtensionService

Partial Public Class StockYahooEntities

Private MyException As Exception

Public Event ErrorMessage(ByVal Message As String)

Public Sub New(ByVal Connection As String)
	MyBase.New(Connection)
End Sub

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

''' <summary>
''' Save a new report to the database
''' </summary>
''' <param name="ThisReport"></param>
''' <returns>
''' True if it succeed
''' </returns>
''' <remarks>
''' the function report an error if a report with the same 
''' name and date exist in the database
'''  </remarks>
Public Function SaveNew(ByRef ThisReport As Report) As Boolean
''get the database report
''save this update in the database under a new report
''check if it has already been saved
'For Each Report In Me.Reports.Items(ThisReport)
'	If Report.Equals(ThisReport) Then
'		Me.Exception = New Exception("Yahoo download report already exist in the database...")
'		ThisReport.Exception = Me.Exception
'		Return False
'	End If
'Next
''check that we do not have any double stock entries
'For Each ThisStock In ThisReport.Stocks
'	Dim ThisSymbol = ThisStock.Symbol
'	Dim ThisStockList = ThisReport.Stocks.Where(Function(ThisR As YahooAccessData.Stock) ThisR.Symbol = ThisSymbol)
'	Debug.Print(ThisStockList.Count.ToString)
'	If ThisStockList.Count <> 1 Then
'		Debug.Print(String.Format("Something is wrong with {0} of {1)", ThisStockList.Count, ThisSymbol))
'	End If
'Next
'load the new report in the database
'add a database report
Me.Reports.Add(ThisReport)
'this is not needed
'For Each ThisStock In ThisReport.Stocks
'	I = I + 1
'	Me.Stocks.Add(ThisStock)
'	Me.Sectors.Add(ThisStock.Sector)
'	Me.Industries.Add(ThisStock.Industry)
'	For Each ThisSplitFactor In ThisStock.SplitFactors
'		Me.SplitFactors.Add(ThisSplitFactor)
'	Next
'	For Each ThisRecord In ThisStock.Records
'		Me.Records.Add(ThisRecord)
'		For Each ThisFinancialHighLight In ThisRecord.FinancialHighlights
'			Me.FinancialHighlights.Add(ThisFinancialHighLight)
'		Next
'		For Each ThisMarketQuoteData In ThisRecord.MarketQuoteDatas
'			Me.MarketQuoteDatas.Add(ThisMarketQuoteData)
'		Next
'		For Each ThisQuoteValue In ThisRecord.QuoteValues
'			Me.QuoteValues.Add(ThisQuoteValue)
'		Next
'		For Each ThisTradeInfo In ThisRecord.TradeInfoes
'			Me.TradeInfoes.Add(ThisTradeInfo)
'		Next
'		For Each ThisValuationMeasure In ThisRecord.ValuationMeasures
'			Me.ValuationMeasures.Add(ThisValuationMeasure)
'		Next
'	Next
	Try
		Me.SaveChanges()
	Catch e As Exception
		Me.Exception = e
		ThisReport.Exception = e
		RaiseEvent ErrorMessage(e.Message)
		Return False
	End Try
'Next
Return True
End Function
End Class

