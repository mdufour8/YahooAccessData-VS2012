Imports System
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports YahooAccessData.ExtensionService
Imports System.Text
Imports System.IO
Imports System.Configuration
Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports System.Linq
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class StockDataBasicTest
  'Inherits Attribute

  '<TestMethod()>
  'Public Sub TestBasic()
  'Dim dbStockEntity As YahooAccessData.StockYahooEntities
  'Dim ThisReport As YahooAccessData.Report

  'Dim Connections As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings

  'For Each Connection As ConnectionStringSettings In Connections
  '  Debug.Print("AppConfig ConnectionString:")
  '  Debug.Print(Connection.ConnectionString)
  'Next
  'Debug.Print("")
  'Debug.Print("AppConfig connection string for StockYahooEntities:")
  'Debug.Print(Connections("StockYahooEntities").ConnectionString)
  'Debug.Print("")

  'dbStockEntity = New YahooAccessData.StockYahooEntities(Connections("StockYahooEntities").ConnectionString)
  'With dbStockEntity
  '  With .Database.Connection
  '    Debug.Print("YahooAccessData.StockYahooEntities ConnectionString:")
  '    Debug.Print("Connection String: " & .ConnectionString)
  '    Debug.Print("Connection state: " & .State.ToString)
  '    .Open()
  '  End With
  '  Try
  '    Debug.Print(String.Format("Database exist:{0}", .Database.Exists()))
  '  Catch ex As Exception
  '    MsgBox(ex.Message)
  '  End Try

  '  ThisReport = .Reports.Create

  '  With ThisReport
  '    .DateStart = Now
  '    .DateStop = .DateStart
  '    .Name = "Test"
  '  End With
  '  .Reports.Add(ThisReport)
  '  .SaveChanges()
  '  'test is the data is saved in the database
  '  'Assert.IsTrue(.Reports.Count = 1)
  '  For Each ThisReport In .Reports
  '    With ThisReport
  '      Debug.Print(.ID.ToString & ":" & .Name)
  '    End With
  '  Next

  '  'can be written this way
  '  Dim Result = .Reports.
  '                Where(Function(ThisR As YahooAccessData.Report) ThisR.Name = "Test").
  '                Select(Function(ThisR) ThisR)

  '  Assert.IsTrue(Result.Count = 1)
  '  'or written with 
  '  Result = From R As YahooAccessData.Report In .Reports
  '           Where (R.Name = "Test")
  '           Select R

  '  Assert.IsTrue(Result.Count = 1)
  '  For Each ThisReport In Result
  '    With ThisReport
  '      Debug.Print(.ID.ToString & ":" & .Name)
  '    End With
  '    .Reports.Remove(ThisReport)
  '  Next
  '  .SaveChanges()

  '  Result = From R In .Reports
  '            Where (R.Name = "Test")
  '            Select R
  '  Assert.IsTrue(Result.Count = 0)
  'End With
  'End Sub

  Public Sub TestBasicOperation()
  Dim ThisReport As YahooAccessData.Report

    ThisReport = New YahooAccessData.Report
    With ThisReport
      .DateStart = Now
      .DateStop = .DateStart
      .Name = "Test"
      .StockAdd("Stock1", "Sector1", "Industry1")
      .StockAdd("Stock2", "Sector1", "Industry1")
      Assert.IsTrue(.Stocks.Count = 2)
      Assert.IsTrue(.Sectors.Count = 1)
      Assert.IsTrue(.Industries.Count = 1)

    End With
  End Sub

  '<TestMethod()>
  Public Sub TestReadSpeed()
    Dim dbStockEntity As YahooAccessData.StockYahooEntities
    Dim ThisReport1 As YahooAccessData.Report
    Dim ThisReport2 As YahooAccessData.Report
    Dim ThisReport3 As YahooAccessData.Report
    Dim ThisReport4 As YahooAccessData.Report
    Dim ThisReportWork As New YahooAccessData.Report
    Dim ThisSector As String = "Financial"
    'Dim ThisSector As String = "Indices"
    'Dim ThisSector As String = "Basic Materials"
    Dim ThisFile1 As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\" & ThisSector & "_ThreadedRead.rep"
    Dim ThisFile2 As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\" & ThisSector & "_UnThreadedRead.rep"
    ThisFile1 = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\Report_Financial_New.rep"
    ThisFile2 = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\Report_Financial.rep"

    Trace.WriteLine(String.Format("Start reading the last database sector {0}", ThisSector))
    dbStockEntity = New YahooAccessData.StockYahooEntities
    ThisReport1 = ReadLastFromDatabase(dbStockEntity, IsProxy:=False, IsThreaded:=True, SectorData:=ThisSector)
    Trace.WriteLine(String.Format("Saving file to {0}", ThisFile1))
    ThisReport1.FileSave(ThisFile1)
    ThisReport2 = ThisReport1.CopyDeep
    Trace.WriteLine("Checking copy deep equality...")
    If ThisReport1.EqualsDeep(ThisReport2) Then
      Trace.WriteLine("Copy deep equality is succesful...")
    Else
      Trace.WriteLine("Copy deep equality failed...")
    End If
    ThisReport3 = New YahooAccessData.Report
    Trace.WriteLine(String.Format("Loading file {0}", ThisFile1))
    ThisReport3 = ThisReport3.FileLoad(ThisFile1)
    Trace.WriteLine("Checking copy deep equality with loaded file...")
    If ThisReport1.EqualsDeep(ThisReport3, IsIgnoreID:=True) Then
      Trace.WriteLine("Copy deep equality with loaded file is succesful...")
    Else
      Trace.WriteLine("Copy deep equality with loaded file failed...")
    End If
    ThisReport4 = New YahooAccessData.Report
    Trace.WriteLine("Checking new file with older file version...")
    ThisReport4 = ThisReport4.FileLoad(ThisFile2)
    If ThisReport3.EqualsDeep(ThisReport4, IsIgnoreID:=True) Then
      Trace.WriteLine("Equality successful with new and old file version...")
    Else
      Trace.WriteLine("Equality with new and old file version failed...")
    End If
  End Sub

  Private Function ReadLastFromDatabase(
    ByVal ThisEntity As YahooAccessData.StockYahooEntities,
    ByVal IsProxy As Boolean,
    Optional ByVal IsThreaded As Boolean = True,
    Optional ByVal SectorData As String = "") As YahooAccessData.Report

    Dim ThisReport As YahooAccessData.Report
    Dim ThisReportCopy As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisStream As Stream = New MemoryStream

    With ThisEntity
      With .Configuration
        '.ValidateOnSaveEnabled = False
        '.AutoDetectChangesEnabled = False
        'LazyLoadingEnabled need to be true to load all the proxies
        .LazyLoadingEnabled = True
        .ProxyCreationEnabled = True
      End With
      ThisStopWatch.Restart()
      If SectorData <> "" Then
        ThisReport = .Reports.ItemTakeLasts(.Sectors, "YahooDownload", SectorData, 1).LastOrDefault()
      Else
        ThisReport = .Reports.ItemTakeLasts("YahooDownload", 1).LastOrDefault
      End If
      Trace.TraceInformation(String.Format("Time to select proxy for last sector data is (ms): {0}", ThisStopWatch.ElapsedMilliseconds))
      If ThisReport Is Nothing Then
        ThisReportCopy = ThisReport
      Else
        If IsProxy Then
          'keep the proxy database reference
          ThisReportCopy = ThisReport
        Else
          'ThisReport.SerializeTo(ThisStream)
          'ThisReportCopy = YahooAccessData.ReportFrom.Load(ThisStream)
          'If ThisReportCopy.EqualsDeep(ThisReport) Then
          '	ThisReport = ThisReport
          'Else
          '	ThisReport = ThisReport
          'End If
          ThisReportCopy = ThisReport.CopyDeep(IsThreaded)
          Trace.TraceInformation(String.Format("Time to select and copydeep last sector data is (ms): {0}", ThisStopWatch.ElapsedMilliseconds))
        End If
      End If
    End With
    Return ThisReportCopy
  End Function

''' <summary>
''' http://www.rightline.net/calendar/market-holidays.html
''' </summary>
''' <remarks></remarks>
'<TestMethod()>
  Public Sub TestTradingDate()
    Dim ThisReport As New YahooAccessData.Report

    'test new year for 2011 and 2012
    Dim ThisDate As Date = DateSerial(2011, 1, 1)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #12/31/2010 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #1/3/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 1, 1)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #12/30/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #1/3/2012 9:45:00 AM#)

    'test Martin Luther King for 2011 and 2012
    ThisDate = DateSerial(2011, 1, 17)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #1/14/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #1/18/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 1, 16)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #1/13/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #1/17/2012 9:45:00 AM#)

    'Washington's Birthday (Presidents' Day)
    ThisDate = DateSerial(2011, 2, 21)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #2/18/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #2/22/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 2, 20)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #2/17/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #2/21/2012 9:45:00 AM#)

    'Good Friday
    ThisDate = DateSerial(2011, 4, 22)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #4/21/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #4/25/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 4, 6)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #4/5/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #4/9/2012 9:45:00 AM#)

    'Memorial Day
    ThisDate = DateSerial(2011, 5, 30)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #5/27/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #5/31/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 5, 28)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #5/25/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #5/29/2012 9:45:00 AM#)

    'Independence Day
    ThisDate = DateSerial(2011, 7, 4)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #7/1/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #7/5/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 7, 4)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #7/3/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #7/5/2012 9:45:00 AM#)

    'Labor day
    ThisDate = DateSerial(2011, 9, 5)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #9/2/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #9/6/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 9, 3)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #8/31/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #9/4/2012 9:45:00 AM#)

    'Thanksgiving Day
    ThisDate = DateSerial(2011, 11, 24)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #11/23/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #11/25/2011 9:45:00 AM#)
    'the market should be closing early
    ThisDate = DateSerial(2011, 11, 25).AddHours(14)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #11/25/2011 1:15:00 PM#)

    ThisDate = DateSerial(2012, 11, 22)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #11/21/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #11/23/2012 9:45:00 AM#)
    'the market should be closing early
    ThisDate = DateSerial(2012, 11, 23).AddHours(14)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #11/23/2012 1:15:00 PM#)

    'Christmas(Day)
    ThisDate = DateSerial(2011, 12, 25)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #12/23/2011 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #12/27/2011 9:45:00 AM#)

    ThisDate = DateSerial(2012, 12, 25)
    Assert.IsTrue(ThisReport.DayOfTrade(ThisDate) = #12/24/2012 4:15:00 PM#)
    Assert.IsTrue(ThisReport.DayOfTradeNext(ThisDate) = #12/26/2012 9:45:00 AM#)

    'test for special holiday Sandy storm
    'test beginning of day
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 9:00:00 AM#) = #10/25/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 9:30:00 AM#) = #10/25/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 9:44:59 AM#) = #10/25/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 9:45:00 AM#) = #10/26/2012 9:45:00 AM#)
    'test closing of day
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 4:00:00 PM#) = #10/26/2012 4:00:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 4:14:59 PM#) = #10/26/2012 4:14:59 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 4:15:01 PM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/26/2012 4:30:00 PM#) = #10/26/2012 4:15:00 PM#)

    'test of special holiday require extensive testing
    'start with the first day of the storm Sandy
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 9:00:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 9:30:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 9:44:59 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 9:45:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 4:00:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/29/2012 4:45:00 AM#) = #10/26/2012 4:15:00 PM#)
    'continue with the second day of the storm Sandy
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 9:00:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 9:30:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 9:44:59 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 9:45:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 4:00:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/30/2012 4:45:00 AM#) = #10/26/2012 4:15:00 PM#)
    'test the start of trading on the third day 
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/31/2012 9:30:00 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/31/2012 9:44:59 AM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/31/2012 9:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/31/2012 4:00:00 PM#) = #10/31/2012 4:00:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#10/31/2012 4:45:00 PM#) = #10/31/2012 4:15:00 PM#)

    'same test but looking at the next trading day
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/26/2012 4:15:00 PM#) = #10/26/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/26/2012 4:15:01 PM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 9:00:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 9:30:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 9:44:59 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 9:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 4:00:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/29/2012 4:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    'continue with the second day of the storm Sandy
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 9:00:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 9:30:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 9:44:59 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 9:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 4:00:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/30/2012 4:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    'test the start of trading on the third day 
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 9:30:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 9:44:59 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 9:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:00:00 PM#) = #10/31/2012 4:00:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:15:00 PM#) = #10/31/2012 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:15:01 PM#) = #11/1/2012 9:45:00 AM#)

    'load the default holiday use for testing and validation for Tuesday-Wedesday January 15-16 1980
    'the market open at 12:00 
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 9:30:00 AM#) = #1/14/1980 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:00:01 PM#) = #1/14/1980 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:14:59 PM#) = #1/14/1980 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:15:00 PM#) = #1/15/1980 12:15:00 PM#)
    'the market close at 13:00 on the 16 but open at 9:30
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 9:30:00 AM#) = #1/15/1980 4:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 9:45:00 AM#) = #1/16/1980 9:45:00 AM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:14:59 PM#) = #1/16/1980 1:14:59 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:15:00 PM#) = #1/16/1980 1:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:15:01 PM#) = #1/16/1980 1:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 4:15:01 PM#) = #1/16/1980 1:15:00 PM#)

    'now same text but with the next function
    'test the opening at 12:00 on the 15 January 1980
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 9:30:00 AM#) = #1/15/1980 12:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 9:45:00 AM#) = #1/15/1980 12:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 12:14:59 PM#) = #1/15/1980 12:15:00 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 12:15:01 PM#) = #1/15/1980 12:15:01 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 4:14:59 PM#) = #1/15/1980 4:14:59 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/15/1980 4:15:01 PM#) = #1/16/1980 9:45:00 AM#)
    'test the closing at 13:00 on the 16 January
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/16/1980 1:14:59 PM#) = #1/16/1980 1:14:59 PM#)
    Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#1/16/1980 1:15:01 PM#) = #1/17/1980 9:45:00 AM#)




    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:00:01 PM#) = #1/14/1980 4:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:14:59 PM#) = #1/14/1980 4:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/15/1980 12:15:00 PM#) = #1/15/1980 12:15:00 PM#)
    ''the market close at 13:00 on the 16 but open at 9:30
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 9:30:00 AM#) = #1/15/1980 4:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 9:45:00 AM#) = #1/16/1980 9:45:00 AM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:14:59 PM#) = #1/16/1980 1:14:59 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:15:00 PM#) = #1/16/1980 1:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 1:15:01 PM#) = #1/16/1980 1:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTrade(#1/16/1980 4:15:01 PM#) = #1/16/1980 1:15:00 PM#)



    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 9:44:59 AM#) = #10/31/2012 9:45:00 AM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 9:45:00 AM#) = #10/31/2012 9:45:00 AM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:00:00 PM#) = #10/31/2012 4:00:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:15:00 PM#) = #10/31/2012 4:15:00 PM#)
    'Assert.IsTrue(YahooAccessData.ReportDate.DayOfTradeNext(#10/31/2012 4:15:01 PM#) = #11/1/2012 9:45:00 AM#)
  End Sub

  '<TestMethod()>
  'Public Sub TestFileLoad()
  '  Dim ThisReportWork = New YahooAccessData.Report
  '  Dim ThisReportNew = New YahooAccessData.Report
  '  Dim Connections As ConnectionStringSettingsCollection = ConfigurationManager.ConnectionStrings
  '  'file version 1.7.9
  '  'Dim ThisFile As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\bin\Debug\Report_Technology.rep"
  '  'file version 1.8.0
  '  Dim ThisFile As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\Report_Financial.rep"
  '  Dim ThisFileNew As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\Report_Financial_New.rep"

  '  Dim ThisReport = ThisReportWork.FileLoad(ThisFile)
  '  ThisReport.FileSave(ThisFileNew)
  '  'load the file again and compare for equality
  '  ThisReportNew = ThisReportWork.FileLoad(ThisFileNew)
  '  If ThisReportNew.EqualsDeep(ThisReport) = False Then
  '    Debug.Print("Report do not match...")
  '  End If

  '  Exit Sub

  '  With ThisReport
  '    .DateStart = Now
  '    .DateStop = .DateStart
  '    .Name = "Test"
  '    Dim ThisSplitFactorFuture = New YahooAccessData.SplitFactorFuture
  '    With ThisSplitFactorFuture
  '      .Name = "Name"
  '      .Symbol = "Symbol"
  '      .Exchange = "Exchange"
  '      .DateEx = Now
  '      .DateAnnounce = .DateEx
  '      .DatePayable = .DateEx
  '    End With
  '    .SplitFactorFutures.Add(ThisSplitFactorFuture)
  '  End With
  '  Dim dbStockEntity = New YahooAccessData.StockYahooEntities(Connections("StockYahooEntities").ConnectionString)
  '  With dbStockEntity
  '    .Reports.Add(ThisReport)
  '    .SaveChanges()
  '  End With
  'End Sub

  '<TestMethod()>
  'Public Sub TestLinkedHashset()
  '	Dim ThisLinkedHashSet As New YahooAccessData.LinkedHashSet(Of YahooAccessData.SearchKey(Of Integer, String), String)

  '	Dim ThisCollection As ICollection(Of YahooAccessData.SearchKey(Of Integer, String))

  '	Trace.WriteLine("Linked Hashset-------------")
  '	Dim ThisElement2 As YahooAccessData.SearchKey(Of Integer, String) = New YahooAccessData.SearchKey(Of Integer, String) With {.Item = 2, .KeyValue = "2"}
  '	With ThisLinkedHashSet
  '		.Add(New YahooAccessData.SearchKey(Of Integer, String) With {.Item = 1, .KeyValue = "1"})
  '		.Add(ThisElement2)
  '		.Add(New YahooAccessData.SearchKey(Of Integer, String) With {.Item = 3, .KeyValue = "3"})
  '		.Remove(ThisElement2)
  '		.Add(New YahooAccessData.SearchKey(Of Integer, String) With {.Item = 4, .KeyValue = "4"})
  '		.Add(New YahooAccessData.SearchKey(Of Integer, String) With {.Item = 5, .KeyValue = "5"})
  '	End With
  '	Trace.WriteLine("The element are enumerated in the correct order")
  '	For Each ThisItem In ThisLinkedHashSet
  '		Trace.WriteLine(ThisItem.Item.ToString)
  '	Next
  '	Trace.WriteLine("The element are also enumerated in the correct order with the ElementAt index loop")
  '	For I = 0 To ThisLinkedHashSet.Count - 1
  '		Trace.WriteLine(ThisLinkedHashSet.ElementAt(I).Item.ToString)
  '	Next
  '	Trace.WriteLine("This also work using the ICollection interface")
  '	ThisCollection = ThisLinkedHashSet
  '	Trace.WriteLine("The element are enumerated in the correct order")
  '	For Each ThisItem In ThisCollection
  '		Trace.WriteLine(ThisItem.Item.ToString)
  '	Next
  '	Trace.WriteLine("The element are also enumerated in the correct order with the ElementAtorDefault index loop")
  '	For I = 0 To ThisCollection.Count - 1
  '		Trace.WriteLine(ThisCollection(I).Item.ToString)
  '	Next
  '	Trace.WriteLine("-------------")
  'End Sub

  '<TestMethod()>
  Public Sub TestCollectionSpeed()
    Dim ThisReport As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisSort As YahooAccessData.ISort(Of YahooAccessData.Stock)
    Const NumberStock As Integer = 10000

    ThisReport = New YahooAccessData.Report
    With ThisReport
      .DateStart = Now
      .DateStop = .DateStart
      .Name = "Test"
      Trace.TraceInformation(String.Format("Creating {0} stock...", NumberStock))
      ThisStopWatch.Restart()
      '.Stocks
      For I = 0 To NumberStock - 1
        .StockAdd("Stock1" & I.ToString, "Sector1", "Industry1")
      Next
      ThisStopWatch.Stop()
      Trace.TraceInformation(String.Format("Time needed to create {0} stock is (ms): {1}", NumberStock, ThisStopWatch.ElapsedMilliseconds))

      Trace.TraceInformation(String.Format("Sorting {0} stock...", NumberStock))
      ThisSort = TryCast(.Stocks, YahooAccessData.ISort(Of YahooAccessData.Stock))
      ThisStopWatch.Restart()
      ThisSort.Sort()
      ThisStopWatch.Stop()
      Trace.TraceInformation(String.Format("Time needed to sort {0} stock is (ms): {1}", NumberStock, ThisStopWatch.ElapsedMilliseconds))
      Trace.TraceInformation(String.Format("First Stock Symbol is {0}", .Stocks(0).Symbol))


      Trace.TraceInformation(String.Format("Reversing {0} stock...", NumberStock))
      ThisStopWatch.Restart()
      ThisSort.Reverse()
      ThisStopWatch.Stop()
      Trace.TraceInformation(String.Format("Time needed to reverse {0} stock is (ms): {1}", NumberStock, ThisStopWatch.ElapsedMilliseconds))
      Trace.TraceInformation(String.Format("First Stock Symbol is {0}", .Stocks(0).Symbol))

      Trace.TraceInformation(String.Format("Sorting again {0} stock...", NumberStock))
      ThisStopWatch.Restart()
      ThisSort.Sort()
      ThisStopWatch.Stop()
      Trace.TraceInformation(String.Format("Time needed to sorting again {0} stock is (ms): {1}", NumberStock, ThisStopWatch.ElapsedMilliseconds))
      Trace.TraceInformation(String.Format("First Stock Symbol is {0}", .Stocks(0).Symbol))

      Assert.IsTrue(.Stocks.Count = NumberStock)
      Assert.IsTrue(.Sectors.Count = 1)
      Assert.IsTrue(.Industries.Count = 1)
    End With
  End Sub

  <TestMethod()>
  Sub TestFileSaveRecordIndexedVersion()
    Dim ThisReport As YahooAccessData.Report
    Dim ThisReportForAdd As YahooAccessData.Report
    Dim ThisReport1 As YahooAccessData.Report
    Dim ThisReport2 As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch

    Const FILE_NAME_ZIP1 As String = "Report_20130130_171005_V6_2_6_3.Zip"
    Const FILE_NAME_ZIP2 As String = "Report_20120709_104109_V5_3_1_3.Zip"
    Const FILE_NAME_TEST_ZIP As String = "TestZip.Zip"
    Const FILE_NAME_REP_STANDARD1 As String = "Report_Standard.rep"
    Const FILE_NAME_REP_STANDARD1_2 As String = "Report_Standard_1_2.rep"
    Const FILE_NAME_REP_NEW_RECORD_INDEXED As String = "Report_Record_Indexed.rep"

    Const FILE_PATH_TEST As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData VS2012\ConsoleTest\FileTest\"
    Const FILE_PATH As String = ""

    Console.WriteLine(String.Format("Loading zip file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP1))
    ThisStopWatch.Restart()
    ThisReport = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP1)
    'Console.WriteLine(String.Format("File Zip loading time is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    Console.WriteLine(String.Format("Saving file to: {0}", FILE_PATH & FILE_NAME_REP_STANDARD1))
    'ThisReport.FileSaveZip(FILE_PATH_TEST & FILE_NAME_TEST_ZIP)
    ThisReport.FileSaveZipAsync(FILE_PATH_TEST & FILE_NAME_TEST_ZIP).Wait()

    ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_TEST_ZIP)
    If ThisReport.EqualsDeep(ThisReport1) = True Then
      Console.WriteLine(String.Format("Zip Test and rep file version match..."))
    Else
      Console.WriteLine(String.Format("Zip Test and rep file version do not match..."))
      Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
      If Console.ReadKey.Key = ConsoleKey.N Then Return
    End If

    Console.WriteLine(String.Format("Writing record indexed file: {0}", FILE_PATH_TEST & FILE_NAME_REP_NEW_RECORD_INDEXED))
    ThisStopWatch.Restart()
    ThisReport.FileSaveAsync(FILE_PATH_TEST & FILE_NAME_REP_NEW_RECORD_INDEXED, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed).Wait()
    Console.WriteLine(String.Format("File save record indexed time is {0} ms", ThisStopWatch.ElapsedMilliseconds))

    Console.WriteLine(String.Format("Open new record indexed file..."))
    ThisReport2 = New YahooAccessData.Report
    ThisReport2.FileOpen(FILE_PATH_TEST & FILE_NAME_REP_NEW_RECORD_INDEXED)
    Console.WriteLine(String.Format("File reading record indexed time is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    Console.WriteLine(String.Format("Verify that the record indexed file match the original zip file"))
    If ThisReport.EqualsDeep(ThisReport2) = True Then
      Console.WriteLine(String.Format("The new record indexed file version match the original zip file..."))
    Else
      Console.WriteLine(String.Format("The new record indexed file version DO NOT MATCH the original zip file..."))
      Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
      If Console.ReadKey.Key = ConsoleKey.N Then Return
    End If
  End Sub

  <TestMethod()>
  Sub TestFileRecordIndexedVirtualReading()
    Dim ThisReport_Indexed As YahooAccessData.Report
    Dim ThisReport1 As YahooAccessData.Report
    Dim ThisReport2 As YahooAccessData.Report
    Dim ThisReport1_2 As YahooAccessData.Report

    Dim ThisStock As YahooAccessData.Stock
    Dim ThisRecord As YahooAccessData.Record
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch

    Const FILE_NAME_ZIP1 As String = "Report_20120709_101006_V5_3_0_0.Zip"
    Const FILE_NAME_ZIP2 As String = "Report_20120709_104109_V5_3_1_3.Zip"
    Const FILE_NAME_ZIP_OLD_FORMAT As String = "Report_20110915_100757_V1_7_7_0.Zip"
    Const FILE_NAME_REP_STANDARD1 As String = "Report_Standard.rep"
    Const FILE_NAME_REP_STANDARD1_2 As String = "Report_Standard_1_2.rep"
    Const FILE_NAME_REP_NEW_RECORD_INDEXED As String = "Report_Record_Indexed.rep"

    Const FILE_PATH_TEST As String = "H:\Test\"
    Const FILE_PATH As String = "H:\"

    Dim ThisDate As Date = Now


    Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP1))
    'test the old file format reading
    ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP_OLD_FORMAT)

    Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP1))
    ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP1)
    ThisReport1_2 = ThisReport1.CopyDeep
    Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP2))
    ThisReport2 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP2)
    Console.WriteLine(String.Format("Adding file {0} and {1}...", FILE_PATH_TEST & FILE_NAME_ZIP1, FILE_PATH_TEST & FILE_NAME_ZIP2))
    ThisReport1_2.Add(ThisReport2)
    'Console.WriteLine(String.Format("Saving added file to {0}...", FILE_PATH_TEST & FILE_NAME_REP_STANDARD1_2))
    'ThisReport1_2.FileSave(FILE_PATH_TEST & FILE_NAME_REP_STANDARD1_2)

    Console.WriteLine(String.Format("Opening File: {0} ", FILE_PATH & FILE_NAME_REP_NEW_RECORD_INDEXED))
    ThisReport_Indexed = New YahooAccessData.Report
    ThisStopWatch.Restart()
    ThisReport_Indexed.FileOpen(FILE_PATH & FILE_NAME_REP_NEW_RECORD_INDEXED)
    Console.WriteLine(ThisReport_Indexed.BondRates.Count.ToString)
    ThisStopWatch.Stop()
    Console.WriteLine(String.Format("File reading record indexed time is {0} ms", ThisStopWatch.ElapsedMilliseconds))

    If ThisReport_Indexed.EqualsDeep(ThisReport1) = False Then
      If ThisReport_Indexed.EqualsDeep(ThisReport1_2) = False Then
        Console.WriteLine(String.Format("File do not match..."))
        Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
        If Console.ReadKey.Key = ConsoleKey.N Then Return
      Else
        Console.WriteLine(String.Format("Files {0} match...", ThisReport_Indexed.ToString))
        Return
      End If
    End If

    Console.WriteLine(String.Format("Load the file again with a different time range..."))
    ThisStopWatch.Restart()
    ThisReport_Indexed.AsDateRange.Refresh(ThisDate, ThisDate)
    Console.WriteLine(String.Format("File reading with date change is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    Console.WriteLine(ThisReport_Indexed.ToString)
    Console.WriteLine(String.Format("Add file {0}", FILE_NAME_ZIP2))
    ThisReport_Indexed.Add(ThisReport2)
    Console.WriteLine(String.Format("Save the added file...", ThisStopWatch.ElapsedMilliseconds))
    ThisReport_Indexed.FileSave()
    Console.WriteLine(String.Format("Read all the data in the file..."))
    ThisReport_Indexed.AsDateRange.Refresh(ThisDate.AddYears(-1), ThisDate)
    Console.WriteLine(String.Format("Check if the indexed file match the original added file {0}", FILE_NAME_REP_STANDARD1_2))
    If ThisReport_Indexed.EqualsDeep(ThisReport1_2) = True Then
      Console.WriteLine(String.Format("Summed file and indexed file match..."))
    Else
      Console.WriteLine(String.Format("Summed file and indexed file DO NOT MATCH..."))
      'Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
      'If Console.ReadKey.Key = ConsoleKey.N Then Return
    End If
  End Sub

  <TestMethod()>
  Async Function TestFileRecordIndexedSavingAsync() As Task(Of Boolean)
    Dim ThisReport_Indexed As YahooAccessData.Report
    Dim ThisReport1 As YahooAccessData.Report
    Dim ThisReport2 As YahooAccessData.Report
    Dim ThisReport1_2 As YahooAccessData.Report
    'Dim ThisReport1_2 As YahooAccessData.Report

    Dim ThisStock As YahooAccessData.Stock
    Dim ThisRecord As YahooAccessData.Record
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch

    Const FILE_NAME_ZIP1 As String = "Report_20120709_101006_V5_3_0_0.Zip"
    Const FILE_NAME_ZIP2 As String = "Report_20120709_104109_V5_3_1_3.Zip"
    Const FILE_NAME_ZIP_OLD_FORMAT As String = "Report_20110915_100757_V1_7_7_0.Zip"
    Const FILE_NAME_REP_STANDARD1 As String = "Report_Standard.rep"
    Const FILE_NAME_REP_STANDARD1_2 As String = "Report_Standard_1_2.rep"
    Const FILE_NAME_REP_NEW_RECORD_INDEXED As String = "Report_Record_Indexed.rep"

    Const FILE_PATH_TEST As String = "H:\Test\"
    Const FILE_PATH As String = "H:\Test\"

    Dim ThisDate As Date = Now


    'Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP1))
    'test the old file format reading
    'ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP_OLD_FORMAT)

    Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP1))
    ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP1)
    ThisReport1_2 = ThisReport1.CopyDeep
    Console.WriteLine(String.Format("Opening file: {0}", FILE_PATH_TEST & FILE_NAME_ZIP2))
    ThisReport2 = YahooAccessData.ReportFrom.LoadZip(FILE_PATH_TEST & FILE_NAME_ZIP2)
    Console.WriteLine(String.Format("Adding file {0} and {1}...", FILE_PATH_TEST & FILE_NAME_ZIP1, FILE_PATH_TEST & FILE_NAME_ZIP2))
    ThisReport1_2.Add(ThisReport2)

    ThisReport_Indexed = New YahooAccessData.Report
    ThisReport_Indexed.FileSave(FILE_PATH_TEST & FILE_NAME_REP_STANDARD1_2, YahooAccessData.IMemoryStream.enuFileType.RecordIndexed)
    ThisReport_Indexed.FileOpen(FILE_PATH_TEST & FILE_NAME_REP_STANDARD1_2)
    ThisReport_Indexed.Add(ThisReport1)
    Console.WriteLine(String.Format("Saving added file to {0}...", FILE_PATH_TEST & FILE_NAME_REP_STANDARD1_2))
    ThisStopWatch.Restart()
    Await ThisReport_Indexed.FileSaveAsync
    'ThisReport_Indexed.FileSave()
    ThisStopWatch.Stop()
    Console.WriteLine(String.Format("File saving record indexed time is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    Console.WriteLine(String.Format("Adding file ThisReport2..."))
    ThisReport_Indexed.AsDateRange.Refresh(Now, Now)
    ThisReport_Indexed.Add(ThisReport2)
    Console.WriteLine(String.Format("Saving file ThisReport2..."))
    Await ThisReport_Indexed.FileSaveAsync
    Console.WriteLine(String.Format("Refresh..."))
    ThisReport_Indexed.AsDateRange.Refresh(ThisReport_Indexed.DateStart, ThisReport_Indexed.DateStop)

    'ThisReport_Indexed.AsDateRange.Refresh()


    'Console.WriteLine(String.Format("Saving File: {0} ", FILE_PATH & FILE_NAME_REP_NEW_RECORD_INDEXED))
    'ThisReport_Indexed = New YahooAccessData.Report
    'ThisStopWatch.Restart()
    'ThisReport_Indexed.FileSave(FILE_PATH & FILE_NAME_REP_NEW_RECORD_INDEXED)
    'Console.WriteLine(ThisReport_Indexed.BondRates.Count.ToString)
    'ThisStopWatch.Stop()
    'Console.WriteLine(String.Format("File reading record indexed time is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    ThisReport_Indexed.Name = ThisReport1_2.Name
    Console.WriteLine(String.Format(ThisReport_Indexed.ToString))
    If ThisReport_Indexed.EqualsDeep(ThisReport1) = False Then
      If ThisReport_Indexed.EqualsDeep(ThisReport1_2) = False Then
        Console.WriteLine(String.Format("File do not match..."))
        Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
        If Console.ReadKey.Key = ConsoleKey.N Then Return True
      Else
        Console.WriteLine(String.Format("Files {0} match...", ThisReport_Indexed.ToString))
        Return True
      End If
    End If
    Return True
    Console.WriteLine(String.Format("Load the file again with a different time range..."))
    ThisStopWatch.Restart()
    ThisReport_Indexed.AsDateRange.Refresh(ThisDate, ThisDate)
    Console.WriteLine(String.Format("File reading with date change is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    Console.WriteLine(ThisReport_Indexed.ToString)
    Console.WriteLine(String.Format("Add file {0}", FILE_NAME_ZIP2))
    ThisReport_Indexed.Add(ThisReport2)
    Console.WriteLine(String.Format("Save the added file...", ThisStopWatch.ElapsedMilliseconds))
    ThisReport_Indexed.FileSave()
    Console.WriteLine(String.Format("Read all the data in the file..."))
    ThisReport_Indexed.AsDateRange.Refresh(ThisDate.AddYears(-1), ThisDate)
    Console.WriteLine(String.Format("Check if the indexed file match the original added file {0}", FILE_NAME_REP_STANDARD1_2))
    If ThisReport_Indexed.EqualsDeep(ThisReport1_2) = True Then
      Console.WriteLine(String.Format("Summed file and indexed file match..."))
    Else
      Console.WriteLine(String.Format("Summed file and indexed file DO NOT MATCH..."))
      'Console.WriteLine(String.Format("Do you want to continue (Y or N)?"))
      'If Console.ReadKey.Key = ConsoleKey.N Then Return
    End If
    Return True
  End Function

  <TestMethod()>
  Sub TestFileAddZip()
    Dim ThisReport As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch

    'Const FILE_PATH As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\AddZip\Report_20111214_163434_V2_2_1_0.Zip"


    Const FILE_NAME_ZIP As String = "Report_20120123_100500_V2_2_1_0.Zip"
    Const FILE_NAME_REP As String = "Report_20120123_100500_V2_2_1_0.rep"
    Const FILE_PATH As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\AddZip\"


    ThisReport = New YahooAccessData.Report
    ThisStopWatch.Restart()
    ThisReport.FileLoadZip(FILE_PATH & FILE_NAME_ZIP)
    Console.WriteLine(String.Format("File load unzip time is {0} ms", ThisStopWatch.ElapsedMilliseconds))
    ThisReport.FileSave(FILE_PATH & FILE_NAME_REP)
    ThisReport.FileSave(FILE_PATH & FILE_NAME_REP)
  End Sub


  <TestMethod()>
  Sub TestFileAddDailyZip()
    Dim ThisReport As New YahooAccessData.Report
    Dim ThisReportLoad As New YahooAccessData.Report
    Dim ThisReportPartial As New YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisFileList As New List(Of String)
    Dim ThisFileListCopy As New List(Of String)
    Dim ThisFileLocalOrderList As New List(Of String)
    Dim ThisFileName As String
    Dim I As Integer

    Const FILE_NAME_REP As String = "ReportDaily_20120123.rep"
    Const FILE_PATH As String = "C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockDataBasicTest\TestFile\AddDailyZip\"

    With My.Computer.FileSystem
      ThisFileList.Clear()
      For Each ThisFileName In .GetFiles(FILE_PATH)
        If UCase(System.IO.Path.GetExtension(ThisFileName)) = ".ZIP" Then
          ThisFileList.Add(ThisFileName)
        End If
      Next
    End With
    If ThisFileList.Count = 0 Then Exit Sub
    'the file should be save in ascending data order
    'read back the file from the list in ascending date from the date tag on the file name
    ThisFileLocalOrderList.Clear()
    With My.Computer.FileSystem
      For Each ThisFileName In ThisFileList.
          Where(Function(ThisFile As String) UCase(.GetFileInfo(ThisFile).Extension) = ".ZIP").
          OrderBy(Of Date)(Function(ThisFile As String) FileDateExtract(.GetFileInfo(ThisFile).Name))
        ThisFileLocalOrderList.Add(ThisFileName)
      Next
    End With
    'we now have all file ordered 
    'ThisReport = New YahooAccessData.Report
    ThisStopWatch.Restart()
    For Each ThisFileName In ThisFileLocalOrderList
      I = I + 1
      Console.WriteLine(String.Format("Loading file {0}", My.Computer.FileSystem.GetFileInfo(ThisFileName).Name))
      ThisReportPartial = YahooAccessData.ReportFrom.LoadZip(ThisFileName)
      If I = 1 Then
        ThisReport = ThisReportPartial
      Else
        ThisReport.Add(ThisReportPartial)
      End If

      Console.WriteLine(String.Format("Number of BondRate is: {0}", ThisReport.BondRates.Count))
      Console.WriteLine(String.Format("Number of SplitFactorFuture is: {0}", ThisReport.SplitFactorFutures.Count))
      Console.WriteLine(String.Format("Number of sector is: {0}", ThisReport.Sectors.Count))
      Console.WriteLine(String.Format("Number of Industry is: {0}", ThisReport.Industries.Count))
      Console.WriteLine(String.Format("Number of Stock is: {0}", ThisReport.Stocks.Count))
      Console.WriteLine(String.Format("Number of Record is: {0}", ThisReport.Stocks(0).Records.Count))
    Next
    Console.WriteLine(String.Format("File daily load unzip time is {0} ms", ThisStopWatch.ElapsedMilliseconds))

    Console.WriteLine(String.Format("Start saving file {0}", FILE_PATH & FILE_NAME_REP))
    ThisStopWatch.Restart()
    ThisReport.FileSave(FILE_PATH & FILE_NAME_REP)
    Console.WriteLine(String.Format("Time to save file is {0} ms", ThisStopWatch.ElapsedMilliseconds))

    Console.WriteLine(String.Format("Start reading file {0}", FILE_PATH & FILE_NAME_REP))
    ThisStopWatch.Restart()
    ThisReportLoad = YahooAccessData.ReportFrom.Load(FILE_PATH & FILE_NAME_REP)
    If ThisReportLoad.Exception IsNot Nothing Then
      MsgBox(ThisReportLoad.Exception.MessageAll)
    End If
    If ThisReport.EqualsDeep(ThisReportLoad) Then
      Console.WriteLine("Report are equal")
    Else
      Console.WriteLine("Report are not equal")
    End If
    Console.WriteLine(String.Format("Time to read file is {0} ms", ThisStopWatch.ElapsedMilliseconds))
  End Sub

  Sub CompareVersion()
    Dim ThisReport As New YahooAccessData.Report
    Dim ThisReport1 As YahooAccessData.Report
    Dim ThisReport2 As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisFileList As New List(Of String)
    Dim ThisFileListCopy As New List(Of String)
    Dim ThisFileLocalOrderList As New List(Of String)
    Dim ThisFileName As String
    Dim I As Integer

    Const FILE_NAME_ZIP1 As String = "C:\Users\mdufour\YahooDownloadTest\Report_20120205_234851_V4_0_0_0.Zip"
    Const FILE_NAME_ZIP2 As String = "C:\Users\mdufour\YahooDownloadTest\Report_20120203_163435_V2_2_1_0.Zip"

    Console.WriteLine(String.Format("Loading file {0}", My.Computer.FileSystem.GetFileInfo(FILE_NAME_ZIP1).Name))
    ThisReport1 = YahooAccessData.ReportFrom.LoadZip(FILE_NAME_ZIP1)
    With ThisReport1
      Console.WriteLine(String.Format("Number of BondRate is: {0}", .BondRates.Count))
      Console.WriteLine(String.Format("Number of SplitFactorFuture is: {0}", .SplitFactorFutures.Count))
      Console.WriteLine(String.Format("Number of sector is: {0}", .Sectors.Count))
      Console.WriteLine(String.Format("Number of Industry is: {0}", .Industries.Count))
      Console.WriteLine(String.Format("Number of Stock is: {0}", .Stocks.Count))
      Console.WriteLine(String.Format("Number of Record is: {0}", .Stocks(0).Records.Count))
    End With
    Console.WriteLine(String.Format("Loading file {0}", My.Computer.FileSystem.GetFileInfo(FILE_NAME_ZIP2).Name))
    ThisReport2 = YahooAccessData.ReportFrom.LoadZip(FILE_NAME_ZIP2)
    With ThisReport2
      Console.WriteLine(String.Format("Number of BondRate is: {0}", .BondRates.Count))
      Console.WriteLine(String.Format("Number of SplitFactorFuture is: {0}", .SplitFactorFutures.Count))
      Console.WriteLine(String.Format("Number of sector is: {0}", .Sectors.Count))
      Console.WriteLine(String.Format("Number of Industry is: {0}", .Industries.Count))
      Console.WriteLine(String.Format("Number of Stock is: {0}", .Stocks.Count))
      Console.WriteLine(String.Format("Number of Record is: {0}", .Stocks(0).Records.Count))
    End With
    If ThisReport1.EqualsDeep(ThisReport2, IsIgnoreID:=True) Then
      Console.WriteLine("Report are equal")
    Else
      Console.WriteLine("Report are not equal")
    End If

    Console.WriteLine(String.Format("Time to read file is {0} ms", ThisStopWatch.ElapsedMilliseconds))
  End Sub

  '<TestMethod()>
  Public Sub TestFileRead()
    Dim ThisReportNew As New YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisProcess As Process = Process.GetCurrentProcess
    Dim I As Integer

    Const FILE_NAME_REP As String = "Report_Financial.rep"
    'Const FILE_NAME_REP As String = "Report_Financial.rep"
    Const FILE_PATH As String = "C:\Users\mdufour\AppData\Roaming\MD\Yahoodownload\1.7.7.0\Download Last\"

    Console.WriteLine(String.Format("Start reading file {0}", FILE_PATH & FILE_NAME_REP))
    ThisStopWatch.Restart()
    ThisProcess.Refresh()
    Console.WriteLine(String.Format("Memory size before reading file is {0:f1} KB", ThisProcess.PrivateMemorySize64 / 1024))
    Console.WriteLine(String.Format("Press any key to continue..."))
    Console.ReadLine()

    Dim ThisReport As New YahooAccessData.Report
    'With ThisReport
    '  For I = 1 To 1000
    '    '.StockAdd("Stock" & I.ToString, "Sector1", "Industry1")
    '    .Sectors.Add(New YahooAccessData.Sector With {.Name = "Sector" & I.ToString})
    '    '.StockAdd("Stock" & I.ToString, "Sector1", "Industry1")
    '  Next
    'End With
    ThisReport = YahooAccessData.ReportFrom.Load(FILE_PATH & FILE_NAME_REP)
    ThisProcess.Refresh()

    Console.WriteLine(ThisReport.ToString)
    Console.WriteLine(String.Format("Memory size after reading file is {0:f1} KB", ThisProcess.PrivateMemorySize64 / 1024))
    Console.WriteLine(String.Format("Press any key to continue..."))
    Console.ReadLine()
    'ThisReport.Dispose()
    'ThisReport = Nothing
    ThisProcess.Refresh()
    Console.WriteLine(String.Format("Memory size after closing file is {0:f1} ", ThisProcess.PrivateMemorySize64 / 1024))
    ThisStopWatch.Stop()
    'ThisReport.FileSave(FILE_PATH & FILE_NAME_REP & ".new")
    'ThisReportNew = YahooAccessData.ReportFrom.Load(FILE_PATH & FILE_NAME_REP & ".new")
    'If ThisReportNew.EqualsDeep(ThisReport) Then
    '	Console.WriteLine("Report are equal")
    'Else
    '	Console.WriteLine("Report are not equal")
    'End If
    Console.WriteLine(String.Format("Time to read file is {0} ms", ThisStopWatch.ElapsedMilliseconds))
  End Sub

  Public Async Sub TestFileNewVersion_1_2()
    Dim ThisReportNew As New YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim I As Integer

    Const FILE_NAME_REP As String = "Report_Financial.rep"
    Const FILE_NAME_REP_NEW As String = "Report_Financial.new.rep"
    'Const FILE_NAME_REP As String = "Report_Financial.rep"
    Const FILE_PATH As String = "C:\Users\mdufour\AppData\Roaming\MD\Yahoodownload\1.7.7.0\Download Last\"

    Dim ThisReport As New YahooAccessData.Report
    ThisReport = YahooAccessData.ReportFrom.Load(FILE_PATH & FILE_NAME_REP)
    Await ThisReport.FileSaveAsync(FILE_PATH & FILE_NAME_REP_NEW)
    ThisReportNew = YahooAccessData.ReportFrom.Load(FILE_PATH & FILE_NAME_REP_NEW)
    If ThisReportNew.EqualsDeep(ThisReport) Then
      Console.WriteLine("Report are equal")
    Else
      Console.WriteLine("Report are not equal")
    End If
  End Sub


  Private Function FileDateExtract(ByVal ThisFile As String) As Date
    Dim ThisFileName As String = My.Computer.FileSystem.GetFileInfo(ThisFile).Name
    Dim ThisFileDate As Date

    Dim ThisPosStartDate As Integer = InStr(ThisFileName, "_") + 1
    Dim ThisPosStartTime As Integer = InStr(ThisPosStartDate + 1, ThisFileName, "_") + 1

    Dim ThisDate As String = Mid(ThisFileName, ThisPosStartDate, 8)
    Dim ThisTime As String = Mid(ThisFileName, ThisPosStartTime, 6)
    ThisFileDate = DateSerial(CInt(Mid(ThisDate, 1, 4)), CInt(Mid(ThisDate, 5, 2)), CInt(Mid(ThisDate, 7, 2)))
    ThisFileDate = ThisFileDate.AddHours(CInt(Mid(ThisTime, 1, 2))).AddMinutes(CInt(Mid(ThisTime, 3, 2))).AddSeconds(CInt(Mid(ThisTime, 5, 2)))
    Return ThisFileDate
  End Function
End Class