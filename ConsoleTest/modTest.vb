Imports System
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.ExtensionService
Imports System.Text
Imports System.IO
Imports System.Configuration
Imports System.Runtime.CompilerServices
Imports System.Collections.Generic
Imports System.Linq

#Const TRACE = True

#Region "modTest"
Module modTest
  Public Const NUMBER_POINT As Integer = 1000

  Public Sub Main()
    Dim ThisConsoleTracer = New ConsoleTraceListener(True)
    Trace.Listeners.Add(ThisConsoleTracer)


    Call TestFilterMemory()

    'Call TestFilterControl()

    'ThisTest.TestTradingDate()
    'TestCollectionSpeed()
    'ThisTest.TestFileAddZip()
    'ThisTest.TestFileAddDailyZip()
    'ThisTest.TestReadSpeed()
    'ThisTest.CompareVersion()
    'ThisTest.TestBasicOperation()
    'ThisTest.TestFileSaveRecordIndexedVersion()
    'ThisTest.TestFileRecordIndexedVirtualReading()
    'ThisTask = ThisTest.TestFileRecordIndexedSavingAsync()
    'Try
    '	ThisTask.Wait()
    'Catch ex As Exception
    '	Console.WriteLine(String.Format(ThisTask.Exception.MessageAll))
    'End Try

    'ThisTest.TestFileNewVersion_1_2()
    'Dim ThisResponse As String
    'Do
    '  ThisTest.TestFileNewVersion_1_2()
    '  Console.WriteLine(String.Format("Do you want to run again? (Yes or No)..."))
    '  ThisResponse = UCase(Console.ReadLine().ToString)
    'Loop Until ThisResponse = "N"

    Console.WriteLine(String.Format("Press any key to continue..."))
    Console.ReadLine()
  End Sub

  Public Sub TestCollectionSpeed()
    Dim ThisReport As YahooAccessData.Report
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisSort As YahooAccessData.ISort(Of YahooAccessData.Stock)
    Const NumberStock As Integer = 10000

    ThisReport = New YahooAccessData.Report

    With ThisReport
      .DateStart = Now
      .DateStop = .DateStart
      .Name = "Test1"
      Trace.TraceInformation(String.Format("Creating {0} stock...", NumberStock))
      ThisStopWatch.Restart()
      '.Stocks
      For I = 0 To NumberStock - 1
        .StockAdd("Stock" & I.ToString, "Sector1", "Industry1")
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

    End With
  End Sub


  Public Sub TestFilterMemory()
    Dim ThisListOfFilter As IList(Of IFilter)
    Dim I As Integer
    Dim ThisTestSignal() As Double = YahooAccessData.MathPlus.WaveForm.Sinus(Amplitude:=10, NumberCycles:=10, PhaseDeg:=0, NumberSamples:=NUMBER_POINT, Mean:=100)
    Dim ThisFilter As IFilter
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisStopWatchTotal = New System.Diagnostics.Stopwatch

    Const NUMBER_FILTER As Integer = 100


    ThisListOfFilter = New List(Of IFilter)

    Console.WriteLine(String.Format("Press any key to start creation of {0} filters...", NUMBER_FILTER))
    Console.ReadLine()
    ThisStopWatchTotal.Restart()
    ThisStopWatch.Restart()
    For I = 0 To NUMBER_FILTER - 1
      ThisFilter = New FilterLowPassExp(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=False)
      ThisListOfFilter.Add(New FilterLowPassExp(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True))
      'ThisListOfFilter.Add(New FilterLowPassExp(FilterRate:=10, IsPredictionEnabled:=True))
      If ThisStopWatch.ElapsedMilliseconds > 1000 Then
        ThisStopWatch.Restart()
        Console.WriteLine(String.Format("Completed {0} new filters", I))
      End If
    Next
    ThisStopWatch.Stop()
    ThisStopWatchTotal.Stop()
    Console.WriteLine(String.Format("Time creation required per filter is {0} ms", ThisStopWatchTotal.ElapsedMilliseconds / NUMBER_FILTER))

    Console.WriteLine(String.Format("Press any key to start refresh of {0} filters...", NUMBER_FILTER))
    Console.ReadLine()

    ThisStopWatchTotal.Restart()
    ThisStopWatch.Restart()
    For I = 0 To NUMBER_FILTER - 1
      With DirectCast(ThisListOfFilter(I), IFilterControl)
        .Refresh(.FilterRate)
      End With
      'ThisListOfFilter(I).Filter(ThisTestSignal)
      If ThisStopWatch.ElapsedMilliseconds > 1000 Then
        ThisStopWatch.Reset()
        Console.WriteLine(String.Format("Completed {0} new filters", I))
      End If
    Next
    ThisStopWatch.Stop()
    ThisStopWatchTotal.Stop()
    Console.WriteLine(String.Format("Time required per filter refresh is {0} ms", ThisStopWatchTotal.ElapsedMilliseconds / NUMBER_FILTER))
  End Sub

  Public Sub TestFilterControl()
    Dim ThisTestSignal() As Double = YahooAccessData.MathPlus.WaveForm.Sinus(Amplitude:=10, NumberCycles:=10, PhaseDeg:=0, NumberSamples:=NUMBER_POINT, Mean:=100)

    Call TestFilter(
      Title:="Filter Low Pass Exponential",
      Filter1:=New FilterLowPassExp(10, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExp(20, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExp(10, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass Exponential with Input Signal",
      Filter1:=New FilterLowPassExp(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExp(FilterRate:=20, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExp(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass Exponential No Delay",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass Exponential No Delay with Unitary Look Ahead",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, NumberLookAheadPoint:=1, IsPredictionEnabled:=True))

    Dim ThisFilter1 As IFilter = New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True)
    Dim ThisFilter2 As IFilter = New FilterLowPassExpNoDelay(FilterRate:=20, InputValue:=ThisTestSignal, IsPredictionEnabled:=True)
    Dim ThisFilterTest As IFilter = New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True)

    Call TestFilter(
      Title:="Filter Low Pass Exponential No Delay with Input Signal",
      Filter1:=ThisFilter1,
      Filter2:=ThisFilter2,
      FilterTest:=ThisFilterTest)

    Call TestFilter(
      Title:="Filter Low Pass Exponential No Delay with Input Signal and Unitary Look Ahead",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass PLL",
      Filter1:=New FilterLowPassPLL(10, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassPLL(20, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassPLL(10, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass PLL with Input Signal",
      Filter1:=New FilterLowPassPLL(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassPLL(FilterRate:=20, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassPLL(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass PLL No Delay",
      Filter1:=New FilterLowPassPLLNoDelay(FilterRate:=10, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassPLLNoDelay(FilterRate:=20, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassPLLNoDelay(FilterRate:=10, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass PLL No Delay with Unitary Look Ahead",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, NumberLookAheadPoint:=1, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass Pll No Delay with Input Signal",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, InputValue:=ThisTestSignal, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, IsPredictionEnabled:=True))

    Call TestFilter(
      Title:="Filter Low Pass Pll No Delay with Input Signal and Unitary Look Ahead",
      Filter1:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      Filter2:=New FilterLowPassExpNoDelay(FilterRate:=20, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True),
      FilterTest:=New FilterLowPassExpNoDelay(FilterRate:=10, InputValue:=ThisTestSignal, NumberLookAheadPoint:=1, IsPredictionEnabled:=True))
  End Sub

  Private Sub TestFilter(ByVal Title As String, ByVal Filter1 As IFilter, ByVal Filter2 As IFilter, ByVal FilterTest As IFilter)
    Dim ThisStopWatch = New System.Diagnostics.Stopwatch
    Dim ThisTestSignal() As Double
    Dim I As Integer
    Dim IsSuccess As Boolean
    Dim ThisFilterControl1 As IFilterControl
    Dim ThisFilterControl2 As IFilterControl
    Dim ThisFilterTestControl As IFilterControl
    Dim ThisFilterPrediction1 As IFilterPrediction
    Dim ThisFilterPrediction2 As IFilterPrediction
    Dim ThisFilterTestPrediction As IFilterPrediction

    Const NUMBER_FILTER_RATE_CHANGE_FOR_SPEED_TEST As Integer = 100

    ThisFilterControl1 = DirectCast(Filter1, IFilterControl)
    ThisFilterControl2 = DirectCast(Filter2, IFilterControl)
    ThisFilterTestControl = DirectCast(FilterTest, IFilterControl)
    ThisFilterPrediction1 = DirectCast(Filter1, IFilterPrediction)
    ThisFilterPrediction2 = DirectCast(Filter2, IFilterPrediction)
    ThisFilterTestPrediction = DirectCast(FilterTest, IFilterPrediction)

    Trace.TraceInformation("---------------------------------------------")
    Trace.TraceInformation(String.Format("{0}{1}{2}", "<<<<<<<<<<", Title, ">>>>>>>>>>"))

    ThisTestSignal = YahooAccessData.MathPlus.WaveForm.Sinus(Amplitude:=10, NumberCycles:=10, PhaseDeg:=0, NumberSamples:=NUMBER_POINT, Mean:=100)

    'If TypeOf Filter1 Is FilterLowPassExpNoDelay Then
    '  Debugger.Break()
    'End If
    If ThisFilterControl1.IsInputEnabled = False Then Filter1.Filter(ThisTestSignal)
    If ThisFilterControl1.IsInputEnabled = False Then Filter2.Filter(ThisTestSignal)
    If ThisFilterControl1.IsInputEnabled = False Then FilterTest.Filter(ThisTestSignal)

    'test that we can succesfully change the filter rate 
    Trace.TraceInformation(String.Format("Checking that we can succesfully change the filter rate for {0}", TypeName(FilterTest)))
    ThisFilterTestControl.Refresh(Filter2.Rate)
    'verify that the two filter result match
    IsSuccess = True
    If ThisFilterControl1.IsInputEnabled = False Then FilterTest.Filter(ThisTestSignal)
    For I = 0 To ThisTestSignal.Length - 1
      If FilterTest.ToList(I) <> Filter2.ToList(I) Then
        IsSuccess = False
        Exit For
      End If
      If ThisFilterTestPrediction.ToListOfGainPerYear(I) <> ThisFilterPrediction2.ToListOfGainPerYear(I) Then
        IsSuccess = False
        Exit For
      End If
    Next
    If IsSuccess Then
      Trace.TraceInformation(String.Format("Changing the filter rate for '{0}' is a SUCCESS!", TypeName(FilterTest)))
    Else
      Trace.TraceInformation(String.Format("Changing the filter rate for {0} FAILED!", TypeName(FilterTest)))
      Return
    End If

    ThisStopWatch.Reset()
    ThisStopWatch.Start()
    For J = 0 To NUMBER_FILTER_RATE_CHANGE_FOR_SPEED_TEST - 1
      If FilterTest.Rate = Filter1.Rate Then
        ThisFilterTestControl.Refresh(Filter2.Rate)
        If ThisFilterControl1.IsInputEnabled = False Then FilterTest.Filter(ThisTestSignal)
      Else
        ThisFilterTestControl.Refresh(Filter1.Rate)
        If ThisFilterControl1.IsInputEnabled = False Then FilterTest.Filter(ThisTestSignal)
      End If
    Next
    ThisStopWatch.Stop()
    Trace.TraceInformation(String.Format("Filter rate changing speed for {0} is {1} ms/1000 Points", TypeName(FilterTest), (1000 / NUMBER_POINT) * ThisStopWatch.ElapsedMilliseconds / NUMBER_FILTER_RATE_CHANGE_FOR_SPEED_TEST))
  End Sub

End Module
#End Region


