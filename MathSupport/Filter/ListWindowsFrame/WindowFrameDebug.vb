Imports System.Diagnostics


''' <summary>
''' Allow debugging of the windowsfarem class as follow:
''' [WF-Add] Index=49 Open=105.00 High=107.00 Low=104.00 Last=106.50 Vol=12000
'''[WF-State] Count=50 HighIdx=12 High=109.20 (Last=108.75) LowIdx=3 Low=101.75 (Last=102.10)
'''[WF-Decimate] Index=49 Raw(Open=105.00, High=107.00, Low=104.00, Last=106.50) 
'''→ Dec(Open=105.00, High=109.20, Low=101.75, Last=106.80, Vol=54321)
'''[WF-Summary] Count=50 First(Open=100.50, Last=101.00) Last(Open=105.00, Last=106.50) High(Idx=12, High=109.20, Last=108.75) Low(Idx=3, Low=101.75, Last=102.10)
''' </summary>
Friend Module WindowFrameDebug

	Public Enum DebugLevel
		Off        ' no logging
		Summary    ' only TraceSummary
		Normal     ' Add + State
		Detailed   ' Add + State + Decimate
	End Enum

	' To enable tracing for debugging:
	'   WindowFrameDebug.Level = WindowFrameDebug.DebugLevel.Normal
	'   WindowFrameDebug.Level = WindowFrameDebug.DebugLevel.Detailed
	'   WindowFrameDebug.Level = WindowFrameDebug.DebugLevel.Summary
	' Default is Off
	Public Property Level As DebugLevel = DebugLevel.Off

	<Conditional("DEBUG")>
	Public Sub TraceAdd(Of T As {Class, IPriceVol, IPricePivotPoint})(item As T, index As Integer)
		If Level >= DebugLevel.Normal Then
			Trace.WriteLine($"[WF-Add] Index={index} " &
														$"Open={item.Open:F2} High={item.High:F2} Low={item.Low:F2} " &
														$"Last={item.Last:F2} Vol={item.Vol}")
		End If
	End Sub

	<Conditional("DEBUG")>
	Public Sub TraceState(Of T As {Class, IPriceVol, IPricePivotPoint})(list As List(Of T), highIdx As Integer, lowIdx As Integer)
		If Level >= DebugLevel.Normal Then
			If list.Count = 0 Then
				Trace.WriteLine("[WF-State] Empty list.")
				Return
			End If
			Dim highItem = list(highIdx)
			Dim lowItem = list(lowIdx)
			Trace.WriteLine($"[WF-State] Count={list.Count} " &
														$"HighIdx={highIdx} High={highItem.High:F2} (Last={highItem.Last:F2}) " &
														$"LowIdx={lowIdx} Low={lowItem.Low:F2} (Last={lowItem.Last:F2})")
		End If
	End Sub

	<Conditional("DEBUG")>
	Public Sub TraceDecimate(Of T As {Class, IPriceVol, IPricePivotPoint})(lastItem As T, decimated As T, index As Integer)
		If Level >= DebugLevel.Detailed Then
			Trace.WriteLine($"[WF-Decimate] Index={index} " &
														$"Raw(Open={lastItem.Open:F2}, High={lastItem.High:F2}, Low={lastItem.Low:F2}, Last={lastItem.Last:F2}) " &
														$"→ Dec(Open={decimated.Open:F2}, High={decimated.High:F2}, Low={decimated.Low:F2}, Last={decimated.Last:F2}, Vol={decimated.Vol})")
		End If
	End Sub

	<Conditional("DEBUG")>
	Public Sub TraceSummary(Of T As {Class, IPriceVol, IPricePivotPoint})(list As List(Of T), highIdx As Integer, lowIdx As Integer)
		If Level >= DebugLevel.Summary Then
			If list.Count = 0 Then
				Trace.WriteLine("[WF-Summary] Empty list.")
				Return
			End If
			Dim first = list.First()
			Dim last = list.Last()
			Dim highItem = list(highIdx)
			Dim lowItem = list(lowIdx)
			Trace.WriteLine($"[WF-Summary] Count={list.Count} " &
														$"First(Open={first.Open:F2}, Last={first.Last:F2}) " &
														$"Last(Open={last.Open:F2}, Last={last.Last:F2}) " &
														$"High(Idx={highIdx}, High={highItem.High:F2}, Last={highItem.Last:F2}) " &
														$"Low(Idx={lowIdx}, Low={lowItem.Low:F2}, Last={lowItem.Last:F2})")
		End If
	End Sub

End Module
