Imports System.Diagnostics
Imports System.Runtime.CompilerServices

Public Module PriceVolIdentityTracker

	''' <summary>
	''' Print the unique instance ID and values of a PriceVol.
	''' ID comes from RuntimeHelpers.GetHashCode (stable per object lifetime).
	''' </summary>
	<Conditional("DEBUG")>
	Public Sub TrackIdentity(label As String, pv As PriceVol)
		If pv Is Nothing Then
			Trace.WriteLine($"[TrackIdentity:{label}] PriceVol = Nothing")
			Return
		End If

		Dim id As Integer = RuntimeHelpers.GetHashCode(pv)
		Trace.WriteLine($"[TrackIdentity:{label}] ID={id} {pv}")
	End Sub

	''' <summary>
	''' Compare two PriceVols: logs whether they are the same reference,
	''' and whether their values differ.
	''' </summary>
	<Conditional("DEBUG")>
	Public Sub CompareIdentity(label As String, pv1 As PriceVol, pv2 As PriceVol)
		If pv1 Is Nothing OrElse pv2 Is Nothing Then
			Trace.WriteLine($"[CompareIdentity:{label}] One or both PriceVols are Nothing.")
			Return
		End If

		Dim id1 As Integer = RuntimeHelpers.GetHashCode(pv1)
		Dim id2 As Integer = RuntimeHelpers.GetHashCode(pv2)

		If Object.ReferenceEquals(pv1, pv2) Then
			Trace.WriteLine($"[CompareIdentity:{label}] SAME reference (ID={id1})")
		Else
			Trace.WriteLine($"[CompareIdentity:{label}] DIFFERENT references (ID1={id1}, ID2={id2})")
		End If

		If Not pv1.Equals(pv2) Then
			Trace.WriteLine($"[CompareIdentity:{label}] Values differ:")
			Trace.WriteLine($"    pv1: {pv1}")
			Trace.WriteLine($"    pv2: {pv2}")
		Else
			Trace.WriteLine($"[CompareIdentity:{label}] Values are equal.")
		End If
	End Sub
End Module

