#Region "Imports directives"

Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Reflection
Imports System.Text
Imports YahooAccessData.ExtensionService

#End Region

<ComVisible(True), _
ClassInterface(ClassInterfaceType.AutoDual), _
Guid("FEFCCEB4-935C-41C6-888E-5A9CF9505FA5")> _
Public Class Functions
  Public Function Sinus(ByVal CellStart As Excel.Range, ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double) As Double(,)
    Dim ThisArray As Double(,) '<-- declared as 2D Array
    Dim ThisSinus As Double()
    Dim K As Integer
    Try
      Dim AC = CellStart.Application.Caller
      Dim ThisRange As Excel.Range = CType(AC, Excel.Range)
      Dim NRows = ThisRange.Rows.Count
      Dim NCols = ThisRange.Columns.Count
      Dim NumberPoints = NCols * NRows

      ReDim ThisArray(0 To NRows - 1, 0 To NCols - 1)
      ReDim ThisSinus(0 To NumberPoints - 1)

      ThisSinus = YahooAccessData.MathPlus.WaveForm.Sinus(Amplitude, NumberCycles, PhaseDeg, NumberPoints)
      For J As Integer = 0 To NCols - 1
        For I As Integer = 0 To NRows - 1
          ThisArray(I, J) = ThisSinus(K)
          K = K + 1
        Next
      Next
    Catch ex As Exception
    End Try
    Return ThisArray
  End Function

  Public Function Square(ByVal CellStart As Excel.Range, ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double) As Double(,)
    Dim ThisArray As Double(,) '<-- declared as 2D Array
    Dim ThisSinus As Double()
    Dim K As Integer
    Try
      Dim AC = CellStart.Application.Caller
      Dim ThisRange As Excel.Range = CType(AC, Excel.Range)
      Dim NRows = ThisRange.Rows.Count
      Dim NCols = ThisRange.Columns.Count
      Dim NumberPoints = NCols * NRows

      ReDim ThisArray(0 To NRows - 1, 0 To NCols - 1)
      ReDim ThisSinus(0 To NumberPoints - 1)

      ThisSinus = YahooAccessData.MathPlus.WaveForm.Square(Amplitude, NumberCycles, PhaseDeg, NumberPoints)
      For J As Integer = 0 To NCols - 1
        For I As Integer = 0 To NRows - 1
          ThisArray(I, J) = ThisSinus(K)
          K = K + 1
        Next
      Next
    Catch ex As Exception
    End Try
    Return ThisArray
  End Function


  Function GetCountIncrement(ByRef CellStart As Excel.Range, Optional InitialCount As Integer = 1, Optional StepSize As Integer = 1) As Integer(,)
    Dim ThisArray As Integer(,) = Nothing '<-- declared as 2D Array
    Try
      Dim AC = CellStart.Application.Caller
      Dim ThisRange As Excel.Range = CType(AC, Excel.Range)
      Dim NRows = ThisRange.Rows.Count
      Dim NCols = ThisRange.Columns.Count

      ReDim ThisArray(0 To NRows - 1, 0 To NCols - 1)

      For J As Integer = 0 To NCols - 1
        For I As Integer = 0 To NRows - 1
          ThisArray(I, J) = InitialCount
          InitialCount = InitialCount + StepSize
        Next
      Next
    Catch ex As Exception
    End Try
    Return ThisArray
  End Function

#Region "Registration of Automation Add-in"
  ''' <summary>
  ''' This is function which is called when we register the dll
  ''' </summary>
  ''' <param name="type"></param>
  ''' <remarks></remarks>
  <ComRegisterFunction()> _
  Public Shared Sub RegisterFunction(ByVal type As Type)

      ' Add the "Programmable" registry key under CLSID
      Registry.ClassesRoot.CreateSubKey(GetCLSIDSubKeyName( _
                                        type, "Programmable"))

      ' Register the full path to mscoree.dll which makes Excel happier.
      Dim key As RegistryKey = Registry.ClassesRoot.OpenSubKey( _
      GetCLSIDSubKeyName(type, "InprocServer32"), True)
      key.SetValue("", (Environment.SystemDirectory & "\mscoree.dll"), _
                    RegistryValueKind.String)

  End Sub

  ''' <summary>
  ''' This is function which is called when we unregister the dll
  ''' </summary>
  ''' <param name="type"></param>
  ''' <remarks></remarks>
  <ComUnregisterFunction()> _
  Public Shared Sub UnregisterFunction(ByVal type As Type)

      ' Remove the "Programmable" registry key under CLSID
      Registry.ClassesRoot.DeleteSubKey( _
      GetCLSIDSubKeyName(type, "Programmable"), False)

  End Sub

  ''' <summary>
  ''' Assistant function used by RegisterFunction/UnregisterFunction
  ''' </summary>
  ''' <param name="type"></param>
  ''' <param name="subKeyName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Shared Function GetCLSIDSubKeyName( _
  ByVal type As Type, ByVal subKeyName As String) As String

      Dim s As New StringBuilder
      s.Append("CLSID\{")
      s.Append(type.GUID.ToString.ToUpper)
      s.Append("}\")
      s.Append(subKeyName)
      Return s.ToString

  End Function
#End Region
End Class
