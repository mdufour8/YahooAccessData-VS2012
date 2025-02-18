#Region "Imports"
Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Threading.Tasks
#End Region

Namespace MathPlus.Filter
	''' <summary>
	''' This is an implementation of a double exponential filtering for simple prediction using the 
	''' Brown's linear exponential smoothing (LES) method. This version is equivalent of 
	''' the second derivative of the function 
	''' See wikipedia: https://en.wikipedia.org/wiki/Exponential_smoothing#Double_exponential_smoothing
	''' </summary>
	<Serializable()>
	Public Class FilterLowPassPLLPredict
		Inherits FilterLowPassExpPredict

		Public Sub New(ByVal FilterRate As Double, ByVal NumberToPredict As Double)
			MyBase.New(
				NumberToPredict:=NumberToPredict,
				FilterHead:=New FilterLowPassPLL(FilterRate, IsPredictionEnabled:=False),
				FilterBase:=New FilterLowPassExp(FilterRate, IsPredictionEnabled:=False))
		End Sub

		Public Sub New(ByVal NumberToPredict As Double, ByVal FilterHead As IFilter, ByVal FilterBase As IFilter)
			MyBase.New(
				NumberToPredict:=NumberToPredict,
				FilterHead:=FilterHead,
				FilterBase:=FilterBase)
		End Sub

		Public Overrides Function ToString() As String
			Return $"{Me.GetType().Name}: FilterRate={Me.FilterRate}"
		End Function

	End Class
End Namespace