Namespace MathPlus.Filter
  Public Interface IFilterPLLDetector
    Function RunErrorDetector(ByVal Input As Double, ByVal InputFeedback As Double) As Double
    Sub RunConvergence(ByVal NumberOfIteration As Integer, ByVal ValueBegin As Double, ByVal ValueEnd As Double)
    ReadOnly Property Status As Boolean
    ReadOnly Property ToCount As Integer
    ReadOnly Property ToErrorLimit As Double
    ReadOnly Property ErrorLast As Double
    ReadOnly Property Count As Double
    ReadOnly Property ValueInit As Double
    ReadOnly Property IsMinimum As Boolean
    ReadOnly Property IsMaximum As Boolean
    ReadOnly Property Minimum As Double
    ReadOnly Property Maximum As Double
    ReadOnly Property ToList() As IList(Of Double)
    ReadOnly Property ToListOfPriceMedianNextDayLow() As IList(Of Double)
    ReadOnly Property ToListOfPriceMedianNextDayHigh() As IList(Of Double)
    ReadOnly Property ToListOfConvergence() As IList(Of Double)
    ReadOnly Property ToListOfVolatility() As IList(Of Double)
    ReadOnly Property DetectorBalance As Double
    Property Tag As String
    Function ValueOutput(ByVal Input As Double, ByVal InputFeedback As Double) As Double
  End Interface
End Namespace