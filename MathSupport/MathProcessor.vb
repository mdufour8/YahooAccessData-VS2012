
Imports YahooAccessData
Imports YahooAccessData.StockPriceVol

Namespace MathPlus

	' This class provides static methods for performing arithmetic operations on operands of different types.
	'The current implementation is thread safe and does not modify the input operands.
	'It returns new instances containing the results of the operations.
	Public Class MathProcessor

		Public Const CORRELATION_HL_RHO_Default As Double = 0.5


#Region " ---Add---"
		''' <summary>
		''' Adds a scalar value to each element of a list of doubles. This implementation creates a new sequence 
		''' without altering the original collection. it returns a new list containing the transformed values and is thread safe.
		''' </summary>
		''' <param name="a"></param>
		''' <param name="b"></param>
		''' <returns>The result in a new list of doubles.</returns>
		Public Shared Function Add(a As Double, b As List(Of Double)) As List(Of Double)
			Return b.Select(Function(x) x + a).ToList()
		End Function

		Public Shared Function Add(a As List(Of Double), b As Double) As List(Of Double)
			Return a.Select(Function(x) x + b).ToList()
		End Function

		Public Shared Function Add(a As List(Of Double), b As List(Of Double)) As List(Of Double)
			If a.Count <> b.Count Then
				Throw New InvalidOperationException("Vector lengths do not match.")
			End If
			Return a.Select(Of Double)(Function(x, i) x + b(i)).ToList()
		End Function

		Public Shared Function Add(a As Double, b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			If _
			b.First.DataType = StockPriceDataType.RawPrice Then
				'already in log space
				Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
			End If

			Return b.Select(Function(x)
												'make a new copy of x to avoid modifying the original StockPriceVol instance,
												'ensuring thread safety
												Dim y = New StockPriceVol(x)
												y.Open = x.Open + a
												y.High = x.High + a
												y.Low = x.Low + a
												y.Last = x.Last + a
												SetHighLow(y)
												Return y
											End Function).ToList()
		End Function

		Public Shared Function Add(a As List(Of StockPriceVol), b As Double) As List(Of StockPriceVol)
			Return Add(b, a)
		End Function

		Public Shared Function Add(a As List(Of StockPriceVol), b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			'try to make the list teh same length if the count difference is only 1 and the shorter list
			'has a added the last value for the missing element, otherwise throw an exception
			If _
				a.First.DataType = StockPriceDataType.RawPrice OrElse
				b.First.DataType = StockPriceDataType.RawPrice Then
				'already in log space
				Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
			End If
			'try to fix the count if only one element is missing and the last element of the shorter list
			'is the same as the last element of the longer list, otherwise throw an exception
			Dim CountDifference As Integer = (a.Count - b.Count)
			Select Case CountDifference
				Case = 0
				Case 1 To 2
					'missing one or more element in b
					'make a copy of the list to avoid modifying the original list, ensuring thread safety
					b = b.ToList
					'add the last element of the list b to the bCopy to make it the same length as the a list
					'also make sure the date correspond to the last element of the a list to maintain the correct date alignment for the addition operation
					For I As Integer = 1 To CountDifference
						'how to add one day to a date in VB.NET? we can use the AddDays method of the DateTime class
						'may be better in fact not to touch the date and just use the last date of the a list, because we are not sure if the last date
						'of the a list is a trading day or not. The difference in point is usually related to the exchange calendar and the trading days, so we will just use the last date of the a list
						'to avoid any potential issues with non-trading days.
						'so keep the saame last day and just add the last value of the b list to the bCopy to make it the same length as the a list
						'also place the volume to zero, because we don't have any volume data for the missing days, and we don't want
						'to introduce any bias in the calculation.
						b.Add(New StockPriceVol(b.Last) With {
							.DateDay = a.Last.DateDay,
							.Open = .Last,
							.High = .Last,
							.Low = .Last,
							.OpenNext = .Last,
							.LastPrevious = .Last,
							.Volume = 0})
					Next
				Case -1, -2
					a = a.ToList
					For I As Integer = 1 To -CountDifference
						a.Add(New StockPriceVol(a.Last) With {
							.DateDay = b.Last.DateDay,
							.Open = .Last,
							.High = .Last,
							.Low = .Last,
							.OpenNext = .Last,
							.LastPrevious = .Last,
							.Volume = 0})
					Next
				Case Else
					Throw New InvalidOperationException("Vector lengths do not match.")
			End Select
			Return a.Select(Of StockPriceVol)(Function(x, i) Add(x, b(index:=i), Rho:=CORRELATION_HL_RHO_Default)).ToList()
		End Function

		Public Shared Function Add(a As StockPriceVol, b As StockPriceVol) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = a.Open + b.Open
				.Last = a.Last + b.Last
				.High = a.High + b.High
				.Low = a.Low + b.Low
			End With
			SetHighLow(z)
			Return z
		End Function

		Public Shared Function Add(a As StockPriceVol, b As StockPriceVol, Rho As Double) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = a.Open + b.Open
				.Last = a.Last + b.Last
			End With
			Dim aHighExc = Math.Max(0, Math.Exp(a.High - a.Open) - 1)
			Dim bHighExc = Math.Max(0, Math.Exp(b.High - b.Open) - 1)
			Dim zHighExcRel = Math.Sqrt(aHighExc * aHighExc + bHighExc * bHighExc + 2 * Rho * aHighExc * bHighExc)
			z.High = z.Open + Math.Log(1 + zHighExcRel)

			Dim aLowExc = Math.Max(0, Math.Exp(a.Open - a.Low) - 1)
			Dim bLowExc = Math.Max(0, Math.Exp(b.Open - b.Low) - 1)
			Dim zLowExcRel = Math.Sqrt(aLowExc * aLowExc + bLowExc * bLowExc + 2 * Rho * aLowExc * bLowExc)
			z.Low = z.Open - Math.Log(1 + zLowExcRel)

			If SetHighLow(z) = False Then
				'For debugging
				'z = z
			End If
			Return z
		End Function


#End Region
#Region "---Subtract---"
		Public Shared Function Subtract(a As Double, b As List(Of Double)) As List(Of Double)
			Return Add(a, Negate(b))
		End Function

		Public Shared Function Subtract(a As List(Of Double), b As Double) As List(Of Double)
			Return Add(a, -b)
		End Function

		Public Shared Function Subtract(a As List(Of Double), b As List(Of Double)) As List(Of Double)
			Return Add(a, Negate(b))
		End Function


		Public Shared Function Subtract(a As Double, b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			Return Add(a, Negate(b))
		End Function

		Public Shared Function Subtract(a As List(Of StockPriceVol), b As Double) As List(Of StockPriceVol)
			Return Add(a, -b)
		End Function

		Public Shared Function Subtract(a As List(Of StockPriceVol), b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			'try to make the list teh same length if the count difference is only 1 and the shorter list
			'has a added the last value for the missing element, otherwise throw an exception
			If _
				a.First.DataType = StockPriceDataType.RawPrice OrElse
				b.First.DataType = StockPriceDataType.RawPrice Then
				'already in log space
				Throw New InvalidOperationException("Adding is only supported for cumulative log return data sources.")
			End If
			'try to fix the count if only one element is missing and the last element of the shorter list
			'is the same as the last element of the longer list, otherwise throw an exception
			Dim CountDifference As Integer = (a.Count - b.Count)
			Select Case CountDifference
				Case = 0
				Case 1 To 2
					'missing one or more element in b
					'make a copy of the list to avoid modifying the original list, ensuring thread safety
					b = b.ToList
					'add the last element of the list b to the bCopy to make it the same length as the a list
					'also make sure the date correspond to the last element of the a list to maintain the correct date alignment for the addition operation
					For I As Integer = 1 To CountDifference
						'how to add one day to a date in VB.NET? we can use the AddDays method of the DateTime class
						'may be better in fact not to touch the date and just use the last date of the a list, because we are not sure if the last date
						'of the a list is a trading day or not. The difference in point is usually related to the exchange calendar and the trading days, so we will just use the last date of the a list
						'to avoid any potential issues with non-trading days.
						'so keep the saame last day and just add the last value of the b list to the bCopy to make it the same length as the a list
						'also place the volume to zero, because we don't have any volume data for the missing days, and we don't want
						'to introduce any bias in the calculation.
						b.Add(New StockPriceVol(b.Last) With {
							.DateDay = a.Last.DateDay,
							.Open = .Last,
							.High = .Last,
							.Low = .Last,
							.OpenNext = .Last,
							.LastPrevious = .Last,
							.Volume = 0})
					Next
				Case -1, -2
					a = a.ToList
					For I As Integer = 1 To -CountDifference
						a.Add(New StockPriceVol(a.Last) With {
							.DateDay = b.Last.DateDay,
							.Open = .Last,
							.High = .Last,
							.Low = .Last,
							.OpenNext = .Last,
							.LastPrevious = .Last,
							.Volume = 0})
					Next
				Case Else
						Throw New InvalidOperationException("Vector lengths do not match.")
      End Select

			Return a.Select(Of StockPriceVol)(Function(x, i) Subtract(x, b(index:=i), Rho:=CORRELATION_HL_RHO_Default)).ToList()
		End Function

		Public Shared Function Subtract(a As StockPriceVol, b As StockPriceVol) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = a.Open - b.Open
				.Last = a.Last - b.Last
				.High = a.High - b.High
				.Low = a.Low - b.Low
			End With
			SetHighLow(z)
			Return z
		End Function

		Public Shared Function Subtract(a As StockPriceVol, b As StockPriceVol, Rho As Double) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = a.Open - b.Open
				.Last = a.Last - b.Last
			End With
			Dim aHighExc = Math.Max(0, Math.Exp(a.High - a.Open) - 1)
			Dim bHighExc = Math.Max(0, Math.Exp(b.High - b.Open) - 1)
			Dim zHighExcRel = Math.Sqrt(aHighExc * aHighExc + bHighExc * bHighExc - 2 * Rho * aHighExc * bHighExc)
			z.High = z.Open + Math.Log(1 + zHighExcRel)

			Dim aLowExc = Math.Max(0, Math.Exp(a.Open - a.Low) - 1)
			Dim bLowExc = Math.Max(0, Math.Exp(b.Open - b.Low) - 1)
			Dim zLowExcRel = Math.Sqrt(aLowExc * aLowExc + bLowExc * bLowExc - 2 * Rho * aLowExc * bLowExc)
			z.Low = z.Open - Math.Log(1 + zLowExcRel)

			If SetHighLow(z) = False Then
				'For debugging
				'z = z
			End If
			Return z
		End Function
#End Region
#Region "---Negate---"
		Public Shared Function Negate(a As List(Of Double)) As List(Of Double)
			Return a.Select(Function(x) -x).ToList()
		End Function

		Public Shared Function Negate(a As List(Of StockPriceVol)) As List(Of StockPriceVol)
			If a.First.DataType = StockPriceDataType.RawPrice Then
				Throw New InvalidOperationException("Negating is only supported for cumulative log return data sources.")
			End If
			Return a.Select(Function(x) Negate(x)).ToList()
		End Function

		Public Shared Function Negate(a As StockPriceVol) As StockPriceVol
			Dim y = New StockPriceVol(a)
			y.Open = -a.Open
			y.High = -a.High
			y.Low = -a.Low
			y.Last = -a.Last
			SetHighLow(y)
			Return y
		End Function
#End Region
#Region "---Multiply---"
		Public Shared Function Multiply(a As Double, b As List(Of Double)) As List(Of Double)
			Return b.Select(Function(x) x * a).ToList()
		End Function

		Public Shared Function Multiply(a As List(Of Double), b As Double) As List(Of Double)
			Return a.Select(Function(x) x * b).ToList()
		End Function

		Public Shared Function Multiply(a As List(Of Double), b As List(Of Double)) As List(Of Double)
			If a.Count <> b.Count Then
				Throw New InvalidOperationException("Vector lengths do not match.")
			End If
			Return a.Select(Function(x, i) x * b(i)).ToList()
		End Function

		Public Shared Function Multiply(a As Double, b As List(Of StockPriceVol)) As List(Of StockPriceVol)

			Return b.Select(Of StockPriceVol)(Function(x)

																					Dim z = New StockPriceVol(x)
																					z.Open = a * x.Open
																					z.High = a * x.High
																					z.Low = a * x.Low
																					z.Last = a * x.Last
																					SetHighLow(z)
																					Return z
																				End Function).ToList()
		End Function

		Public Shared Function Multiply(a As List(Of StockPriceVol), b As Double) As List(Of StockPriceVol)
			Return Multiply(b, a)
		End Function

		Public Shared Function Multiply(a As List(Of StockPriceVol), b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			'try to fix the count if only one element is missing and the last element of the shorter list
			'is the same as the last element of the longer list, otherwise throw an exception
			Select Case (a.Count - b.Count)
				Case = 0
				Case 1
					b = b.ToList
					b.Add(New StockPriceVol(b.Last) With {
						.DateDay = a.Last.DateDay,
						.Open = .Last,
						.High = .Last,
						.Low = .Last,
						.OpenNext = .Last,
						.LastPrevious = .Last})
				Case -1
					a = a.ToList
					a.Add(New StockPriceVol(a.Last) With {
						.DateDay = b.Last.DateDay,
						.Open = .Last,
						.High = .Last,
						.Low = .Last,
						.OpenNext = .Last,
						.LastPrevious = .Last})
				Case Else
					Throw New InvalidOperationException("Vector lengths do not match.")
			End Select

			Return a.Select(Of StockPriceVol)(Function(x, i) Multiply(x, b(i))).ToList()
		End Function

		Public Shared Function Multiply(a As StockPriceVol, b As StockPriceVol) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = a.Open * b.Open
				.High = a.High * b.High
				.Low = a.Low * b.Low
				.Last = a.Last * b.Last
			End With
			SetHighLow(z)
			Return z
		End Function

#End Region
#Region "---Divide---"
		Public Shared Function Divide(a As Double, b As List(Of Double)) As List(Of Double)
			Return b.Select(Function(x) a / x).ToList()
		End Function

		Public Shared Function Divide(a As List(Of Double), b As Double) As List(Of Double)
			Return a.Select(Function(x) x / b).ToList()
		End Function

		Public Shared Function Divide(a As List(Of Double), b As List(Of Double)) As List(Of Double)
			If a.Count <> b.Count Then
				Throw New InvalidOperationException("Vector lengths do not match.")
			End If
			Return a.Select(Function(x, i) x / b(i)).ToList()
		End Function


		Public Shared Function Divide(a As Double, b As List(Of StockPriceVol)) As List(Of StockPriceVol)
			Return b.Select(Of StockPriceVol)(Function(x)

																					Dim z = New StockPriceVol(x)
																					With z
																						.Open = a / .Open
																						.High = a / .High
																						.Low = a / .Low
																						.Last = a / .Last
																					End With
																					SetHighLow(z)
																					Return z
																				End Function).ToList()
		End Function

		Public Shared Function Divide(a As List(Of StockPriceVol), b As Double) As List(Of StockPriceVol)
			If b = 0 Then
				Throw New InvalidOperationException("Division by zero is not allowed.")
			End If
			Return a.Select(Of StockPriceVol)(Function(x)
																					Dim z = New StockPriceVol(x)
																					With z
																						.Open = .Open / b
																						.High = .High / b
																						.Low = .Low / b
																						.Last = .Last / b
																					End With
																					SetHighLow(z)
																					Return z
																				End Function).ToList()
		End Function

		Public Shared Function Divide(a As List(Of StockPriceVol), b As List(Of StockPriceVol)) As List(Of StockPriceVol)

			'try to fix the count if only one element is missing and the last element of the shorter list
			'is the same as the last element of the longer list, otherwise throw an exception
			Select Case (a.Count - b.Count)
				Case = 0
				Case 1
					b = b.ToList
					b.Add(New StockPriceVol(b.Last) With {
						.DateDay = a.Last.DateDay,
						.Open = .Last,
						.High = .Last,
						.Low = .Last,
						.OpenNext = .Last,
						.LastPrevious = .Last})
				Case -1
					a = a.ToList
					a.Add(New StockPriceVol(a.Last) With {
						.DateDay = b.Last.DateDay,
						.Open = .Last,
						.High = .Last,
						.Low = .Last,
						.OpenNext = .Last,
						.LastPrevious = .Last})
				Case Else
					Throw New InvalidOperationException("Vector lengths do not match.")
			End Select

			Try
				Return a.Select(Of StockPriceVol)(Function(x, i) Divide(x, b(i))).ToList()
			Catch ex As Exception
				Throw New ArgumentOutOfRangeException(message:="An error occurred during division operation!", innerException:=ex)
			End Try
		End Function

		Public Shared Function Divide(a As StockPriceVol, b As StockPriceVol) As StockPriceVol
			'what should we do with division by zero cases?
			'for now we will just let it throw an exception, but we may want to consider
			'returning some special value or handling it differently in the future
			'also what would be the purpose of dividing two series of cumulative log gain?

			'the test is already done in the RPN calculator, so we don't need to do it here
			'If _
			'	a.DataType <> StockPriceDataType.RawPrice OrElse
			'	b.DataType <> StockPriceDataType.RawPrice Then

			'	Throw New InvalidOperationException(
			'		"Divide is only valid for RawPrice data. " &
			'		"Use subtraction for relative comparison in cumulative log-return space.")
			'End If

			' division logic make sense for raw price only...
			Dim z = New StockPriceVol(a)
			z.Open = a.Open / b.Open
			z.Last = a.Last / b.Last
			z.High = a.High / b.High
			z.Low = a.Low / b.Low
			SetHighLow(z)
			Return z
		End Function
#End Region
#Region "Math Support"
		'Create a new list of CumulativeLogReturn with the same number of elements and date layout as the input list,
		'but with all values set to zero (or a specified default value). The list created essentially represent a stream with a gain of zero across all time periods,
		'which can be useful as a baseline or starting point for various calculations and comparisons in financial analysis.
		Public Shared Function Clear(a As List(Of StockPriceVol), Optional DefaultValue As Double = 0.0) As List(Of StockPriceVol)
			Return a.Select(Function(x) Clear(x, DefaultValue)).ToList()
		End Function

		Public Shared Function Clear(a As StockPriceVol, Optional DefaultValue As Double = 0.0) As StockPriceVol
			Dim z = New StockPriceVol(a)
			With z
				.Open = DefaultValue
				.High = DefaultValue
				.Low = DefaultValue
				.Last = DefaultValue
				.OpenNext = DefaultValue
				.LastPrevious = DefaultValue
			End With
			Return z
		End Function
#End Region

		Public Shared Sub SetHighLow(ByRef a As StockPriceVol, ByRef b As StockPriceVol, ByRef z As StockPriceVol, Rho As Double)
			Dim aMid = (a.Open + a.Last) / 2
			Dim bMid = (b.Open + b.Last) / 2
			Dim zMid = (z.Open + z.Last) / 2

			'note we use the exponential space here to fix some issue of 0/0
			'this model is correct mathematically
			'So the old ratio formula Is wrong for log-gain values because it treats Mid as if it were a price level.
			'The better form
			'Dim highExc = Math.Exp(High - Mid()) - 1
			'has no division by Mid, And it correctly interpretsthe log normal price model, where the
			'High And Low are multiplicative factors of the Mid, rather than additive offsets.
			'So yes, the exponential model fixes both issues
			'mathematically correct in log space
			'avoids the 0 / 0 flat-gain problem with n/a values even when the High and Mid is zero
			Dim aHighExc = Math.Max(0, Math.Exp(a.High - aMid) - 1)
			Dim bHighExc = Math.Max(0, Math.Exp(b.High - bMid) - 1)
			Dim zHighExcRel = Math.Sqrt(aHighExc * aHighExc + bHighExc * bHighExc + 2 * Rho * aHighExc * bHighExc)
			z.High = zMid + Math.Log(1 + zHighExcRel)


			Dim aLowExc = Math.Max(0, Math.Exp(aMid - a.Low) - 1)
			Dim bLowExc = Math.Max(0, Math.Exp(bMid - b.Low) - 1)
			Dim zLowExcRel = Math.Sqrt(aLowExc * aLowExc + bLowExc * bLowExc + 2 * Rho * aLowExc * bLowExc)

			z.Low = zMid - Math.Log(1 + zLowExcRel)

			'sanity test, may be removed later, but for now we will keep it to catch any potential issues with the data
			'You can place Stop statements anywhere in procedures to suspend execution. Using the Stop statement Is similar to setting
			'a breakpoint in the code. The Stop statement suspends execution, but unlike End, it does Not close any files Or clear
			'any variables, unless it Is encountered in a compiled executable (.exe) file.
			'should not happen, but if it does, we will stop execution to catch the issue
			If SetHighLow(z, IsApplyCorrection:=False) = False Then Stop
			'If z.High <Math.Max(z.Open, z.Last) Then Stop
			'If z.Low > Math.Min(z.Open, z.Last) Then Stop
				'If z.Low > z.High Then Stop
				End Sub

		''' <summary>
		''' fixing the High and Low values. Also return a boolean indicating if the values were consistent.
		''' </summary>
		''' <param name="a"></param>
		''' <returns></returns>
		Public Shared Function SetHighLow(ByRef a As StockPriceVol, Optional IsApplyCorrection As Boolean = True) As Boolean
			Dim IsConsistent As Boolean = True
			Dim OLHigh As Double = Math.Max(a.Open, a.Last)
			Dim OLLow As Double = Math.Min(a.Open, a.Last)

			If a.High < a.Low Then
				IsConsistent = False
				If IsApplyCorrection Then
					'if the High is less than the Low, we will swap them
					Dim Temp As Double = a.High
					a.High = a.Low
					a.Low = Temp
				Else
					Return IsConsistent
				End If
			End If
			If a.High < OLHigh Then
				IsConsistent = False
				If IsApplyCorrection Then
					a.High = OLHigh
				Else
					Return IsConsistent
				End If
			End If
			If a.Low > OLLow Then
				IsConsistent = False
				If IsApplyCorrection Then
					a.Low = OLLow
				Else
					Return IsConsistent
				End If
			End If
			Return IsConsistent
		End Function

		Public Shared Sub SetNextLast(ByRef a As List(Of StockPriceVol))
			a(0).LastPrevious = a(0).Open
			For i As Integer = 0 To a.Count - 2
				a(i).OpenNext = a(i + 1).Open
				a(i + 1).LastPrevious = a(i).Last
			Next i
			With a.Last
				.OpenNext = .Last
			End With
		End Sub
	End Class
End Namespace
