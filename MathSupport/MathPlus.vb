Imports MathNet.Numerics
Imports MathNet.Numerics.RootFinding
Imports YahooAccessData.MathPlus.Filter
Imports YahooAccessData.OptionValuation
Imports YahooAccessData.MathPlus.Measure.Measure
Imports YahooAccessData.ExtensionService.Extensions
Imports System.Runtime.CompilerServices

Namespace MathPlus
#Const DebugPrediction = False
#Region "Definition"

	Public Module VectorList

		' In VB.NET, assigning `Nothing` to a value type (like Double or Integer) sets it to its default value (0.0 for Double).
		' This allows us to use Enumerable.Repeat(Of T)(Nothing, count) to efficiently create a List(Of T) filled with default values.
		' For example, Enumerable.Repeat(Of Double)(Nothing, 10) creates a list of 10 zeroes.
		' This pattern is safe and idiomatic in VB.NET, and works correctly with value types in generic functions.
		Public Function Create(Of T)(ByVal Count As Integer, Optional defaultValue As T = Nothing) As List(Of T)
			Return Enumerable.Repeat(defaultValue, count:=Count).ToList()
		End Function



		''' <summary>
		''' Calculate the following equation:
		''' ValueSource(I) = ValueSource(I) + ((Weight * (Value(I) - Mean)))
		''' </summary>
		''' <param name="ValueSource"></param>
		''' <param name="Value"></param>
		''' <param name="Weight"></param>
		''' <param name="ValueOffset"></param>
		<Extension>
		Public Sub VectorAdd(
			ByVal ValueSource As IList(Of Double),
			ByVal Value As IList(Of Double),
			ByVal Weight As Double,
			ByVal ValueOffset As Double)

			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) + (Weight * (Value(I) - ValueOffset))
			Next
		End Sub

		<Extension>
		Public Sub VectorAdd(
			ByVal ValueSource As IList(Of Double),
			ByVal Value As IList(Of Double))

			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) + Value(I)
			Next
		End Sub

		<Extension>
		Public Sub VectorAdd(ByVal ValueSource As IList(Of Double), ByVal Value As Double)
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) + Value
			Next
		End Sub

		<Extension>
		Public Sub VectorSub(ByVal ValueSource As IList(Of Double), ByVal Value As IList(Of Double))
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) - Value(I)
			Next
		End Sub

		<Extension>
		Public Sub VectorSub(ByVal ValueSource As IList(Of Double), ByVal Value As Double)
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) - Value
			Next
		End Sub

		<Extension>
		Public Sub VectorDivide(ByVal ValueSource As IList(Of Double), ByVal Value As Double)
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) / Value
			Next
		End Sub

		<Extension>
		Public Sub VectorCorrelationDivide(ByVal ValueSource As IList(Of Double), ByVal Value As Double, ByVal Mean As Double)
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ((ValueSource(I) - Mean) / Value) + Mean
			Next
		End Sub

		<Extension>
		Public Sub VectorMultiply(ByVal ValueSource As IList(Of Double), ByVal Value As Double)
			Dim I As Integer

			For I = 0 To ValueSource.Count - 1
				ValueSource(I) = ValueSource(I) * Value
			Next
		End Sub
	End Module

	Public Module General
		Public Const PI As Double = Math.PI  '3.14159265358979
		Public Const TwoPI As Double = 2 * PI
		Public Const PIOver2 As Double = PI / 2.0#
		Public Const LOGBase10 As Double = 0.434294481903252  '1 / Log(10)
		Public Const LOG10Base10 As Double = 10 * LOGBase10
		Public Const LOG20Base10 As Double = 20 * LOGBase10
		Public Const LOGBase2 As Double = 1.44269504088896    '1/Log(2)
		Public Const DEG_TO_RAD As Double = PI / 180
		Public Const RAD_TO_DEG As Double = 1 / DEG_TO_RAD
		Public Const NUMBER_WORKDAY_PER_YEAR As Integer = 260
		Public Const NUMBER_TRADINGDAY_PER_YEAR As Integer = 252
		Public Const NUMBER_TRADINGDAY_PER_MONTH As Integer = NUMBER_TRADINGDAY_PER_YEAR \ 12
		Public Const NUMBER_SECOND_PER_DAY As Integer = 24 * 3600
		Public Const STATISTICAL_SIGMA_DAILY_TO_YEARLY_RATIO As Double = 15.874507866387544   'Math.Sqrt(NUMBER_TRADINGDAY_PER_YEAR)

#Region "Friend Local function"
		''' <summary>
		''' note to test pseudo log
		''' pseudoLog10  function(x) { asinh(x/2)/log(10) }
		''' </summary>
		''' <param name="Value"></param>
		''' <param name="ValueRef"></param>
		''' <returns></returns>
		Friend Function LogPriceReturn(ByVal Value As Double, ByVal ValueRef As Double) As Double
			Dim ThisReturnLog As Double
			'filter for value less than zero
			If ValueRef <= 0 Then
				ThisReturnLog = 0
			Else
				If Value <= 0 Then
					ThisReturnLog = 0
				Else
					ThisReturnLog = Math.Log(Value / ValueRef)
					If Double.IsNaN(ThisReturnLog) Or Double.IsInfinity(ThisReturnLog) Then
						'ThisResult = Double.NaN
						ThisReturnLog = 0.0
					End If
				End If
			End If
			Return ThisReturnLog
		End Function

		'Volatility adjusted for discrete dividends European model
		''' <summary>
		''' See Haug Book on option p 369. This function cna be used to make a volatility adjustment to correct for the effect 
		''' of discrete dividend payment on the price of the option. This is normally only valid fo the Europeen type of option
		''' </summary>
		''' <param name="S">StockPrice</param>
		''' <param name="T">Time to maturity in years</param>
		''' <param name="r">Risk-free Rate</param>
		''' <param name="DividendPaymentValues">An array of dividend payment value</param>
		''' <param name="DividendTimesInYear">An array of dividend payment time in year</param>
		''' <param name="v"></param>
		''' <returns>The function return the adjusted volatility</returns>
		''' <remarks></remarks>
		Friend Function HaugHaugDividendVolatilityCorrection(
			S As Double,
			T As Double,
			r As Double,
			DividendPaymentValues() As Double,
			DividendTimesInYear() As Double,
			v As Double) As Double

			Dim SumDividends As Double
			Dim sumVolatilities As Double
			Dim n As Integer
			Dim j As Integer
			Dim i As Integer

			n = DividendPaymentValues.Length ' number of Dividend

			sumVolatilities = 0
			For j = 0 To n
				SumDividends = 0
				For i = j To n - 1
					SumDividends = SumDividends + DividendPaymentValues(i) * Math.Exp(-r * DividendTimesInYear(i))
				Next
				If j = 0 Then
					sumVolatilities = sumVolatilities + (S * v / (S - SumDividends)) ^ 2 * DividendTimesInYear(j)
				ElseIf j < n Then
					sumVolatilities = sumVolatilities + (S * v / (S - SumDividends)) ^ 2 * (DividendTimesInYear(j) - DividendTimesInYear(j - 1))
				Else
					sumVolatilities = sumVolatilities + v ^ 2 * (T - DividendTimesInYear(j - 1))
				End If
			Next
			Return Math.Sqrt(sumVolatilities / T)
		End Function

		''' <summary>
		''' See Chapter of Haug Option Pricing Book
		''' </summary>
		''' <param name="S"></param>
		''' <param name="X"></param>
		''' <param name="T"></param>
		''' <param name="r"></param>
		''' <param name="b"></param>
		''' <param name="v"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Function BSAmericanCallApprox2002(
			ByVal S As Double,
			ByVal X As Double,
			ByVal T As Double,
			ByVal r As Double,
			ByVal b As Double,
			ByVal v As Double) As Double

			Dim BInfinity As Double
			Dim B0 As Double
			Dim ht1 As Double
			Dim ht2 As Double
			Dim I1 As Double
			Dim I2 As Double
			Dim alfa1 As Double
			Dim alfa2 As Double
			Dim Beta As Double
			Dim t1 As Double
			Dim ThisResult As Double


			t1 = 1 / 2 * (Math.Sqrt(5) - 1) * T

			If b >= r Then  '// Never optimal to exersice before maturity
				ThisResult = Measure.Measure.BlackScholes(Measure.Measure.enuOptionType._Call, S, X, T, r, r - b, v)
			Else
				Beta = (1 / 2 - b / v ^ 2) + Math.Sqrt((b / v ^ 2 - 1 / 2) ^ 2 + 2 * r / v ^ 2)
				BInfinity = Beta / (Beta - 1) * X
				B0 = Max(X, r / (r - b) * X)

				ht1 = -(b * t1 + 2 * v * Math.Sqrt(t1)) * X ^ 2 / ((BInfinity - B0) * B0)
				ht2 = -(b * T + 2 * v * Math.Sqrt(T)) * X ^ 2 / ((BInfinity - B0) * B0)
				I1 = B0 + (BInfinity - B0) * (1 - Math.Exp(ht1))
				I2 = B0 + (BInfinity - B0) * (1 - Math.Exp(ht2))
				alfa1 = (I1 - X) * I1 ^ (-Beta)
				alfa2 = (I2 - X) * I2 ^ (-Beta)

				If S >= I2 Then
					ThisResult = S - X
				Else
					ThisResult =
						alfa2 * S ^ Beta -
						alfa2 * phi(S, t1, Beta, I2, I2, r, b, v) +
						phi(S, t1, 1, I2, I2, r, b, v) - phi(S, t1, 1, I1, I2, r, b, v) -
						X * phi(S, t1, 0, I2, I2, r, b, v) +
						X * phi(S, t1, 0, I1, I2, r, b, v) +
						alfa1 * phi(S, t1, Beta, I1, I2, r, b, v) -
						alfa1 * ksi(S, T, Beta, I1, I2, I1, t1, r, b, v) +
						ksi(S, T, 1, I1, I2, I1, t1, r, b, v) -
						ksi(S, T, 1, X, I2, I1, t1, r, b, v) -
						X * ksi(S, T, 0, I1, I2, I1, t1, r, b, v) +
						X * ksi(S, T, 0, X, I2, I1, t1, r, b, v)
				End If
			End If
			Return ThisResult
		End Function
		Private Function phi(S As Double, T As Double, gamma As Double, h As Double, i As Double, r As Double, b As Double, v As Double) As Double
			Dim lambda As Double, kappa As Double
			Dim d As Double

			lambda = (-r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2) * T
			d = -(Math.Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T) / (v * Math.Sqrt(T))
			kappa = 2 * b / v ^ 2 + 2 * gamma - 1
			Return Math.Exp(lambda) * S ^ gamma * (CND(d) - (i / S) ^ kappa * CND(d - 2 * Math.Log(i / S) / (v * Math.Sqrt(T))))
		End Function

		' Muligens forskjellig fra phi i Bjerksun Stensland 1993
		Private Function phi2(S As Double, T2 As Double, gamma As Double, h As Double, i As Double, r As Double, b As Double, v As Double) As Double
			Dim lambda As Double, kappa As Double
			Dim d As Double, d2 As Double

			lambda = -r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2
			kappa = 2 * b / v ^ 2 + 2 * gamma - 1

			d = (Math.Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))
			d2 = (Math.Log(i ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))

			Return Math.Exp(lambda * T2) * S ^ gamma * (CND(-d) - (i / S) ^ kappa * CND(-d2))
		End Function

		Private Function ksi(S As Double, T2 As Double, gamma As Double, h As Double, I2 As Double, I1 As Double, t1 As Double, r As Double, b As Double, v As Double) As Double
			Dim e1 As Double, e2 As Double, e3 As Double, e4 As Double
			Dim f1 As Double, f2 As Double, f3 As Double, f4 As Double
			Dim rho As Double, kappa As Double, lambda As Double

			e1 = (Math.Log(S / I1) + (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Math.Sqrt(t1))
			e2 = (Math.Log(I2 ^ 2 / (S * I1)) + (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Math.Sqrt(t1))
			e3 = (Math.Log(S / I1) - (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Math.Sqrt(t1))
			e4 = (Math.Log(I2 ^ 2 / (S * I1)) - (b + (gamma - 0.5) * v ^ 2) * t1) / (v * Math.Sqrt(t1))

			f1 = (Math.Log(S / h) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))
			f2 = (Math.Log(I2 ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))
			f3 = (Math.Log(I1 ^ 2 / (S * h)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))
			f4 = (Math.Log(S * I1 ^ 2 / (h * I2 ^ 2)) + (b + (gamma - 0.5) * v ^ 2) * T2) / (v * Math.Sqrt(T2))

			rho = Math.Sqrt(t1 / T2)
			lambda = -r + gamma * b + 0.5 * gamma * (gamma - 1) * v ^ 2
			kappa = 2 * b / (v ^ 2) + (2 * gamma - 1)

			Return Math.Exp(lambda * T2) * S ^ gamma * (CBND(-e1, -f1, rho) -
						(I2 / S) ^ kappa * CBND(-e2, -f2, rho) -
						(I1 / S) ^ kappa * CBND(-e3, -f3, -rho) + (I1 / I2) ^ kappa * CBND(-e4, -f4, -rho))
		End Function


		''' <summary>
		''' Cummulative Normal Distribution double precision algorithm based on Hart 1968
		''' Based on implementation by Graeme West
		''' See chapter 13 of book
		''' </summary>
		''' <param name="X"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function CND(X As Double) As Double
			Dim y As Double, Exponential As Double, SumA As Double, SumB As Double

			y = Math.Abs(X)
			If y > 37 Then
				CND = 0
			Else
				Exponential = Math.Exp(-y ^ 2 / 2)
				If y < 7.07106781186547 Then
					SumA = 0.0352624965998911 * y + 0.700383064443688
					SumA = SumA * y + 6.37396220353165
					SumA = SumA * y + 33.912866078383
					SumA = SumA * y + 112.079291497871
					SumA = SumA * y + 221.213596169931
					SumA = SumA * y + 220.206867912376
					SumB = 0.0883883476483184 * y + 1.75566716318264
					SumB = SumB * y + 16.064177579207
					SumB = SumB * y + 86.7807322029461
					SumB = SumB * y + 296.564248779674
					SumB = SumB * y + 637.333633378831
					SumB = SumB * y + 793.826512519948
					SumB = SumB * y + 440.413735824752
					CND = Exponential * SumA / SumB
				Else
					SumA = y + 0.65
					SumA = y + 4 / SumA
					SumA = y + 3 / SumA
					SumA = y + 2 / SumA
					SumA = y + 1 / SumA
					CND = Exponential / (SumA * 2.506628274631)
				End If
			End If

			If X > 0 Then CND = 1 - CND
		End Function

		''' <summary>
		''' The cumulative bivariate normal distribution function
		''' </summary>
		''' <param name="X"></param>
		''' <param name="y"></param>
		''' <param name="rho"></param>
		''' <returns></returns>
		''' <remarks>
		''' A function for computing bivariate normal probabilities.
		'''       Alan Genz
		'''   Department of Mathematics
		'''   Washington State University
		'''   Pullman, WA 99164-3113
		'''   Email : alangenz@wsu.edu
		'''
		'''   This function is based on the method described by
		'''   Drezner, Z and G.O. Wesolowsky, (1990),
		'''   On the computation of the bivariate normal integral,
		'''   Journal of Statist. Comput. Simul. 35, pp. 101-107,
		'''   with major modifications for double precision, and for |R| close to 1.
		'''   This code was originally transelated into VBA by Graeme West
		''' </remarks>
		Private Function CBND(X As Double, y As Double, rho As Double) As Double


			Dim i As Integer, ISs As Integer, LG As Integer, NG As Integer
			Dim XX(10, 3) As Double, W(10, 3) As Double
			Dim h As Double, k As Double, hk As Double, hs As Double, BVN As Double, Ass As Double, asr As Double, sn As Double
			Dim A As Double, b As Double, bs As Double, c As Double, d As Double
			Dim xs As Double, rs As Double

			W(1, 1) = 0.17132449237917
			XX(1, 1) = -0.932469514203152
			W(2, 1) = 0.360761573048138
			XX(2, 1) = -0.661209386466265
			W(3, 1) = 0.46791393457269
			XX(3, 1) = -0.238619186083197

			W(1, 2) = 0.0471753363865118
			XX(1, 2) = -0.981560634246719
			W(2, 2) = 0.106939325995318
			XX(2, 2) = -0.904117256370475
			W(3, 2) = 0.160078328543346
			XX(3, 2) = -0.769902674194305
			W(4, 2) = 0.203167426723066
			XX(4, 2) = -0.587317954286617
			W(5, 2) = 0.233492536538355
			XX(5, 2) = -0.36783149899818
			W(6, 2) = 0.249147045813403
			XX(6, 2) = -0.125233408511469

			W(1, 3) = 0.0176140071391521
			XX(1, 3) = -0.993128599185095
			W(2, 3) = 0.0406014298003869
			XX(2, 3) = -0.963971927277914
			W(3, 3) = 0.0626720483341091
			XX(3, 3) = -0.912234428251326
			W(4, 3) = 0.0832767415767048
			XX(4, 3) = -0.839116971822219
			W(5, 3) = 0.10193011981724
			XX(5, 3) = -0.746331906460151
			W(6, 3) = 0.118194531961518
			XX(6, 3) = -0.636053680726515
			W(7, 3) = 0.131688638449177
			XX(7, 3) = -0.510867001950827
			W(8, 3) = 0.142096109318382
			XX(8, 3) = -0.37370608871542
			W(9, 3) = 0.149172986472604
			XX(9, 3) = -0.227785851141645
			W(10, 3) = 0.152753387130726
			XX(10, 3) = -0.0765265211334973

			If Math.Abs(rho) < 0.3 Then
				NG = 1
				LG = 3
			ElseIf Math.Abs(rho) < 0.75 Then
				NG = 2
				LG = 6
			Else
				NG = 3
				LG = 10
			End If

			h = -X
			k = -y
			hk = h * k
			BVN = 0

			If Math.Abs(rho) < 0.925 Then
				If Math.Abs(rho) > 0 Then
					hs = (h * h + k * k) / 2
					asr = Math.Asin(rho)
					For i = 1 To LG
						For ISs = -1 To 1 Step 2
							sn = Math.Sin(asr * (ISs * XX(i, NG) + 1) / 2)
							BVN = BVN + W(i, NG) * Math.Exp((sn * hk - hs) / (1 - sn * sn))
						Next ISs
					Next i
					BVN = BVN * asr / (4 * PI)
				End If
				BVN = BVN + CND(-h) * CND(-k)
			Else
				If rho < 0 Then
					k = -k
					hk = -hk
				End If
				If Math.Abs(rho) < 1 Then
					Ass = (1 - rho) * (1 + rho)
					A = Math.Sqrt(Ass)
					bs = (h - k) ^ 2
					c = (4 - hk) / 8
					d = (12 - hk) / 16
					asr = -(bs / Ass + hk) / 2
					If asr > -100 Then BVN = A * Math.Exp(asr) * (1 - c * (bs - Ass) * (1 - d * bs / 5) / 3 + c * d * Ass * Ass / 5)
					If -hk < 100 Then
						b = Math.Sqrt(bs)
						BVN = BVN - Math.Exp(-hk / 2) * Math.Sqrt(2 * PI) * CND(-b / A) * b * (1 - c * bs * (1 - d * bs / 5) / 3)
					End If
					A = A / 2
					For i = 1 To LG
						For ISs = -1 To 1 Step 2
							xs = (A * (ISs * XX(i, NG) + 1)) ^ 2
							rs = Math.Sqrt(1 - xs)
							asr = -(bs / xs + hk) / 2
							If asr > -100 Then
								BVN = BVN + A * W(i, NG) * Math.Exp(asr) * (Math.Exp(-hk * (1 - rs) / (2 * (1 + rs))) / rs - (1 + c * xs * (1 + d * xs)))
							End If
						Next ISs
					Next i
					BVN = -BVN / (2 * PI)
				End If
				If rho > 0 Then
					BVN = BVN + CND(-Max(h, k))
				Else
					BVN = -BVN
					If k > h Then BVN = BVN + CND(k) - CND(h)
				End If
			End If
			Return BVN
		End Function

		''' <summary>
		''' Equivalent function of Excel Application.Max
		''' </summary>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Function Max(ByVal X As Double, ByVal Y As Double) As Double
			'newval = IIf(val1 - val2 < 0, 0, val1 - val2)
			Max = X
			If (Y > Max) Then Max = Y
		End Function
#End Region
	End Module
#End Region
#Region "Phase Measurement"
	Public Class PhaseMeasurement
		Private Shared Function PhaseSubDegLocal(ByVal PhaseRef As Double, ByVal Phase As Double) As Double
			'the result is always between +-180 degrees
			Dim DeltaPhase As Double

			DeltaPhase = PhaseRef - Phase
			If DeltaPhase > 180.0# Then
				DeltaPhase = DeltaPhase - 360.0#
			ElseIf DeltaPhase < -180.0# Then
				DeltaPhase = DeltaPhase + 360.0#
			End If
			PhaseSubDegLocal = DeltaPhase
		End Function

		''' <summary>
		''' Calculate the difference in degrees between between a PhaseRef and a Phase value (PhaseRef-Phase). The returned output
		''' is always between +- 180 degrees
		''' </summary>
		''' <param name="PhaseRef">in degrees</param>
		''' <param name="Phase">in degrees</param>
		''' <returns>The result between +- 180 degrees</returns>
		''' <remarks></remarks>
		Public Shared Function PhaseSub(ByVal PhaseRef As Integer, ByVal Phase As Integer) As Integer
			'the result is always between +-180 degrees
			Dim DeltaPhase As Integer

			DeltaPhase = PhaseRef - Phase
			If DeltaPhase > 180 Then
				DeltaPhase = DeltaPhase - 360
			ElseIf DeltaPhase < -180 Then
				DeltaPhase = DeltaPhase + 360
			End If
			PhaseSub = DeltaPhase
		End Function

		''' <summary>
		''' Clip a phase input in degree between 0 to 360 degree
		''' </summary>
		''' <param name="Phase"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Private Shared Function PhaseClip(ByVal Phase As Integer) As Integer
			'bound the output between 0 and 360 degree
			Phase = Phase Mod 360
			If Phase < 0 Then
				Phase = Phase + 360
			End If
			PhaseClip = Phase
		End Function

		''' <summary>
		''' Calculate the deviation rms from a PhaseRef from an array of phase measurement 
		''' </summary>
		''' <param name="PhaseMeasure"></param>
		''' <param name="PhaseRef"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function PhaseDevRMS(ByRef PhaseMeasure() As Integer, ByVal PhaseRef As Integer) As Double
			Dim I As Integer
			Dim NPoint As Integer
			Dim Sum2 As Double
			Dim Result As Double

			NPoint = PhaseMeasure.Length
			Sum2 = 0
			For I = 0 To NPoint - 1
				Sum2 = Sum2 + (PhaseSub(PhaseMeasure(I), PhaseRef)) ^ 2
			Next
			Result = Math.Sqrt(Sum2 / NPoint)
			PhaseDevRMS = Result
		End Function

		''' <summary>
		''' Calculate the phase average from an array of phase measurement. The average phase output
		''' can optionally be rotated by a certain amount in degree.
		''' </summary>
		''' <param name="PhaseMeasure"></param>
		''' <param name="RotationDeg"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function PhaseAverage(ByRef PhaseMeasure() As Integer, Optional ByVal RotationDeg As Integer = 0) As Double
			Dim I As Integer
			Dim NPoint As Integer
			Dim SumX As Double
			Dim SumY As Double

			NPoint = PhaseMeasure.Length
			For I = 0 To NPoint - 1
				SumX = SumX + Math.Cos(DEG_TO_RAD * PhaseMeasure(I))
				SumY = SumY + Math.Sin(DEG_TO_RAD * PhaseMeasure(I))
			Next
			PhaseAverage = ArcTanForDouble(SumX, SumY, RotationDeg)
		End Function

		''' <summary>
		''' Calculate the phase Statistic from an array
		''' </summary>
		''' <param name="PhaseMeasure">The array in degrees</param>
		''' <param name="PhaseMean">The output Phase mean of th array</param>
		''' <param name="PhaseDevRMS">The output phase deviation RMS</param>
		''' <param name="PhaseRef">The phase reference point for the calculation</param>
		''' <remarks>All parameters are in degrees</remarks>
		Public Shared Sub PhaseStatistic(ByRef PhaseMeasure() As Integer, ByRef PhaseMean As Double, ByRef PhaseDevRMS As Double, Optional ByVal PhaseRef As Integer = 0)
			Dim I As Integer
			Dim NPoint As Integer
			Dim SumX As Double
			Dim SumY As Double
			Dim Sum2 As Double

			NPoint = PhaseMeasure.Length
			For I = 0 To NPoint - 1
				SumX = SumX + Math.Cos(DEG_TO_RAD * PhaseMeasure(I))
				SumY = SumY + Math.Sin(DEG_TO_RAD * PhaseMeasure(I))
				Sum2 = Sum2 + (PhaseSub(PhaseMeasure(I), PhaseRef)) ^ 2
			Next
			If NPoint > 0 Then
				PhaseMean = PhaseSubDegLocal(ArcTanForDouble(SumX, SumY), 0)
				PhaseDevRMS = Math.Sqrt(Sum2 / NPoint)
			Else
				PhaseMean = 0
				PhaseDevRMS = 0
			End If
		End Sub

		''' <summary>
		''' Calculate the arc tangente over 0 to 360 degrees. A rotation angle can optionally be applied after the calculation.
		''' </summary>
		''' <param name="X"></param>
		''' <param name="Y"></param>
		''' <param name="RotationDeg"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Shared Function ArcTan(ByVal X As Double, ByVal Y As Double, Optional ByVal RotationDeg As Integer = 0) As Integer
			Dim Pxy As Integer

			If X = 0 Then
				If Y > 0 Then
					Pxy = 90
				ElseIf Y < 0 Then
					Pxy = 270
				End If
			Else
				Pxy = CInt(RAD_TO_DEG * Math.Atan(Y / X))
				If X > 0 Then
					If Y < 0 Then Pxy = 360 + Pxy
				Else
					Pxy = 180 + Pxy
					If Pxy >= 360 Then Pxy = Pxy - 360
				End If
			End If
			If RotationDeg <> 0 Then
				Return PhaseClip(PhaseSub(Pxy, RotationDeg))
			Else
				Return Pxy
			End If
		End Function

		'Output the Arc Tangent in degrees from 0 to 360 degrees.
		Public Shared Function ArcTanForDouble(X As Double, Y As Double, Optional ByVal RotationDeg As Double = 0) As Double
			Dim Pxy As Double

			If X = 0 Then
				If Y > 0 Then
					Pxy = 90.0#
				ElseIf Y < 0 Then
					Pxy = 270.0#
				End If
			Else
				Pxy = RAD_TO_DEG * Math.Atan(Y / X)
				If X > 0 Then
					If Y < 0 Then Pxy = 360.0# + Pxy
				Else
					Pxy = 180.0# + Pxy
					If Pxy >= 360.0# Then Pxy = Pxy - 360.0#
				End If
			End If
			If RotationDeg <> 0 Then
				Return PhaseClipForDouble(PhaseSubForDouble(Pxy, RotationDeg))
			Else
				Return Pxy
			End If
		End Function

		Public Shared Function PhaseClipForDouble(ByVal Phase As Double) As Double
			'bound the output between 0 and 360 degree
			Phase = Phase Mod 360.0#
			If Phase < 0.0# Then
				Phase = Phase + 360.0#
			End If
			PhaseClipForDouble = Phase
		End Function

		Public Shared Function PhaseSubForDouble(ByVal PhaseRef As Double, ByVal Phase As Double) As Double
			'the result is always between +-180 degrees
			Dim DeltaPhase As Double

			DeltaPhase = PhaseRef - Phase
			If DeltaPhase > 180.0# Then
				DeltaPhase = DeltaPhase - 360.0#
			ElseIf DeltaPhase < -180 Then
				DeltaPhase = DeltaPhase + 360.0#
			End If
			PhaseSubForDouble = DeltaPhase
		End Function
	End Class
#End Region
#Region "WaveForm"
	Public Class WaveForm
		Public Shared Function SignalScale(ByRef Value As Double(), ByVal Slope As Double, ByVal Offset As Double) As Double()
			Dim I As Integer

			For I = 0 To Value.Length - 1
				Value(I) = Slope * Value(I) + Offset
			Next
			Return Value
		End Function


		''' <summary>
		''' This function is part of the family of the standard sigmoid limiting function. The particularity 
		''' of this implementation is to always maintain a unit slope at around zero and to extend the maximum output range to 
		''' +-Scale while maintaining a symetric output around zero. La compression de gain est de 0.76 at MaxScale input and of
		''' 0.48 for an input of twice the MaxScale.
		''' </summary>
		''' <param name="Value">The input value to be limited</param>
		''' <param name="MaxScale">The maximum range output of the function</param>
		''' <returns>the limited output value between +-MaxScale</returns>
		''' <remarks>See the file Sigmoid Simulation.xlsx for evaluation</remarks>
		Public Shared Function SignalLimit(ByVal Value As Double, ByVal MaxScale As Double) As Double
			If MaxScale = 0 Then
				Return 0.0
			Else
				Return (2 * MaxScale / (1 + (System.Math.Exp(-2 * Value / MaxScale)))) - MaxScale
			End If
		End Function

		''' <summary>
		''' This function is part of the family of the standard sigmoid limiting function. The particularity 
		''' of this implementation is to always maintain a unit slope at around Offset and to extend the maximum output range to 
		''' Offset +-Scale while maintaining a symetric output around the Offset. La compression de gain est de 0.76 at MaxScale input and of
		''' 0.48 for an input of twice the MaxScale relative to the offset.
		''' </summary>
		''' <param name="Value">The input value to be limited</param>
		''' <param name="MaxScale">The maximum range output of the function</param>
		''' <param name="Offset">The center of the function</param>
		''' <returns>the limited output value between Offset +-MaxScale</returns>
		''' <remarks>See the file Sigmoid Simulation.xlsx for evaluation</remarks>
		Public Shared Function SignalLimit(ByVal Value As Double, ByVal MaxScale As Double, ByVal Offset As Double) As Double
			Return SignalLimit((Value - Offset), MaxScale) + Offset
		End Function


		''' <summary>
		''' This function is part of the family of the standard sigmoid limiting function. The particularity 
		''' of this implementation is to always maintain a unit slope at around Offset and to extend the maximum output range from Offset -MinScale to
		''' Offset +MaxScale while maintaining a symetric output around the Offset. 
		''' </summary>
		''' <param name="Value">The input value to be limited</param>
		''' <param name="MinScale">The minimum range output of the function around the offset expressed as a positive number</param>
		''' <param name="MaxScale">The maximum range output of the function around the offset</param>
		''' <param name="Offset">The center of the function</param>
		''' <returns>the limited output value from Offset - MinScale to Offset + MaxScale</returns>
		''' <remarks>See the file Sigmoid Simulation.xlsx for evaluation</remarks>
		Public Shared Function SignalLimit(ByVal Value As Double, ByVal MinScale As Double, ByVal MaxScale As Double, ByVal Offset As Double) As Double
			If Value >= Offset Then
				Return SignalLimit((Value - Offset), MaxScale) + Offset
			Else
				Return SignalLimit((Value - Offset), MinScale) + Offset
			End If
		End Function

		''' <summary>
		''' Return an sinusoidal array with the specific characterictic given by x(k)=A*Sin(PI*Fdigital*k + Phase)
		''' </summary>
		''' <param name="Amplitude"></param>
		''' Maximum amplitude level A
		''' <param name="NumberCycles">
		''' Represent the number of cycle over the specified number of samples defined. It is related to Fd, the Digital Frequency
		''' by Fd=2*(NumberCycles)/(NumberSamples-1)
		''' </param>
		''' <param name="PhaseDeg"></param>
		''' The offset phase of the sinusoide 
		''' <param name="NumberSamples"></param>
		''' The total number of samples for the sinusoide 
		''' <returns>the simusoide table in an array</returns>
		''' <remarks></remarks>
		Public Shared Function Sinus(ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double, ByVal NumberSamples As Integer) As Double()
			Dim ThisSinus(0 To NumberSamples - 1) As Double
			Dim PhaseRad As Double = MathPlus.DEG_TO_RAD * PhaseDeg
			Dim I As Integer

			Dim OmegaFreqDig As Double = Math.PI * ((2 * NumberCycles) / (NumberSamples - 1))

			For I = 0 To NumberSamples - 1
				ThisSinus(I) = Amplitude * Math.Sin(OmegaFreqDig * I + PhaseRad)
			Next
			Return ThisSinus
		End Function

		''' <summary>
		''' Return an sinusoidal array with the specific characterictic given by x(k)=A*Sin(PI*Fdigital*k + Phase)
		''' </summary>
		''' <param name="Amplitude"></param>
		''' Maximum amplitude level A
		''' <param name="NumberCycles">
		''' Represent the number of cycle over the specified number of samples defined. It is related to Fd, the Digital Frequency
		''' by Fd=2*(NumberCycles)/(NumberSamples-1)
		''' </param>
		''' <param name="PhaseDeg"></param>
		''' The offset phase of the sinusoide 
		''' <param name="NumberSamples"></param>
		''' The total number of samples for the sinusoide 
		''' <returns>the simusoide table in an array</returns>
		''' <remarks></remarks>
		Public Shared Function Sinus(ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double, ByVal NumberSamples As Integer, ByVal Mean As Double) As Double()
			Dim ThisSinus(0 To NumberSamples - 1) As Double
			Dim PhaseRad As Double = MathPlus.DEG_TO_RAD * PhaseDeg
			Dim I As Integer

			Dim OmegaFreqDig As Double = Math.PI * ((2 * NumberCycles) / (NumberSamples - 1))

			For I = 0 To NumberSamples - 1
				ThisSinus(I) = Amplitude * Math.Sin(OmegaFreqDig * I + PhaseRad) + Mean
			Next
			Return ThisSinus
		End Function

		Public Shared Function Cosinus(ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double, ByVal NumberSamples As Integer, ByVal Mean As Double) As Double()
			Return MathPlus.WaveForm.Sinus(Amplitude, NumberCycles, PhaseDeg + 90, NumberSamples, Mean)
		End Function

		Public Shared Function Cosinus(ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double, ByVal NumberSamples As Integer) As Double()
			Return MathPlus.WaveForm.Sinus(Amplitude, NumberCycles, PhaseDeg + 90, NumberSamples)
		End Function

		''' <summary>
		''' Return an Square waveform array with the specific characterictic given by x(k)=A*Sign(Sin(PI*Fdigital*k + Phase))
		''' </summary>
		''' <param name="Amplitude"></param>
		''' Maximum amplitude level A
		''' <param name="NumberCycles">
		''' Represent the number of cycle over the specified number of samples defined. It is related to Fd, the Digital Frequency
		''' by Fd=2*(NumberCycles)/(NumberSamples-1)
		''' </param>
		''' <param name="PhaseDeg"></param>
		''' The offset phase of the sinusoide 
		''' <param name="NumberSamples"></param>
		''' The total number of samples for the sinusoide 
		''' <returns>the simusoide table in an array</returns>
		''' <remarks></remarks>
		Public Shared Function Square(ByVal Amplitude As Double, ByVal NumberCycles As Double, ByVal PhaseDeg As Double, ByVal NumberSamples As Integer) As Double()
			Dim ThisSquare(0 To NumberSamples - 1) As Double
			Dim PhaseRad As Double = MathPlus.DEG_TO_RAD * PhaseDeg
			Dim I As Integer

			Dim OmegaFreqDig As Double = Math.PI * ((2 * NumberCycles) / (NumberSamples - 1))
			For I = 0 To NumberSamples - 1
				Select Case Math.Sin(OmegaFreqDig * I + PhaseRad)
					Case Is > 0
						ThisSquare(I) = Amplitude
					Case Is < 0
						ThisSquare(I) = -Amplitude
				End Select
			Next
			Return ThisSquare
		End Function
	End Class
#End Region
#Region "Filter"
	Namespace Filter
#Region "FilterModule"

#Disable Warning BC42300 ' XML comment block must immediately precede the language element to which it applies
#Disable Warning BC42300 ' XML comment block must immediately precede the language element to which it applies
		''' <summary>
		''' Not fully implemented yet
		''' </summary>
		''' <remarks></remarks>
		'Public Module FilterGenerate
		'  Public Function CreateFilter(ByVal FilterType As IFilterType.enuFilterType) As IFilter
		'    Select Case FilterType
		'      Case IFilterType.enuFilterType.LowPassExp
		'        Dim ThisDialog As New DialogFilterBasic("Filter Lowpass Exponential")
		'        If ThisDialog.ShowDialog = Windows.Forms.DialogResult.OK Then
		'          Return New FilterLowPassExp(ThisDialog.Rate)
		'        Else
		'          Return Nothing
		'        End If
		'      Case Else
		'        Return Nothing
		'    End Select
		'  End Function
		'End Module
#End Region
#Region "FilterHighPassExp"
		<Serializable()>
		Public Class FilterHighPassExp
#Enable Warning BC42300 ' XML comment block must immediately precede the language element to which it applies
#Enable Warning BC42300 ' XML comment block must immediately precede the language element to which it applies
			Private MyFilterLast As Double
			Private MyFilterLP As FilterLowPassExp
			Private IsValueInitial As Boolean
			Private MyListOfValue As ListScaled

			Public Sub New(ByVal FilterRate As Double)
				MyListOfValue = New ListScaled
				MyFilterLP = New FilterLowPassExp(FilterRate)
			End Sub

			'Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Double)
			'  MyFilterLP = New FilterLowPassExp(FilterRate, ValueInitial)
			'  MyFilterLast = ValueInitial
			'End Sub

			'Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Single)
			'  MyFilterLP = New FilterLowPassExp(FilterRate, ValueInitial)
			'End Sub

			Public Function Filter(ByVal Value As Double) As Double
#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual <> ThisResultPrediction Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
				MyFilterLast = Value - MyFilterLP.Filter(Value)
				MyListOfValue.Add(MyFilterLast)
				Return MyFilterLast
			End Function

			Public Function Filter(ByRef Value() As Double) As Double()
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function Filter(ByVal Value As Single) As Double
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function FilterPredictionNext(ByVal Value As Double) As Double
				Return Value - MyFilterLP.FilterPredictionNext(Value)
			End Function

			Public Function FilterPredictionNext(ByVal Value As Single) As Double
				Return Me.FilterPredictionNext(CDbl(Value))
			End Function

			Public Function FilterLast() As Double
				Return MyFilterLast
			End Function

			Public Function Last() As Double
				Return MyFilterLP.Last
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return MyFilterLP.Rate
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double)
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double()
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterRateInterpolated"
		<Serializable()>
		Public Class FilterRateInterpolated
			Implements IFilterRateInterpolated


			Private MyListOfPriceVol As List(Of IList(Of Double))
			Private MyListOfOutputType As List(Of IList(Of Double))
			Private MyOutputType As IFilterRateInterpolated.enuOutputType
			Private MyRateMin As Integer
			Private MyRateMax As Integer
			Private MyFilterRange As Double
			Private MySelectedItem As Integer
			Private MyFilterRate As Double

			Public Sub New(
				ByVal Filter As IFilter,
				ByVal MinRate As Integer,
				ByVal MaxRate As Integer,
				ByVal OutputType As IFilterRateInterpolated.enuOutputType)

				Dim ThisFilterControl As IFilterControl
				Dim ThisFilterPrediction As IFilterPrediction = Nothing
				Dim I As Integer

				MyOutputType = OutputType
				MyListOfPriceVol = New List(Of IList(Of Double))

				If Not TypeOf Filter Is IFilterControl Then Throw New NotSupportedException("IFilterControl interface is not supported...")
				If TypeOf Filter Is IFilterPrediction Then
					ThisFilterPrediction = DirectCast(Filter, IFilterPrediction)
				End If
				ThisFilterControl = DirectCast(Filter, IFilterControl)

				MyRateMin = MinRate
				MyRateMax = MaxRate
				MyFilterRange = MyRateMax - MyRateMin
				If MyFilterRange <= 0 Then Throw New InvalidConstraintException("Invalid filter range...")
				If ThisFilterControl.IsInputEnabled = False Then Throw New NotSupportedException("Filter has no input data...")
				Select Case MyOutputType
					Case IFilterRateInterpolated.enuOutputType.Standard
						For I = MyRateMin To MyRateMax
							ThisFilterControl.Refresh(I)
							MyListOfPriceVol.Add(Filter.ToList)
						Next
					Case IFilterRateInterpolated.enuOutputType.GainPerYear
						If ThisFilterPrediction Is Nothing Then Throw New NotSupportedException("Filter prediction has no data...")
						For I = MyRateMin To MyRateMax
							ThisFilterControl.Refresh(I)
							MyListOfPriceVol.Add(ThisFilterPrediction.ToListOfGainPerYear)
						Next
					Case IFilterRateInterpolated.enuOutputType.GainPerYearDerivative
						If ThisFilterPrediction Is Nothing Then Throw New NotSupportedException("Filter prediction derivative has no data...")
						For I = MyRateMin To MyRateMax
							ThisFilterControl.Refresh(I)
							MyListOfPriceVol.Add(ThisFilterPrediction.ToListOfGainPerYearDerivative)
						Next
					Case IFilterRateInterpolated.enuOutputType.NotDefined
						Throw New NotSupportedException("Invalid Filter output type...")
				End Select
				Me.Refresh(Filter.Rate)
			End Sub

			Friend Sub New(
				ByVal ListOfFilterValue As List(Of IList(Of Double)),
				ByVal Rate As Integer,
				ByVal MinRate As Integer,
				ByVal OutputType As IFilterRateInterpolated.enuOutputType)


				MyOutputType = OutputType
				MyListOfPriceVol = New List(Of IList(Of Double))(ListOfFilterValue)
				MyRateMin = MinRate
				MyRateMax = MyRateMin + MyListOfPriceVol.Count - 1
				MyFilterRange = MyRateMax - MyRateMin
				Me.Refresh(Rate)
			End Sub

			Public Sub New(
				ByVal FilterValue As IEnumerable(Of IList(Of Double)),
				ByVal Rate As Integer,
				ByVal MinRate As Integer,
				ByVal OutputType As IFilterRateInterpolated.enuOutputType)


				MyOutputType = IFilterRateInterpolated.enuOutputType.NotDefined
				MyListOfPriceVol = New List(Of IList(Of Double))(FilterValue)
				MyRateMin = MinRate
				MyRateMax = MyRateMin + MyListOfPriceVol.Count - 1
				MyFilterRange = MyRateMax - MyRateMin
				Me.Refresh(Rate)
			End Sub

			Public ReadOnly Property Rate As Double Implements IFilterRateInterpolated.Rate
				Get
					Return MyFilterRate
				End Get
			End Property

			Public ReadOnly Property RateMinimum As Double Implements IFilterRateInterpolated.RateMinimum
				Get
					Return MyRateMin
				End Get
			End Property

			Public ReadOnly Property RateMaximum As Double Implements IFilterRateInterpolated.RateMaximum
				Get
					Return MyRateMax
				End Get
			End Property

			Public Sub Refresh(ByVal FilterRate As Double) Implements IFilterRateInterpolated.Refresh
				'interpolate the result
				If FilterRate < MyRateMin Then
					FilterRate = MyRateMin
				ElseIf FilterRate > MyRateMax Then
					FilterRate = MyRateMax
				End If
				MyFilterRate = FilterRate
				'select the closest item
				MySelectedItem = CInt(MyFilterRate) - MyRateMin
			End Sub

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilterRateInterpolated.ToList
				Get
					Return MyListOfPriceVol(MySelectedItem)
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal FilterRate As Double) As IList(Of Double) Implements IFilterRateInterpolated.ToList
				Get
					Me.Refresh(FilterRate)
					Return MyListOfPriceVol(MySelectedItem)
				End Get
			End Property

			Public Function CopyFrom() As IFilterRateInterpolated Implements IFilterRateInterpolated.CopyFrom
				Return New FilterRateInterpolated(
					MyListOfPriceVol,
					CInt(Me.Rate),
					CInt(Me.RateMinimum),
					Me.OutputType)
			End Function

			Public ReadOnly Property OutputType As IFilterRateInterpolated.enuOutputType Implements IFilterRateInterpolated.OutputType
				Get
					Return MyOutputType
				End Get
			End Property

			''' <summary>
			''' Transform a value in Percent (between 0 to 1) in the filter rate format
			''' </summary>
			''' <param name="Value">value in percent contain between 0 and 1</param>
			''' <returns>The filter rate</returns>
			''' <remarks></remarks>
			Public Function ToFilterRate(Value As Double) As Double Implements IFilterRateInterpolated.ToFilterRate
				If Value < 0 Then
					Value = 0.0
				ElseIf Value > 1.0 Then
					Value = 1.0
				End If
				Return MyFilterRange * Value + MyRateMin
			End Function

			Public Function ToFilterRateInverse(Value As Double) As Double Implements IFilterRateInterpolated.ToFilterRateInverse
				If Value < MyRateMin Then
					Value = MyRateMin
				ElseIf Value > MyRateMax Then
					Value = MyRateMax
				End If
				Return (Value - MyRateMin) / MyFilterRange
			End Function
		End Class
#End Region
#Region "FilterLowPassExp(Of T As {Structure, IPriceVolLarge})"
		<Serializable()>
		Friend Class FilterLowPassExp(Of T As {New, IPriceVol, PriceVol})
			Private MyRate As Double
			'Private A As Double
			'Private B As Double
			Private FilterValueLast As IPriceVol
			Private FilterValueLastK1 As IPriceVol
			Private ValueLast As IPriceVol
			Private IsValueInitial As Boolean
			Private MyListOfValue As List(Of IPriceVol)
			Private MyFilterLast As IFilterRun
			Private MyFilterOpen As IFilterRun
			Private MyFilterHigh As IFilterRun
			Private MyFilterLow As IFilterRun

			Public Sub New(ByVal FilterRate As Integer)
				Me.New(CDbl(FilterRate))
			End Sub

			Public Sub New(ByVal FilterRate As Double)
				MyListOfValue = New List(Of IPriceVol)
				If FilterRate < 1 Then FilterRate = 1
				MyRate = FilterRate
				MyFilterLast = New FilterExp(FilterRate)
				MyFilterOpen = New FilterExp(FilterRate)
				MyFilterHigh = New FilterExp(FilterRate)
				MyFilterLow = New FilterExp(FilterRate)

				FilterValueLast = Nothing
				ValueLast = FilterValueLast
				'A = CSng(2 / (FilterRate + 1))
				'B = 1 - A
				IsValueInitial = False
			End Sub

			Public Function Filter(ByVal Value As IPriceVol) As IPriceVol
				'this is safe if the ValueLast is not changed
				ValueLast = Value
				MyFilterLast = New FilterExp(Value.Last)
				MyFilterOpen = New FilterExp(Value.Open)
				MyFilterHigh = New FilterExp(Value.High)
				MyFilterLow = New FilterExp(Value.Low)
				'copy from the input value and then update from teh filter result
				If FilterValueLast IsNot Nothing Then
					FilterValueLastK1 = DirectCast(FilterValueLast, PriceVol).CopyFrom
				Else
					FilterValueLastK1 = Nothing
				End If
				FilterValueLast = DirectCast(Value, PriceVol).CopyFrom
				With FilterValueLast
					.Last = CSng(MyFilterLast.FilterLast)
					.Open = CSng(MyFilterOpen.FilterLast)
					.High = CSng(MyFilterHigh.FilterLast)
					.Low = CSng(MyFilterLow.FilterLast)
					.Vol = .Vol
				End With
				FilterValueLast.LastWeighted = RecordPrices.CalculateLastWeighted(FilterValueLast)
				MyListOfValue.Add(FilterValueLast)
				Return FilterValueLast
			End Function

			Public Function Filter(ByRef Value() As IPriceVol) As IPriceVol()
				Dim ThisValue As IPriceVol
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function FilterLast() As IPriceVol
				Return FilterValueLast
			End Function

			Public Function Last() As YahooAccessData.IPriceVol
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return CInt(MyRate)
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of IPriceVol)
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As IPriceVol()
				Return MyListOfValue.ToArray
			End Function

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterLowPassPLL(Of T As {Structure, IPriceVolLarge})"
		<Serializable()>
		Friend Class FilterLowPassPLL(Of T As {Structure, IPriceVolLarge})
			Private MyRate As Integer
			'Private A As Double
			'Private B As Double
			Private FilterValueLast As IPriceVolLarge
			Private FilterValueLastK1 As IPriceVolLarge
			Private ValueLast As IPriceVolLarge
			Private IsValueInitial As Boolean
			Private MyListOfValue As List(Of IPriceVolLarge)
			Private MyFilter As FilterLowPassPLL

			Public Sub New(ByVal FilterRate As Integer)
				MyFilter = New FilterLowPassPLL(FilterRate)
				MyListOfValue = New List(Of IPriceVolLarge)
				MyRate = MyFilter.Rate
				FilterValueLast = Nothing
				ValueLast = FilterValueLast
				'A = CSng(2 / (FilterRate + 1))
				'B = 1 - A
				IsValueInitial = False
			End Sub

			Public Function Filter(ByVal Value As T) As T
				Return DirectCast(Me.Filter(DirectCast(Value, IPriceVolLarge)), T)
			End Function

			Public Function Filter(ByVal Value As IPriceVolLarge) As IPriceVolLarge
#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As IPriceVol = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual.LastWeighted <> ThisResultPrediction.LastWeighted Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
				MyFilter.Filter(Value.Last)
				ValueLast = Value
				FilterValueLast = Value
				With FilterValueLast
					.Last = MyFilter.FilterLast
					.Open = .Last + (Value.Open - Value.Last)
					.High = .Last + (Value.High - Value.Last)
					.Low = .Last + (Value.Low - Value.Last)
					.FilterLast = .Last
					.LastWeighted = RecordPrices.CalculateLastWeighted(FilterValueLast)
					If MyListOfValue.Count > 0 Then
						.LastPrevious = FilterValueLastK1.Last
					Else
						.LastPrevious = .Last
					End If
				End With
				FilterValueLastK1 = FilterValueLast
				MyListOfValue.Add(FilterValueLast)
				Return FilterValueLast
			End Function

			Public Function Filter(ByRef Value() As IPriceVolLarge) As IPriceVolLarge()
				Dim ThisValue As T
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function Filter(ByRef Value() As T) As T()
				Dim ThisValue As T
				Dim ThisValues(0 To Value.Length - 1) As T
				Dim I As Integer

				For Each ThisValue In Value
					ThisValues(I) = Me.Filter(ThisValue)
					I = I + 1
				Next
				Return ThisValues
			End Function

			Public Function Filter(ByRef Value() As IPriceVolLarge, ByVal DelayRemovedToItem As Integer) As IPriceVolLarge()
				Dim ThisValues(0 To Value.Length - 1) As IPriceVolLarge
				Dim I As Integer
				Dim J As Integer

				Dim ThisFilterLeft As New FilterLowPassPLL(Of T)(Me.Rate)
				Dim ThisFilterRight As New FilterLowPassPLL(Of T)(Me.Rate)
				Dim ThisFilterRightList As New List(Of IPriceVolLarge)
				Dim ThisFilterLeftItem As IPriceVolLarge
				Dim ThisFilterRightItem As IPriceVolLarge
				Dim ThisPriceVol As IPriceVolLarge

				'filter from the left
				ThisFilterLeft.Filter(Value)
				'filter from the right the section with the reverse filtering
				For I = DelayRemovedToItem To 0 Step -1
					ThisFilterRightList.Add(ThisFilterRight.Filter(Value(I)))
				Next
				'the data in ThisFilterRightList is reversed
				'need to look at it in reverse order using J
				J = DelayRemovedToItem
				For I = 0 To Value.Length - 1
					ThisPriceVol = New T
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						With ThisPriceVol
							.Last = ThisFilterLeftItem.Last
							.Open = ThisFilterLeftItem.Open
							.High = ThisFilterLeftItem.High
							.Low = ThisFilterLeftItem.Low
							.FilterLast = .Last
						End With
					Else
						ThisFilterRightItem = ThisFilterRightList(J)
						With ThisPriceVol
							.Last = (ThisFilterLeftItem.Last + ThisFilterRightItem.Last) / 2
							.Open = (ThisFilterLeftItem.Open + ThisFilterRightItem.Open) / 2
							.High = (ThisFilterLeftItem.High + ThisFilterRightItem.High) / 2
							.Low = (ThisFilterLeftItem.Low + ThisFilterRightItem.Low) / 2
							.FilterLast = .Last
						End With
					End If
					ThisPriceVol.LastWeighted = RecordPrices.CalculateLastWeighted(ThisPriceVol)
					MyListOfValue.Add(ThisPriceVol)
					ThisValues(I) = ThisPriceVol
					J = J - 1
				Next
				Return ThisValues
			End Function

			Public Function Filter(ByRef Value() As T, ByVal DelayRemovedToItem As Integer) As T()
				Dim ThisValues(0 To Value.Length - 1) As T
				Dim I As Integer
				Dim J As Integer

				Dim ThisFilterLeft As New FilterLowPassPLL(Of T)(Me.Rate)
				Dim ThisFilterRight As New FilterLowPassPLL(Of T)(Me.Rate)
				Dim ThisFilterRightList As New List(Of T)
				Dim ThisFilterLeftItem As IPriceVolLarge
				Dim ThisFilterRightItem As IPriceVolLarge
				Dim ThisPriceVol As IPriceVolLarge

				'filter from the left
				ThisFilterLeft.Filter(Value)
				'filter from the right the section with the reverse filtering
				For I = DelayRemovedToItem To 0 Step -1
					ThisFilterRightList.Add(ThisFilterRight.Filter(Value(I)))
				Next
				'the data in ThisFilterRightList is reversed
				'need to look at it in reverse order using J
				J = DelayRemovedToItem
				For I = 0 To Value.Length - 1
					ThisPriceVol = New T
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						With ThisPriceVol
							.Last = ThisFilterLeftItem.Last
							.Open = ThisFilterLeftItem.Open
							.High = ThisFilterLeftItem.High
							.Low = ThisFilterLeftItem.Low
							.FilterLast = .Last
						End With
					Else
						ThisFilterRightItem = ThisFilterRightList(J)
						With ThisPriceVol
							.Last = (ThisFilterLeftItem.Last + ThisFilterRightItem.Last) / 2
							.Open = (ThisFilterLeftItem.Open + ThisFilterRightItem.Open) / 2
							.High = (ThisFilterLeftItem.High + ThisFilterRightItem.High) / 2
							.Low = (ThisFilterLeftItem.Low + ThisFilterRightItem.Low) / 2
							.FilterLast = .Last
						End With
					End If
					ThisPriceVol.LastWeighted = RecordPrices.CalculateLastWeighted(ThisPriceVol)
					MyListOfValue.Add(ThisPriceVol)
					ThisValues(I) = DirectCast(ThisPriceVol, T)
					J = J - 1
				Next
				Return ThisValues
			End Function

			Public Function FilterBackTo(ByRef Value As IPriceVolLarge) As YahooAccessData.IPriceVolLarge
				Dim ThisValue As IPriceVolLarge = Me.ValueLast
				With ThisValue
					.Last = MyFilter.FilterBackTo(Value.Last)
					.Open = .Last + (Value.Open - Value.Last)
					.High = .Last + (Value.High - Value.Last)
					.Low = .Last + (Value.Low - Value.Last)
					.Range = RecordPrices.CalculateTrueRange(ThisValue)
					.LastWeighted = RecordPrices.CalculateLastWeighted(ThisValue)
				End With
				Return ThisValue
			End Function

			Public Function FilterPredictionNext(ByVal Value As T) As YahooAccessData.IPriceVolLarge
				Dim ThisFilterValueLast As YahooAccessData.IPriceVolLarge = FilterValueLast

				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						ThisFilterValueLast = Value
					Else
						ThisFilterValueLast = FilterValueLast
					End If
				End If
				With ThisFilterValueLast
					.Last = MyFilter.FilterPredictionNext(Value.Last)
					.Open = .Last + (Value.Open - Value.Last)
					.High = .Last + (Value.High - Value.Last)
					.Low = .Last + (Value.Low - Value.Last)
					.Range = RecordPrices.CalculateTrueRange(ThisFilterValueLast)
					.LastWeighted = RecordPrices.CalculateLastWeighted(ThisFilterValueLast)
				End With
				Return ThisFilterValueLast
			End Function

			Public Function FilterLast() As YahooAccessData.IPriceVolLarge
				Return FilterValueLast
			End Function

			Public Function Last() As YahooAccessData.IPriceVolLarge
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of IPriceVolLarge)
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As IPriceVolLarge()
				Return MyListOfValue.ToArray
			End Function

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterLowPassPLLNoDelay"
		<Serializable()>
		Public Class FilterLowPassPLLNoDelay
			Implements IFilter
			Implements IFilterPrediction
			Implements IFilterControl
			Implements IFilterControlRate
			Implements IRegisterKey(Of String)
			Implements IFilterCopy

			Private MyFilterLowPassPLL As FilterLowPassPLL

			Private MyNumberLookAheadPoint As Integer
			Private MyDampingFactor As Double

			Public Sub New(ByVal FilterRate As Double, Optional ByVal DampingFactor As Double = FilterLowPassPLL.DAMPING_FACTOR, Optional IsPredictionEnabled As Boolean = False)
				MyFilterLowPassPLL = New FilterLowPassPLL(FilterRate, DampingFactor, IsPredictionEnabled)
				MyNumberLookAheadPoint = -1
				MyDampingFactor = DampingFactor
			End Sub

			Public Sub New(
				ByVal FilterRate As Double,
				ByVal NumberLookAheadPoint As Integer,
				Optional ByVal DampingFactor As Double = FilterLowPassPLL.DAMPING_FACTOR,
				Optional IsPredictionEnabled As Boolean = False)

				Me.New(FilterRate, DampingFactor, IsPredictionEnabled)
				If NumberLookAheadPoint < 0 Then NumberLookAheadPoint = -1
				MyNumberLookAheadPoint = NumberLookAheadPoint
				MyDampingFactor = DampingFactor
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByRef InputValue() As Double, Optional ByVal DampingFactor As Double = FilterLowPassPLL.DAMPING_FACTOR, Optional IsPredictionEnabled As Boolean = False)
				MyFilterLowPassPLL = New FilterLowPassPLL(FilterRate, InputValue, IsRunFilter:=False, DampingFactor:=DampingFactor, IsPredictionEnabled:=IsPredictionEnabled)
				MyNumberLookAheadPoint = -1
				MyDampingFactor = DampingFactor
				Me.Filter(InputValue)
			End Sub

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyFilterLowPassPLL.Count
				End Get
			End Property

			Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				Return Me.Filter(Value, Value.Length - 1)
			End Function

			Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Dim I As Integer
				Dim K As Integer

				Dim ThisFilterLeft As New FilterLowPassPLL(MyFilterLowPassPLL.ASIFilterControl.FilterRate, MyDampingFactor)
				Dim ThisFilterRight As New FilterLowPassPLL(MyFilterLowPassPLL.ASIFilterControl.FilterRate, MyDampingFactor)
				Dim ThisPriceFiltered As Double

				If MyNumberLookAheadPoint < 0 Then
					'direct call here
					Return MyFilterLowPassPLL.Filter(Value, DelayRemovedToItem)
				Else
					For I = 0 To Value.Length - 1
						'If I = 1000 Then
						'  Debugger.Break()
						'End If
						ThisFilterLeft.Filter(Value(I))
						If I > DelayRemovedToItem Then
							ThisPriceFiltered = ThisFilterLeft.FilterLast
						Else
							'filter backward looking ahead
							ThisFilterRight.ASIFilterControl.Clear()
							K = I + MyNumberLookAheadPoint
							Do
								If K >= Value.Length Then
									K = Value.Length - 1
								End If
								ThisFilterRight.Filter(Value(K))
								K = K - 1
							Loop Until K < I
							ThisPriceFiltered = (ThisFilterLeft.FilterLast + ThisFilterRight.FilterLast) / 2
						End If
						MyFilterLowPassPLL.Filter(Value(I), ThisPriceFiltered)
					Next
					Return MyFilterLowPassPLL.ToArray
				End If
			End Function

			Private Function Filter(Value As Double) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function Filter(Value As Single) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				Throw New NotSupportedException
			End Function

			Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Throw New NotSupportedException
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return MyFilterLowPassPLL.FilterLast
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Return MyFilterLowPassPLL.FilterLastToPriceVol
			End Function

			Public Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
				Return MyFilterLowPassPLL.FilterPredictionNext(Value)
			End Function

			Public Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
				Return MyFilterLowPassPLL.FilterPredictionNext(Value)
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return MyFilterLowPassPLL.Last
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Return MyFilterLowPassPLL.LastToPriceVol
			End Function

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyFilterLowPassPLL.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyFilterLowPassPLL.Min
				End Get
			End Property

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyFilterLowPassPLL.Rate
				End Get
			End Property

			Public Property Tag As String Implements IFilter.Tag

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyFilterLowPassPLL.ToArray
			End Function

			Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyFilterLowPassPLL.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyFilterLowPassPLL.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public ReadOnly Property ToList As System.Collections.Generic.IList(Of Double) Implements IFilter.ToList
				Get
					Return MyFilterLowPassPLL.ToList
				End Get
			End Property

			Public ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of Double) Implements IFilter.ToListOfError
				Get
					Return MyFilterLowPassPLL.ToListOfError
				End Get
			End Property

			Public ReadOnly Property ToListScaled As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyFilterLowPassPLL.ToListScaled
				End Get
			End Property

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return MyFilterLowPassPLL.ToString
			End Function
#Region "IFilterPrediction"
			Public Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
				Return Me
			End Function

			Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassPLL.AsIFilterPrediction.FilterPrediction(NumberOfPrediction)
			End Function

			Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassPLL.AsIFilterPrediction.FilterPrediction(NumberOfPrediction, GainPerYear)
			End Function

			Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassPLL.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction)
			End Function

			Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassPLL.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction, GainPerYear)
			End Function

			Private ReadOnly Property IFilterPrediction_IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
				Get
					Return MyFilterLowPassPLL.AsIFilterPrediction.IsEnabled
				End Get
			End Property

			Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
				Get
					Return MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYear
				End Get
			End Property

			Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
				Get
					Return MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYearDerivative
				End Get
			End Property
#End Region
#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function

			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID

			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
#Region "IFilterControl"
			Public Function ASIFilterControl() As IFilterControl Implements IFilterControl.AsIFilterControl
				Return Me
			End Function

			Private Sub IFilterControl_Clear() Implements IFilterControl.Clear
				MyFilterLowPassPLL.ASIFilterControl.Clear()
			End Sub

			Private Sub IFilterControl_Refresh(FilterRate As Double) Implements IFilterControl.Refresh
				MyFilterLowPassPLL.ASIFilterControl.Refresh(FilterRate)
			End Sub

			Private Sub IFilterControl_Refresh(Rate As Integer) Implements IFilterControl.Refresh
				MyFilterLowPassPLL.ASIFilterControl.Refresh(Rate)
			End Sub

			Private ReadOnly Property IFilterControl_FilterRate As Double Implements IFilterControl.FilterRate
				Get
					Return MyFilterLowPassPLL.ASIFilterControl.FilterRate
				End Get
			End Property

			Private Function IFilterControl_InputValue() As Double() Implements IFilterControl.InputValue
				Return MyFilterLowPassPLL.ASIFilterControl.InputValue
			End Function

			Private ReadOnly Property IFilterControl_IsInputEnabled As Boolean Implements IFilterControl.IsInputEnabled
				Get
					Return MyFilterLowPassPLL.ASIFilterControl.IsInputEnabled
				End Get
			End Property
#End Region
#Region "IFilterControlRate"
			Public Function AsIFilterControlRate() As IFilterControlRate Implements IFilterControlRate.AsIFilterControlRate
				Return Me
			End Function

			Private Sub IFilterControlRate_UpdateRate(Rate As Double) Implements IFilterControlRate.UpdateRate
				MyFilterLowPassPLL.AsIFilterControlRate.UpdateRate(Rate)
			End Sub

			Private Sub IFilterControlRate_UpdateRate(Rate As Integer) Implements IFilterControlRate.UpdateRate
				MyFilterLowPassPLL.AsIFilterControlRate.UpdateRate(Rate)
			End Sub

			Private Property IFilterControlRate_Enabled As Boolean Implements IFilterControlRate.Enabled
				Get
					Return MyFilterLowPassPLL.AsIFilterControlRate.Enabled
				End Get
				Set(value As Boolean)
					MyFilterLowPassPLL.AsIFilterControlRate.Enabled = value
				End Set
			End Property
#End Region
#Region "IFilterCopy"
			Public Function AsIFilterCopy() As IFilterCopy Implements IFilterCopy.AsIFilterCopy
				Return Me
			End Function

			Private Function IFilterCopy_CopyFrom() As IFilter Implements IFilterCopy.CopyFrom

				Throw New NotSupportedException
				Dim ThisFilter As FilterLowPassExpNoDelay
				Dim ThisFilterControl As IFilterControl = Me
				If ThisFilterControl.IsInputEnabled Then
					ThisFilter = New FilterLowPassExpNoDelay(ThisFilterControl.FilterRate, ThisFilterControl.InputValue, IsPredictionEnabled:=MyFilterLowPassPLL.AsIFilterPrediction.IsEnabled)
				Else
					ThisFilter = New FilterLowPassExpNoDelay(ThisFilterControl.FilterRate, IsPredictionEnabled:=MyFilterLowPassPLL.AsIFilterPrediction.IsEnabled)
				End If
				Return ThisFilter
			End Function
#End Region
		End Class
#End Region
#Region "FilterLowPassExpNoDelay"
		<Serializable()>
		Public Class FilterLowPassExpNoDelay
			Implements IFilter
			Implements IFilterPrediction
			Implements IFilterControl
			Implements IFilterControlRate
			Implements IRegisterKey(Of String)
			Implements IFilterCopy

			'Inherits FilterLowPassExp

			Private MyFilterLowPassExp As FilterLowPassExp
			Private MyNumberLookAheadPoint As Integer

			Public Sub New(ByVal FilterRate As Double, Optional IsPredictionEnabled As Boolean = False)
				MyFilterLowPassExp = New FilterLowPassExp(FilterRate, IsPredictionEnabled)
				MyNumberLookAheadPoint = -1
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByVal NumberLookAheadPoint As Integer, Optional IsPredictionEnabled As Boolean = False)
				Me.New(FilterRate, IsPredictionEnabled)
				If NumberLookAheadPoint < 0 Then NumberLookAheadPoint = -1
				MyNumberLookAheadPoint = NumberLookAheadPoint
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByRef InputValue() As Double, Optional IsPredictionEnabled As Boolean = False)
				MyFilterLowPassExp = New FilterLowPassExp(FilterRate, InputValue, IsRunFilter:=False, IsPredictionEnabled:=IsPredictionEnabled)
				MyNumberLookAheadPoint = -1
				'ready to run the filter
				Me.Filter(InputValue)
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByRef InputValue() As Double, ByVal NumberLookAheadPoint As Integer, Optional IsPredictionEnabled As Boolean = False)
				MyFilterLowPassExp = New FilterLowPassExp(FilterRate, InputValue, IsRunFilter:=False, IsPredictionEnabled:=IsPredictionEnabled)
				If NumberLookAheadPoint < 0 Then NumberLookAheadPoint = -1
				MyNumberLookAheadPoint = NumberLookAheadPoint
				'ready to run the filter
				Me.Filter(InputValue)
			End Sub

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyFilterLowPassExp.Count
				End Get
			End Property

			Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				Return Me.Filter(Value, Value.Length - 1)
			End Function

			Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				If MyNumberLookAheadPoint < 0 Then
					'direct call here
					Return MyFilterLowPassExp.Filter(Value, DelayRemovedToItem)
				Else
					Dim I As Integer
					Dim K As Integer

					Dim ThisFilterLeft As New FilterLowPassExp(MyFilterLowPassExp.ASIFilterControl.FilterRate)
					Dim ThisFilterRight As New FilterLowPassExp(MyFilterLowPassExp.ASIFilterControl.FilterRate)
					Dim ThisPriceFiltered As Double
					For I = 0 To Value.Length - 1
						ThisFilterLeft.Filter(Value(I))
						If I > DelayRemovedToItem Then
							ThisPriceFiltered = ThisFilterLeft.FilterLast
						Else
							'filter backward looking ahead
							ThisFilterRight.ASIFilterControl.Clear()
							K = I + MyNumberLookAheadPoint
							Do
								If K >= Value.Length Then
									K = Value.Length - 1
								End If
								ThisFilterRight.Filter(Value(K))
								K = K - 1
							Loop Until K < I
							ThisPriceFiltered = (ThisFilterLeft.FilterLast + ThisFilterRight.FilterLast) / 2
						End If
						MyFilterLowPassExp.Filter(Value(I), ThisPriceFiltered)
					Next
					Return MyFilterLowPassExp.ToArray
				End If
			End Function

			Private Function Filter(Value As Double) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function Filter(Value As Single) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Throw New NotSupportedException
			End Function

			Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				Throw New NotSupportedException
			End Function

			Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Throw New NotSupportedException
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return MyFilterLowPassExp.FilterLast
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Return MyFilterLowPassExp.FilterLastToPriceVol
			End Function

			Public Function FilterPredictionNext(Value As Double) As Double Implements IFilter.FilterPredictionNext
				Return MyFilterLowPassExp.FilterPredictionNext(Value)
			End Function

			Public Function FilterPredictionNext(Value As Single) As Double Implements IFilter.FilterPredictionNext
				Return MyFilterLowPassExp.FilterPredictionNext(Value)
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return MyFilterLowPassExp.Last
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Return MyFilterLowPassExp.LastToPriceVol
			End Function

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyFilterLowPassExp.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyFilterLowPassExp.Min
				End Get
			End Property

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyFilterLowPassExp.Rate
				End Get
			End Property

			Public Property Tag As String Implements IFilter.Tag

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyFilterLowPassExp.ToArray
			End Function

			Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyFilterLowPassExp.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyFilterLowPassExp.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public ReadOnly Property ToList As System.Collections.Generic.IList(Of Double) Implements IFilter.ToList
				Get
					Return MyFilterLowPassExp.ToList
				End Get
			End Property

			Public ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of Double) Implements IFilter.ToListOfError
				Get
					Return MyFilterLowPassExp.ToListOfError
				End Get
			End Property

			Public ReadOnly Property ToListScaled As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyFilterLowPassExp.ToListScaled
				End Get
			End Property

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return MyFilterLowPassExp.ToString
			End Function
#Region "IFilterPrediction"
			Public Function AsIFilterPrediction() As IFilterPrediction Implements IFilterPrediction.AsIFilterPrediction
				Return Me
			End Function

			Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassExp.AsIFilterPrediction.FilterPrediction(NumberOfPrediction)
			End Function

			Private Function IFilterPrediction_FilterPrediction(NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
				Throw New NotSupportedException
				Return MyFilterLowPassExp.AsIFilterPrediction.FilterPrediction(NumberOfPrediction, GainPerYear)
			End Function

			Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassExp.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction)
			End Function

			Private Function IFilterPrediction_FilterPrediction(Index As Integer, NumberOfPrediction As Integer, GainPerYear As Double) As Double Implements IFilterPrediction.FilterPrediction
				Return MyFilterLowPassExp.AsIFilterPrediction.FilterPrediction(Index, NumberOfPrediction, GainPerYear)
			End Function

			Private ReadOnly Property IFilterPrediction_IsEnabled As Boolean Implements IFilterPrediction.IsEnabled
				Get
					Return MyFilterLowPassExp.AsIFilterPrediction.IsEnabled
				End Get
			End Property

			Private ReadOnly Property IFilterPrediction_ToListOfGainPerYear As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYear
				Get
					Return MyFilterLowPassExp.AsIFilterPrediction.ToListOfGainPerYear
				End Get
			End Property

			Private ReadOnly Property IFilterPrediction_ToListOfGainPerYearDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPrediction.ToListOfGainPerYearDerivative
				Get
					Return MyFilterLowPassExp.AsIFilterPrediction.ToListOfGainPerYearDerivative
				End Get
			End Property
#End Region
#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function

			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID

			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
#Region "IFilterControl"
			Public Function ASIFilterControl() As IFilterControl Implements IFilterControl.AsIFilterControl
				Return Me
			End Function

			Private Sub IFilterControl_Clear() Implements IFilterControl.Clear
				MyFilterLowPassExp.ASIFilterControl.Clear()
			End Sub

			Private Sub IFilterControl_Refresh(FilterRate As Double) Implements IFilterControl.Refresh
				MyFilterLowPassExp.ASIFilterControl.Refresh(FilterRate)
			End Sub

			Private Sub IFilterControl_Refresh(Rate As Integer) Implements IFilterControl.Refresh
				MyFilterLowPassExp.ASIFilterControl.Refresh(Rate)
			End Sub

			Private ReadOnly Property IFilterControl_FilterRate As Double Implements IFilterControl.FilterRate
				Get
					Return MyFilterLowPassExp.ASIFilterControl.FilterRate
				End Get
			End Property

			Private Function IFilterControl_InputValue() As Double() Implements IFilterControl.InputValue
				Return MyFilterLowPassExp.ASIFilterControl.InputValue
			End Function

			Private ReadOnly Property IFilterControl_IsInputEnabled As Boolean Implements IFilterControl.IsInputEnabled
				Get
					Return MyFilterLowPassExp.ASIFilterControl.IsInputEnabled
				End Get
			End Property
#End Region
#Region "IFilterControlRate"
			Public Function AsIFilterControlRate() As IFilterControlRate Implements IFilterControlRate.AsIFilterControlRate
				Return Me
			End Function

			Private Sub IFilterControlRate_UpdateRate(Rate As Double) Implements IFilterControlRate.UpdateRate
				MyFilterLowPassExp.AsIFilterControlRate.UpdateRate(Rate)
			End Sub

			Private Sub IFilterControlRate_UpdateRate(Rate As Integer) Implements IFilterControlRate.UpdateRate
				MyFilterLowPassExp.AsIFilterControlRate.UpdateRate(Rate)
			End Sub

			Private Property IFilterControlRate_Enabled As Boolean Implements IFilterControlRate.Enabled
				Get
					Return MyFilterLowPassExp.AsIFilterControlRate.Enabled
				End Get
				Set(value As Boolean)
					MyFilterLowPassExp.AsIFilterControlRate.Enabled = value
				End Set
			End Property
#End Region
#Region "IFilterCopy"
			Public Function AsIFilterCopy() As IFilterCopy Implements IFilterCopy.AsIFilterCopy
				Return Me
			End Function

			Private Function IFilterCopy_CopyFrom() As IFilter Implements IFilterCopy.CopyFrom
				Dim ThisFilter As FilterLowPassExpNoDelay
				Dim ThisFilterControl As IFilterControl = Me
				If ThisFilterControl.IsInputEnabled Then
					ThisFilter = New FilterLowPassExpNoDelay(ThisFilterControl.FilterRate, ThisFilterControl.InputValue, IsPredictionEnabled:=MyFilterLowPassExp.AsIFilterPrediction.IsEnabled)
				Else
					ThisFilter = New FilterLowPassExpNoDelay(ThisFilterControl.FilterRate, IsPredictionEnabled:=MyFilterLowPassExp.AsIFilterPrediction.IsEnabled)
				End If
				Return ThisFilter
			End Function
#End Region
		End Class
#End Region
#Region "FilterLowPassRMS"
		''' <summary>
		''' Not yet completed
		''' </summary>
		''' <remarks></remarks>
		<Serializable()>
		Public Class FilterLowPassRMS
			Implements IFilter

			Private MyRate As Integer
			Private A As Double
			Private B As Double
			Private FilterValueLastK1 As Double
			Private FilterValueLast As Double
			Private ValueLast As Double
			Private ValueLastK1 As Double
			'Private MyValueSumForInit As Double
			Private IsValueInitial As Boolean
			Private MyListOfValue As ListScaled
			Private MyFilterValueSum As Double
			Private MyListWindows As ListWindowFrame

			Public Sub New(ByVal FilterRate As Double)
				Throw New NotImplementedException
				MyListOfValue = New ListScaled
				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				FilterValueLast = 0
				FilterValueLastK1 = 0
				ValueLast = 0
				ValueLastK1 = 0
				'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
				'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
				'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
				A = CDbl((2 / (FilterRate + 1)))

				'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
				B = 1 - A
				IsValueInitial = False
				'MyValueSumForInit = 0
			End Sub

			Public Sub New(ByVal FilterRate As Integer)
				Me.New(CDbl(FilterRate))
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Double)
				Me.New(FilterRate)
				FilterValueLast = ValueInitial
				FilterValueLastK1 = FilterValueLast
				ValueLast = ValueInitial
				ValueLastK1 = ValueLast
				IsValueInitial = True
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal ValueInitial As Single)
				Me.New(FilterRate, CDbl(ValueInitial))
			End Sub

			Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						FilterValueLast = Value
					End If
				End If
				FilterValueLastK1 = FilterValueLast

				MyFilterValueSum = MyFilterValueSum + Value
				If MyListOfValue.Count = Rate Then


				Else

				End If

				FilterValueLast = A * Value + B * FilterValueLast
				MyListOfValue.Add(FilterValueLast)
				ValueLastK1 = ValueLast
				ValueLast = Value
				Return FilterValueLast
			End Function

			Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Return Me.Filter(CDbl(Value.Last))
			End Function

			Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			''' <summary>
			''' Special filtering that can be used to remove the delay starting at a specific point
			''' </summary>
			''' <param name="Value">The value to be filtered</param>
			''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
			''' <returns>The result</returns>
			''' <remarks></remarks>
			Public Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Dim ThisValues(0 To Value.Length - 1) As Double
				Dim I As Integer
				Dim J As Integer

				Dim ThisFilterLeft As New FilterLowPassExp(Me.Rate)
				Dim ThisFilterRight As New FilterLowPassExp(Me.Rate)
				Dim ThisFilterLeftItem As Double
				Dim ThisFilterRightItem As Double
				Dim ThisPriceVol As Double

				'filter from the left
				ThisFilterLeft.Filter(Value)
				'filter from the right the section with the reverse filtering
				For I = DelayRemovedToItem To 0 Step -1
					ThisFilterRight.Filter(Value(I))
				Next
				'the data in ThisFilterRightList is reversed
				'need to look at it in reverse order using J
				J = DelayRemovedToItem
				For I = 0 To Value.Length - 1
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						ThisPriceVol = ThisFilterLeftItem
					Else
						ThisFilterRightItem = ThisFilterRight.ToList(J)
						ThisPriceVol = (ThisFilterLeftItem + ThisFilterRightItem) / 2
					End If
					MyListOfValue.Add(ThisPriceVol)
					ThisValues(I) = ThisPriceVol
					J = J - 1
				Next
				Return ThisValues
			End Function

			Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Return 0.0
			End Function

			Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				If A > 0 Then
					Return (Value - B * FilterValueLastK1) / A
				Else
					Return Value
				End If
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
				With ThisPriceVol
					.LastPrevious = CSng(FilterValueLastK1)
					If Me.Last > .Last Then
						.High = CSng(Me.Last)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf Me.Last < .Last Then
						.Low = CSng(Me.Last)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
				With ThisPriceVol
					.LastPrevious = CSng(ValueLastK1)
					If Me.FilterLast > .Last Then
						.High = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf Me.FilterLast < .Last Then
						.Low = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
				Return CSng(Me.Filter(CDbl(Value)))
			End Function

			Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
				Dim ThisFilterValueLast As Double = FilterValueLast
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						ThisFilterValueLast = Value
					Else
						ThisFilterValueLast = FilterValueLast
					End If
				End If
				ThisFilterValueLast = A * Value + B * ThisFilterValueLast
				Return ThisFilterValueLast
			End Function

			Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
				Return Me.FilterPredictionNext(CDbl(Value))
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return FilterValueLast
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String Implements IFilter.Tag

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterPriceDecimation"
		<Serializable()>
		Public Class FilterPriceDecimation
			Implements IFilter(Of IPriceVol)
			Implements IFilterPivotPointList
			Implements IRegisterKey(Of String)


			Private MyRate As Integer
			Private MyRatePivotPoint As Integer
			Private FilterValueLast As IPriceVol
			Private ValueLast As IPriceVol
			Private MyListOfPriceVol As List(Of IPriceVol)
			Private MyListOfPivotPointHigh As List(Of Integer)
			Private MyListOfPivotPointLow As List(Of Integer)
			Private MyListOfPivotPointHighDistance As List(Of Double)
			Private MyListOfPivotPointLowDistance As List(Of Double)
			Private MyListOfPivotPointCompositeDistance As List(Of Double)
			Private MyListOfPivotPointBuySell As List(Of Double)
			Private MyListForPivotPointCompositeDistanceWithGain As List(Of Double)
			Private MyFilterLowPassPLL As Filter.FilterLowPassPLL
			Private MyFilterLowPassExp As Filter.FilterLowPassExp
			Private MyListOfPriceGainDerivativeDifference As List(Of Double)
			Private MyListOfPriceGainDifference As List(Of Double)

			Private MyListOfWindowFramePivotPoint As ListWindowFrameAsClass(Of PriceVol)
			Private MyHighIndexForeCast As Nullable(Of Integer)
			Private MyLowIndexForeCast As Nullable(Of Integer)
			Private MyHighIndexForeCastDistanceToPivot As Integer
			Private MyLowIndexForeCastDistanceToPivot As Integer
			Private MyGainVariationWeight As Double


			Public Sub New(ByVal FilterRate As Double, ByVal GainVariationWeight As Double)
				Me.New(FilterRate)
				MyGainVariationWeight = GainVariationWeight
			End Sub

			Public Sub New(ByVal FilterRate As Double)

				MyListOfPriceVol = New List(Of IPriceVol)
				MyListOfPivotPointHigh = New List(Of Integer)
				MyListOfPivotPointLow = New List(Of Integer)
				MyListOfPivotPointHighDistance = New List(Of Double)
				MyListOfPivotPointLowDistance = New List(Of Double)
				MyListOfPivotPointCompositeDistance = New List(Of Double)
				MyListForPivotPointCompositeDistanceWithGain = New List(Of Double)
				MyListOfPriceGainDerivativeDifference = New List(Of Double)
				MyListOfPriceGainDifference = New List(Of Double)

				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				MyRatePivotPoint = 2 * MyRate + 1
				MyListOfWindowFramePivotPoint = New ListWindowFrameAsClass(Of PriceVol)(MyRatePivotPoint)
				MyFilterLowPassPLL = New Filter.FilterLowPassPLL(MyRatePivotPoint, IsPredictionEnabled:=True)
				MyFilterLowPassExp = New Filter.FilterLowPassExp(MyRatePivotPoint, IsPredictionEnabled:=True)

				'MyListOfWindowFramePivotPointHalf = New ListWindowFrame(Of PriceVol)(MyRatePivotPointHalf)
				FilterValueLast = Nothing
				ValueLast = Nothing
				MyHighIndexForeCast = Nothing
				MyLowIndexForeCast = Nothing
				MyHighIndexForeCastDistanceToPivot = 0
				MyLowIndexForeCastDistanceToPivot = 0
				MyGainVariationWeight = 1.0
			End Sub

			Public Sub New(ByVal FilterRate As Integer)
				Me.New(CDbl(FilterRate))
			End Sub

			Public ReadOnly Property Count As Integer Implements IFilter(Of IPriceVol).Count
				Get
					Return MyListOfPriceVol.Count
				End Get
			End Property

			Public Function Filter(ByRef Value() As Double) As IPriceVol() Implements IFilter(Of IPriceVol).Filter
				For Each ThisValue In Value
					Me.Filter(New PriceVol(CSng(ThisValue)))
				Next
				Return Me.ToArray
			End Function

			Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As IPriceVol() Implements IFilter(Of IPriceVol).Filter
				Throw New NotImplementedException
			End Function

			Public Function Filter(Value As Double) As IPriceVol Implements IFilter(Of IPriceVol).Filter
				Return Me.Filter(New PriceVol(CSng(Value)))
			End Function

			Public Function Filter(Value As Single) As IPriceVol Implements IFilter(Of IPriceVol).Filter
				Return Me.Filter(New PriceVol(Value))
			End Function

			Public Function Filter(Value As IPriceVol) As IPriceVol Implements IFilter(Of IPriceVol).Filter
				Dim ThisDistanceComposite As Double
				Dim ThisDistanceCompositeFiltered As Double
				Dim ThisItemLevel As Double

				If MyListOfPriceVol.Count = 0 Then
					FilterValueLast = Value
				End If
				MyListOfWindowFramePivotPoint.Add(DirectCast(Value, PriceVol))
				'remove the error handle eventually
				Try
					FilterValueLast = MyListOfWindowFramePivotPoint.ItemDecimate
				Catch ex As Exception
					Debugger.Break()
				End Try
				If FilterValueLast.Range > 0 Then
					ThisItemLevel = (FilterValueLast.Last - FilterValueLast.Low) / FilterValueLast.Range
				Else
					ThisItemLevel = 1.0
				End If

				MyListOfPriceVol.Add(FilterValueLast)
				MyFilterLowPassExp.Filter(Value.Last)
				MyFilterLowPassPLL.Filter(Value.Last)

				'MyListOfWindowFramePivotPointHalf.Add(DirectCast(Value, PriceVol))
				'If MyListOfPriceVol.Count = 1312 Then
				'  Debugger.Break()
				'End If
				'note that MyRatePivotPoint = 2 * MyRate + 1
				'we start the process only when teh windows is full
				If MyListOfWindowFramePivotPoint.Count = MyRatePivotPoint Then
					If MyListOfWindowFramePivotPoint.ItemHighIndex = MyRate Then
						'got a confirmed high pivot point
						MyListOfPivotPointHigh.Add(MyListOfPriceVol.Count - MyRate - 1)
						MyHighIndexForeCastDistanceToPivot = 0
						MyListOfPivotPointHighDistance.Add(0.0)
					Else
						'prediction on the next coming point
						If MyListOfWindowFramePivotPoint.ItemHighIndex > MyRate Then
							'a new high is confirmed that may become a new pivot point
							'calculate the distance to a pivot point before it become confirmed 
							MyHighIndexForeCastDistanceToPivot = MyListOfWindowFramePivotPoint.ItemHighIndex - MyRate
							MyHighIndexForeCast = MyListOfPriceVol.Count - ((MyListOfWindowFramePivotPoint.Count - MyListOfWindowFramePivotPoint.ItemHighIndex) - 1) - 1
							MyListOfPivotPointHighDistance.Add(CalculatePivotPointDistanceRatioWeight(MyHighIndexForeCastDistanceToPivot / MyRate))
						Else
							MyHighIndexForeCastDistanceToPivot = 0
							MyHighIndexForeCast = Nothing
							MyListOfPivotPointHighDistance.Add(0.0)
						End If
					End If

					If MyListOfWindowFramePivotPoint.ItemLowIndex = MyRate Then
						'got a confirmed low pivot point
						MyListOfPivotPointLow.Add(MyListOfPriceVol.Count - MyRate - 1)
						MyLowIndexForeCastDistanceToPivot = 0
						MyListOfPivotPointLowDistance.Add(0.0)
					Else
						'prediction on the next coming point
						If MyListOfWindowFramePivotPoint.ItemLowIndex > MyRate Then
							MyLowIndexForeCastDistanceToPivot = MyListOfWindowFramePivotPoint.ItemLowIndex - MyRate
							MyLowIndexForeCast = MyListOfPriceVol.Count - ((MyListOfWindowFramePivotPoint.Count - MyListOfWindowFramePivotPoint.ItemLowIndex) - 1) - 1
							MyListOfPivotPointLowDistance.Add(CalculatePivotPointDistanceRatioWeight(MyLowIndexForeCastDistanceToPivot / MyRate))
						Else
							MyLowIndexForeCastDistanceToPivot = 0
							MyLowIndexForeCast = Nothing
							MyListOfPivotPointLowDistance.Add(0.0)
						End If
					End If
				Else
					MyListOfPivotPointHighDistance.Add(0.0)
					MyListOfPivotPointLowDistance.Add(0.0)
				End If
				'ThisDistanceComposite = (
				'  MyListOfPivotPointHighDistance(MyListOfPivotPointHighDistance.Count - 1) -
				'  MyListOfPivotPointLowDistance(MyListOfPivotPointLowDistance.Count - 1)) / 2
				ThisDistanceComposite = (MyListOfPivotPointHighDistance.Last - MyListOfPivotPointLowDistance.Last) / 2
				If ThisDistanceComposite > 0 Then
					ThisDistanceComposite = ThisItemLevel * ThisDistanceComposite
				ElseIf ThisDistanceComposite < 0 Then
					ThisDistanceComposite = (1 - ThisItemLevel) * ThisDistanceComposite
				End If

				MyListOfPivotPointCompositeDistance.Add(ThisDistanceComposite)
				MyListOfPriceGainDifference.Add(MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYear.Last - MyFilterLowPassExp.AsIFilterPrediction.ToListOfGainPerYear.Last)
				MyListOfPriceGainDerivativeDifference.Add(MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last - MyFilterLowPassExp.AsIFilterPrediction.ToListOfGainPerYearDerivative.Last)

				ThisDistanceCompositeFiltered = ThisDistanceComposite + MyGainVariationWeight * (MyListOfPriceGainDifference.Last + MyListOfPriceGainDerivativeDifference.Last)
				MyListForPivotPointCompositeDistanceWithGain.Add(WaveForm.SignalLimit(ThisDistanceCompositeFiltered, 1))
				ValueLast = Value
				Return FilterValueLast
			End Function



			Public Function FilterBackTo(ByRef Value As IPriceVol) As Double Implements IFilter(Of IPriceVol).FilterBackTo
				Throw New NotImplementedException
			End Function

			Public Function FilterErrorLast() As IPriceVol Implements IFilter(Of IPriceVol).FilterErrorLast
				Throw New NotImplementedException
			End Function

			Public Function FilterLast() As IPriceVol Implements IFilter(Of IPriceVol).FilterLast
				Return FilterValueLast
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter(Of IPriceVol).FilterLastToPriceVol
				Return FilterValueLast
			End Function

			Public Function FilterPredictionNext(Value As Double) As IPriceVol Implements IFilter(Of IPriceVol).FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Public Function FilterPredictionNext(Value As Single) As IPriceVol Implements IFilter(Of IPriceVol).FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Public Function Last() As Double Implements IFilter(Of IPriceVol).Last
				Return ValueLast.Last
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter(Of IPriceVol).LastToPriceVol
				Return ValueLast
			End Function

			Public ReadOnly Property Max As Double Implements IFilter(Of IPriceVol).Max
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter(Of IPriceVol).Min
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property Rate As Integer Implements IFilter(Of IPriceVol).Rate
				Get
					Return MyRate
				End Get
			End Property

			Public Property Tag As String Implements IFilter(Of IPriceVol).Tag

			Public Function ToArray() As IPriceVol() Implements IFilter(Of IPriceVol).ToArray
				Return MyListOfPriceVol.ToArray
			End Function

			Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As IPriceVol() Implements IFilter(Of IPriceVol).ToArray
				Return MyListOfPriceVol.ToArray
			End Function

			Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As IPriceVol() Implements IFilter(Of IPriceVol).ToArray
				Return MyListOfPriceVol.ToArray
			End Function

			Public ReadOnly Property ToList As System.Collections.Generic.IList(Of IPriceVol) Implements IFilter(Of IPriceVol).ToList
				Get
					Return MyListOfPriceVol
				End Get
			End Property

			Public ReadOnly Property ToListOfError As System.Collections.Generic.IList(Of IPriceVol) Implements IFilter(Of IPriceVol).ToListOfError
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled As ListScaled Implements IFilter(Of IPriceVol).ToListScaled
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public Overrides Function ToString() As String Implements IFilter(Of IPriceVol).ToString
				Return Me.FilterLast.ToString
			End Function

			Public ReadOnly Property AsIFilterPivotPointList As IFilterPivotPointList Implements IFilterPivotPointList.AsIFilterPivotPointList
				Get
					Return Me
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfHighIndex As System.Collections.Generic.IList(Of Integer) Implements IFilterPivotPointList.ToListOfHighIndex
				Get
					Return MyListOfPivotPointHigh
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfLowIndex As System.Collections.Generic.IList(Of Integer) Implements IFilterPivotPointList.ToListOfLowIndex
				Get
					Return MyListOfPivotPointLow
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfCompositeDistance As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfCompositeDistance
				Get
					Return MyListOfPivotPointCompositeDistance
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfHighDistance As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfHighDistance
				Get
					Return MyListOfPivotPointHighDistance
				End Get
			End Property


			Private ReadOnly Property IFilterPivotPointList_ToListOfLowDistance As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfLowDistance
				Get
					Return MyListOfPivotPointLowDistance
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_HighIndexForeCast As Integer? Implements IFilterPivotPointList.HighIndexForeCast
				Get
					Return MyHighIndexForeCast
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_LowIndexForeCast As Integer? Implements IFilterPivotPointList.LowIndexForeCast
				Get
					Return MyLowIndexForeCast
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_HighIndexForeCastDistanceToPivot As Integer Implements IFilterPivotPointList.HighIndexForeCastDistanceToPivot
				Get
					Return MyHighIndexForeCastDistanceToPivot
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_LowIndexForeCastDistanceToPivot As Integer Implements IFilterPivotPointList.LowIndexForeCastDistanceToPivot
				Get
					Return MyLowIndexForeCastDistanceToPivot
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfFilteredCompositeDistance As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfFilteredCompositeDistance
				Get
					Return MyListForPivotPointCompositeDistanceWithGain
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfPriceGain As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfPriceGain
				Get
					Return MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYear
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfPriceGainDerivative As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfPriceGainDerivative
				Get
					Return MyFilterLowPassPLL.AsIFilterPrediction.ToListOfGainPerYearDerivative
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfPriceGainDerivativeDifference As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfPriceGainDerivativeDifference
				Get
					Return MyListOfPriceGainDerivativeDifference
				End Get
			End Property

			Private ReadOnly Property IFilterPivotPointList_ToListOfPriceGainDifference As System.Collections.Generic.IList(Of Double) Implements IFilterPivotPointList.ToListOfPriceGainDifference
				Get
					Return MyListOfPriceGainDifference
				End Get
			End Property

			Private Function CalculatePivotPointDistanceRatioWeight(ByVal DistanceRatio As Double) As Double
				Return 1 - Math.Exp(-3 * (DistanceRatio ^ 2))
			End Function
#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function
			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
		End Class
#End Region
#Region "FilterBrownianStatistic"
		<Serializable()>
		Public Class FilterBrownianStatistic
			Implements IFilter
			Implements IRegisterKey(Of String)


			Private Const GAIN_RATE As Integer = 25
			Private MyRate As Integer
			Private A As Double
			Private B As Double
			Private FilterValueLastK1 As Double
			Private FilterValueLast As Double
			Private ValueLast As Double
			Private ValueLastK1 As Double
			'Private MyValueSumForInit As Double
			Private IsValueInitial As Boolean
			Private MyListOfValue As ListScaled
			Private ThisFilterVolatilityForBrownianStatistic As Filter.FilterVolatilityYangZhang
			Private ThisFilterLPForPricePLLBrownianProbability As FilterLowPassPLL
			Private ThisListOfTargetPricePredictionGainYieldPerYear As IList(Of Double)
			Private ThisListOfTargetPricePredictionGainPerYear As IList(Of Double)
			Private ThisFilterLPForTargetPricePredictionForK0 As FilterLowPassExpPredict
			Dim ThisProbability As Double
			Private ThisProbabilityInverseInSigmaOfThirdRatio As Double
			Private ThisFilterPLLProbabilitySlow As FilterLowPassPLL
			Private ThisFilterPLLProbabilityFast As FilterLowPassPLL
			Private ThisPositionForPrediction As Integer

			Public Sub New(ByVal FilterRate As Integer)
				Me.New(CDbl(FilterRate))
			End Sub

			Public Sub New(
				ByVal FilterRate As Double)

				ThisFilterVolatilityForBrownianStatistic = New FilterVolatilityYangZhang(
					StockOption.VolatilityFilterRate(IStockOption.enuVolatilityStandardYearlyType.Monthly),
					 FilterVolatility.enuVolatilityStatisticType.Exponential)

				ThisFilterLPForPricePLLBrownianProbability = New FilterLowPassPLL(5)
				ThisListOfTargetPricePredictionGainYieldPerYear = New List(Of Double)
				ThisListOfTargetPricePredictionGainPerYear = New YahooAccessData.ListScaled
				ThisFilterLPForTargetPricePredictionForK0 = New FilterLowPassExpPredict(GAIN_RATE, 0)

				MyListOfValue = New ListScaled
				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				FilterValueLast = 0
				FilterValueLastK1 = 0
				ValueLast = 0
				ValueLastK1 = 0
				'this is the factor A that will give the same bandwidth than a moving average with a flat windows of FilterRate points
				'see https://en.wikipedia.org/wiki/Exponential_smoothing  section: Comparison with moving average
				'this result come from the fact that the delay for a square window moving average is given by (N+1)/2 and 1/Alpha for an exponential filter
				A = CDbl((2 / (FilterRate + 1)))

				'Seek also:https://en.wikipedia.org/wiki/Low-pass_filter
				B = 1 - A
				IsValueInitial = False
				'MyValueSumForInit = 0
			End Sub

			Public Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
#If DebugPrediction Then
            Static IsHere As Boolean
            If IsHere = False Then
              IsHere = True
              Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
              Dim ThisResultActual = Me.Filter(Value)
              IsHere = False
              If ThisResultActual <> ThisResultPrediction Then
                Debugger.Break()
              End If
              Return ThisResultActual
            End If
#End If
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						FilterValueLast = Value
					End If
				End If
				'not sure with this
				'If IsValueInitial = False Then
				'  If MyListOfValue.Count < MyRate Then
				'    MyValueSumForInit = MyValueSumForInit + Value
				'    FilterValueLast = MyValueSumForInit / (MyListOfValue.Count + 1)
				'  End If
				'End If
				FilterValueLastK1 = FilterValueLast
				FilterValueLast = A * Value + B * FilterValueLast
				MyListOfValue.Add(FilterValueLast)
				ValueLastK1 = ValueLast
				ValueLast = Value
				Return FilterValueLast
			End Function

			Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Throw New NotImplementedException
				'ThisProbability = StockOption.StockPricePredictionInverse(I - ThisPositionForPrediction, ThisStockPrice, ThisGainValue, ThisVolatilityForBrownianStatistic, ThisFilterLPForPricePLLBrownianProbability.ToList(I))
				'ThisProbabilityInverseInSigmaOfThirdRatio = ((Measure.InverseCDFGaussian(0.5, 0.5 / 3, ThisProbability)))
				'ThisFilterPLLProbabilitySlow.Filter(ThisProbabilityInverseInSigmaOfThirdRatio)
				'ThisStockPriceProbabilitySlow.Add(ThisFilterPLLProbabilitySlow.FilterLast)

				'ThisProbability = StockOption.StockPricePredictionInverse(I - ThisPositionForPredictionFast, ThisStockPriceForFast, ThisGainValue, ThisVolatilityForBrownianStatistic, ThisFilterLPForPricePLLBrownianProbability.ToList(I))
				'ThisProbabilityInverseInSigmaOfThirdRatio = ((Measure.InverseCDFGaussian(0.5, 0.5 / 3, ThisProbability)))
				'ThisFilterPLLProbabilityFast.Filter(ThisProbabilityInverseInSigmaOfThirdRatio)
				'ThisStockPriceProbabilityFast.Add(ThisFilterPLLProbabilityFast.FilterLast)

			End Function

			Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			''' <summary>
			''' Special filtering that can be used to remove the delay starting at a specific point
			''' </summary>
			''' <param name="Value">The value to be filtered</param>
			''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
			''' <returns>The result</returns>
			''' <remarks></remarks>
			Public Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Dim ThisValues(0 To Value.Length - 1) As Double
				Dim I As Integer
				Dim J As Integer

				Dim ThisFilterLeft As New FilterLowPassExp(Me.Rate)
				Dim ThisFilterRight As New FilterLowPassExp(Me.Rate)
				Dim ThisFilterLeftItem As Double
				Dim ThisFilterRightItem As Double
				Dim ThisPriceVol As Double

				'filter from the left
				ThisFilterLeft.Filter(Value)
				'filter from the right the section with the reverse filtering
				For I = DelayRemovedToItem To 0 Step -1
					ThisFilterRight.Filter(Value(I))
				Next
				'the data in ThisFilterRightList is reversed
				'need to look at it in reverse order using J
				J = DelayRemovedToItem
				For I = 0 To Value.Length - 1
					ThisFilterLeftItem = ThisFilterLeft.ToList(I)
					If I > DelayRemovedToItem Then
						ThisPriceVol = ThisFilterLeftItem
					Else
						ThisFilterRightItem = ThisFilterRight.ToList(J)
						ThisPriceVol = (ThisFilterLeftItem + ThisFilterRightItem) / 2
					End If
					MyListOfValue.Add(ThisPriceVol)
					ThisValues(I) = ThisPriceVol
					J = J - 1
				Next
				Return ThisValues
			End Function

			Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Return 0.0
			End Function

			Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				If A > 0 Then
					Return (Value - B * FilterValueLastK1) / A
				Else
					Return Value
				End If
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
				With ThisPriceVol
					.LastPrevious = CSng(FilterValueLastK1)
					If Me.Last > .Last Then
						.High = CSng(Me.Last)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf Me.Last < .Last Then
						.Low = CSng(Me.Last)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
				With ThisPriceVol
					.LastPrevious = CSng(ValueLastK1)
					If Me.FilterLast > .Last Then
						.High = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf Me.FilterLast < .Last Then
						.Low = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
				Return CSng(Me.Filter(CDbl(Value)))
			End Function

			Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
				Dim ThisFilterValueLast As Double = FilterValueLast
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						ThisFilterValueLast = Value
					Else
						ThisFilterValueLast = FilterValueLast
					End If
				End If
				ThisFilterValueLast = A * Value + B * ThisFilterValueLast
				Return ThisFilterValueLast
			End Function

			Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
				Return Me.FilterPredictionNext(CDbl(Value))
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return FilterValueLast
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String Implements IFilter.Tag

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return Me.FilterLast.ToString
			End Function

#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function
			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
		End Class
#End Region
#Region "FilterAttackDecayExp"
		<Serializable()>
		Public Class FilterAttackDecayExp
			Private MyRateAttack As Integer
			Private MyRateDecay As Integer
			Private AForAttack As Double
			Private BForAttack As Double
			Private AForDecay As Double
			Private BForDecay As Double
			Private FilterValueLastK1 As Double
			Private FilterValueLast As Double
			Private ValueLast As Double
			Private ValueLastK1 As Double
			'Private MyValueSumForInit As Double
			Private IsValueInitial As Boolean
			Private MyListOfValue As ListScaled

			Public Sub New(ByVal FilterRateAttack As Double, ByVal FilterRateDecay As Double)
				MyListOfValue = New ListScaled
				If FilterRateAttack < 1 Then FilterRateAttack = 1
				MyRateAttack = CInt(FilterRateAttack)
				If FilterRateDecay < 1 Then FilterRateDecay = 1
				MyRateDecay = CInt(FilterRateAttack)

				FilterValueLast = 0
				FilterValueLastK1 = 0
				ValueLast = 0
				ValueLastK1 = 0

				AForAttack = CDbl((2 / (FilterRateAttack + 1)))
				BForAttack = 1 - AForAttack
				AForDecay = CDbl((2 / (FilterRateDecay + 1)))
				BForDecay = 1 - AForDecay
				IsValueInitial = False
				IsValueInitial = False

				'MyValueSumForInit = 0
			End Sub

			Public Sub New(ByVal FilterRateAttack As Integer, ByVal FilterRateDecay As Integer)
				Me.New(CDbl(FilterRateAttack), CDbl(FilterRateDecay))
			End Sub

			Public Sub New(ByVal FilterRateAttack As Integer, ByVal FilterRateDecay As Integer, ByVal ValueInitial As Double)
				Me.New(CDbl(FilterRateAttack), CDbl(FilterRateDecay))
				FilterValueLast = ValueInitial
				FilterValueLastK1 = FilterValueLast
				ValueLast = ValueInitial
				ValueLastK1 = ValueLast
				IsValueInitial = True
			End Sub

			Public Sub New(ByVal FilterRateAttack As Integer, ByVal FilterRateDecay As Integer, ByVal ValueInitial As Single)
				Me.New(FilterRateAttack, FilterRateDecay, CDbl(ValueInitial))
			End Sub

			Public Function Filter(ByVal Value As Double) As Double
#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual <> ThisResultPrediction Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						FilterValueLast = Value
						ValueLast = Value
						ValueLastK1 = ValueLast
					End If
				End If
				FilterValueLastK1 = FilterValueLast
				If (Value - ValueLast) >= 0 Then
					'use the attack variable
					FilterValueLast = AForAttack * Value + BForAttack * FilterValueLast
				Else
					'use the decay variable
					FilterValueLast = AForDecay * Value + BForDecay * FilterValueLast
				End If
				MyListOfValue.Add(FilterValueLast)
				ValueLastK1 = ValueLast
				ValueLast = Value
				Return FilterValueLast
			End Function

			Public Function Filter(ByRef Value() As Double) As Double()
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function FilterBackTo(ByRef Value As Double) As Double
				If (Value - ValueLast) >= 0 Then
					'use the attack variable
					If AForAttack > 0 Then
						Return (Value - BForAttack * FilterValueLastK1) / AForAttack
					Else
						Return Value
					End If
				Else
					'use the decay variable
					If AForDecay > 0 Then
						Return (Value - BForDecay * FilterValueLastK1) / AForDecay
					Else
						Return Value
					End If
				End If
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
				With ThisPriceVol
					.LastPrevious = CSng(FilterValueLastK1)
					If ValueLast > .Last Then
						.High = CSng(ValueLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf ValueLast < .Last Then
						.Low = CSng(ValueLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function LastToPriceVol() As IPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.Last))
				With ThisPriceVol
					.LastPrevious = CSng(ValueLastK1)
					If Me.FilterLast > .Last Then
						.High = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf Me.FilterLast < .Last Then
						.Low = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				Return ThisPriceVol
			End Function

			Public Function Filter(ByVal Value As Single) As Double
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function FilterPredictionNext(ByVal Value As Double) As Double
				Dim ThisFilterValueLast As Double = FilterValueLast
				Dim ThisValueLast = ValueLast
				If MyListOfValue.Count = 0 Then
					'initialization
					If IsValueInitial = False Then
						ThisFilterValueLast = Value
						ValueLast = Value
					End If
				End If
				If (Value - ValueLast) >= 0 Then
					'use the attack variable
					ThisFilterValueLast = AForAttack * Value + BForAttack * ThisFilterValueLast
				Else
					'use the decay variable
					ThisFilterValueLast = AForDecay * Value + BForDecay * ThisFilterValueLast
				End If
				Return ThisFilterValueLast
			End Function

			Public Function FilterPredictionNext(ByVal Value As Single) As Double
				Return Me.FilterPredictionNext(CDbl(Value))
			End Function

			Public Function FilterLast() As Double
				Return FilterValueLast
			End Function

			Public Function Last() As Double
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return (MyRateAttack + MyRateDecay) \ 2
				End Get
			End Property

			Public ReadOnly Property RateAttack As Integer
				Get
					Return MyRateAttack
				End Get
			End Property

			Public ReadOnly Property RateDecay As Integer
				Get
					Return MyRateDecay
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double)
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double()
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "StatisticalDistribution"
		Public Class StatisticalDistribution
			Implements IStatisticalDistribution

			'Private MyPDF() As Integer
			'Private MyLCR() As Integer
			'Private MyCDF() As Integer
			'Private MyFD() As Integer
			Private MyPDF As ExtensionService.VBArray(Of Integer)
			Private MyLCR As ExtensionService.VBArray(Of Integer)
			Private MyCDF As ExtensionService.VBArray(Of Integer)
			Private MyFD As ExtensionService.VBArray(Of Integer)
			Private MyListOfPDF As List(Of Double)
			Private MyListOfLCR As List(Of Double)
			Private MyListOfCDF As List(Of Double)
			Private MyListOfFD As List(Of Double)
			Private MyBucketLimitHigh As Integer
			Private MyBucketLimitLow As Integer
			Private MyBucketHigh As Integer
			Private MyBucketLow As Integer
			Private MyNumberBucket As Integer
			Private MyNumberPoint As Integer
			Private MyNumberPointMaximum As Integer
			Private MyBucketValueLast As Integer
			Private MyListOfWindowsFrame As ListWindowFrame
			Private MyStatisticalDistributionForVolatilityFunction As IStatisticalDistributionFunction
			Private MyMean As Double
			Private MyMeanOfX2 As Double
			Private MyVariance As Double

			''' <summary>
			''' Calculate the distribution statistic based on a bounded integer bin bucket
			''' </summary>
			''' <param name="BucketLimitLow">The minimum value for an integer bin bucket</param>
			''' <param name="BucketLimitHigh">The maximum value for an integer bin bucket</param>
			''' <remarks></remarks>
			Public Sub New(ByVal BucketLimitLow As Integer, ByVal BucketLimitHigh As Integer)
				Me.New(0, New StatisticalDistributionFunctionBasic(BucketLimitLow, BucketLimitHigh))
			End Sub

			''' <summary>
			''' Calculate the distribution statistic based on a bounded integer bin bucket
			''' </summary>
			''' <param name="NumberPointMaximum">
			''' Limit the calculation to this number of points. If the number is less or egual to zero the calculation use all the points.
			'''  </param>
			''' <param name="BucketLimitLow">The minimum value for an integer bin bucket</param>
			''' <param name="BucketLimitHigh">The maximum value for an integer bin bucket</param>
			''' <remarks></remarks>
			Public Sub New(ByVal NumberPointMaximum As Integer, ByVal BucketLimitLow As Integer, ByVal BucketLimitHigh As Integer)
				Me.New(NumberPointMaximum, New StatisticalDistributionFunctionBasic(BucketLimitLow, BucketLimitHigh))
			End Sub

			''' <summary>
			''' 
			''' </summary>
			''' <param name="StatisticalDistributionFunction"></param>
			''' <remarks></remarks>
			Public Sub New(ByVal NumberPointMaximum As Integer, ByVal StatisticalDistributionFunction As IStatisticalDistributionFunction)
				MyStatisticalDistributionForVolatilityFunction = StatisticalDistributionFunction

				'extract the data for direct access
				MyBucketLimitHigh = MyStatisticalDistributionForVolatilityFunction.BucketLimitHigh
				MyBucketLimitLow = MyStatisticalDistributionForVolatilityFunction.BucketLimitLow
				MyNumberBucket = MyStatisticalDistributionForVolatilityFunction.NumberBucket
				MyNumberPointMaximum = NumberPointMaximum

				MyPDF = New ExtensionService.VBArray(Of Integer)(MyBucketLimitLow, MyBucketLimitHigh)
				MyLCR = New ExtensionService.VBArray(Of Integer)(MyBucketLimitLow, MyBucketLimitHigh)
				MyCDF = New ExtensionService.VBArray(Of Integer)(MyBucketLimitLow, MyBucketLimitHigh)
				MyFD = New ExtensionService.VBArray(Of Integer)(MyBucketLimitLow, MyBucketLimitHigh)
				MyListOfPDF = New List(Of Double)
				MyListOfLCR = New List(Of Double)
				MyListOfCDF = New List(Of Double)
				MyListOfFD = New List(Of Double)
				If MyNumberPointMaximum > 0 Then
					MyListOfWindowsFrame = New ListWindowFrame(MyNumberPointMaximum)
				Else
					MyListOfWindowsFrame = Nothing
				End If
			End Sub


			Public ReadOnly Property BucketHigh As Integer Implements IStatisticalDistribution.BucketHigh
				Get
					Return MyBucketHigh
				End Get
			End Property

			Public ReadOnly Property BucketLow As Integer Implements IStatisticalDistribution.BucketLow
				Get
					Return MyBucketLow
				End Get
			End Property

			Public ReadOnly Property BucketLimitHigh As Integer Implements IStatisticalDistribution.BucketLimitHigh
				Get
					Return MyBucketLimitHigh
				End Get
			End Property

			Public ReadOnly Property BucketLimitLow As Integer Implements IStatisticalDistribution.BucketLimitLow
				Get
					Return MyBucketLimitLow
				End Get
			End Property

			Public Sub BucketFill(Value As Double) Implements IStatisticalDistribution.BucketFill
				Dim ThisBucketValue As Integer
				Dim ThisBucketValueLast As Integer
				Dim I As Integer

				ThisBucketValue = MyStatisticalDistributionForVolatilityFunction.ToBucket(Value)
				If MyNumberPoint = 0 Then
					MyBucketValueLast = ThisBucketValue
					MyBucketHigh = ThisBucketValue
					MyBucketLow = ThisBucketValue
				End If
				MyPDF(ThisBucketValue) = MyPDF(ThisBucketValue) + 1
				Select Case ThisBucketValue - MyBucketValueLast
					Case Is > 0
						For I = MyBucketValueLast To ThisBucketValue - 1
							MyLCR(I) = MyLCR(I) + 1
						Next
					Case Is < 0
						For I = MyBucketValueLast To ThisBucketValue + 1 Step -1
							MyLCR(I) = MyLCR(I) + 1
						Next
				End Select
				MyNumberPoint = MyNumberPoint + 1
				MyBucketValueLast = ThisBucketValue
				'check if we need to remove the old elements
				If MyNumberPointMaximum > 0 Then
					'this is a windows statistic distribution and some old elements may need to be removed
					MyListOfWindowsFrame.Add(ThisBucketValue)
					If MyListOfWindowsFrame.ItemRemoved.HasValue Then
						ThisBucketValue = CInt(MyListOfWindowsFrame.ItemFirst.Value)
						ThisBucketValueLast = CInt(MyListOfWindowsFrame.ItemRemoved.Value)
						MyPDF(ThisBucketValueLast) = MyPDF(ThisBucketValueLast) - 1
						Select Case ThisBucketValue - ThisBucketValueLast
							Case Is > 0
								For I = ThisBucketValueLast To ThisBucketValue - 1
									MyLCR(I) = MyLCR(I) - 1
								Next
							Case Is < 0
								For I = ThisBucketValueLast To ThisBucketValue + 1 Step -1
									MyLCR(I) = MyLCR(I) - 1
								Next
						End Select
						MyNumberPoint = MyNumberPoint - 1
					End If
					MyBucketHigh = CInt(MyListOfWindowsFrame.ItemHigh.Value)
					MyBucketLow = CInt(MyListOfWindowsFrame.ItemLow.Value)
				Else
					'add to the distribution
					If ThisBucketValue > MyBucketHigh Then
						MyBucketHigh = ThisBucketValue
					End If
					If ThisBucketValue < MyBucketHigh Then
						MyBucketLow = ThisBucketValue
					End If
				End If
			End Sub

			Public ReadOnly Property Mean As Double Implements IStatisticalDistribution.Mean
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Private ReadOnly Property MeanOfSquare As Double Implements IStatisticalDistribution.MeanOfSquare
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property NumberOfBucket As Integer Implements IStatisticalDistribution.NumberOfBucket
				Get
					Return MyNumberBucket
				End Get
			End Property

			Public ReadOnly Property NumberPoint As Integer Implements IStatisticalDistribution.NumberPoint
				Get
					Return MyNumberPoint
				End Get
			End Property

			Public ReadOnly Property StandardDeviation As Double Implements IStatisticalDistribution.StandardDeviation
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property ToListOfCDF As System.Collections.Generic.IList(Of Double) Implements IStatisticalDistribution.ToListOfCDF
				Get
					Return MyListOfCDF
				End Get
			End Property

			Public ReadOnly Property ToListOfFD As System.Collections.Generic.IList(Of Double) Implements IStatisticalDistribution.ToListOfFD
				Get
					Return MyListOfFD
				End Get
			End Property

			Public ReadOnly Property ToListOfLCR As System.Collections.Generic.IList(Of Double) Implements IStatisticalDistribution.ToListOfLCR
				Get
					Return MyListOfLCR
				End Get
			End Property

			Public ReadOnly Property ToListOfPDF As System.Collections.Generic.IList(Of Double) Implements IStatisticalDistribution.ToListOfPDF
				Get
					Return MyListOfPDF
				End Get
			End Property

			Public ReadOnly Property Variance As Double Implements IStatisticalDistribution.Variance
				Get
					Throw New NotImplementedException
				End Get
			End Property

			Public ReadOnly Property AsIStatisticalDistribution As IStatisticalDistribution Implements IStatisticalDistribution.AsIStatisticalDistribution
				Get
					Return Me
				End Get
			End Property

			Public Sub Refresh(Optional Type As IStatisticalDistribution.enuRefreshType = IStatisticalDistribution.enuRefreshType.PerCent) Implements IStatisticalDistribution.Refresh
				Dim I As Integer
				Dim ThisCDFSum As Double
				Dim ThisPDF As Double
				Dim ThisLCR As Double
				Dim ThisFD As Double
				Dim ThisXPDF As Double

				ThisCDFSum = 0
				ThisFD = 0
				MyMean = 0
				MyMeanOfX2 = 0
				MyListOfPDF.Clear()
				MyListOfCDF.Clear()
				MyListOfLCR.Clear()
				MyListOfFD.Clear()
				Select Case Type
					Case IStatisticalDistribution.enuRefreshType.NumberOfPoint
						For I = MyBucketLow To MyBucketHigh
							ThisPDF = MyPDF(I)
							ThisXPDF = CDbl(I) * (ThisPDF / MyNumberPoint)
							MyMean = MyMean + ThisXPDF
							MyMeanOfX2 = MyMeanOfX2 + (CDbl(I) * ThisXPDF)
							'divide by 2 to because we assume symetry and measure the positive and negative LCR in this object
							'and LCR is by definition the negative LCR only
							ThisLCR = MyLCR(I) / 2
							ThisCDFSum = ThisCDFSum + ThisPDF
							MyListOfPDF.Add(ThisPDF)
							MyListOfCDF.Add(ThisCDFSum)
							MyListOfLCR.Add(ThisLCR)
							'assuming that we are dealing with daily data 
							'ThisCDFSum / ThisLCR is the fading duration in number of days
							'it may be appropriate to represent the value on a log scale per year
							'output the fading duration in the log of the Number of day/per year 
							'note that the LCR is zero at the edge of the distribution
							If ThisLCR > 0 Then
								ThisFD = 10 * Math.Log10((ThisCDFSum / ThisLCR) / NUMBER_WORKDAY_PER_YEAR)
							Else
								'use the previous value
							End If
							MyListOfFD.Add(ThisFD)
						Next
					Case IStatisticalDistribution.enuRefreshType.PerCent
						For I = MyBucketLow To MyBucketHigh
							ThisPDF = MyPDF(I) / MyNumberPoint
							ThisXPDF = CDbl(I) * ThisPDF
							MyMean = MyMean + ThisXPDF
							MyMeanOfX2 = MyMeanOfX2 + CDbl(I) * ThisXPDF
							'divide by 2 to because we assume symetry and measure the positive and negative LCR in this object
							'and LCR is by definition the negative LCR only
							ThisLCR = MyLCR(I) / MyNumberPoint / 2
							ThisCDFSum = ThisCDFSum + ThisPDF
							MyListOfPDF.Add(ThisPDF)
							MyListOfCDF.Add(ThisCDFSum)
							MyListOfLCR.Add(ThisLCR)
							'assuming that we are dealing with daily data 
							'ThisCDFSum / ThisLCR is the fading duration in number of days
							'it may be appropriate to represent the value on a log scale per year
							'output the fading duration in the log of the Number of day/per year 
							'note that the LCR is zero at the edge of the distribution
							If ThisLCR > 0 Then
								ThisFD = 10 * Math.Log10((ThisCDFSum / ThisLCR) / NUMBER_WORKDAY_PER_YEAR)
							Else
								'use the previous value
							End If
							MyListOfFD.Add(ThisFD)
						Next
					Case IStatisticalDistribution.enuRefreshType.ArrayStandard
						MyListOfPDF.Add(Me.NumberPoint)
						MyListOfCDF.Add(Me.NumberPoint)
						MyListOfLCR.Add(Me.NumberPoint)
						MyListOfFD.Add(Me.NumberPoint)
						ThisFD = 0
						For I = MyBucketLimitLow To MyBucketLimitHigh
							ThisPDF = MyPDF(I)
							ThisXPDF = CDbl(I) * (ThisPDF / MyNumberPoint)
							MyMean = MyMean + ThisXPDF
							MyMeanOfX2 = MyMeanOfX2 + (CDbl(I) * ThisXPDF)
							'divide by 2 to because we assume symetry and measure the positive and negative LCR in this object
							'and LCR is by definition the negative LCR only
							ThisLCR = MyLCR(I) / 2
							ThisCDFSum = ThisCDFSum + ThisPDF
							MyListOfPDF.Add(ThisPDF)
							MyListOfCDF.Add(ThisCDFSum)
							MyListOfLCR.Add(ThisLCR)
							'assuming that we are dealing with daily data 
							'ThisCDFSum / ThisLCR is the fading duration in number of days
							'it may be appropriate to represent the value on a log scale per year
							'output the fading duration in the log of the Number of day/per year 
							'note that the LCR is zero at the edge of the distribution
							If ThisLCR > 0 Then
								ThisFD = 10 * Math.Log10((ThisCDFSum / ThisLCR) / NUMBER_WORKDAY_PER_YEAR)
							Else
								'use the previous value
							End If
							MyListOfFD.Add(ThisFD)
						Next
				End Select
			End Sub

			Public Property Tag As String Implements IStatisticalDistribution.Tag
		End Class

		Public Class StatisticalDistributionFunctionBasic
			Implements IStatisticalDistributionFunction


			Private MyBucketLimitLow As Integer
			Private MyBucketLimitHigh As Integer
			Private MyBucketLow As Integer
			Private MyBucketHigh As Integer
			Private MyNumberBucket As Integer

			Public Sub New(ByVal BucketLimitLow As Integer, ByVal BucketLimitHigh As Integer)
				If MyBucketLimitLow < MyBucketLimitHigh Then
					Throw New InvalidConstraintException
				End If
				MyBucketLimitLow = BucketLimitLow
				MyBucketLimitHigh = BucketLimitHigh
				MyNumberBucket = MyBucketLimitHigh - MyBucketLimitLow + 1
			End Sub

			Public ReadOnly Property BucketLimitHigh As Integer Implements IStatisticalDistributionFunction.BucketLimitHigh
				Get
					Return MyBucketLimitHigh
				End Get
			End Property

			Public ReadOnly Property BucketLimitLow As Integer Implements IStatisticalDistributionFunction.BucketLimitLow
				Get
					Return MyBucketLimitLow
				End Get
			End Property

			Public ReadOnly Property NumberBucket As Integer Implements IStatisticalDistributionFunction.NumberBucket
				Get
					Return MyNumberBucket
				End Get
			End Property

			Public Overridable Function ToBucket(Value As Double) As Integer Implements IStatisticalDistributionFunction.ToBucket
				Dim ThisBucketValue As Integer
				ThisBucketValue = CInt(Value)
				If ThisBucketValue < MyBucketLimitLow Then
					ThisBucketValue = MyBucketLimitLow
				End If
				If ThisBucketValue > MyBucketLimitHigh Then
					ThisBucketValue = MyBucketLimitHigh
				End If
				Return ThisBucketValue
			End Function

			Public Overridable Function FromBucket(Index As Integer) As Double Implements IStatisticalDistributionFunction.FromBucket
				If Index < 0 Then
					Index = 0
				End If
				If Index > MyNumberBucket - 1 Then
					Index = MyNumberBucket - 1
				End If
				Return CDbl(Index)
			End Function

			Public ReadOnly Property BucketHigh As Integer Implements IStatisticalDistributionFunction.BucketHigh
				Get
					Return MyBucketHigh
				End Get
			End Property

			Public ReadOnly Property BucketLow As Integer Implements IStatisticalDistributionFunction.BucketLow
				Get
					Return MyBucketLow
				End Get
			End Property

			Public Overrides Function ToString() As String
				Return TypeName(Me)
			End Function
		End Class

		Public Class StatisticalDistributionFunctionLog
			Implements IStatisticalDistributionFunction

			Private MyStatisticalDistributionForVolatilityFunctionBasic As StatisticalDistributionFunctionBasic

			Public Sub New(ByVal BucketLimitLow As Integer, ByVal BucketLimitHigh As Integer)
				MyStatisticalDistributionForVolatilityFunctionBasic = New StatisticalDistributionFunctionBasic(BucketLimitLow, BucketLimitHigh)
			End Sub

			Public ReadOnly Property BucketHigh As Integer Implements IStatisticalDistributionFunction.BucketHigh
				Get
					Return MyStatisticalDistributionForVolatilityFunctionBasic.BucketHigh
				End Get
			End Property

			Public ReadOnly Property BucketLimitHigh As Integer Implements IStatisticalDistributionFunction.BucketLimitHigh
				Get
					Return MyStatisticalDistributionForVolatilityFunctionBasic.BucketLimitHigh
				End Get
			End Property

			Public ReadOnly Property BucketLimitLow As Integer Implements IStatisticalDistributionFunction.BucketLimitLow
				Get
					Return MyStatisticalDistributionForVolatilityFunctionBasic.BucketLimitLow
				End Get
			End Property

			Public ReadOnly Property BucketLow As Integer Implements IStatisticalDistributionFunction.BucketLow
				Get
					Return MyStatisticalDistributionForVolatilityFunctionBasic.BucketLow
				End Get
			End Property

			Public Function ToBucket(Value As Double) As Integer Implements IStatisticalDistributionFunction.ToBucket
				Return MyStatisticalDistributionForVolatilityFunctionBasic.ToBucket(10 * Math.Log10(Value))
			End Function

			Public Function FromBucket(Index As Integer) As Double Implements IStatisticalDistributionFunction.FromBucket
				Return 10 ^ (MyStatisticalDistributionForVolatilityFunctionBasic.FromBucket(Index) / 10)
			End Function

			Public ReadOnly Property NumberBucket As Integer Implements IStatisticalDistributionFunction.NumberBucket
				Get
					Return MyStatisticalDistributionForVolatilityFunctionBasic.NumberBucket
				End Get
			End Property

			Public Overloads Overrides Function ToString() As String
				Return TypeName(Me)
			End Function
		End Class

#End Region

#Region "Filter Delay"
		''' <summary>
		''' This Filter delay the incoming stream by the specified amount
		''' </summary>
		''' <remarks></remarks>
		<Serializable()>
		Public Class FilterSignalDelay
			Implements IFilter
			Implements IRegisterKey(Of String)

			Private MyRate As Integer
			Private FilterValueLastK1 As Double
			Private FilterValueLast As Double

			Private ValueLast As Double
			Private ValueLastK1 As Double
			Private MyListOfValue As List(Of Double)

			''' <summary>
			''' Calculate the statistical information from all value 
			''' </summary>
			''' <remarks></remarks>
			Public Sub New()
				Me.New(0)
			End Sub

			Public Sub New(ByVal FilterRate As Integer)
				MyListOfValue = New List(Of Double)
				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				FilterValueLast = 0
				FilterValueLastK1 = 0
				ValueLast = 0
				ValueLastK1 = 0
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal StartPoint As Integer)
				Me.New(FilterRate)
			End Sub

			Public Function Filter(Value As Double) As Double Implements IFilter.Filter
				If MyListOfValue.Count = 0 Then
					'initialization
					FilterValueLast = Value
				End If
				FilterValueLastK1 = FilterValueLast
				MyListOfValue.Add(Value)
				If (MyListOfValue.Count - 1) >= MyRate Then
					FilterValueLast = MyListOfValue(0)
					MyListOfValue.RemoveAt(0)
				Else
					MyListOfValue.Add(Value)
					FilterValueLast = MyListOfValue(0)
				End If
				ValueLastK1 = ValueLast
				ValueLast = Value
				Return FilterValueLast
			End Function

			Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Return Me.Filter(Value.Last)
			End Function

			Public Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				Dim ThisResult(0 To Value.Length - 1) As Double
				For I = 0 To Value.Length - 1
					ThisResult(I) = Me.Filter(Value(I))
				Next
				Return ThisResult
			End Function

			Public Function Filter(ByRef Value() As Double, DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Throw New NotImplementedException
			End Function

			Public Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Throw New NotImplementedException
			End Function

			Public Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				Throw New NotImplementedException
			End Function

			Public Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Me.FilterLast))
				With ThisPriceVol
					.LastPrevious = CSng(FilterValueLastK1)
					If Me.FilterLast > .High Then
						.High = CSng(Me.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
					If Me.FilterLast > .Low Then
						.Low = CSng(Me.FilterLast)
					End If
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
				End With
				Return ThisPriceVol
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Throw New NotSupportedException
			End Function

			Public Function Filter(Value As Single) As Double Implements IFilter.Filter
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
				Throw New NotSupportedException
			End Function

			Public Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
				Return Me.FilterPredictionNext(CDbl(Value))
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return FilterValueLast
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(MinValueInitial As Double, MaxValueInitial As Double, ScaleToMinValue As Double, ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Property Tag As String Implements IFilter.Tag

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return Me.FilterLast.ToString
			End Function

#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function
			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
		End Class
#End Region
#Region "FilterStochasticHHS"
		''' <summary>
		'''   http://tlc.thinkorswim.ca/center/charting/studies/studies-library/G-L/HHLLS.html
		''' 
		''' The HHLLS (Higher High Lower Low Stochastic) study is a momentum-based technical indicator 
		''' developed by Vitali Apirine. It consists of two stochastic lines, the calculation of which 
		''' is inspired by StochasticFull and Williams%R. The main purpose of HHLLS is to recognize trend 
		''' behavior: emergence, corrections, and reversals. This can be done by spotting well-known signals:
		'''  divergences, crossovers, and overbought/oversold conditions.
		''' 
		'''The HHLLS indicator is calculated as follows. First, the study calculates two values, HH and LL:
		'''    HH. For bars with the high price greater than the previous high (otherwise 0):
		'''    HH = (Hc - Hl)/(Hh-Hl),
		'''
		''' where Hc is the current bar’s high price, Hl and Hh are the lowest and the highest high prices over
		'''a certain period.
		'''
		''' LL. For bars with the low price less than the previous low (otherwise 0):
		'''
		'''   LL = (Lh - Lc)/(Lh-Ll),
		'''where Lc is the current low price, Ll and Lh are the lowest and the highest low prices over
		'''	a certain period. Both variables are then smoothed with a moving average, resulting in two main 
		''' plots HHS and LLS. 
		'''	The values of these plots range from zero to 100, rarely reaching either boundary. 
		'''	Thus, by default, overbought and oversold levels are set at 60 and 10, respectively.
		'''
		'''The divergence between the price and HHLLS plots might prove useful in recognition of trend 
		'''reversals or corrections. 
		'''A bearish divergence is indicated when the price is trending up but HHS fails to confirm this move. 
		'''Conversely, price making a new low when LLS goes up can be considered a bullish divergence. 
		'''Divergence of either type may need additional confirmation: a signal may prove stronger 
		'''when one of the lines forms the divergence and the other crosses above the level of 50.
		'''
		'''Another technique of using HHLLS is analyzing the behavior of the main plots in relation to each other. 
		'''When HHS is rising while the LLS is making new lows, the price may be entering an uptrend. 
		'''The opposite situation may lead to emergence of a downtrend. 
		'''The crossovers of the two lines may also indicate important trading signals.
		''' 
		''' Speudo implementation:
		''' 
		''' HIGHER HIGHS AND LOWER LOWS
		''' Author: Vitali Aprine, TASC Feb 2016
		''' Coded by: Richard Denning, 12/10/2015
		''' TradersEdgeSystems com
		''' 
		''' </summary>
		<Serializable()>
		Public Class FilterStochasticHHS
			Implements IFilter
			Implements IRegisterKey(Of String)


			Private MyRate As Integer
			Private FilterValueLast As Double
			Private ValueLast As Double
			Private MyListOfValue As ListScaled
			Private MyFilterPost As IFilter
			Private MyListOfWindowFrame As ListWindowFrame

			Public Sub New()


			End Sub

			Public Sub New(
										ByVal FilterRate As Double,
										Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)
				Me.New(FilterRate, FilterRate, FilterOutputType)
			End Sub

			Public Sub New(
										ByVal FilterRate As Double,
										ByVal PostFilterRate As Double,
										Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)

				MyListOfValue = New ListScaled
				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				FilterValueLast = 0
				ValueLast = 0.0
				MyListOfWindowFrame = New ListWindowFrame(CInt(FilterRate))
				Select Case FilterOutputType
					Case FilterStochasticHHSLLS.FilterOutputType.Exponential
						MyFilterPost = New FilterLowPassExp(PostFilterRate)
					Case FilterStochasticHHSLLS.FilterOutputType.Hull
						MyFilterPost = New FilterLowPassExpHull(PostFilterRate)
					Case FilterStochasticHHSLLS.FilterOutputType.PLL
						MyFilterPost = New FilterLowPassPLL(PostFilterRate)
				End Select
			End Sub

			Private Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
				'input type not supported
				Throw New NotImplementedException
			End Function

			''' <summary>
			'''The HHLLS indicator is calculated as follows. First, the study calculates two values, HH and LL:
			'''    HH. For bars with the high price greater than the previous high (otherwise 0):
			'''    HH = (Hc - Hl)/(Hh-Hl),
			'''
			''' where Hc is the current bar’s high price, Hl and Hh are the lowest and the highest high prices over a certain period.
			'''
			''' LL. For bars with the low price less than the previous low (otherwise 0):
			'''
			'''   LL = (Lh - Lc)/(Lh-Ll),
			'''where Lc is the current low price, Ll and Lh are the lowest and the highest low prices over a certain period.
			'''
			'''Both variables are then smoothed with a moving average, resulting in two main plots HHS and LLS. 
			'''The values of these plots range from zero to 100, rarely reaching either boundary. 
			'''Thus, by default, overbought and oversold levels are set at 60 and 10, respectively.
			''' </summary>
			''' <param name="Value"></param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Dim ThisDirectionPressure As Double
				Dim ThisRange As Double
				Dim ThisHigh As Double
				Dim ThisLow As Double
				Dim ThisValue As Double = Value.High

				If MyListOfWindowFrame.Count = 0 Then
					ValueLast = ThisValue
				End If
				MyListOfWindowFrame.Add(ThisValue)
				ThisHigh = MyListOfWindowFrame.ItemHigh.Value
				ThisLow = MyListOfWindowFrame.ItemLow.Value
				ThisRange = ThisHigh - ThisLow
				If ThisRange > 0 Then
					If ThisValue > ValueLast Then
						ThisDirectionPressure = (ThisValue - ThisLow) / ThisRange
						'ThisDirectionPressure = 0.5 * (1 + (ThisValue - ThisLow) / ThisRange)
					Else
						'ThisDirectionPressure = 0.5 * ((ThisValue - ThisLow) / ThisRange)
						ThisDirectionPressure = 0.0
					End If
				Else
					ThisDirectionPressure = 0.5
				End If
				FilterValueLast = MyFilterPost.Filter(ThisDirectionPressure)
				MyListOfValue.Add(FilterValueLast)
				ValueLast = ThisValue
				Return FilterValueLast
			End Function
#Region "Friend Function"
			Friend ReadOnly Property ToFilterPost As IFilter
				Get
					Return MyFilterPost
				End Get
			End Property

			Friend ReadOnly Property ToListOfWindowFrame As ListWindowFrame
				Get
					Return MyListOfWindowFrame
				End Get
			End Property

			Friend Function Filter(ByVal InputValueLast As Double, ByVal Result As Double) As Double
				FilterValueLast = Result
				MyListOfValue.Add(FilterValueLast)
				ValueLast = InputValueLast
				Return FilterValueLast
			End Function
#End Region
			Private Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				'input type not supported
				Throw New NotImplementedException
			End Function

			''' <summary>
			''' Special filtering that can be used to remove the delay starting at a specific point
			''' </summary>
			''' <param name="Value">The value to be filtered</param>
			''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
			''' <returns>The result</returns>
			''' <remarks></remarks>
			Private Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Throw New NotImplementedException
			End Function

			Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Return 0.0
			End Function

			Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				Throw New NotImplementedException
			End Function

			Private Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Throw New NotImplementedException
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Return CType(New PriceVol(CSng(ValueLast)), IPriceVol)
			End Function

			Private Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
				Throw New NotImplementedException
			End Function

			Private Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Private Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return FilterValueLast
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return ValueLast
			End Function

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String Implements IFilter.Tag

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return Me.FilterLast.ToString
			End Function
#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function
			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
		End Class
#End Region
#Region "FilterStochasticLLS"
		<Serializable()>
		Public Class FilterStochasticLLS
			Inherits FilterStochasticHHS

			Public Sub New(ByVal FilterRate As Double, Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)
				MyBase.New(FilterRate, FilterRate, FilterOutputType)
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByVal PostFilterRate As Double, Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)
				MyBase.New(FilterRate, PostFilterRate, FilterOutputType)
			End Sub

			Public Overloads Function Filter(Value As IPriceVol) As Double
				Dim ThisDirectionPressure As Double
				Dim ThisRange As Double
				Dim ThisHigh As Double
				Dim ThisLow As Double
				Dim ThisValueLow As Double = Value.Low
				Dim ThisValueLowLast As Double = Me.Last
				Dim ThisFilterPost As Double

				If Me.Count = 0 Then
					ThisValueLowLast = ThisValueLow
				End If
				If Me.Count = 1000 Then
					ThisValueLowLast = ThisValueLowLast
				End If
				With Me.ToListOfWindowFrame
					.Add(ThisValueLow)
					ThisHigh = .ItemHigh.Value
					ThisLow = .ItemLow.Value
				End With
				ThisRange = ThisHigh - ThisLow
				If ThisRange > 0 Then
					If ThisValueLow < ThisValueLowLast Then
						ThisDirectionPressure = (ThisHigh - ThisValueLow) / ThisRange
					Else
						ThisDirectionPressure = 0.0
					End If
				Else
					ThisDirectionPressure = 0.5
				End If
				ThisFilterPost = Me.ToFilterPost.Filter(ThisDirectionPressure)
				Return Me.Filter(ThisValueLow, ThisFilterPost)
			End Function
		End Class
#End Region
#Region "FilterStochasticHHSLLS"
		<Serializable()>
		Public Class FilterStochasticHHSLLS
			Inherits FilterStochasticHHS

			Private MyFilterStochasticHHS As FilterStochasticHHS
			Private MyFilterStochasticLLS As FilterStochasticLLS


			Public Enum FilterOutputType
				Exponential
				Hull
				PLL
			End Enum



			Public Sub New(ByVal FilterRate As Double, Optional ByVal FilterOutputType As FilterOutputType = FilterOutputType.Exponential)
				MyBase.New(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochasticHHS = New FilterStochasticHHS(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochasticLLS = New FilterStochasticLLS(FilterRate, FilterRate, FilterOutputType)
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByVal PostFilterRate As Double, Optional ByVal FilterOutputType As FilterOutputType = FilterOutputType.Exponential)
				MyBase.New(FilterRate, PostFilterRate, FilterOutputType)
				MyFilterStochasticHHS = New FilterStochasticHHS(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochasticLLS = New FilterStochasticLLS(FilterRate, FilterRate, FilterOutputType)
			End Sub

			Public Overloads Function Filter(Value As IPriceVol) As Double
				Dim ThisFilterValueHHS As Double
				Dim ThisFilterValueLLS As Double
				Dim ThisFilterSum As Double

				ThisFilterValueHHS = MyFilterStochasticHHS.Filter(Value)
				ThisFilterValueLLS = MyFilterStochasticLLS.Filter(Value)
				ThisFilterSum = ThisFilterValueHHS + ThisFilterValueLLS
				If ThisFilterSum > 0 Then
					Return Me.Filter((Value.High + Value.Low) / 2, ThisFilterValueHHS / ThisFilterSum)
				Else
					Return Me.Filter((Value.High + Value.Low) / 2, 0.5)
				End If
			End Function
		End Class
#End Region
#Region "FilterStochasticDeltaHHSLLS"
		<Serializable()>
		Public Class FilterStochasticDeltaHHSLLS
			Inherits FilterStochasticHHS

			Private MyFilterStochasticHHSLLS As FilterStochasticHHSLLS
			Private MyFilterStochastic As FilterStochastic
			Private MyFilterStatistic As FilterStatistical
			Private ThisNormalDist As MathNet.Numerics.Distributions.Normal

			Public Sub New(ByVal FilterRate As Double, Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)
				MyBase.New(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochasticHHSLLS = New FilterStochasticHHSLLS(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochastic = New FilterStochastic(MyFilterStochasticHHSLLS.Rate, MyFilterStochasticHHSLLS.Rate, 2) With {.Tag = "FilterStochasticDeltaHHSLLS", .IsFilterPeak = True, .IsFilterRange = True}
				MyFilterStatistic = New FilterStatistical(MyFilterStochasticHHSLLS.Rate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				ThisNormalDist = New MathNet.Numerics.Distributions.Normal(0, 1)
			End Sub

			Public Sub New(ByVal FilterRate As Double, ByVal PostFilterRate As Double, Optional ByVal FilterOutputType As FilterStochasticHHSLLS.FilterOutputType = FilterStochasticHHSLLS.FilterOutputType.Exponential)
				MyBase.New(FilterRate, PostFilterRate, FilterOutputType)
				MyFilterStochasticHHSLLS = New FilterStochasticHHSLLS(FilterRate, FilterRate, FilterOutputType)
				MyFilterStochastic = New FilterStochastic(MyFilterStochasticHHSLLS.Rate, MyFilterStochasticHHSLLS.Rate, 2) With {.Tag = "FilterStochasticDeltaHHSLLS", .IsFilterPeak = True, .IsFilterRange = True}
				MyFilterStatistic = New FilterStatistical(MyFilterStochasticHHSLLS.Rate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				ThisNormalDist = New MathNet.Numerics.Distributions.Normal(0, 1)
			End Sub

			Public Overloads Function Filter(Value As IPriceVol) As Double
				Dim ThisFilterStatisticLast As IStatistical
				Dim ThisFilterDiff As Double

				MyFilterStochasticHHSLLS.Filter(Value)
				MyFilterStochastic.Filter(Value)

				ThisFilterDiff = MyFilterStochasticHHSLLS.FilterLast - MyFilterStochastic.FilterLast
				ThisFilterStatisticLast = MyFilterStatistic.Filter(ThisFilterDiff)
				If ThisFilterStatisticLast.StandardDeviation > 0 Then
					Return Me.Filter(Value.Last, ThisNormalDist.CumulativeDistribution(ThisFilterDiff / ThisFilterStatisticLast.StandardDeviation))
				Else
					Return Me.Filter(Value.Last, 0.0)
				End If
			End Function
		End Class
#End Region
#Region "FilterUltimateOscillator"
		'''<summary>
		'''see:
		''' http://stockcharts.com/school/doku.php?id=chart_school:technical_indicators:ultimate_oscillator
		''' https://en.wikipedia.org/wiki/Ultimate_oscillator
		'''BP = Close - Minimum(Low or Prior Close).
		'''TR = Maximum(High or Prior Close)  -  Minimum(Low or Prior Close)
		'''Average7 = (7-period BP Sum) / (7-period TR Sum)
		'''Average14 = (14-period BP Sum) / (14-period TR Sum)
		'''Average28 = (28-period BP Sum) / (28-period TR Sum)
		'''UO = 100 x [(4 x Average7)+(2 x Average14)+Average28]/(4+2+1)
		''' </summary>
		''' <remarks>The current implementation support any rate or output filter</remarks>
		<Serializable()>
		Public Class FilterUltimateOscillatorExp
			Implements IFilter
			Implements IRegisterKey(Of String)

			Public Enum FilterUltimateOscillatorType
				Exponential
				Hull
				PLL
			End Enum


			Private MyRate As Integer
			Private FilterValueLast As Double
			Private ValueLast As IPriceVol
			Private MyListOfValue As ListScaled
			Private MyFilterBP1 As IFilter
			Private MyFilterBP2 As IFilter
			Private MyFilterBP3 As IFilter
			Private MyFilterTR1 As IFilter
			Private MyFilterTR2 As IFilter
			Private MyFilterTR3 As IFilter
			Private MyFilterPost As IFilter

			Public Sub New(ByVal FilterRate As Double, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential)
				MyListOfValue = New ListScaled
				If FilterRate < 1 Then FilterRate = 1
				MyRate = CInt(FilterRate)
				FilterValueLast = 0
				ValueLast = New PriceVol(0)
				Select Case FilterType
					Case FilterUltimateOscillatorType.Exponential
						MyFilterBP1 = New FilterLowPassExp(FilterRate)
						MyFilterBP2 = New FilterLowPassExp(2 * FilterRate)
						MyFilterBP3 = New FilterLowPassExp(4 * FilterRate)
						MyFilterTR1 = New FilterLowPassExp(MyFilterBP1.Rate)
						MyFilterTR2 = New FilterLowPassExp(MyFilterBP2.Rate)
						MyFilterTR3 = New FilterLowPassExp(MyFilterBP3.Rate)
					Case FilterUltimateOscillatorType.Hull
						MyFilterBP1 = New FilterLowPassExpHull(FilterRate)
						MyFilterBP2 = New FilterLowPassExpHull(2 * FilterRate)
						MyFilterBP3 = New FilterLowPassExpHull(4 * FilterRate)
						MyFilterTR1 = New FilterLowPassExpHull(MyFilterBP1.Rate)
						MyFilterTR2 = New FilterLowPassExpHull(MyFilterBP2.Rate)
						MyFilterTR3 = New FilterLowPassExpHull(MyFilterBP3.Rate)
					Case FilterUltimateOscillatorType.PLL
						MyFilterBP1 = New FilterLowPassPLL(FilterRate)
						MyFilterBP2 = New FilterLowPassPLL(2 * FilterRate)
						MyFilterBP3 = New FilterLowPassPLL(4 * FilterRate)
						MyFilterTR1 = New FilterLowPassPLL(MyFilterBP1.Rate)
						MyFilterTR2 = New FilterLowPassPLL(MyFilterBP2.Rate)
						MyFilterTR3 = New FilterLowPassPLL(MyFilterBP3.Rate)
				End Select
				MyFilterPost = Nothing
			End Sub

			Public Sub New(ByVal FilterRate As Integer, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential)
				Me.New(CDbl(FilterRate), FilterType)
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal PostFilterRate As Integer, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential)
				Me.New(CDbl(FilterRate), FilterType)
				'post filter is always a Hull filter
				If PostFilterRate > 0 Then
					MyFilterPost = New FilterLowPassExpHull(PostFilterRate)
				End If
			End Sub

			Private Function Filter(ByVal Value As Double) As Double Implements IFilter.Filter
				'input type is not supported
				Throw New NotImplementedException
			End Function

			Public Function Filter(Value As IPriceVol) As Double Implements IFilter.Filter
				Dim ThisBuyPressure As Double
				Dim ThisBuyPressureRatio1 As Double
				Dim ThisBuyPressureRatio2 As Double
				Dim ThisBuyPressureRatio3 As Double

				If Value.LastPrevious < Value.Low Then
					ThisBuyPressure = Value.Last - Value.LastPrevious
				Else
					ThisBuyPressure = Value.Last - Value.Low
				End If
				MyFilterBP1.Filter(ThisBuyPressure)
				MyFilterBP2.Filter(ThisBuyPressure)
				MyFilterBP3.Filter(ThisBuyPressure)
				MyFilterTR1.Filter(Value.Range)
				MyFilterTR2.Filter(Value.Range)
				MyFilterTR3.Filter(Value.Range)

				If MyFilterTR1.FilterLast > 0 Then
					ThisBuyPressureRatio1 = MyFilterBP1.FilterLast / MyFilterTR1.FilterLast
					ThisBuyPressureRatio2 = MyFilterBP2.FilterLast / MyFilterTR2.FilterLast
					ThisBuyPressureRatio3 = MyFilterBP3.FilterLast / MyFilterTR3.FilterLast
				Else
					ThisBuyPressureRatio1 = 0.5
					ThisBuyPressureRatio2 = 0.5
					ThisBuyPressureRatio3 = 0.5
				End If
				FilterValueLast = (4 * ThisBuyPressureRatio1 + 2 * ThisBuyPressureRatio2 + ThisBuyPressureRatio3) / 7
				If MyFilterPost IsNot Nothing Then
					FilterValueLast = MyFilterPost.Filter(FilterValueLast)
				End If
				MyListOfValue.Add(FilterValueLast)
				ValueLast = Value
				Return FilterValueLast
			End Function

			Friend Function FilterDiffExp(Value As IPriceVol, ByVal FilterResult As Double) As Double
				FilterValueLast = FilterResult
				MyListOfValue.Add(FilterValueLast)
				ValueLast = Value
				Return FilterValueLast
			End Function

			Private Function Filter(ByRef Value() As Double) As Double() Implements IFilter.Filter
				'input type not supported
				Throw New NotImplementedException
			End Function

			''' <summary>
			''' Special filtering that can be used to remove the delay starting at a specific point
			''' </summary>
			''' <param name="Value">The value to be filtered</param>
			''' <param name="DelayRemovedToItem">The point where the delay stop to be removed</param>
			''' <returns>The result</returns>
			''' <remarks></remarks>
			Private Function Filter(ByRef Value() As Double, ByVal DelayRemovedToItem As Integer) As Double() Implements IFilter.Filter
				Throw New NotImplementedException
			End Function

			Private Function FilterErrorLast() As Double Implements IFilter.FilterErrorLast
				Return 0.0
			End Function

			Private Function FilterBackTo(ByRef Value As Double) As Double Implements IFilter.FilterBackTo
				Throw New NotImplementedException
			End Function

			Private Function FilterLastToPriceVol() As IPriceVol Implements IFilter.FilterLastToPriceVol
				Throw New NotImplementedException
			End Function

			Public Function LastToPriceVol() As IPriceVol Implements IFilter.LastToPriceVol
				Return ValueLast
			End Function

			Private Function Filter(ByVal Value As Single) As Double Implements IFilter.Filter
				Throw New NotImplementedException
			End Function

			Private Function FilterPredictionNext(ByVal Value As Double) As Double Implements IFilter.FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Private Function FilterPredictionNext(ByVal Value As Single) As Double Implements IFilter.FilterPredictionNext
				Throw New NotImplementedException
			End Function

			Public Function FilterLast() As Double Implements IFilter.FilterLast
				Return FilterValueLast
			End Function

			Public Function Last() As Double Implements IFilter.Last
				Return ValueLast.Last
			End Function

			Public ReadOnly Property Rate As Integer Implements IFilter.Rate
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer Implements IFilter.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements IFilter.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements IFilter.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IFilter.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property ToListOfError() As IList(Of Double) Implements IFilter.ToListOfError
				Get
					Throw New NotSupportedException
				End Get
			End Property

			Public ReadOnly Property ToListScaled() As ListScaled Implements IFilter.ToListScaled
				Get
					Return MyListOfValue
				End Get
			End Property

			Public Function ToArray() As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double() Implements IFilter.ToArray
				Return MyListOfValue.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Property Tag As String Implements IFilter.Tag

			Public Overrides Function ToString() As String Implements IFilter.ToString
				Return Me.FilterLast.ToString
			End Function

#Region "IRegisterKey"
			Public Function AsIRegisterKey() As IRegisterKey(Of String)
				Return Me
			End Function
			Private Property IRegisterKey_KeyID As Integer Implements IRegisterKey(Of String).KeyID
			Dim MyKeyValue As String
			Private Property IRegisterKey_KeyValue As String Implements IRegisterKey(Of String).KeyValue
				Get
					Return MyKeyValue
				End Get
				Set(value As String)
					MyKeyValue = value
				End Set
			End Property
#End Region
		End Class
#End Region
#Region "FilterUltimateOscillatorDeltaExp"
		<Serializable()>
		Public Class FilterUltimateOscillatorDeltaExp
			Inherits FilterUltimateOscillatorExp

			Private MyFilterSlow As FilterUltimateOscillatorExp
			Private MyFilterFast As FilterUltimateOscillatorExp
			Private MyFilterStatistic As FilterStatistical
			'Private ThisNormalDist = New MathNet.Numerics.Distributions.Normal(Mean, StandardDeviation)
			Private ThisNormalDist As MathNet.Numerics.Distributions.Normal



			Public Sub New(ByVal FilterRate As Double, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential, Optional ByVal FilterDifferenceRatio As Double = 2.0)
				MyBase.New(FilterRate, FilterType)
				MyFilterFast = New FilterUltimateOscillatorExp(FilterRate, FilterType)
				MyFilterSlow = New FilterUltimateOscillatorExp(FilterDifferenceRatio * FilterRate, FilterType)
				MyFilterStatistic = New FilterStatistical(MyFilterSlow.Rate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				ThisNormalDist = New MathNet.Numerics.Distributions.Normal(0, 1)
			End Sub

			Public Sub New(ByVal FilterRate As Integer, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential, Optional ByVal FilterDifferenceRatio As Double = 2.0)
				MyBase.New(FilterRate, FilterType)

				Dim ThisFilterRateSlow As Integer = CInt(FilterDifferenceRatio * FilterRate)
				MyFilterFast = New FilterUltimateOscillatorExp(FilterRate, FilterRate, FilterType)
				MyFilterSlow = New FilterUltimateOscillatorExp(ThisFilterRateSlow, ThisFilterRateSlow, FilterType)
				MyFilterStatistic = New FilterStatistical(MyFilterSlow.Rate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				ThisNormalDist = New MathNet.Numerics.Distributions.Normal(0, 1)
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal PostFilterRate As Integer, Optional ByVal FilterType As FilterUltimateOscillatorType = FilterUltimateOscillatorType.Exponential, Optional ByVal FilterDifferenceRatio As Double = 2.0)
				MyBase.New(FilterRate, FilterType)
				MyFilterFast = New FilterUltimateOscillatorExp(FilterRate, PostFilterRate, FilterType)
				MyFilterSlow = New FilterUltimateOscillatorExp(CInt(FilterDifferenceRatio * FilterRate), PostFilterRate, FilterType)
				MyFilterStatistic = New FilterStatistical(MyFilterSlow.Rate, StatisticType:=FilterVolatility.enuVolatilityStatisticType.Exponential)
				ThisNormalDist = New MathNet.Numerics.Distributions.Normal(0, 1)
			End Sub

			Public Overloads Function Filter(Value As IPriceVol) As Double
				Dim ThisFilterFast As Double = MyFilterFast.Filter(Value)
				Dim ThisFilterSlow As Double = MyFilterSlow.Filter(Value)
				Dim ThisFilterSum As Double = ThisFilterFast + ThisFilterSlow
				Dim ThisFilterDiff As Double = ThisFilterFast - ThisFilterSlow
				Dim ThisFilterLast As IStatistical = MyFilterStatistic.Filter(ThisFilterDiff)

				If ThisFilterLast.StandardDeviation > 0 Then
					Return MyBase.FilterDiffExp(Value, ThisNormalDist.CumulativeDistribution(ThisFilterDiff / ThisFilterLast.StandardDeviation))
				Else
					Return MyBase.FilterDiffExp(Value, 0.0)
				End If
			End Function
		End Class
#End Region
#Region "FilterStochasticBrownian"

		Public Class FilterStochasticBrownianUltimate
			Inherits FilterStochastic

			Private Const FILTER_RATE_BAND As Integer = 5

			Private MyFilterStochasticBrownian As FilterStochasticBrownian
			Private MyFilterStochasticBrownian2 As FilterStochasticBrownian
			Private MyFilterStochasticBrownian4 As FilterStochasticBrownian

			Private MyFilterVolatilityYangZhangForStatistic As FilterVolatilityYangZhang
			Private MyListOfProbabilityBandHigh As List(Of Double)
			Private MyListOfProbabilityBandLow As List(Of Double)
			Private MyListOfPriceVolatilityHigh As List(Of Double)
			Private MyListOfPriceVolatilityLow As List(Of Double)

			Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
				Me.New(FilterRate, CDbl(FilterOutputRate))
			End Sub

			Public Sub New(ByVal IsPreFilter As Boolean, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				Me.New(FilterRate, FilterRate, FilterOutputRate)
			End Sub

			Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
				Me.New(PreFilterRate, FilterRate, CDbl(FilterOutputRate))
			End Sub

			Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				MyBase.New(PreFilterRate, FilterRate, FilterOutputRate)

				MyFilterStochasticBrownian = New FilterStochasticBrownian(FilterRate, FilterOutputRate)
				MyFilterStochasticBrownian2 = New FilterStochasticBrownian(2 * FilterRate, FilterOutputRate)
				MyFilterStochasticBrownian4 = New FilterStochasticBrownian(4 * FilterRate, FilterOutputRate)

				MyListOfProbabilityBandHigh = New List(Of Double)
				MyListOfProbabilityBandLow = New List(Of Double)
				MyListOfPriceVolatilityHigh = New List(Of Double)
				MyListOfPriceVolatilityLow = New List(Of Double)
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				Me.New(FilterRate, FilterRate, FilterOutputRate)
			End Sub

			Public Overloads Function Filter(ByRef Value As IPriceVol) As Double
				Return Me.FilterLocal(Value)
			End Function

			Private Function FilterLocal(ByVal Value As IPriceVol) As Double
				Dim ThisValueRemoved As IPriceVolLarge = Nothing
				Dim ThisValueHigh As Double
				Dim ThisValueLow As Double
				Dim ThisStocRangeVolatility As Double
				Dim ThisProbHigh As Double
				Dim ThisProbLow As Double
				Dim ThisStochasticResult As Double
				Dim ThisGainPerYear As Double
				Dim ThisGainPerYearDerivative As Double
				Dim ThisPriceVolatilityHigh As Double
				Dim ThisPriceVolatilityLow As Double
				Dim ThisRate As Integer = Me.Rate
				Dim ThisPriceFilter As Double


				MyFilterStochasticBrownian.Filter(Value)
				MyFilterStochasticBrownian2.Filter(Value)
				MyFilterStochasticBrownian4.Filter(Value)
				ThisValueLow = Me.FilterUltimate(MyFilterStochasticBrownian.FilterPriceBandLow, MyFilterStochasticBrownian2.FilterPriceBandLow, MyFilterStochasticBrownian4.FilterPriceBandLow)
				ThisValueHigh = Me.FilterUltimate(MyFilterStochasticBrownian.FilterPriceBandHigh, MyFilterStochasticBrownian2.FilterPriceBandHigh, MyFilterStochasticBrownian4.FilterPriceBandHigh)

				ThisPriceFilter = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToFilterPrice.FilterLast,
					MyFilterStochasticBrownian2.ToFilterPrice.FilterLast,
					MyFilterStochasticBrownian4.ToFilterPrice.FilterLast)

				ThisGainPerYear = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToListOfGain.Last,
					MyFilterStochasticBrownian2.ToListOfGain.Last,
					MyFilterStochasticBrownian4.ToListOfGain.Last)

				ThisGainPerYearDerivative = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToListOfGainDerivative.Last,
					MyFilterStochasticBrownian2.ToListOfGainDerivative.Last,
					MyFilterStochasticBrownian4.ToListOfGainDerivative.Last)

				ThisStocRangeVolatility = Me.FilterUltimate(
					MyFilterStochasticBrownian.FilterLast(Type:=IStochastic.enuStochasticType.RangeVolatility),
					MyFilterStochasticBrownian2.FilterLast(Type:=IStochastic.enuStochasticType.RangeVolatility),
					MyFilterStochasticBrownian4.FilterLast(Type:=IStochastic.enuStochasticType.RangeVolatility))

				ThisProbHigh = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToList(Type:=IStochastic.enuStochasticType.ProbabilityHigh).Last,
					MyFilterStochasticBrownian2.ToList(Type:=IStochastic.enuStochasticType.ProbabilityHigh).Last,
					MyFilterStochasticBrownian4.ToList(Type:=IStochastic.enuStochasticType.ProbabilityHigh).Last)

				ThisProbLow = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToList(Type:=IStochastic.enuStochasticType.ProbabilityLow).Last,
					MyFilterStochasticBrownian2.ToList(Type:=IStochastic.enuStochasticType.ProbabilityLow).Last,
					MyFilterStochasticBrownian4.ToList(Type:=IStochastic.enuStochasticType.ProbabilityLow).Last)

				ThisPriceVolatilityHigh = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityHigh).Last,
					MyFilterStochasticBrownian2.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityHigh).Last,
					MyFilterStochasticBrownian4.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityHigh).Last)

				ThisPriceVolatilityLow = Me.FilterUltimate(
					MyFilterStochasticBrownian.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityLow).Last,
					MyFilterStochasticBrownian2.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityLow).Last,
					MyFilterStochasticBrownian4.ToList(Type:=IStochastic.enuStochasticType.PriceBandVolatilityLow).Last)

				ThisStochasticResult = ThisProbHigh / (ThisProbHigh + ThisProbLow)
				MyListOfProbabilityBandHigh.Add(ThisProbHigh)
				MyListOfProbabilityBandLow.Add(ThisProbLow)
				If MyListOfPriceVolatilityHigh.Count = 0 Then
					'Dim I As Integer
					''For I = 1 To FILTER_RATE_BAND - 1
					'For I = 1 To ThisRate - 1
					'  MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
					'  MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
					'Next
				End If
				MyListOfPriceVolatilityHigh.Add(ThisPriceVolatilityHigh)
				MyListOfPriceVolatilityLow.Add(ThisPriceVolatilityLow)
				Return Me.ListDataUpdate(ThisValueHigh, ThisValueLow, ThisStocRangeVolatility, ThisStochasticResult)
			End Function

			Public Overrides ReadOnly Property ToList(ByVal Type As IStochastic.enuStochasticType) As IList(Of Double)
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.ProbabilityHigh
							Return MyListOfProbabilityBandHigh
						Case IStochastic.enuStochasticType.ProbabilityLow
							Return MyListOfProbabilityBandLow
						Case IStochastic.enuStochasticType.PriceBandVolatilityHigh
							Return MyListOfPriceVolatilityHigh
						Case IStochastic.enuStochasticType.PriceBandVolatilityLow
							Return MyListOfPriceVolatilityLow
						Case Else
							Return MyBase.ToList(Type)
					End Select
				End Get
			End Property

			Private Function FilterUltimate(ByRef Value1 As Double, ByRef Value2 As Double, ByRef Value3 As Double) As Double
				Return (4 * Value1 + 2 * Value2 + Value3) / 7
			End Function
		End Class
#End Region
#Region "FilterStochasticSD1"
		''' <summary>
		''' NOT USE ANYMORE. DELETE WHEN READY SINCE THE IMPLEMENTATION IS NOW IN THE STANDARD FilterStochastic
		''' USING THE .IsfilterRange OPTION TO TRUE
		''' This is a modified version of the standard Stochactic taking in acount the 
		''' standard deviation of the signal in a special manner
		''' The standard definition is for N=14 and an output filter of 3 given by
		''' read more at: http://www.investopedia.com/terms/s/stochasticoscillator.asp#ixzz2MnmOFLnS
		''' %K = 100[(C - L14)/(H14 - L14)] 
		''' Where
		''' C = the most recent closing price
		''' L14 = the low of the 14 previous trading sessions
		''' H14 = the highest price traded during the same 14-day period.
		''' %D = 3-period moving average of %K
		''' </summary>
		''' <remarks>The current implementation support any rate or output filter</remarks>
		Public Class FilterStochasticSD1
			Implements IStochastic

			Private MyRate As Integer
			Private MyRateOutput As Integer
			Private MyRatePreFilter As Integer
			Private MyValueSum As Double
			Private MyValueSumSquare As Double
			Private MyStocFastLast As Double
			Private MyStocFastLastHisteresis As Double
			Private MyStocRangeVolatility As Double
			Private MyStocFastSlowLast As Double
			Private MyValueMax As FilterData
			Private MyValueMin As FilterData
			Private MyValueLast As Double
			Private MyRangeLast As Double
			Private MyFilter As FilterLowPassExp
			Private MyFilterPLL As FilterLowPassPLL
			Private MyListOfValueWindows As List(Of FilterData)
			Private MyListOfPVValueWindows As List(Of FilterData)
			Private MyFilterLPOfStochasticSlow As FilterLowPassExp
			Private MyPriceVolForRangeLast As IPriceVol
			Private MyFilterLPOfRange As FilterLowPassExp
			Private MyListOfStochasticFast As ListScaled
			Private MyListOfStochasticFastSlow As ListScaled
			Private MyFilterLPOfFilterBackTo As FilterLowPassExp
			Private MyListOfPriceBandHigh As List(Of Double)
			Private MyListOfPriceBandLow As List(Of Double)
			Private MyListOfPriceRangeVolatility As ListScaled
			Private MyFilterHighAttackDecay As FilterAttackDecayExp
			Private MyFilterLowAttackDecay As FilterAttackDecayExp
			'Private MyFilterBollinger As FilterBollingerBand

#Region "New"
			Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
				Me.New(FilterRate, CDbl(FilterOutputRate))
			End Sub
			Public Sub New(ByVal IsPreFilter As Boolean, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				Me.New(FilterRate, FilterRate, FilterOutputRate)
			End Sub
			Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Integer)
				Me.New(PreFilterRate, FilterRate, CDbl(FilterOutputRate))
			End Sub
			Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				Me.New(FilterRate, FilterOutputRate)
				MyRatePreFilter = PreFilterRate
				If MyRatePreFilter < 1 Then MyRatePreFilter = 1
				If MyRatePreFilter > 1 Then
					MyFilter = New FilterLowPassExp(MyRatePreFilter)
				End If
			End Sub
			Public Sub New(ByVal FilterRate As Integer, ByVal FilterOutputRate As Double)
				If FilterRate < 1 Then FilterRate = 1
				If FilterOutputRate < 1 Then FilterOutputRate = 1
				MyRatePreFilter = 1
				MyRate = FilterRate
				MyRateOutput = CInt(FilterOutputRate)
				MyListOfValueWindows = New List(Of FilterData)(capacity:=FilterRate)
				MyListOfStochasticFast = New ListScaled
				MyFilterLPOfStochasticSlow = New FilterLowPassExp(FilterOutputRate)
				MyFilterLPOfRange = New FilterLowPassExp(MyRate)
				MyFilterHighAttackDecay = New FilterAttackDecayExp(1, MyRate \ 2)
				MyFilterLowAttackDecay = New FilterAttackDecayExp(MyRate \ 2, 1)
				MyListOfStochasticFastSlow = New ListScaled
				MyListOfPriceBandHigh = New List(Of Double)
				MyListOfPriceBandLow = New List(Of Double)
				MyListOfPriceRangeVolatility = New ListScaled
				'MyFilterBollinger = New FilterBollingerBand(FilterRate)
			End Sub
#End Region

			Public Function Filter(ByVal Value As Single) As Double Implements IStochastic.Filter
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function Filter(ByRef Value As Double) As Double Implements IStochastic.Filter
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Value))

				'If Me.Tag = "AAPL:TargetPrice" Then
				'  Me.Tag = Me.Tag
				'End If
				MyFilterLPOfRange.Filter(Value)
				With ThisPriceVol
					If MyPriceVolForRangeLast IsNot Nothing Then
						.LastPrevious = MyPriceVolForRangeLast.Last
					End If
					If MyFilterLPOfRange.FilterLast > .Last Then
						.High = CSng(MyFilterLPOfRange.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf MyFilterLPOfRange.FilterLast < .Last Then
						.Low = CSng(MyFilterLPOfRange.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				MyPriceVolForRangeLast = ThisPriceVol
				Return Me.Filter(ThisPriceVol)
			End Function

			'Public Function Filter(ByRef Value As IPriceVol, ByVal ValueToAddAndAverage As Single) As Double Implements IStochastic.Filter
			'  Dim ThisPriceVol As IPriceVol = New PriceVol(0)

			'  With ThisPriceVol
			'    .Last = (Value.Last + ValueToAddAndAverage) / 2
			'    .Open = (Value.Open + ValueToAddAndAverage) / 2
			'    .High = .Last + (Value.High - Value.Last)
			'    .Low = .Last + (Value.Low - Value.Last)
			'    If MyPriceVolForRangeLast IsNot Nothing Then
			'      .LastPrevious = MyPriceVolForRangeLast.Last
			'    Else
			'      .LastPrevious = .Open
			'    End If
			'    .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
			'  End With
			'  MyPriceVolForRangeLast = ThisPriceVol
			'  Return Me.Filter(ThisPriceVol)
			'End Function

			Public Function Filter(ByRef Value As IPriceVol, FilterRate As Integer) As Double Implements IStochastic.Filter
				Return Me.Filter(Value)
			End Function

			Public Function Filter(ByRef Value As IPriceVol) As Double Implements IStochastic.Filter
				Dim ThisPriceDelta As Double
				Dim ThisValueRemoved As FilterData = Nothing
				Dim ThisData As FilterData
				Dim ThisVariance As Double
				Dim ThisMean As Double
				Dim ThisRange As Double
				Dim ThisMeanSquared As Double
				Dim ThisSigmaMean As Double

#If DebugPrediction Then
        Static IsHere As Boolean
        If IsHere = False Then
          IsHere = True
          Dim ThisResultPrediction As Double = Me.FilterPredictionNext(Value)
          Dim ThisResultActual = Me.Filter(Value)
          IsHere = False
          If ThisResultActual <> ThisResultPrediction Then
            Debugger.Break()
          End If
          Return ThisResultActual
        End If
#End If
				MyValueLast = Value.Last
				'ThisRange = Value.High - Value.Low
				ThisRange = Value.Range
				'ThisRange = 0
				If MyFilter Is Nothing Then
					ThisData = New FilterData(Value.Last, Value.Last, ThisRange)
				Else
					ThisData = New FilterData(MyFilter.Filter(Value.Last), Value.Last, ThisRange)
				End If
				If MyListOfValueWindows.Count = 0 Then
					'initialization
					MyValueMax = ThisData
					MyValueMin = ThisData
					MyValueSum = 0
					MyValueSumSquare = 0
					MyStocFastLast = 0.5
					MyStocFastSlowLast = 0.5
				Else
					With MyListOfValueWindows
						If .Count = MyRate Then
							ThisValueRemoved = .First
							MyValueSum = MyValueSum - ThisValueRemoved.Range
							MyValueSumSquare = MyValueSumSquare - (ThisValueRemoved.Range * ThisValueRemoved.Range)
							.RemoveAt(0)
						End If
						If .Count > 0 Then
							'if the element removed was a min or a max we need to find another one
							'the min and the max are not necessary located at the same index
							If ThisValueRemoved IsNot Nothing Then
								If ThisValueRemoved Is MyValueMax Then
									If ThisValueRemoved Is MyValueMin Then
										'need to search for a maximum and a minimum at the same time
										'should be a rare occurence
										MyValueMax = MyListOfValueWindows.First
										MyValueMin = MyValueMax
										For Each ThisDataLocal In MyListOfValueWindows
											If ThisDataLocal.FilterLast > MyValueMax.FilterLast Then
												MyValueMax = ThisDataLocal
											End If
											If ThisDataLocal.FilterLast < MyValueMin.FilterLast Then
												MyValueMin = ThisDataLocal
											End If
										Next
									Else
										'search only for a maximum
										MyValueMax = MyListOfValueWindows.First
										For Each ThisDataLocal In MyListOfValueWindows
											If ThisDataLocal.FilterLast > MyValueMax.FilterLast Then
												MyValueMax = ThisDataLocal
											End If
										Next
									End If
								Else
									If ThisValueRemoved Is MyValueMin Then
										'need to search for a minimum
										'search only for a maximum
										MyValueMin = .First
										For Each ThisDataLocal In MyListOfValueWindows
											If ThisDataLocal.FilterLast < MyValueMin.FilterLast Then
												MyValueMin = ThisDataLocal
											End If
										Next
									End If
								End If
							End If
						Else
							MyValueMax = ThisData
							MyValueMin = ThisData
						End If
						'update the max and min with the latest data
						If ThisData.FilterLast > MyValueMax.FilterLast Then
							MyValueMax = ThisData
						End If
						If ThisData.FilterLast < MyValueMin.FilterLast Then
							MyValueMin = ThisData
						End If
					End With
				End If
				MyListOfValueWindows.Add(ThisData)
				'Debug.Print(String.Format("{0},{1},{2},{3}", MyListOfPriceRangeVolatility.Count, MyValueMin.FilterLast, MyValueMax.FilterLast, MyValueMax.FilterLast - MyValueMin.FilterLast))
				MyValueSum = MyValueSum + ThisData.Range
				MyValueSumSquare = MyValueSumSquare + (ThisData.Range * ThisData.Range)
				ThisMean = MyValueSum / MyListOfValueWindows.Count
				ThisMeanSquared = ThisMean * ThisMean
				ThisSigmaMean = MyValueSumSquare / MyListOfValueWindows.Count
				If ThisSigmaMean > ThisMeanSquared Then
					ThisVariance = Math.Sqrt(ThisSigmaMean - ThisMeanSquared)
				Else
					ThisVariance = 0
				End If
				MyRangeLast = ThisMean + ThisVariance
				'If Me.Tag = "AAPLPrice" Then
				'  If MyListOfStochasticFast.Count = 370 Then
				'    ThisVariance = ThisVariance
				'  End If
				'End If

				'we now have the min and max over the period
				MyFilterHighAttackDecay.Filter(MyValueMax.FilterLast)
				MyFilterLowAttackDecay.Filter(MyValueMin.FilterLast)
				If Me.IsFilterPeak Then
					ThisPriceDelta = MyFilterHighAttackDecay.FilterLast - MyFilterLowAttackDecay.FilterLast + MyRangeLast
					If ThisPriceDelta <> 0 Then
						MyStocRangeVolatility = MyRangeLast / ThisPriceDelta
						MyStocFastLast = (ThisData.FilterLast - (MyFilterLowAttackDecay.FilterLast - MyRangeLast / 2)) / ThisPriceDelta
					End If
				Else
					ThisPriceDelta = MyValueMax.FilterLast - MyValueMin.FilterLast + MyRangeLast
					If ThisPriceDelta <> 0 Then
						MyStocRangeVolatility = MyRangeLast / ThisPriceDelta
						MyStocFastLast = (ThisData.FilterLast - (MyValueMin.FilterLast - MyRangeLast / 2)) / ThisPriceDelta
					End If
				End If

				MyListOfPriceRangeVolatility.Add(MyStocRangeVolatility)
				MyFilterLPOfStochasticSlow.Filter(MyStocFastLast)
				MyStocFastSlowLast = MyStocFastLast - MyFilterLPOfStochasticSlow.FilterLast
				MyListOfStochasticFast.Add(MyStocFastLast)
				MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
				MyListOfPriceBandHigh.Add(MyValueMax.FilterLast)
				MyListOfPriceBandLow.Add(MyValueMin.FilterLast)
				Return MyFilterLPOfStochasticSlow.FilterLast
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double Implements IStochastic.Filter
				Throw New NotSupportedException
			End Function


			Public Function Filter(ByRef Value() As Double) As Double() Implements IStochastic.Filter
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function FilterPredictionNext(ByRef Value As Double) As Double Implements IStochastic.FilterPredictionNext
				Dim I As Integer
				Dim ThisPriceDelta As Double
				Dim ThisValueRemoved As FilterData = Nothing
				Dim ThisData As FilterData
				Dim ThisDataLocal As FilterData

				Dim ThisValueMax As FilterData = MyValueMax
				Dim ThisValueMin As FilterData = MyValueMin
				Dim ThisStocFastLast As Double = 0
				Dim ThisStocFastSlowLast As Double = 0

				'If Me.Tag = "AAPL" Then
				'Debugger.Break()
				'If MyListOfStochasticFast.Count = 370 Then
				'Debugger.Break()
				'End If
				'End If
				If MyFilter Is Nothing Then
					ThisData = New FilterData(Value)
				Else
					If MyFilterPLL IsNot Nothing Then
						ThisData = New FilterData(MyFilterPLL.Filter(Value))
					Else
						ThisData = New FilterData(MyFilter.Filter(Value))
					End If
				End If
				If MyListOfValueWindows.Count = 0 Then
					'initialization
					ThisValueMax = ThisData
					ThisValueMin = ThisData
					ThisStocFastLast = 0
					ThisStocFastSlowLast = 0
				Else
					With MyListOfValueWindows
						If .Count = MyRate Then
							ThisValueRemoved = .First
						End If
						If .Count > 1 Then
							'if the element removed was a min or a max we need to find another one
							'the min and the max are not necessary located at the same index
							If ThisValueRemoved IsNot Nothing Then
								If ThisValueRemoved Is ThisValueMax Then
									If ThisValueRemoved Is ThisValueMin Then
										'need to search for a maximum and a minimum at the same time
										'should be a rare occurence
										ThisValueMax = MyListOfValueWindows(1)
										ThisValueMin = ThisValueMax
										For I = 1 To MyListOfValueWindows.Count - 1
											ThisDataLocal = MyListOfValueWindows(I)
											If ThisDataLocal.FilterLast > ThisValueMax.FilterLast Then
												ThisValueMax = ThisDataLocal
											End If
											If ThisDataLocal.FilterLast < ThisValueMin.FilterLast Then
												ThisValueMin = ThisDataLocal
											End If
										Next
									Else
										'search only for a maximum
										ThisValueMax = MyListOfValueWindows(1)
										For I = 1 To MyListOfValueWindows.Count - 1
											ThisDataLocal = MyListOfValueWindows(I)
											If ThisDataLocal.FilterLast > ThisValueMax.FilterLast Then
												ThisValueMax = ThisDataLocal
											End If
										Next
									End If
								Else
									If ThisValueRemoved Is ThisValueMin Then
										'need to search for a minimum
										'search only for a maximum
										ThisValueMin = MyListOfValueWindows(1)
										For I = 1 To MyListOfValueWindows.Count - 1
											ThisDataLocal = MyListOfValueWindows(I)
											If ThisDataLocal.FilterLast < ThisValueMin.FilterLast Then
												ThisValueMin = ThisDataLocal
											End If
										Next
									End If
								End If
							End If
						Else
							ThisValueMax = ThisData
							ThisValueMin = ThisData
						End If
						'update the max and min with the latest data
						If ThisData.FilterLast > ThisValueMax.FilterLast Then
							ThisValueMax = ThisData
						End If
						If ThisData.FilterLast < ThisValueMin.FilterLast Then
							ThisValueMin = ThisData
						End If
					End With
				End If
				'we now have the min and max over the period
				ThisPriceDelta = ThisValueMax.FilterLast - ThisValueMin.FilterLast
				If ThisPriceDelta <> 0 Then
					ThisStocFastLast = (ThisData.FilterLast - ThisValueMin.FilterLast) / ThisPriceDelta
				End If
				Return MyFilterLPOfStochasticSlow.FilterPredictionNext(ThisStocFastLast)
			End Function

			Public Function FilterBackTo(ByRef Value As Double, Optional ByVal IsPreFilter As Boolean = True) As Double Implements IStochastic.FilterBackTo
				Dim ThisStocFastLast As Double
				Dim ThisValue As Double

				ThisStocFastLast = MyFilterLPOfStochasticSlow.FilterBackTo(Value)
				If ThisStocFastLast > 1.0 Then
					ThisStocFastLast = 1
				ElseIf ThisStocFastLast < 0 Then
					ThisStocFastLast = 0
				End If
				If Me.IsFilterPeak Then
					ThisValue = ((MyFilterHighAttackDecay.FilterLast - MyFilterLowAttackDecay.FilterLast + MyRangeLast) * ThisStocFastLast) + (MyFilterLowAttackDecay.FilterLast - MyRangeLast / 2)
				Else
					ThisValue = ((MyValueMax.FilterLast - MyValueMin.FilterLast + MyRangeLast) * ThisStocFastLast) + (MyValueMin.FilterLast - MyRangeLast / 2)
				End If
				If MyFilter IsNot Nothing Then
					If IsPreFilter Then
						ThisValue = MyFilter.FilterBackTo(ThisValue)
					End If
				End If
				Return ThisValue
			End Function

			Public Function FilterPriceBandHigh() As Double Implements IStochastic.FilterPriceBandHigh
				Return MyValueMax.FilterLast
			End Function

			Public Function FilterPriceBandLow() As Double Implements IStochastic.FilterPriceBandLow
				Return MyValueMax.FilterLast
			End Function

			Public Function FilterLast() As Double Implements IStochastic.FilterLast
				Return MyFilterLPOfStochasticSlow.FilterLast
			End Function

			''' <summary>
			''' return the last result of the fast-slow stochastic calculation
			''' </summary>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Function FilterLast(ByVal Type As IStochastic.enuStochasticType) As Double Implements IStochastic.FilterLast
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyStocFastSlowLast
					Case IStochastic.enuStochasticType.Fast
						Return MyStocFastLast
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.FilterLast
					Case IStochastic.enuStochasticType.RangeVolatility
						Return MyListOfPriceRangeVolatility.Last
					Case IStochastic.enuStochasticType.PriceBandHigh
						Return MyValueMax.FilterLast
					Case IStochastic.enuStochasticType.PriceBandLow
						Return MyValueMin.FilterLast
					Case Else
						Return MyStocFastSlowLast
				End Select
			End Function

			Public Function Last() As Double Implements IStochastic.Last
				Return MyValueLast
			End Function

			Public Property Rate(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Integer Implements IStochastic.Rate
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyRate
						Case IStochastic.enuStochasticType.Fast
							Return MyRatePreFilter
						Case IStochastic.enuStochasticType.Slow
							Return MyRateOutput
						Case Else
							Return 1
					End Select
				End Get
				Set(value As Integer)
					'do not set the rate here
				End Set
			End Property

			Public ReadOnly Property Count As Integer Implements IStochastic.Count
				Get
					Return MyFilterLPOfStochasticSlow.Count
				End Get
			End Property

			Public ReadOnly Property Max(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Max
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow.Max
						Case IStochastic.enuStochasticType.Fast
							Return 1.0
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.Max
						Case Else
							Return MyListOfStochasticFastSlow.Max
					End Select
				End Get
			End Property

			Public ReadOnly Property Min(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Min
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow.Min
						Case IStochastic.enuStochasticType.Fast
							Return 0.0
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.Min
						Case Else
							Return MyListOfStochasticFastSlow.Min
					End Select
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IStochastic.ToList
				Get
					Return MyFilterLPOfStochasticSlow.ToList
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal Type As IStochastic.enuStochasticType) As IList(Of Double) Implements IStochastic.ToList
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow
						Case IStochastic.enuStochasticType.Fast
							Return MyListOfStochasticFast
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.ToList
						Case IStochastic.enuStochasticType.PriceBandHigh
							Return MyListOfPriceBandHigh
						Case IStochastic.enuStochasticType.PriceBandLow
							Return MyListOfPriceBandLow
						Case IStochastic.enuStochasticType.RangeVolatility
							Return MyListOfPriceRangeVolatility
						Case Else
							Return MyListOfStochasticFastSlow
					End Select
				End Get
			End Property

			Public Function ToArray(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.ToArray
					Case IStochastic.enuStochasticType.Fast
						Return MyListOfStochasticFast.ToArray
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.ToArray
					Case Else
						Return MyListOfStochasticFastSlow.ToArray
				End Select
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Return Me.ToArray(Me.Min(Type), Me.Max(Type), ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case IStochastic.enuStochasticType.Fast
						Return MyListOfStochasticFast.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case Else
						Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
				End Select
			End Function

			Public Property Tag As String Implements IStochastic.Tag

			Public Overrides Function ToString() As String Implements IStochastic.ToString
				Return Me.FilterLast.ToString
			End Function

			Public Property IsFilterPeak As Boolean Implements IStochastic.IsFilterPeak

			Public Property IsFilterRange As Boolean Implements IStochastic.IsFilterRange
		End Class
#End Region
#Region "ListWindow"
		Friend Class ListWindow(Of T As {Class, IPriceVolLarge})
			Implements IList(Of T)

			Private MyListOfIPriceVol As List(Of T)
			Private MyWindowSize As Integer
			Private MyItemHigh As T
			Private MyItemLow As T
			Private MyItemRemoved As T
			Private MyItemHighLast As T
			Private MyItemLowLast As T

#Region "New"
			Public Sub New(ByVal WindowSize As Integer)
				MyListOfIPriceVol = New List(Of T)(WindowSize)
				MyItemHigh = Nothing
				MyItemLow = Nothing
				MyItemRemoved = Nothing
				MyWindowSize = WindowSize
			End Sub
#End Region
#Region "Main Properties"
			ReadOnly Property ItemLow As T
				Get
					Return MyItemLow
				End Get
			End Property

			ReadOnly Property ItemHigh As T
				Get
					Return MyItemHigh
				End Get
			End Property

			ReadOnly Property WindowSize As Integer
				Get
					Return MyWindowSize
				End Get
			End Property

			ReadOnly Property ItemRemoved As T
				Get
					Return MyItemRemoved
				End Get
			End Property

			Public Sub RemoveLast()
				With MyListOfIPriceVol
					.RemoveAt(.Count - 1)       'remove on top
					.Insert(0, MyItemRemoved)   'insert at the beginning
					're-store the max and min
					MyItemHigh = MyItemHighLast
					MyItemLow = MyItemLowLast
				End With
			End Sub
#End Region
#Region "ICollection"
			Public Sub Add(item As T) Implements ICollection(Of T).Add
				With MyListOfIPriceVol

					'If TypeOf item Is ValueType Then
					'  MyItemRemoved = .First
					'End If

					MyItemHighLast = MyItemHigh
					MyItemLowLast = MyItemLow
					If .Count = MyWindowSize Then
						MyItemRemoved = .First
						.RemoveAt(0)
					Else
						MyItemRemoved = Nothing
					End If

					If .Count > 0 Then
						'if the element removed was a min or a max we need to find another one
						'the min and the max are not necessary located at the same index
						If MyItemRemoved IsNot Nothing Then
							If MyItemRemoved Is MyItemHigh Then
								If MyItemRemoved Is MyItemLow Then
									'need to search for a maximum and a minimum at the same time
									'should be a rare occurence
									MyItemHigh = .First
									MyItemLow = MyItemHigh
									For Each ThisPriceVol As T In MyListOfIPriceVol
										If ThisPriceVol.FilterLast >= MyItemHigh.FilterLast Then
											MyItemHigh = ThisPriceVol
										End If
										If ThisPriceVol.FilterLast <= MyItemLow.FilterLast Then
											MyItemLow = ThisPriceVol
										End If
									Next
								Else
									'search only for a maximum
									MyItemHigh = .First
									For Each ThisPriceVol In MyListOfIPriceVol
										If ThisPriceVol.FilterLast >= MyItemHigh.FilterLast Then
											MyItemHigh = ThisPriceVol
										End If
									Next
								End If
							Else
								If MyItemRemoved Is MyItemLow Then
									'need to search for a minimum
									'search only for a minimum
									MyItemLow = .First
									For Each ThisPriceVol In MyListOfIPriceVol
										If ThisPriceVol.FilterLast <= MyItemLow.FilterLast Then
											MyItemLow = ThisPriceVol
										End If
									Next
								End If
							End If
						End If
					Else
						MyItemHigh = item
						MyItemLow = item
					End If
					'update the max and min with the latest data
					If item.FilterLast >= MyItemHigh.FilterLast Then
						MyItemHigh = item
					End If
					If item.FilterLast <= MyItemLow.FilterLast Then
						MyItemLow = item
					End If
					.Add(item)
				End With
			End Sub

			Public Sub Clear() Implements ICollection(Of T).Clear
				MyListOfIPriceVol.Clear()
				MyItemHigh = Nothing
				MyItemHighLast = Nothing
				MyItemLow = Nothing
				MyItemLowLast = Nothing
				MyItemRemoved = Nothing
			End Sub

			Public Function Contains(item As T) As Boolean Implements ICollection(Of T).Contains
				Return MyListOfIPriceVol.Contains(item)
			End Function

			Public Sub CopyTo(array() As T, arrayIndex As Integer) Implements ICollection(Of T).CopyTo
				MyListOfIPriceVol.CopyTo(array, arrayIndex)
			End Sub

			Public ReadOnly Property Count As Integer Implements ICollection(Of T).Count
				Get
					Return MyListOfIPriceVol.Count
				End Get
			End Property

			Public ReadOnly Property IsReadOnly As Boolean Implements ICollection(Of T).IsReadOnly
				Get
					Return False
				End Get
			End Property

			Public Function Remove(item As T) As Boolean Implements ICollection(Of T).Remove
				Throw New NotImplementedException
				'Return MyListOfIPriceVol.Remove(item)
			End Function
#End Region
#Region "IEnumerable"
			Public Function GetEnumerator() As IEnumerator(Of T) Implements IEnumerable(Of T).GetEnumerator
				Return MyListOfIPriceVol.GetEnumerator
			End Function

			''' <summary>
			''' non generic implementation does not need to be public
			''' </summary>
			''' <returns></returns>
			''' <remarks></remarks>
			Private Function IList_GetEnumerator() As System.Collections.IEnumerator Implements System.Collections.IEnumerable.GetEnumerator
				Return Me.GetEnumerator()
			End Function
#End Region
#Region "IList"
			Public Function IndexOf(item As T) As Integer Implements IList(Of T).IndexOf
				Return MyListOfIPriceVol.IndexOf(item)
			End Function

			Public Sub Insert(index As Integer, item As T) Implements IList(Of T).Insert
				Throw New NotImplementedException
				'MyListOfIPriceVol.Insert(index, item)
			End Sub

			Default Public Property Item(index As Integer) As T Implements IList(Of T).Item
				Get
					Return MyListOfIPriceVol.Item(index)
				End Get
				Set(value As T)
					Throw New NotImplementedException
					'MyListOfIPriceVol.Item(index) = value
				End Set
			End Property

			Public Sub RemoveAt(index As Integer) Implements IList(Of T).RemoveAt
				Throw New NotImplementedException
				'MyListOfIPriceVol.RemoveAt(index)
			End Sub
#End Region
		End Class
#End Region
#Region "FilterData"
		Friend Class FilterData
			Implements IFilterData

			Public Sub New()
				Me.FilterLast = 0.0
			End Sub

			Public Sub New(ByVal FilterLast As Double)
				Me.FilterLast = FilterLast
				Me.FilterInput = FilterLast
				Me.Range = 0
			End Sub

			Public Sub New(ByVal FilterLast As Double, ByVal FilterInput As Double)
				Me.FilterLast = FilterLast
				Me.FilterInput = FilterInput
				Me.Range = 0
			End Sub

			Public Sub New(ByVal FilterLast As Double, ByVal FilterInput As Double, ByVal Range As Double)
				Me.FilterInput = FilterInput
				Me.FilterLast = FilterLast
				Me.Range = Range
			End Sub

			Public Property FilterInput As Double Implements IFilterData.FilterInput
			Public Property FilterLast As Double Implements IFilterData.FilterLast
			Public Property Range As Double Implements IFilterData.Range

			'Private Property FilterInput As Double
			'Public Property FilterLast As Double
			'Public Property Range As Double

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterStochasticOptimized"
		Public Class FilterStochasticOptimized
			Implements IStochastic

			Private MyFilterRate As Integer
			Private MyFilterLPOfRate As FilterLowPassExp
			Private MyRate As Integer
			Private MyRateMin As Integer
			Private MyRateMax As Integer
			Private MyRateOutput As Integer
			Private MyRatePreFilter As Integer
			Private MyValueSum As Double
			Private MyValueSumSquare As Double
			Private MyStocFastLast As Double
			Private MyStocFastLastHisteresis As Double
			Private MyStocRangeVolatility As Double
			Private MyStocFastSlowLast As Double
			Private MyValueLast As Double
			'Private MyValueMax As FilterData
			'Private MyValueMin As FilterData
			Private MyFilter As FilterLowPassExp
			Private MyFilterPLL As FilterLowPassPLL
			Private MyListOfValueWindows As List(Of FilterData)
			Private MyListOfPVValueWindows As List(Of FilterData)
			Private MyFilterLPOfStochasticSlow As FilterLowPassExp
			Private MyPriceVolForRangeLast As IPriceVol
			Private MyFilterLPOfRange As FilterLowPassExp
			Private MyListOfStochasticFast As ListScaled
			Private MyListOfStochasticFastSlow As ListScaled
			Private MyFilterLPOfFilterBackTo As FilterLowPassExp
			Private MyListOfPriceBandHigh As List(Of Double)
			Private MyListOfPriceBandLow As List(Of Double)
			Private MyListOfPriceRangeVolatility As ListScaled
			Private IsPreFilterEnabled As Boolean
			Private MyListOfStochastic As List(Of IStochastic)
			Private MyDictionaryOfStochastic As Dictionary(Of Integer, IStochastic)
			Private MyStochastic As IStochastic
			Private IsFilterPeakLocal As Boolean
			Private IsFilterRangeLocal As Boolean

			Public Sub New(ByRef ListOfStochastic As IList(Of IStochastic))
				If ListOfStochastic.Count = 0 Then
					Throw New ArgumentOutOfRangeException
				End If

				MyRateMin = ListOfStochastic.First.Rate
				MyRateMax = ListOfStochastic.Last.Rate
				IsFilterPeakLocal = ListOfStochastic.First.IsFilterPeak
				IsFilterRangeLocal = ListOfStochastic.First.IsFilterPeak

				MyRate = MyRateMin - 1
				MyListOfStochastic = New List(Of IStochastic)
				MyDictionaryOfStochastic = New Dictionary(Of Integer, IStochastic)
				For Each Stochastic In ListOfStochastic
					If Stochastic.Rate - MyRate <> 1 Then
						Throw New ArgumentOutOfRangeException
					End If
					MyRate = Stochastic.Rate
					Stochastic.IsFilterPeak = Me.IsFilterPeak
					Stochastic.IsFilterRange = Me.IsFilterRange
					MyListOfStochastic.Add(Stochastic)
					MyDictionaryOfStochastic.Add(MyRate, Stochastic)
				Next
				MyRate = MyRateMin
				MyStochastic = MyListOfStochastic.First
				MyFilterLPOfStochasticSlow = New FilterLowPassExp(2)
				MyListOfStochasticFast = New ListScaled
				MyListOfStochasticFastSlow = New ListScaled
				MyListOfPriceBandHigh = New List(Of Double)
				MyListOfPriceBandLow = New List(Of Double)
				MyListOfPriceRangeVolatility = New ListScaled
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByRef ListOfStochastic As IList(Of IStochastic))
				Me.New(ListOfStochastic)
				MyFilterRate = FilterRate
				MyFilterLPOfRate = New FilterLowPassExp(MyFilterRate)
			End Sub

			Public Function Filter(ByVal Value As Single) As Double Implements IStochastic.Filter
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function Filter(ByRef Value As Double) As Double Implements IStochastic.Filter
				Dim ThisPriceVol As IPriceVol = New PriceVol(CSng(Value))

				'If Me.Tag = "AAPL:TargetPrice" Then
				'  Me.Tag = Me.Tag
				'End If
				MyFilterLPOfRange.Filter(Value)
				With ThisPriceVol
					If MyPriceVolForRangeLast IsNot Nothing Then
						.LastPrevious = MyPriceVolForRangeLast.Last
					End If
					If MyFilterLPOfRange.FilterLast > .Last Then
						.High = CSng(MyFilterLPOfRange.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					ElseIf MyFilterLPOfRange.FilterLast < .Last Then
						.Low = CSng(MyFilterLPOfRange.FilterLast)
						.Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
					End If
				End With
				MyPriceVolForRangeLast = ThisPriceVol
				Return Me.Filter(ThisPriceVol)
			End Function

			'Public Function Filter(ByRef Value As IPriceVol, ByVal ValueToAddAndAverage As Single) As Double Implements IStochastic.Filter
			'  Dim ThisPriceVol As IPriceVol = New PriceVol(0)

			'  With ThisPriceVol
			'    .Last = (Value.Last + ValueToAddAndAverage) / 2
			'    .Open = (Value.Open + ValueToAddAndAverage) / 2
			'    .High = .Last + (Value.High - Value.Last)
			'    .Low = .Last + (Value.Low - Value.Last)
			'    If MyPriceVolForRangeLast IsNot Nothing Then
			'      .LastPrevious = MyPriceVolForRangeLast.Last
			'    Else
			'      .LastPrevious = .Open
			'    End If
			'    .Range = RecordPrices.CalculateTrueRange(ThisPriceVol)
			'  End With
			'  MyPriceVolForRangeLast = ThisPriceVol
			'  Return Me.Filter(ThisPriceVol)
			'End Function

			Public Function Filter(ByRef Value As IPriceVol, FilterRate As Integer) As Double Implements IStochastic.Filter
				If MyFilterLPOfRate IsNot Nothing Then
					FilterRate = CInt(MyFilterLPOfRate.Filter(CDbl(FilterRate)))
				End If
				If MyRate <> FilterRate Then
					MyRate = FilterRate
					If MyDictionaryOfStochastic.ContainsKey(MyRate) = False Then
						Throw New ArgumentOutOfRangeException
					End If
					MyStochastic = MyDictionaryOfStochastic(MyRate)
				End If
				Return Me.Filter(Value)
			End Function

			Public Function Filter(ByRef Value As IPriceVol) As Double Implements IStochastic.Filter
				For Each Stochastic As IStochastic In MyListOfStochastic
					Stochastic.Filter(Value)
				Next
				MyStocRangeVolatility = MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.RangeVolatility)
				MyStocFastLast = MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.Fast)
				MyListOfPriceRangeVolatility.Add(MyStocRangeVolatility)
				MyFilterLPOfStochasticSlow.Filter(MyStocFastLast)
				MyStocFastSlowLast = MyStocFastLast - MyFilterLPOfStochasticSlow.FilterLast
				MyListOfStochasticFast.Add(MyStocFastLast)
				MyListOfStochasticFastSlow.Add(MyStocFastSlowLast)
				MyListOfPriceBandHigh.Add(MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.PriceBandHigh))
				MyListOfPriceBandLow.Add(MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.PriceBandLow))
				Return MyFilterLPOfStochasticSlow.FilterLast
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ValueExpectedMin As Double, ValueExpectedMax As Double) As Double Implements IStochastic.Filter
				Throw New NotSupportedException
			End Function

			Public Function Filter(ByRef Value() As Double) As Double() Implements IStochastic.Filter
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return Me.ToArray
			End Function

			Public Function FilterPredictionNext(ByRef Value As Double) As Double Implements IStochastic.FilterPredictionNext
				Return MyStochastic.FilterPredictionNext(Value)
			End Function

			Public Function FilterBackTo(ByRef Value As Double, Optional ByVal IsPreFilter As Boolean = True) As Double Implements IStochastic.FilterBackTo
				Return MyStochastic.FilterBackTo(Value, IsPreFilter)
			End Function

			Public Function FilterPriceBandHigh() As Double Implements IStochastic.FilterPriceBandHigh
				Return MyStochastic.FilterPriceBandHigh
			End Function

			Public Function FilterPriceBandLow() As Double Implements IStochastic.FilterPriceBandLow
				Return MyStochastic.FilterPriceBandLow
			End Function

			Public Function FilterLast() As Double Implements IStochastic.FilterLast
				Return MyFilterLPOfStochasticSlow.FilterLast
			End Function

			''' <summary>
			''' return the last result of the fast-slow stochastic calculation
			''' </summary>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Function FilterLast(ByVal Type As IStochastic.enuStochasticType) As Double Implements IStochastic.FilterLast
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyStocFastSlowLast
					Case IStochastic.enuStochasticType.Fast
						Return MyStocFastLast
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.FilterLast
					Case IStochastic.enuStochasticType.RangeVolatility
						Return MyListOfPriceRangeVolatility.Last
					Case IStochastic.enuStochasticType.PriceBandHigh
						Return MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.PriceBandHigh)
					Case IStochastic.enuStochasticType.PriceBandLow
						Return MyStochastic.FilterLast(Type:=IStochastic.enuStochasticType.PriceBandLow)
					Case Else
						Return MyStocFastSlowLast
				End Select
			End Function

			Public Function Last() As Double Implements IStochastic.Last
				Return MyValueLast
			End Function

			Public Property Rate(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Integer Implements IStochastic.Rate
				Get
					Return MyStochastic.Rate(Type)
				End Get
				Set(value As Integer)
					If MyRate <> value Then
						MyRate = value
						If MyDictionaryOfStochastic.ContainsKey(MyRate) = False Then
							Throw New ArgumentOutOfRangeException
						End If
						MyStochastic = MyDictionaryOfStochastic(MyRate)
					End If
				End Set
			End Property

			Public ReadOnly Property Count As Integer Implements IStochastic.Count
				Get
					Return MyFilterLPOfStochasticSlow.Count
				End Get
			End Property

			Public ReadOnly Property Max(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Max
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow.Max
						Case IStochastic.enuStochasticType.Fast
							Return 1.0
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.Max
						Case Else
							Return MyListOfStochasticFastSlow.Max
					End Select
				End Get
			End Property

			Public ReadOnly Property Min(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double Implements IStochastic.Min
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow.Min
						Case IStochastic.enuStochasticType.Fast
							Return 0.0
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.Min
						Case Else
							Return MyListOfStochasticFastSlow.Min
					End Select
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements IStochastic.ToList
				Get
					Return MyFilterLPOfStochasticSlow.ToList
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal Type As IStochastic.enuStochasticType) As IList(Of Double) Implements IStochastic.ToList
				Get
					Select Case Type
						Case IStochastic.enuStochasticType.FastSlow
							Return MyListOfStochasticFastSlow
						Case IStochastic.enuStochasticType.Fast
							Return MyListOfStochasticFast
						Case IStochastic.enuStochasticType.Slow
							Return MyFilterLPOfStochasticSlow.ToList
						Case IStochastic.enuStochasticType.PriceBandHigh
							Return MyListOfPriceBandHigh
						Case IStochastic.enuStochasticType.PriceBandLow
							Return MyListOfPriceBandLow
						Case IStochastic.enuStochasticType.RangeVolatility
							Return MyListOfPriceRangeVolatility
						Case Else
							Return MyListOfStochasticFastSlow
					End Select
				End Get
			End Property

			Public Function ToArray(Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.ToArray
					Case IStochastic.enuStochasticType.Fast
						Return MyListOfStochasticFast.ToArray
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.ToArray
					Case Else
						Return MyListOfStochasticFastSlow.ToArray
				End Select
			End Function

			Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Return Me.ToArray(Me.Min(Type), Me.Max(Type), ScaleToMinValue, ScaleToMaxValue)
			End Function

			Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double, Optional ByVal Type As IStochastic.enuStochasticType = IStochastic.enuStochasticType.FastSlow) As Double() Implements IStochastic.ToArray
				Select Case Type
					Case IStochastic.enuStochasticType.FastSlow
						Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case IStochastic.enuStochasticType.Fast
						Return MyListOfStochasticFast.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case IStochastic.enuStochasticType.Slow
						Return MyFilterLPOfStochasticSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
					Case Else
						Return MyListOfStochasticFastSlow.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)
				End Select
			End Function

			Public Property Tag As String Implements IStochastic.Tag

			Public Overrides Function ToString() As String Implements IStochastic.ToString
				Return Me.FilterLast.ToString
			End Function

			Public Property IsFilterPeak As Boolean Implements IStochastic.IsFilterPeak
				Get
					Return IsFilterPeakLocal
				End Get
				Set(value As Boolean)
					IsFilterPeakLocal = value
					For Each Stochastic In MyListOfStochastic
						Stochastic.IsFilterPeak = IsFilterPeakLocal
					Next
				End Set
			End Property
			Public Property IsFilterRange As Boolean Implements IStochastic.IsFilterRange
				Get
					Return IsFilterRangeLocal
				End Get
				Set(value As Boolean)
					IsFilterRangeLocal = value
					For Each Stochastic In MyListOfStochastic
						Stochastic.IsFilterRange = IsFilterRangeLocal
					Next
				End Set
			End Property
		End Class
#End Region
#Region "FilterBandHighLow"
		<Serializable()>
		Public Class FilterBandHighLow(Of T As YahooAccessData.IPriceVol)
			Private MyRate As Integer
			Private MyPriceVolMax As IPriceVol
			Private MyPriceVolMin As IPriceVol
			Private MyPriceVolLast As IPriceVol
			Private MyPriceVolResultLast As IPriceVol
			Private MyListOfPriceVolWindows As List(Of IPriceVol)
			Private MyListOfPriceVolHighLow As List(Of IPriceVol)

			Public Sub New(ByVal FilterRate As Integer)
				If FilterRate < 1 Then FilterRate = 1
				MyRate = FilterRate
				MyListOfPriceVolWindows = New List(Of IPriceVol)(capacity:=FilterRate)
				MyListOfPriceVolHighLow = New List(Of IPriceVol)(capacity:=FilterRate)
			End Sub

			Public Function Filter(ByVal PriceVol As T) As IPriceVol
				Dim ThisPriceVolRemoved As IPriceVol = Nothing
				Dim ThisPriceVolFiltered As IPriceVol
				Dim ThisPriceVol As IPriceVol

				MyPriceVolLast = PriceVol
				ThisPriceVolFiltered = PriceVol
				If MyListOfPriceVolWindows.Count = 0 Then
					'initialization
					MyPriceVolMax = ThisPriceVolFiltered
					MyPriceVolMin = ThisPriceVolFiltered
					MyPriceVolResultLast = New PriceVol(0)
					With MyPriceVolResultLast
						.High = MyPriceVolMax.High
						.Open = MyPriceVolMax.Open
						.Low = MyPriceVolMax.Low
						.Last = MyPriceVolMax.Last
						.LastWeighted = MyPriceVolMax.LastWeighted
					End With
					MyListOfPriceVolHighLow.Add(MyPriceVolResultLast)
				Else
					MyPriceVolResultLast = New PriceVol(0)
					With MyPriceVolResultLast
						.High = MyPriceVolMax.High
						.Open = MyListOfPriceVolWindows.First.Open
						.Low = MyPriceVolMin.Low
						.Last = ThisPriceVolFiltered.Last
						.LastWeighted = RecordPrices.CalculateLastWeighted(MyPriceVolResultLast)
					End With
					MyListOfPriceVolHighLow.Add(MyPriceVolResultLast)
					With MyListOfPriceVolWindows
						If MyListOfPriceVolWindows.Count = MyRate Then
							ThisPriceVolRemoved = .First
							.RemoveAt(0)
						End If
						'if the element removed was a min or a max we need to find another one
						'the min and the max are not necessary located at the same index
						If ThisPriceVolRemoved IsNot Nothing Then
							If ThisPriceVolRemoved Is MyPriceVolMax Then
								If ThisPriceVolRemoved Is MyPriceVolMin Then
									'need to search for a maximum and a minimum at the same time
									'should be a rare occurence
									MyPriceVolMax = MyListOfPriceVolWindows.First
									MyPriceVolMin = MyPriceVolMax
									For Each ThisPriceVol In MyListOfPriceVolWindows
										If ThisPriceVol.High > MyPriceVolMax.High Then
											MyPriceVolMax = ThisPriceVol
										End If
										If ThisPriceVol.Low < MyPriceVolMin.Low Then
											MyPriceVolMin = ThisPriceVol
										End If
									Next
								Else
									'search only for a maximum
									MyPriceVolMax = MyListOfPriceVolWindows.First
									For Each ThisPriceVol In MyListOfPriceVolWindows
										If ThisPriceVol.High > MyPriceVolMax.High Then
											MyPriceVolMax = ThisPriceVol
										End If
									Next
								End If
							Else
								If ThisPriceVolRemoved Is MyPriceVolMin Then
									'need to search for a minimum
									'search only for a maximum
									MyPriceVolMin = MyListOfPriceVolWindows.First
									For Each ThisPriceVol In MyListOfPriceVolWindows
										If ThisPriceVol.Low < MyPriceVolMin.Low Then
											MyPriceVolMin = ThisPriceVol
										End If
									Next
								End If
							End If
						End If
						'update the max and min with the latest data
						If ThisPriceVolFiltered.High > MyPriceVolMax.High Then
							MyPriceVolMax = ThisPriceVolFiltered
						End If
						If ThisPriceVolFiltered.Low < MyPriceVolMin.Low Then
							MyPriceVolMin = ThisPriceVolFiltered
						End If
					End With
				End If
				MyListOfPriceVolWindows.Add(ThisPriceVolFiltered)
				'we now have the min and max over the period
				'If Me.Tag = "AAPL" Then
				'  If MyListOfStochasticFast.Count = 370 Then
				'    'Debugger.Break()
				'  End If
				'End If
				Return MyPriceVolResultLast
			End Function

			Public Function Filter(ByRef Value() As T) As YahooAccessData.IPriceVol()
				Dim ThisValue As T
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return MyListOfPriceVolWindows.ToArray
			End Function

			Public Function FilterLast() As IPriceVol
				Return MyPriceVolResultLast
			End Function

			Public Function Last() As IPriceVol
				Return MyPriceVolLast
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfPriceVolHighLow.Count
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of IPriceVol)
				Get
					Return MyListOfPriceVolHighLow.ToList
				End Get
			End Property

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterBollingerBand"
		<Serializable()>
		Public Class FilterBollingerBand
			Public Enum enuBollingerBandType
				Band
				BandShiftedUp
				BandPriceRange
				BandPriceRangeShiftedUp
			End Enum
			Public Enum enuBollingerDataType
				Volatility
				VolatilityYearlyCorrected
				BandPerCent
				PriceToMomentumRatio
			End Enum

			Private MyRate As Integer
			Private MyKFactor As Double
			Private MyBollingerLast As IPriceVol
			Private MyBollingerPriceRangeLast As IPriceVol
			Private MyBollingerLastPrevious As IPriceVol
			Private MyBollingerPriceRangeLastPrevious As IPriceVol
			Private MyValueLast As Double
			Private MyFilterX0 As FilterLowPassExp
			Private MyFilterX As FilterLowPassExp
			Private MyFilterX2 As FilterLowPassExp
			Private MyFilterBand As FilterLowPassExp
			'Private MyFilterSigmaHighLow As FilterLowPassExp
			Private MyFilterXHighLow As FilterLowPassExp
			Private MyFilterX2HighLow As FilterLowPassExp
			Private MyFilterXHull As FilterLowPassExpHull
			Private MyCountOfPriceMomentumRatio As Integer

			Private MyListOfBollingerBand As List(Of IPriceVol)
			Private MyListOfBollingerBandPriceRange As List(Of IPriceVol)
			Private MyListOfBollingerBandPriceRangeLastPrevious As List(Of IPriceVol)
			Private MyListOfBollingerBandLastPrevious As List(Of IPriceVol)

			Private MyListOfPriceMomentumRatio As ListScaled
			Private MyListOfPriceVolatility As ListScaled
			Private MyListOfPriceVolatilityYearlyCorrected As ListScaled
			Private MyListOfBandPercent As ListScaled
			Private MyListOfPriceToMomentumRatio As ListScaled
			Private MyVolatilityCorrectionFactor As Double

			Public Sub New(ByVal FilterRate As Integer)
				If FilterRate < 1 Then FilterRate = 1
				MyRate = FilterRate
				MyVolatilityCorrectionFactor = Math.Sqrt(NUMBER_WORKDAY_PER_YEAR / (2 * FilterRate))
				MyKFactor = 2.0
				MyFilterX0 = New FilterLowPassExp(FilterRate)
				MyFilterX = New FilterLowPassExp(FilterRate)
				MyFilterX2 = New FilterLowPassExp(FilterRate)
				MyFilterXHull = New FilterLowPassExpHull(FilterRate)
				MyFilterXHighLow = New FilterLowPassExp(FilterRate)
				MyFilterX2HighLow = New FilterLowPassExp(FilterRate)
				'MyFilterSigmaHighLow = New FilterLowPassExp(FilterRate)
				MyListOfBollingerBand = New List(Of IPriceVol)(capacity:=FilterRate)
				MyListOfBollingerBandPriceRange = New List(Of IPriceVol)(capacity:=FilterRate)
				MyListOfBollingerBandPriceRangeLastPrevious = New List(Of IPriceVol)(capacity:=FilterRate)
				MyListOfBollingerBandLastPrevious = New List(Of IPriceVol)(capacity:=FilterRate)
				MyListOfPriceMomentumRatio = New ListScaled(capacity:=FilterRate)
				MyListOfPriceVolatility = New ListScaled(capacity:=FilterRate)
				MyListOfPriceVolatilityYearlyCorrected = New ListScaled(capacity:=FilterRate)
				MyListOfBandPercent = New ListScaled(capacity:=FilterRate)
				MyListOfPriceToMomentumRatio = New ListScaled(capacity:=FilterRate)
				MyFilterBand = New FilterLowPassExp(FilterRate)
			End Sub

			Public Sub New(ByVal FilterRate As Integer, ByVal KFactor As Double)
				Me.New(FilterRate)
				MyKFactor = KFactor
			End Sub

			Public Function Filter(ByVal Value As IPriceVol) As IPriceVol
				Dim ThisPriceVol As IPriceVol = New PriceVol(0)
				Dim ThisSigmaPriceRange As Double
				Dim ThisSigmaMomentum As Double
				Dim ThisSigmaTotal As Double
				Dim ThisHighLow As Single

				MyValueLast = Value.Last

				MyFilterXHull.Filter(Value.Last)
				MyFilterBand.Filter(Value.Last)
				MyFilterX.Filter(MyFilterXHull.FilterLast)
				MyFilterX2.Filter(MyFilterXHull.FilterLast * MyFilterXHull.FilterLast)
				'ThisHighLow = Value.High - Value.Low
				'If Me.Tag = "AAPL" Then
				'  If Value.Range <> 0 Then
				'    Me.Tag = Me.Tag
				'  End If
				'End If
				ThisHighLow = Value.Range
				MyFilterXHighLow.Filter(ThisHighLow)
				MyFilterX2HighLow.Filter(ThisHighLow * ThisHighLow)

				ThisSigmaMomentum = Math.Sqrt(MyFilterX2.FilterLast - (MyFilterX.FilterLast * MyFilterX.FilterLast))
				ThisSigmaPriceRange = (Math.Sqrt(MyFilterX2HighLow.FilterLast - (MyFilterXHighLow.FilterLast * MyFilterXHighLow.FilterLast)))
				'ThisSigmaPriceRange = 0
				'ThisSigmaTotal = ThisSigmaMomentum + ThisSigmaPriceRange
				ThisSigmaTotal = ThisSigmaMomentum + ThisSigmaPriceRange
				If ThisSigmaPriceRange = 0 Then
					MyListOfPriceMomentumRatio.Add(0)
				Else
					'start to record the price momentun only when it is stabilized
					MyCountOfPriceMomentumRatio = MyCountOfPriceMomentumRatio + 1
					If MyCountOfPriceMomentumRatio > Me.Rate Then
						MyListOfPriceMomentumRatio.Add(ThisSigmaPriceRange / ThisSigmaTotal)
						'MyListOfPriceMomentumRatio.Add(Math.Log10(ThisSigmaTotal / ThisSigmaPriceRange))
					Else
						MyListOfPriceMomentumRatio.Add(0)
					End If
				End If

				'calculate the price range first
				With ThisPriceVol
					.Last = CSng(MyFilterXHull.FilterLast)
					.Open = .Last
					.High = CSng(.Last + (MyKFactor * ThisSigmaPriceRange))
					.Low = CSng(.Last - (MyKFactor * ThisSigmaPriceRange))
					If MyBollingerPriceRangeLast IsNot Nothing Then
						.LastPrevious = MyBollingerPriceRangeLast.Last
						MyBollingerPriceRangeLast.OpenNext = .Open
					Else
						.LastPrevious = .Last
					End If
					.OpenNext = .Last
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol.AsIPriceVol)
				End With
				If MyListOfBollingerBandPriceRangeLastPrevious.Count = 0 Then
					MyBollingerPriceRangeLast = ThisPriceVol
				End If
				MyBollingerPriceRangeLastPrevious = MyBollingerPriceRangeLast
				MyBollingerPriceRangeLast = ThisPriceVol
				MyListOfBollingerBandPriceRangeLastPrevious.Add(MyBollingerPriceRangeLastPrevious)
				MyListOfBollingerBandPriceRange.Add(MyBollingerPriceRangeLast)

				'proceed with the main calculation momentum + price range
				With ThisPriceVol
					'.High = CSng(.Last + (MyKFactor * ThisSigmaTotal))
					'.Low = CSng(.Last - (MyKFactor * ThisSigmaTotal))
					.High = CSng(.Last + (MyKFactor * ThisSigmaMomentum))
					.Low = CSng(.Last - (MyKFactor * ThisSigmaMomentum))
					If MyBollingerLast IsNot Nothing Then
						.LastPrevious = MyBollingerLast.Last
						MyBollingerLast.OpenNext = .Open
					End If
					.Range = RecordPrices.CalculateTrueRange(ThisPriceVol.AsIPriceVol)
				End With
				If MyListOfBollingerBandLastPrevious.Count = 0 Then
					MyBollingerLast = ThisPriceVol
				End If
				MyBollingerLastPrevious = MyBollingerLast
				MyBollingerLast = ThisPriceVol
				MyListOfBollingerBandLastPrevious.Add(MyBollingerLastPrevious)
				MyListOfBollingerBand.Add(MyBollingerLast)
				If MyBollingerLast.Last <> 0 Then
					MyListOfPriceVolatility.Add(MyBollingerLast.Range / MyBollingerLast.Last)
				Else
					MyListOfPriceVolatility.Add(0)
				End If
				If MyBollingerLast.Range <> 0 Then
					MyListOfBandPercent.Add((((MyBollingerLast.Last - MyFilterBand.FilterLast) / (MyBollingerLast.Range / 2)) / 2) + 0.5)
				Else
					MyListOfBandPercent.Add(0.5)
				End If
				MyListOfPriceVolatilityYearlyCorrected.Add(MyVolatilityCorrectionFactor * MyListOfPriceVolatility.Last)
				Return MyBollingerLast
			End Function

			Public Function Filter(ByVal Value As Double) As IPriceVol
				Return Me.Filter(New PriceVol(CSng(Value)))
			End Function

			Public Function Filter(ByRef Value() As Double) As IPriceVol()
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return MyListOfBollingerBand.ToArray
			End Function

			Public Function FilterLast() As IPriceVol
				Return MyBollingerLast
			End Function

			Public Function FilterLastPrevious() As IPriceVol
				Return MyBollingerLast
			End Function

			Public Function Last() As Double
				Return MyValueLast
			End Function

			Public ReadOnly Property Rate As Integer
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property KFactor As Double
				Get
					Return MyKFactor
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfBollingerBand.Count
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of IPriceVol)
				Get
					Return MyListOfBollingerBand
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal ListBollingerBandType As enuBollingerBandType) As IList(Of IPriceVol)
				Get
					Select Case ListBollingerBandType
						Case enuBollingerBandType.Band
							Return MyListOfBollingerBand
						Case enuBollingerBandType.BandShiftedUp
							Return MyListOfBollingerBandLastPrevious
						Case enuBollingerBandType.BandPriceRange
							Return MyListOfBollingerBandPriceRange
						Case enuBollingerBandType.BandPriceRangeShiftedUp
							Return MyListOfBollingerBandPriceRangeLastPrevious
						Case Else
							Return MyListOfBollingerBand
					End Select
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal ListBollingerDataType As enuBollingerDataType) As IList(Of Double)
				Get
					Select Case ListBollingerDataType
						Case enuBollingerDataType.Volatility
							Return MyListOfPriceVolatility
						Case enuBollingerDataType.VolatilityYearlyCorrected
							Return MyListOfPriceVolatilityYearlyCorrected
						Case enuBollingerDataType.BandPerCent
							Return MyListOfBandPercent
						Case enuBollingerDataType.PriceToMomentumRatio
							Return MyListOfPriceMomentumRatio
						Case Else
							Return MyListOfPriceVolatility
					End Select
				End Get
			End Property

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterRSI"
		''' <summary>
		''' https://en.wikipedia.org/wiki/Relative_strength_index
		''' </summary>
		''' <remarks></remarks>
		<Serializable()>
		Public Class FilterRSI
			Public Enum enuListDataType
				RSI
				ADX
				DIPlus
				DIMinus
			End Enum

			Public Enum SlopeDirection
				NotSpecified
				Positive
				Zero
				Negative
			End Enum

			Private MyRate As Double
			Private MyRatePreFilter As Double
			Private A As Double
			Private B As Double
			Private MyTickRSIUp As Double
			Private MyTickRSIDown As Double
			Private MyDIUp As Double
			Private MyDIDown As Double
			Private MyDIRangeUp As Double
			Private MyDIRangeDown As Double
			Private MyADXFiltered As Double
			Private MyTickSum As Double
			Private MyFilterRangeLast As Double
			Private MyTickPVUpDown As IPriceVolLarge

			Private MyValueLast As Double
			Private MyValuePriceVolFilteredLast As IPriceVol
			Private MyValueFilteredLast As Double
			Private MyValueMax As Double
			Private MyValueMin As Double
			Private RSILast As Double
			Private MyFilterExp As FilterExp
			Private MyFilterPV As FilterLowPassExp(Of PriceVol)
			Private MyPostFilterHighPassExp As FilterHighPassExp
			Private MyListOfRSI As List(Of Double)
			Private MyListOfADX As List(Of Double)
			Private MyListOfDIPlus As List(Of Double)
			Private MyListOfDIMinus As List(Of Double)

			Public Sub New(ByVal FilterRate As Double)
				MyListOfRSI = New List(Of Double)
				MyListOfADX = New List(Of Double)
				MyListOfDIPlus = New List(Of Double)
				MyListOfDIMinus = New List(Of Double)
				If FilterRate < 1 Then FilterRate = 1
				MyRate = FilterRate
				MyRatePreFilter = 1
				A = CDbl((2 / (FilterRate + 1)))
				B = 1 - A
				MyTickRSIUp = 0
				MyTickRSIDown = 0
				MyValueLast = 0
				MyValueMax = 1.0
				MyValueMin = 0.0

				MyPostFilterHighPassExp = Nothing
			End Sub

			Public Sub New(ByVal PreFilterRate As Integer, ByVal FilterRate As Integer)
				Me.New(PreFilterRate:=PreFilterRate, FilterRate:=FilterRate, PostFilterHighPassRate:=0)
			End Sub

			Public Sub New(ByVal PreFilterRate As Double, ByVal FilterRate As Double, PostFilterHighPassRate As Double)
				Me.New(FilterRate)
				If PreFilterRate < 1 Then PreFilterRate = 1
				MyRatePreFilter = PreFilterRate
				If MyRatePreFilter > 1 Then
					MyFilterExp = New FilterExp(MyRatePreFilter)
					MyFilterPV = New FilterLowPassExp(Of PriceVol)(MyRatePreFilter)
				End If
				If PostFilterHighPassRate > 2 Then
					MyPostFilterHighPassExp = New FilterHighPassExp(PostFilterHighPassRate)
				Else
					PostFilterHighPassRate = Nothing
				End If
			End Sub


			Public Function Filter(ByVal Value As Double, Optional ByVal Direction As SlopeDirection = SlopeDirection.NotSpecified) As Double
				Dim ValueFiltered As Double
				Dim InputDelta As Double

				MyValueLast = Value
				If MyFilterPV Is Nothing Then
					ValueFiltered = Value
				Else
					ValueFiltered = MyFilterExp.FilterRun(Value)
				End If
				If MyListOfRSI.Count = 0 Then
					MyValueFilteredLast = ValueFiltered
					RSILast = 0.5
					MyTickRSIUp = 0
					MyTickRSIDown = 0
				End If

				InputDelta = ValueFiltered - MyValueFilteredLast
				MyValueFilteredLast = ValueFiltered
				Select Case Direction
					Case SlopeDirection.NotSpecified
						Select Case InputDelta
							Case > 0
								MyTickRSIUp = B * MyTickRSIUp + A * InputDelta
								MyTickRSIDown = B * MyTickRSIDown
							Case < 0
								MyTickRSIDown = B * MyTickRSIDown - A * InputDelta
								MyTickRSIUp = B * MyTickRSIUp
						End Select
					Case SlopeDirection.Positive
						MyTickRSIUp = B * MyTickRSIUp + A * Math.Abs(InputDelta)
						MyTickRSIDown = B * MyTickRSIDown
					Case SlopeDirection.Negative
						MyTickRSIDown = B * MyTickRSIDown + A * Math.Abs(InputDelta)
						MyTickRSIUp = B * MyTickRSIUp
					Case SlopeDirection.Zero
						'do nothing
				End Select
				MyTickSum = MyTickRSIUp + MyTickRSIDown
				If MyTickSum = 0 Then
					RSILast = 0.5
				Else
					RSILast = MyTickRSIUp / MyTickSum
				End If
				If MyPostFilterHighPassExp IsNot Nothing Then
					RSILast = MyPostFilterHighPassExp.Filter(RSILast) + 0.5
				End If
				MyListOfRSI.Add(RSILast)
				Return RSILast
			End Function

			''' <summary>
			''' In addition to calculate the RSI, calculate the ADX Trend detector and DI+ and DI-
			''' as defined by J. Welles Wilder in is book New Concepts in Technical Trading Systems. Greensboro, NC: Trend Research, 1978. 
			''' </summary>
			''' <param name="Value"></param>
			''' <returns></returns>
			''' <remarks></remarks>
			Public Function Filter(ByVal Value As IPriceVol) As Double
				Dim ThisValuePVFiltered As IPriceVol
				Dim ThisDelta As Double
				Dim ThisDeltaHigh As Double
				Dim ThisDeltaLow As Double
				Dim ThisDIUp As Double
				Dim ThisDIDown As Double
				Dim ThisADX As Double
				Dim ThisDISum As Double


				If MyFilterPV Is Nothing Then
					'no filtering
					ThisValuePVFiltered = DirectCast(Value, PriceVol).CopyFrom
				Else
					ThisValuePVFiltered = MyFilterPV.Filter(Value)
				End If
				If MyListOfRSI.Count = 0 Then
					MyValuePriceVolFilteredLast = DirectCast(ThisValuePVFiltered, PriceVol).CopyFrom
					RSILast = 0.5
					MyADXFiltered = 0
					MyTickRSIUp = 0
					MyTickRSIDown = 0
					MyDIUp = 0
					MyDIDown = 0
					MyFilterRangeLast = ThisValuePVFiltered.Range
				End If

				'Process the Last and calculate the standard RSI
				ThisDelta = ThisValuePVFiltered.Last - MyValuePriceVolFilteredLast.Last
				If ThisDelta > 0 Then
					'If Me.Tag = "Price" Then
					'  Me.Tag = "Price"
					'End If
					MyTickRSIUp = B * MyTickRSIUp + A * ThisDelta
					MyTickRSIDown = B * MyTickRSIDown
				ElseIf ThisDelta < 0 Then
					'If Me.Tag = "Price" Then
					'  Me.Tag = "Price"
					'End If
					MyTickRSIDown = B * MyTickRSIDown - A * ThisDelta
					MyTickRSIUp = B * MyTickRSIUp
				Else
					MyTickRSIDown = B * MyTickRSIDown
					MyTickRSIUp = B * MyTickRSIUp
				End If
				MyTickSum = MyTickRSIUp + MyTickRSIDown
				If MyTickSum = 0 Then
					RSILast = 0.5
				Else
					'this is the standard definition of RSI
					'ThisADX = MyTickRSIUp / MyTickSum
					'however we prefer the following definition which is slightly less subject to whipsaw
					RSILast = MyTickRSIUp / Math.Sqrt(MyTickRSIUp ^ 2 + MyTickRSIDown ^ 2)
				End If
				'process the High amd low DI
				ThisDeltaHigh = ThisValuePVFiltered.High - MyValuePriceVolFilteredLast.High
				ThisDeltaLow = -(ThisValuePVFiltered.Low - MyValuePriceVolFilteredLast.Low)
				If ThisValuePVFiltered.Range > 0 Then
					If MyFilterRangeLast = 0 Then
						MyFilterRangeLast = ThisValuePVFiltered.Range
					End If
					If (ThisDeltaHigh > ThisDeltaLow) Then
						ThisDIDown = 0
						ThisDIUp = ThisDeltaHigh
					Else
						ThisDIUp = 0
						ThisDIDown = ThisDeltaLow
					End If
					MyDIDown = B * MyDIDown + A * ThisDIDown
					MyDIUp = B * MyDIUp + A * ThisDIUp
					'Process the true range
					MyFilterRangeLast = B * MyFilterRangeLast + A * ThisValuePVFiltered.Range
					MyDIRangeDown = MyDIDown / MyFilterRangeLast
					MyDIRangeUp = MyDIUp / MyFilterRangeLast
				End If
				ThisDISum = MyDIRangeUp + MyDIRangeDown
				If ThisDISum > 0 Then
					'this is the real defition of ADX
					'ThisADX = Math.Abs((MyDIRangeUp - MyDIRangeDown) / (ThisDISum))
					'however we prefer this relation 
					'ThisADX = ((((MyDIRangeUp - MyDIRangeDown)) / (MyDIRangeUp + MyDIRangeDown)) + 1) / 2
					'which is almost equivalent ot the RSI above except the up and down tick are calculated on the high and low
					'rather than the last measurement
					'even better this may be a better calculation more standard with the concept of energy and bollinger band
					'this is the one we use here
					'ThisADX = ((((MyDIRangeUp - MyDIRangeDown)) / Math.Sqrt(MyDIRangeUp ^ 2 + MyDIRangeDown ^ 2)) + 1) / 2
					ThisADX = Math.Sqrt(MyDIRangeUp ^ 2 + MyDIRangeDown ^ 2)
					MyADXFiltered = ThisADX
				Else
					'MyADXFiltered = B * MyADXFiltered
				End If
				'Debug.Print(String.Format("DI-:{0}, DI+:{1}, Range:{2}", MyDIDown, MyDIUp, MyFilterRangeLast))

				MyListOfRSI.Add(RSILast)
				'MyListOfRSI.Add(MyADXFiltered)
				MyListOfADX.Add(MyADXFiltered)
				MyListOfDIPlus.Add(MyDIRangeUp)
				MyListOfDIMinus.Add(MyDIRangeDown)
				MyValuePriceVolFilteredLast = ThisValuePVFiltered
				Return RSILast
			End Function

			Public Function Filter(ByRef Value() As Double) As Double()
				Dim ThisValue As Double
				For Each ThisValue In Value
					Me.Filter(ThisValue)
				Next
				Return MyListOfRSI.ToArray
			End Function

			Public Function Filter(ByVal Value As Single) As Double
				Return Me.Filter(CDbl(Value))
			End Function

			Public Function FilterLast() As Double
				Return RSILast
			End Function

			Public Function Last() As Double
				Return MyValueLast
			End Function

			Public ReadOnly Property RateInternal As Double
				Get
					Return MyRate
				End Get
			End Property

			Public ReadOnly Property Rate As Integer
				Get
					Return CInt(MyRate)
				End Get
			End Property

			Public ReadOnly Property Count As Integer
				Get
					Return MyListOfRSI.Count
				End Get
			End Property

			Public ReadOnly Property Max As Double
				Get
					Return 1.0
				End Get
			End Property

			Public ReadOnly Property Min As Double
				Get
					Return 0.0
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double)
				Get
					Return MyListOfRSI
				End Get
			End Property

			Public ReadOnly Property ToList(ByVal ListDataType As enuListDataType) As IList(Of Double)
				Get
					Select Case ListDataType
						Case enuListDataType.ADX
							Return MyListOfADX
						Case enuListDataType.DIMinus
							Return MyListOfDIMinus
						Case enuListDataType.DIPlus
							Return MyListOfDIPlus
						Case Else
							Return MyListOfRSI
					End Select
				End Get
			End Property

			'Public Function ToArray() As Double()
			'	Return MyListOfRSI.ToArray
			'End Function


			'Public Function ToArray(ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
			'	Return MyListOfRSI.ToArray(ScaleToMinValue, ScaleToMaxValue)
			'End Function

			'Public Function ToArray(ByVal MinValueInitial As Double, ByVal MaxValueInitial As Double, ByVal ScaleToMinValue As Double, ByVal ScaleToMaxValue As Double) As Double()
			'	Return MyListOfRSI.ToArray(MinValueInitial, MaxValueInitial, ScaleToMinValue, ScaleToMaxValue)

			'End Function

			Public Property Tag As String

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function
		End Class
#End Region
#Region "FilterArrayExp"
		Public Class FilterArrayExp
			''' <summary>
			''' Implements an exponential in place high pass array filtering
			''' </summary>
			''' <param name="InputToOutput">
			''' The single dimension input vector data. The result is retuned in place in the same array
			''' </param>
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering
			''' </param>
			''' <remarks>
			''' The algorithm is based on the low pass filtering algorithm result as Yhp(k)=X(k)-Ylp(k)
			''' </remarks>
			Public Shared Sub HighPass(ByRef InputToOutput() As Single, ByVal FilterRate As Integer)
				Dim FilterLowPass(0 To UBound(InputToOutput)) As Single
				Dim I As Integer

				If FilterRate = 0 Then Return
				Call FilterArrayExp.LowPass(InputToOutput, FilterLowPass, FilterRate)
				For I = 0 To UBound(InputToOutput)
					InputToOutput(I) = InputToOutput(I) - FilterLowPass(I)
				Next
			End Sub

			Public Shared Sub HighPass(ByRef InputToOutput() As Double, ByVal FilterRate As Integer)
				Dim FilterLowPass(0 To UBound(InputToOutput)) As Double
				Dim I As Integer

				If FilterRate = 0 Then Return
				Call FilterArrayExp.LowPass(InputToOutput, FilterLowPass, FilterRate)
				For I = 0 To UBound(InputToOutput)
					InputToOutput(I) = InputToOutput(I) - FilterLowPass(I)
				Next
			End Sub

			''' <summary>
			''' Implements an exponential high pass array filtering
			''' </summary>
			''' <param name="Input">
			''' The single dimension input vector data.
			''' </param>
			''' <param name="Output">
			''' The single dimension output vector data.
			''' </param>
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering
			''' </param>
			''' <remarks>
			''' The algorithm is based on the low pass filtering algorithm result as Yhp(k)=X(k)-Ylp(k)
			''' </remarks>
			Public Shared Sub HighPass(ByRef Input() As Single, ByRef Output() As Single, ByVal FilterRate As Integer)
				Dim I As Integer

				If FilterRate = 0 Then Return
				Call FilterArrayExp.LowPass(Input, Output, FilterRate)
				For I = 0 To UBound(Input)
					Output(I) = Input(I) - Output(I)
				Next I
			End Sub

			''' <summary>
			''' Implements an exponential high pass array filtering
			''' </summary>
			''' <param name="Input">
			''' The single dimension input vector data.
			''' </param>
			''' <param name="Output">
			''' The single dimension output vector data.
			''' </param>
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering
			''' </param>
			''' <remarks>
			''' The algorithm is based on the low pass filtering algorithm result as Yhp(k)=X(k)-Ylp(k)
			''' </remarks>
			Public Shared Sub HighPass(ByRef Input() As Double, ByRef Output() As Double, ByVal FilterRate As Integer)
				Dim I As Integer

				If FilterRate = 0 Then Return
				Call FilterArrayExp.LowPass(Input, Output, FilterRate)
				For I = 0 To UBound(Input)
					Output(I) = Input(I) - Output(I)
				Next I
			End Sub

			''' <summary>
			''' Implements an exponential in place low pass array filtering
			''' </summary>
			''' <param name="InputToOutput"></param>
			''' The single dimension input vector data. The result is retuned in place in the same array
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering 
			''' </param>
			''' <remarks>
			''' The algorith implements InputToOutput(k)=A*InputToOutput(k)+(1-A)*InputToOutput(k-1)
			''' Where A=2/(FilterRate+1)
			''' 
			''' The algorithm start the filtering at the first element differents than zero
			''' </remarks>
			Public Shared Sub LowPass(ByRef InputToOutput() As Single, ByVal FilterRate As Integer)
				Dim A As Single
				Dim B As Single
				Dim I As Integer
				Dim InputLast As Single

				If FilterRate = 0 Then Return

				'set the initial condition
				A = CSng((2 / (FilterRate + 1)))
				B = 1 - A

				For I = 0 To UBound(InputToOutput)
					If InputToOutput(I) <> 0 Then
						InputLast = InputToOutput(I)
						Exit For
					End If
				Next
				If I > UBound(InputToOutput) Then Return
				For I = I + 1 To UBound(InputToOutput)
					InputLast = A * InputToOutput(I) + B * InputLast
					InputToOutput(I) = InputLast
				Next
			End Sub

			''' <summary>
			''' Implements an exponential in place low pass array filtering
			''' </summary>
			''' <param name="InputToOutput"></param>
			''' The single dimension input vector data. The result is retuned in place in the same array
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering 
			''' </param>
			''' <remarks>
			''' The algorith implements InputToOutput(k)=A*InputToOutput(k)+(1-A)*InputToOutput(k-1)
			''' Where A=2/(FilterRate+1)
			''' 
			''' The algorithm start the filtering at the first element differents than zero
			''' </remarks>
			Public Shared Sub LowPass(ByRef InputToOutput() As Double, ByVal FilterRate As Integer)
				Dim A As Double
				Dim B As Double
				Dim I As Integer
				Dim InputLast As Double

				If FilterRate = 0 Then Return

				A = (2 / (FilterRate + 1))
				B = 1 - A
				'set the initial condition
				I = 0
				For I = 0 To UBound(InputToOutput)
					If InputToOutput(I) <> 0 Then
						InputLast = InputToOutput(I)
						Exit For
					End If
				Next
				If I > UBound(InputToOutput) Then Return
				For I = I + 1 To UBound(InputToOutput)
					InputLast = A * InputToOutput(I) + B * InputLast
					InputToOutput(I) = InputLast
				Next
			End Sub


			''' <summary>
			''' Implements an exponential array filtering
			''' </summary>
			''' <param name="Input">
			''' The single dimension input vector data
			''' </param>
			''' <param name="Output"></param>
			''' The filtered vector data
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering 
			''' </param>
			''' <remarks>
			''' The algorith implements Output(k)=A*Input(k)+(1-A)*Output(k-1)
			''' Where A=2/(FilterRate+1)
			''' </remarks>
			Public Shared Sub LowPass(ByRef Input() As Single, Output() As Single, ByVal FilterRate As Integer)
				Dim A As Single
				Dim B As Single
				Dim I As Integer

				If FilterRate = 0 Then Exit Sub

				A = CSng(2 / (FilterRate + 1))
				B = 1 - A
				'set the initial condition
				I = 0
				For I = 0 To UBound(Input)
					If Input(I) <> 0 Then
						Output(I) = Input(I)
						Exit For
					Else
						Output(I) = 0
					End If
				Next
				If I > UBound(Input) Then Return
				For I = I + 1 To UBound(Input)
					Output(I) = A * Input(I) + B * Output(I - 1)
				Next
			End Sub

			''' <summary>
			''' Implements an exponential array filtering
			''' </summary>
			''' <param name="Input">
			''' The single dimension input vector data
			''' </param>
			''' <param name="Output"></param>
			''' The filtered vector data
			''' <param name="FilterRate">
			''' Filter rate in number of point equivalent RMS filtering 
			''' </param>
			''' <remarks>
			''' The algorith implements Output(k)=A*Input(k)+(1-A)*Output(k-1)
			''' Where A=2/(FilterRate+1)
			''' </remarks>
			Public Shared Sub LowPass(ByRef Input() As Double, Output() As Double, ByVal FilterRate As Integer)
				Dim A As Double
				Dim B As Double
				Dim I As Integer

				If FilterRate = 0 Then Exit Sub

				A = 2 / (FilterRate + 1)
				B = 1 - A
				'set the initial condition
				I = 0
				For I = 0 To UBound(Input)
					If Input(I) <> 0 Then
						Output(I) = Input(I)
						Exit For
					Else
						Output(I) = 0
					End If
				Next
				If I > UBound(Input) Then Return
				For I = I + 1 To UBound(Input)
					Output(I) = A * Input(I) + B * Output(I - 1)
				Next
			End Sub
		End Class
#End Region
#Region "TransactionStockBuy"
		Public Class TransactionStockBuy
			Implements ITransaction

			Private MyPriceStart As Double
			Private MyPriceLast As IPriceVol
			Private MyFilterLast As Double
			Private MyTransactionCost As Double
			Private MyTransactionCount As Integer
			Private MyPriceTransactionStart As Double
			Private MyPriceTransactionStop As Double
			Private MyListOfValue As ListScaled
			Private MyPriceStop As Double
			Private MyPriceStopCount As Integer
			Private MyPriceVolStop As IPriceVol
			Private IsStopped As Boolean

			Public Sub New()
				Me.New(0)
			End Sub

			Public Sub New(ByVal TransactionCost As Double)
				MyTransactionCost = TransactionCost
				MyListOfValue = New ListScaled
				Me.IsStopReverse = True
				Me.IsStopEnabled = True  'by default
			End Sub
			''' <summary>
			''' Calculate the gain of a transaction in Percent of the original buy value
			''' </summary>
			''' <param name="Value">The price</param>
			''' <returns>The price gain in Percent</returns>
			''' <remarks></remarks>
			Public Function Filter(ByRef Value As IPriceVol) As Double Implements ITransaction.Filter
				If MyListOfValue.Count = 0 Then
					MyPriceStart = Value.OpenNext + MyTransactionCost
					MyPriceStop = Value.Last
					MyPriceStopCount = 0
				End If
				If IsStopped Then
					MyPriceStopCount = MyPriceStopCount + 1
				Else
					MyTransactionCount = MyTransactionCount + 1
					MyPriceLast = Value
					MyPriceTransactionStop = Value.Last
					MyPriceStop = Value.Last
					MyFilterLast = (MyPriceLast.Last - MyPriceStart) / MyPriceStart
				End If
				MyListOfValue.Add(MyFilterLast)
				Return MyFilterLast
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As IPriceVol) As Double Implements ITransaction.Filter
				If (ValueStop Is Nothing) Or (Me.IsStopEnabled = False) Then
					Return Me.Filter(Value)
				Else
					Return Me.Filter(Value, ValueStop.Low)
				End If
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As Double) As Double Implements ITransaction.Filter
				If Me.IsStopEnabled = False Then Return Me.Filter(Value)
				If ValueStop = 0 Then Return Me.Filter(Value)
				If MyListOfValue.Count = 0 Then
					MyPriceStart = Value.OpenNext + MyTransactionCost
					MyPriceStopCount = 0
					MyPriceStop = ValueStop
				End If
				If IsStopped Then
					MyPriceStopCount = MyPriceStopCount + 1
				Else
					If Me.IsStopReverse Then
						'allow the stop price to decrease
						MyPriceStop = ValueStop
					Else
						If ValueStop > MyPriceStop Then
							MyPriceStop = ValueStop
						End If
					End If
					MyTransactionCount = MyTransactionCount + 1
					MyPriceLast = Value
					If Value.Low <= MyPriceStop Then
						IsStopped = True
						MyPriceStopCount = 0
						MyPriceVolStop = Value
						MyPriceTransactionStop = MyPriceStop
						MyFilterLast = (MyPriceTransactionStop - MyPriceStart) / MyPriceStart
					Else
						MyPriceTransactionStop = Value.Last
						MyFilterLast = (MyPriceLast.Last - MyPriceStart) / MyPriceStart
					End If
				End If
				MyListOfValue.Add(MyFilterLast)
				Return MyFilterLast
			End Function

			Public ReadOnly Property IsStop As Boolean Implements ITransaction.IsStop
				Get
					Return IsStopped
				End Get
			End Property

			Public ReadOnly Property PriceStop As Double Implements ITransaction.PriceStop
				Get
					Return MyPriceStop
				End Get
			End Property

			Public ReadOnly Property PriceStopValue As IPriceVol Implements ITransaction.PriceStopValue
				Get
					Return MyPriceVolStop
				End Get
			End Property

			Public ReadOnly Property PriceTransactionStart As Double Implements ITransaction.PriceTransactionStart
				Get
					Return MyPriceTransactionStart
				End Get
			End Property

			Public ReadOnly Property PriceTransactionStop As Double Implements ITransaction.PriceTransactionStop
				Get
					Return MyPriceTransactionStop
				End Get
			End Property

			Public Property IsStopReverse As Boolean Implements ITransaction.IsStopReverse

			Public Function FilterLast() As Double Implements ITransaction.FilterLast
				Return MyFilterLast
			End Function

			Public Function Last() As IPriceVol Implements ITransaction.Last
				Return MyPriceLast
			End Function

			Public ReadOnly Property Count As Integer Implements ITransaction.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property CountStop As Integer Implements ITransaction.CountStop
				Get
					Return MyPriceStopCount
				End Get
			End Property

			Public ReadOnly Property TransactionCount As Integer Implements ITransaction.TransactionCount
				Get
					Return MyTransactionCount
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements ITransaction.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements ITransaction.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements ITransaction.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property TransactionCost As Double Implements ITransaction.TransactionCost
				Get
					Return MyTransactionCost
				End Get
			End Property

			Public Property Tag As String Implements ITransaction.Tag

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function

			Public ReadOnly Property Type As ITransaction.enuTransactionType Implements ITransaction.Type
				Get
					Return ITransaction.enuTransactionType.StockBuy
				End Get
			End Property

			Public Property IsStopEnabled As Boolean Implements ITransaction.IsStopEnabled
		End Class
#End Region
#Region "TransactionStockSell"
		Public Class TransactionStockSell
			Implements ITransaction

			Private MyPriceStart As Double
			Private MyPriceLast As IPriceVol
			Private MyFilterLast As Double
			Private MyTransactionCost As Double
			Private MyTransactionCount As Integer
			Private MyPriceTransactionStart As Double
			Private MyPriceTransactionStop As Double
			Private MyListOfValue As ListScaled
			Private MyPriceStop As Double
			Private IsStopped As Boolean
			Private MyPriceStopCount As Integer
			Private MyPriceVolStop As IPriceVol

			Public Sub New()
				Me.New(0)
			End Sub

			Public Sub New(ByVal TransactionCost As Double)
				MyTransactionCost = TransactionCost
				MyListOfValue = New ListScaled
				Me.IsStopEnabled = True  'by default
			End Sub

			Public Function Filter(ByRef Value As IPriceVol) As Double Implements ITransaction.Filter
				If MyListOfValue.Count = 0 Then
					MyPriceStart = Value.OpenNext - MyTransactionCost
					MyPriceStopCount = 0
					MyPriceStop = Value.Last
				End If
				If IsStopped Then
					MyPriceStopCount = MyPriceStopCount + 1
				Else
					MyTransactionCount = MyTransactionCount + 1
					MyPriceLast = Value
					MyPriceStop = Value.Last
					MyPriceTransactionStop = MyPriceStop
					MyFilterLast = (MyPriceStart - MyPriceLast.Last) / MyPriceStart
				End If
				MyListOfValue.Add(MyFilterLast)
				Return MyFilterLast
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As Double) As Double Implements ITransaction.Filter
				If Me.IsStopEnabled = False Then Return Me.Filter(Value)
				If ValueStop = 0 Then Return Me.Filter(Value)
				If MyListOfValue.Count = 0 Then
					MyPriceStart = Value.OpenNext - MyTransactionCost
					MyPriceStopCount = 0
					MyPriceStop = ValueStop
				End If
				If IsStopped Then
					'check if we should get out of the stop condition
					MyPriceStopCount = MyPriceStopCount + 1
				Else
					If Me.IsStopReverse Then
						MyPriceStop = ValueStop
					Else
						If ValueStop < MyPriceStop Then
							MyPriceStop = ValueStop
						End If
					End If
					MyTransactionCount = MyTransactionCount + 1
					MyPriceLast = Value
					If Value.High >= MyPriceStop Then
						IsStopped = True
						MyPriceStopCount = 0
						MyPriceVolStop = Value
						MyPriceTransactionStop = MyPriceStop
						MyFilterLast = (MyPriceStart - MyPriceTransactionStop) / MyPriceStart
					Else
						MyPriceTransactionStop = Value.Last
						MyFilterLast = (MyPriceStart - MyPriceLast.Last) / MyPriceStart
					End If
				End If
				MyListOfValue.Add(MyFilterLast)
				Return MyFilterLast
			End Function

			Public Function Filter(ByRef Value As IPriceVol, ByVal ValueStop As IPriceVol) As Double Implements ITransaction.Filter
				If (ValueStop Is Nothing) Or (Me.IsStopEnabled = False) Then
					Return Me.Filter(Value)
				Else
					Return Me.Filter(Value, ValueStop.High)
				End If
			End Function

			Public ReadOnly Property IsStop As Boolean Implements ITransaction.IsStop
				Get
					Return IsStopped
				End Get
			End Property

			Public ReadOnly Property PriceStop As Double Implements ITransaction.PriceStop
				Get
					Return MyPriceStop
				End Get
			End Property

			Public ReadOnly Property PriceTransactionStart As Double Implements ITransaction.PriceTransactionStart
				Get
					Return MyPriceTransactionStart
				End Get
			End Property

			Public ReadOnly Property PriceTransactionStop As Double Implements ITransaction.PriceTransactionStop
				Get
					Return MyPriceTransactionStop
				End Get
			End Property

			Public Function FilterLast() As Double Implements ITransaction.FilterLast
				Return MyFilterLast
			End Function

			Public Function Last() As IPriceVol Implements ITransaction.Last
				Return MyPriceLast
			End Function

			Public ReadOnly Property Count As Integer Implements ITransaction.Count
				Get
					Return MyListOfValue.Count
				End Get
			End Property

			Public ReadOnly Property CountStop As Integer Implements ITransaction.CountStop
				Get
					Return MyPriceStopCount
				End Get
			End Property

			Public ReadOnly Property PriceStopValue As IPriceVol Implements ITransaction.PriceStopValue
				Get
					Return MyPriceVolStop
				End Get
			End Property

			Public ReadOnly Property TransactionCount As Integer Implements ITransaction.TransactionCount
				Get
					Return MyTransactionCount
				End Get
			End Property

			Public ReadOnly Property Max As Double Implements ITransaction.Max
				Get
					Return MyListOfValue.Max
				End Get
			End Property

			Public ReadOnly Property Min As Double Implements ITransaction.Min
				Get
					Return MyListOfValue.Min
				End Get
			End Property

			Public ReadOnly Property ToList() As IList(Of Double) Implements ITransaction.ToList
				Get
					Return MyListOfValue
				End Get
			End Property

			Public ReadOnly Property TransactionCost As Double Implements ITransaction.TransactionCost
				Get
					Return MyTransactionCost
				End Get
			End Property

			Public Property Tag As String Implements ITransaction.Tag

			Public Overrides Function ToString() As String
				Return Me.FilterLast.ToString
			End Function


			Public ReadOnly Property Type As ITransaction.enuTransactionType Implements ITransaction.Type
				Get
					Return ITransaction.enuTransactionType.StockSell
				End Get
			End Property

			Public Property IsStopReverse As Boolean Implements ITransaction.IsStopReverse

			Public Property IsStopEnabled As Boolean Implements ITransaction.IsStopEnabled
		End Class
#End Region


	End Namespace  'Filter
#End Region
End Namespace  'MathPlus

