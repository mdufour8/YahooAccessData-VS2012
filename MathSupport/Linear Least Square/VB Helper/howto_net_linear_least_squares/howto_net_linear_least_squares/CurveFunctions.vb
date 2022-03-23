Module CurveFunctions
    ' Find the least squares linear fit.
    ' Return the total error.
    Public Function FindLinearLeastSquaresFit(ByVal points As List(Of PointF), ByRef m As Double, ByRef b As Double) As Double
        ' Perform the calculation.
        ' Find the values S1, Sx, Sy, Sxx, and Sxy.
        Dim S1 As Double = points.Count
        Dim Sx As Double = 0
        Dim Sy As Double = 0
        Dim Sxx As Double = 0
        Dim Sxy As Double = 0
        For Each pt As PointF In points
            Sx += pt.X
            Sy += pt.Y
            Sxx += pt.X * pt.X
            Sxy += pt.X * pt.Y
        Next pt

        ' Solve for m and b.
        m = (Sxy * S1 - Sx * Sy) / (Sxx * S1 - Sx * Sx)
        b = (Sxy * Sx - Sy * Sxx) / (Sx * Sx - S1 * Sxx)

        Return Math.Sqrt(ErrorSquared(points, m, b))
    End Function

    ' Return the error squared.
    Public Function ErrorSquared(ByVal points As List(Of PointF), ByVal m As Double, ByVal b As Double) As Double
        Dim total As Double = 0
        For Each pt As PointF In points
            Dim dy As Double = pt.Y - (m * pt.X + b)
            total += dy * dy
        Next pt
        Return total
    End Function
End Module
