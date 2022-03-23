Imports System.Drawing.Drawing2D

Public Class Form1
    ' Drawing constants.
    Private Const Xmin As Single = -10.0F
    Private Const Xmax As Single = 10.0F
    Private Const Ymin As Single = -10.0F
    Private Const Ymax As Single = 10.0F
    Private DrawingTransform, InverseTransform As Matrix

    Private HasSolution As Boolean = False
    Private BestM, BestB As Double
    Private Points As New List(Of PointF)()

    ' Make a drawing transformation.
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim world_rect As New RectangleF(Xmin, Ymin, Xmax - Xmin, Ymax - Ymin)
        Dim pts() As PointF = _
        { _
            New PointF(0, picGraph.ClientSize.Height), _
            New PointF(picGraph.ClientSize.Width, picGraph.ClientSize.Height), _
            New PointF(0, 0) _
        }
        DrawingTransform = New Matrix(world_rect, pts)
        InverseTransform = DrawingTransform.Clone()
        InverseTransform.Invert()
    End Sub

    ' Save a new point.
    Private Sub picGraph_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles picGraph.MouseClick
        ' Transform the point to world coordinates.
        Dim pts() As PointF = {New PointF(e.X, e.Y)}
        InverseTransform.TransformPoints(pts)

        ' Save the point.
        Points.Add(pts(0))
        picGraph.Refresh()
    End Sub

    ' Draw the points and best fit curve.
    Private Sub picGraph_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles picGraph.Paint
        ' Use the drawing transformation.
        e.Graphics.Transform = DrawingTransform

        ' Draw the axes.
        DrawAxes(e.Graphics)

        ' Draw the curve.
        If (HasSolution) Then
            Using thin_pen As New Pen(Color.Blue, 0)
                Dim y0 As Double = BestM * Xmin + BestB
                Dim y1 As Double = BestM * Xmax + BestB
                e.Graphics.DrawLine(thin_pen, _
                    Xmin, CSng(y0), Xmax, CSng(y1))
            End Using
        End If

        ' Draw the points.
        Const dx As Single = (Xmax - Xmin) / 100
        Const dy As Single = (Ymax - Ymin) / 100
        Using thin_pen As New Pen(Color.Black, 0)
            For Each pt As PointF In Points
                e.Graphics.FillRectangle(Brushes.White, _
                    pt.X - dx, pt.Y - dy, 2 * dx, 2 * dy)
                e.Graphics.DrawRectangle(thin_pen, _
                    pt.X - dx, pt.Y - dy, 2 * dx, 2 * dy)
            Next pt
        End Using
    End Sub

    ' Draw the axes.
    Private Sub DrawAxes(ByVal gr As Graphics)
        Using thin_pen As New Pen(Color.Black, 0)
            Const xthick As Single = 0.2F
            Const ythick As Single = 0.2F
            gr.DrawLine(thin_pen, Xmin, 0, Xmax, 0)
            For x As Single = Xmin To Xmax Step 1.0F
                gr.DrawLine(thin_pen, x, -ythick, x, ythick)
            Next x
            gr.DrawLine(thin_pen, 0, Ymin, 0, Ymax)
            For y As Single = Ymin To Ymax Step 1.0F
                gr.DrawLine(thin_pen, -xthick, y, xthick, y)
            Next y
        End Using
    End Sub

    ' Clear the points.
    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        Points = New List(Of PointF)()
        HasSolution = False
        picGraph.Refresh()

        txtM.Clear()
        txtB.Clear()
        txtError.Clear()
    End Sub

    ' Find parameters for a curve fit.
    Private Sub btnFit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFit.Click
        Me.Cursor = Cursors.WaitCursor
        txtM.Clear()
        txtB.Clear()
        txtError.Clear()
        Application.DoEvents()
        Dim start_time As Date = Date.Now

        ' Find a good fit.
        FindLinearLeastSquaresFit(Points, BestM, BestB)

        Dim stop_time As Date = Date.Now
        Dim elapsed As TimeSpan = stop_time - start_time
        Console.WriteLine("Time: " & _
            elapsed.TotalSeconds.ToString("0.00") & " seconds")

        txtM.Text = BestM.ToString()
        txtB.Text = BestB.ToString()

        ' Display the error.
        ShowError()

        ' We have a solution.
        HasSolution = True
        picGraph.Refresh()

        Me.Cursor = Cursors.Default
    End Sub

    ' Regraph with the given parameters.
    Private Sub btnGraph_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGraph.Click
        BestM = Double.Parse(txtM.Text)
        BestB = Double.Parse(txtB.Text)
        ShowError()
        picGraph.Refresh()
    End Sub

    ' Display the error.
    Private Sub ShowError()
        ' Get the error.
        Dim err As Double = Math.Sqrt(ErrorSquared( _
            Points, BestM, BestB))
        txtError.Text = err.ToString()
    End Sub

Private Sub picGraph_Click(sender As Object, e As EventArgs) Handles picGraph.Click

End Sub
End Class
