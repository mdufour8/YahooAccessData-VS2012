<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
    Me.btnGraph = New System.Windows.Forms.Button()
    Me.rchEquation = New System.Windows.Forms.RichTextBox()
    Me.label3 = New System.Windows.Forms.Label()
    Me.txtError = New System.Windows.Forms.TextBox()
    Me.label2 = New System.Windows.Forms.Label()
    Me.txtB = New System.Windows.Forms.TextBox()
    Me.label1 = New System.Windows.Forms.Label()
    Me.txtM = New System.Windows.Forms.TextBox()
    Me.btnClear = New System.Windows.Forms.Button()
    Me.btnFit = New System.Windows.Forms.Button()
    Me.picGraph = New System.Windows.Forms.PictureBox()
    CType(Me.picGraph, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'btnGraph
    '
    Me.btnGraph.Location = New System.Drawing.Point(426, 337)
    Me.btnGraph.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.btnGraph.Name = "btnGraph"
    Me.btnGraph.Size = New System.Drawing.Size(112, 35)
    Me.btnGraph.TabIndex = 45
    Me.btnGraph.Text = "Graph"
    Me.btnGraph.UseVisualStyleBackColor = True
    '
    'rchEquation
    '
    Me.rchEquation.BackColor = System.Drawing.SystemColors.Control
    Me.rchEquation.BorderStyle = System.Windows.Forms.BorderStyle.None
    Me.rchEquation.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.rchEquation.Location = New System.Drawing.Point(404, 18)
    Me.rchEquation.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.rchEquation.Multiline = False
    Me.rchEquation.Name = "rchEquation"
    Me.rchEquation.Size = New System.Drawing.Size(198, 71)
    Me.rchEquation.TabIndex = 50
    Me.rchEquation.Text = "y = m * x + b"
    '
    'label3
    '
    Me.label3.AutoSize = True
    Me.label3.Location = New System.Drawing.Point(369, 222)
    Me.label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
    Me.label3.Name = "label3"
    Me.label3.Size = New System.Drawing.Size(48, 20)
    Me.label3.TabIndex = 49
    Me.label3.Text = "Error:"
    '
    'txtError
    '
    Me.txtError.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtError.Location = New System.Drawing.Point(426, 217)
    Me.txtError.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.txtError.Name = "txtError"
    Me.txtError.Size = New System.Drawing.Size(175, 26)
    Me.txtError.TabIndex = 44
    '
    'label2
    '
    Me.label2.AutoSize = True
    Me.label2.Location = New System.Drawing.Point(369, 182)
    Me.label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
    Me.label2.Name = "label2"
    Me.label2.Size = New System.Drawing.Size(22, 20)
    Me.label2.TabIndex = 48
    Me.label2.Text = "b:"
    '
    'txtB
    '
    Me.txtB.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtB.Location = New System.Drawing.Point(426, 177)
    Me.txtB.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.txtB.Name = "txtB"
    Me.txtB.Size = New System.Drawing.Size(175, 26)
    Me.txtB.TabIndex = 43
    '
    'label1
    '
    Me.label1.AutoSize = True
    Me.label1.Location = New System.Drawing.Point(369, 142)
    Me.label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
    Me.label1.Name = "label1"
    Me.label1.Size = New System.Drawing.Size(26, 20)
    Me.label1.TabIndex = 47
    Me.label1.Text = "m:"
    '
    'txtM
    '
    Me.txtM.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtM.Location = New System.Drawing.Point(426, 137)
    Me.txtM.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.txtM.Name = "txtM"
    Me.txtM.Size = New System.Drawing.Size(175, 26)
    Me.txtM.TabIndex = 42
    '
    'btnClear
    '
    Me.btnClear.Location = New System.Drawing.Point(489, 92)
    Me.btnClear.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.btnClear.Name = "btnClear"
    Me.btnClear.Size = New System.Drawing.Size(112, 35)
    Me.btnClear.TabIndex = 41
    Me.btnClear.Text = "Clear"
    Me.btnClear.UseVisualStyleBackColor = True
    '
    'btnFit
    '
    Me.btnFit.Location = New System.Drawing.Point(368, 92)
    Me.btnFit.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.btnFit.Name = "btnFit"
    Me.btnFit.Size = New System.Drawing.Size(112, 35)
    Me.btnFit.TabIndex = 40
    Me.btnFit.Text = "Fit"
    Me.btnFit.UseVisualStyleBackColor = True
    '
    'picGraph
    '
    Me.picGraph.BackColor = System.Drawing.Color.White
    Me.picGraph.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
    Me.picGraph.Location = New System.Drawing.Point(18, 18)
    Me.picGraph.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.picGraph.Name = "picGraph"
    Me.picGraph.Size = New System.Drawing.Size(340, 349)
    Me.picGraph.TabIndex = 46
    Me.picGraph.TabStop = False
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(621, 391)
    Me.Controls.Add(Me.btnGraph)
    Me.Controls.Add(Me.rchEquation)
    Me.Controls.Add(Me.label3)
    Me.Controls.Add(Me.txtError)
    Me.Controls.Add(Me.label2)
    Me.Controls.Add(Me.txtB)
    Me.Controls.Add(Me.label1)
    Me.Controls.Add(Me.txtM)
    Me.Controls.Add(Me.btnClear)
    Me.Controls.Add(Me.btnFit)
    Me.Controls.Add(Me.picGraph)
    Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.Name = "Form1"
    Me.Text = "howto_net_linear_least_squares"
    CType(Me.picGraph, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

End Sub
    Private WithEvents btnGraph As System.Windows.Forms.Button
    Private WithEvents rchEquation As System.Windows.Forms.RichTextBox
    Private WithEvents label3 As System.Windows.Forms.Label
    Private WithEvents txtError As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Private WithEvents txtB As System.Windows.Forms.TextBox
    Private WithEvents label1 As System.Windows.Forms.Label
    Private WithEvents txtM As System.Windows.Forms.TextBox
    Private WithEvents btnClear As System.Windows.Forms.Button
    Private WithEvents btnFit As System.Windows.Forms.Button
    Private WithEvents picGraph As System.Windows.Forms.PictureBox

End Class
