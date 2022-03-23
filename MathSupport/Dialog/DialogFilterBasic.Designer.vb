<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DialogFilterBasic
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
    Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
    Me.OK_Button = New System.Windows.Forms.Button()
    Me.Cancel_Button = New System.Windows.Forms.Button()
    Me.LabelFilterRate = New System.Windows.Forms.Label()
    Me.TextBoxNumericFilterRate = New TextBoxNumeric.TextBoxNumeric()
    Me.TableLayoutPanel1.SuspendLayout()
    Me.SuspendLayout()
    '
    'TableLayoutPanel1
    '
    Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.TableLayoutPanel1.ColumnCount = 2
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
    Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
    Me.TableLayoutPanel1.Location = New System.Drawing.Point(69, 118)
    Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
    Me.TableLayoutPanel1.RowCount = 1
    Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
    Me.TableLayoutPanel1.Size = New System.Drawing.Size(219, 45)
    Me.TableLayoutPanel1.TabIndex = 0
    '
    'OK_Button
    '
    Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.OK_Button.Location = New System.Drawing.Point(4, 5)
    Me.OK_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.OK_Button.Name = "OK_Button"
    Me.OK_Button.Size = New System.Drawing.Size(100, 35)
    Me.OK_Button.TabIndex = 0
    Me.OK_Button.Text = "OK"
    '
    'Cancel_Button
    '
    Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
    Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Cancel_Button.Location = New System.Drawing.Point(114, 5)
    Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.Cancel_Button.Name = "Cancel_Button"
    Me.Cancel_Button.Size = New System.Drawing.Size(100, 35)
    Me.Cancel_Button.TabIndex = 1
    Me.Cancel_Button.Text = "Cancel"
    '
    'LabelFilterRate
    '
    Me.LabelFilterRate.AutoSize = True
    Me.LabelFilterRate.Location = New System.Drawing.Point(24, 30)
    Me.LabelFilterRate.Name = "LabelFilterRate"
    Me.LabelFilterRate.Size = New System.Drawing.Size(87, 20)
    Me.LabelFilterRate.TabIndex = 1
    Me.LabelFilterRate.Text = "Filter Rate:"
    '
    'TextBoxNumericFilterRate
    '
    Me.TextBoxNumericFilterRate.Format = "{0:n0}"
    Me.TextBoxNumericFilterRate.IsBeepOnOutOfRange = False
    Me.TextBoxNumericFilterRate.IsMaximum = True
    Me.TextBoxNumericFilterRate.IsMinimum = True
    Me.TextBoxNumericFilterRate.Location = New System.Drawing.Point(37, 54)
    Me.TextBoxNumericFilterRate.MaximumValue = 1000.0R
    Me.TextBoxNumericFilterRate.MinimumValue = 1.0R
    Me.TextBoxNumericFilterRate.Name = "TextBoxNumericFilterRate"
    Me.TextBoxNumericFilterRate.Size = New System.Drawing.Size(100, 26)
    Me.TextBoxNumericFilterRate.TabIndex = 2
    Me.TextBoxNumericFilterRate.Text = "10"
    Me.TextBoxNumericFilterRate.Value = 10.0R
    '
    'DialogFilterBasic
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.CancelButton = Me.Cancel_Button
    Me.ClientSize = New System.Drawing.Size(305, 181)
    Me.Controls.Add(Me.TextBoxNumericFilterRate)
    Me.Controls.Add(Me.LabelFilterRate)
    Me.Controls.Add(Me.TableLayoutPanel1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "DialogFilterBasic"
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
    Me.Text = "Filter Low Pass Exponential"
    Me.TableLayoutPanel1.ResumeLayout(False)
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
  Friend WithEvents OK_Button As System.Windows.Forms.Button
  Friend WithEvents Cancel_Button As System.Windows.Forms.Button
  Friend WithEvents LabelFilterRate As System.Windows.Forms.Label
  Friend WithEvents TextBoxNumericFilterRate As TextBoxNumeric.TextBoxNumeric

End Class
