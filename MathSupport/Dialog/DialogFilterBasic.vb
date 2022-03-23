Imports System.Windows.Forms

Friend Class DialogFilterBasic
  Private MyTitle As String
  Private MyRate As Integer

  Public Sub New(ByVal Title As String)
    Me.New(Title, 10)
  End Sub

  Public Sub New(ByVal Title As String, ByVal Rate As Integer)

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.
    MyTitle = Title
    MyRate = Rate
  End Sub

  Private Sub DialogFilterLowPassExp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Me.Text = MyTitle
    TextBoxNumericFilterRate.Value = MyRate
  End Sub

  Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.Close()
  End Sub

  Public ReadOnly Property Rate As Integer
    Get
      Return CInt(TextBoxNumericFilterRate.Value)
    End Get
  End Property

  Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
    Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.Close()
  End Sub

  Private Sub TextBoxNumericFilterRate_TextChanged(sender As Object, e As EventArgs) Handles TextBoxNumericFilterRate.TextChanged

  End Sub
End Class
