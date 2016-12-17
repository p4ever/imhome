Public Class PopupForm
    Public Sub New()
        InitializeComponent()
        Icon = My.Resources.TrayIcon
    End Sub

    Private Sub CancelButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
    Handles CancelFormButton.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub CloseAppButton_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        Me.DialogResult = System.Windows.Forms.DialogResult.Abort
        Me.Close()
    End Sub

    Private Sub PopupForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Application.Exit()
    End Sub
End Class