Public Class load1

    Private Sub load1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Increment(1)
        If ProgressBar1.Value = 100 Then
            Dim a As New Login
            a.Show()
            Me.Finalize()
            Timer1.Stop()
        End If
        Label1.Text = ProgressBar1.Value & (" %")
    End Sub
End Class