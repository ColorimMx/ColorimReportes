Public Class VenProCol

    Private Sub CrystalReportViewer1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CrystalReportViewer1.Load
        Dim CR1 As New ProCol
        Me.CrystalReportViewer1.ReportSource = CR1
        Me.CrystalReportViewer1.LogOnInfo.Item(0).ConnectionInfo.UserID = "sa"
        Me.CrystalReportViewer1.LogOnInfo.Item(0).ConnectionInfo.Password = "sa"
    End Sub
End Class