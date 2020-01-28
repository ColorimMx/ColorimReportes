Public Class Login

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim Response As MsgBoxResult
        msg = "Esta seguro que desea salir?"
        style = MsgBoxStyle.DefaultButton2 Or
            MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Sair"
        Response = MsgBox(msg, style, title)
        If Response = MsgBoxResult.Yes Then
            Me.Close()
            End
        End If
    End Sub

    Private Sub btnAccesar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAccesar.Click
        Module1.nombre = txtUser.Text
        Dim ac As New Acceso
        Dim usuario As String = txtUser.Text
        Dim contraseña As String = txtPassword.Text
        Dim P As New MenuR

        Dim acceso As String = ac.logear(usuario, contraseña)
        Select acceso
            Case "admin"
                P.acceso = "admin"
                P.Show()
                Me.Dispose(False)
            Case "user"
                P.acceso = "user"
                P.Show()
                Me.Dispose(False)
            Case Else
                txtUser.Clear()
                txtPassword.Clear()
                txtUser.Focus()
        End Select
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class