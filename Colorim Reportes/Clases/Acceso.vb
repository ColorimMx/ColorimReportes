Imports System.Data
Imports System.Data.SqlClient

Public Class Acceso
    Private Property conexion As New SqlConnection

    Public Sub New()
        Me.conexion.ConnectionString = "Data Source=100.48.0.7;Initial Catalog=REPORTS; Persist Security Info=True;User ID=sa; Password=sa;"
    End Sub

    Public Function logear(ByVal USUARIO As String, ByVal PASS As String) As String
        Dim adapter As SqlDataAdapter
        Dim acceso As String = ""
        Dim ds As New DataSet
        Dim sql As String
        sql = "select acceso from r_users where usuario ='" & USUARIO & "' and contrasena ='" & PASS & "' and habilitado = 1"
        Try
            Me.conexion.Open()
            adapter = New SqlDataAdapter(sql, conexion.ConnectionString)
            adapter.Fill(ds)
            acceso = ds.Tables(0).Rows(0).Item(0).ToString
        Catch ex As Exception
            MsgBox(" Usuario/Contraseña Invalidos ", vbCritical)
        Finally
            Me.conexion.Close()
        End Try
        Return acceso
    End Function
End Class