Imports System.Data
Imports System.Data.SqlClient

Public Class MenuR
    Public acceso As String
    Dim conexion As New SqlConnection("Data Source=100.48.0.7;Initial Catalog=REPORTS; Persist Security Info=True;User ID=sa; Password=sa;")

    Private Sub SalirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalirToolStripMenuItem.Click
        Dim msg As String
        Dim title As String
        Dim style As MsgBoxStyle
        Dim Response As MsgBoxResult
        msg = "Esta seguro que desea salir?"
        style = MsgBoxStyle.DefaultButton2 Or
            MsgBoxStyle.Critical Or MsgBoxStyle.YesNo
        title = "Salir"
        Response = MsgBox(msg, style, title)
        If Response = MsgBoxResult.Yes Then
            Me.Close()
            End
        End If
    End Sub


    Private Sub MenuR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '********************************************************************************************************
        'Query acceso a Menus
        Dim menus As New SqlCommand("SELECT R_USERS.ID, R_USERS.NOMBRE, R_USERS.APELLIDO_P, R_USERS.APELLIDO_M, R_USERS.USUARIO," +
                                    "R_USERS.CONTRASENA, R_USERS.ACCESO, R_MENUS.M_ALMACEN, R_MENUS.M_COMPRAS," +
        "R_MENUS.M_CONTABILIDA, R_MENUS.M_CREDITO, R_MENUS.M_PRODUCCION, R_MENUS.M_VENTAS, R_MENUS.M_COMISIONES,R_MENUS.M_DIRECCION," +
        "R_MENUS.M_RHUMANOS,R_USERS.DEPTO,R_USERS.EMPRESA FROM R_USERS INNER JOIN R_MENUS ON R_USERS.ID = R_MENUS.ID WHERE USUARIO = @usuario", conexion)
        menus.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre


        Dim adapMenus As New SqlDataAdapter(menus)
        Dim table As New DataTable()
        adapMenus.Fill(table)

        '********************************************************************************************************
        'Asignacion Variables a tablas Menus
        Dim nombre = table.Rows(0)(1).ToString()
        Dim apellido_p = table.Rows(0)(2).ToString()
        Dim acceso = table.Rows(0)(6).ToString()
        Dim almacen = table.Rows(0)(7).ToString()
        Dim compras = table.Rows(0)(8).ToString()
        Dim contabilidad = table.Rows(0)(9).ToString()
        Dim credito = table.Rows(0)(10).ToString()
        Dim produccion = table.Rows(0)(11).ToString()
        Dim ventas = table.Rows(0)(12).ToString()
        Dim comisiones = table.Rows(0)(13).ToString()
        Dim direccion = table.Rows(0)(14).ToString()
        Dim rhumanos = table.Rows(0)(15).ToString()
        Dim dpto = table.Rows(0)(16).ToString()
        Dim empresa = table.Rows(0)(17).ToString()

        Module1.dptos = dpto
        'Module1.Empresas = empresa

        Label1.Text = nombre + " " + apellido_p
        Text = "Colorantes Importados S.A. de C.V."


        '********************************************************************************************************
        'Acceso a Menu Almacen
        If almacen = True Then
            AlmacenToolStripMenuItem.Enabled = True
        Else
            AlmacenToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Credito
        If credito = True Then
            CreditoToolStripMenuItem.Enabled = True
        Else
            CreditoToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Compras
        If compras = True Then
            ComprasToolStripMenuItem.Enabled = True
        Else
            ComprasToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Contabilidad
        If contabilidad = True Then
            ContabilidadToolStripMenuItem.Enabled = True
        Else
            ContabilidadToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Produccion
        If produccion = True Then
            ProduccionToolStripMenuItem.Enabled = True
        Else
            ProduccionToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Ventas
        If ventas = True Then
            VentasToolStripMenuItem.Enabled = True
        Else
            VentasToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Comisiones
        If comisiones = True Then
            ComisionesToolStripMenuItem.Enabled = True
        Else
            ComisionesToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu Direccion
        If direccion = True Then
            DireccionToolStripMenuItem.Enabled = True
        Else
            DireccionToolStripMenuItem.Enabled = False
        End If

        'Acceso a Menu R.Humanos
        If rhumanos = True Then
            RHumanosToolStripMenuItem.Enabled = True
        Else
            RHumanosToolStripMenuItem.Enabled = False
        End If
        '********************************************************************************************************
        'Query acceso a SubMenu Almacen
        Dim subAlmacen As New SqlCommand("SELECT R_USERS.ID, R_USERS.USUARIO, R_M_ALMACEN.A_EXI_GEN," +
                                         "R_M_ALMACEN.A_EXI_ALM,R_M_ALMACEN.A_EXI_DIA," +
                                         "R_M_ALMACEN.A_OC_CON FROM R_USERS INNER JOIN R_M_ALMACEN ON R_USERS.ID = R_M_ALMACEN.ID WHERE R_USERS.USUARIO = @usuario", conexion)
        subAlmacen.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre

        Dim adapAlmacen As New SqlDataAdapter(subAlmacen)
        Dim rowAlmacen As New DataTable()
        adapAlmacen.Fill(rowAlmacen)

        'Asignacion Variables a tablas Almacen
        Dim A_EXI_GEN = rowAlmacen.Rows(0)(2).ToString()
        Dim A_EXI_ALM = rowAlmacen.Rows(0)(3).ToString()
        Dim A_EXI_DIA = rowAlmacen.Rows(0)(4).ToString()
        Dim A_OC_CON = rowAlmacen.Rows(0)(5).ToString()

        'Acceso SubMenu ExistenciaGeneral
        If A_EXI_GEN = True Then
            ExistenciaGeneralToolStripMenuItem.Enabled = True
        Else
            ExistenciaGeneralToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu ExistenciaAlmacen
        If A_EXI_ALM = True Then
            ExistenciaAlmacenToolStripMenuItem.Enabled = True
        Else
            ExistenciaAlmacenToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu ExistenciaProductoDiasDeInventario
        If A_EXI_DIA = True Then
            ExistenciaProductoDiasDeInventarioToolStripMenuItem.Enabled = True
        Else
            ExistenciaProductoDiasDeInventarioToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu ReciboOrdenCompraContable

        If A_OC_CON = True Then
            ReciboOCContableToolStripMenuItem.Enabled = True
        Else
            ReciboOCContableToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Credito
        Dim subCredito As New SqlCommand("SELECT R_USERS.ID,R_USERS.USUARIO, R_M_CREDITO.C_COB_GEN, R_M_CREDITO.C_COB_DET," +
                                       "R_M_CREDITO.C_NOT_CRE, R_M_CREDITO.C_APL_NOT_CRE, R_M_CREDITO.C_VEN FROM R_USERS INNER JOIN R_M_CREDITO ON R_USERS.ID = R_M_CREDITO.ID WHERE R_USERS.USUARIO = @usuario", conexion)
        subCredito.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre


        Dim adapCredito As New SqlDataAdapter(subCredito)
        Dim rowCredito As New DataTable()
        adapCredito.Fill(rowCredito)

        'Asignacion Variables a tablas Credito
        Dim C_COB_GEN = rowCredito.Rows(0)(2).ToString()
        Dim C_COB_DET = rowCredito.Rows(0)(3).ToString()
        Dim C_NOT_CRE = rowCredito.Rows(0)(4).ToString()
        Dim C_APL_NOT_CRE = rowCredito.Rows(0)(5).ToString()
        Dim C_VEN = rowCredito.Rows(0)(6).ToString()

        'Acceso SubMenu CobranzaGeneral
        If C_COB_GEN = True Then
            CobranzaGeneralToolStripMenuItem.Enabled = True
        Else
            CobranzaGeneralToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu CobranzaDetalle
        If C_COB_DET = True Then
            CobranzaDetalleToolStripMenuItem.Enabled = True
        Else
            CobranzaDetalleToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu NotasDeCredito
        If C_NOT_CRE = True Then
            NotasCreditoPrincipalToolStripMenuItem.Enabled = True
        Else
            NotasCreditoPrincipalToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu AplicacionNotasDeCredit
        If C_APL_NOT_CRE = True Then
            AplicacionNotasDeCreditoToolStripMenuItem.Enabled = True
        Else
            AplicacionNotasDeCreditoToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu Vencimientos
        If C_VEN = True Then
            VencimientosToolStripMenuItem.Enabled = True
        Else
            VencimientosToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Ventas
        Dim subVentas As New SqlCommand("SELECT R_USERS.ID, R_USERS.USUARIO, R_M_VENTAS.V_TAL_EMB, R_M_VENTAS.V_PED_DET_UNI," +
                                       "R_M_VENTAS.V_DET_VEN, R_M_VENTAS.V_PRO_COL,R_M_VENTAS.V_CLI_FIC FROM R_USERS INNER JOIN R_M_VENTAS ON R_USERS.ID = R_M_VENTAS.ID WHERE R_USERS.USUARIO = @usuario", conexion)
        subVentas.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre


        Dim adapVentas As New SqlDataAdapter(subVentas)
        Dim rowVentas As New DataTable()
        adapVentas.Fill(rowVentas)

        'Asignacion Variables a tablas Ventas
        Dim V_TAL_EMB = rowVentas.Rows(0)(2).ToString()
        Dim V_PED_DET_UNI = rowVentas.Rows(0)(3).ToString()
        Dim V_DET_VEN = rowVentas.Rows(0)(4).ToString()
        Dim V_PRO_COL = rowVentas.Rows(0)(5).ToString()
        Dim V_CLI_FIC = rowVentas.Rows(0)(6).ToString()

        'Acceso SubMenu TalonDe
        If V_TAL_EMB = True Then
            TalonDeToolStripMenuItem.Enabled = True
        Else
            TalonDeToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu Pedidos
        If V_PED_DET_UNI = True Then
            PedidosToolStripMenuItem.Enabled = True
        Else
            PedidosToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu DetalleDeVentas
        If V_DET_VEN = True Then
            DetalleDeVentasToolStripMenuItem.Enabled = True
        Else
            DetalleDeVentasToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu PromediosColorante
        If V_PRO_COL = True Then
            PromediosColoranteToolStripMenuItem.Enabled = True
        Else
            PromediosColoranteToolStripMenuItem.Enabled = False
        End If

        'Acceso a SubMenu FichaClientes
        If V_CLI_FIC = True Then
            FichaDeClientesToolStripMenuItem.Enabled = True
        Else
            FichaDeClientesToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Comisiones 
        Dim subComisiones As New SqlCommand("SELECT        R_USERS.ID AS Expr1, R_USERS.USUARIO, R_M_COMISIONES.C_MBA_COM_CLI, R_M_COMISIONES.C_MBA_COM_COB, R_M_COMISIONES.C_MBA_COM_FAM FROM R_USERS INNER JOIN R_M_COMISIONES ON R_USERS.ID = R_M_COMISIONES.ID WHERE R_USERS.USUARIO = @usuario", conexion)
        subComisiones.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre


        Dim adapComisiones As New SqlDataAdapter(subComisiones)
        Dim rowComisiones As New DataTable()
        adapComisiones.Fill(rowComisiones)

        'Asignacion Variables a tablas Comisiones

        Dim C_ORD_CLI = rowComisiones.Rows(0)(2).ToString()
        Dim C_ORD_COB = rowComisiones.Rows(0)(3).ToString()
        Dim C_ORD_FAM = rowComisiones.Rows(0)(4).ToString()

        'Acceso SubMenu MBA3 Orden Cliente
        If C_ORD_CLI = True Then
            MBA3OrdenClienteToolStripMenuItem.Enabled = True
        Else
            MBA3OrdenClienteToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu MBA3 Orden Cobro
        If C_ORD_COB = True Then
            MBA3OrdenCobroToolStripMenuItem.Enabled = True
        Else
            MBA3OrdenCobroToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu MBA3 Orden Familia
        If C_ORD_FAM = True Then
            MBA3OrdenFamiliaToolStripMenuItem.Enabled = True
        Else
            MBA3OrdenFamiliaToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Contabilidad
        Dim subContabilidad As New SqlCommand("SELECT R_USERS.ID, R_USERS.USUARIO, R_M_CONTABILIDAD.C_PRE, R_M_CONTABILIDAD.C_PRE_CUE," +
                                              "R_M_CONTABILIDAD.C_PRE_COM, R_M_CONTABILIDAD.C_SAL_BOD_COS FROM R_USERS INNER JOIN R_M_CONTABILIDAD ON R_USERS.ID = R_M_CONTABILIDAD.ID WHERE R_USERS.USUARIO = @usuario", conexion)
        subContabilidad.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre

        Dim adapContabilidad As New SqlDataAdapter(subContabilidad)
        Dim rowContabilidad As New DataTable()
        adapContabilidad.Fill(rowContabilidad)

        'Asignacion Variables a tablas Contabilidad
        Dim C_PRE = rowContabilidad.Rows(0)(2).ToString()
        Dim C_PRE_CUE = rowContabilidad.Rows(0)(3).ToString()
        Dim C_PRE_COM = rowContabilidad.Rows(0)(4).ToString()
        Dim C_SAL_BOD_COS = rowContabilidad.Rows(0)(5).ToString()

        Module1.pre_com = C_PRE_COM
        'Acceso SubMenu Presupuesto
        If C_PRE = True Then
            PresupuestoToolStripMenuItem.Enabled = True
        Else
            PresupuestoToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu Ppto. Cuentas

        If C_PRE_CUE = True Then
            CuentasToolStripMenuItem.Enabled = True
        Else
            CuentasToolStripMenuItem.Enabled = False
        End If


        'Acceso SubMenu Saldo Bodega Costo

        If C_SAL_BOD_COS = True Then
            SaldosBodegaCostosToolStripMenuItem.Enabled = True
        Else
            SaldosBodegaCostosToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Compras
        Dim subCompras As New SqlCommand("SELECT R_USERS.ID, R_USERS.USUARIO, R_M_COMPRAS.C_REC_MPR FROM R_USERS INNER JOIN R_M_COMPRAS ON R_USERS.ID = R_M_COMPRAS.ID WHERE R_USERS.USUARIO = @usuario", conexion)

        subCompras.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre

        Dim adapCompras As New SqlDataAdapter(subCompras)
        Dim rowCompras As New DataTable()
        adapCompras.Fill(rowCompras)

        'Asignacion Variables a tablas Compras
        Dim C_REC_MPR = rowCompras.Rows(0)(2).ToString()

        'Acceso SubMenu Recepcion MP
        If C_REC_MPR = True Then
            RecepcionesMPToolStripMenuItem.Enabled = True
        Else
            RecepcionesMPToolStripMenuItem.Enabled = False
        End If

        '********************************************************************************************************
        'Query acceso a SubMenu Direccion
        Dim subDireccion As New SqlCommand("SELECT R_USERS.ID, R_USERS.USUARIO, R_M_DIRECCION.D_VEN_AECP,R_M_DIRECCION.D_VEN_AECP_D FROM R_USERS INNER JOIN R_M_DIRECCION ON R_USERS.ID = R_M_DIRECCION.ID WHERE R_USERS.USUARIO = @usuario", conexion)

        subDireccion.Parameters.Add("@usuario", SqlDbType.VarChar).Value = Module1.nombre

        Dim adapDireccion As New SqlDataAdapter(subDireccion)
        Dim rowDireccion As New DataTable()
        adapDireccion.Fill(rowDireccion)

        'Asignacion Variables a tablas Compras
        Dim D_VEN_AECP = rowDireccion.Rows(0)(2).ToString()
        Dim D_VEN_AECP_D = rowDireccion.Rows(0)(3).ToString()

        'Acceso SubMenu Ventas AECP
        If D_VEN_AECP = True Then
            VentasAECPToolStripMenuItem.Enabled = True
        Else
            VentasAECPToolStripMenuItem.Enabled = False
        End If

        'Acceso SubMenu Ventas AECP D
        If D_VEN_AECP_D = True Then
            VentasAECPDatosToolStripMenuItem.Enabled = True
        Else
            VentasAECPDatosToolStripMenuItem.Enabled = False
        End If
    End Sub


    Private Sub TalonDeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TalonDeToolStripMenuItem.Click
        Dim a As New VenFacTalEmb
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub PedidosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PedidosToolStripMenuItem.Click
        Dim a As New VenPedDetUni
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub CobranzaGeneralToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CobranzaGeneralToolStripMenuItem.Click
        Dim a As New CreCobGen
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub CobranzaDetalleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CobranzaDetalleToolStripMenuItem.Click
        Dim a As New CreCobDet
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub ExistenciaGeneralToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExistenciaGeneralToolStripMenuItem.Click
        Dim a As New AlmExiGen
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub ExistenciaAlmacenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExistenciaAlmacenToolStripMenuItem.Click
        Dim a As New AlmExiAlm
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub ExistenciaProductoDiasDeInventarioToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExistenciaProductoDiasDeInventarioToolStripMenuItem.Click
        Dim a As New AlmExiProDiaInv
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub DetalleDeVentasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DetalleDeVentasToolStripMenuItem.Click
        Dim a As New VenDetVen
        a.Show()
        Me.FindForm()
    End Sub
    Private Sub AplicacionNotasDeCreditoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AplicacionNotasDeCreditoToolStripMenuItem.Click
        Dim a As New CreAplNotCre
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub PromediosColoranteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PromediosColoranteToolStripMenuItem.Click
        Dim FrmProCol As New VenProCol
        FrmProCol.Show()
        Me.FindForm()
    End Sub

    Private Sub MBA3OrdenClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MBA3OrdenClienteToolStripMenuItem.Click
        Dim FrmMBACli As New ComMBACli
        FrmMBACli.Show()
        Me.FindForm()
    End Sub

    Private Sub MBA3OrdenCobroToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MBA3OrdenCobroToolStripMenuItem.Click
        Dim FrmMBACob As New ComMBACob
        FrmMBACob.Show()
        Me.FindForm()
    End Sub

    Private Sub MBA3OrdenFamiliaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MBA3OrdenFamiliaToolStripMenuItem.Click
        Dim FrmMBAFam As New ComMBAFam
        FrmMBAFam.Show()
        Me.FindForm()
    End Sub

    Private Sub VencAgenteClienteRToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub NotasCreditoPrincipalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NotasCreditoPrincipalToolStripMenuItem.Click
        Dim FrmNotCrePri As New CreNotCrePri
        FrmNotCrePri.Show()
        Me.FindForm()
    End Sub

    Private Sub TablaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TablaToolStripMenuItem.Click
        Dim FrmVenTab As New CreVenTab
        FrmVenTab.Show()
        Me.FindForm()
    End Sub

    Private Sub AgenteClienteRangoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgenteClienteRangoToolStripMenuItem.Click
        Dim FrnVenAgeCliRan As New CreVenAgeCliRang
        FrnVenAgeCliRan.Show()
        Me.FindForm()
    End Sub

    Private Sub AgenteRangoClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AgenteRangoClienteToolStripMenuItem.Click
        Dim FrmVenAgeRanCli As New CreVenAgeRanCli
        FrmVenAgeRanCli.Show()
        Me.FindForm()
    End Sub

    Private Sub RangoAgenteClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RangoAgenteClienteToolStripMenuItem.Click
        Dim FrmVenRanAgeCli As New CreVenRanAgeCli
        FrmVenRanAgeCli.Show()
        Me.FindForm()
    End Sub

    Private Sub GrafAgenteClienteRangoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrafAgenteClienteRangoToolStripMenuItem.Click
        Dim FrmVenGrafAgeCliRan As New CreVenGrafAgeCliRan
        FrmVenGrafAgeCliRan.Show()
        Me.FindForm()
    End Sub

    Private Sub GrafAgenteRangoClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrafAgenteRangoClienteToolStripMenuItem.Click
        Dim FrmVenGrafAgeRanCli As New CreVenGrafAgeRanCli
        FrmVenGrafAgeRanCli.Show()
        Me.FindForm()
    End Sub

    Private Sub GrafRangoAgenteClienteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GrafRangoAgenteClienteToolStripMenuItem.Click
        Dim FrmVenGraRan As New CreVenGrafRanAgeCli
        FrmVenGraRan.Show()
        Me.FindForm()
    End Sub

    Private Sub FichaDeClientesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FichaDeClientesToolStripMenuItem.Click
        Dim FrmVenCliFic As New VenCliFic
        FrmVenCliFic.Show()
        Me.FindForm()
    End Sub

    Private Sub CuentasToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CuentasToolStripMenuItem.Click
        Dim FrmVenGrafAgeCliRan As New ConPreCue
        FrmVenGrafAgeCliRan.Show()
        Me.FindForm()
    End Sub
    Private Sub ComparativoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComparativoToolStripMenuItem.Click
        Dim FrmVenGrafAgeCliRan As New ConPreCom
        FrmVenGrafAgeCliRan.Show()
        Me.FindForm()
    End Sub

    Private Sub RecepcionesMPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RecepcionesMPToolStripMenuItem.Click
        Dim a As New ComRecMpr
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub VentasAECPToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VentasAECPToolStripMenuItem.Click
        Dim a As New DirVenAECP
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub VentasAECPDatosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VentasAECPDatosToolStripMenuItem.Click
        Dim a As New DirVentasAECPDAT
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub SaldosBodegaCostosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaldosBodegaCostosToolStripMenuItem.Click
        Dim a As New ConSalBodCos
        a.Show()
        Me.FindForm()
    End Sub

    Private Sub ReciboOCContableToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReciboOCContableToolStripMenuItem.Click
        Dim a As New AlmRecOcC
        a.Show()
        Me.FindForm()
    End Sub
End Class