Imports System.Diagnostics
Imports System.Windows.Forms

Public Class ApuntesPeriodicos

    Public vConcepto, vtipoSql, vtipoGrid, vTxtNombre As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, filaSelec As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub ApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
        Me.KeyPreview = True
        vFecha1Enero = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker2.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        vFecha31Diciembre = Val(vAñoEjercicio) + 20
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)

        Dim TL(15) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnFiltroCuenta, "Aplica el filtro a los Registros")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSinFiltroCuenta, "Quitar el filtro a los Registros")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnFiltroConcepto, "Aplica el filtro a los Registros")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnSinFiltroConcepto, "Quitar el filtro a los Registros")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnFiltroFecha, "Aplica el filtro a los Registros")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnSinFiltroFecha, "Quitar el filtro a los Registros")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnAñadirRegistro, "Añadir Registro")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnEditarRegistro, "Editar Registro")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnEliminarRegistro, "Eliminar Registro")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnBuscarRegistro, "Buscar")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnSeguirBuscando, "Pulsar para Seguir Buscando o F3")
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnImprimir, "Imprimir")
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnGraficos, "Mostrar Gráficos")
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnCalculadora, "Activar la Calculadora")
        TL(14) = New ToolTip
        TL(14).SetToolTip(Me.BtnSalir, "Salir de Apuntes Contables")
        TL(15) = New ToolTip
        TL(15).SetToolTip(Me.BtnEliminaSeleccion, "Suprimir las Filas Seleccionadas de la parilla, " & vbCrLf & "No se eliminan de la Base de Datos")

        ' Llenar Grid de APUNTES al cargra el programa
        '**********************************************
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Llenar el Combo Concepto
        '*************************
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbConcepto.Items.Add(drMySql1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMySql1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbCuenta.Items.Add(drMySql1.GetValue(0))
                End While
                CmbCuenta.Text = CmbCuenta.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Cuenta")
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvApuper.Columns
            If columna.Name <> "ImporteAPP" And columna.Name <> "ImporteAPP1" And columna.Name <> "CuentaAPP" And columna.Name <> "CodigoAPP" Then
                frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
            End If
        Next
    End Sub

    Private Sub BtnFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnFiltroCuenta.Click
        BtnFiltroCuenta.Enabled = False
        BtnSinFiltroCuenta.Enabled = True
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        BtnFiltroConcepto.Enabled = False
        BtnSinFiltroConcepto.Enabled = True
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnFiltroFecha.Click
        BtnFiltroFecha.Enabled = False
        BtnSinFiltroFecha.Enabled = True
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vDate1 = DateTimePicker1.Value.Date
        vDate2 = DateTimePicker2.Value.Date
        vtipoSql += " And apuper.FechaAPP >= ?"
        vtipoSql += " And apuper.FechaAPP <= ?"
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnSinFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroCuenta.Click
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " And apuper.CuentaAPP <> '' "
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' Llenar el Combo Concepto
        '*************************
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbConcepto.Items.Add(drMySql1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " And apuper.ConceptoAPP <> '' "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnSinFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroFecha.Click
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub CmbCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCuenta.SelectedIndexChanged
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
            End If
            If BtnFiltroFecha.Enabled = False Then
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuper.FechaAPP >= ?"
                vtipoSql += " And apuper.FechaAPP <= ?"
            End If
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            vtipoGrid = "APUNTES_PERIODICOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        If BtnFiltroFecha.Enabled = False Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
            If BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
            End If
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
            End If
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            vtipoGrid = "APUNTES_PERIODICOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If BtnFiltroFecha.Enabled = False Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
            If BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
            End If
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
            End If
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            vtipoGrid = "APUNTES_PERIODICOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMySql1.Close()
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMySql1 = cmdMySql1cr.ExecuteReader()
        drMySql1.Read()
        TxtConcepto.Text = drMySql1.GetValue(1)
        drMySql1.Close()
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
            If BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
            End If
            If BtnFiltroFecha.Enabled = False Then
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuper.FechaAPP >= ?"
                vtipoSql += " And apuper.FechaAPP <= ?"
            End If
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            vtipoGrid = "APUNTES_PERIODICOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
        vTxtNombre = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarApuntesPeriodicos Is Nothing) OrElse (Not frmEditarApuntesPeriodicos.IsHandleCreated)) Then
            frmEditarApuntesPeriodicos = New EditarApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "NO"  ' Eliminar
        frmEditarApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarApuntesPeriodicos.Dispose()

        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & frmApuntesPeriodicos.CmbCuenta.Text & "' "
        End If
        If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & frmApuntesPeriodicos.CmbConcepto.Text & "' "
        End If
        If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            vDate1 = frmApuntesPeriodicos.DateTimePicker1.Value.Date
            vDate2 = frmApuntesPeriodicos.DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = 0
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoInformeApuntesPeriodicos Is Nothing) OrElse (Not frmTipoInformeApuntesPeriodicos.IsHandleCreated)) Then
            frmTipoInformeApuntesPeriodicos = New TipoInformeApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera modal.
        frmTipoInformeApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoInformeApuntesPeriodicos.Dispose()
    End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoGraficoPeriodico Is Nothing) OrElse (Not frmTipoGraficoPeriodico.IsHandleCreated)) Then
            frmTipoGraficoPeriodico = New TipoGraficoPeriodico
        End If
        ' Llamamos al formulario de manera modal.
        frmTipoGraficoPeriodico.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoGraficoPeriodico.Dispose()
    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        'Elimina las Filas Seleccionadas
        '*******************************
        For Each r As DataGridViewRow In DgvApuper.SelectedRows
            If DgvApuper.Rows.Count > 1 Then
                DgvApuper.Rows.Remove(r)
            End If
        Next
        filaSelec = DgvApuper.CurrentRow.Index
        For i = 0 To DgvApuper.Rows.Count - 1
            DgvApuper.Rows(i).Selected = False
        Next
        'Variable que guardara el valor
        'Dim iTotal As Integer = Me.DgvApuper.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
        'Definimos la variable i para controlar el ciclo for
        'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
        DgvApuntesContables(3, 4)
        DgvApuper.Select()
        DgvApuper.CurrentRow.Selected = True
        DgvApuper.Refresh()
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
        vTxtNombre = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarApuntesPeriodicos Is Nothing) OrElse (Not frmEditarApuntesPeriodicos.IsHandleCreated)) Then
            frmEditarApuntesPeriodicos = New EditarApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "SI"
        frmEditarApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarApuntesPeriodicos.Dispose()
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        DgvApuper.CurrentCell = DgvApuper.Rows(filaActual).Cells(0)
        DgvApuper.Rows(filaActual).Selected = True
    End Sub

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        vRow = 0
        For Each row As DataGridViewRow In DgvApuper.Rows
            If frmBuscar.ChkPrimerRegistro.Checked = True Then 'Desde el primer registro
                If vCampo = 0 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                ElseIf vCampo = 1 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                ElseIf vCampo = 2 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                ElseIf vCampo = 3 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                ElseIf vCampo = 4 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(5).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                End If
            Else ' desde donde está la fila seleccionada
                vRow = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
                If row.Index > vRow Then
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 1 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 2 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 3 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 4 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(5).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    End If
                End If
            End If
        Next
        If vRow = -1 Then
            MsgBox("No hay ninguna Coincidencia con los datos Introducidos")
            BtnSeguirBuscando.Enabled = False
        Else
            If vCampo = 0 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
            ElseIf vCampo = 1 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
            ElseIf vCampo = 2 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(1)
            ElseIf vCampo = 3 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(2)
            ElseIf vCampo = 4 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(5)
            End If
        End If
    End Sub

    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        SeguirF3()
    End Sub

    Private Sub ApuntesPeriodicos_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                SeguirF3()
            End If
        End If
    End Sub

    Private Sub SeguirF3()
        vCantidadFilas = DgvApuper.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox("No hay más Registros que Coincidencian con los datos Introducidos")
            BtnSeguirBuscando.Enabled = False
        Else
            vContador = -1
            For Each row As DataGridViewRow In DgvApuper.Rows
                vContador += 1
                If vContador > vRow Then
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 1 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 2 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 3 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 4 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(5).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    End If
                End If
            Next
            If vRowSeguir = -1 Then
                MsgBox("No hay más Registros que Coincidencian con los datos Introducidos")
                BtnSeguirBuscando.Enabled = False
            Else
                vRow = vRowSeguir
                If vCampo = 0 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(1)
                ElseIf vCampo = 3 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(2)
                ElseIf vCampo = 4 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(5)
                End If
                vRowSeguir = 0
            End If
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = vFilaActual - 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        frmPrincipal.TsLabelFormulario.Text = "Introducción de Apuntes Periódicos"
        ' Comprobamos si existe un identificador asociado.
        If ((frmIntroApuntesPeriodicos Is Nothing) OrElse (Not frmIntroApuntesPeriodicos.IsHandleCreated)) Then
            frmIntroApuntesPeriodicos = New IntroApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera modal.
        frmIntroApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmIntroApuntesPeriodicos.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = DgvApuper.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = vFilaActual + 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = DgvApuper.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = DgvApuper.RowCount - 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub DgvApuper_DoubleClick(sender As Object, e As EventArgs) Handles DgvApuper.DoubleClick
        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
        vTxtNombre = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarApuntesPeriodicos Is Nothing) OrElse (Not frmEditarApuntesPeriodicos.IsHandleCreated)) Then
            frmEditarApuntesPeriodicos = New EditarApuntesPeriodicos
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "SI"
        frmEditarApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarApuntesPeriodicos.Dispose()
        vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & CmbConcepto.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
        End If
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        DgvApuper.CurrentCell = DgvApuper.Rows(filaActual).Cells(0)
        DgvApuper.Rows(filaActual).Selected = True
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class