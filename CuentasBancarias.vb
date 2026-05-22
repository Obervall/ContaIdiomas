Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class CuentasBancarias

    Public vtipoSql, vtipoGrid, vTxtNombre, filaActual As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, PrintLine, Contador, filaSelec As Integer
    Public TL(13) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub CuentasBancarias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Me.KeyPreview = True
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnFiltroTipoCuenta, resManager.GetString("ToolTipAplicarFiltro"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSinFiltroTipoCuenta, resManager.GetString("ToolTipQuitarFiltro"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAñadirRegistro, resManager.GetString("ToolTipAñadir"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnEditarRegistro, resManager.GetString("ToolTipEditar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnEliminarRegistro, resManager.GetString("ToolTipEliminar"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnEliminaSeleccion, resManager.GetString("ToolTipEliminaSeleccion"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))

        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox4.MouseMove, AddressOf VerificarFiltrosDesactivados

        ' Llenar el Combo Tipo Cuenta desde Access (Modelo Híbrido)
        '********************************************************
        cmdMdb1cr.CommandText = "SELECT tipocuentas.CodigoTIP FROM tipocuentas"
        cmdMdb1cr.CommandText += " ORDER BY tipocuentas.CodigoTIP ASC"

        Try
            ' 1. Guardamos la posición seleccionada actual para que no salte el combo
            Dim indiceSeleccionado As Integer = CmbTipoCuenta.SelectedIndex

            ' 2. Limpiamos los ítems viejos para evitar duplicados
            CmbTipoCuenta.Items.Clear()

            ' 3. Ejecutamos el lector de Access
            drMdb1 = cmdMdb1cr.ExecuteReader()

            If drMdb1.HasRows Then
                While drMdb1.Read()
                    ' Obtenemos el texto guardado en Access (quitando espacios en blanco extras)
                    Dim valorBD As String = drMdb1.GetValue(0).ToString().Trim()

                    ' Buscamos si tú programaste una traducción para este valor en el .resx
                    Dim textoTraducido As String = rmse.GetString(valorBD)

                    ' PROTECCIÓN: Si no hay traducción (porque lo creó el usuario), 
                    ' se usa el valor original de la base de datos de Access
                    If String.IsNullOrEmpty(textoTraducido) Then
                        textoTraducido = valorBD
                    End If

                    ' Añadimos el ítem definitivo al ComboBox
                    CmbTipoCuenta.Items.Add(textoTraducido)
                End While

                ' 4. Restauramos la selección del usuario o dejamos el primer ítem por defecto
                If indiceSeleccionado >= 0 AndAlso indiceSeleccionado < CmbTipoCuenta.Items.Count Then
                    CmbTipoCuenta.SelectedIndex = indiceSeleccionado
                ElseIf CmbTipoCuenta.Items.Count > 0 Then
                    CmbTipoCuenta.SelectedIndex = 0
                End If
            Else
                MsgBox(resManager.GetString("NoExistenRegistros") & " " & cmdMdb1cr.CommandText)
            End If

            drMdb1.Close()

        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorAlLlenarComboTipoCuenta"))
            MsgBox(ex.ToString)
        Finally
            ' Bloque de seguridad: cerramos el reader por si ocurre algún fallo en el bucle
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
                drMdb1.Close()
            End If
        End Try

        CmbTipoCuenta.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoCuenta.SelectedIndex = 0

        ' Llenar Grid de CUENTAS
        '***********************
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvCuentas.Columns
            If columna.Name <> "NotasCUE1" And columna.Name <> "Expr1003" Then
                frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
            End If
        Next
    End Sub

    Private Sub BtnFiltroTipoCuenta_Click(sender As Object, e As EventArgs) Handles BtnFiltroTipoCuenta.Click
        BtnFiltroTipoCuenta.Enabled = False
        BtnSinFiltroTipoCuenta.Enabled = True
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " WHERE "
        vtipoSql += "cuentas.TipoCUE = '" & CmbTipoCuenta.Text & "' "
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnSinFiltroTipoCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroTipoCuenta.Click
        BtnFiltroTipoCuenta.Enabled = True
        BtnSinFiltroTipoCuenta.Enabled = False
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " WHERE "
        vtipoSql += "cuentas.TipoCUE <> '' "
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        vRow = 0
        For Each row As DataGridViewRow In DgvCuentas.Rows
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
                        If CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    End If
                End If
            Else ' desde donde está la fila seleccionada
                vRow = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
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
                            If CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
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
                DgvCuentas.Rows(vRow).Selected = True
                DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(0)
            ElseIf vCampo = 1 Then
                DgvCuentas.Rows(vRow).Selected = True
                DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(0)
            ElseIf vCampo = 2 Then
                DgvCuentas.Rows(vRow).Selected = True
                DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(1)
            ElseIf vCampo = 3 Then
                DgvCuentas.Rows(vRow).Selected = True
                DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(2)
            ElseIf vCampo = 4 Then
                DgvCuentas.Rows(vRow).Selected = True
                DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(4)
            End If
        End If
    End Sub

    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        SeguirF3()
    End Sub

    Private Sub ApuntesContables_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                SeguirF3()
            End If
        End If
    End Sub

    Private Sub SeguirF3()
        vCantidadFilas = DgvCuentas.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox("No hay más Registros que Coincidencian con los datos Introducidos")
            BtnSeguirBuscando.Enabled = False
        Else
            vContador = -1
            For Each row As DataGridViewRow In DgvCuentas.Rows
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
                            If CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
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
                    DgvCuentas.Rows(vRow).Selected = True
                    DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvCuentas.Rows(vRow).Selected = True
                    DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvCuentas.Rows(vRow).Selected = True
                    DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(1)
                ElseIf vCampo = 3 Then
                    DgvCuentas.Rows(vRow).Selected = True
                    DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(2)
                ElseIf vCampo = 4 Then
                    DgvCuentas.Rows(vRow).Selected = True
                    DgvCuentas.CurrentCell = DgvCuentas.Rows(vRow).Cells(4)
                End If
                vRowSeguir = 0
            End If
        End If
    End Sub

    Private Sub CmbTipoCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTipoCuenta.SelectedIndexChanged
        If BtnFiltroTipoCuenta.Enabled = False Then
            vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
            vtipoSql += " WHERE "
            vtipoSql += "cuentas.TipoCUE = '" & CmbTipoCuenta.Text & "' "
            vtipoSql += " ORDER BY  cuentas.NombreCUE ASC"
            vtipoGrid = "CUENTAS_BANCARIAS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmNuevaCuentaBancaria Is Nothing) OrElse (Not frmNuevaCuentaBancaria.IsHandleCreated)) Then
            frmNuevaCuentaBancaria = New NuevaCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        frmNuevaCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmNuevaCuentaBancaria.Dispose()
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        If BtnFiltroTipoCuenta.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "cuentas.TipoCUE = '" & CmbTipoCuenta.Text & "' "
        End If
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarCuentaBancaria Is Nothing) OrElse (Not frmEditarCuentaBancaria.IsHandleCreated)) Then
            frmEditarCuentaBancaria = New EditarCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "NO"  ' Eliminar
        frmEditarCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarCuentaBancaria.Dispose()
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        If BtnFiltroTipoCuenta.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "cuentas.TipoCUE = '" & CmbTipoCuenta.Text & "' "
        End If
        vtipoSql += " ORDER BY  cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        'Elimina las Filas Seleccionadas
        '*******************************
        For Each r As DataGridViewRow In DgvCuentas.SelectedRows
            If DgvCuentas.Rows.Count > 1 Then
                DgvCuentas.Rows.Remove(r)
            End If
        Next
        filaSelec = DgvCuentas.CurrentRow.Index
        For i = 0 To DgvCuentas.Rows.Count - 1
            DgvCuentas.Rows(i).Selected = False
        Next
        'Variable que guardara el valor
        'Dim iTotal As Integer = Me.DgvCuentas.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
        'Definimos la variable i para controlar el ciclo for
        'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
        DgvCuentasBancarias()
        DgvCuentas.Select()
        DgvCuentas.CurrentRow.Selected = True
        DgvCuentas.Refresh()
    End Sub

    Private Sub DgvCuentas_KeyDown(sender As Object, e As KeyEventArgs) Handles DgvCuentas.KeyDown
        'MsgBox(e.KeyCode)
        'Elimina las Filas Seleccionadas
        '*******************************
        If e.KeyCode = 46 Then  'Tecla Supr
            For Each r As DataGridViewRow In DgvCuentas.SelectedRows
                If DgvCuentas.Rows.Count > 1 Then
                    DgvCuentas.Rows.Remove(r)
                End If
            Next
            filaSelec = DgvCuentas.CurrentRow.Index
            For i = 0 To DgvCuentas.Rows.Count - 1
                DgvCuentas.Rows(i).Selected = False
            Next
            'Variable que guardara el valor
            'Dim iTotal As Integer = Me.DgvCuentas.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
            'Definimos la variable i para controlar el ciclo for
            'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
            DgvCuentasBancarias()
            DgvCuentas.Select()
            DgvCuentas.CurrentRow.Selected = True
            DgvCuentas.Refresh()
        End If
        If e.KeyCode = 117 Then 'Tecla F6
            'Vuelve a Refrecar el DataGrid y dejar los Btn de los Filtros sin Filtrar
            '************************************************************************
            vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
            vtipoSql += " ORDER BY  cuentas.NombreCUE ASC"
            vtipoGrid = "CUENTAS_BANCARIAS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            BtnFiltroTipoCuenta.Enabled = True
            BtnSinFiltroTipoCuenta.Enabled = False
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click

        'Llenamos la tabla de ImprimirForm con los cálculos realizados
        '*************************************************************
        vValor = 0
        frmImprimirForm.LblTotal.Text = "Total: 0,00 " & vMoneda
        For Each fila As DataGridViewRow In frmCuentasBancarias.DgvCuentas.Rows
            vValor += fila.Cells(3).Value
            frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
        Next

        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString

        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        'Para ver la plantilla de impresión
        'frmImprimirForm.Show()

        If My.Settings.Previsualizar = True Then
            'Te deja ver un preview del reporte antes de imprimir
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        If My.Settings.ElegirImpresora = True Then
            'Te deja elegir la impresora
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        End If

        If My.Settings.DirectoImpresora = True Then
            'Imprime en la impresora por defecto
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Cualquier variable que desees que conserve su valor debes declararla fuera del Printdocument
        'Todas las variable declaradas dentro de printdocument pierden su valor al cambiar de pagina
        'Definimos los tipos de letras a utilizar en el reporte
        '******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}

        If BtnSinFiltroTipoCuenta.Enabled = True Then
            frmImprimirForm.LblTitulo.Text = "Listado de Cuentas Bancarias. Filtrado por: " & CmbTipoCuenta.Text
        Else
            frmImprimirForm.LblTitulo.Text = "Listado de Cuentas Bancarias"
        End If

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString("Tipo:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString("Nombre Cuenta:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString("Número Cuenta:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        e.Graphics.DrawString("Saldo(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'
        Do While PrintLine < frmCuentasBancarias.DgvCuentas.Rows.Count
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                'Esta parte se activa solo si 'startY' que es la posicion vertical almacenada supera el borde inferior de la pagina
                'Este se reinicia con cada pagina necesitada
                e.HasMorePages = True
                Exit Do
            End If
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(0).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(2).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(3).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY, sf)
            startY += frmImprimirForm.LblFecha.Height
            e.Graphics.DrawString("Notas:  ", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)  ' frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(4).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 50, startY)
            'Aqui estoy usando un tipo de letras mas grande
            'LabelCodigo' mas grande que 'Punto1' para crear mas espacio entre filas

            'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
            'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
            ''Imprimimos los valores que salen despues del datagridview al final del reporte
            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop
        'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
        'Imprimimos los valores que salen despues del datagridview al final del reporte
        If Contador >= frmCuentasBancarias.DgvCuentas.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY + 15, sf)
            'e.Graphics.DrawString(vValor, FuenteNegrita, Brushes.Black, frmImprimirForm.Punto4.Left, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.Punto4.Left - 50, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.lbCursadas.Text, FuenteDetalles, Brushes.Black, ImprimirForm.lbCursadas.Left, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.lbPromedio.Text, FuenteDetalles, Brushes.Black, ImprimirForm.lbPromedio.Left, startY + 30)

            'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
            PrintLine = 0
            Contador = 0
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

        'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
        If Contador = 0 Then
            frmImprimirForm.LblNumeroPagina.Text = "0"
        End If
    End Sub


    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
        vTxtNombre = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarCuentaBancaria Is Nothing) OrElse (Not frmEditarCuentaBancaria.IsHandleCreated)) Then
            frmEditarCuentaBancaria = New EditarCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "SI"
        frmEditarCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarCuentaBancaria.Dispose()
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        If BtnFiltroTipoCuenta.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "cuentas.TipoCUE = '" & CmbTipoCuenta.Text & "' "
        End If
        vtipoSql += " ORDER BY  cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        DgvCuentas.CurrentCell = DgvCuentas.Rows(filaActual).Cells(0)
        DgvCuentas.Rows(filaActual).Selected = True
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
        BtnFiltroTipoCuenta.Enabled = True
        BtnSinFiltroTipoCuenta.Enabled = False
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        BtnFiltroTipoCuenta.Enabled = True
        BtnSinFiltroTipoCuenta.Enabled = False
        Me.Close()
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = 0
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = vFilaActual - 1
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = DgvCuentas.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = vFilaActual + 1
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = DgvCuentas.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = DgvCuentas.RowCount - 1
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnSinFiltroTipoCuenta, TL(1)},
            {Me.BtnSeguirBuscando, TL(6)}
        }

        For Each par In botonesBloqueados
            Dim boton As Button = par.Key
            Dim tool As ToolTip = par.Value

            If Not boton.Enabled Then
                ' Traducimos la posición del ratón al contenedor nativo del botón (su GroupBox)
                Dim posRatonRelativaAlBoton As Point = boton.Parent.PointToClient(Cursor.Position)

                ' Si el ratón está sobre el botón desactivado
                If boton.Bounds.Contains(posRatonRelativaAlBoton) Then
                    ' Calculamos la posición respecto al formulario para dibujar el cartelito en el lugar correcto
                    Dim posRatonRelativaAlForm As Point = Me.PointToClient(Cursor.Position)
                    'tool.Show(resManager.GetString("ToolTipQuitarFiltro"), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso
                    Dim textoKey As String = If(boton Is Me.BtnSeguirBuscando, "ToolTipSeguirBuscando", "ToolTipQuitarFiltro")
                    tool.Show(resManager.GetString(textoKey), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón no está sobre ningún botón bloqueado, ocultamos los tres
        TL(1).Hide(Me)
        TL(6).Hide(Me)
    End Sub

End Class