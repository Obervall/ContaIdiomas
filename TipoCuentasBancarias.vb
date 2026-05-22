Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class TipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, filaActual As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, PrintLine, Contador As Integer
    Public TL(10) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub TipoCuentasBancarias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Me.KeyPreview = True
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAñadirRegistro, "Añadir Registro")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEditarRegistro, "Editar Registro")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnEliminarRegistro, "Eliminar Registro")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnBuscarRegistro, "Buscar")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnSeguirBuscando, "Pulsar para Seguir Buscando o F3")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnImprimir, "Imprimir")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnSalir, "Salir de Apuntes Contables")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnPrimero, "Ir al Primer Registro")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnAnterior, "Ir al Anterior Registro")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnSiguiente, "Ir al Siguiente Registro")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnUltimo, "Ir al Ultimo Registro")

        ' Llenar Grid de TIPO CUENTAS BANCARIAS
        '**************************************
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvTipoCuentasBancarias.Columns
            frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
        Next
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

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        vRow = 0
        For Each row As DataGridViewRow In DgvTipoCuentasBancarias.Rows
            If frmBuscar.ChkPrimerRegistro.Checked = True Then 'Desde el primer registro
                If vCampo = 0 Then
                    If frmBuscar.ChkExacta.Checked = False Then
                        If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                            row.Selected = True
                            vRow = row.Index
                            Exit For
                        Else
                            vRow = -1
                        End If
                    Else
                        If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
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
                End If
            Else ' desde donde está la fila seleccionada
                vRow = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
                If row.Index > vRow Then
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
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
                    End If
                End If
            End If
        Next
        If vRow = -1 Then
            MsgBox("No hay ninguna Coincidencia con los datos Introducidos")
            BtnSeguirBuscando.Enabled = False
        Else
            If vCampo = 0 Then
                DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(0)
            ElseIf vCampo = 1 Then
                DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(0)
            ElseIf vCampo = 2 Then
                DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(1)
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
        vCantidadFilas = DgvTipoCuentasBancarias.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox("No hay más Registros que Coincidencian con los datos Introducidos")
            BtnSeguirBuscando.Enabled = False
        Else
            vContador = -1
            For Each row As DataGridViewRow In DgvTipoCuentasBancarias.Rows
                vContador += 1
                If vContador > vRow Then
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
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
                    End If
                End If
            Next
            If vRowSeguir = -1 Then
                MsgBox("No hay más Registros que Coincidencian con los datos Introducidos")
                BtnSeguirBuscando.Enabled = False
            Else
                vRow = vRowSeguir
                If vCampo = 0 Then
                    DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                    DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                    DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvTipoCuentasBancarias.Rows(vRow).Selected = True
                    DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vRow).Cells(1)
                    vRowSeguir = 0
                End If
            End If
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click

        vtipoSql = "SELECT * FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"

        LlenarGrid(vtipoSql, "PRINT_TIPO_CUENTAS", 1)
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
        frmImprimirForm.LblTitulo.Text = "Listado Tipo de Cuentas Bancarias"

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString("Tipo Cuenta:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString("Descripción:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'
        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                'Esta parte se activa solo si 'startY' que es la posicion vertical almacenada supera el borde inferior de la pagina
                'Este se reinicia con cada pagina necesitada
                e.HasMorePages = True
                Exit Do
            End If
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
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
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
            'e.Graphics.DrawString(vValor, FuenteNegrita, Brushes.Black, frmImprimirForm.Punto4.Left, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.Punto4.Left - 50, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.lbCursadas.Text, FuenteDetalles, Brushes.Black, ImprimirForm.lbCursadas.Left, startY + 15)
            'e.Graphics.DrawString(frmImprimirForm.lbPromedio.Text, FuenteDetalles, Brushes.Black, ImprimirForm.lbPromedio.Left, startY + 30)
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        vTxtNombre = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarTipoCuentaBancaria Is Nothing) OrElse (Not frmEditarTipoCuentaBancaria.IsHandleCreated)) Then
            frmEditarTipoCuentaBancaria = New EditarTipoCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        vEditar = "SI"
        frmEditarTipoCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarTipoCuentaBancaria.Dispose()
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(filaActual).Cells(0)
        DgvTipoCuentasBancarias.Rows(filaActual).Selected = True
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        MsgBox("No se pueden Eliminar Tipos de Cuentas Bancarias")

        ' DEJO EL TEXTO POR SI SE TIENE QUE USAR ALGUNA VEZ, FALTARIA EL VOLVER A LLENAR EL GRID ACTUALIZADO
        '***************************************************************************************************
        'filaActual = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        'vTxtNombre = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString

        'respuesta = MsgBox("¿Estas seguro de Eliminar el Tipo Cuenta Bancaria '" & vTxtNombre & "' ", vbQuestion & vbYesNo & vbDefaultButton2, "Eliminar Tipo Cuenta Bancaria")
        'If respuesta = vbYes Then
        '    ' Eliminar Registro Tipo Cuentas
        '    vtipoSql = "DELETE FROM tipocuentas"
        '    vtipoSql += " WHERE tipocuentas.CodigoTIP = '" & vTxtNombre & "' "
        '    cmdMdb1cr.CommandText = vtipoSql

        '    Try
        '        cmdMdb1cr.ExecuteNonQuery()
        '        MsgBox("Registro Tipo Cuenta Bancaria, Borrada !!!")
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        'End If

    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmNuevoTipoCuentaBancaria Is Nothing) OrElse (Not frmNuevoTipoCuentaBancaria.IsHandleCreated)) Then
            frmNuevoTipoCuentaBancaria = New NuevoTipoCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        frmNuevoTipoCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmNuevoTipoCuentaBancaria.Dispose()
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = 0
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = vFilaActual - 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = DgvTipoCuentasBancarias.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = vFilaActual + 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = DgvTipoCuentasBancarias.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = DgvTipoCuentasBancarias.RowCount - 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub
End Class