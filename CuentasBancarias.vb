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

        CargarComboTipoCuentaGlobal(Me.CmbTipoCuenta, rmse)

        CmbTipoCuenta.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoCuenta.SelectedIndex = 0

        ' Llenar Grid de CUENTAS
        '***********************
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirColumnasGridCuentas(DgvCuentas, rmse)

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

        ' 1. Por defecto, asumimos que el valor a buscar es el texto visible del combo
        Dim valorOriginalBD As String = CmbTipoCuenta.Text

        ' 2. ¡Truco de inversión! Buscamos en el historial de recursos cuál era el código original de fábrica
        ' Recorremos todas las claves del archivo de recursos para ver cuál coincide con la traducción actual
        Dim recursos As System.Resources.ResourceSet = rmse.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)
        If recursos IsNot Nothing Then
            For Each elemento As System.Collections.DictionaryEntry In recursos
                ' Si el valor traducido en el .resx coincide con lo que el usuario ve en el combo...
                If elemento.Value.ToString() = CmbTipoCuenta.Text Then
                    ' ¡Encontramos el código original de la base de datos! (ej: "CAJA EFECTIVO")
                    valorOriginalBD = elemento.Key.ToString()
                    Exit For
                End If
            Next
        End If

        ' 3. Armamos la consulta SQL usando el valor original que Access sí entiende
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " WHERE "
        vtipoSql += "cuentas.TipoCUE = '" & valorOriginalBD & "' "
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"
        ' 4. Ejecutamos tu función para llenar el Grid
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirColumnasGridCuentas(DgvCuentas, rmse)
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
        TraducirColumnasGridCuentas(DgvCuentas, rmse)
        ' Si quieres que al quitar el filtro el combo vuelva a su posición inicial o se limpie:
        If CmbTipoCuenta.Items.Count > 0 Then
            CmbTipoCuenta.SelectedIndex = 0 ' O pon -1 si prefieres que se quede en blanco
        End If
    End Sub

    ' BOTÓN BUSCAR: Abre la ventana para configurar los parámetros y busca desde el principio
    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        ' Llamamos al motor pasándole True para que respete si el usuario marcó "Desde el primer registro"
        EjecutarBusqueda(forzarDesdeInicio:=True)
    End Sub

    ' BOTÓN SEGUIR BUSCANDO: No abre ventana, busca directamente la siguiente coincidencia
    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        ' Llamamos al motor pasándole False para obligarle a saltar a la siguiente fila desde donde esté parado
        EjecutarBusqueda(forzarDesdeInicio:=False)
    End Sub

    Private Sub ApuntesContables_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                EjecutarBusqueda(forzarDesdeInicio:=False)
            End If
        End If
    End Sub

    Private Sub EjecutarBusqueda(ByVal forzarDesdeInicio As Boolean)
        vBuscar = frmBuscar.CmbTextoBuscar.Text.ToLower()
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        Dim buscarExacto As Boolean = frmBuscar.ChkExacta.Checked
        Dim desdePrimerRegistro As Boolean = frmBuscar.ChkPrimerRegistro.Checked

        ' CORRECCIÓN AQUÍ: Determinamos el punto de inicio real combinando los estados
        Dim filaInicio As Integer = 0

        ' Si se pulsa "Seguir Buscando" O si el usuario NO quiere empezar desde el primer registro
        If (Not forzarDesdeInicio OrElse Not desdePrimerRegistro) AndAlso DgvCuentas.CurrentRow IsNot Nothing Then
            filaInicio = DgvCuentas.CurrentRow.Index + 1
        End If

        ' Si al avanzar nos salimos del total de filas, avisamos y salimos
        If filaInicio >= DgvCuentas.Rows.Count Then
            MsgBox(resManager.GetString("FinalTabla"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            Exit Sub
        End If

        ' Mapeamos los índices de las celdas
        Dim columnasEvaluar As Integer()
        Select Case vCampo
            Case 0 : columnasEvaluar = {0, 1, 2, 4}
            Case 1 : columnasEvaluar = {0}
            Case 2 : columnasEvaluar = {1}
            Case 3 : columnasEvaluar = {2}
            Case 4 : columnasEvaluar = {4}
            Case Else : columnasEvaluar = {}
        End Select

        vRow = -1

        ' Busqueda en el DataGridView
        For i As Integer = filaInicio To DgvCuentas.Rows.Count - 1
            Dim row As DataGridViewRow = DgvCuentas.Rows(i)

            ' Evitamos evaluar la fila nueva vacía que añade DataGridView al final de forma automática
            If row.IsNewRow Then Continue For

            Dim coincide As Boolean = False

            For Each colIdx As Integer In columnasEvaluar
                ' Protección extra por si alguna celda está vacía (Nothing)
                Dim valorCelda As String = ""
                If row.Cells(colIdx).Value IsNot Nothing Then
                    valorCelda = row.Cells(colIdx).Value.ToString().ToLower()
                End If

                If buscarExacto Then
                    coincide = (valorCelda = vBuscar)
                Else
                    coincide = valorCelda.Contains(vBuscar)
                End If

                If coincide Then Exit For
            Next

            If coincide Then
                DgvCuentas.ClearSelection()
                row.Selected = True

                ' Intentamos enfocar la celda para que la cuadrícula se mueva (scrolleé) automáticamente
                Try
                    DgvCuentas.CurrentCell = row.Cells(columnasEvaluar(0))
                Catch
                    DgvCuentas.CurrentCell = row.Cells(0)
                End Try

                vRow = row.Index
                Exit For
            End If
        Next

        ' Avisar al usuario si no encontró nada en todo el recorrido
        If vRow = -1 Then
            MsgBox(resManager.GetString("MsgDatos2"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
        End If
    End Sub

    Private Sub CmbTipoCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTipoCuenta.SelectedIndexChanged
        ' Si el botón de filtro está deshabilitado (significa que el filtro está actualmente activo)
        If BtnFiltroTipoCuenta.Enabled = False Then
            ' Refrescamos el grid automáticamente con el nuevo tipo seleccionado
            EjecutarFiltroTipoCuenta()
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
        TraducirColumnasGridCuentas(DgvCuentas, rmse)
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarCuentaBancaria Is Nothing) OrElse (Not frmEditarCuentaBancaria.IsHandleCreated)) Then
            frmEditarCuentaBancaria = New EditarCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        ' Llamamos al formulario de manera modal.
        If vEditar = "NO" Then
            vEditar = "NO"  ' Eliminar
        Else
            vEditar = "NO"
        End If
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
        TraducirColumnasGridCuentas(DgvCuentas, rmse)
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
            TraducirColumnasGridCuentas(DgvCuentas, rmse)

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
            ' Saltamos la fila vacía automática si existiera al final
            If fila.IsNewRow Then Continue For

            ' Protección si el valor de la celda de saldo está vacío
            If fila.Cells(3).Value IsNot Nothing AndAlso IsNumeric(fila.Cells(3).Value) Then
                vValor += Convert.ToDouble(fila.Cells(3).Value)
            End If
            frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
        Next

        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString

        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        If My.Settings.Previsualizar = True Then
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        If My.Settings.ElegirImpresora = True Then
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        End If

        If My.Settings.DirectoImpresora = True Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}

        If BtnSinFiltroTipoCuenta.Enabled = True Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoCuentasBancariasFiltrado") & " " & CmbTipoCuenta.Text
        Else
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoCuentasBancarias")
        End If

        'Imprimimos el encabezado
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        'Imprimimos los títulos de columnas
        e.Graphics.DrawString(resManager.GetString("Tipo") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Nombre") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Numero") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Saldo") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)

        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        Dim startX As Integer = frmImprimirForm.Punto1.Left
        Dim startY As Integer = frmImprimirForm.Punto1.Top

        Do While PrintLine < frmCuentasBancarias.DgvCuentas.Rows.Count
            ' Evitamos procesar la fila vacía automática si existe al final
            If frmCuentasBancarias.DgvCuentas.Rows(PrintLine).IsNewRow Then
                PrintLine += 1
                Contador += 1
                Continue Do
            End If

            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' ✨ EL ÚNICO TRUCO: Leemos el nombre de la celda (1) y buscamos su traducción Nom_
            Dim nombreCelda As String = frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(1).Value.ToString().Trim()
            Dim tradNombre As String = resManager.GetString("Nom_" & nombreCelda.Replace(" ", "_"))
            Dim nombreFinal As String = If(Not String.IsNullOrEmpty(tradNombre), tradNombre, nombreCelda)

            ' Imprimimos los datos de las columnas principales (Usamos 'nombreFinal' en la celda 1)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(0).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(nombreFinal, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(2).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
            e.Graphics.DrawString(frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(3).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY, sf)

            ' Avanzamos la coordenada vertical para pintar la fila de Notas
            startY += frmImprimirForm.LblFecha.Height

            ' 1. Obtener y limpiar el texto de Notas
            Dim textoNotas As String = ""
            If frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(4).Value IsNot Nothing Then
                textoNotas = frmCuentasBancarias.DgvCuentas.Rows(PrintLine).Cells(4).Value.ToString()
            End If
            textoNotas = textoNotas.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ")

            While textoNotas.Contains("  ")
                textoNotas = textoNotas.Replace("  ", " ")
            End While

            ' 2. Ajuste estricto a un máximo de 100 caracteres
            If textoNotas.Length > 100 Then
                textoNotas = textoNotas.Substring(0, 95) & "..."
            End If

            ' 3. Cálculo dinámico de coordenadas para evitar solapamientos
            Dim etiquetaNotas As String = resManager.GetString("Notas") & ": "
            Dim tamañoEtiqueta As SizeF = e.Graphics.MeasureString(etiquetaNotas, FuenteSubrayada)
            Dim posicionXTexto As Integer = frmImprimirForm.Punto1.Left + CInt(tamañoEtiqueta.Width) + 10

            ' 4. Imprimir la fila de Notas
            e.Graphics.DrawString(etiquetaNotas, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(textoNotas, FuenteDetalles, Brushes.Black, posicionXTexto, startY)

            ' Avanzamos espacio vertical para la siguiente iteración de fila
            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop

        If Contador >= frmCuentasBancarias.DgvCuentas.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY + 15, sf)
        End If

        ' Contador de páginas
        frmImprimirForm.LblNumeroPagina.Text = (CInt(frmImprimirForm.LblNumeroPagina.Text) + 1).ToString()
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub


    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
        vTxtNombre = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarCuentaBancaria Is Nothing) OrElse (Not frmEditarCuentaBancaria.IsHandleCreated)) Then
            frmEditarCuentaBancaria = New EditarCuentaBancaria
        End If
        ' Llamamos al formulario de manera modal.
        ' Llamamos al formulario de manera modal.
        If vEditar = "NO" Then
            vEditar = "SI"  ' EDITAR
        Else
            vEditar = "SI"
        End If
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
        TraducirColumnasGridCuentas(DgvCuentas, rmse)

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
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = 0
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = vFilaActual - 1
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = DgvCuentas.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = vFilaActual + 1
            DgvCuentas.Rows(vFila).Selected = True
            DgvCuentas.CurrentCell = DgvCuentas.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvCuentas.CurrentRow.Index
        If vFilaActual = DgvCuentas.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
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

    Private Sub EjecutarFiltroTipoCuenta()
        ' 1. Evitamos ejecutar si el combo está vacío o no hay selección válida
        If CmbTipoCuenta.SelectedIndex = -1 Then Exit Sub

        ' 2. Buscamos el valor original en Access usando la traducción inversa
        Dim valorOriginalBD As String = CmbTipoCuenta.Text
        Dim recursos As System.Resources.ResourceSet = rmse.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)

        If recursos IsNot Nothing Then
            For Each elemento As System.Collections.DictionaryEntry In recursos
                If elemento.Value.ToString() = CmbTipoCuenta.Text Then
                    valorOriginalBD = elemento.Key.ToString()
                    Exit For
                End If
            Next
        End If

        ' 3. Armamos la consulta SQL con el valor original
        vtipoSql = "SELECT cuentas.TipoCUE, cuentas.NombreCUE, cuentas.NumeroCUE, cuentas.NotasCUE, cuentas.NotasCUE FROM cuentas"
        vtipoSql += " WHERE "
        vtipoSql += "cuentas.TipoCUE = '" & valorOriginalBD & "' "
        vtipoSql += " ORDER BY cuentas.NombreCUE ASC"
        vtipoGrid = "CUENTAS_BANCARIAS"

        ' 4. Llenamos el Grid y ejecutamos tu función universal de traducción
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirColumnasGridCuentas(DgvCuentas, rmse)
    End Sub
End Class