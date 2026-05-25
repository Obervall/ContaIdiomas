Imports System.Collections.Generic
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
        TL(0).SetToolTip(Me.BtnAñadirRegistro, resManager.GetString("ToolTipAñadir"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEditarRegistro, resManager.GetString("ToolTipEditar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnEliminarRegistro, resManager.GetString("ToolTipEliminar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))

        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox2.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados

        ' Llenar Grid de TIPO CUENTAS BANCARIAS
        '**************************************
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirContenidoGridTiposCuenta(DgvTipoCuentasBancarias, rmse)

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

    ''' <summary>
    ''' Función auxiliar que centraliza las reglas lógicas de comparación de texto.
    ''' </summary>
    Private Function EvaluarCoincidencia(celda0 As String, celda1 As String, buscar As String, campo As Integer, exacta As Boolean) As Boolean
        Select Case campo
            Case 0 ' Todos los campos (Celda 0 o Celda 1)
                If exacta Then
                    Return celda0 = buscar OrElse celda1 = buscar
                Else
                    Return celda0.Contains(buscar) OrElse celda1.Contains(buscar)
                End If

            Case 1 ' Solo Nombre / Código (Celda 0)
                If exacta Then Return celda0 = buscar Else Return celda0.Contains(buscar)

            Case 2 ' Solo Descripción (Celda 1)
                If exacta Then Return celda1 = buscar Else Return celda1.Contains(buscar)

            Case Else
                Return False
        End Select
    End Function

    ' BOTÓN BUSCAR: Abre la ventana de parámetros y busca desde el principio si el Check lo pide
    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        ' Llamamos al motor pasándole True para que respete el estado inicial del formulario de búsqueda
        EjecutarBusquedaTipos(forzarDesdeInicio:=True)
    End Sub

    ' BOTÓN SEGUIR BUSCANDO: No abre ventana, busca directamente la siguiente coincidencia
    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        ' Llamamos al motor pasándole False para obligarle a saltar a la siguiente fila
        EjecutarBusquedaTipos(forzarDesdeInicio:=False)
    End Sub

    ' DETECTAR TECLA F3 EN EL FORMULARIO
    Private Sub frmTipoCuentaBancaria_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                EjecutarBusquedaTipos(forzarDesdeInicio:=False)
            End If
        End If
    End Sub

    Private Sub EjecutarBusquedaTipos(ByVal forzarDesdeInicio As Boolean)
        ' Protegemos el código capturando el texto de búsqueda de forma segura
        vBuscar = If(frmBuscar.CmbTextoBuscar.Text, "").ToLower().Trim()
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        Dim buscarExacto As Boolean = frmBuscar.ChkExacta.Checked
        Dim desdePrimerRegistro As Boolean = frmBuscar.ChkPrimerRegistro.Checked

        ' Si no hay nada que buscar, cancelamos
        If String.IsNullOrEmpty(vBuscar) Then Exit Sub

        ' Determinamos el punto de inicio real en la tabla
        Dim filaInicio As Integer = 0

        ' Si se pulsa "Seguir Buscando" O si el usuario NO marcó empezar desde el primer registro
        If (Not forzarDesdeInicio OrElse Not desdePrimerRegistro) AndAlso DgvTipoCuentasBancarias.CurrentRow IsNot Nothing Then
            filaInicio = DgvTipoCuentasBancarias.CurrentRow.Index + 1
        End If

        ' Si al intentar avanzar nos salimos del límite del Grid, avisamos y salimos
        If filaInicio >= DgvTipoCuentasBancarias.Rows.Count Then
            MsgBox(resManager.GetString("MsgDatos2"), MsgBoxStyle.Information, Me.Text)
            BtnSeguirBuscando.Enabled = False
            Exit Sub
        End If

        ' Mapeamos los índices de las columnas para Tipo de Cuentas (0 = Nombre/Código, 1 = Descripción)
        Dim columnasEvaluar As Integer()
        Select Case vCampo
            Case 0 : columnasEvaluar = {0, 1} ' Ambos campos
            Case 1 : columnasEvaluar = {0}    ' Solo el Nombre/Código del Tipo
            Case 2 : columnasEvaluar = {1}    ' Solo la Descripción del Tipo
            Case Else : columnasEvaluar = {}
        End Select

        vRow = -1

        ' Bucle de búsqueda en el DataGridView de tipos de cuentas
        For i As Integer = filaInicio To DgvTipoCuentasBancarias.Rows.Count - 1
            Dim row As DataGridViewRow = DgvTipoCuentasBancarias.Rows(i)

            ' Evitamos evaluar la fila nueva vacía automática del final
            If row.IsNewRow Then Continue For

            Dim coincide As Boolean = False

            For Each colIdx As Integer In columnasEvaluar
                ' Protección contra celdas vacías (Nothing)
                Dim valorCelda As String = ""
                If row.Cells(colIdx).Value IsNot Nothing Then
                    valorCelda = row.Cells(colIdx).Value.ToString().ToLower().Trim()
                End If

                ' Evaluamos según el checkbox de coincidencia exacta
                If buscarExacto Then
                    coincide = (valorCelda = vBuscar)
                Else
                    coincide = valorCelda.Contains(vBuscar)
                End If

                If coincide Then Exit For
            Next

            ' Si encontramos la coincidencia, movemos el foco visual de la pantalla
            If coincide Then
                DgvTipoCuentasBancarias.ClearSelection()
                row.Selected = True

                Try
                    ' Selecciona la celda del campo por el que buscó para forzar el scroll visual
                    DgvTipoCuentasBancarias.CurrentCell = row.Cells(columnasEvaluar(0))
                Catch
                    DgvTipoCuentasBancarias.CurrentCell = row.Cells(0)
                End Try

                ' Sincronizamos el scroll automático
                DgvTipoCuentasBancarias.FirstDisplayedScrollingRowIndex = row.Index

                vRow = row.Index
                Exit For
            End If
        Next

        ' Avisar al usuario si no encontró absolutamente nada en todo el recorrido
        If vRow = -1 Then
            MsgBox(resManager.GetString("MsgDatos1"), MsgBoxStyle.Information, Me.Text)
            BtnSeguirBuscando.Enabled = False
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        vtipoSql = "SELECT * FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"

        LlenarGrid(vtipoSql, "PRINT_TIPO_CUENTAS", 1)
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString

        'Para ver la plantilla de impresión
        'frmImprimirForm.Show()

        ' 1. Reinicia las variables globales antes de empezar a imprimir
        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        ' 2. Lanza el proceso de impresión (esto activa automáticamente el evento PrintPage)
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
        frmImprimirForm.LblTitulo.Text = rmse.GetString("TituloReporte")

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        ' Encabezado Columna 0: Tomamos el texto y lo recortamos si supera los 30 caracteres
        Dim textoEncabezado0 As String = resManager.GetString("Tipo") & ":"
        If textoEncabezado0.Length > 30 Then textoEncabezado0 = textoEncabezado0.Substring(0, 30)
        e.Graphics.DrawString(textoEncabezado0, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)

        ' Encabezado Columna 1: Se queda igual en su posición fija
        e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'

        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' --- COLUMNA 0 (Tipo de cuenta - Máx 30 caracteres) ---
            Dim valorBD0 As String = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString()
            Dim textoCelda0 As String = TraducirDinamico(valorBD0, False) ' Busca "Cuenta_Corriente" o devuelve el texto íntegro

            If textoCelda0.Length > 30 Then
                textoCelda0 = textoCelda0.Substring(0, 30)
            End If
            e.Graphics.DrawString(textoCelda0, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)


            ' --- COLUMNA 1 (Descripción - Ajuste con ... antes del borde) ---
            Dim valorBD1 As String = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString()

            ' 1. Tomamos el TIPO (Celda 0) para armar la llave de la descripción (ej: "Cuenta_Corriente")
            Dim tipoParaLlave As String = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString()

            ' 2. Buscamos en ResX Manager la combinación "Desc_" & "Cuenta_Corriente"
            Dim textoCelda1 As String = TraducirDinamico(tipoParaLlave, True)

            ' 3. Si la función nos devuelve el mismo nombre del tipo (porque no encontró la clave en ResX),
            ' significa que es un tipo nuevo del usuario. Por lo tanto, usamos la descripción original de la BD.
            If textoCelda1 = tipoParaLlave Then
                textoCelda1 = valorBD1
            End If

            ' 4. Dibujamos en la hoja (este código se queda igual)
            Dim anchoDisponibleCol1 As Integer = e.MarginBounds.Right - frmImprimirForm.Punto2.Left
            Dim formatoCortado As New StringFormat()
            formatoCortado.Trimming = StringTrimming.EllipsisCharacter
            formatoCortado.FormatFlags = StringFormatFlags.NoWrap

            Dim rectanguloCelda1 As New RectangleF(frmImprimirForm.Punto2.Left, startY, anchoDisponibleCol1, frmImprimirForm.Punto1.Height)
            e.Graphics.DrawString(textoCelda1, FuenteDetalles, Brushes.Black, rectanguloCelda1, formatoCortado)

            ' Control de renglones y páginas
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

    Private Function TraducirDinamico(textoOriginal As String, esDescripcion As Boolean) As String
        ' 1. Validamos que no venga vacío
        If String.IsNullOrEmpty(textoOriginal) Then Return ""

        ' 2. Formateamos el texto para que coincida con la Key de ResX Manager (ej: "Cuenta Corriente" -> "Cuenta_Corriente")
        Dim llave As String = textoOriginal.Trim().Replace(" ", "_")
        If esDescripcion Then llave = "Desc_" & llave

        Try
            ' 3. Buscamos en el gestor de recursos (reemplaza 'resManager' por tu objeto ResourceManager activo)
            Dim textoTraducido As String = rmse.GetString(llave)

            ' 4. Si existe traducción en el .resx la devolvemos; si no, devolvemos el texto original de la BD
            If Not String.IsNullOrEmpty(textoTraducido) Then
                Return textoTraducido
            Else
                Return textoOriginal
            End If
        Catch ex As Exception
            ' En caso de cualquier error imprevisto de lectura, no rompemos la app, devolvemos el dato original
            Return textoOriginal
        End Try
    End Function

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
        ' 1. Verificar si hay alguna fila seleccionada en el Grid
        If frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow Is Nothing Then
            MsgBox("Por favor, seleccione un tipo de cuenta de la lista.", vbExclamation, "Atención")
            Exit Sub
        End If

        ' 2. Obtener el nombre/código del tipo de cuenta de la celda (0)
        Dim filaActual As Integer = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        Dim vTxtNombre As String = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString()

        ' 3. VALIDACIÓN: Comprobar si este Tipo de Cuenta está en uso en la tabla 'cuentas'
        Dim vSqlVerificar As String = "SELECT COUNT(*) FROM cuentas WHERE TipoCUE = '" & vTxtNombre & "'"

        ' Reutilizamos tu comando por defecto asignándole la consulta de verificación
        cmdMdb1cr.CommandText = vSqlVerificar
        Dim cuentasAsociadas As Integer

        Try
            ' Ejecutamos el conteo sobre tu comando de siempre
            cuentasAsociadas = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())
        Catch ex As Exception
            MsgBox("Error al verificar la integridad de los datos: " & ex.Message, vbCritical, "Error")
            Exit Sub
        End Try

        ' 4. Bloquear el borrado si está asignado a alguna cuenta
        If cuentasAsociadas > 0 Then
            MsgBox("No se puede eliminar el tipo '" & vTxtNombre & "' porque está asignado a " & cuentasAsociadas & " cuenta(s) bancaria(s).", vbExclamation, "Acción Cancelada")
            Exit Sub
        End If

        ' 5. Confirmación de borrado seguro
        If MsgBox("¿Está seguro de eliminar el Tipo de Cuenta Bancaria '" & vTxtNombre & "'?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Borrado") = vbYes Then
            ' Limpiamos cualquier rastro de la consulta anterior en el comando
            cmdMdb1cr.Parameters.Clear()

            ' Asignamos la nueva sentencia de borrado
            Dim vtipoSql As String = "DELETE FROM tipocuentas WHERE CodigoTIP = '" & vTxtNombre & "'"
            cmdMdb1cr.CommandText = vtipoSql
            MsgBox(vtipoSql) ' Solo para depuración, puedes eliminar esta línea después de verificar que la consulta es correcta)
            Try
                ' Ejecutamos el borrado físico
                Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()

                ' Validamos si realmente la base de datos eliminó algo
                If filasAfectadas > 0 Then
                    MsgBox("¡El registro ha sido borrado correctamente!", vbInformation, "Registro Eliminado")
                Else
                    MsgBox("El registro no se pudo borrar. Verifique si el nombre '" & vTxtNombre & "' coincide exactamente en la base de datos.", vbExclamation, "Atención")
                End If

                ' 6. Recarga automática del Grid con tu método exacto
                '*****************************************************
                vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP FROM tipocuentas"
                vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
                vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
                LlenarGrid(vtipoSql, vtipoGrid, "1")

            Catch ex As Exception
                MsgBox("Error al intentar eliminar el registro: " & ex.Message, vbCritical, "Error")
            End Try
        Else
            MsgBox("Eliminación cancelada. El tipo de cuenta '" & vTxtNombre & "' no ha sido eliminado.", vbInformation, "Acción Cancelada")
        End If
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

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnEliminarRegistro, TL(2)},
            {Me.BtnSeguirBuscando, TL(4)}
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
                    'tool.Show(resManager.GetString("ToolTipEliminar"), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso
                    Dim textoKey As String = If(boton Is Me.BtnSeguirBuscando, "ToolTipSeguirBuscando", "ToolTipEliminar")
                    tool.Show(resManager.GetString(textoKey), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón no está sobre ningún botón bloqueado, ocultamos los tres
        TL(2).Hide(Me)
        TL(4).Hide(Me)
    End Sub

End Class