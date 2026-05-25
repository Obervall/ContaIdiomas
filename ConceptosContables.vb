Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class ConceptosContables

    Public vtipoSql, vtipoGrid, vTxtNombre, filaActual As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas As Integer
    Public PrintLine, Contador As Integer
    Public TL(12) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub ConceptosContables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
        ' Al abrirse la ventana, se traduce el combo automáticamente con el idioma actual
        ActualizarIdiomaComboConcepto()

        Me.KeyPreview = True
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnFiltroTipoConcepto, resManager.GetString("ToolTipAplicarFiltro"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSinFiltroTipoConcepto, resManager.GetString("ToolTipQuitarFiltro"))
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
        TL(9).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))

        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox4.MouseMove, AddressOf VerificarFiltrosDesactivados

        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoConcepto.SelectedIndex = 0

        ' Llenar Grid de CONCEPTOS
        '*************************
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvConceptos.Columns
            frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
        Next
    End Sub

    Private Sub BtnSinFiltroTipoConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroTipoConcepto.Click
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnFiltroTipoConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroTipoConcepto.Click
        BtnFiltroTipoConcepto.Enabled = False
        BtnSinFiltroTipoConcepto.Enabled = True
        ' 1. Traducimos la posición del combo al valor real que entiende la DB
        Dim tipoParaDB As String = ""
        Select Case CmbTipoConcepto.SelectedIndex
            Case 0
                tipoParaDB = "GASTO"
            Case 1
                tipoParaDB = "INGRESO"
            Case 2
                tipoParaDB = "ESPECIAL"
        End Select
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        vtipoSql += " WHERE "
        vtipoSql += "conceptos.TipoCON = '" & tipoParaDB & "' "
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub


    Private Sub CmbTipoConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTipoConcepto.SelectedIndexChanged
        If BtnFiltroTipoConcepto.Enabled = False Then
            Dim tipoParaDB As String = ""
            Select Case CmbTipoConcepto.SelectedIndex
                Case 0
                    tipoParaDB = "GASTO"
                Case 1
                    tipoParaDB = "INGRESO"
                Case 2
                    tipoParaDB = "ESPECIAL"
            End Select

            vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & tipoParaDB & "' "
            vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
            vtipoGrid = "CONCEPTOS_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
    End Sub

    ' BOTÓN BUSCAR: Abre la ventana de configuración y busca
    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        ' Ejecutamos la búsqueda respetando si el usuario marcó "Desde el primer registro"
        EjecutarBusquedaConceptos(forzarDesdeInicio:=True)
    End Sub

    ' BOTÓN SEGUIR BUSCANDO: Busca la siguiente coincidencia directamente
    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        EjecutarBusquedaConceptos(forzarDesdeInicio:=False)
    End Sub

    ' EVENTO DE TECLADO: Captura la tecla F3 para seguir buscando
    Private Sub ConceptosContables_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled AndAlso e.KeyCode = Keys.F3 Then
            EjecutarBusquedaConceptos(forzarDesdeInicio:=False)
        End If
    End Sub

    ' EL MOTOR UNIFICADO DE BÚSQUEDA (El estilo limpio de hoy)
    Private Sub EjecutarBusquedaConceptos(ByVal forzarDesdeInicio As Boolean)
        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex

        Dim buscarTexto As String = vBuscar.ToLower()
        Dim exacta As Boolean = frmBuscar.ChkExacta.Checked
        Dim desdePrimerRegistro As Boolean = frmBuscar.ChkPrimerRegistro.Checked

        ' 1. Calcular dinámicamente desde qué fila empezar a buscar
        Dim filaInicio As Integer = 0
        If Not forzarDesdeInicio OrElse Not desdePrimerRegistro Then
            ' Si ya tenemos una fila guardada en vRow, empezamos desde la siguiente
            If vRow >= 0 AndAlso vRow < DgvConceptos.Rows.Count Then
                filaInicio = vRow + 1
            ElseIf DgvConceptos.CurrentRow IsNot Nothing Then
                filaInicio = DgvConceptos.CurrentRow.Index + 1
            End If
        End If

        ' Si el punto de partida supera las filas actuales, avisamos que terminó
        If filaInicio >= DgvConceptos.Rows.Count Then
            MsgBox(resManager.GetString("NoHayMasCoincidencias"))
            BtnSeguirBuscando.Enabled = False
            Exit Sub
        End If

        ' 2. Mapear las celdas a evaluar según el ComboBox vCampo
        Dim celdasAEvaluar As Integer()
        If vCampo = 0 Then
            celdasAEvaluar = {0, 1, 2, 3} ' Todos los campos
        Else
            celdasAEvaluar = {vCampo - 1} ' Campos individuales (1->0, 2->1, etc.)
        End If

        Dim encontrado As Boolean = False

        ' 3. Bucle de búsqueda
        For i As Integer = filaInicio To DgvConceptos.Rows.Count - 1
            Dim row As DataGridViewRow = DgvConceptos.Rows(i)

            ' Saltamos la fila vacía automática si existe al final del grid
            If row.IsNewRow Then Continue For

            Dim coincide As Boolean = False

            ' Evaluar coincidencias en las celdas asignadas
            For Each idx As Integer In celdasAEvaluar
                ' Protección contra celdas con valores nulos (Nothing)
                Dim valorCelda As String = ""
                If row.Cells(idx).Value IsNot Nothing Then
                    valorCelda = row.Cells(idx).Value.ToString().ToLower()
                End If

                If exacta Then
                    coincide = (valorCelda = buscarTexto)
                Else
                    coincide = valorCelda.Contains(buscarTexto)
                End If

                If coincide Then Exit For ' Si coincide en una celda, saltamos a la fila
            Next

            ' 4. Si encontramos coincidencia, aplicamos la selección visual y guardamos la posición
            If coincide Then
                DgvConceptos.ClearSelection()
                row.Selected = True

                ' Foco inteligente: Si busca en "Todos", va a la celda 0. Si no, va a su celda correspondiente.
                Dim celdaFoco As Integer = If(vCampo = 0, 0, vCampo - 1)

                Try
                    DgvConceptos.CurrentCell = row.Cells(celdaFoco)
                Catch
                    DgvConceptos.CurrentCell = row.Cells(0)
                End Try

                vRow = row.Index ' Guardamos la fila global para la siguiente llamada de F3
                encontrado = True
                Exit For
            End If
        Next

        ' 5. Gestión del resultado final si no hubo éxito
        If Not encontrado Then
            If forzarDesdeInicio Then
                MsgBox(resManager.GetString("NoHayCoincidencia"))
            Else
                MsgBox(resManager.GetString("NoHayMasCoincidencias"))
            End If
            BtnSeguirBuscando.Enabled = False
            vRow = -1 ' Reseteamos la posición
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click

        Dim tipoParaDB As String = ""
        Select Case CmbTipoConcepto.SelectedIndex
            Case 0
                tipoParaDB = "GASTO"
            Case 1
                tipoParaDB = "INGRESO"
            Case 2
                tipoParaDB = "ESPECIAL"
        End Select

        vtipoSql = "SELECT * FROM conceptos"
        If BtnFiltroTipoConcepto.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & tipoParaDB & "' "
        End If
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"

        LlenarGrid(vtipoSql, "PRINT_CONCEPTOS", 1)
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
        If BtnFiltroTipoConcepto.Enabled = False Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoConceptosContablesFiltrado") & " " & CmbTipoConcepto.Text
        Else
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoConceptosContables")
        End If

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString(resManager.GetString("Tipo") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString(frmConceptosContables.rmse.GetString("Concepto_Contable") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        'e.Graphics.DrawString(resManager.GetString("Notas") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto1.Top - 30)

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

            Dim tipoOriginal As String = ""
            If frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value IsNot Nothing Then
                tipoOriginal = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value.ToString().Trim()
            End If

            ' Construimos la clave del archivo .resx según el valor de la BD
            Dim claveRecurso As String
            Select Case tipoOriginal.ToUpper()
                Case "GASTO" : claveRecurso = "Tipo_Gasto"
                Case "INGRESO" : claveRecurso = "Tipo_Ingreso"
                Case "ESPECIAL" : claveRecurso = "Tipo_Especial"
                Case Else : claveRecurso = tipoOriginal ' Por si añades otros en el futuro
            End Select

            ' Buscamos la traducción en tu resManager (o rmse según uses global o local)
            Dim tipoTraducido As String = resManager.GetString(claveRecurso)

            ' Si no se encuentra traducción en los .resx, dejamos el texto original
            If String.IsNullOrEmpty(tipoTraducido) Then
                tipoTraducido = tipoOriginal
            End If

            ' Imprimimos la variable traducida en lugar del valor directo del Grid
            e.Graphics.DrawString(tipoTraducido, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
            startY += frmImprimirForm.LblFecha.Height

            ' 1. Obtener el texto y limpiar los saltos de línea
            Dim textoNotas As String = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value.ToString()
            textoNotas = textoNotas.Replace(ControlChars.Cr, " ").Replace(ControlChars.Lf, " ")

            ' 2. Reemplazar espacios dobles por uno solo
            While textoNotas.Contains("  ")
                textoNotas = textoNotas.Replace("  ", " ")
            End While

            ' 3. Recortar a un máximo de 120 caracteres si es necesario
            If textoNotas.Length > 100 Then
                textoNotas = textoNotas.Substring(0, 95) & "..."
            End If

            ' 4. Preparar el prefijo de notas y medir su tamaño en píxeles
            Dim etiquetaNotas As String = resManager.GetString("Notas") & ": "

            ' Medimos cuánto mide la palabra con la fuente específica que va a usar
            Dim tamañoEtiqueta As SizeF = e.Graphics.MeasureString(etiquetaNotas, FuenteSubrayada)

            ' Calculamos la posición X exacta agregando un margen de seguridad de 10 píxeles
            Dim posicionXTexto As Integer = frmImprimirForm.Punto1.Left + CInt(tamañoEtiqueta.Width) + 10

            ' 5. Imprimir la etiqueta "Notas:"
            e.Graphics.DrawString(etiquetaNotas, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)

            ' 6. Imprimir el texto limpio usando la coordenada dinámica calculada (ya no se montará)
            e.Graphics.DrawString(textoNotas, FuenteDetalles, Brushes.Black, posicionXTexto, startY)


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
        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index
        vTxtNombre = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString

        If CmbTipoConcepto.Text = resManager.GetString("Tipo_Especial") Or vTxtNombre = resManager.GetString("Tipo_Traspaso") Then
            MsgBox(frmConceptosContables.rmse.GetString("Concepto_No_Editable"), vbCritical)
        Else
            ' Comprobamos si existe un identificador asociado.
            If ((frmEditarConceptoContable Is Nothing) OrElse (Not frmEditarConceptoContable.IsHandleCreated)) Then
                frmEditarConceptoContable = New EditarConceptoContable
            End If
            ' Llamamos al formulario de manera modal.
            If vEditar = "NO" Then
                vEditar = "NO"  ' Eliminar
            Else
                vEditar = "SI"
            End If
            frmEditarConceptoContable.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmEditarConceptoContable.Dispose()
            vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
            If BtnFiltroTipoConcepto.Enabled = False Then
                vtipoSql += " WHERE "
                vtipoSql += "conceptos.TipoCON = '" & CmbTipoConcepto.Text & "' "
            End If
            vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
            vtipoGrid = "CONCEPTOS_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            DgvConceptos.CurrentCell = DgvConceptos.Rows(filaActual).Cells(0)
            DgvConceptos.Rows(filaActual).Selected = True
        End If
    End Sub

    Private Sub DgvConceptos_DoubleClick(sender As Object, e As EventArgs) Handles DgvConceptos.DoubleClick
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        vEditar = "NO"  ' Eliminar
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmNuevoConceptoContable Is Nothing) OrElse (Not frmNuevoConceptoContable.IsHandleCreated)) Then
            frmNuevoConceptoContable = New NuevoConceptoContable
        End If
        ' Llamamos al formulario de manera modal.
        frmNuevoConceptoContable.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmNuevoConceptoContable.Dispose()
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        If BtnFiltroTipoConcepto.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & CmbTipoConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = 0
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = vFilaActual - 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = DgvConceptos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = vFilaActual + 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = DgvConceptos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = DgvConceptos.RowCount - 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnSinFiltroTipoConcepto, TL(1)},
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

    Private Sub ActualizarIdiomaComboConcepto()
        ' 1. Guardamos la posición seleccionada actual para no perder el dato
        Dim indiceSeleccionado As Integer = Me.CmbTipoConcepto.SelectedIndex

        ' 2. Limpiamos la lista vieja para que no se acumulen elementos
        Me.CmbTipoConcepto.Items.Clear()

        ' 3. Añadimos las traducciones individuales desde tu resx manager
        Me.CmbTipoConcepto.Items.Add(rmse.GetString("CmbTipoConcepto.Items"))
        Me.CmbTipoConcepto.Items.Add(rmse.GetString("CmbTipoConcepto.Items1")) ' Revisa si tu segunda clave es Items1, Items2 o ItemsA
        Me.CmbTipoConcepto.Items.Add(rmse.GetString("CmbTipoConcepto.Items2"))

        ' 4. Restauramos el ítem que el usuario estaba viendo
        If indiceSeleccionado >= 0 AndAlso indiceSeleccionado < Me.CmbTipoConcepto.Items.Count Then
            Me.CmbTipoConcepto.SelectedIndex = indiceSeleccionado
        ElseIf Me.CmbTipoConcepto.Items.Count > 0 Then
            Me.CmbTipoConcepto.SelectedIndex = 0 ' Por defecto seleccionamos el primero si no había nada
        End If
    End Sub

    Private Sub DgvConceptos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgvConceptos.CellFormatting
        ' 1. Filtramos para actuar SOLO en la columna "TIPO" y aseguramos que la celda tenga datos
        If DgvConceptos.Columns(e.ColumnIndex).Name = "TipoCON" AndAlso e.Value IsNot Nothing Then

            ' 2. Pasamos el valor original de la DB a texto y le quitamos espacios
            Dim valorOriginal As String = e.Value.ToString().Trim()

            ' 3. Traducimos "al vuelo" según lo que venga de la base de datos antigua
            Select Case valorOriginal
                Case "GASTO"
                    ' Cambiamos lo que ve el usuario por el recurso traducido
                    e.Value = resManager.GetString("Tipo_Gasto")
                    e.FormattingApplied = True ' Le decimos al sistema que ya aplicamos el cambio

                Case "INGRESO"
                    e.Value = resManager.GetString("Tipo_Ingreso")
                    e.FormattingApplied = True

                Case "ESPECIAL"
                    e.Value = resManager.GetString("Tipo_Especial")
                    e.FormattingApplied = True
            End Select
        End If
    End Sub

End Class