Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class SeleccionFechas

    Public vtipoSql As String
    Public PrintLine, Contador As Integer
    Public PosicionSinEncabezado As Integer = frmImprimirForm.Punto1.Top
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub SeleccionFechas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        ' 1. Convertimos el año a número entero de forma segura
        Dim anio As Integer
        If Not Integer.TryParse(vAñoEjercicio, anio) Then
            ' Salvavidas: si falla o está vacío, usa el año actual
            anio = Date.Today.Year
        End If

        ' 2. Asignamos el año numérico puro a tus variables
        vFecha1Enero = anio
        vFecha31Diciembre = anio

        ' 3. Creamos los objetos de fecha límites una sola vez en memoria
        Dim fechaInicio As New Date(anio, 1, 1)
        Dim fechaFin As New Date(anio, 12, 31)

        ' 4. Configuramos de forma limpia el primer DateTimePicker
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker1.MaxDate = fechaFin
        DateTimePicker1.Value = fechaInicio

        ' 5. Configuramos de forma limpia el segundo DateTimePicker
        DateTimePicker2.MinDate = fechaInicio
        DateTimePicker2.MaxDate = fechaFin
        DateTimePicker2.Value = fechaFin


        Dim TL(2) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnHoy2, resManager.GetString("IrAHoy"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))

    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 BLOQUE 1: INFORME DE APUNTES CONTABLES (INNER JOIN CON CONCEPTOS)
        ' =========================================================================
        If vOrdenadoPorFechasAPU = 1 Or vOrdenadoPorConceptosAPU = 1 Or vOrdenadoPorImportesAPU = 1 Then
            ' Cruzamos apuntes con conceptos para extraer el CodigoCON legible en la segunda columna (Celdas(1))
            vtipoSql = "SELECT apuntes.FechaAPU, conceptos.CodigoCON, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU " &
                       "FROM apuntes " &
                       "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON " &
                       "WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString

            If vSoloIngresosAPU = 1 Then
                vtipoSql += " And apuntes.ImporteAPU > 0 "
            End If
            If vSoloGastosAPU = 1 Then
                vtipoSql += " And apuntes.ImporteAPU < 0 "
            End If

            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            frmImprimirForm.LblEntreFechas.Text = resManager.GetString("Desde") & ": " & DateTimePicker1.Value.ToShortDateString() & "    " & resManager.GetString("Hasta") & ": " & DateTimePicker2.Value.ToShortDateString()

            ' 🚀 EL TRUCO DE ALTA ESCUELA CONTABLE: Convertimos las fechas al formato nativo estandarizado de Access (#AAAA-MM-DD#)
            ' Esto destruye cualquier interferencia de puntos alemanes o barras catalanas en la ordenación.
            Dim fechaInicioAccess As String = "#" & vDate1.ToString("yyyy-MM-dd") & "#"
            Dim fechaFinAccess As String = "#" & vDate2.ToString("yyyy-MM-dd") & "#"

            ' Inyectamos las condiciones directamente con el formato blindado de Microsoft Access
            vtipoSql += " And apuntes.FechaAPU >= " & fechaInicioAccess
            vtipoSql += " And apuntes.FechaAPU <= " & fechaFinAccess

            ' Configuramos la ordenación alfabética real por el nombre del concepto
            If vOrdenadoPorFechasAPU = 1 Then
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC"
            End If
            If vOrdenadoPorConceptosAPU = 1 Then
                vtipoSql += " ORDER BY conceptos.CodigoCON ASC, apuntes.FechaAPU ASC "
            End If
            If vOrdenadoPorImportesAPU = 1 Then
                vtipoSql += " ORDER BY apuntes.ImporteAPU ASC, apuntes.FechaAPU ASC "
            End If
        End If

        ' =========================================================================
        ' 🌟 BLOQUE 2: INFORME DE APUNTES PERIÓDICOS (INNER JOIN CON CONCEPTOS)
        ' =========================================================================
        If vOrdenadoPorFechasAPP = 1 Or vOrdenadoPorConceptosAPP = 1 Or vOrdenadoPorImportesAPP = 1 Then
            ' Cruzamos apuper con conceptos para extraer el CodigoCON legible en la segunda columna (Celdas(1))
            vtipoSql = "SELECT apuper.FechaAPP, conceptos.CodigoCON, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP " &
                       "FROM apuper " &
                       "INNER JOIN conceptos ON apuper.ConceptoAPP = conceptos.IdConceptoCON " &
                       "WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString

            If vSoloIngresosAPP = 1 Then
                vtipoSql += " And apuper.ImporteAPP > 0 "
            End If
            If vSoloGastosAPP = 1 Then
                vtipoSql += " And apuper.ImporteAPP < 0 "
            End If

            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            frmImprimirForm.LblEntreFechas.Text = "Desde: " & DateTimePicker1.Value.ToShortDateString() & "    Hasta: " & DateTimePicker2.Value.ToShortDateString()

            Dim fechaInicioAccess As String = "#" & vDate1.ToString("yyyy-MM-dd") & "#"
            Dim fechaFinAccess As String = "#" & vDate2.ToString("yyyy-MM-dd") & "#"

            vtipoSql += " And apuper.FechaAPP >= " & fechaInicioAccess
            vtipoSql += " And apuper.FechaAPP <= " & fechaFinAccess

            ' Configuramos la ordenación alfabética real por el nombre del concepto
            If vOrdenadoPorFechasAPP = 1 Then
                vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            End If
            If vOrdenadoPorConceptosAPP = 1 Then
                vtipoSql += " ORDER BY conceptos.CodigoCON ASC, apuper.FechaAPP ASC"
            End If
            If vOrdenadoPorImportesAPP = 1 Then
                vtipoSql += " ORDER BY apuper.ImporteAPP ASC, apuper.FechaAPP ASC"
            End If
        End If

        ' 🌟 IMPORTANTE: Ejecutamos el volcado dócil en la macro de tu reporte
        vtipoGrid = "PRINT_INFORME_APUNTES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString()

        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

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

        Me.Close()
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

        frmImprimirForm.LblTitulo.Text = vTituloInforme

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Right, frmImprimirForm.LblEntreFechas.Top)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString(resManager.GetString("Fecha") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Concepto") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes (Tu estructura intacta)
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

            ' 🌟 1. CAPTURAMOS LOS VALORES DE LA FILA ACTUAL EN VARIABLES VOLANTILES DE LA RAM
            ' =========================================================================
            ' 🌟 EXTRACCIÓN Y LIMPIEZA DE LA FECHA (Celda 0 - ¡Adiós al horario 00:00:00!)
            ' =========================================================================
            ' =========================================================================
            ' 🌟 EXTRACCIÓN Y LIMPIEZA DE LA FECHA UNIFORME (¡Adiós desorden por idiomas!)
            ' =========================================================================
            Dim textoFechaPapel As String = ""
            If frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value IsNot Nothing Then
                Dim fechaFila As Date
                ' Parseamos la celda de Access de forma segura
                If Date.TryParse(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString(), fechaFila) Then

                    ' 🚀 EL TRUCO MAESTRO: Forzamos el formato europeo rígido (Día/Mes/Año) en el folio.
                    ' Esto impide que en inglés se barajen los números visualmente en el papel.
                    textoFechaPapel = fechaFila.ToString("d/MM/yyyy")

                Else
                    textoFechaPapel = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString()
                End If
            End If

            Dim textoConceptoPapel As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "").Trim()
            Dim textoDescripcionPapel As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value?.ToString(), "").Trim()

            ' 🌟 2. EL CORTAFUEGOS BIOLÓGICO MULTIIDIOMA ANTES DE TOCAR EL LIENZO GRÁFICO
            If resManager IsNot Nothing Then
                ' Leemos el chip regional que el hilo de Preferencias tiene activo en este microsegundo
                Dim culturaActivaEnVivo As System.Globalization.CultureInfo = Threading.Thread.CurrentThread.CurrentUICulture

                ' 🚀 A. TRADUCCIÓN DE LA COLUMNA CONCEPTO (Celda 1)
                ' Usamos tu excelente escáner inverso para cazar la Key neutral (ej: "LUZ", "SALDO")
                Dim claveConceptoNeutral As String = ObtenerClaveNeutral(textoConceptoPapel, resManager)

                If Not String.IsNullOrEmpty(claveConceptoNeutral) Then
                    Dim tradConcepto As String = resManager.GetString(claveConceptoNeutral, culturaActivaEnVivo)
                    If Not String.IsNullOrEmpty(tradConcepto) Then textoConceptoPapel = tradConcepto.Trim().ToUpper()
                Else
                    ' Plan B de respaldo por la ortografía de los recursos
                    Dim tradDirecta As String = resManager.GetString(textoConceptoPapel.Replace(" ", "_"), culturaActivaEnVivo)
                    If Not String.IsNullOrEmpty(tradDirecta) Then textoConceptoPapel = tradDirecta.Trim().ToUpper()
                End If

                ' 🚀 B. TRADUCCIÓN DE LA COLUMNA DESCRIPCIÓN (Celda 2 - Captura tu Key "Desc_SALDO")
                ' Si el escáner inverso nos chiva que este registro es el Saldo Inicial histórico de Access:
                If ObtenerClaveNeutral(textoDescripcionPapel, resManager) = "Desc_SALDO" OrElse
                   textoDescripcionPapel.Equals("Saldo Inicial", StringComparison.OrdinalIgnoreCase) OrElse
                   textoDescripcionPapel.Equals("Initial Balance", StringComparison.OrdinalIgnoreCase) Then

                    Dim tradSaldo As String = resManager.GetString("Desc_SALDO", culturaActivaEnVivo)
                    If Not String.IsNullOrEmpty(tradSaldo) Then textoDescripcionPapel = tradSaldo.Trim()
                End If
            End If

            ' 🌟 3. DIBUJO DEFINITIVO EN LA HOJA PAPEL (Imprimimos los textos ya traducidos en vivo)
            e.Graphics.DrawString(textoFechaPapel, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(textoConceptoPapel, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
            e.Graphics.DrawString(textoDescripcionPapel, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)

            ' =========================================================================
            ' VALIDACIÓN SEGURA PARA LA CELDA 3 (REMPLAZA TU LÍNEA ANTERIOR)
            ' =========================================================================
            Dim textoImporte As String = "0,00"

            If frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value IsNot DBNull.Value AndAlso frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value IsNot Nothing Then

                Dim valorCelda As Object = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value
                Dim numeroDecimal As Decimal

                ' Si el valor ya es numérico, aplicamos tu formato numérico universal
                If TypeOf valorCelda Is Decimal OrElse TypeOf valorCelda Is Double OrElse TypeOf valorCelda Is Integer Then
                    textoImporte = Convert.ToDecimal(valorCelda).ToString("###,##0.00")
                    ' Si viene como texto, intentamos convertirlo de forma segura
                ElseIf Decimal.TryParse(valorCelda.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, numeroDecimal) Then
                    textoImporte = numeroDecimal.ToString("###,##0.00")
                Else
                    ' Si por error hay un texto no numérico, pintamos lo que haya para no perder el dato
                    textoImporte = valorCelda.ToString()
                End If
            End If

            ' Imprimimos usando la variable segura que ya tiene el formato "###,##0.00"
            e.Graphics.DrawString(textoImporte, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 50, startY, sf)
            ' =========================================================================

            ' Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
            ' Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
            '' Imprimimos los valores que salen despues del datagridview al final del reporte

            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop
        'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
        'Imprimimos los valores que salen despues del datagridview al final del reporte
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 50, startY + 15, sf)

            'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
            PrintLine = 0
            Contador = 0
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

        'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
        If Contador = 0 Then
            frmImprimirForm.LblNumeroPagina.Text = "0"
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox(frmIntroApuntes.rmse.GetString("EjercicioActual"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            DateTimePicker1.Value = New Date(vAñoEjercicio, 1, 1)
        Else
            DateTimePicker1.Value = DateTime.Today
        End If
    End Sub

    Private Sub BtnHoy2_Click(sender As Object, e As EventArgs) Handles BtnHoy2.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox(frmIntroApuntes.rmse.GetString("EjercicioActual"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            DateTimePicker2.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker2.Value = DateTime.Today
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub DateTimePicker2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DateTimePicker2.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub
End Class