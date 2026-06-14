Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosCuentas

    Public Property EsGrafico3D As Boolean = False
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImportePrimero, vImporteSegundo As Double
    Private b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub GraficosCuentas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Iniciamos Tabla Tempapu
        '***********************
        LimpiarTempApu()

        'Ordenamos la columna Cuenta, antes de calcular los totales parciales.
        '***********************************************************************
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(6), System.ComponentModel.ListSortDirection.Ascending)
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(6), System.ComponentModel.ListSortDirection.Ascending)
        End If

        'Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
        '***********************************************************************
        vNombreConcepto = ""

        If vGrafico <> "" Then
            'Viene de Apuntes Periódicos
            LlenarTempApuCuentas("CUENTAS_APUNTES_PERIODICOS")
        Else
            ' Viene de Apuntes Contables
            LlenarTempApuCuentas("CUENTAS_APUNTES_CONTABLES")
        End If

        miDataTable.Columns.Add("Cuenta")
        miDataTable.Columns.Add("Importe")
        Dim unused As DataRow = miDataTable.NewRow()
        vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            'Guardamos los datos en un database
            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Cuenta") = fila.Cells(0).Value.ToString
            vValor = fila.Cells(1).Value
            vValor = Math.Truncate(vValor)
            Renglon("Importe") = vValor.ToString
            miDataTable.Rows.Add(Renglon)
        Next

        DibujarGraficoColumnas()
    End Sub

    Public Sub CrearEstilos()
        ' 1. Creamos el estilo de la fuente (por ejemplo: Arial, Tamaño 12, Negrita)
        Dim fuenteEjes As New Font("Arial", 12, FontStyle.Bold)

        ' 2. Aplicamos la fuente al Eje X (Conceptos)
        Chart1.ChartAreas("ChartArea1").AxisX.TitleFont = fuenteEjes

        ' 3. Aplicamos la fuente al Eje Y (Moneda)
        Chart1.ChartAreas("ChartArea1").AxisY.TitleFont = fuenteEjes

        ' Traducción segura del título principal de la gráfica
        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        End If
        ' Eje X (Horizontal - Abajo del gráfico)
        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Cuentas")

        ' Eje Y (Vertical - A la izquierda del gráfico)
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"
        Chart1.Series("Gastos").Points.Clear()
        ' 1. Primero traducimos la leyenda leyendo el resManager
        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
        Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = Me.EsGrafico3D

    End Sub

    Private Sub TsBtnColumnas_Click(sender As Object, e As EventArgs) Handles TsBtnColumnas.Click
        TsBtnColumnas.Checked = True
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        DibujarGraficoColumnas()
    End Sub

    Private Sub DibujarGraficoColumnas()
        ' 1. Cargamos fuentes, ejes y la propiedad EsGrafico3D
        CrearEstilos()

        ' 2. Limpieza obligatoria de puntos previos (Sin XValueMember ni YValueMembers fantasmas)
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            vContador += 1
            ' Conversión segura multiidioma del importe (¡Centralizado en tu módulo!)
            Dim importeActual As Decimal = ConvertirDecimalSeguro(miView(x)("Importe"))

            vImporteConcepto = importeActual

            If (vContador Mod 2) <> 0 Then
                ' --- El número es impar ---
                vImportePrimero = vImporteConcepto

                ' Protección contra errores de desbordamiento de índice al final de la tabla
                If x + 1 < miView.Count Then
                    Dim importeSiguiente As Decimal = 0.0D
                    If miView(x + 1)("Importe") IsNot DBNull.Value AndAlso miView(x + 1)("Importe") IsNot Nothing Then
                        Decimal.TryParse(miView(x + 1)("Importe").ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importeSiguiente)
                    End If
                    vImporteSegundo = importeSiguiente
                Else
                    vImporteSegundo = 0
                End If

                ' Evaluaciones de Impares
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                End If
            Else
                ' --- El número es par ---
                vImporteSegundo = vImporteConcepto

                ' Evaluaciones de Pares
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Column)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Column)
                End If
            End If
        Next
    End Sub

    ' Función auxiliar universal compartida para Columnas, Áreas y Líneas
    Private Sub AñadirPuntoGrafico(nombreSerie As String, cuenta As Object, valor As Decimal, colorPunto As Color, tipoGrafico As SeriesChartType)
        With Chart1.Series(nombreSerie)
            Dim i As Integer = .Points.AddXY(cuenta, Math.Abs(valor))
            .Points(i).Color = colorPunto
            .ChartType = tipoGrafico
        End With
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        ' 1. Sincronización visual estricta de los botones
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        ' 2. Cargamos fuentes, ejes e idiomas
        CrearEstilos()

        ' 3. Limpieza de puntos y miembros fantasmas
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            vContador += 1

            ' Conversión segura multiidioma del importe actual (¡Centralizado en tu módulo!)
            Dim importeActual As Decimal = ConvertirDecimalSeguro(miView(x)("Importe"))

            vImporteConcepto = importeActual

            If (vContador Mod 2) <> 0 Then
                ' --- El número es impar ---
                vImportePrimero = vImporteConcepto

                ' Protección contra errores de desbordamiento de índice al final de la tabla
                If x + 1 < miView.Count Then
                    Dim importeSiguiente As Decimal = 0.0D
                    If miView(x + 1)("Importe") IsNot DBNull.Value AndAlso miView(x + 1)("Importe") IsNot Nothing Then
                        Decimal.TryParse(miView(x + 1)("Importe").ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importeSiguiente)
                    End If
                    vImporteSegundo = importeSiguiente
                Else
                    vImporteSegundo = 0
                End If

                ' Evaluaciones de Impares (Enviamos "SeriesChartType.Area")
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                End If
            Else
                ' --- El número es par ---
                vImporteSegundo = vImporteConcepto

                ' Evaluaciones de Pares (Enviamos "SeriesChartType.Area")
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Area)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Area)
                End If
            End If
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        ' 1. Sincronización visual estricta de la botonera
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True
        TsBtnPastel.Checked = False

        ' 2. Cargamos fuentes, ejes e idiomas
        CrearEstilos()

        ' 3. Limpieza de puntos y propiedades fantasma
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            vContador += 1

            ' Conversión segura multiidioma del importe actual (¡Centralizado en tu módulo!)
            Dim importeActual As Decimal = ConvertirDecimalSeguro(miView(x)("Importe"))

            vImporteConcepto = importeActual

            If (vContador Mod 2) <> 0 Then
                ' --- El número es impar ---
                vImportePrimero = vImporteConcepto

                ' Protección contra errores de desbordamiento de índice al final de la tabla
                If x + 1 < miView.Count Then
                    Dim importeSiguiente As Decimal = 0.0D
                    If miView(x + 1)("Importe") IsNot DBNull.Value AndAlso miView(x + 1)("Importe") IsNot Nothing Then
                        Decimal.TryParse(miView(x + 1)("Importe").ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importeSiguiente)
                    End If
                    vImporteSegundo = importeSiguiente
                Else
                    vImporteSegundo = 0
                End If

                ' Evaluaciones de Impares (Enviamos "SeriesChartType.Line")
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                End If
            Else
                ' --- El número es par ---
                vImporteSegundo = vImporteConcepto

                ' Evaluaciones de Pares (Enviamos "SeriesChartType.Line")
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Cuenta"), vImporteConcepto, Color.Blue, SeriesChartType.Line)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Cuenta"), vImporteConcepto, Color.Red, SeriesChartType.Line)
                End If
            End If
        Next
    End Sub

    Public Sub CrearEstilosPastelCuentas()
        ' Limpieza obligatoria de los títulos de los ejes para el modo Pastel
        Chart1.ChartAreas("ChartArea1").AxisX.Title = ""
        Chart1.ChartAreas("ChartArea1").AxisY.Title = ""

        ' Forzar visibilidad en leyenda y mapear dinámicamente con las etiquetas de las cuentas (#VALX)
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").LegendText = "#VALX"
        Chart1.Series("Ingresos").LegendText = "#VALX"

        ' Mapeo de miembros de datos
        Chart1.Series("Gastos").XValueMember = "Cuenta"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Cuenta"
        Chart1.Series("Ingresos").YValueMembers = "Importe"
    End Sub

    Private Sub DibujarGraficoPastelCuentas()
        ' 1. Aplicamos los estilos y mapeos específicos
        CrearEstilosPastelCuentas()

        ' 2. Limpieza estricta de puntos previos
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        ' 3. Recorrido seguro de la vista
        For x = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                ' Conversión segura multiidioma del importe (¡Centralizado en tu módulo!)
                Dim importePuro As Decimal = ConvertirDecimalSeguro(miView(x)("Importe"))

                If importePuro <= 0 Then
                    vImporteConcepto = Math.Abs(importePuro)
                    .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                Else
                    .Points.AddXY(miView(x)("Cuenta"), importePuro)
                End If
                .ChartType = SeriesChartType.Pie
            End With

            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Pie
            End With
        Next
    End Sub

    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        ' 1. GESTIÓN VISUAL (Haz esto antes que nada)
        ' Apagamos TODOS de forma estricta
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False

        ' Encendemos únicamente el Pastel
        TsBtnPastel.Checked = True

        ' 2. DIBUJAR EL GRÁFICO
        ' Si hay un fallo de datos aquí dentro, los botones ya se habrán corregido en pantalla
        DibujarGraficoPastelCuentas()
    End Sub

    Private Sub TSBtnImprimir_Click(sender As Object, e As EventArgs) Handles TSBtnImprimir.Click
        ' PREGUNTA DE ORIENTACIÓN: Preguntamos si desea imprimir en Horizontal (Landscape)
        Dim respuesta As DialogResult = MessageBox.Show(
            resManager.GetString("PreguntaHorizontal"), ' O pon el texto directo: "¿Deseas imprimir el gráfico en orientación Horizontal?"
            resManager.GetString("TituloPregunta"),     ' O pon el texto directo: "Orientación de página"
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If respuesta = DialogResult.Yes Then
            PrintDocument1.DefaultPageSettings.Landscape = True  ' Horizontal
        Else
            PrintDocument1.DefaultPageSettings.Landscape = False ' Vertical (Defecto)
        End If

        'Iniciamos Código para Imprimir (Tu código intacto)
        '******************************
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString
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

        'Definimos los tipos de letras a utilizar en el reporte (Tus fuentes intactas)
        '******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 14)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)

        'Imprimimos el encabezado los datos que están antes del dibujo (Cambiado a Me.Chart1 para que sea universal)
        '*************************************************************
        e.Graphics.DrawString(Me.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)

        Dim posXFecha As Integer = e.MarginBounds.Right - 150
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, posXFecha, frmImprimirForm.LblFecha.Top)

        ' =======================================================================
        ' 1. CAPTURA: Usamos Me.Chart1 para capturar de forma segura el gráfico actual
        ' =======================================================================
        b = New Bitmap(Me.Chart1.Width, Me.Chart1.Height)
        Me.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))

        ' =======================================================================
        ' 2. ESCALA: Calculamos las dimensiones optimizadas para Vertical y Horizontal
        ' =======================================================================
        ' Tomamos el ancho útil disponible de la hoja según su orientación
        Dim anchoDestino As Integer = e.MarginBounds.Width

        ' Calculamos la altura proporcional base
        Dim altoDestino As Integer = CInt((anchoDestino / b.Width) * b.Height)

        ' CONTROL PARA VERTICAL: Si el papel está en vertical, calculamos el espacio útil hacia abajo
        If Not PrintDocument1.DefaultPageSettings.Landscape Then
            ' Calculamos el alto máximo disponible en el folio (restando el encabezado de arriba)
            Dim altoMaximoDisponible As Integer = e.MarginBounds.Height - 150

            ' Si el gráfico es muy pequeño y sobra mucho espacio, lo expandimos un 35% más a lo alto
            If altoDestino < (altoMaximoDisponible * 0.6) Then
                altoDestino = CInt(altoMaximoDisponible * 0.65)
            End If
        End If

        ' =======================================================================
        ' 2. ESCALA: Máxima expansión aprovechando los bordes del papel
        ' =======================================================================
        If PrintDocument1.DefaultPageSettings.Landscape = True Then
            ' --- CONFIGURACIÓN PARA HORIZONTAL (Se mantiene como te gustaba) ---
            anchoDestino = e.MarginBounds.Width
            altoDestino = CInt((anchoDestino / b.Width) * b.Height)

            ' Creamos el rectángulo alineado al margen izquierdo estándar
            Dim rectanguloPapel As New Rectangle(e.MarginBounds.Left, 100, anchoDestino, altoDestino)
            e.Graphics.DrawImage(b, rectanguloPapel)
        Else
            ' --- CONFIGURACIÓN PARA VERTICAL (Agrandado al límite de la hoja) ---
            ' 1. Tomamos el ancho total absoluto físico del papel (Saltamos el margen)
            Dim anchoPapelTotal As Integer = e.PageBounds.Width

            ' 2. Dejamos solo un pequeño borde estético de seguridad (ej: 25 píxeles por lado)
            anchoDestino = anchoPapelTotal - 50

            ' 3. Calculamos el alto de forma estrictamente proporcional para que no se deforme
            altoDestino = CInt((anchoDestino / b.Width) * b.Height)

            ' 4. Creamos el rectángulo centrado (X=25 para equilibrar los bordes)
            Dim rectanguloPapelVertical As New Rectangle(25, 100, anchoDestino, altoDestino)
            e.Graphics.DrawImage(b, rectanguloPapelVertical)
        End If

        'Si deseamos poner un contador de páginas (Tu código intacto)
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

End Class