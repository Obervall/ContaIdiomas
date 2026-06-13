Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosMeses

    Public Property EsGrafico3D As Boolean = False
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImportePrimero, vImporteSegundo As Double
    Private b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub GraficosMeses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. Vaciamos la tabla temporal llamando a tu módulo
        LimpiarTempApu()

        ' 2. Ordenamos el DataGridView de origen según el formulario activo
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            LlenarTempApuMeses("MESES_APUNTES_PERIODICOS")
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            LlenarTempApuMeses("MESES_APUNTES_CONTABLES")
        End If

        ' 3. Estructuramos la tabla local en memoria para el gráfico
        miDataTable.Rows.Clear()
        If miDataTable.Columns.Count = 0 Then
            miDataTable.Columns.Add("Fecha")
            miDataTable.Columns.Add("Importe")
        End If

        ' 4. Rellenamos la cuadrícula intermedia de impresión
        vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")

        ' 5. Recorremos las filas analizando el texto plano (ej: "25-01" o "26-02")
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow Then Continue For
            If fila.Cells(0).Value Is Nothing OrElse IsDBNull(fila.Cells(0).Value) Then Continue For

            Dim Renglon As DataRow = miDataTable.NewRow()
            Dim textoBaseDatos As String = fila.Cells(0).Value.ToString()

            ' Troceamos las posiciones fijas del texto de la base de datos
            Dim añoDosDigitos As String = Mid(textoBaseDatos, 1, 2)
            Dim mesExtraido As String = Mid(textoBaseDatos, 4, 2)

            ' Traducimos el mes de forma correlativa manteniendo el año real
            Select Case mesExtraido
                Case "01" : Renglon("Fecha") = rmse.GetString("Enero") & añoDosDigitos
                Case "02" : Renglon("Fecha") = rmse.GetString("Febrero") & añoDosDigitos
                Case "03" : Renglon("Fecha") = rmse.GetString("Marzo") & añoDosDigitos
                Case "04" : Renglon("Fecha") = rmse.GetString("Abril") & añoDosDigitos
                Case "05" : Renglon("Fecha") = rmse.GetString("Mayo") & añoDosDigitos
                Case "06" : Renglon("Fecha") = rmse.GetString("Junio") & añoDosDigitos
                Case "07" : Renglon("Fecha") = rmse.GetString("Julio") & añoDosDigitos
                Case "08" : Renglon("Fecha") = rmse.GetString("Agosto") & añoDosDigitos
                Case "09" : Renglon("Fecha") = rmse.GetString("Septiembre") & añoDosDigitos
                Case "10" : Renglon("Fecha") = rmse.GetString("Octubre") & añoDosDigitos
                Case "11" : Renglon("Fecha") = rmse.GetString("Noviembre") & añoDosDigitos
                Case "12" : Renglon("Fecha") = rmse.GetString("Diciembre") & añoDosDigitos
                Case Else : Renglon("Fecha") = rmse.GetString("Mes") & mesExtraido
            End Select

            ' Capturamos el importe quitando decimales de forma segura
            If fila.Cells(1).Value IsNot Nothing AndAlso IsNumeric(fila.Cells(1).Value) Then
                Renglon("Importe") = Math.Truncate(Convert.ToDouble(fila.Cells(1).Value))
            Else
                Renglon("Importe") = 0
            End If
            miDataTable.Rows.Add(Renglon)
        Next

        ' 6. Inicialización automática simulando el clic en el botón de Columnas
        TsBtnColumnas.PerformClick()
    End Sub

    ' =======================================================================
    ' CONFIGURACIÓN CENTRAL DE ESTILOS E IDIOMAS (2D y 3D Conmutable)
    ' =======================================================================
    Public Sub CrearEstilos()
        ' Aseguramos que existan las series para evitar el ArgumentException
        If Chart1.Series.IndexOf("Gastos") = -1 Then Chart1.Series.Add("Gastos")
        If Chart1.Series.IndexOf("Ingresos") = -1 Then Chart1.Series.Add("Ingresos")

        Dim fuenteEjes As New Font("Arial", 12, FontStyle.Bold)
        Chart1.ChartAreas("ChartArea1").AxisX.TitleFont = fuenteEjes
        Chart1.ChartAreas("ChartArea1").AxisY.TitleFont = fuenteEjes

        ' CORRECCIÓN EFECTUADA: "Fecha" en lugar de "Fechas" para coincidir con tu Resource Manager
        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Meses")
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

        ' CONMUTADOR 2D/3D AUTOMÁTICO: Enciende o apaga el relieve según la propiedad
        Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = Me.EsGrafico3D

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        End If

        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
    End Sub

    ' =======================================================================
    ' MÉTODOS DE DIBUJADO COMPACTOS (Leyendo de miDataTable manual)
    ' =======================================================================
    Private Sub DibujarGraficoColumnas()
        CrearEstilos()

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        For x = 0 To miDataTable.Rows.Count - 1
            Dim nombreEjeX As String = miDataTable.Rows(x)("Fecha").ToString()
            Dim importeMes As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Importe"))

            If importeMes <= 0 Then
                With Chart1.Series("Gastos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, Math.Abs(importeMes))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Column
                End With
            Else
                With Chart1.Series("Ingresos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, importeMes)
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Column
                End With
            End If
        Next
    End Sub

    Private Sub TsBtnColumnas_Click(sender As Object, e As EventArgs) Handles TsBtnColumnas.Click
        TsBtnColumnas.Checked = True
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        DibujarGraficoColumnas()
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        CrearEstilos()

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        For x = 0 To miDataTable.Rows.Count - 1
            Dim nombreEjeX As String = miDataTable.Rows(x)("Fecha").ToString()
            Dim importeMes As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Importe"))

            If importeMes <= 0 Then
                With Chart1.Series("Gastos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, Math.Abs(importeMes))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Area
                End With
            Else
                With Chart1.Series("Ingresos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, importeMes)
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Area
                End With
            End If
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True
        TsBtnPastel.Checked = False

        CrearEstilos()

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        For x = 0 To miDataTable.Rows.Count - 1
            Dim nombreEjeX As String = miDataTable.Rows(x)("Fecha").ToString()
            Dim importeMes As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Importe"))

            If importeMes <= 0 Then
                With Chart1.Series("Gastos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, Math.Abs(importeMes))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Line
                End With
            Else
                With Chart1.Series("Ingresos")
                    Dim i As Integer = .Points.AddXY(nombreEjeX, importeMes)
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Line
                End With
            End If
        Next
    End Sub

    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = True

        ' Limpieza de títulos de ejes para el modo tarta
        Chart1.ChartAreas("ChartArea1").AxisX.Title = ""
        Chart1.ChartAreas("ChartArea1").AxisY.Title = ""

        ' Mapeamos dinámicamente las leyendas con los nombres de los meses (#VALX)
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").LegendText = "#VALX"
        Chart1.Series("Ingresos").LegendText = "#VALX"

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        For x = 0 To miDataTable.Rows.Count - 1
            Dim nombreEjeX As String = miDataTable.Rows(x)("Fecha").ToString()
            Dim importeMes As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Importe"))

            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), miView(x)("Importe"))
                End If
                .ChartType = SeriesChartType.Pie
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Pie
            End With
        Next
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