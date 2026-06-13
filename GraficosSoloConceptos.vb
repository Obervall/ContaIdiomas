Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosSoloConceptos

    Public Property EsGrafico3D As Boolean = False
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x As Integer
    Private b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub GraficosSoloConceptos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Iniciamos Tabla Tempapu
        '***********************
        LimpiarTempApu()

        ''Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
        ''***********************************************************************
        vNombreConcepto = ""

        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow Then Continue For

            ' 🚨 VARIABLES NUMÉRICAS PURAS: Forzamos el uso de Double para evitar problemas de String
            Dim vImporteConceptoNum As Double = 0
            Dim vExistenteImporteConceptoNum As Double = 0
            Dim vNewImporteConceptoNum As Double = 0

            ' 1. Conversión segura del importe que viene de la celda (Columna 3)
            If fila.Cells(3).Value IsNot Nothing Then
                Dim textoImporte As String = fila.Cells(3).Value.ToString().Trim()
                ' Conversión segura respetando comas y puntos regionales
                If Not Double.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, vImporteConceptoNum) Then
                    ' Plan B de seguridad por si viene en formato universal invariante
                    Double.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, vImporteConceptoNum)
                End If
            End If

            ' Comprobamos si el concepto cambia
            If vNombreConcepto <> fila.Cells(1).Value.ToString() Then
                vNombreConcepto = fila.Cells(1).Value.ToString().Trim()

                ' INSERT PARAMETRIZADO: Evita fallos por comillas simples en el nombre (ej: "L'Alimentació")
                vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                cmdMdb1cr.CommandType = CommandType.Text
                cmdMdb1cr.CommandText = vAñadir
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)
                cmdMdb1cr.Parameters.AddWithValue("@Importe", vImporteConceptoNum)

                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorGrabarTemporal") & ": " & ex.Message)
                End Try
            Else
                ' Si el concepto es el mismo, buscamos el acumulado actual en la MDB
                cmdMdb1cr.CommandType = CommandType.Text
                cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE tempapu.ConceptoAPU = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                Try
                    Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                        If drMdb1.Read() Then
                            If drMdb1.GetValue(0) IsNot DBNull.Value Then
                                ' Extraemos el valor numérico puro directamente de la MDB
                                Double.TryParse(drMdb1.GetValue(0).ToString(), vExistenteImporteConceptoNum)
                            End If
                        End If
                    End Using
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorVerificarIntegridad") & ": " & ex.Message)
                End Try

                ' 🔥 MATEMÁTICAS SEGURAS: Ahora que ambos son numéricos puros, sumamos sin miedo a solapamientos de String
                vNewImporteConceptoNum = vImporteConceptoNum + vExistenteImporteConceptoNum

                ' UPDATE PARAMETRIZADO: Guardamos el nuevo total de forma limpia
                vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ?"
                cmdMdb1cr.CommandText = vAñadir2
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@NuevoImporte", vNewImporteConceptoNum)
                cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorGrabarTemporal") & ": " & ex.Message)
                End Try
            End If
        Next

        miDataTable.Columns.Add("Concepto")
        miDataTable.Columns.Add("Importe")
        Dim unused As DataRow = miDataTable.NewRow()
        vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            'Guardamos los datos en un database
            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Concepto") = fila.Cells(0).Value.ToString
            vValor = fila.Cells(1).Value
            vValor = Math.Truncate(vValor)
            Renglon("Importe") = vValor.ToString
            miDataTable.Rows.Add(Renglon)
        Next
        TsBtnColumnas.Checked = True
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
        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Conceptos")

        ' Eje Y (Vertical - A la izquierda del gráfico)
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

        If vGraficoSolo = "GastosPorConcepto" Then
            Chart1.Series("Gastos").IsVisibleInLegend = True
            Chart1.Series("Gastos").XValueMember = "Concepto"
            Chart1.Series("Gastos").Points.Clear()
            Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Else 'IngresosPorConcepto
            Chart1.Series("Ingresos").IsVisibleInLegend = True
            Chart1.Series("Ingresos").XValueMember = "Concepto"
            Chart1.Series("Ingresos").Points.Clear()
            Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
        End If
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
        CrearEstilos()

        ' 1. DETERMINACIÓN DINÁMICA DE SERIES
        Dim serieActiva As String = "Gastos"
        Dim serieOcultar As String = "Ingresos"

        If vGraficoSolo = "IngresosPorConcepto" Then
            serieActiva = "Ingresos"
            serieOcultar = "Gastos"
        End If

        ' 2. RESTAURACIÓN DE SEGURIDAD TRAS EL PASTEL
        Chart1.Series(serieActiva).Enabled = True
        Chart1.Series(serieActiva).ChartType = SeriesChartType.Column
        Chart1.Series(serieActiva).Points.Clear()
        Chart1.Series(serieActiva).LegendText = resManager.GetString(serieActiva)

        ' Desactivamos la serie que no corresponde para que no interfiera
        If Chart1.Series.FindByName(serieOcultar) IsNot Nothing Then
            Chart1.Series(serieOcultar).Points.Clear()
            Chart1.Series(serieOcultar).Enabled = False
        End If

        ' 3. LLENADO DINÁMICO DE PUNTOS
        For x = 0 To miView.Count - 1
            With Chart1.Series(serieActiva)
                Dim nombreConcepto As String = "Sin Nombre"
                If miView(x)("Concepto") IsNot DBNull.Value AndAlso miView(x)("Concepto") IsNot Nothing Then
                    nombreConcepto = miView(x)("Concepto").ToString()
                End If

                If miView(x)("Importe") <= 0 Then
                    Dim importePuro As Decimal = 0.0D
                    If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                        Dim textoImporte As String = miView(x)("Importe").ToString()
                        If Not Decimal.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importePuro) Then
                            Decimal.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importePuro)
                        End If
                    End If
                    vImporteConcepto = Math.Abs(importePuro)

                    Dim i As Integer = .Points.AddXY(nombreConcepto, vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(nombreConcepto, miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
            End With
        Next
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        CrearEstilos()

        Dim serieActiva As String = "Gastos"
        Dim serieOcultar As String = "Ingresos"

        If vGraficoSolo = "IngresosPorConcepto" Then
            serieActiva = "Ingresos"
            serieOcultar = "Gastos"
        End If

        Chart1.Series(serieActiva).Enabled = True
        Chart1.Series(serieActiva).ChartType = SeriesChartType.Area
        Chart1.Series(serieActiva).Points.Clear()
        Chart1.Series(serieActiva).LegendText = resManager.GetString(serieActiva)

        If Chart1.Series.FindByName(serieOcultar) IsNot Nothing Then
            Chart1.Series(serieOcultar).Points.Clear()
            Chart1.Series(serieOcultar).Enabled = False
        End If

        For x = 0 To miView.Count - 1
            With Chart1.Series(serieActiva)
                Dim nombreConcepto As String = "Sin Nombre"
                If miView(x)("Concepto") IsNot DBNull.Value AndAlso miView(x)("Concepto") IsNot Nothing Then
                    nombreConcepto = miView(x)("Concepto").ToString()
                End If

                If miView(x)("Importe") <= 0 Then
                    Dim importePuro As Decimal = 0.0D
                    If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                        Dim textoImporte As String = miView(x)("Importe").ToString()
                        If Not Decimal.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importePuro) Then
                            Decimal.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importePuro)
                        End If
                    End If
                    vImporteConcepto = Math.Abs(importePuro)
                    Dim i As Integer = .Points.AddXY(nombreConcepto, vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(nombreConcepto, miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
            End With
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True
        TsBtnPastel.Checked = False

        CrearEstilos()

        Dim serieActiva As String = "Gastos"
        Dim serieOcultar As String = "Ingresos"

        If vGraficoSolo = "IngresosPorConcepto" Then
            serieActiva = "Ingresos"
            serieOcultar = "Gastos"
        End If

        Chart1.Series(serieActiva).Enabled = True
        Chart1.Series(serieActiva).ChartType = SeriesChartType.Line
        Chart1.Series(serieActiva).Points.Clear()
        Chart1.Series(serieActiva).LegendText = resManager.GetString(serieActiva)

        If Chart1.Series.FindByName(serieOcultar) IsNot Nothing Then
            Chart1.Series(serieOcultar).Points.Clear()
            Chart1.Series(serieOcultar).Enabled = False
        End If

        For x = 0 To miView.Count - 1
            With Chart1.Series(serieActiva)
                Dim nombreConcepto As String = "Sin Nombre"
                If miView(x)("Concepto") IsNot DBNull.Value AndAlso miView(x)("Concepto") IsNot Nothing Then
                    nombreConcepto = miView(x)("Concepto").ToString()
                End If

                If miView(x)("Importe") <= 0 Then
                    Dim importePuro As Decimal = 0.0D
                    If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                        Dim textoImporte As String = miView(x)("Importe").ToString()
                        If Not Decimal.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importePuro) Then
                            Decimal.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importePuro)
                        End If
                    End If
                    vImporteConcepto = Math.Abs(importePuro)
                    Dim i As Integer = .Points.AddXY(nombreConcepto, vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(nombreConcepto, miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
            End With
        Next
    End Sub


    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        ' Estado de los botones de la barra de herramientas
        TsBtnPastel.Checked = True
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False

        ' El pastel no utiliza títulos de ejes
        Chart1.ChartAreas("ChartArea1").AxisX.Title = ""
        Chart1.ChartAreas("ChartArea1").AxisY.Title = ""

        ' DETERMINACIÓN DE SERIES SEGÚN TU VARIABLE vGraficoSolo
        Dim serieActiva As String = "Gastos"
        Dim serieOcultar As String = "Ingresos"

        If vGraficoSolo = "IngresosPorConcepto" Then
            serieActiva = "Ingresos"
            serieOcultar = "Gastos"
        End If

        ' Aseguramos que la leyenda exista y esté encendida
        If Chart1.Legends.Count = 0 Then Chart1.Legends.Add(New Legend("Default"))
        Chart1.Legends(0).Enabled = True

        ' CONFIGURACIÓN DE LA SERIE ACTIVA
        Chart1.Series(serieActiva).Enabled = True
        Chart1.Series(serieActiva).Legend = Chart1.Legends(0).Name

        ' Solución a los ceros: Forzamos a leer el texto del eje X
        Chart1.Series(serieActiva).LegendText = "#AXISLABEL"
        Chart1.Series(serieActiva).Points.Clear()

        ' Desactivamos la serie opuesta para evitar el error "Pie cannot be combined"
        If Chart1.Series.FindByName(serieOcultar) IsNot Nothing Then
            Chart1.Series(serieOcultar).Points.Clear()
            Chart1.Series(serieOcultar).Enabled = False
        End If

        ' Llenado de datos desde el DataView
        For x = 0 To miView.Count - 1
            With Chart1.Series(serieActiva)
                Dim nombreConcepto As String = "Sin Nombre"
                If miView(x)("Concepto") IsNot DBNull.Value AndAlso miView(x)("Concepto") IsNot Nothing Then
                    nombreConcepto = miView(x)("Concepto").ToString()
                End If

                If miView(x)("Importe") <= 0 Then
                    Dim importePuro As Decimal = 0.0D
                    If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                        Dim textoImporte As String = miView(x)("Importe").ToString()
                        If Not Decimal.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importePuro) Then
                            Decimal.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importePuro)
                        End If
                    End If
                    vImporteConcepto = Math.Abs(importePuro)
                    .Points.AddXY(nombreConcepto, vImporteConcepto)
                Else
                    .Points.AddXY(nombreConcepto, miView(x)("Importe"))
                End If

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
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

End Class