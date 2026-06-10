Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosPresupuestos

    Public Property EsGrafico3D As Boolean = False
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImporteConcepto, vNewImporteConcepto, vImporteConcepto2, vExistenteImporteConcepto As Double
    Private b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    ' =======================================================================
    ' EVENTO LOAD: Procesa los conceptos y presupuestos en la tabla temporal
    ' =======================================================================
    Private Sub GraficosPresupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ActualizarTextosFormulario(Me)
        ' 1. Iniciamos Tabla Tempapu
        LimpiarTempPrint()
        ' 2. Ordenamos la columna Concepto de origen
        frmPresupuestos.DgvPresupuestos.Sort(frmPresupuestos.DgvPresupuestos.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        ' 3. Llenamos la tabla tmpprint con los Conceptos Agrupados desde DgvPresupuestos
        vNombreConcepto = ""
        For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
            If fila.IsNewRow Then Continue For ' Evitamos la fila en blanco automática
            vImporteConcepto = fila.Cells(3).Value
            If vNombreConcepto <> fila.Cells(0).Value.ToString Then
                vNombreConcepto = fila.Cells(0).Value.ToString
                vImporteConcepto = 0
                vImporteConcepto = fila.Cells(3).Value
                ' Buscamos la suma total del Concepto
                vtipoSql = "SELECT * FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                vtipoSql += " And apuntes.ConceptoAPU = '" & vNombreConcepto & "' "
                vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC"
                LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")
                vImporteConcepto2 = 0
                For Each filas As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                    vImporteConcepto2 += Math.Abs(filas.Cells(4).Value)
                Next
                vAñadir = "INSERT INTO tmpprint(FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) "
                vAñadir += "VALUES ('01/01/1900', '" & vNombreConcepto & "', '', '', '' , '" & vImporteConcepto2 & "', '" & vImporteConcepto & "')"
                cmdMdb1cr.CommandText = vAñadir
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("Error al grabar el Concepto en Tmpprint")
                End Try
            Else
                cmdMdb1cr.CommandType = CommandType.Text
                cmdMdb1cr.CommandText = "SELECT * FROM tmpprint WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            vExistenteImporteConcepto = drMdb1.GetValue(6)
                        End While
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                End Try
                vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                vAñadir2 = "UPDATE tmpprint SET SaldoTMP = '" & vNewImporteConcepto & "' WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                cmdMdb1cr.CommandText = vAñadir2
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("Error al actualizar el Importe del Concepto en tmpprint")
                End Try
            End If
        Next
        ' 4. Volcado final a la cuadrícula de datos para la gráfica
        vtipoSql = "SELECT * FROM tmpprint ORDER BY tmpprint.ConceptoTMP ASC"
        LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")
        '5. Estructuramos y vaciamos la tabla en memoria miDataTable
        miDataTable.Rows.Clear()
        miDataTable.Columns.Clear()
        miDataTable.Columns.Add("Concepto")
        miDataTable.Columns.Add("Real")
        miDataTable.Columns.Add("Presupuestado")
        vValor = 0

        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow Then Continue For
            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Concepto") = fila.Cells(1).Value.ToString
            vValor = fila.Cells(5).Value
            Renglon("Real") = Math.Truncate(vValor).ToString
            vValor = fila.Cells(6).Value
            Renglon("Presupuestado") = Math.Truncate(vValor).ToString
            miDataTable.Rows.Add(Renglon)
        Next

        '' Aseguramos que las columnas existan en la tabla en memoria
        'If miDataTable.Columns.Count = 0 Then
        '    miDataTable.Columns.Add("Concepto")
        '    miDataTable.Columns.Add("Real")
        '    miDataTable.Columns.Add("Presupuestado")
        'End If
        '' 5. Estructuramos y volcamos los datos al DataTable de la gráfica
        '' *=============================================================
        'miDataTable.Rows.Clear()

        '' Aseguramos que las columnas existan en la tabla en memoria
        'If miDataTable.Columns.Count = 0 Then
        '    miDataTable.Columns.Add("Concepto")
        '    miDataTable.Columns.Add("Real")
        '    miDataTable.Columns.Add("Presupuestado")
        'End If

        'For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
        '    If fila.IsNewRow Then Continue For
        '    If fila.Cells(1).Value Is Nothing Then Continue For ' Si no hay concepto, saltamos

        '    Dim Renglon As DataRow = miDataTable.NewRow()

        '    ' Guardamos el nombre del concepto (Alimentación, Luz, etc.)
        '    Renglon("Concepto") = fila.Cells(1).Value.ToString()

        '    ' LEER GASTO REAL: Capturamos de la celda 5 (ImporteTMP) de forma numérica pura
        '    If fila.Cells(5).Value IsNot Nothing AndAlso IsNumeric(fila.Cells(5).Value) Then
        '        Dim valorReal As Double = Convert.ToDouble(fila.Cells(5).Value)
        '        Renglon("Real") = Math.Truncate(valorReal).ToString()
        '    Else
        '        Renglon("Real") = "0"
        '    End If

        '    ' LEER PRESUPUESTO: Capturamos de la celda 6 (SaldoTMP) de forma numérica pura
        '    If fila.Cells(6).Value IsNot Nothing AndAlso IsNumeric(fila.Cells(6).Value) Then
        '        Dim valorPresupuesto As Double = Convert.ToDouble(fila.Cells(6).Value)
        '        Renglon("Presupuestado") = Math.Truncate(valorPresupuesto).ToString()
        '    Else
        '        Renglon("Presupuestado") = "0"
        '    End If

        '    miDataTable.Rows.Add(Renglon)
        'Next

        '' 6. Forzamos la carga del gráfico de columnas al arrancar
        TsBtnColumnas.PerformClick()
    End Sub

    ' =======================================================================
    ' CONFIGURACIÓN CENTRAL DE ESTILOS E IDIOMAS (2D y 3D Conmutable)
    ' =======================================================================
    Public Sub CrearEstilos()
            ' Aseguramos que existan las series en el control para evitar errores
            If Chart1.Series.IndexOf("Real") = -1 Then Chart1.Series.Add("Real")
            If Chart1.Series.IndexOf("Presupuestado") = -1 Then Chart1.Series.Add("Presupuestado")

            ' Fuentes en negrita para los ejes de la gráfica
            Dim fuenteEjes As New Font("Arial", 12, FontStyle.Bold)
            Chart1.ChartAreas("ChartArea1").AxisX.TitleFont = fuenteEjes
            Chart1.ChartAreas("ChartArea1").AxisY.TitleFont = fuenteEjes

        ' Asignación segura de textos desde tus archivos de recursos Keys
        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Conceptos")
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

            ' CONMUTADOR AUTOMÁTICO: Enciende o apaga el volumen 3D según la selección del usuario
            Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = Me.EsGrafico3D

            ' Título superior traducido
            If Chart1.Titles.Count > 0 Then
                Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
            End If

            ' Activamos las leyendas con los textos correspondientes del resManager
            Chart1.Series("Real").IsVisibleInLegend = True
            Chart1.Series("Presupuestado").IsVisibleInLegend = True
        Chart1.Series("Real").LegendText = resManager.GetString("Realidad")
        Chart1.Series("Presupuestado").LegendText = resManager.GetString("Presupuestado")
        End Sub

        ' =======================================================================
        ' MÉTODOS DE DIBUJADO COMPACTOS (Lectura manual pura desde miDataTable)
        ' =======================================================================
        Private Sub DibujarGraficoColumnas()
            ' Cargamos la configuración visual antes de pintar
            CrearEstilos()

            ' Limpiamos los puntos de datos previos para evitar acumulaciones
            Chart1.Series("Real").Points.Clear()
            Chart1.Series("Presupuestado").Points.Clear()

            ' Recorremos las filas limpias de la tabla agrupada de presupuestos
            For x = 0 To miDataTable.Rows.Count - 1
                Dim nombreConcepto As String = miDataTable.Rows(x)("Concepto").ToString()
                Dim valorReal As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Real"))
                Dim valorPresupuesto As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Presupuestado"))

                ' Serie Real (Barras Rojas)
                With Chart1.Series("Real")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorReal))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Column
                End With

                ' Serie Presupuestado (Barras Azules)
                With Chart1.Series("Presupuestado")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorPresupuesto))
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Column
                End With
            Next
        End Sub

        ' =======================================================================
        ' MANEJADORES DE EVENTOS DE LA BOTONERA (TOOLSTRIP)
        ' =======================================================================

        ' BOTÓN COLUMNAS
        Private Sub TsBtnColumnas_Click(sender As Object, e As EventArgs) Handles TsBtnColumnas.Click
            TsBtnColumnas.Checked = True
            TsBtnAreas.Checked = False
            TsBtnLineas.Checked = False

            DibujarGraficoColumnas()
        End Sub

        ' BOTÓN ÁREAS
        Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
            TsBtnColumnas.Checked = False
            TsBtnAreas.Checked = True
            TsBtnLineas.Checked = False

            CrearEstilos()

            Chart1.Series("Real").Points.Clear()
            Chart1.Series("Presupuestado").Points.Clear()

            For x = 0 To miDataTable.Rows.Count - 1
                Dim nombreConcepto As String = miDataTable.Rows(x)("Concepto").ToString()
                Dim valorReal As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Real"))
                Dim valorPresupuesto As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Presupuestado"))

                With Chart1.Series("Real")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorReal))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Area
                End With

                With Chart1.Series("Presupuestado")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorPresupuesto))
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Area
                End With
            Next
        End Sub

        ' BOTÓN LÍNEAS
        Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
            TsBtnColumnas.Checked = False
            TsBtnAreas.Checked = False
            TsBtnLineas.Checked = True

            CrearEstilos()

            Chart1.Series("Real").Points.Clear()
            Chart1.Series("Presupuestado").Points.Clear()

            For x = 0 To miDataTable.Rows.Count - 1
                Dim nombreConcepto As String = miDataTable.Rows(x)("Concepto").ToString()
                Dim valorReal As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Real"))
                Dim valorPresupuesto As Decimal = Convert.ToDecimal(miDataTable.Rows(x)("Presupuestado"))

                With Chart1.Series("Real")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorReal))
                    .Points(i).Color = Color.Red
                    .ChartType = SeriesChartType.Line
                End With

                With Chart1.Series("Presupuestado")
                    Dim i As Integer = .Points.AddXY(nombreConcepto, Math.Abs(valorPresupuesto))
                    .Points(i).Color = Color.Blue
                    .ChartType = SeriesChartType.Line
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