Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting


Public Class GraficosConceptos

    Public vAñadir, vAñadir2, vTempapu, vImporteConcepto, vNewImporteConcepto As String
    Public vExistenteImporteConcepto As String
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x As Integer
    Public b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub GraficosConceptos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' =========================================================================
        ' 1. INITIALIZACIÓN CRÍTICA DEL GRÁFICO (EVITA EL ERROR NULL / ARGUMENT)
        ' =========================================================================
        ' Borramos cualquier residuo del diseñador y creamos las series limpias
        Chart1.Series.Clear()
        Chart1.Series.Add("Gastos")
        Chart1.Series.Add("Ingresos")

        ' Establecemos el tipo inicial (Columnas) para que herede los ejes cartesianos
        Chart1.Series("Gastos").ChartType = SeriesChartType.Column
        Chart1.Series("Ingresos").ChartType = SeriesChartType.Column

        ' Traducción automática del formulario y títulos
        ActualizarTextosFormulario(Me)

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        Else
            Chart1.Titles.Add(rmse.GetString("TituloGrafico"))
        End If
        Chart1.Titles(0).Font = New Font("Arial", 12, FontStyle.Italic)

        ' Fuentes y títulos de los ejes (Multiidioma)
        Dim fuenteTitulosGrafico As New Font("Arial", 12, FontStyle.Bold)
        Chart1.ChartAreas(0).AxisX.TitleFont = fuenteTitulosGrafico
        Chart1.ChartAreas(0).AxisY.TitleFont = fuenteTitulosGrafico

        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Concepto")
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

        ' =========================================================================
        ' 2. LIMPIEZA DE LA TABLA TEMPORAL EN LA BASE DE DATOS
        ' =========================================================================
        vTempapu = "DELETE FROM tempapu"
        cmdMdb1cr.CommandText = vTempapu
        Try
            cmdMdb1cr.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' =========================================================================
        ' 3. ORDENACIÓN DINÁMICA DE LOS DATAGRIDVIEW SEGÚN EL ORIGEN
        ' =========================================================================
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
            ' Aseguramos que la función de refresco reciba los parámetros correctos
            DgvApuntesContables(3, 4)
        End If

        ' =========================================================================
        ' 4. PROCESAMIENTO Y AGRUPACIÓN DE CONCEPTOS (CON REPARACIÓN DE EXECUTES)
        ' =========================================================================
        vNombreConcepto = ""

        If vGrafico <> "" Then
            ' --- ORIGEN: APUNTES PERIÓDICOS ---
            For Each fila As DataGridViewRow In frmApuntesPeriodicos.DgvApuper.Rows
                If fila.IsNewRow Then Continue For ' Seguridad: Evitamos la fila vacía

                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString() Then
                    vNombreConcepto = fila.Cells(1).Value.ToString()
                    vImporteConcepto = fila.Cells(3).Value

                    vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMdb1cr.CommandText = vAñadir
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vExistenteImporteConcepto = drMdb1.GetValue(1)
                            End While
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto)
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMdb1cr.CommandText = vAñadir2
                    Try
                        ' CORRECCIÓN CRÍTICA: Los UPDATE usan ExecuteNonQuery, no ExecuteReader
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
        Else
            ' --- ORIGEN: APUNTES CONTABLES ---
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For ' Seguridad: Evitamos la fila vacía

                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString() Then
                    vNombreConcepto = fila.Cells(1).Value.ToString()
                    vImporteConcepto = fila.Cells(3).Value

                    vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMdb1cr.CommandText = vAñadir
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vExistenteImporteConcepto = drMdb1.GetValue(1)
                            End While
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try

                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto)
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMdb1cr.CommandText = vAñadir2
                    Try
                        ' CORRECCIÓN CRÍTICA: Los UPDATE usan ExecuteNonQuery, no ExecuteReader
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
        End If

        ' =========================================================================
        ' 5. VOLCADO DE DATOS AL DATATABLE E INYECCIÓN EN EL GRÁFICO
        ' =========================================================================
        miDataTable.Columns.Add("Concepto")
        miDataTable.Columns.Add("Importe")

        vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")

        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow Then Continue For ' Seguridad

            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Concepto") = fila.Cells(0).Value.ToString()
            vValor = fila.Cells(1).Value
            vValor = Math.Truncate(vValor)
            Renglon("Importe") = vValor.ToString()
            miDataTable.Rows.Add(Renglon)
        Next

        ' ¡LAS LÍNEAS QUE FALTABAN PARA DETECTAR LOS DATOS AL ABRIR!:
        ' Vinculamos el gráfico con tu DataTable y le asignamos las columnas correspondientes
        Chart1.DataSource = miDataTable
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        ' Forzamos al gráfico a dibujarse con los datos cargados
        Chart1.DataBind()

        ' Mostramos las leyendas de forma segura
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
    End Sub



    Private Sub TSBtnImprimir_Click(sender As Object, e As EventArgs) Handles TSBtnImprimir.Click
        'Iniciamos Código para Imprimir
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

        'Definimos los tipos de letras a utilizar en el reporte
        '******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 14)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)

        'Imprimimos el encabezado los datos que están antes del dibujo
        '*************************************************************
        e.Graphics.DrawString(frmGraficosConceptos.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        b = New Bitmap(frmGraficosConceptos.Chart1.Width, frmGraficosConceptos.Chart1.Height)
        frmGraficosConceptos.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))
        e.Graphics.DrawImage(b, 0, 100)

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

    Private Sub TsBtnColumnas_Click(sender As Object, e As EventArgs) Handles TsBtnColumnas.Click
        TsBtnColumnas.Checked = True
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        ' Recreación desde cero absoluto
        Chart1.Series.Clear()
        Chart1.Series.Add("Gastos")
        Chart1.Series.Add("Ingresos")

        ' Traducción segura de títulos principales y textos de ejes
        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        Else
            Chart1.Titles.Add(rmse.GetString("TituloGrafico"))
        End If
        Chart1.ChartAreas(0).AxisX.Title = resManager.GetString("Conceptos")
        Chart1.ChartAreas(0).AxisY.Title = resManager.GetString("Moneda")

        Dim fuenteTitulos As New Font("Arial", 12, FontStyle.Bold)
        Chart1.ChartAreas(0).AxisX.TitleFont = fuenteTitulos
        Chart1.ChartAreas(0).AxisY.TitleFont = fuenteTitulos

        ' Configuración de la estructura visual de barras
        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
        Chart1.Series("Gastos").ChartType = SeriesChartType.Column
        Chart1.Series("Ingresos").ChartType = SeriesChartType.Column

        ' Bucle de llenado de datos
        For x As Integer = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Convert.ToDouble(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
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

        Chart1.Series.Clear()
        Chart1.Series.Add("Gastos")
        Chart1.Series.Add("Ingresos")

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        Else
            Chart1.Titles.Add(rmse.GetString("TituloGrafico"))
        End If
        Chart1.ChartAreas(0).AxisX.Title = resManager.GetString("Conceptos")
        Chart1.ChartAreas(0).AxisY.Title = resManager.GetString("Moneda")

        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
        Chart1.Series("Gastos").ChartType = SeriesChartType.Line
        Chart1.Series("Ingresos").ChartType = SeriesChartType.Line
        Chart1.Series("Gastos").Color = Color.Red
        Chart1.Series("Ingresos").Color = Color.Blue

        For x As Integer = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Convert.ToDouble(miView(x)("Importe")))
                    .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                Else
                    .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                End If
            End With
        Next
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False

        Chart1.Series.Clear()
        Chart1.Series.Add("Gastos")
        Chart1.Series.Add("Ingresos")

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        Else
            Chart1.Titles.Add(rmse.GetString("TituloGrafico"))
        End If
        Chart1.ChartAreas(0).AxisX.Title = resManager.GetString("Conceptos")
        Chart1.ChartAreas(0).AxisY.Title = resManager.GetString("Moneda")

        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")
        Chart1.Series("Gastos").ChartType = SeriesChartType.Area
        Chart1.Series("Ingresos").ChartType = SeriesChartType.Area

        For x As Integer = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Convert.ToDouble(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
            End With
        Next
    End Sub

    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = True

        ' Eliminamos todo residuo de ejes creando únicamente la serie Gastos
        Chart1.Series.Clear()
        Chart1.Series.Add("Gastos")

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        Else
            Chart1.Titles.Add(rmse.GetString("TituloGrafico"))
        End If

        ' El modo pastel no trabaja con ejes cartesianos, limpiamos títulos residuales
        Chart1.ChartAreas(0).AxisX.Title = ""
        Chart1.ChartAreas(0).AxisY.Title = ""

        Chart1.Series("Gastos").LegendText = "#VALX"
        Chart1.Series("Gastos").ChartType = SeriesChartType.Pie

        For x As Integer = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Convert.ToDouble(miView(x)("Importe")))
                    .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                Else
                    .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                End If
            End With
        Next
    End Sub

End Class