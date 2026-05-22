Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosConceptos3D

    Public vAñadir, vAñadir2, vTempapu, vImporteConcepto, vNewImporteConcepto As String
    Public vExistenteImporteConcepto As String
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x As Integer
    Private b As Bitmap

    Private Sub GraficosConceptos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        'Iniciamos Tabla Tempapu
        '***********************
        vTempapu = "DELETE FROM tempapu"
        cmdMySql1cr.CommandText = vTempapu
        Try
            cmdMySql1cr.ExecuteNonQuery()
            'MsgBox("Registros Tempapu, Borrados !!!")
        Catch ex As Exception
            MsgBox("Error al limpiar la tabla Tempapu")
            MsgBox(ex.ToString)
        End Try

        'Ordenamos la columna Concepto, antes de calcular los totales parciales.
        '***********************************************************************
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
            DgvApuntesContables(3, 4)
        End If

        'Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
        '***********************************************************************
        vNombreConcepto = ""
        If vGrafico <> "" Then
            For Each fila As DataGridViewRow In frmApuntesPeriodicos.DgvApuper.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = ""
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMySql1cr.CommandText = vAñadir
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al Grabar el Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMySql1cr.CommandType = CommandType.Text
                    cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        If drMySql1.HasRows Then
                            While drMySql1.Read()
                                vExistenteImporteConcepto = drMySql1.GetValue(1)
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
                        End If
                        drMySql1.Close()
                    Catch ex As Exception
                        MsgBox("Error al verificar el Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMySql1cr.CommandText = vAñadir2
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al actualizar el Importe del Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    drMySql1.Close()
                End If
            Next
        Else
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = ""
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMySql1cr.CommandText = vAñadir
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al Grabar el Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMySql1cr.CommandType = CommandType.Text
                    cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        If drMySql1.HasRows Then
                            While drMySql1.Read()
                                vExistenteImporteConcepto = drMySql1.GetValue(1)
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
                        End If
                        drMySql1.Close()
                    Catch ex As Exception
                        MsgBox("Error al verificar el Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMySql1cr.CommandText = vAñadir2
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al actualizar el Importe del Concepto en Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    drMySql1.Close()
                End If
            Next
        End If

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
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        If TsBtnPastel.Checked Then
            Chart1.Series("Gastos").Points.Clear()
            'Enviamos a un dataview los datos
            For x = 0 To miView.Count - 1
                'Tomamos los datos de DataView para la gráfica
                With Chart1.Series("Gastos")
                    If miView(x)("Importe") <= 0 Then
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    Else
                        Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                    End If
                    .ChartType = SeriesChartType.Pie
                End With
                With Chart1.Series("Ingresos")
                    .ChartType = SeriesChartType.Pie
                End With
            Next
        Else
            Chart1.Series("Gastos").XValueMember = "Concepto"
            Chart1.Series("Ingresos").YValueMembers = "Importe"

            Chart1.Series("Gastos").Points.Clear()
            For x = 0 To miView.Count - 1
                'Tomamos los datos de DataView para la gráfica
                With Chart1.Series("Gastos")
                    If miView(x)("Importe") <= 0 Then
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                    Else
                        Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                        .Points(i).Color = Color.Blue
                    End If
                    .ChartType = SeriesChartType.Column
                End With
                With Chart1.Series("Ingresos")
                    .ChartType = SeriesChartType.Column
                End With
            Next
        End If
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
        e.Graphics.DrawString(frmGraficosConceptos3D.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        b = New Bitmap(frmGraficosConceptos3D.Chart1.Width, frmGraficosConceptos3D.Chart1.Height)
        frmGraficosConceptos3D.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))
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
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"
        Chart1.Series("Gastos").Points.Clear()

        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
                .ChartType = SeriesChartType.Column
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Column
            End With
        Next
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").Points.Clear()
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
                .ChartType = SeriesChartType.Area
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Area
            End With
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True
        TsBtnPastel.Checked = False
        Chart1.Series("Gastos").XValueMember = "Concepto"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").Points.Clear()
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                    .Points(i).Color = Color.Red
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                    .Points(i).Color = Color.Blue
                End If
                .ChartType = SeriesChartType.Line
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Line
            End With
        Next
    End Sub

    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = True
        Chart1.Series("Gastos").Points.Clear()
        'Enviamos a un dataview los datos
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Gastos")
                If miView(x)("Importe") <= 0 Then
                    vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
                End If
                .ChartType = SeriesChartType.Pie
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Pie
            End With
        Next
    End Sub

    'Dim ChartArea2 As New ChartArea()
    'Dim Legend2 As New Legend()
    'Dim Series2 As New Series()
    'Dim Chart2 = New Chart()
    'Me.Controls.Add(chart2)

    'ChartArea2.Name = "ChartArea2"
    'Chart2.ChartAreas.Add(ChartArea2)
    'Legend2.Name = "Legend2"
    'Chart2.Legends.Add(Legend2)
    'Chart2.Location = New Point(0, 31)
    'Chart2.Name = "Chart2"
    'Series2.ChartArea = "ChartArea2"
    'Series2.Legend = "Legend2"
    'Series2.Name = "Series2"
    'Series2.ChartType = SeriesChartType.Column
    'Chart2.Series.Add(Series2)
    'Chart2.Size = New System.Drawing.Size(1360, 604)
    'Chart2.TabIndex = 0
    'Chart2.Text = "Evaluación Chart"

    'Chart2.Series("Series2").XValueMember = "Concepto"
    'Chart2.Series("Series2").YValueMembers = "Importe"
    'Chart2.Series("Series2").Points.Clear()


    'For x = 0 To miView.Count - 1
    '    'Tomamos los datos de DataView para la gráfica
    '    With Chart2.Series("Series2")
    '        Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), miView(x)("Importe"))
    '        .Points(i).Color = Color.Blue
    '    End With
    'Next

End Class