Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosPresupuestos

    Public vAñadir, vAñadir2, vTempapu, vTmpprint, vPositivo As String
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImportePrimero, vImporteSegundo, vImporteConcepto, vNewImporteConcepto, vImporteConcepto2, vExistenteImporteConcepto As Double
    Private b As Bitmap

    Private Sub GraficosPresupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        'Iniciamos Tabla Tmpprint
        '************************
        vTmpprint = "DELETE FROM tmpprint"
        cmdMySql1cr.CommandText = vTmpprint
        Try
            cmdMySql1cr.ExecuteNonQuery()
            'MsgBox("Registros Tmpprint, Borrados !!!")
        Catch ex As Exception
            MsgBox("Error al borrar los registros de Tmpprint")
            MsgBox(ex.ToString)
        End Try

        'Ordenamos la columna Concepto, antes de calcular los totales parciales.
        '***********************************************************************
        frmPresupuestos.DgvPresupuestos.Sort(frmPresupuestos.DgvPresupuestos.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

        'Llenamos la tabla tmpprint con los Conceptos Agrupados desde DgvPresupuestos
        '****************************************************************************
        vNombreConcepto = ""
        For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
            vImporteConcepto = fila.Cells(3).Value
            If vNombreConcepto <> fila.Cells(0).Value.ToString Then
                vNombreConcepto = fila.Cells(0).Value.ToString
                vImporteConcepto = 0
                vImporteConcepto = fila.Cells(3).Value

                'Buscamos la suma total del Concepto
                '***********************************
                vtipoSql = "SELECT * FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                vtipoSql += " And apuntes.ConceptoAPU = '" & vNombreConcepto & "' "
                vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC"
                LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")  'Agrupado Conceptos

                vImporteConcepto2 = "0,00"
                For Each filas As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                    vImporteConcepto2 += Math.Abs(filas.Cells(4).Value)
                Next
                vAñadir = "INSERT INTO tmpprint"
                vAñadir += "(FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) "
                vAñadir += "VALUES ('01/01/1900', '" & vNombreConcepto & "', '', '', '' , '" & vImporteConcepto2 & "', '" & vImporteConcepto & "')"
                cmdMySql1cr.CommandText = vAñadir
                Try
                    cmdMySql1cr.ExecuteNonQuery()
                    'MsgBox("Registro1, Grabado Correctamente " & vNombreConcepto)
                Catch ex As Exception
                    MsgBox("Error al grabar el Concepto en Tmpprint")
                    MsgBox(ex.ToString)
                End Try
            Else
                cmdMySql1cr.CommandType = CommandType.Text
                cmdMySql1cr.CommandText = "SELECT * FROM tmpprint WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                Try
                    drMySql1 = cmdMySql1cr.ExecuteReader()
                    If drMySql1.HasRows Then
                        While drMySql1.Read()
                            vExistenteImporteConcepto = drMySql1.GetValue(6)
                        End While
                    Else
                        'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
                    End If
                    drMySql1.Close()
                Catch ex As Exception
                    MsgBox("Error al buscar el Importe del Concepto en Tempprint")
                    MsgBox(ex.ToString)
                End Try
                vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                vAñadir2 = "UPDATE tmpprint SET SaldoTMP = '" & vNewImporteConcepto & "' "
                vAñadir2 += " WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                cmdMySql1cr.CommandText = vAñadir2
                Try
                    drMySql1 = cmdMySql1cr.ExecuteReader()
                    'MsgBox(vImporteConcepto & " Registro2, Grabado Correctamente " & vNombreConcepto & " " & vNewImporteConcepto)
                Catch ex As Exception
                    MsgBox("Error al actualizar el Importe del Concepto en tmpprint")
                    MsgBox(ex.ToString)
                End Try
                drMySql1.Close()
            End If
        Next

        vtipoSql = "SELECT * FROM tmpprint"
        vtipoSql += " ORDER BY tmpprint.ConceptoTMP ASC"
        LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")

        miDataTable.Columns.Add("Concepto")
        miDataTable.Columns.Add("Real")
        miDataTable.Columns.Add("Presupuestado")

        Dim unused As DataRow = miDataTable.NewRow()
        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            'Guardamos los datos en un database
            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Concepto") = fila.Cells(1).Value.ToString
            vValor = fila.Cells(5).Value
            vValor = Math.Truncate(vValor)
            Renglon("Real") = vValor.ToString
            vValor = fila.Cells(6).Value
            vValor = Math.Truncate(vValor)
            Renglon("Presupuestado") = vValor.ToString
            miDataTable.Rows.Add(Renglon)
        Next
        Chart1.Series("Real").IsVisibleInLegend = True
        Chart1.Series("Presupuestado").IsVisibleInLegend = True

        Chart1.Series("Real").XValueMember = "Concepto"
        Chart1.Series("Real").YValueMembers = "Real"
        Chart1.Series("Presupuestado").XValueMember = "Concepto"
        Chart1.Series("Presupuestado").YValueMembers = "Presupuestado"

        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Real")
                vImporteConcepto = Math.Abs(Val(miView(x)("Real")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Red
                .ChartType = SeriesChartType.Column
            End With
            With Chart1.Series("Presupuestado")
                vImporteConcepto = Math.Abs(Val(miView(x)("Presupuestado")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Blue
                .ChartType = SeriesChartType.Column
            End With
        Next
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
        e.Graphics.DrawString(frmGraficosPresupuestos.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        b = New Bitmap(frmGraficosPresupuestos.Chart1.Width, frmGraficosPresupuestos.Chart1.Height)
        frmGraficosPresupuestos.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))
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

        Chart1.Series("Real").IsVisibleInLegend = True
        Chart1.Series("Presupuestado").IsVisibleInLegend = True

        Chart1.Series("Real").XValueMember = "Concepto"
        Chart1.Series("Real").YValueMembers = "Real"
        Chart1.Series("Presupuestado").XValueMember = "Concepto"
        Chart1.Series("Presupuestado").YValueMembers = "Presupuestado"

        Chart1.Series("Real").Points.Clear()
        Chart1.Series("Presupuestado").Points.Clear()

        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Real")
                vImporteConcepto = Math.Abs(Val(miView(x)("Real")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Red
                .ChartType = SeriesChartType.Column
            End With
            With Chart1.Series("Presupuestado")
                vImporteConcepto = Math.Abs(Val(miView(x)("Presupuestado")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Blue
                .ChartType = SeriesChartType.Column
            End With
        Next
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False

        Chart1.Series("Real").IsVisibleInLegend = True
        Chart1.Series("Presupuestado").IsVisibleInLegend = True

        Chart1.Series("Real").XValueMember = "Concepto"
        Chart1.Series("Real").YValueMembers = "Real"
        Chart1.Series("Presupuestado").XValueMember = "Concepto"
        Chart1.Series("Presupuestado").YValueMembers = "Presupuestado"

        Chart1.Series("Real").Points.Clear()
        Chart1.Series("Presupuestado").Points.Clear()

        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Real")
                vImporteConcepto = Math.Abs(Val(miView(x)("Real")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Red
                .ChartType = SeriesChartType.Area
            End With
            With Chart1.Series("Presupuestado")
                vImporteConcepto = Math.Abs(Val(miView(x)("Presupuestado")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Blue
                .ChartType = SeriesChartType.Area
            End With
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True

        Chart1.Series("Real").IsVisibleInLegend = True
        Chart1.Series("Presupuestado").IsVisibleInLegend = True

        Chart1.Series("Real").XValueMember = "Concepto"
        Chart1.Series("Real").YValueMembers = "Real"
        Chart1.Series("Presupuestado").XValueMember = "Concepto"
        Chart1.Series("Presupuestado").YValueMembers = "Presupuestado"

        Chart1.Series("Real").Points.Clear()
        Chart1.Series("Presupuestado").Points.Clear()

        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            With Chart1.Series("Real")
                vImporteConcepto = Math.Abs(Val(miView(x)("Real")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Red
                .ChartType = SeriesChartType.Line
            End With
            With Chart1.Series("Presupuestado")
                vImporteConcepto = Math.Abs(Val(miView(x)("Presupuestado")))
                Dim i As Integer = .Points.AddXY(miView(x)("Concepto"), vImporteConcepto)
                .Points(i).Color = Color.Blue
                .ChartType = SeriesChartType.Line
            End With
        Next
    End Sub
End Class