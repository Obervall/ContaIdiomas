Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosMeses

    Public vAñadir, vAñadir2, vTempapu, vImporteConcepto, vNewImporteConcepto As String
    Public vExistenteImporteConcepto, vPositivo As String
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImportePrimero, vImporteSegundo As Double
    Private b As Bitmap

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
        e.Graphics.DrawString(frmGraficosMeses.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        b = New Bitmap(frmGraficosMeses.Chart1.Width, frmGraficosMeses.Chart1.Height)
        frmGraficosMeses.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))
        e.Graphics.DrawImage(b, 0, 100)

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

    Private Sub GraficosCuentas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        'Iniciamos Tabla Tempapu
        '***********************
        vTempapu = "DELETE FROM tempapu"
        cmdMySql1cr.CommandText = vTempapu
        Try
            cmdMySql1cr.ExecuteNonQuery()
            'MsgBox("Registros Tempapu, Borrados !!!")
        Catch ex As Exception
            MsgBox("Error al borrar los registros de Tempapu")
            MsgBox(ex.ToString)
        End Try

        'Ordenamos la columna Fecha, antes de calcular los totales parciales.
        '***********************************************************************
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        End If

        'Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
        '***********************************************************************
        vNombreConcepto = ""
        If vGrafico <> "" Then
            For Each fila As DataGridViewRow In frmApuntesPeriodicos.DgvApuper.Rows
                If fila.Cells(3).Value <> 0 Then
                    vImporteConcepto = fila.Cells(3).Value
                    If vNombreConcepto <> Mid(fila.Cells(0).Value, 9, 10) & "-" & Mid(fila.Cells(0).Value, 4, 2).ToString Then
                        vNombreConcepto = Mid(fila.Cells(0).Value, 9, 10) & "-" & Mid(fila.Cells(0).Value, 4, 2).ToString
                        vImporteConcepto = ""
                        vImporteConcepto = fila.Cells(3).Value
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                        cmdMySql1cr.CommandText = vAñadir
                        Try
                            cmdMySql1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "',' 0 ')"
                        cmdMySql1cr.CommandText = vAñadir
                        Try
                            cmdMySql1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                    Else ' Si el Concepto existe y hay importe diferente a cero, si es positivo o negativo se suma
                        cmdMySql1cr.CommandType = CommandType.Text
                        If Val(vImporteConcepto) > 0 Then
                            cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                            cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU > 0 "
                        ElseIf Val(vImporteConcepto) < 0 Then
                            cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                            cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU < 0 "
                        End If
                        Try
                            drMySql1 = cmdMySql1cr.ExecuteReader()
                            If drMySql1.HasRows Then 'Significa que existe con las condiciones
                                While drMySql1.Read()
                                    vExistenteImporteConcepto = drMySql1.GetValue(1)
                                End While
                                drMySql1.Close()
                                vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                If Val(vImporteConcepto) > 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU > 0 "
                                ElseIf Val(vImporteConcepto) < 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU < 0 "
                                End If
                                cmdMySql1cr.CommandText = vAñadir2
                                Try
                                    drMySql1 = cmdMySql1cr.ExecuteReader()
                                Catch ex As Exception
                                    MsgBox("Error al actualizar el Concepto en Tempapu")
                                    MsgBox(ex.ToString)
                                End Try
                                drMySql1.Close()

                            Else   'NO existe, lo añadimos al cero
                                'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
                                drMySql1.Close()
                                cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU = 0 "
                                drMySql1 = cmdMySql1cr.ExecuteReader()
                                If drMySql1.HasRows Then 'Significa que existe con las condiciones
                                    While drMySql1.Read()
                                        vExistenteImporteConcepto = drMySql1.GetValue(1)
                                    End While
                                    drMySql1.Close()
                                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU = 0 "
                                    cmdMySql1cr.CommandText = vAñadir2
                                    Try
                                        drMySql1 = cmdMySql1cr.ExecuteReader()
                                    Catch ex As Exception
                                        MsgBox("Error al actualizar el Concepto en Tempapu")
                                        MsgBox(ex.ToString)
                                    End Try
                                    drMySql1.Close()
                                End If
                                drMySql1.Close()
                            End If
                        Catch ex As Exception
                            MsgBox("Error al verificar que el Concepto existe en Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                    End If
                End If
            Next
        Else
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                If fila.Cells(3).Value <> 0 Then
                    vImporteConcepto = fila.Cells(3).Value
                    If vNombreConcepto <> Mid(fila.Cells(0).Value, 9, 10) & "-" & Mid(fila.Cells(0).Value, 4, 2).ToString Then
                        vNombreConcepto = Mid(fila.Cells(0).Value, 9, 10) & "-" & Mid(fila.Cells(0).Value, 4, 2).ToString
                        vImporteConcepto = ""
                        vImporteConcepto = fila.Cells(3).Value
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                        cmdMySql1cr.CommandText = vAñadir
                        Try
                            cmdMySql1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "',' 0 ')"
                        cmdMySql1cr.CommandText = vAñadir
                        Try
                            cmdMySql1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                    Else ' Si el Concepto existe y hay importe diferente a cero, si es positivo o negativo se suma
                        cmdMySql1cr.CommandType = CommandType.Text
                        If Val(vImporteConcepto) > 0 Then
                            cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                            cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU > 0 "
                        ElseIf Val(vImporteConcepto) < 0 Then
                            cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                            cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU < 0 "
                        End If
                        Try
                            drMySql1 = cmdMySql1cr.ExecuteReader()
                            If drMySql1.HasRows Then 'Significa que existe con las condiciones
                                While drMySql1.Read()
                                    vExistenteImporteConcepto = drMySql1.GetValue(1)
                                End While
                                drMySql1.Close()
                                vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                If Val(vImporteConcepto) > 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU > 0 "
                                ElseIf Val(vImporteConcepto) < 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU < 0 "
                                End If
                                cmdMySql1cr.CommandText = vAñadir2
                                Try
                                    drMySql1 = cmdMySql1cr.ExecuteReader()
                                Catch ex As Exception
                                    MsgBox("Error al actualizar el Concepto en Tempapu")
                                    MsgBox(ex.ToString)
                                End Try
                                drMySql1.Close()

                            Else   'NO existe, lo añadimos al cero
                                'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
                                drMySql1.Close()
                                cmdMySql1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                cmdMySql1cr.CommandText += "And tempapu.SumaImporteAPU = 0 "
                                drMySql1 = cmdMySql1cr.ExecuteReader()
                                If drMySql1.HasRows Then 'Significa que existe con las condiciones
                                    While drMySql1.Read()
                                        vExistenteImporteConcepto = drMySql1.GetValue(1)
                                    End While
                                    drMySql1.Close()
                                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU = 0 "
                                    cmdMySql1cr.CommandText = vAñadir2
                                    Try
                                        drMySql1 = cmdMySql1cr.ExecuteReader()
                                    Catch ex As Exception
                                        MsgBox("Error al actualizar el Concepto en Tempapu")
                                        MsgBox(ex.ToString)
                                    End Try
                                    drMySql1.Close()
                                End If
                                drMySql1.Close()
                            End If
                        Catch ex As Exception
                            MsgBox("Error al verificar que el Concepto existe en Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                    End If
                End If
            Next
        End If

        miDataTable.Columns.Add("Fecha")
        miDataTable.Columns.Add("Importe")
        Dim unused As DataRow = miDataTable.NewRow()
        vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            'Guardamos los datos en un database
            Dim Renglon As DataRow = miDataTable.NewRow()
            If Mid(fila.Cells(0).Value, 4, 5) = "01" Then
                Renglon("Fecha") = "Enero-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "02" Then
                Renglon("Fecha") = "Febrero-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "03" Then
                Renglon("Fecha") = "Marzo-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "04" Then
                Renglon("Fecha") = "Abril-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "05" Then
                Renglon("Fecha") = "Mayo-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "06" Then
                Renglon("Fecha") = "Junio-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "07" Then
                Renglon("Fecha") = "Julio-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "08" Then
                Renglon("Fecha") = "Agosto-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "09" Then
                Renglon("Fecha") = "Septiembre-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "10" Then
                Renglon("Fecha") = "Octubre-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "11" Then
                Renglon("Fecha") = "Noviembre-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            If Mid(fila.Cells(0).Value, 4, 5) = "12" Then
                Renglon("Fecha") = "Diciembre-" & Mid(fila.Cells(0).Value, 1, 2)
            End If
            vValor = fila.Cells(1).Value
            vValor = Math.Truncate(vValor)
            Renglon("Importe") = vValor
            miDataTable.Rows.Add(Renglon)
        Next
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True

        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        vContador = 0
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            vContador += 1
            vImporteConcepto = Val(miView(x)("Importe"))
            If (vContador Mod 2) <> 0 Then
                'El número es impar.
                vImportePrimero = Val(miView(x)("Importe"))
                vImporteSegundo = Val(miView(x + 1)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
            Else
                'El número es par.
                vImporteSegundo = Val(miView(x)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
            End If
        Next
    End Sub

    Private Sub TsBtnColumnas_Click(sender As Object, e As EventArgs) Handles TsBtnColumnas.Click
        TsBtnColumnas.Checked = True
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False
        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            vContador += 1
            vImporteConcepto = Val(miView(x)("Importe"))
            If (vContador Mod 2) <> 0 Then
                'El número es impar.
                vImportePrimero = Val(miView(x)("Importe"))
                vImporteSegundo = Val(miView(x + 1)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
            Else
                'El número es par.
                vImporteSegundo = Val(miView(x)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
            End If
        Next
    End Sub

    Private Sub TsBtnAreas_Click(sender As Object, e As EventArgs) Handles TsBtnAreas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = True
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = False
        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            vContador += 1
            vImporteConcepto = Val(miView(x)("Importe"))
            If (vContador Mod 2) <> 0 Then
                'El número es impar.
                vImportePrimero = Val(miView(x)("Importe"))
                vImporteSegundo = Val(miView(x + 1)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
            Else
                'El número es par.
                vImporteSegundo = Val(miView(x)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
            End If
        Next
    End Sub

    Private Sub TsBtnLineas_Click(sender As Object, e As EventArgs) Handles TsBtnLineas.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = True
        TsBtnPastel.Checked = False
        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
            vContador += 1
            vImporteConcepto = Val(miView(x)("Importe"))
            If (vContador Mod 2) <> 0 Then
                'El número es impar.
                vImportePrimero = Val(miView(x)("Importe"))
                vImporteSegundo = Val(miView(x + 1)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
            Else
                'El número es par.
                vImporteSegundo = Val(miView(x)("Importe"))
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
            End If
        Next
    End Sub

    Private Sub TsBtnPastel_Click(sender As Object, e As EventArgs) Handles TsBtnPastel.Click
        TsBtnColumnas.Checked = False
        TsBtnAreas.Checked = False
        TsBtnLineas.Checked = False
        TsBtnPastel.Checked = True

        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()
        For x = 0 To miView.Count - 1
            'Tomamos los datos de DataView para la gráfica
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
End Class