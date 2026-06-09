Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosFechas

    Public Property EsGrafico3D As Boolean = False
    Public vAñadir, vAñadir2, vPositivo As String
    Public miDataTable As New DataTable
    Public miView As New DataView(miDataTable)
    Public x, vContador As Integer
    Public vImportePrimero, vImporteSegundo, vImporteConcepto, vNewImporteConcepto, vExistenteImporteConcepto As Double
    Private b As Bitmap
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub GraficosCuentas_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'Iniciamos Tabla tmpprint
        '************************
        LimpiarTempPrint()

        'Ordenamos la columna Fecha, antes de calcular los totales parciales.
        '***********************************************************************
        If vGrafico <> "" Then
            frmApuntesPeriodicos.DgvApuper.Sort(frmApuntesPeriodicos.DgvApuper.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        Else
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
        End If

        'Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
        '***********************************************************************
        If vGrafico <> "" Then
            'Viene de Apuntes Periódicos
            LlenarTempApuFechas("FECHAS_APUNTES_PERIODICOS")
        Else
            ' Viene de Apuntes Contables
            LlenarTempApuFechas("FECHAS_APUNTES_CONTABLES")
        End If

        miDataTable.Columns.Add("Fecha")
        miDataTable.Columns.Add("Importe")
        Dim unused As DataRow = miDataTable.NewRow()
        vtipoSql = "SELECT * FROM tmpprint ORDER BY tmpprint.FechaTMP ASC"
        LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        vValor = 0
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            'Guardamos los datos en un database
            Dim Renglon As DataRow = miDataTable.NewRow()
            Renglon("Fecha") = Mid(fila.Cells(0).Value, 1, 10).ToString
            vValor = fila.Cells(5).Value
            vValor = Math.Truncate(vValor)
            Renglon("Importe") = vValor
            miDataTable.Rows.Add(Renglon)
        Next

        DibujarGraficoColumnas()
    End Sub

    Private Sub DibujarGraficoColumnas()
        ' 1. Aplicamos los estilos e internacionalización de leyendas
        CrearEstilos()

        ' 2. Limpieza obligatoria de puntos previos
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        vContador = 0

        For x = 0 To miView.Count - 1
            vContador += 1

            ' Conversión segura del importe actual (Elemento X)
            Dim importeActual As Decimal = 0.0D
            If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                Decimal.TryParse(miView(x)("Importe").ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importeActual)
            End If

            vImporteConcepto = importeActual

            ' Lógica de comparación de impares (Verificando que exista un elemento siguiente)
            If (vContador Mod 2) <> 0 Then
                vImportePrimero = vImporteConcepto

                ' Protección contra errores de desbordamiento de índice (Evita que estalle al final de la tabla)
                If x + 1 < miView.Count Then
                    Dim importeSiguiente As Decimal = 0.0D
                    If miView(x + 1)("Importe") IsNot DBNull.Value AndAlso miView(x + 1)("Importe") IsNot Nothing Then
                        Decimal.TryParse(miView(x + 1)("Importe").ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importeSiguiente)
                    End If
                    vImporteSegundo = importeSiguiente
                Else
                    vImporteSegundo = 0 ' Si es el último registro impar aislado, el segundo se asume como 0
                End If

                ' --- Evaluaciones de Impares ---
                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                End If
            Else
                ' --- Evaluaciones de Pares ---
                vImporteSegundo = vImporteConcepto

                If vImportePrimero = 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                ElseIf vImportePrimero = 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                ElseIf vImportePrimero > 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                ElseIf vImportePrimero < 0 And vImporteSegundo = 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                ElseIf vImportePrimero < 0 And vImporteSegundo > 0 Then
                    AñadirPuntoGrafico("Ingresos", miView(x)("Fecha"), vImporteConcepto, Color.Blue)
                ElseIf vImportePrimero > 0 And vImporteSegundo < 0 Then
                    AñadirPuntoGrafico("Gastos", miView(x)("Fecha"), vImporteConcepto, Color.Red)
                End If
            End If
        Next
    End Sub

    ' Función auxiliar para compactar tu código y que no quede repetitivo
    Private Sub AñadirPuntoGrafico(nombreSerie As String, fecha As Object, valor As Decimal, colorPunto As Color)
        With Chart1.Series(nombreSerie)
            Dim i As Integer = .Points.AddXY(fecha, Math.Abs(valor))
            .Points(i).Color = colorPunto
            .ChartType = SeriesChartType.Column
        End With
    End Sub

    Public Sub CrearEstilos()
        ' 1. Configuración de fuentes y títulos
        Dim fuenteEjes As New Font("Arial", 12, FontStyle.Bold)
        Chart1.ChartAreas("ChartArea1").AxisX.TitleFont = fuenteEjes
        Chart1.ChartAreas("ChartArea1").AxisY.TitleFont = fuenteEjes

        If Chart1.Titles.Count > 0 Then
            Chart1.Titles(0).Text = rmse.GetString("TituloGrafico")
        End If

        Chart1.ChartAreas("ChartArea1").AxisX.Title = resManager.GetString("Fecha")
        Chart1.ChartAreas("ChartArea1").AxisY.Title = resManager.GetString("Moneda") & ": " & vMoneda

        ' 2. Forzar visibilidad en la leyenda antes de traducir
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True

        Chart1.Series("Gastos").LegendText = resManager.GetString("Gastos")
        Chart1.Series("Ingresos").LegendText = resManager.GetString("Ingresos")

        ' 3. Mapeo de miembros de datos de la serie
        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"
        Chart1.ChartAreas("ChartArea1").Area3DStyle.Enable3D = Me.EsGrafico3D

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

        CrearEstilos()

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

    Public Sub CrearEstilosPastelFechas()
        ' Limpieza obligatoria de los títulos de los ejes para el modo Pastel
        Chart1.ChartAreas("ChartArea1").AxisX.Title = ""
        Chart1.ChartAreas("ChartArea1").AxisY.Title = ""

        ' Forzar visibilidad en leyenda y mapear dinámicamente con las etiquetas de las fechas (#VALX)
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True
        Chart1.Series("Gastos").LegendText = "#VALX"
        Chart1.Series("Ingresos").LegendText = "#VALX"

        ' Mapeo de miembros de datos
        Chart1.Series("Gastos").XValueMember = "Fecha"
        Chart1.Series("Gastos").YValueMembers = "Importe"
        Chart1.Series("Ingresos").XValueMember = "Fecha"
        Chart1.Series("Ingresos").YValueMembers = "Importe"
    End Sub

    Private Sub DibujarGraficoPastelFechas()
        ' 1. Aplicamos los estilos y mapeos específicos
        CrearEstilosPastelFechas()

        ' 2. Limpieza estricta de puntos previos
        Chart1.Series("Gastos").Points.Clear()
        Chart1.Series("Ingresos").Points.Clear()

        ' 3. Recorrido seguro de la vista
        For x = 0 To miView.Count - 1
            With Chart1.Series("Gastos")
                ' Conversión segura multiidioma del importe
                Dim importePuro As Decimal = 0.0D
                If miView(x)("Importe") IsNot DBNull.Value AndAlso miView(x)("Importe") IsNot Nothing Then
                    Dim textoImporte As String = miView(x)("Importe").ToString()
                    If Not Decimal.TryParse(textoImporte, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, importePuro) Then
                        Decimal.TryParse(textoImporte.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importePuro)
                    End If
                End If

                If importePuro <= 0 Then
                    vImporteConcepto = Math.Abs(importePuro)
                    .Points.AddXY(miView(x)("Fecha"), vImporteConcepto)
                Else
                    .Points.AddXY(miView(x)("Fecha"), importePuro)
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
        DibujarGraficoPastelFechas()
    End Sub

End Class