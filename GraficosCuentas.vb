Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports System.Windows.Forms.DataVisualization.Charting

Public Class GraficosCuentas

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
        e.Graphics.DrawString(frmGraficosCuentas.Chart1.Titles.Item(0).Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        b = New Bitmap(frmGraficosCuentas.Chart1.Width, frmGraficosCuentas.Chart1.Height)
        frmGraficosCuentas.Chart1.DrawToBitmap(b, New Rectangle(0, 0, b.Width, b.Height))
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
        cmdMdb1cr.CommandText = vTempapu
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Registros Tempapu, Borrados !!!")
        Catch ex As Exception
            MsgBox("Error al borrar los registros de Tempapu")
            MsgBox(ex.ToString)
        End Try

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
            For Each fila As DataGridViewRow In frmApuntesPeriodicos.DgvApuper.Rows
                If fila.Cells(3).Value <> 0 Then
                    vImporteConcepto = fila.Cells(3).Value
                    If vNombreConcepto <> fila.Cells(6).Value.ToString Then
                        vNombreConcepto = fila.Cells(6).Value.ToString
                        vImporteConcepto = ""
                        vImporteConcepto = fila.Cells(3).Value
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "',' 0 ')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu")
                            MsgBox(ex.ToString)
                        End Try
                    Else ' Si el Concepto existe y hay importe diferente a cero, si es positivo o negativo se suma
                        cmdMdb1cr.CommandType = CommandType.Text
                        ' 1. Convertimos el importe a Decimal de forma segura (multiidioma)
                        Dim importeDecimal As Decimal = 0.0D
                        If vImporteConcepto IsNot Nothing Then
                            Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."),
                     System.Globalization.NumberStyles.Any,
                     System.Globalization.CultureInfo.InvariantCulture,
                     importeDecimal)
                        End If

                        ' 2. Limpiamos los parámetros previos del comando
                        cmdMdb1cr.Parameters.Clear()

                        ' 3. Evaluamos de forma exacta usando el número decimal puro
                        If importeDecimal > 0 Then
                            ' Consulta usando parámetros (?) para evitar fallos por comillas o caracteres raros
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU > 0"

                            ' Añadimos el parámetro que sustituye al primer "?"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto.ToString())

                        ElseIf importeDecimal < 0 Then
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU < 0"

                            ' Añadimos el parámetro que sustituye al primer "?"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto.ToString())
                        End If

                        'If Val(vImporteConcepto) > 0 Then
                        '    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                        '    cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU > 0 "
                        'ElseIf Val(vImporteConcepto) < 0 Then
                        '    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                        '    cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU < 0 "
                        'End If
                        Try
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then 'Significa que existe con las condiciones
                                While drMdb1.Read()
                                    vExistenteImporteConcepto = drMdb1.GetValue(1)
                                End While
                                drMdb1.Close()
                                'vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                ' 1. Convertimos ambos importes a variables decimales exactas
                                Dim importe1 As Decimal = 0.0D
                                Dim importe2 As Decimal = 0.0D

                                ' Conversión segura del primer importe
                                If vImporteConcepto IsNot Nothing Then
                                    Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe1)
                                End If

                                ' Conversión segura del segundo importe
                                If vExistenteImporteConcepto IsNot Nothing Then
                                    Decimal.TryParse(vExistenteImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe2)
                                End If

                                ' 2. Sumamos los números reales de forma exacta
                                vNewImporteConcepto = importe1 + importe2

                                If importe1 > 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU > 0 "
                                ElseIf importe1 < 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU < 0 "
                                End If
                                cmdMdb1cr.CommandText = vAñadir2
                                Try
                                    drMdb1 = cmdMdb1cr.ExecuteReader()
                                Catch ex As Exception
                                    MsgBox("Error al actualizar el Concepto en Tempapu con las condiciones")
                                    MsgBox(ex.ToString)
                                End Try
                                drMdb1.Close()

                            Else   'NO existe, lo añadimos al cero
                                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                                drMdb1.Close()
                                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU = 0 "
                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then 'Significa que existe con las condiciones
                                    While drMdb1.Read()
                                        vExistenteImporteConcepto = drMdb1.GetValue(1)
                                    End While
                                    drMdb1.Close()
                                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                    ' 1. Convertimos ambos importes a variables decimales exactas
                                    Dim importe1 As Decimal = 0.0D
                                    Dim importe2 As Decimal = 0.0D

                                    ' Conversión segura del primer importe
                                    If vImporteConcepto IsNot Nothing Then
                                        Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe1)
                                    End If

                                    ' Conversión segura del segundo importe
                                    If vExistenteImporteConcepto IsNot Nothing Then
                                        Decimal.TryParse(vExistenteImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe2)
                                    End If

                                    ' 2. Sumamos los números reales de forma exacta
                                    vNewImporteConcepto = importe1 + importe2

                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU = 0 "
                                    cmdMdb1cr.CommandText = vAñadir2
                                    Try
                                        drMdb1 = cmdMdb1cr.ExecuteReader()
                                    Catch ex As Exception
                                        MsgBox("Error al actualizar el Concepto en Tempapu con las condiciones")
                                        MsgBox(ex.ToString)
                                    End Try
                                    drMdb1.Close()
                                End If
                                drMdb1.Close()
                            End If
                        Catch ex As Exception
                            'MsgBox("Error al verificar que el Concepto existe en Tempapu con las condiciones")
                            MsgBox(ex.ToString)
                        End Try
                    End If
                End If
            Next
        Else
            'For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
            '    If fila.Cells(3).Value <> 0 Then
            '        vImporteConcepto = fila.Cells(3).Value
            '        If vNombreConcepto <> fila.Cells(6).Value.ToString Then
            '            vNombreConcepto = fila.Cells(6).Value.ToString
            '            vImporteConcepto = ""
            '            vImporteConcepto = fila.Cells(3).Value
            '            vAñadir = "INSERT INTO tempapu"
            '            vAñadir += "(ConceptoAPU, SumaImporteAPU) "
            '            vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
            '            cmdMdb1cr.CommandText = vAñadir
            '            Try
            '                cmdMdb1cr.ExecuteNonQuery()
            '            Catch ex As Exception
            '                MsgBox("Error al añadir el Concepto a Tempapu")
            '                MsgBox(ex.ToString)
            '            End Try
            '            vAñadir = "INSERT INTO tempapu"
            '            vAñadir += "(ConceptoAPU, SumaImporteAPU) "
            '            vAñadir += "VALUES ('" & vNombreConcepto & "',' 0 ')"
            '            cmdMdb1cr.CommandText = vAñadir
            '            Try
            '                cmdMdb1cr.ExecuteNonQuery()
            '            Catch ex As Exception
            '                'MsgBox("Error al añadir el Concepto a Tempapu")
            '                MsgBox(ex.ToString)
            '            End Try
            '        Else ' Si el Concepto existe y hay importe diferente a cero, si es positivo o negativo se suma
            '            cmdMdb1cr.CommandType = CommandType.Text
            '            If Val(vImporteConcepto) > 0 Then
            '                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU > 0 "
            '            ElseIf Val(vImporteConcepto) < 0 Then
            '                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU < 0 "
            '            End If
            '            Try
            '                drMdb1 = cmdMdb1cr.ExecuteReader()
            '                If drMdb1.HasRows Then 'Significa que existe con las condiciones
            '                    While drMdb1.Read()
            '                        vExistenteImporteConcepto = drMdb1.GetValue(1)
            '                    End While
            '                    drMdb1.Close()
            '                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
            '                    If Val(vImporteConcepto) > 0 Then
            '                        vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
            '                        vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                        vAñadir2 += "And tempapu.SumaImporteAPU > 0 "
            '                    ElseIf Val(vImporteConcepto) < 0 Then
            '                        vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
            '                        vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                        vAñadir2 += "And tempapu.SumaImporteAPU < 0 "
            '                    End If
            '                    cmdMdb1cr.CommandText = vAñadir2
            '                    Try
            '                        drMdb1 = cmdMdb1cr.ExecuteReader()
            '                    Catch ex As Exception
            '                        MsgBox("Error al actualizar el Concepto en Tempapu con las condiciones")
            '                        MsgBox(ex.ToString)
            '                    End Try
            '                    drMdb1.Close()

            '                Else   'NO existe, lo añadimos al cero
            '                    'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            '                    drMdb1.Close()
            '                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                    cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU = 0 "
            '                    drMdb1 = cmdMdb1cr.ExecuteReader()
            '                    If drMdb1.HasRows Then 'Significa que existe con las condiciones
            '                        While drMdb1.Read()
            '                            vExistenteImporteConcepto = drMdb1.GetValue(1)
            '                        End While
            '                        drMdb1.Close()
            '                        vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
            '                        vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
            '                        vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
            '                        vAñadir2 += "And tempapu.SumaImporteAPU = 0 "
            '                        cmdMdb1cr.CommandText = vAñadir2
            '                        Try
            '                            drMdb1 = cmdMdb1cr.ExecuteReader()
            '                        Catch ex As Exception
            '                            MsgBox("Error al actualizar el Concepto en Tempapu con las condiciones")
            '                            MsgBox(ex.ToString)
            '                        End Try
            '                        drMdb1.Close()
            '                    End If
            '                    drMdb1.Close()
            '                End If
            '            Catch ex As Exception
            '                MsgBox("Error al verificar que el Concepto existe en Tempapu con las condiciones")
            '                MsgBox(ex.ToString)
            '            End Try
            '        End If
            '    End If
            'Next
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                ' Aseguramos que la fila no sea la fila vacía del final del DataGridView
                If fila.IsNewRow Then Continue For

                ' 1. Leemos el importe de la celda 3 de forma segura (como Decimal)
                Dim importeFila As Decimal = 0.0D
                If fila.Cells(3).Value IsNot Nothing AndAlso Not IsDBNull(fila.Cells(3).Value) Then
                    Decimal.TryParse(fila.Cells(3).Value.ToString().Replace(",", "."),
                         System.Globalization.NumberStyles.Any,
                         System.Globalization.CultureInfo.InvariantCulture,
                         importeFila)
                End If

                ' Si el importe es diferente de cero, procesamos
                If importeFila <> 0 Then

                    ' Leemos el nombre del concepto de la celda 6 de forma segura
                    Dim conceptoFila As String = ""
                    If fila.Cells(6).Value IsNot Nothing Then
                        conceptoFila = fila.Cells(6).Value.ToString()
                    End If

                    If vNombreConcepto <> conceptoFila Then
                        vNombreConcepto = conceptoFila

                        ' A) INSERT de la fila con su importe real
                        cmdMdb1cr.CommandText = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)
                        cmdMdb1cr.Parameters.AddWithValue("@Importe", importeFila) ' .NET maneja la coma/punto de forma nativa

                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox("Error al añadir el Concepto a Tempapu" & vbCrLf & ex.Message)
                        End Try

                        ' B) INSERT de la fila testigo con importe 0
                        cmdMdb1cr.CommandText = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, 0)"
                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.Message)
                        End Try

                    Else ' Si el concepto coincide, buscamos si ya existe para sumar
                        cmdMdb1cr.CommandType = CommandType.Text
                        cmdMdb1cr.Parameters.Clear()

                        If importeFila > 0 Then
                            cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE ConceptoAPU = ? And SumaImporteAPU > 0"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)
                        Else
                            cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE ConceptoAPU = ? And SumaImporteAPU < 0"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)
                        End If

                        Dim existeConCondicion As Boolean = False
                        Dim existenteImporte As Decimal = 0.0D

                        Try
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then
                                existeConCondicion = True
                                While drMdb1.Read()
                                    If Not drMdb1.IsDBNull(0) Then
                                        Decimal.TryParse(drMdb1.GetValue(0).ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, existenteImporte)
                                    End If
                                End While
                            End If
                            drMdb1.Close()

                            If existeConCondicion Then
                                ' Realizamos la suma matemática limpia en Decimal
                                Dim nuevaSuma As Decimal = importeFila + existenteImporte

                                cmdMdb1cr.Parameters.Clear()
                                If importeFila > 0 Then
                                    cmdMdb1cr.CommandText = "UPDATE tempapu SET SumaImporteAPU = ? WHERE ConceptoAPU = ? And SumaImporteAPU > 0"
                                Else
                                    cmdMdb1cr.CommandText = "UPDATE tempapu SET SumaImporteAPU = ? WHERE ConceptoAPU = ? And SumaImporteAPU < 0"
                                End If

                                ' IMPORTANTE EN ACCESS: Añadir parámetros en el orden exacto del SQL
                                cmdMdb1cr.Parameters.AddWithValue("@Suma", nuevaSuma)
                                cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                                Try
                                    cmdMdb1cr.ExecuteNonQuery() ' CORREGIDO: UPDATE usa ExecuteNonQuery, no ExecuteReader
                                Catch ex As Exception
                                    MsgBox("Error al actualizar el Concepto en Tempapu" & vbCrLf & ex.Message)
                                End Try

                            Else ' NO existe con esa condición, buscamos el registro con importe 0
                                cmdMdb1cr.Parameters.Clear()
                                cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE ConceptoAPU = ? And SumaImporteAPU = 0"
                                cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                                Dim existeCero As Boolean = False
                                Dim existenteImporteCero As Decimal = 0.0D

                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then
                                    existeCero = True
                                    While drMdb1.Read()
                                        If Not drMdb1.IsDBNull(0) Then
                                            Decimal.TryParse(drMdb1.GetValue(0).ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, existenteImporteCero)
                                        End If
                                    End While
                                End If
                                drMdb1.Close()

                                If existeCero Then
                                    Dim nuevaSumaCero As Decimal = importeFila + existenteImporteCero

                                    cmdMdb1cr.Parameters.Clear()
                                    cmdMdb1cr.CommandText = "UPDATE tempapu SET SumaImporteAPU = ? WHERE ConceptoAPU = ? And SumaImporteAPU = 0"
                                    cmdMdb1cr.Parameters.AddWithValue("@Suma", nuevaSumaCero)
                                    cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                                    Try
                                        cmdMdb1cr.ExecuteNonQuery() ' CORREGIDO: ExecuteNonQuery
                                    Catch ex As Exception
                                        MsgBox("Error al actualizar el Concepto cero en Tempapu" & vbCrLf & ex.Message)
                                    End Try
                                End If
                            End If

                        Catch ex As Exception
                            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                            MsgBox("Error en el proceso de verificación del Concepto" & vbCrLf & ex.Message)
                        End Try
                    End If
                End If
            Next

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
        Chart1.Series("Gastos").IsVisibleInLegend = True
        Chart1.Series("Ingresos").IsVisibleInLegend = True

        'Chart1.Series("Gastos").XValueMember = "Cuenta"
        'Chart1.Series("Ingresos").YValueMembers = "Importe"
        'Chart1.Series("Ingresos").XValueMember = "Cuenta"
        'Chart1.Series("Ingresos").YValueMembers = "Importe"

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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
        Chart1.Series("Gastos").XValueMember = "Cuenta"
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Column
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
        Chart1.Series("Gastos").XValueMember = "Cuenta"
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Area
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
        Chart1.Series("Gastos").XValueMember = "Cuenta"
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero = 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Red
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo = 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero < 0 And vImporteSegundo > 0 Then
                    With Chart1.Series("Ingresos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                        .Points(i).Color = Color.Blue
                        .ChartType = SeriesChartType.Line
                    End With
                End If
                If vImportePrimero > 0 And vImporteSegundo < 0 Then
                    With Chart1.Series("Gastos")
                        vImporteConcepto = Math.Abs(Val(miView(x)("Importe")))
                        Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
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
                    Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), vImporteConcepto)
                Else
                    Dim i As Integer = .Points.AddXY(miView(x)("Cuenta"), miView(x)("Importe"))
                End If
                .ChartType = SeriesChartType.Pie
            End With
            With Chart1.Series("Ingresos")
                .ChartType = SeriesChartType.Pie
            End With
        Next
    End Sub
End Class