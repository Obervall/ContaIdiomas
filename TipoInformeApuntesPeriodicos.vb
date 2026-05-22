Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class TipoInformeApuntesPeriodicos

    Private vtipoSql, vAñadir, vAñadir2 As String
    Private PrintLine, Contador As Integer
    Public Property DgvApuntes As Object

    Private Sub TipoInformeApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vDate1 = Format(frmApuntesPeriodicos.DateTimePicker1.Value, "yyyy/MM/dd")
        vDate2 = Format(frmApuntesPeriodicos.DateTimePicker2.Value, "yyyy/MM/dd")

        'Comienzo del SELECT
        '********************
        If RadioButton1.Checked = True Or RadioButton2.Checked = True Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP,  apuper.CuentaAPP,  apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
        Else
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
        End If

        'Comienzo del Título
        '********************
        If RadioButton1.Checked = True Then
            frmImprimirForm.LblTitulo.Text = "Listado Completo."
        End If
        If RadioButton2.Checked = True Then
            frmImprimirForm.LblTitulo.Text = "Listado por Conceptos."
        End If
        If RadioButton3.Checked = True Then
            frmImprimirForm.LblTitulo.Text = "Listado por Cuentas."
        End If
        If RadioButton4.Checked = True Then
            frmImprimirForm.LblTitulo.Text = "Listado por Fechas."
        End If

        'Siguiente parte, General, del Título, si hay algún Filtro checkeado
        '*****************************************************************
        If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Or frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Or frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            frmImprimirForm.LblTitulo.Text += " Filtrado:"
        End If

        'Siguiente parte del Título con el texto del componente filtrado, según el Combo
        '*******************************************************************************
        If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuper.ConceptoAPP = '" & frmApuntesPeriodicos.CmbConcepto.Text & "' "
            frmImprimirForm.LblTitulo.Text += "  " & frmApuntesPeriodicos.CmbConcepto.Text & "."
        End If
        If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuper.CuentaAPP = '" & frmApuntesPeriodicos.CmbCuenta.Text & "' "
            frmImprimirForm.LblTitulo.Text += "  " & frmApuntesPeriodicos.CmbCuenta.Text & "."
        End If
        If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            vtipoSql += " And apuper.FechaAPP >= #" & vDate1 & "#"
            vtipoSql += " And apuper.FechaAPP <= #" & vDate2 & "#"
            frmImprimirForm.LblTitulo.Text += "  FECHAS."
            frmImprimirForm.LblEntreFechas.Text = "Desde: " & frmApuntesPeriodicos.DateTimePicker1.Value & "    Hasta: " & frmApuntesPeriodicos.DateTimePicker2.Value
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** COMPLETO ***
        '*****************************************************************************
        If RadioButton1.Checked = True Then
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "1")
            vIngresos = 0
            vGastos = 0
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vSaldo = fila.Cells(3).Value + vValor
                fila.Cells(4).Value = vSaldo
                vValor = fila.Cells(4).Value
                If fila.Cells(3).Value >= 0 Then
                    vIngresos += fila.Cells(3).Value
                Else
                    vGastos += fila.Cells(3).Value
                End If
                frmImprimirForm.LblTotal.Text = "Total Ingresos: " & Format(vIngresos, "###,##0.00").ToString & "  -  Total Gastos: " & Format(vGastos, "###,##0.00").ToString & "                        TOTAL: " & Format(vValor, "###,##0.00 ").ToString & vMoneda
            Next
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** CONCEPTOS ***
        '******************************************************************************
        If RadioButton2.Checked = True Then
            vtipoSql += " ORDER BY apuper.ConceptoAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")  'Agrupado Conceptos
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double = 0
            Dim vImporteTempapu As Double = 0
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMySql1cr.CommandText = vTempapu
            Try
                cmdMySql1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                MsgBox("Error al iniciar la tabla Tempapu")
                MsgBox(ex.ToString)
            End Try
            'Llenamos la tabla Temporal con los Conceptos Agrupados
            '******************************************************
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMySql1cr.CommandText = vAñadir
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al añadir el Concepto a la tabla Tempapu")
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
                        MsgBox("Error al verificar si el Concepto ya existe en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMySql1cr.CommandText = vAñadir2
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al actualizar el Importe del Concepto en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    drMySql1.Close()
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
                frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
            Next
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** CUENTAS ***
        '******************************************************************************
        If RadioButton3.Checked = True Then 'Por Cuentas
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = '" & frmApuntesPeriodicos.CmbConcepto.Text & "' "
            End If
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = '" & frmApuntesPeriodicos.CmbCuenta.Text & "' "
            End If
            If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                vtipoSql += " And apuper.FechaAPP >= #" & vDate1 & "#"
                vtipoSql += " And apuper.FechaAPP <= #" & vDate2 & "#"
            End If
            vtipoSql += " ORDER BY apuper.CuentaAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_CUENTAS_PERIODICAS", "3")  'Agrupado por Cuentas
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double = 0
            Dim vImporteTempapu As Double = 0
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMySql1cr.CommandText = vTempapu
            Try
                cmdMySql1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                MsgBox("Error al iniciar la tabla Tempapu")
                MsgBox(ex.ToString)
            End Try
            'Llenamos la tabla Temporal con los Conceptos Agrupados
            '******************************************************

            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(6).Value.ToString Then
                    vNombreConcepto = fila.Cells(6).Value.ToString
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMySql1cr.CommandText = vAñadir
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al añadir el Concepto a la tabla Tempapu")
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
                        MsgBox("Error al verificar si el Concepto ya existe en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMySql1cr.CommandText = vAñadir2
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al actualizar el Importe del Concepto en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    drMySql1.Close()
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
                frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
            Next
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** FECHAS ***
        '******************************************************************************
        If RadioButton4.Checked = True Then 'Por Fechas
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = '" & frmApuntesPeriodicos.CmbConcepto.Text & "' "
            End If
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = '" & frmApuntesPeriodicos.CmbCuenta.Text & "' "
            End If
            If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                vtipoSql += " And apuper.FechaAPP >= #" & vDate1 & "#"
                vtipoSql += " And apuper.FechaAPP <= #" & vDate2 & "#"
            End If
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_CUENTAS_PERIODICAS", "3")  'Agrupado por Cuentas
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double = 0
            Dim vImporteTempapu As Double = 0
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMySql1cr.CommandText = vTempapu
            Try
                cmdMySql1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                MsgBox("Error al iniciar la tabla Tempapu")
                MsgBox(ex.ToString)
            End Try
            'Llenamos la tabla Temporal con los Conceptos Agrupados
            '******************************************************

            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(0).Value.ToString Then
                    vNombreConcepto = fila.Cells(0).Value.ToString
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                    cmdMySql1cr.CommandText = vAñadir
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al añadir el Concepto a la tabla Tempapu")
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
                        MsgBox("Error al verificar si el Concepto ya existe en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    cmdMySql1cr.CommandText = vAñadir2
                    Try
                        drMySql1 = cmdMySql1cr.ExecuteReader()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al actualizar el Importe del Concepto en la tabla Tempapu")
                        MsgBox(ex.ToString)
                    End Try
                    drMySql1.Close()
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
                frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
            Next
        End If

        'Iniciamos Código para Imprimir
        '******************************
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString
        PrintLine = 0
        Contador = 0
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
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Right, frmImprimirForm.LblEntreFechas.Top)
        End If

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        If RadioButton1.Checked = True Then
            e.Graphics.DrawString("Fecha:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString("Concepto:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
            e.Graphics.DrawString("Descripción:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
            e.Graphics.DrawString("Importe(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto4.Left, frmImprimirForm.Punto4.Top - 30)
            e.Graphics.DrawString("Saldo(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)
        End If

        If RadioButton2.Checked = True Then
            e.Graphics.DrawString("Concepto:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString("Importe(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        If RadioButton3.Checked = True Then
            e.Graphics.DrawString("Grupo Cuentas:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString("Importe(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        If RadioButton4.Checked = True Then
            e.Graphics.DrawString("Grupo Fechas:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString("Importe(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'
        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                'Esta parte se activa solo si 'startY' que es la posicion vertical almacenada supera el borde inferior de la pagina
                'Este se reinicia con cada pagina necesitada
                e.HasMorePages = True
                Exit Do
            End If
            If RadioButton1.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
                e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto4.Right + 50, startY, sf)
                e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(4).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 30, startY, sf)
                startY += frmImprimirForm.LblFecha.Height
                e.Graphics.DrawString("Cuenta:  ", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(6).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 50, startY)
                startY += frmImprimirForm.LblFecha.Height
                e.Graphics.DrawString("Notas:   ", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(5).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 50, startY)
                startY += frmImprimirForm.LblFecha.Height
                e.Graphics.DrawString("---------------------------------------------------------------------------------------------------------------------------------------------------------------------", FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            End If

            If RadioButton2.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
            End If

            If RadioButton3.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
            End If

            If RadioButton4.Checked = True Then
                e.Graphics.DrawString(Mid(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value, 1, 10), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
            End If

            'Aqui estoy usando un tipo de letras mas grande
            'LabelCodigo' mas grande que 'Punto1' para crear mas espacio entre filas

            'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
            'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
            ''Imprimimos los valores que salen despues del datagridview al final del reporte
            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop
        'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
        'Imprimimos los valores que salen despues del datagridview al final del reporte
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            If RadioButton1.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 30, startY + 15, sf)
            End If
            If RadioButton2.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY + 15, sf)
            End If
            If RadioButton3.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY + 15, sf)
            End If
            If RadioButton4.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY + 15, sf)
            End If

            'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
            PrintLine = 0
            Contador = 0
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

        'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
        If Contador = 0 Then
            frmImprimirForm.LblNumeroPagina.Text = "0"
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class