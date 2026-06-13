Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class TipoInformeApuntes

    Private vtipoSql, vAñadir, vAñadir2 As String
    Private PrintLine, Contador As Integer
    Public Property DgvApuntes As Object
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub TipoInformeApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        If frmApuntesContables.BtnFiltroChekedList.Enabled = False Then
            frmTipoInformeApuntes.RadioButton2.Enabled = False
            frmTipoInformeApuntes.RadioButton5.Enabled = True
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
        vDate2 = frmApuntesContables.DateTimePicker2.Value.Date

        'Comienzo del Título
        '********************
        If RadioButton1.Checked = True Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoCompleto")
        End If
        If RadioButton2.Checked = True Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoConceptos")
        End If
        If RadioButton3.Checked = True Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoCuentas")
        End If
        If RadioButton4.Checked = True Then
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoFechas")
        End If

        'Siguiente parte, General, del Título, si hay algún Filtro checkeado
        '*****************************************************************
        ' Filtrado: si hay algún filtro checkeado, añadimos la palabra "Filtrado" al título,
        ' para diferenciarlo del listado completo sin filtros.
        If frmApuntesContables.BtnFiltroConcepto.Enabled = False Or frmApuntesContables.BtnFiltroCuenta.Enabled = False Or frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            frmImprimirForm.LblTitulo.Text += " " & resManager.GetString("Filtrado")
        End If

        'Siguiente parte del Título con el texto del componente filtrado, según el Combo
        '*******************************************************************************
        If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
            If frmApuntesContables.ListBox1.SelectedItems.Count >= 2 Then
                frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoConceptosChequeados")
            Else
                frmImprimirForm.LblTitulo.Text += "  " & frmApuntesContables.CmbConcepto.Text & "."
            End If
        End If
        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
            If frmApuntesContables.ListBox1.SelectedItems.Count >= 2 Then
                frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoCuentasConceptosChequeados")
            Else
                frmImprimirForm.LblTitulo.Text += "  " & frmApuntesContables.CmbCuenta.Text & "."
            End If
        End If
        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            If frmApuntesContables.ListBox1.SelectedItems.Count >= 2 Then
                frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoFechasConceptosChequeados")
            Else
                frmImprimirForm.LblTitulo.Text += "  " & resManager.GetString("FECHA_UPPER") & "."
            End If
            frmImprimirForm.LblEntreFechas.Text = resManager.GetString("Desde") & ": " & frmApuntesContables.DateTimePicker1.Value & "    " & resManager.GetString("Hasta") & ": " & frmApuntesContables.DateTimePicker2.Value
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** COMPLETO ***
        '*****************************************************************************
        If RadioButton1.Checked = True Then
            vIngresos = 0
            vGastos = 0
            vValor = 0
            ' 1. Saca la actualización del LblTotal fuera del bucle (por rendimiento)
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                ' Evita errores si la fila es la de inserción (nueva fila vacía al final)
                If Not fila.IsNewRow AndAlso fila.Cells(3).Value IsNot Nothing Then

                    ' Convertimos el valor de la celda a Double de forma segura
                    Dim valorCelda As Double = Convert.ToDouble(fila.Cells(3).Value)
                    ' Cálculo del saldo acumulado
                    vSaldo = valorCelda + vValor
                    fila.Cells(4).Value = vSaldo
                    vValor = vSaldo ' Asignamos directamente el número, no el valor de la celda
                    ' Acumuladores de Ingresos y Gastos
                    If valorCelda >= 0 Then
                        vIngresos += valorCelda
                    Else
                        vGastos += valorCelda
                    End If
                End If
            Next
            ' 2. Formateo multiidioma automático con .ToString("N2") fuera del bucle
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1}  -  {2}: {3}                        {4}: {5}{6}",
                resManager.GetString("TotalIngresos"),
                vIngresos.ToString("N2"),
                resManager.GetString("TotalGastos"),
                vGastos.ToString("N2"),
                resManager.GetString("TOTAL"),
                vValor.ToString("N2"),
                vMoneda
            )
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** CONCEPTOS ***
        '******************************************************************************
        If RadioButton2.Checked = True Then
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double
            Dim vImporteTempapu As String = ""
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMdb1cr.CommandText = vTempapu
            Try
                cmdMdb1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                'MsgBox("Error al Borrar los Registros de Tempapu")
                MsgBox(ex.ToString)
            End Try

            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
            DgvApuntesContables(3, 4)

            'Llenamos la tabla Temporal con los Conceptos Agrupados
            '******************************************************
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES (?, ?)"
                    cmdMdb1cr.CommandText = vAñadir
                    cmdMdb1cr.Parameters.Clear() ' Limpia parámetros anteriores
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vImporteConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Grabar el Concepto: " & vNombreConcepto)
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
                        Else
                            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        'MsgBox("Error al Leer el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = ? "
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vNewImporteConcepto
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Actualizar el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** CUENTAS ***
        '******************************************************************************
        If RadioButton3.Checked = True Then 'Por Cuentas
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double
            Dim vImporteTempapu As String = ""
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMdb1cr.CommandText = vTempapu
            Try
                cmdMdb1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                'MsgBox("Error al Borrar los Registros de Tempapu")
                MsgBox(ex.ToString)
            End Try

            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(6), System.ComponentModel.ListSortDirection.Ascending)
            DgvApuntesContables(3, 4)

            'Llenamos la tabla Temporal con las Cuentas Agrupadas
            '****************************************************
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(6).Value Then
                    vNombreConcepto = fila.Cells(6).Value
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES (?, ?)"
                    cmdMdb1cr.CommandText = vAñadir
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vImporteConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Grabar la Cuenta: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ?"
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vExistenteImporteConcepto = drMdb1.GetValue(1)
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        'MsgBox("Error al Leer el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = ? "
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vNewImporteConcepto
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Actualizar la Cuenta: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** FECHAS ***
        '***************************************************************************
        If RadioButton4.Checked = True Then 'Por Fechas
            Dim vTmpprint As String
            Dim vImporteFecha As String
            Dim vImporteTmpprint As String = ""
            Dim vNewImporteFechas As String
            'Iniciamos Tabla Tmpprint
            vTmpprint = "DELETE FROM tmpprint"
            cmdMdb1cr.CommandText = vTmpprint
            Try
                cmdMdb1cr.ExecuteNonQuery()
                'MsgBox("Registros Tmpprint, Borrados !!!")
            Catch ex As Exception
                'MsgBox("Error al Borrar los Registros de Tmpprint")
                MsgBox(ex.ToString)
            End Try

            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(0), System.ComponentModel.ListSortDirection.Ascending)
            DgvApuntesContables(3, 4)

            vValor = 0
            vFechaTemp = CDate("01/01/1900")
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                vValor += fila.Cells(3).Value
                If vFechaTemp <> CDate(fila.Cells(0).Value) Then
                    vFechaTemp = CDate(fila.Cells(0).Value)
                    vImporteFecha = ""
                    vImporteFecha = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tmpprint"
                    vAñadir += "(FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) "
                    vAñadir += "VALUES (?, ?, ?, ?, ?, ?, ?)"
                    cmdMdb1cr.CommandText = vAñadir
                    cmdMdb1cr.Parameters.Clear() ' Limpia parámetros anteriores
                    cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vFechaTemp
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = ""
                    cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = ""
                    cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.VarWChar).Value = ""
                    cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = ""
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vImporteFecha
                    cmdMdb1cr.Parameters.Add("@sal", OleDb.OleDbType.Currency).Value = vImporteFecha
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Grabar la Fecha: " & vFechaTemp)
                        MsgBox(ex.ToString)
                    End Try
                Else
                    vFechaTemp2 = vFechaTemp
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.Parameters.Clear() ' Limpia parámetros anteriores
                    cmdMdb1cr.CommandText = "SELECT * FROM tmpprint WHERE tmpprint.FechaTMP = ?"
                    cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vFechaTemp2
                    If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                        vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                        vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                        cmdMdb1cr.CommandText += " And tmpprint.FechaTMP >= ?"
                        cmdMdb1cr.CommandText += " And tmpprint.FechaTMP <= ?"
                        cmdMdb1cr.Parameters.Add("@date1", OleDb.OleDbType.Date).Value = vDate1
                        cmdMdb1cr.Parameters.Add("@date2", OleDb.OleDbType.Date).Value = vDate2
                    End If
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            'MsgBox(cmdMdb1cr.CommandText)
                            While drMdb1.Read()
                                vImporteTmpprint = drMdb1.GetValue(5)
                                'MsgBox(vImporteTmpprint)
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        'MsgBox("Error al Leer la Fecha: " & vFechaTemp)
                        MsgBox(ex.ToString)
                    End Try
                    'vNewImporteFechas = (vImporteTmpprint + Val(fila.Cells(3).Value)).ToString
                    ' 1. Convertimos el acumulador actual a Decimal de forma segura
                    Dim importeAcumulado As Decimal = 0.0D
                    If vImporteTmpprint IsNot Nothing Then
                        Decimal.TryParse(vImporteTmpprint.ToString().Replace(",", "."),
                     System.Globalization.NumberStyles.Any,
                     System.Globalization.CultureInfo.InvariantCulture,
                     importeAcumulado)
                    End If

                    ' 2. Convertimos el importe de la celda actual a Decimal
                    Dim importeCelda As Decimal = 0.0D
                    If fila.Cells(3).Value IsNot Nothing AndAlso Not IsDBNull(fila.Cells(3).Value) Then
                        Decimal.TryParse(fila.Cells(3).Value.ToString().Replace(",", "."),
                     System.Globalization.NumberStyles.Any,
                     System.Globalization.CultureInfo.InvariantCulture,
                     importeCelda)
                    End If

                    ' 3. Realizamos la suma matemática exacta
                    Dim sumaTotal As Decimal = importeAcumulado + importeCelda

                    ' 4. Guardamos el resultado en formato texto para tu variable de impresión
                    vNewImporteFechas = sumaTotal.ToString(System.Globalization.CultureInfo.InvariantCulture)

                    vAñadir2 = "UPDATE tmpprint SET ImporteTMP = ? WHERE tmpprint.FechaTMP = ?"
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vNewImporteFechas
                    cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vFechaTemp2
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Actualizar la Fecha: " & vFechaTemp)
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)
            'Llenamos la tabla Temporal con los cálculos realizados
            vtipoSql = "SELECT * FROM tmpprint ORDER BY tmpprint.FechaTMP ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES_FECHAS", "4")
        End If

        'Llenar el Grid de ImprimirForm para leerlo luego en el Print *** CONCEPTOS CHECKEADOS ***
        '*****************************************************************************************
        If RadioButton5.Checked = True Then
            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNewImporteConcepto As Double
            Dim vExistenteImporteConcepto As Double
            Dim vImporteTempapu As String = ""
            vNombreConcepto = ""
            'Iniciamos Tabla Tempapu
            '***********************
            vTempapu = "DELETE FROM tempapu"
            cmdMdb1cr.CommandText = vTempapu
            Try
                cmdMdb1cr.ExecuteNonQuery()
                'MsgBox("Registros Tempapu, Borrados !!!")
            Catch ex As Exception
                'MsgBox("Error al Borrar los Registros de Tempapu")
                MsgBox(ex.ToString)
            End Try
            'Ordenamos la columna Concepto, antes de calcular los totales parciales.
            '***********************************************************************
            frmApuntesContables.DgvApuntes.Sort(frmApuntesContables.DgvApuntes.Columns(1), System.ComponentModel.ListSortDirection.Ascending)

            'Llenamos la tabla Temporal con los Conceptos Agrupados desde DgvApuntes
            '***********************************************************************
            For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = 0
                    vImporteConcepto = fila.Cells(3).Value
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES (?, ?)"
                    cmdMdb1cr.CommandText = vAñadir
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vImporteConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Grabar el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ?"
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vExistenteImporteConcepto = drMdb1.GetValue(1)
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        'MsgBox("Error al Leer el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? "
                    vAñadir2 += " WHERE tempapu.ConceptoAPU = ?"
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = vNewImporteConcepto
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.VarWChar).Value = vNombreConcepto
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        'MsgBox("Error al Actualizar el Concepto: " & vNombreConcepto)
                        MsgBox(ex.ToString)
                    End Try
                End If
            Next
            'Llenamos la tabla de ImprimirForm con los cálculos realizados
            '*************************************************************
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(1).Value
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)
            frmImprimirForm.LblTitulo.Text = rmse.GetString("ListadoConceptosChequeados")
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
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Right, frmImprimirForm.LblEntreFechas.Top)
        End If

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        If RadioButton1.Checked = True Then
            e.Graphics.DrawString(resManager.GetString("Fecha") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Concepto") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto4.Left, frmImprimirForm.Punto4.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Saldo") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)
        End If

        If RadioButton2.Checked = True Then
            e.Graphics.DrawString(resManager.GetString("Concepto") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        If RadioButton3.Checked = True Then
            e.Graphics.DrawString(resManager.GetString("Cuenta") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        If RadioButton4.Checked = True Then
            e.Graphics.DrawString(resManager.GetString("Fecha") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        If RadioButton5.Checked = True Then
            e.Graphics.DrawString(resManager.GetString("Concepto") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(resManager.GetString("Importe") & "(" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        End If

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'
        If RadioButton1.Checked = True Then
            Do While PrintLine < frmApuntesContables.DgvApuntes.Rows.Count
                If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                    'Esta parte se activa solo si 'startY' que es la posicion vertical almacenada supera el borde inferior de la pagina
                    'Este se reinicia con cada pagina necesitada
                    e.HasMorePages = True
                    Exit Do
                End If
                If RadioButton1.Checked = True Then
                    e.Graphics.DrawString(DirectCast(frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(0).Value, DateTime).ToString("d"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                    e.Graphics.DrawString(frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
                    e.Graphics.DrawString(frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(2).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
                    ' Convertimos a Double y aplicamos "N2" de forma nativa
                    Dim importeFormateado As String = Convert.ToDouble(frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(3).Value).ToString("N2")
                    Dim importeFormateado2 As String = Convert.ToDouble(frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(4).Value).ToString("N2")
                    ' Lo pintamos usando tu variable de alineación 'sf'
                    e.Graphics.DrawString(importeFormateado, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto4.Right + 50, startY, sf)
                    e.Graphics.DrawString(importeFormateado2, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY, sf)
                    ' --- IMPRESIÓN DE CUENTA ---
                    startY += frmImprimirForm.LblFecha.Height
                    ' 1. Guardamos la etiqueta y el valor de forma segura
                    Dim etiquetaCuenta As String = resManager.GetString("Cuenta") & ":   "
                    Dim valorCuenta As String = ""
                    If frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(6).Value IsNot Nothing Then
                        valorCuenta = frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(6).Value.ToString()
                    End If
                    ' 2. Pintamos la etiqueta ("Cuenta:  ")
                    e.Graphics.DrawString(etiquetaCuenta, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                    ' 3. Calculamos matemáticamente el ancho de la etiqueta traducida
                    Dim anchoEtiquetaCuenta As Single = e.Graphics.MeasureString(etiquetaCuenta, FuenteSubrayada).Width
                    ' 4. Pintamos el valor sumando el ancho exacto (nunca se solapará)
                    e.Graphics.DrawString(valorCuenta, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + anchoEtiquetaCuenta, startY)
                    ' --- IMPRESIÓN DE NOTAS ---
                    startY += frmImprimirForm.LblFecha.Height
                    ' 1. Guardamos la etiqueta y el valor de forma segura
                    Dim etiquetaNotas As String = resManager.GetString("Notas") & ":    "
                    Dim valorNotas As String = ""
                    If frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(5).Value IsNot Nothing Then
                        valorNotas = frmApuntesContables.DgvApuntes.Rows(PrintLine).Cells(5).Value.ToString()
                    End If
                    ' 2. Pintamos la etiqueta ("Notas:   ")
                    e.Graphics.DrawString(etiquetaNotas, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                    ' 3. Calculamos matemáticamente el ancho de la etiqueta traducida
                    Dim anchoEtiquetaNotas As Single = e.Graphics.MeasureString(etiquetaNotas, FuenteSubrayada).Width
                    ' 4. Pintamos el valor en su posición exacta calculada dinámicamente
                    e.Graphics.DrawString(valorNotas, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + anchoEtiquetaNotas, startY)
                    startY += frmImprimirForm.LblFecha.Height
                    e.Graphics.DrawString("---------------------------------------------------------------------------------------------------------------------------------------------------------------------", FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
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
            If Contador >= frmApuntesContables.DgvApuntes.Rows.Count Then
                If RadioButton1.Checked = True Then
                    e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 40, startY, sf)
                End If
                'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
                PrintLine = 0
                Contador = 0
            End If
            'Si deseamos poner un contador de páginas
            'Esta parte siempre va a salir en todas las paginas
            frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
            e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
            e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
            'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
            If Contador = 0 Then
                frmImprimirForm.LblNumeroPagina.Text = "0"
            End If
        Else
            Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
                Dim filaImprimir As DataGridViewRow = frmImprimirForm.DgvApuntes.Rows(PrintLine)
                ' 1. SEGURIDAD: Si es la fila nueva vacía del final, la saltamos y avanzamos
                If filaImprimir.IsNewRow Then
                    PrintLine += 1
                    Continue Do
                End If
                ' 2. CONTROL DE SALTO DE PÁGINA (¡Arreglado!)
                ' Si no cabe la línea actual, marcamos que hay más páginas y salimos. 
                ' Al NO avanzar PrintLine aquí, la siguiente página empezará exactamente en esta fila.
                If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                    e.HasMorePages = True
                    Exit Do
                End If
                ' Variables para almacenar dinámicamente qué vamos a pintar en esta vuelta
                Dim textoCelda0 As String = ""
                Dim valorImporte As Double = 0
                Dim celdaImporte As Object = Nothing
                ' 3. LOGICA DE FILTROS SEGÚN RADIOBUTTONS
                ' Obtenemos los valores de forma segura controlando los nulos (Nothing)
                If RadioButton2.Checked OrElse RadioButton3.Checked OrElse RadioButton5.Checked Then
                    If filaImprimir.Cells(0).Value IsNot Nothing Then textoCelda0 = filaImprimir.Cells(0).Value.ToString()
                    celdaImporte = filaImprimir.Cells(1).Value
                ElseIf RadioButton4.Checked Then
                    If filaImprimir.Cells(0).Value IsNot Nothing Then
                        textoCelda0 = DirectCast(filaImprimir.Cells(0).Value, DateTime).ToString("d")
                    End If
                    ' En el RadioButton4 usabas la celda 5
                    celdaImporte = filaImprimir.Cells(5).Value
                End If
                ' Convertimos el importe a número solo si la celda tiene datos
                If celdaImporte IsNot Nothing AndAlso IsNumeric(celdaImporte) Then
                    valorImporte = Convert.ToDouble(celdaImporte)
                End If
                ' 4. PINTADO EN EL DOCUMENTO (Se ejecuta una sola vez en lugar de repetir código)
                e.Graphics.DrawString(textoCelda0, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(valorImporte.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
                ' 5. AVANCE DE LÍNEA Y CONTADORES
                startY += frmImprimirForm.LblFecha.Height
                PrintLine += 1
                Contador += 1
            Loop
            ' 6. FINALIZACIÓN DE LA IMPRESIÓN
            ' Si el bucle termina de forma natural porque recorrió todas las filas, cerramos la paginación
            If PrintLine >= frmImprimirForm.DgvApuntes.Rows.Count Then
                e.HasMorePages = False
                PrintLine = 0 ' Reiniciamos para la próxima vez que el usuario pulse Imprimir
            End If
            'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
            'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
            'Imprimimos los valores que salen despues del datagridview al final del reporte
            If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
                If RadioButton2.Checked = True Then
                    e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
                End If
                If RadioButton3.Checked = True Then
                    e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
                End If
                If RadioButton4.Checked = True Then
                    e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
                End If
                If RadioButton5.Checked = True Then
                    e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Right + 50, startY, sf)
                End If
                'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
                PrintLine = 0
                Contador = 0
            End If
            'Si deseamos poner un contador de páginas
            'Esta parte siempre va a salir en todas las paginas
            frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
            e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
            e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
            'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
            If Contador = 0 Then
                frmImprimirForm.LblNumeroPagina.Text = "0"
            End If
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class