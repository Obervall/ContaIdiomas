Option Explicit On
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class TipoInformeApuntesPeriodicos

    Private vtipoSql, vAñadir, vAñadir2 As String
    Private PrintLine, Contador As Integer
    Public Property DgvApuntes As Object
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub TipoInformeApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' =========================================================================
        ' 🌟 CORTAFUEGOS INDESTRUCTIBLE DE PURGA DE RAM (¡La estocada definitiva!)
        ' =========================================================================
        ' Vaciamos las variables globales y el motor de comandos para eliminar 
        ' cualquier consulta zombi que se haya quedado congelada en la memoria.
        vtipoSql = ""
        vtipoSqlChk = ""
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.CommandText = ""

        ' Recuperamos tus variables cronológicas dóciles de fábrica
        vDate1 = frmApuntesPeriodicos.DateTimePicker1.Value.Date
        vDate2 = frmApuntesPeriodicos.DateTimePicker2.Value.Date

        ' =========================================================================
        ' 🌟 1. RECONSTRUCCIÓN LIMPIA DEL SELECT BASE (Con el espacio de seguridad inicial)
        ' =========================================================================
        If RadioButton1.Checked = True OrElse RadioButton2.Checked = True Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP As [SaldoAPP], apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
        Else
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
        End If

        If RadioButton1.Checked = True Then
            frmImprimirForm.LblTitulo.Text = frmTipoInformeApuntes.rmse.GetString("ListadoCompleto")
        End If
        If RadioButton2.Checked = True Then
            frmImprimirForm.LblTitulo.Text = frmTipoInformeApuntes.rmse.GetString("ListadoConceptos")
        End If
        If RadioButton3.Checked = True Then
            frmImprimirForm.LblTitulo.Text = frmTipoInformeApuntes.rmse.GetString("ListadoCuentas")
        End If
        If RadioButton4.Checked = True Then
            frmImprimirForm.LblTitulo.Text = frmTipoInformeApuntes.rmse.GetString("ListadoFechas")
        End If

        If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False OrElse frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False OrElse frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            frmImprimirForm.LblTitulo.Text += " " & resManager.GetString("Filtrado") & ":"
        End If
        ' -------------------------------------------------------------------------
        ' 🌟 TRAMO 2 DE 4: FILTROS RELACIONALES INTELIGENTES (Inmunes a fallos de tipos)
        ' -------------------------------------------------------------------------
        If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
            Dim textoConceptoSel As String = frmApuntesPeriodicos.CmbConcepto.Text.Trim()
            Dim idConcepto As Integer = 0

            Using con As New OleDbConnection(conexion1.ConnectionString)
                Using cmd As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? OR DescripcionCON = ? OR CodigoCON = ?", con)
                    cmd.Parameters.Add("@c1", OleDbType.VarWChar).Value = textoConceptoSel.Replace(" ", "_").ToUpper()
                    cmd.Parameters.Add("@d1", OleDbType.VarWChar).Value = textoConceptoSel
                    cmd.Parameters.Add("@c2", OleDbType.VarWChar).Value = textoConceptoSel.ToUpper()
                    Try
                        con.Open()
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not IsDBNull(res) Then idConcepto = Convert.ToInt32(res)
                    Catch
                    End Try
                End Using
            End Using

            If idConcepto > 0 Then
                vtipoSql += " And apuper.ConceptoAPP = " & idConcepto.ToString()
            Else
                vtipoSql += " And apuper.ConceptoAPP = -1"
            End If
            frmImprimirForm.LblTitulo.Text += " " & textoConceptoSel & "."
        End If

        If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
            Dim textoCuentaSel As String = frmApuntesPeriodicos.CmbCuenta.Text.Trim()
            Dim idCuenta As Integer = 0

            Using con As New OleDbConnection(conexion1.ConnectionString)
                Using cmd As New OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", con)
                    cmd.Parameters.Add("@n", OleDbType.VarWChar).Value = textoCuentaSel
                    Try
                        con.Open()
                        Dim res = cmd.ExecuteScalar()
                        If res IsNot Nothing AndAlso Not IsDBNull(res) Then idCuenta = Convert.ToInt32(res)
                    Catch
                    End Try
                End Using
            End Using

            If idCuenta > 0 Then
                vtipoSql += " And apuper.CuentaAPP = " & idCuenta.ToString()
            Else
                vtipoSql += " And apuper.CuentaAPP = -1"
            End If
            frmImprimirForm.LblTitulo.Text += " " & textoCuentaSel & "."
        End If

        If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            Dim fInicio As String = "#" & frmApuntesPeriodicos.DateTimePicker1.Value.ToString("yyyy-MM-dd") & "#"
            Dim fFin As String = "#" & frmApuntesPeriodicos.DateTimePicker2.Value.ToString("yyyy-MM-dd") & "#"

            vtipoSql += " And apuper.FechaAPP >= " & fInicio
            vtipoSql += " And apuper.FechaAPP <= " & fFin

            frmImprimirForm.LblTitulo.Text += " " & resManager.GetString("Fechas") & "."
            frmImprimirForm.LblEntreFechas.Text = resManager.GetString("Desde") & ": " & frmApuntesPeriodicos.DateTimePicker1.Value.ToShortDateString() & " " & resManager.GetString("Hasta") & ": " & frmApuntesPeriodicos.DateTimePicker2.Value.ToShortDateString()
        End If

        ' -------------------------------------------------------------------------
        ' 🌟 TRAMO 3 DE 4: LISTADO COMPLETO (RB1) E INICIO DE TEMPORAL (RB2)
        ' -------------------------------------------------------------------------
        If RadioButton1.Checked = True Then
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "1")
            vIngresos = 0
            vGastos = 0
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For
                vSaldo = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value)) + vValor
                fila.Cells(4).Value = vSaldo
                vValor = fila.Cells(4).Value
                If CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value)) >= 0 Then
                    vIngresos += CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))
                Else
                    vGastos += CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))
                End If
                frmImprimirForm.LblTotal.Text = resManager.GetString("TotalIngresos") & ": " & vIngresos.ToString("N2") & "  -  " & resManager.GetString("TotalGastos") & ": " & vGastos.ToString("N2") & "                        " & resManager.GetString("TOTAL") & " :  " & vValor.ToString("N2") & vMoneda
            Next
        End If

        'If RadioButton2.Checked = True Then
        '    vtipoSql += " ORDER BY apuper.ConceptoAPP ASC"
        '    LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")

        '    Dim vTempapu As String
        '    Dim vImporteConcepto As Double
        '    Dim vNombreConcepto As String = ""

        '    vTempapu = "DELETE FROM tempapu"
        '    cmdMdb1cr.CommandText = vTempapu
        '    Try
        '        cmdMdb1cr.ExecuteNonQuery()
        '    Catch ex As Exception
        '        MsgBox(ex.ToString())
        '    End Try

        '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
        '        If fila.IsNewRow Then Continue For

        '        vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

        '        If vNombreConcepto <> fila.Cells(1).Value.ToString() Then
        '            vNombreConcepto = fila.Cells(1).Value.ToString().Trim()
        '            vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

        '            vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
        '            cmdMdb1cr.CommandType = CommandType.Text
        '            cmdMdb1cr.CommandText = vAñadir
        '            cmdMdb1cr.Parameters.Clear()
        '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

        '            Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            paramImpTemp.Value = Math.Round(vImporteConcepto, 2)

        '            Try
        '                cmdMdb1cr.ExecuteNonQuery()
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try
        '        Else
        '            cmdMdb1cr.CommandType = CommandType.Text
        '            cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE tempapu.ConceptoAPU = ?"
        '            cmdMdb1cr.Parameters.Clear()
        '            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

        '            Dim vExistenteImporteConcepto As Double = 0
        '            Try
        '                Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
        '                    If drMdb1.Read() Then
        '                        If drMdb1.GetValue(0) IsNot DBNull.Value Then
        '                            Double.TryParse(drMdb1.GetValue(0).ToString(), vExistenteImporteConcepto)
        '                        End If
        '                    End If
        '                End Using
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try

        '            vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto

        '            vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ?"
        '            cmdMdb1cr.CommandText = vAñadir2
        '            cmdMdb1cr.Parameters.Clear()

        '            'Dim paramSumaTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            'paramSumaTemp.Value = Math.Round(vNewImporteConcepto, 2)
        '            ' 🚀 REPARADO: Convertimos primero a Double puro y luego redondeamos sin errores
        '            Dim vImporteDouble As Double = CDbl(ConvertirDecimalSeguro(vImporteConcepto))
        '            Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            paramImpTemp.Value = Math.Round(vImporteDouble, 2)

        '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

        '            Try
        '                cmdMdb1cr.ExecuteNonQuery()
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try
        '        End If
        '    Next

        '    vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        '    LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        '    vValor = 0
        '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
        '        If fila.IsNewRow Then Continue For
        '        vValor += CDbl(ConvertirDecimalSeguro(fila.Cells(1).Value))
        '        frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
        '    Next
        'End If

        ' =========================================================================
        ' 🌟 CASO 2 SANEADO DE ALTA INGENIERÍA: CONCEPTOS P_U_R_O_S (Aceptar)
        ' =========================================================================
        If RadioButton2.Checked = True Then
            ' 1. VACIADO PREVENTIVO: Limpiamos la tabla intermedia del disco duro
            cmdMdb1cr.CommandText = "DELETE FROM tempapu"
            Try
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try

            ' 2. 🚀 LA JUGADA MAESTRA (GROUP BY): Forzamos a Access a agrupar y sumar los conceptos de forma nativa
            ' Sembramos la consulta agrupando por el campo exacto del ID del concepto
            Dim sqlAgruparConceptos As String = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) " &
                                                "SELECT CStr(apuper.ConceptoAPP), Sum(apuper.ImporteAPP) FROM apuper " &
                                                "WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString & " " &
                                                "GROUP BY apuper.ConceptoAPP"

            cmdMdb1cr.CommandType = CommandType.Text
            cmdMdb1cr.CommandText = sqlAgruparConceptos
            cmdMdb1cr.Parameters.Clear()
            Try
                ' 🎯 ¡BUM!: De un solo golpe en el disco duro, Access procesa, suma y consolida todos los conceptos
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error al consolidar conceptos nativo: " & ex.ToString())
            End Try

            ' 3. VOLCADO DIRECTO AL CANVAS DE IMPRESIÓN 
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")

            ' Calculamos la barra de totales finales del folio
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For
                vValor += CDbl(ConvertirDecimalSeguro(fila.Cells(1).Value))
                frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
            Next
        End If


        ' =========================================================================
        ' 🌟 CASO 3 SANEADO DE ALTA INGENIERÍA: CUENTAS BANCARIAS P_U_R_A_S (Aceptar)
        ' =========================================================================
        If RadioButton3.Checked = True Then
            ' 1. VACIADO PREVENTIVO: Limpiamos la tabla intermedia del disco duro
            cmdMdb1cr.CommandText = "DELETE FROM tempapu"
            Try
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try

            ' 2. 🚀 LA JUGADA MAESTRA (GROUP BY): Forzamos a Access a agrupar y sumar los bancos de forma nativa
            ' Sembramos la consulta agrupando por el campo exacto de la cuenta corriente
            Dim sqlAgruparBancos As String = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) " &
                                             "SELECT CStr(apuper.CuentaAPP), Sum(apuper.ImporteAPP) FROM apuper " &
                                             "WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString & " " &
                                             "GROUP BY apuper.CuentaAPP"

            cmdMdb1cr.CommandType = CommandType.Text
            cmdMdb1cr.CommandText = sqlAgruparBancos
            cmdMdb1cr.Parameters.Clear()
            Try
                ' 🎯 ¡BUM!: De un solo golpe en el disco duro, Access procesa, suma y consolida todos los bancos
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error al consolidar bancos nativo: " & ex.ToString())
            End Try

            ' 3. VOLCADO DIRECTO AL CANVAS DE IMPRESIÓN (Tu misma llamada dócil del RadioButton2)
            vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
            LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")

            ' Calculamos la barra de totales finales del folio
            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For
                vValor += CDbl(ConvertirDecimalSeguro(fila.Cells(1).Value))
                frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
            Next
        End If


        'If RadioButton4.Checked = True Then
        '    vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
        '    vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString

        '    If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
        '        Dim textoConceptoSel As String = frmApuntesPeriodicos.CmbConcepto.Text.Trim()
        '        Dim idConcepto As Integer = 0
        '        Using con As New OleDbConnection(conexion1.ConnectionString)
        '            Using cmd As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? OR DescripcionCON = ?", con)
        '                cmd.Parameters.Add("@c", OleDbType.VarWChar).Value = textoConceptoSel.Replace(" ", "_").ToUpper()
        '                cmd.Parameters.Add("@d", OleDbType.VarWChar).Value = textoConceptoSel
        '                Try
        '                    con.Open()
        '                    Dim res = cmd.ExecuteScalar()
        '                    If res IsNot Nothing Then idConcepto = Convert.ToInt32(res)
        '                Catch
        '                End Try
        '            End Using
        '        End Using
        '        vtipoSql += " And apuper.ConceptoAPP = " & If(idConcepto > 0, idConcepto.ToString(), "-1")
        '    End If

        '    If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
        '        Dim textoCuentaSel As String = frmApuntesPeriodicos.CmbCuenta.Text.Trim()
        '        Dim idCuenta As Integer = 0
        '        Using con As New OleDbConnection(conexion1.ConnectionString)
        '            Using cmd As New OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", con)
        '                cmd.Parameters.Add("@n", OleDbType.VarWChar).Value = textoCuentaSel
        '                Try
        '                    con.Open()
        '                    Dim res = cmd.ExecuteScalar()
        '                    If res IsNot Nothing Then idCuenta = Convert.ToInt32(res)
        '                Catch
        '                End Try
        '            End Using
        '        End Using
        '        vtipoSql += " And apuper.CuentaAPP = " & If(idCuenta > 0, idCuenta.ToString(), "-1")
        '    End If

        '    If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
        '        Dim fInicio As String = "#" & frmApuntesPeriodicos.DateTimePicker1.Value.ToString("yyyy-MM-dd") & "#"
        '        Dim fFin As String = "#" & frmApuntesPeriodicos.DateTimePicker2.Value.ToString("yyyy-MM-dd") & "#"
        '        vtipoSql += " And apuper.FechaAPP >= " & fInicio
        '        vtipoSql += " And apuper.FechaAPP <= " & fFin
        '    End If
        '    vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        '    LlenarGrid(vtipoSql, "PRINT_CUENTAS_PERIODICAS", "3")

        '    Dim vTempapu As String
        '    Dim vImporteConcepto As Double
        '    Dim vNombreConcepto As String = ""

        '    vTempapu = "DELETE FROM tempapu"
        '    cmdMdb1cr.CommandText = vTempapu
        '    Try
        '        cmdMdb1cr.ExecuteNonQuery()
        '    Catch ex As Exception
        '        MsgBox(ex.ToString())
        '    End Try

        '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
        '        If fila.IsNewRow Then Continue For

        '        vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

        '        If vNombreConcepto <> fila.Cells(0).Value.ToString() Then
        '            vNombreConcepto = fila.Cells(0).Value.ToString().Trim()
        '            vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

        '            vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
        '            cmdMdb1cr.CommandType = CommandType.Text
        '            cmdMdb1cr.CommandText = vAñadir
        '            cmdMdb1cr.Parameters.Clear()
        '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

        '            Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            paramImpTemp.Value = Math.Round(vImporteConcepto, 2)

        '            Try
        '                cmdMdb1cr.ExecuteNonQuery()
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try
        '        Else
        '            cmdMdb1cr.CommandType = CommandType.Text
        '            cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE tempapu.ConceptoAPU = ?"
        '            cmdMdb1cr.Parameters.Clear()
        '            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

        '            Dim vExistenteImporteConcepto As Double = 0
        '            Try
        '                Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
        '                    If drMdb1.Read() Then
        '                        If drMdb1.GetValue(0) IsNot DBNull.Value Then
        '                            Double.TryParse(drMdb1.GetValue(0).ToString(), vExistenteImporteConcepto)
        '                        End If
        '                    End If
        '                End Using
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try

        '            vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto

        '            vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ?"
        '            cmdMdb1cr.CommandText = vAñadir2
        '            cmdMdb1cr.Parameters.Clear()

        '            'Dim paramSumaTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            'paramSumaTemp.Value = Math.Round(vNewImporteConcepto, 2)
        '            ' 🚀 REPARADO: Convertimos primero a Double puro y luego redondeamos sin errores
        '            Dim vNewImporteDouble As Double = CDbl(ConvertirDecimalSeguro(vNewImporteConcepto))
        '            Dim paramSumaTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
        '            paramSumaTemp.Value = Math.Round(vNewImporteDouble, 2)

        '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

        '            Try
        '                cmdMdb1cr.ExecuteNonQuery()
        '            Catch ex As Exception
        '                MsgBox(ex.ToString())
        '            End Try
        '        End If
        '    Next

        '    vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPU ASC"
        '    LlenarGrid(vtipoSql, "PRINT_TEMP_APUNTES", "0")
        '    vValor = 0
        '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
        '        If fila.IsNewRow Then Continue For
        '        vValor += CDbl(ConvertirDecimalSeguro(fila.Cells(1).Value))
        '        frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
        '    Next
        'End If

        ' =========================================================================
        ' 🌟 CASO 4 CORREGIDO: FILTRADO ACUMULADO POR BLOQUES DE FECHAS (RB4)
        ' =========================================================================
        If RadioButton4.Checked = True Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString

            ' Filtros elásticos que neutralizan el reventón de tipos relacionales
            'If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
            '    Dim textoConceptoSel As String = frmApuntesPeriodicos.CmbConcepto.Text.Trim()
            '    Dim idConcepto As Integer = 0
            '    Using con As New OleDbConnection(conexion1.ConnectionString)
            '        Using cmd As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? OR DescripcionCON = ?", con)
            '            cmd.Parameters.Add("@c", OleDbType.VarWChar).Value = textoConceptoSel.Replace(" ", "_").ToUpper()
            '            cmd.Parameters.Add("@d", OleDbType.VarWChar).Value = textoConceptoSel
            '            Try
            '                con.Open()
            '                Dim res = cmd.ExecuteScalar()
            '                If res IsNot Nothing Then idConcepto = Convert.ToInt32(res)
            '            Catch
            '            End Try
            '        End Using
            '    End Using
            '    vtipoSql += " And apuper.ConceptoAPP = " & If(idConcepto > 0, idConcepto.ToString(), "-1")
            'End If

            'If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
            '    Dim textoCuentaSel As String = frmApuntesPeriodicos.CmbCuenta.Text.Trim()
            '    Dim idCuenta As Integer = 0
            '    Using con As New OleDbConnection(conexion1.ConnectionString)
            '        Using cmd As New OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", con)
            '            cmd.Parameters.Add("@n", OleDbType.VarWChar).Value = textoCuentaSel
            '            Try
            '                con.Open()
            '                Dim res = cmd.ExecuteScalar()
            '                If res IsNot Nothing Then idCuenta = Convert.ToInt32(res)
            '            Catch
            '            End Try
            '        End Using
            '    End Using
            '    vtipoSql += " And apuper.CuentaAPP = " & If(idCuenta > 0, idCuenta.ToString(), "-1")
            'End If

            ' 🚀 A. FILTRO DE CONCEPTO (2 comodines '?' = 2 parámetros exactos)
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                Dim textoConceptoSel As String = frmApuntesPeriodicos.CmbConcepto.Text.Trim()
                Dim idConcepto As Integer = 0
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Using cmd As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? OR DescripcionCON = ?", con)
                        cmd.Parameters.Clear() ' Limpieza preventiva local
                        cmd.Parameters.Add("@c", OleDbType.VarWChar).Value = textoConceptoSel.Replace(" ", "_").ToUpper()
                        cmd.Parameters.Add("@d", OleDbType.VarWChar).Value = textoConceptoSel
                        Try
                            con.Open()
                            Dim res = cmd.ExecuteScalar()
                            If res IsNot Nothing AndAlso Not IsDBNull(res) Then idConcepto = Convert.ToInt32(res)
                        Catch
                        End Try
                    End Using
                End Using
                vtipoSql += " And apuper.ConceptoAPP = " & If(idConcepto > 0, idConcepto.ToString(), "-1")
            End If

            ' 🚀 B. FILTRO DE CUENTA (1 comodín '?' = 1 parámetro exacto)
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                Dim textoCuentaSel As String = frmApuntesPeriodicos.CmbCuenta.Text.Trim()
                Dim idCuenta As Integer = 0
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Using cmd As New OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", con)
                        cmd.Parameters.Clear() ' Limpieza preventiva local
                        cmd.Parameters.Add("@n", OleDbType.VarWChar).Value = textoCuentaSel
                        Try
                            con.Open()
                            Dim res = cmd.ExecuteScalar()
                            If res IsNot Nothing AndAlso Not IsDBNull(res) Then idCuenta = Convert.ToInt32(res)
                        Catch
                        End Try
                    End Using
                End Using
                vtipoSql += " And apuper.CuentaAPP = " & If(idCuenta > 0, idCuenta.ToString(), "-1")
            End If

            If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                Dim fInicio As String = "#" & frmApuntesPeriodicos.DateTimePicker1.Value.ToString("yyyy-MM-dd") & "#"
                Dim fFin As String = "#" & frmApuntesPeriodicos.DateTimePicker2.Value.ToString("yyyy-MM-dd") & "#"
                vtipoSql += " And apuper.FechaAPP >= " & fInicio
                vtipoSql += " And apuper.FechaAPP <= " & fFin
            End If

            ' 🚀 CORRECCIÓN CLAVE 1: Forzamos el volcado unificado al canal estable "PRINT_APUNTES_CONTABLES"
            vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "1")

            Dim vTempapu As String
            Dim vImporteConcepto As Double
            Dim vNombreConcepto As String = ""

            vTempapu = "DELETE FROM tempapu"
            cmdMdb1cr.CommandText = vTempapu
            Try
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try

            ' Bucle acumulador cronológico sobre el Grid unificado
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For

                vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

                ' Forzamos la comparación estricta de la fecha limpia decapitando horas en la RAM
                Dim fechaFilaLimpia As String = ""
                If fila.Cells(0).Value IsNot Nothing Then
                    Dim fechaTemp As Date
                    If Date.TryParse(fila.Cells(0).Value.ToString(), fechaTemp) Then
                        fechaFilaLimpia = fechaTemp.ToShortDateString().Trim()
                    Else
                        fechaFilaLimpia = fila.Cells(0).Value.ToString().Trim()
                    End If
                End If

                If vNombreConcepto <> fechaFilaLimpia Then
                    vNombreConcepto = fechaFilaLimpia
                    vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

                    vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = vAñadir
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                    Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                    paramImpTemp.Value = Math.Round(vImporteConcepto, 2)

                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE tempapu.ConceptoAPU = ?"
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)

                    Dim vExistenteImporteConcepto As Double = 0
                    Try
                        Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                            If drMdb1.Read() Then
                                If drMdb1.GetValue(0) IsNot DBNull.Value Then
                                    Double.TryParse(drMdb1.GetValue(0).ToString(), vExistenteImporteConcepto)
                                End If
                            End If
                        End Using
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try

                    vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto

                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ?"
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()

                    Dim vNewImporteDouble As Double = CDbl(ConvertirDecimalSeguro(vNewImporteConcepto))
                    Dim paramSumaTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                    paramSumaTemp.Value = Math.Round(vNewImporteDouble, 2)

                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox(ex.ToString())
                    End Try
                End If
            Next

            ' =========================================================================
            ' 🌟 ORDENACIÓN CRONOLÓGICA BLINDADA EN EL VOLCADO FINAL (¡La estocada de cierre!)
            ' =========================================================================
            ' Forzamos el casteo con CDate para que ordene por tiempo real y no por texto alfabético
            vtipoSql = "SELECT * FROM tempapu ORDER BY CDate(tempapu.ConceptoAPU) ASC"
            LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "0")

            vValor = 0
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                If fila.IsNewRow Then Continue For

                ' Leemos Cells(1) que es donde viaja la SumaImporteAPU en tempapu
                Dim importeAcumulado As Double = 0
                If fila.Cells(1).Value IsNot Nothing Then
                    Double.TryParse(ConvertirDecimalSeguro(fila.Cells(1).Value), importeAcumulado)
                End If

                vValor += importeAcumulado
                frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
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
        ' Cualquier variable que desees que conserve su valor debes declararla fuera del Printdocument
        ' Todas las variables declaradas dentro de printdocument pierden su valor al cambiar de pagina

        ' Definimos los tipos de letras a utilizar en el reporte
        ' ******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 14)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}

        ' Imprimimos el encabezado los datos que están antes del datagridview
        ' *******************************************************************
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)

        Dim newImage As Image = frmImprimirForm.PictureBox1.Image
        e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Right, frmImprimirForm.LblEntreFechas.Top)
        End If

        ' =========================================================================
        ' 🌟 ENCABEZADOS DE COLUMNA ASOCIADOS A RECURSOS (¡Adiós fallos de acentos!)
        ' =========================================================================
        ' Capturamos los textos limpios directamente de tus archivos .resx en caliente
        Dim txtFecha As String = If(resManager.GetString("Fecha"), "Fecha") & ":"
        Dim txtConcepto As String = If(resManager.GetString("Concepto"), "Concepto") & ":"
        Dim txtDescripcion As String = If(resManager.GetString("Descripcion"), "Descripción") & ":" ' 🚀 REPARADO: Sin tilde en la Key del .resx
        Dim txtImporte As String = If(resManager.GetString("Importe"), "Importe") & " (" & vMoneda & "):"
        Dim txtSaldo As String = If(resManager.GetString("Saldo"), "Saldo") & " (" & vMoneda & "):"
        Dim txtCuenta As String = If(resManager.GetString("Cuenta"), "Cuenta") & ":"

        ' --- CASO 1: Listado Completo por Fechas ---
        If RadioButton1.Checked = True Then
            e.Graphics.DrawString(txtFecha, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(txtConcepto, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
            e.Graphics.DrawString(txtDescripcion, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
            e.Graphics.DrawString(txtImporte, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto4.Left, frmImprimirForm.Punto4.Top - 30)
            e.Graphics.DrawString(txtSaldo, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)
        End If

        ' --- CASO 2: Listado Agrupado por Conceptos ---
        ' 🚀 ALINEACIÓN CLAVE: Mapeamos la columna económica en el Punto 2 de forma fija 
        ' para que el bucle de impresión coincida al milímetro con el renderizado físico.
        If RadioButton2.Checked = True Then
            e.Graphics.DrawString(txtConcepto, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(txtImporte, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto1.Top - 30)
        End If

        ' --- CASO 3: Listado Agrupado por Cuentas Bancarias ---
        If RadioButton3.Checked = True Then
            e.Graphics.DrawString(txtCuenta, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(txtImporte, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto1.Top - 30)
        End If

        ' --- CASO 4: Listado Agrupado por Bloques de Fechas ---
        If RadioButton4.Checked = True Then
            e.Graphics.DrawString(txtFecha, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
            e.Graphics.DrawString(txtImporte, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto1.Top - 30)
        End If

        ' Imprimimos la linea debajo de los encabezados
        ' *********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        ' Imprimimos los detalles del reporte, es decir el listado de Apuntes
        ' *******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left ' Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top  ' Tomamos la posicion vertical de la letra 'Punto1'

        ' Fijamos una altura de salto constante y elegante basada en tu fuente de detalles
        Dim alturaFila As Integer = CInt(FuenteDetalles.Height * 1.5)

        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If frmImprimirForm.DgvApuntes.Rows(PrintLine).IsNewRow Then
                PrintLine += 1
                Continue Do
            End If

            ' Cortafuegos de salto de página seguro si superamos el margen inferior del folio
            If startY + (alturaFila * 3) > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' -------------------------------------------------------------------------
            ' 🌟 CASO 1: Listado Completo por Fechas (Saneado, Traducido y Relacional)
            ' -------------------------------------------------------------------------
            If RadioButton1.Checked = True Then
                ' 🚀 A. DECAPITACIÓN DE HORAS: Convertimos de forma segura a Date y extraemos solo el día corto
                Dim fechaVisual As String = ""
                If frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value IsNot Nothing Then
                    Dim fechaTemp As Date
                    If Date.TryParse(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString(), fechaTemp) Then
                        fechaVisual = fechaTemp.ToShortDateString() ' 🔴 ¡ÉXITO!: Elimina el 12:00:00 AM
                    Else
                        fechaVisual = frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value.ToString()
                    End If
                End If

                ' 🚀 B. PUENTE RELACIONAL PARA EL CONCEPTO (De ID a Nombre Corto Traducido)
                Dim idConceptoTexto As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "").Trim()
                Dim conceptoVisual As String = idConceptoTexto
                Dim idConceptoNum As Integer = 0

                If Integer.TryParse(idConceptoTexto, idConceptoNum) Then
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT CodigoCON FROM conceptos WHERE IdConceptoCON = ?", con)
                            cmd.Parameters.Add("@id", OleDbType.Integer).Value = idConceptoNum
                            Try
                                con.Open()
                                Dim res = cmd.ExecuteScalar()
                                If res IsNot Nothing Then conceptoVisual = res.ToString().Trim()
                            Catch
                            End Try
                        End Using
                    End Using
                End If

                ' Traducción automática del Concepto al inglés (o al idioma de la sesión)
                If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(conceptoVisual) Then
                    Dim claveRecurso As String = conceptoVisual.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)
                    If Not String.IsNullOrEmpty(traduccion) Then conceptoVisual = traduccion.Trim()
                End If
                conceptoVisual = conceptoVisual.Replace("_", " ").ToUpper()

                ' Capturamos la descripción larga original
                Dim valCelda2 As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value?.ToString(), "")

                ' Formateo de importes económicos
                Dim impNum As Double = 0
                Dim salNum As Double = 0
                Double.TryParse(If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value?.ToString(), "0"), impNum)
                Double.TryParse(If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(4).Value?.ToString(), "0"), salNum)

                ' 🚀 IMPRESIÓN DEL CUERPO PRINCIPAL DE LA FILA
                e.Graphics.DrawString(fechaVisual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(conceptoVisual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
                e.Graphics.DrawString(valCelda2, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
                e.Graphics.DrawString(impNum.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto4.Right + 50, startY, sf)
                e.Graphics.DrawString(salNum.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 30, startY, sf)

                ' 🚀 C. PUENTE RELACIONAL PARA LA CUENTA BANCARIA (De ID a Nombre de Banco)
                Dim idCuentaTexto As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(6).Value?.ToString(), "").Trim()
                Dim cuentaVisual As String = idCuentaTexto
                Dim idCuentaNum As Integer = 0

                If Integer.TryParse(idCuentaTexto, idCuentaNum) Then
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT NombreCUE FROM cuentas WHERE IdCuentaCUE = ?", con)
                            cmd.Parameters.Add("@id", OleDbType.Integer).Value = idCuentaNum
                            Try
                                con.Open()
                                Dim res = cmd.ExecuteScalar()
                                If res IsNot Nothing Then cuentaVisual = res.ToString().Trim()
                            Catch
                            End Try
                        End Using
                    End Using
                End If

                ' Traducción del Banco si es una cuenta universal (ej: Efectivo Casa)
                If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(cuentaVisual) Then
                    Dim claveBase As String = cuentaVisual.Replace(" ", "_")
                    Dim tradCuenta As String = resManager.GetString("Desc_" & claveBase)
                    If String.IsNullOrEmpty(tradCuenta) Then tradCuenta = resManager.GetString(claveBase)
                    If Not String.IsNullOrEmpty(tradCuenta) Then cuentaVisual = tradCuenta.Trim()
                End If
                cuentaVisual = cuentaVisual.Replace("_", " ").ToUpper()

                ' Imprimimos la fila de la Cuenta Bancaria
                startY += alturaFila
                e.Graphics.DrawString(resManager.GetString("Cuenta") & ": ", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(cuentaVisual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 65, startY)

                ' Imprimimos la fila de las Notas
                startY += alturaFila
                e.Graphics.DrawString(resManager.GetString("Notas") & ": ", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(5).Value?.ToString(), ""), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 65, startY)

                ' Línea divisoria de final de asiento
                startY += alturaFila
                e.Graphics.DrawString("---------------------------------------------------------------------------------------------------------------------------------------------------------------------", FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                startY += alturaFila
            End If


            ' -------------------------------------------------------------------------
            ' 🌟 CASO 2: Listado Agrupado por Conceptos (¡De ID a Texto Traducido!)
            ' -------------------------------------------------------------------------
            If RadioButton2.Checked = True Then
                ' 1. CAPTURAMOS EL ID RELACIONAL QUE VIENE DEL DATASET (ej: "1" o "54")
                Dim idConceptoTexto As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString(), "").Trim()
                Dim conceptoVisual As String = idConceptoTexto ' Salvavidas por si acaso
                Dim idConceptoNum As Integer = 0

                ' 2. 🚀 TRUCO MAESTRO RELACIONAL: Buscamos el código corto en el maestro de conceptos
                If Integer.TryParse(idConceptoTexto, idConceptoNum) Then
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT CodigoCON FROM conceptos WHERE IdConceptoCON = ?", con)
                            cmd.Parameters.Add("@id", OleDbType.Integer).Value = idConceptoNum
                            Try
                                con.Open()
                                Dim res = cmd.ExecuteScalar()
                                If res IsNot Nothing Then conceptoVisual = res.ToString().Trim()
                            Catch
                            End Try
                        End Using
                    End Using
                End If

                ' 3. 🚀 TRADUCCIÓN AUTOMÁTICA EN CALIENTE (Muta al catalán, alemán o inglés en vivo)
                If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(conceptoVisual) Then
                    Dim claveRecurso As String = conceptoVisual.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)
                    If Not String.IsNullOrEmpty(traduccion) Then conceptoVisual = traduccion.Trim()
                End If

                ' 4. LIMPIEZA VISUAL DE GUIONES: Formateamos "PENSIO_ES" en "PENSIO ES" en mayúsculas contables
                conceptoVisual = conceptoVisual.Replace("_", " ").Trim().ToUpper()

                ' Extraemos el importe acumulado puro
                Dim importeTxt As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "0")
                Dim impNum As Double = 0
                Double.TryParse(importeTxt, impNum)

                ' 🖨️ IMPRESIÓN CORREGIDA EN EL FOLIO
                e.Graphics.DrawString(conceptoVisual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(impNum.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY, sf)

                startY += alturaFila
            End If

            '' =========================================================================
            '' 🌟 TRAMO 3 CORREGIDO: PROCESADO INDEPENDIENTE DE CONCEPTOS (RB2) Y CUENTAS (RB3)
            '' =========================================================================
            'If RadioButton1.Checked = True Then
            '    LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "1")
            '    vIngresos = 0
            '    vGastos = 0
            '    vValor = 0
            '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            '        If fila.IsNewRow Then Continue For
            '        vSaldo = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value)) + vValor
            '        fila.Cells(4).Value = vSaldo
            '        vValor = fila.Cells(4).Value
            '        If CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value)) >= 0 Then
            '            vIngresos += CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))
            '        Else
            '            vGastos += CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))
            '        End If
            '        frmImprimirForm.LblTotal.Text = resManager.GetString("TotalIngresos") & ": " & vIngresos.ToString("N2") & "  -  " & resManager.GetString("TotalGastos") & ": " & vGastos.ToString("N2") & "                        " & resManager.GetString("TOTAL") & " :  " & vValor.ToString("N2") & vMoneda
            '    Next
            'End If

            '' --- 🚀 LOGICA EXCLUSIVA PARA EL RADIOBUTTON 2 (Conceptos Puros) ---
            'If RadioButton2.Checked = True Then
            '    'vtipoSql += " ORDER BY apuper.ConceptoAPP ASC"
            '    MsgBox("Se ha detectado un error en la lógica de agrupación por conceptos. Se recomienda revisar la función LlenarGrid y la construcción de vtipoSql para asegurar que los conceptos se muestren correctamente.", MsgBoxStyle.Critical, "Error de Agrupación")
            '    LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")

            '    'Dim vTempapu As String
            '    Dim vImporteConcepto As Double
            '    Dim vNombreConcepto As String = ""

            '    cmdMdb1cr.CommandText = "DELETE FROM tempapu"
            '    Try
            '        cmdMdb1cr.ExecuteNonQuery()
            '    Catch ex As Exception
            '        MsgBox(ex.ToString())
            '    End Try

            '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            '        If fila.IsNewRow Then Continue For
            '        vImporteConcepto = CDbl(ConvertirDecimalSeguro(fila.Cells(3).Value))

            '        ' 🎯 Leemos la celda 1 (Concepto) de forma estricta
            '        If vNombreConcepto <> fila.Cells(1).Value.ToString() Then
            '            vNombreConcepto = fila.Cells(1).Value.ToString().Trim()

            '            cmdMdb1cr.CommandText = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
            '            cmdMdb1cr.Parameters.Clear()
            '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
            '            Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
            '            paramImpTemp.Value = Math.Round(CDbl(ConvertirDecimalSeguro(vImporteConcepto)), 2)
            '            Try
            '                cmdMdb1cr.ExecuteNonQuery()
            '            Catch ex As Exception
            '                MsgBox(ex.ToString())
            '            End Try
            '        Else
            '            cmdMdb1cr.CommandText = "SELECT SumaImporteAPU FROM tempapu WHERE tempapu.ConceptoAPU = ?"
            '            cmdMdb1cr.Parameters.Clear()
            '            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto)
            '            Dim vExistente As Double = 0
            '            Try
            '                Using drMdb1 As OleDbDataReader = cmdMdb1cr.ExecuteReader()
            '                    If drMdb1.Read() AndAlso drMdb1.GetValue(0) IsNot DBNull.Value Then vExistente = Convert.ToDouble(drMdb1.GetValue(0))
            '                End Using
            '            Catch
            '            End Try

            '            cmdMdb1cr.CommandText = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ?"
            '            cmdMdb1cr.Parameters.Clear()
            '            Dim paramSumaTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
            '            paramSumaTemp.Value = Math.Round(vImporteConcepto + vExistente, 2)
            '            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
            '            Try
            '                cmdMdb1cr.ExecuteNonQuery()
            '            Catch ex As Exception
            '                MsgBox(ex.ToString())
            '            End Try
            '        End If
            '    Next

            '    ' =========================================================================
            '    ' 🌟 LECTURA CORREGIDA DE LA TABLA TEMPORAL (¡Inmune a desbordamientos!)
            '    ' =========================================================================
            '    vtipoSql = "SELECT * FROM tempapu ORDER BY tempapu.ConceptoAPP ASC"
            '    LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "0")

            '    vValor = 0
            '    For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            '        If fila.IsNewRow Then Continue For

            '        ' 🚀 LA CORRECCIÓN: Leemos Cells(1) que es donde viaja la SumaImporteAPU en tempapu
            '        Dim importeAcumulado As Double = 0
            '        If fila.Cells(1).Value IsNot Nothing Then
            '            Double.TryParse(ConvertirDecimalSeguro(fila.Cells(1).Value), importeAcumulado)
            '        End If

            '        vValor += importeAcumulado
            '        frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ":  " & vValor.ToString("N2") & vMoneda
            '    Next
            'End If

            ' -------------------------------------------------------------------------
            ' 🌟 CASO 3 CONFIGURADO: CLON EXACTO DE CONCEPTOS ADAPTADO A CUENTAS (Print)
            ' -------------------------------------------------------------------------
            If RadioButton3.Checked = True Then
                ' 1. CAPTURAMOS EL ID RELACIONAL DE LA CUENTA DESDE EL REPORTE CONSOLIDADO (Celda 0)
                Dim idCuentaTexto As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString(), "").Trim()
                Dim cuentaVisual As String = idCuentaTexto ' Salvavidas por defecto
                Dim idCuentaNum As Integer = 0

                ' 2. 🚀 TRUCO MAESTRO RELACIONAL: Buscamos el nombre legible en el maestro de cuentas
                If Integer.TryParse(idCuentaTexto, idCuentaNum) Then
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT NombreCUE FROM cuentas WHERE IdCuentaCUE = ?", con)
                            cmd.Parameters.Add("@id", OleDbType.Integer).Value = idCuentaNum
                            Try
                                con.Open()
                                Dim res = cmd.ExecuteScalar()
                                If res IsNot Nothing Then cuentaVisual = res.ToString().Trim()
                            Catch
                            End Try
                        End Using
                    End Using
                End If

                ' 3. 🚀 TRADUCCIÓN AUTOMÁTICA EN CALIENTE (Para soporte multiidioma universal)
                If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(cuentaVisual) Then
                    Dim claveRecurso As String = cuentaVisual.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)
                    If Not String.IsNullOrEmpty(traduccion) Then cuentaVisual = traduccion.Trim()
                End If

                ' 4. LIMPIEZA COSMÉTICA FINANCIERA: Mayúsculas puras limpias de guiones bajos
                cuentaVisual = cuentaVisual.Replace("_", " ").Trim().ToUpper()

                ' Extraemos el importe acumulado consolidado por el banco desde la Celda 1
                Dim importeTxt As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "0")
                Dim impNum As Double = 0
                Double.TryParse(importeTxt, impNum)

                ' 🖨️ IMPRESIÓN RECTIFICADA Y COMPACTA EN EL LIENZO
                e.Graphics.DrawString(cuentaVisual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(impNum.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY, sf)

                ' 🚀 SEPARACIÓN VISUAL ENTRE DATOS: Espaciado doble para que las filas respiren elegantemente
                startY += (alturaFila * 2)
            End If

            ' -------------------------------------------------------------------------
            ' 🌟 CASO 4: Listado Agrupado por Bloques de Fechas
            ' -------------------------------------------------------------------------
            If RadioButton4.Checked = True Then
                Dim fechaTxt As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString(), "")
                If fechaTxt.Length > 10 Then fechaTxt = fechaTxt.Substring(0, 10)

                Dim importeTxt As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "0")
                Dim impNum As Double = 0
                Dim unused As Boolean = Double.TryParse(importeTxt, impNum)

                e.Graphics.DrawString(fechaTxt, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
                e.Graphics.DrawString(impNum.ToString("N2"), FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY, sf)
                startY += alturaFila
            End If

            PrintLine += 1
            Contador += 1
        Loop

        ' =========================================================================
        ' 🌟 TRAMO 3 DE 3: IMPRESIÓN DE TOTALES Y SECTOR DE PÁGINAS (¡Cierre del Folio!)
        ' =========================================================================
        ' Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then

            ' --- CASO 1: Listado Completo por Fechas (Tu coordenada original intacta) ---
            If RadioButton1.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 30, startY + 15, sf)
            End If

            ' --- CASO 2: Listado Agrupado por Conceptos (🚀 ALINEADO AL PUNTO 2) ---
            If RadioButton2.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY + 15, sf)
            End If

            ' --- CASO 3: Listado Agrupado por Cuentas Bancarias (🚀 ALINEADO AL PUNTO 2) ---
            If RadioButton3.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY + 15, sf)
            End If

            ' --- CASO 4: Listado Agrupado por Bloques de Fechas (🚀 ALINEADO AL PUNTO 2) ---
            If RadioButton4.Checked = True Then
                e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left + 80, startY + 15, sf)
            End If

            ' Reseteamos punteros de control para que la Vista Previa no se encasquille
            PrintLine = 0
            Contador = 0
        End If

        ' --- SECCIÓN DE PAGINACIÓN MULTIIDIOMA INDESTRUCTIBLE ---
        ' Forzamos la conversión controlando cadenas vacías de forma elástica
        Dim numPagActual As Integer = 0
        Integer.TryParse(frmImprimirForm.LblNumeroPagina.Text, numPagActual)
        numPagActual += 1
        frmImprimirForm.LblNumeroPagina.Text = numPagActual.ToString()

        ' Extraemos de tus diccionarios .resx la etiqueta "Pagina" traducida en vivo
        Dim txtEtiquetaPagina As String = If(resManager.GetString("Pagina"), "Página")

        e.Graphics.DrawString(txtEtiquetaPagina, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

        ' Reseteo preventivo de la paginación para el refresco en caliente de la pantalla
        If Contador = 0 Then
            frmImprimirForm.LblNumeroPagina.Text = "0"
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class