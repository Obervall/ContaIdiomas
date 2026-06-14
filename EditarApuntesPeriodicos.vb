Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class EditarApuntesPeriodicos

    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU As String
    Public vCodigoAPU As Integer
    Public vimporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        ActualizarTextosFormulario(Me)

        Label7.Text = vMoneda
        Dim TL(8) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, "Ir a Hoy")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEliminar, "Eliminar Registro")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, "Aceptar")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, "Seleccionar el Concepto a la que se refiere la transacción")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, "Seleccionar la Cuenta a la que se refiere la transacción")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, "Introducir una descripción para el Asiento")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, "Importe del Asiento")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, "Activar la Calculadora")

        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbConcepto.Items.Add(drMdb1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Descripción
        '****************************
        cmdMdb1cr.CommandText = "SELECT * FROM apuntes ORDER BY apuntes.DescripcionAPU ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                primero = 1
                While drMdb1.Read()
                    If Trim(drMdb1.GetValue(3)) <> "Saldo Inicial" Then
                        If primero = 1 Then
                            CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                            primero = 2
                        Else
                            nuevo = 0
                            For i = 0 To CmbDescripcion.Items.Count - 1
                                If Trim(drMdb1.GetValue(3)) = Trim(CmbDescripcion.Items(i)) Then
                                    nuevo = 0
                                    Exit For
                                Else
                                    nuevo = 1
                                End If
                            Next
                            If nuevo = 1 Then
                                CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                                nuevo = 0
                            End If
                        End If
                    End If
                End While
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMdb1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuenta.Items.Add(drMdb1.GetValue(0))
                End While
                CmbCuenta.Text = CmbCuenta.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
        DateTimePicker1.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(0).Value.ToString
        CmbConcepto.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(1).Value.ToString
        CmbDescripcion.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(2).Value.ToString
        vimporteAPU = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(3).Value.ToString
        vimporteAPU = vimporteAPU
        TxtImporte.Text = Math.Abs(vimporteAPU).ToString("N2")
        TxtNota.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(5).Value.ToString
        CmbCuenta.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(6).Value.ToString
        vCodigoAPU = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(7).Value

        If vEditar = "SI" Then
            LblEditando.Text = "EDITANDO APUNTE PERIODICO"
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = "¡¡ ELIMINAR APUNTE PERIODICO !!"
            DateTimePicker1.Enabled = False
            CmbConcepto.Enabled = False
            CmbDescripcion.Enabled = False
            TxtImporte.Enabled = False
            CmbCuenta.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("SeguroEliminarRegistro"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("$this.Text"))
        If respuesta = vbYes Then
            ' Eliminar Registro Apunte
            vtipoSql = "DELETE FROM apuper"
            vtipoSql += " WHERE apuper.CodigoAPP = " & vCodigoAPU.ToString
            cmdMdb1cr.CommandText = vtipoSql

            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("RegistroApuntePeriódicoBorrado"))
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        Else
            frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Selected = True
            frmApuntesPeriodicos.DgvApuper.CurrentCell = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(0)
        End If
        Me.Close()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged

        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMdb1.Close()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMdb1 = cmdMdb1cr.ExecuteReader()
        drMdb1.Read()
        TxtTipoConcepto.Text = drMdb1.GetValue(2)
        CmbDescripcion.Text = drMdb1.GetValue(1)
        drMdb1.Close()
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            CmbCuenta.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
        If TxtImporte.Text = "" Then
            TxtImporte.Text = 0
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Normalización inicial del cuadro de texto
        If String.IsNullOrWhiteSpace(TxtImporte.Text) Then TxtImporte.Text = "0"

        ' 2. Extraemos el importe de forma segura usando tu función del módulo
        Dim importeDecimal As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)

        If importeDecimal <> 0 Then
            ' 3. Aplicamos el signo aritmético puro (sin concatenar texto "-")
            If TxtTipoConcepto.Text = "GASTO" Then
                vimporteAPU = -Math.Abs(importeDecimal)
            Else
                vimporteAPU = Math.Abs(importeDecimal)
            End If

            vDate3 = DateTimePicker1.Value.Date

            ' --- FASE 1: EJECUCIÓN DEL UPDATE PARAMETRIZADO ---
            vtipoSql = "UPDATE apuper SET FechaAPP = ?, ConceptoAPP = ?, DescripcionAPP = ?, ImporteAPP = ?, CuentaAPP = ?, NotasAPP = ? " &
                       "WHERE apuper.CodigoAPP = ?"
            cmdMdb1cr.CommandText = vtipoSql

            ' Los parámetros de Access se asocian estrictamente por el orden de los comodines '?'
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("@FechaAPP", vDate3)
            cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPP", CmbConcepto.Text.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@DescripcionAPP", CmbDescripcion.Text.Trim())

            ' Forzamos formato Currency para evitar conflictos de precisión decimal en Access
            Dim paramImp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteAPP", OleDb.OleDbType.Currency)
            paramImp.Value = Math.Round(vimporteAPU, 2)

            cmdMdb1cr.Parameters.AddWithValue("@CuentaAPP", CmbCuenta.Text.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@NotasAPP", TxtNota.Text.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@CodigoAPP", CInt(vCodigoAPU))

            Try
                cmdMdb1cr.ExecuteNonQuery()
                Me.Close()
            Catch ex As Exception
                MessageBox.Show("Error al actualizar registro: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub ' Si falla el guardado, detenemos el flujo para no romper la grilla
            End Try


            ' --- FASE 2: REFRESCO DE LA GRILLA DINÁMICA Y PARAMETRIZADA ---
            ' Iniciamos la consulta base
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper " &
                       "WHERE apuper.EjercicioAPP = ?"

            ' Preparamos una lista temporal para guardar los valores de los parámetros en orden de aparición
            Dim valoresFiltros As New List(Of Object)()
            valoresFiltros.Add(CInt(vAñoEjercicio)) ' El primer '?' corresponde al Ejercicio

            ' Filtro Cuenta
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                vtipoSql += " And apuper.CuentaAPP = ?"
                valoresFiltros.Add(frmApuntesPeriodicos.CmbCuenta.Text.Trim())
            End If

            ' Filtro Concepto
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuper.ConceptoAPP = ?"
                valoresFiltros.Add(frmApuntesPeriodicos.CmbConcepto.Text.Trim())
            End If

            ' Filtro Fechas (¡Adiós almohadillas!)
            If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                vDate1 = frmApuntesPeriodicos.DateTimePicker1.Value.Date
                vDate2 = frmApuntesPeriodicos.DateTimePicker2.Value.Date
                vtipoSql += " And apuper.FechaAPP >= ?"
                vtipoSql += " And apuper.FechaAPP <= ?"
                valoresFiltros.Add(vDate1)
                valoresFiltros.Add(vDate2)
            End If

            vtipoSql += " ORDER BY apuper.FechaAPP ASC, apuper.ImporteAPP ASC"
            cmdMdb1cr.CommandText = vtipoSql

            ' Inyectamos los parámetros en el comando siguiendo el orden secuencial exacto
            cmdMdb1cr.Parameters.Clear()
            For idx As Integer = 0 To valoresFiltros.Count - 1
                cmdMdb1cr.Parameters.AddWithValue("@P" & idx, valoresFiltros(idx))
            Next

            ' Cargamos los datos limpios en el Grid
            vtipoGrid = "APUNTES_PERIODICOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")

            ' Reposicionamos la fila seleccionada por el usuario de forma segura
            If frmApuntesPeriodicos.DgvApuper.Rows.Count > 0 AndAlso vFilaActual < frmApuntesPeriodicos.DgvApuper.Rows.Count Then
                frmApuntesPeriodicos.DgvApuper.Rows(vFilaActual).Selected = True
                frmApuntesPeriodicos.DgvApuper.CurrentCell = frmApuntesPeriodicos.DgvApuper.Rows(vFilaActual).Cells(0)
            End If
        Else
            MessageBox.Show(frmIntroApuntes.rmse.GetString("NoQuantityAmount"), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtImporte.Select()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs)
        frmPrincipal.TsLabelFormulario.Text = "Conceptos Contables"

        ' Comprobamos si existe un identificador asociado.
        If ((frmConceptosContables Is Nothing) OrElse (Not frmConceptosContables.IsHandleCreated)) Then
            frmConceptosContables = New ConceptosContables
        End If

        ' Llamamos al formulario de manera modal.
        frmConceptosContables.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmConceptosContables.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs)
        frmPrincipal.TsLabelFormulario.Text = "Cuentas Bancarias"

        ' Comprobamos si existe un identificador asociado.
        If ((frmCuentasBancarias Is Nothing) OrElse (Not frmCuentasBancarias.IsHandleCreated)) Then
            frmCuentasBancarias = New CuentasBancarias
        End If

        ' Llamamos al formulario de manera modal.
        frmCuentasBancarias.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmCuentasBancarias.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class