Imports System.Diagnostics
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntesPeriodicos

    Public vConcepto, vtipoSql, vtipoGrid, vAñadirSql As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, vAnexo, vbOK As String
    Public vNumeroPagos, vDate3Year As Integer
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Label7.Text = vMoneda

        vFecha1Enero = Val(vAñoEjercicio)
        ' si el año del ejercicio es diferente al año actual, se pone como fecha máxima el 31 de diciembre del año del ejercicio + 1,
        ' para que se puedan introducir apuntes periódicos con fecha hasta el 31 de diciembre del año siguiente al ejercicio
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio + 1, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If

        Dim TL(12) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, "Ir a Hoy")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.DateTimePicker1, "Seleccionar la Fecha del Primer Pago/Cobro")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptarSalir, "Aceptar y Salir")
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
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, "Añade, Edita, Borra o Consulta Conceptos Contables")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, "Añade, Edita, Borra o Consulta Cuentas Bancarias")
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.TxtNumeroPagos, "Introducir el Número de Pagos/Cobros que se harán en este Apunte Periódico")
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.CmbPeriocidad, "Seleccionar el Periodo entre cada Pago/Cobro")

        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    If drMdb1.GetValue(0) <> "TRASPASO" Then
                        CmbConcepto.Items.Add(drMdb1.GetValue(0))
                    End If
                End While
                If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                    CmbConcepto.Text = CmbConcepto.Items(frmApuntesPeriodicos.CmbConcepto.SelectedIndex)
                Else
                    CmbConcepto.Text = CmbConcepto.Items(0)
                End If
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Concepto")
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
            MsgBox("Error al llenar el Combo Descripción")
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
            MsgBox("Error al llenar el Combo Cuenta")
            MsgBox(ex.ToString)
        End Try
        TxtImporte.Text = 0
        CmbConcepto.Select()
        CmbPeriocidad.Text = CmbPeriocidad.Items(3)
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
        If TxtTipoConcepto.Text = "GASTO" Then
            LblNumeroPagosCobros.Text = "Nº de Pagos:"
            LblFechaPagoCobro.Text = "1er Pago:"
        Else
            LblNumeroPagosCobros.Text = "Nº de Cobros:"
            LblFechaPagoCobro.Text = "1er Cobro:"
        End If
        drMdb1.Close()
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            CmbCuenta.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptarSalir.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptarSalir_Click(sender As Object, e As EventArgs) Handles BtnAceptarSalir.Click
        If CmbDescripcion.Text <> "" Then
            If TxtNumeroPagos.Text <> "" Then
                If TxtImporte.Text <> "" And TxtImporte.Text <> "0" Then
                    'Las veces que se tiene que guardar = Nº de Pagos/Cobros
                    vNumeroPagos = TxtNumeroPagos.Text
                    For i = 0 To vNumeroPagos - 1
                        If CmbPeriocidad.Text = "Diaria" Then
                            vDate3 = DateTimePicker1.Value.Date.AddDays(i)
                        End If
                        If CmbPeriocidad.Text = "Semanal" Then
                            vDate3 = DateTimePicker1.Value.Date.AddDays(i * 7)
                        End If
                        If CmbPeriocidad.Text = "Quincenal" Then
                            vDate3 = DateTimePicker1.Value.Date.AddDays(i * 15)
                        End If
                        If CmbPeriocidad.Text = "Mensual" Then
                            vDate3 = DateTimePicker1.Value.Date.AddMonths(i)
                        End If
                        If CmbPeriocidad.Text = "Bimensual" Then
                            vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 2)
                        End If
                        If CmbPeriocidad.Text = "Trimestral" Then
                            vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 3)
                        End If
                        If CmbPeriocidad.Text = "Semestral" Then
                            vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 6)
                        End If
                        If CmbPeriocidad.Text = "Anual" Then
                            vDate3 = DateTimePicker1.Value.Date.AddYears(i)
                        End If
                        vDate3Year = Date.Now.Year
                        vAnexo = (i + 1).ToString
                        vDescripcionAPU = CmbDescripcion.Text & "  (" & vAnexo & " de " & vNumeroPagos.ToString & ")".ToString
                        vImporteAPU = TxtImporte.Text
                        If TxtTipoConcepto.Text = "GASTO" Then
                            vImporteAPU = "-" & vImporteAPU.ToString
                        End If
                        vNotasAPU = TxtNota.Text
                        vCuentaAPU = CmbCuenta.Text.ToString
                        vAñadirSql = "INSERT INTO apuper "
                        vAñadirSql += "(FechaAPP, ConceptoAPP, DescripcionAPP, ImporteAPP, EjercicioAPP, NotasAPP, CuentaAPP) "
                        vAñadirSql += "VALUES (#" & vDate3 & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vDate3Year & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                        cmdMdb1cr.CommandText = vAñadirSql
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                            vbOK = "SI"
                        Catch ex As Exception
                            MsgBox("Error al Grabar el Apunte Periódico Nº " & (i + 1).ToString)
                            MsgBox(ex.ToString)
                        End Try
                    Next
                    If vbOK = "SI" Then
                        'MsgBox("Registro, Grabado Correctamente")
                    End If

                    vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
                    vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
                    If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                        vtipoSql += " And apuper.CuentaAPP = '" & frmApuntesPeriodicos.CmbCuenta.Text & "' "
                    End If
                    If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                        vtipoSql += " And apuper.ConceptoAPP = '" & frmApuntesPeriodicos.CmbConcepto.Text & "' "
                    End If
                    If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                        vDate1 = frmApuntesPeriodicos.DateTimePicker1.Value
                        vDate2 = frmApuntesPeriodicos.DateTimePicker2.Value
                        vtipoSql += " And apuper.FechaAPP >= #" & vDate1 & "#"
                        vtipoSql += " And apuper.FechaAPP <= #" & vDate2 & "#"
                    End If
                    vtipoSql += " ORDER BY apuper.FechaAPP ASC"
                    vtipoGrid = "APUNTES_PERIODICOS"
                    LlenarGrid(vtipoSql, vtipoGrid, "1")
                    Me.Close()
                Else
                    MsgBox("NO hay Cantidad en Importe ...", vbExclamation)
                    TxtImporte.Select()
                End If
            Else
                MsgBox("NO hay Número de Pagos/Cobros ...", vbExclamation)
                TxtNumeroPagos.Select()
            End If
        Else
            MsgBox("La Descripción NO puede estar vacia.", vbExclamation)
            CmbDescripcion.Select()
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

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
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

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs) Handles BtnCuenta.Click
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

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub

    Private Sub CmbCuenta_Click(sender As Object, e As EventArgs) Handles CmbCuenta.Click
        CmbCuenta.DroppedDown = True
    End Sub
End Class