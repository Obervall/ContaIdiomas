Imports System.Diagnostics
Imports System.Windows.Forms

Public Class TraspasoCuentas

    Public vConcepto, vAñadirOrigenSql, vAñadirDestinoSql As String
    Public vImporteAPU As Double
    Public vDescripcionAPU, vNotasAPU, vCuentaOrigenAPU, vCuentaDestinoAPU As String
    Public vfechaHoyOrigen As Date = DateTime.Today
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Label7.Text = vMoneda

        Dim TL(11) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoyOrigen, "Ir a Hoy Origen")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptar, "Aceptar y Salir")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.CmbConcepto, "Seleccionar el Concepto a la que se refiere la transacción")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbCuentaOrigen, "Seleccionar la Cuenta Origen a la que se refiere la transacción")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuentaDestino, "Seleccionar la Cuenta Destino a la que se refiere la transacción")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtDescripcion, "Introducir una descripción para el Asiento")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, "Importe del Asiento")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, "Activar la Calculadora")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, "Añade, Edita, Borra o Consulta Conceptos Contables")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuentaOrigen, "Añade, Edita, Borra o Consulta Cuentas Bancarias")
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnCuentaDestino, "Añade, Edita, Borra o Consulta Cuentas Bancarias")


        ' Llenar el Combo Concepto
        '*************************
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    If drMySql1.GetValue(0) = "TRASPASO" Then
                        CmbConcepto.Items.Add(drMySql1.GetValue(0))
                    End If
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al cargar el Combo Concepto: " & ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMySql1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbCuentaOrigen.Items.Add(drMySql1.GetValue(0))
                    CmbCuentaDestino.Items.Add(drMySql1.GetValue(0))
                End While
                CmbCuentaOrigen.Text = CmbCuentaOrigen.Items(0)
                CmbCuentaDestino.Text = CmbCuentaOrigen.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al cargar el Combo Cuenta: " & ex.ToString)
        End Try
        TxtImporte.Text = 0
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMySql1.Close()
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMySql1 = cmdMySql1cr.ExecuteReader()
        drMySql1.Read()
        TxtTipoConcepto.Text = drMySql1.GetValue(2)
        drMySql1.Close()
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtNota.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnHoyOrigen_Click(sender As Object, e As EventArgs) Handles BtnHoyOrigen.Click
        DtpOrigen.Value = vfechaHoyOrigen
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        If CmbCuentaOrigen.Text <> CmbCuentaDestino.Text Then
            If TxtDescripcion.Text <> "" Then
                If TxtImporte.Text <> "" And TxtImporte.Text <> "0" Then
                    vConcepto = CmbConcepto.Text ' & " ORIGEN"
                    vDescripcionAPU = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
                    vImporteAPU = TxtImporte.Text
                    vImporteAPU = "-" & vImporteAPU.ToString
                    vNotasAPU = TxtNota.Text
                    vCuentaOrigenAPU = CmbCuentaOrigen.Text.ToString
                    vAñadirOrigenSql = "INSERT INTO apuntes "
                    vAñadirOrigenSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                    vAñadirOrigenSql += "VALUES (#" & CDate(DtpOrigen.Value).ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaOrigenAPU & "')"
                    cmdMySql1cr.CommandText = vAñadirOrigenSql
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        MsgBox("Registro Origen, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al Grabar el Registro Origen: " & ex.ToString)
                    End Try

                    vConcepto = CmbConcepto.Text '  & " DESTINO"
                    vImporteAPU = TxtImporte.Text
                    vCuentaDestinoAPU = CmbCuentaDestino.Text.ToString
                    vAñadirDestinoSql = "INSERT INTO apuntes "
                    vAñadirDestinoSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                    vAñadirDestinoSql += "VALUES (#" & CDate(DtpDestino.Value).ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaDestinoAPU & "')"
                    cmdMySql1cr.CommandText = vAñadirDestinoSql
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        MsgBox("Registro Destino, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox("Error al Grabar el Registro Destino: " & ex.ToString)
                    End Try
                    Me.Close()
                Else
                    MsgBox("NO hay Cantidad en Importe ...", vbExclamation)
                    TxtImporte.Select()
                End If
            Else
                MsgBox("La Descripción NO puede estar vacia.", vbExclamation)
            End If
        Else
            MsgBox("Las cuentas de Origen y Destino tienen que ser Diferentes.", vbExclamation)
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCuentaOrigen_Click(sender As Object, e As EventArgs) Handles BtnCuentaOrigen.Click
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

    Private Sub BtnCuentaDestino_Click(sender As Object, e As EventArgs) Handles BtnCuentaDestino.Click
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

    Private Sub DtpOrigen_ValueChanged(sender As Object, e As EventArgs) Handles DtpOrigen.ValueChanged
        DtpDestino.Value = DtpOrigen.Value
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
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

    Private Sub CmbCuentaOrigen_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuentaOrigen.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuentaDestino_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuentaDestino.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class