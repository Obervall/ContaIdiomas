Imports System.Diagnostics
Imports System.Windows.Forms

Public Class TraspasoCuentas

    Public vConcepto, vAñadirOrigenSql, vAñadirDestinoSql As String
    Public vImporteAPU As Double
    Public vDescripcionAPU, vNotasAPU, vCuentaOrigenAPU, vCuentaDestinoAPU As String
    Public vfechaHoyOrigen As Date = DateTime.Today
    Private TL(11) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        ActualizarTextosFormulario(Me)

        Label7.Text = vMoneda
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoyOrigen, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.CmbConcepto, rmse.GetString("ToolTipSeleccionarConcepto"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbCuentaOrigen, rmse.GetString("SeleccionarCuentaOrigen"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuentaDestino, rmse.GetString("SeleccionarCuentaDestino"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtDescripcion, rmse.GetString("IntroDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteApunte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuentaOrigen, resManager.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnCuentaDestino, resManager.GetString("BtnCuenta"))

        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    If drMdb1.GetValue(0) = "TRASPASO" Then
                        CmbConcepto.Items.Add(drMdb1.GetValue(0))
                    End If
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMdb1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuentaOrigen.Items.Add(drMdb1.GetValue(0))
                    CmbCuentaDestino.Items.Add(drMdb1.GetValue(0))
                End While
                CmbCuentaOrigen.Text = CmbCuentaOrigen.Items(0)
                CmbCuentaDestino.Text = CmbCuentaOrigen.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al cargar el Combo Cuenta: " & ex.ToString)
        End Try
        TxtImporte.Text = 0
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
        drMdb1.Close()
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
                    cmdMdb1cr.CommandText = vAñadirOrigenSql
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        MsgBox(rmse.GetString("RegistroOrigenGrabadoCorrectamente"))
                    Catch ex As Exception
                        MsgBox(resManager.GetString("Error") & ": " & ex.ToString)
                    End Try

                    vConcepto = CmbConcepto.Text '  & " DESTINO"
                    vImporteAPU = TxtImporte.Text
                    vCuentaDestinoAPU = CmbCuentaDestino.Text.ToString
                    vAñadirDestinoSql = "INSERT INTO apuntes "
                    vAñadirDestinoSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                    vAñadirDestinoSql += "VALUES (#" & CDate(DtpDestino.Value).ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaDestinoAPU & "')"
                    cmdMdb1cr.CommandText = vAñadirDestinoSql
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        MsgBox(rmse.GetString("RegistroDestinoGrabadoCorrectamente"))
                    Catch ex As Exception
                        MsgBox(resManager.GetString("Error") & ": " & ex.ToString)
                    End Try
                    Me.Close()
                Else
                    MsgBox(rmse.GetString("NoHayImporte") & "...", vbExclamation)
                    TxtImporte.Select()
                End If
            Else
                MsgBox(rmse.GetString("DescripcionVacia"), vbExclamation)
            End If
        Else
            MsgBox(rmse.GetString("CuentasDiferentes"), vbExclamation)
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCuentaOrigen_Click(sender As Object, e As EventArgs) Handles BtnCuentaOrigen.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnCuentaDestino_Click(sender As Object, e As EventArgs) Handles BtnCuentaDestino.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()
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