Imports System.Data
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports MySqlConnector

Public Class NuevaCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtNumero, vTxtTipo, vTxtNotas As String
    Public TL(4) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub NuevaCuentaBancaria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(4) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoCuenta, resManager.GetString("ToolTipTipoCuenta"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNumero, resManager.GetString("ToolTipIBAN"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, resManager.GetString("ToolTipNombre"))

        ' Aunque el combo se llame diferente en este formulario, funcionará igual:
        CargarComboTipoCuentaGlobal(Me.CmbTipoCuenta, rmse)
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        TxtNombre.Text = TxtNombre.Text.ToUpper
        TxtNombre.SelectionStart = Len(TxtNombre.Text)
        Dim vBusca As String
        vBusca = TxtNombre.Text.ToString
        DgvExistente.Visible = True

        ' Llenar Grid de Nombre/Código EXISTENTES en CUENTAS BANCARIAS
        '*************************************************************
        vtipoSql = "SELECT cuentas.NombreCUE "
        vtipoSql += "FROM cuentas WHERE cuentas.NombreCUE Like '" & vBusca & "%' ORDER BY cuentas.NombreCUE"
        vtipoGrid = "NOMBRESEXISTENTES2"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtNumero.Select()
        End If
    End Sub

    Private Sub TxtNumero_TextChanged(sender As Object, e As EventArgs) Handles TxtNumero.TextChanged
        TxtNumero.Text = TxtNumero.Text.ToUpper
        TxtNumero.SelectionStart = Len(TxtNumero.Text)
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtNumero.Select()
    End Sub

    Private Sub TxtNumero_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNumero.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtNota.Select()
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vTxtNombre = TxtNombre.Text
        vTxtNumero = TxtNumero.Text
        vTxtTipo = CmbTipoCuenta.Text
        vTxtNotas = TxtNota.Text

        ' Verificar que no se repite Nombre/Código en Cuentas Bancarias
        '**************************************************************
        vtipoSql = "SELECT * FROM cuentas WHERE cuentas.NombreCUE = '" & vTxtNombre & "' "
        cmdMdb1cr.CommandText = vtipoSql

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                drMdb1.Close()
                MsgBox("El Nombre: " & vTxtNombre & ", ya existe en Cuentas Bancarias", vbOKOnly, "Cuenta Existente")
                TxtNombre.Select()
            Else
                drMdb1.Close()
                vtipoSql = "INSERT INTO cuentas "
                vtipoSql += "(NombreCUE, NumeroCUE, TipoCUE, NotasCUE) "
                vtipoSql += "VALUES ('" & vTxtNombre & "','" & vTxtNumero & "','" & vTxtTipo & "','" & vTxtNotas & "')"

                cmdMdb1cr.CommandText = vtipoSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                    Me.Close()
                Catch ex As Exception
                    MsgBox("Error al Grabar la Nueva Cuenta Bancaria")
                    MsgBox(ex.ToString)
                End Try
            End If
        Catch ex As Exception
            MsgBox("Error al verificar que el Nombre no se repite en Cuentas Bancarias")
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtNombre_GotFocus(sender As Object, e As EventArgs) Handles TxtNombre.GotFocus
        PintaTxt()
    End Sub

    Public Sub PintaTxt()
        Dim Texto As TextBox
        Texto = Me.ActiveControl
        Texto.SelectionStart = 0
        Texto.SelectionLength = Texto.Text.Length
    End Sub

End Class