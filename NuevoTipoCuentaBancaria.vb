Imports System.Data
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports MySqlConnector

Public Class NuevoTipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtDescripcion As String

    Private Sub NuevaCuentaBancaria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(3) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, "Aceptar y Salir")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.TxtDescripcion, "Introducir Descripción del Tipo de Cuenta Bancaria")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNombre, "Introducir Nombre del Tipo de Cuenta Bancaria")

        TxtNombre.Select()
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        Dim vBusca As String
        vBusca = TxtNombre.Text.ToString
        DgvExistente.Visible = True

        ' Llenar Grid de Nombre/Código EXISTENTES en TIPO CUENTAS BANCARIAS
        '******************************************************************
        vtipoSql = "SELECT tipocuentas.CodigoTIP "
        vtipoSql += "FROM tipocuentas WHERE tipocuentas.CodigoTIP Like '" & vBusca & "%' ORDER BY tipocuentas.CodigoTIP"
        vtipoGrid = "NOMBRESEXISTENTES3"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtDescripcion.Select()
        End If
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtDescripcion.Select()
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        If TxtNombre.Text <> "" Then
            vTxtNombre = TxtNombre.Text
            vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)

            ' Verificar que no se repite Nombre/Código en Cuentas Bancarias
            '**************************************************************
            vtipoSql = "SELECT * FROM tipocuentas WHERE tipocuentas.CodigoTIP = '" & vTxtNombre & "' "
            vtipoGrid = "NOMBRESEXISTENTES3"
            cmdMdb1cr.CommandText = vtipoSql

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    drMdb1.Close()
                    MsgBox("El Nombre: " & vTxtNombre & ", ya existe en Tipo Cuentas Bancarias", vbOKOnly, "Tipo Cuenta Existente")
                    TxtNombre.Select()
                Else
                    drMdb1.Close()
                    vtipoSql = "INSERT INTO tipocuentas "
                    vtipoSql += "(CodigoTIP, DescripcionTIP) "
                    vtipoSql += "VALUES ('" & vTxtNombre & "','" & vTxtDescripcion & "')"

                    cmdMdb1cr.CommandText = vtipoSql
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro, Grabado Correctamente")
                        Me.Close()
                    Catch ex As Exception
                        MsgBox("Error al insertar el nuevo Tipo de Cuenta Bancaria")
                        MsgBox(ex.ToString)
                    End Try
                End If
            Catch ex As Exception
                MsgBox("Error al verificar que el Nombre no se repite en Tipo Cuentas Bancarias")
                MsgBox(ex.ToString)
            End Try
        Else
            MsgBox("NO hay Datos en Nombre ...")
            TxtNombre.Select()
        End If
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