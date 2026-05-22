Imports System.Data
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports MySqlConnector

Public Class EditarCuentaBancaria

    Public vtipoSql, vTxtNombre, vTxtNumero, vTxtTipo, vTxtNotas As String
    Public filaActual As Integer

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(4) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, "Aceptar y Salir")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoCuenta, "Seleccionar el Tipo de Cuenta")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNumero, "Introducir Número de Cuenta o IBAN")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, "Introducir Nombre de la Cuenta")

        ' Llenar el Combo Tipo Cuenta
        '****************************
        cmdMySql1cr.CommandText = "SELECT tipocuentas.CodigoTIP FROM tipocuentas"
        cmdMySql1cr.CommandText += " ORDER BY tipocuentas.CodigoTIP ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbTipoCuenta.Items.Add(drMySql1.GetValue(0))
                End While
                CmbTipoCuenta.Text = CmbTipoCuenta.Items(0)
            Else
                MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Tipo Cuenta")
            MsgBox(ex.ToString)
        End Try

        CmbTipoCuenta.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoCuenta.SelectedIndex = 0

        filaActual = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
        CmbTipoCuenta.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(0).Value.ToString
        TxtNombre.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(1).Value.ToString
        TxtNumero.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(2).Value.ToString
        TxtNota.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(4).Value.ToString

        If vEditar = "SI" Then
            LblEditando.Text = "EDITANDO CUENTA BANCARIA"
            CmbTipoCuenta.Enabled = True
            TxtNombre.Enabled = False
            TxtNumero.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = "¡¡ ELIMINAR CUENTA BANCARIA !!"
            CmbTipoCuenta.Enabled = False
            TxtNombre.Enabled = False
            TxtNumero.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNumero.KeyPress
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

        ' Modificar Registro
        '*******************
        vtipoSql = "UPDATE cuentas SET NumeroCUE = '" & vTxtNumero & "' , TipoCUE = '" & vTxtTipo & "' , NotasCUE = '" & vTxtNotas & "' "
        vtipoSql += " WHERE cuentas.NombreCUE = '" & vTxtNombre & "' "
        cmdMySql1cr.CommandText = vtipoSql

        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            'MsgBox("Registro, Grabado Correctamente")
            Me.Close()
        Catch ex As Exception
            MsgBox("Error al Modificar el Registro")
            MsgBox(ex.ToString)
        End Try
        drMySql1.Close()
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        vTxtNombre = TxtNombre.Text
        respuesta = MsgBox("¿Estas seguro de Eliminar la Cuenta Bancaria '" & vTxtNombre & "' y todos sus Apuntes de Todos los Ejercicios?.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar Cuenta Bancaria")
        If respuesta = vbYes Then
            ' Eliminar Registro Cuentas
            vtipoSql = "DELETE FROM cuentas"
            vtipoSql += " WHERE cuentas.NombreCUE = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox("Registro Cuenta Bancaria, Borrada !!!")
            Catch ex As Exception
                MsgBox("Error al Eliminar el Registro de la Cuenta Bancaria")
                MsgBox(ex.ToString)
            End Try

            ' Eliminar Registros Apuntes
            vtipoSql = "DELETE FROM apuntes"
            vtipoSql += " WHERE apuntes.CuentaAPU = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox("Apuntes, Borrados !!!")
            Catch ex As Exception
                MsgBox("Error al Eliminar los Apuntes de la Cuenta Bancaria")
                MsgBox(ex.ToString)
            End Try

            ' Eliminar Registros Apuntes Periódicos
            vtipoSql = "DELETE FROM apuper"
            vtipoSql += " WHERE apuper.CuentaAPP = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox("Apuntes Periodicos, Borrados !!!")
            Catch ex As Exception
                MsgBox("Error al Eliminar los Apuntes Periódicos de la Cuenta Bancaria")
                MsgBox(ex.ToString)
            End Try
        End If
        Me.Close()
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

End Class