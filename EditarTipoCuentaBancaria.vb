Imports System.Data
Imports System.Diagnostics
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports MySqlConnector

Public Class EditarTipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtDescripcion As String

    Public filaActual As Integer

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

        filaActual = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        TxtNombre.Text = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString
        TxtDescripcion.Text = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(1).Value.ToString

        If vEditar = "SI" Then
            LblEditando.Text = "EDITANDO TIPO CUENTA BANCARIA"
            TxtNombre.Enabled = False
            TxtDescripcion.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = "¡¡ ELIMINAR TIPO CUENTA BANCARIA !!"
            TxtNombre.Enabled = False
            TxtDescripcion.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vTxtNombre = TxtNombre.Text
        vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)

        ' Modificar Registro
        '*******************
        vtipoSql = "UPDATE tipocuentas SET DescripcionTIP = '" & vTxtDescripcion & "' "
        vtipoSql += " WHERE tipocuentas.CodigoTIP = '" & vTxtNombre & "' "
        cmdMdb1cr.CommandText = vtipoSql

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            'MsgBox("Registro, Grabado Correctamente")
            Me.Close()
        Catch ex As Exception
            MsgBox("Error al Modificar el Registro")
            MsgBox(ex.ToString)
        End Try
        drMdb1.Close()
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        MsgBox("No se pueden Eliminar Cuentas Bancarias")
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