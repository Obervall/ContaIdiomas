Imports System.Windows.Forms

Public Class EditarTipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtDescripcion As String
    Public filaActual As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        ActualizarTextosFormulario(Me)

        Dim TL(3) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.TxtDescripcion, resManager.GetString("Descripcion"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNombre, resManager.GetString("Nombre"))

        filaActual = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        TxtNombre.Text = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString
        TxtDescripcion.Text = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(1).Value.ToString

        If vEditar = "SI" Then
            'LblEditando.Text = "EDITANDO TIPO CUENTA BANCARIA"
            TxtNombre.Enabled = False
            TxtDescripcion.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
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
        ' 1. Limpias la consulta cambiando las comillas y concatenaciones por "?"
        vtipoSql = "UPDATE tipocuentas SET DescripcionTIP = ? WHERE tipocuentas.CodigoTIP = ?"
        cmdMdb1cr.CommandText = vtipoSql

        ' 2. Limpias los parámetros anteriores para que no se acumulen
        cmdMdb1cr.Parameters.Clear()

        ' 3. Añades los nuevos valores en el orden exacto de los "?"
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtDescripcion)
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre)

        ' 4. Tu comando ya está listo para hacer el ExecuteNonQuery como siempre
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Registro, Grabado Correctamente")
            Me.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        MsgBox(rmse.GetString("EliminarCuenta5"))
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