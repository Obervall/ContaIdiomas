Imports System.Windows.Forms

Public Class EditarTipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtDescripcion As String
    Public filaActual As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

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
        ' 1. Capturamos los textos limpios de los cuadros del formulario
        vTxtNombre = TxtNombre.Text.Trim()
        vTxtDescripcion = TxtDescripcion.Text.Trim() ' Los parámetros gestionan comillas y apóstrofes solos

        ' =========================================================================
        ' 2. OBTENER EL ID NUMÉRICO REAL DESDE EL GRID DE LA PANTALLA ANTERIOR
        ' =========================================================================
        Dim idTipoModificar As Integer = 0
        Try
            Dim filaActual As Integer = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
            ' Recuperamos el Id numérico que viaja en la Celda 2 del Grid maestro
            idTipoModificar = Convert.ToInt32(frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(2).Value)
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorRecuperarID"), resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' =========================================================================
        ' 3. CONFIGURAR EL UPDATE USANDO EL ID NUMÉRICO (MÁXIMA CONSISTENCIA)
        ' =========================================================================
        ' Cambiamos el filtro WHERE para que busque estrictamente por el IdTipoCUE
        vtipoSql = "UPDATE tipocuentas SET DescripcionTIP = ? WHERE IdTipoCUE = ?"
        cmdMdb1cr.CommandText = vtipoSql

        ' En Access/OleDb el orden de los parámetros debe ser EXACTO al del SQL
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionTIP", vTxtDescripcion)
        cmdMdb1cr.Parameters.AddWithValue("@IdTipoCUE", idTipoModificar) ' Filtro WHERE numérico

        Try
            Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()

            If filasAfectadas > 0 Then
                Me.Close() ' Guardado con éxito, cierra la ventana modal
            Else
                MessageBox.Show(resManager.GetString("NoEncuentraRegistro"), resManager.GetString("Atencion"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorModificarRegistro") & ": " & vbNewLine & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
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