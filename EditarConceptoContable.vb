Imports System.Windows.Forms

Public Class EditarConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtTipo, vTxtNotas As String
    Public filaActual As Integer
    Public TL(2) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.TxtDescripcion, resManager.GetString("ToolTipDescripcion"))

        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoConcepto.SelectedIndex = 0

        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index
        CmbTipoConcepto.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(0).Value.ToString
        TxtNombre.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString
        TxtDescripcion.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(2).Value.ToString
        TxtNota.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(3).Value.ToString

        If vEditar = "SI" Then
            'LblEditando.Text = Por defecto creado en el diseño del formulario, se le asigna el texto de Editando
            CmbTipoConcepto.Enabled = False
            TxtNombre.Enabled = False
            TxtDescripcion.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
            CmbTipoConcepto.Enabled = False
            TxtNombre.Enabled = False
            TxtDescripcion.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
            vEditar = "SI"
        End If
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
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
        vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
        vTxtTipo = CmbTipoConcepto.Text
        vTxtNotas = TxtNota.Text

        ' Modificar Registro
        '*******************
        vtipoSql = "UPDATE conceptos SET DescripcionCON = '" & vTxtDescripcion & "' , NotasCON = '" & vTxtNotas & "' "
        vtipoSql += " WHERE conceptos.CodigoCON = '" & vTxtNombre & "' "
        cmdMySql1cr.CommandText = vtipoSql

        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            'MsgBox("Registro, Grabado Correctamente")
            Me.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorModificarRegistro"))
            MsgBox(ex.ToString)
        End Try
        drMySql1.Close()
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        vTxtNombre = TxtNombre.Text
        respuesta = MsgBox(resManager.GetString("EliminarConcepto") & " " & vTxtNombre & " " & resManager.GetString("EliminarConcepto2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))
        If respuesta = vbYes Then
            ' Eliminar Registro Conceptos
            vtipoSql = "DELETE FROM conceptos"
            vtipoSql += " WHERE conceptos.CodigoCON = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql

            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox(resManager.GetString("EliminarConcepto3"))
            Catch ex As Exception
                MsgBox(resManager.GetString("EliminarConcepto4"))
                MsgBox(ex.ToString)
            End Try

            ' Eliminar Registros Apuntes
            vtipoSql = "DELETE FROM apuntes"
            vtipoSql += " WHERE apuntes.ConceptoAPU = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox(resManager.GetString("EliminarApuntes"))
            Catch ex As Exception
                MsgBox(resManager.GetString("EliminarApuntesError"))
                MsgBox(ex.ToString)
            End Try

            ' Eliminar Registros Apuntes Periódicos
            vtipoSql = "DELETE FROM apuper"
            vtipoSql += " WHERE apuper.ConceptoAPP = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox(resManager.GetString("EliminarApuntesPeriodicos"))
            Catch ex As Exception
                MsgBox(resManager.GetString("EliminarApuntesPeriodicosError"))
                MsgBox(ex.ToString)
            End Try

            ' Eliminar Registros Presupuestos
            vtipoSql = "DELETE FROM presupuesto"
            vtipoSql += " WHERE presupuesto.ConceptoPRE = '" & vTxtNombre & "' "
            cmdMySql1cr.CommandText = vtipoSql
            Try
                cmdMySql1cr.ExecuteNonQuery()
                MsgBox(resManager.GetString("EliminarPresupuestos"))
            Catch ex As Exception
                MsgBox(resManager.GetString("EliminarPresupuestosError"))
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