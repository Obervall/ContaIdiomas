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
        TL(2).SetToolTip(Me.TxtDescripcion, rmse.GetString("ToolTipDescripcion"))

        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoConcepto.SelectedIndex = 0

        ' 1. Capturar datos actuales del Grid
        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index

        Dim tipoTextoGrid As String = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(0).Value.ToString().Trim()
        TxtNombre.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString()
        TxtDescripcion.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(2).Value.ToString()
        TxtNota.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(3).Value.ToString()

        ' 2. REVERTIR IDIOMA: Averiguar el TipoCON original de la Base de Datos
        Dim tipoOriginalBD As String = tipoTextoGrid
        Dim recursos As System.Resources.ResourceSet = rmse.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)
        If recursos IsNot Nothing Then
            For Each elemento As System.Collections.DictionaryEntry In recursos
                If elemento.Value.ToString().Trim().ToUpper() = tipoTextoGrid.ToUpper() Then
                    ' Extraemos la Key original (ej: "Tipo_Especial" o "Tipo_Gasto")
                    Dim keyEncontrada As String = elemento.Key.ToString().ToUpper()

                    ' Identificamos si mapea con tu regla de negocio de conceptos
                    If keyEncontrada = "TIPO_GASTO" Then tipoOriginalBD = "GASTO"
                    If keyEncontrada = "TIPO_INGRESO" Then tipoOriginalBD = "INGRESO"
                    If keyEncontrada = "TIPO_ESPECIAL" Then tipoOriginalBD = "ESPECIAL"
                    Exit For
                End If
            Next
        End If

        ' Asignamos el texto traducido al ComboBox para mantener la consistencia visual
        CmbTipoConcepto.Text = tipoTextoGrid

        ' 3. EVALUAR MODO (Editar o Eliminar) CON EXCEPCIÓN PARA "ESPECIAL"
        If vEditar = "SI" Then
            ' Si es del sistema (ESPECIAL), bloqueamos su edición de raíz
            If tipoOriginalBD = "ESPECIAL" Then
                CmbTipoConcepto.Enabled = False
                TxtNombre.Enabled = False
                TxtDescripcion.Enabled = False
                TxtNota.Enabled = False
                BtnAceptar.Enabled = False
                BtnEliminar.Enabled = False

                ' Mostramos un aviso visual opcional en el formulario (debes añadir esta Key en ResX Manager)
                LblEditando.Text = rmse.GetString("ConceptoSistemaProtegido")
                BtnCancelar.Select()
            Else
                ' Flujo normal para conceptos modificables del usuario
                CmbTipoConcepto.Enabled = False
                TxtNombre.Enabled = False
                TxtDescripcion.Select()
                BtnEliminar.Enabled = False
            End If
        Else
            ' MODO ELIMINAR (Muestra los datos bloqueados listos para pulsar Eliminar)
            LblEditando.Text = rmse.GetString("LblEliminando")
            CmbTipoConcepto.Enabled = False
            TxtNombre.Enabled = False
            TxtDescripcion.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False

            ' Si el tipo es ESPECIAL, el botón de eliminar también se apaga por seguridad
            If tipoOriginalBD = "ESPECIAL" Then
                BtnEliminar.Enabled = False
                BtnCancelar.Select()
            Else
                BtnEliminar.Select()
            End If

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
        cmdMdb1cr.CommandText = vtipoSql

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            'MsgBox("Registro, Grabado Correctamente")
            Me.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorModificarRegistro"))
            MsgBox(ex.ToString)
        End Try
        drMdb1.Close()
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Obtener el nombre del concepto que se pretende eliminar
        vTxtNombre = TxtNombre.Text.Trim()

        ' 2. REVERTIR EL IDIOMA (Búsqueda inversa para obtener el nombre original de la BD)
        Dim nombreOriginalBD As String = vTxtNombre
        Dim recursos As System.Resources.ResourceSet = rmse.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)
        If recursos IsNot Nothing Then
            For Each elemento As System.Collections.DictionaryEntry In recursos
                If elemento.Value.ToString().Trim().ToUpper() = vTxtNombre.ToUpper() Then
                    nombreOriginalBD = elemento.Key.ToString().Replace("_", " ")
                    Exit For
                End If
            Next
        End If

        ' 3. VALIDACIÓN: Verificar si el concepto tiene el tipo original "ESPECIAL"
        ' Buscamos directamente con el nombre original de la base de datos
        Dim vSqlVerificarEspecial As String = "SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & nombreOriginalBD & "'"
        cmdMdb1cr.CommandText = vSqlVerificarEspecial

        Try
            Dim tipoOrigen As Object = cmdMdb1cr.ExecuteScalar()

            ' Si el resultado no es nulo y es igual a "ESPECIAL", bloqueamos la acción
            If tipoOrigen IsNot Nothing AndAlso tipoOrigen.ToString().Trim().ToUpper() = "ESPECIAL" Then
                ' Recuerda dar de alta "ConceptoSistemaNoBorrar" en tu ResX Manager
                MsgBox(rmse.GetString("ConceptoSistemaNoBorrar"), vbExclamation, resManager.GetString("AccionCancelada"))
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorVerificarIntegridad") & ": " & ex.Message, vbCritical, rmse.GetString("$this.Text"))
            Exit Sub
        End Try

        ' 4. Mensaje de confirmación (Muestra el nombre traducido para el usuario)
        respuesta = MsgBox(rmse.GetString("EliminarConcepto") & " " & vTxtNombre & " " & rmse.GetString("EliminarConcepto2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))

        If respuesta = vbYes Then
            ' 5. Eliminar Registro Conceptos utilizando el nombre real de la BD
            vtipoSql = "DELETE FROM conceptos"
            vtipoSql += " WHERE conceptos.CodigoCON = '" & nombreOriginalBD & "' "
            cmdMdb1cr.CommandText = vtipoSql

            ' Eliminar Registros Apuntes
            vtipoSql = "DELETE FROM apuntes"
            vtipoSql += " WHERE apuntes.ConceptoAPU = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntes"))
            Catch ex As Exception
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntesError") & vbNewLine & ex.Message)
            End Try

            ' Eliminar Registros Apuntes Periódicos
            vtipoSql = "DELETE FROM apuper"
            vtipoSql += " WHERE apuper.ConceptoAPP = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
            Catch ex As Exception
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicosError") & vbNewLine & ex.Message)
            End Try

            ' Eliminar Registros Presupuestos
            vtipoSql = "DELETE FROM presupuesto"
            vtipoSql += " WHERE presupuesto.ConceptoPRE = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmPresupuestos.rmse.GetString("EliminarPresupuestos"))
            Catch ex As Exception
                MsgBox(frmPresupuestos.rmse.GetString("EliminarPresupuestosError"))
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