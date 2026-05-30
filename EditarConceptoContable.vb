Imports System.Windows.Forms

Public Class EditarConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtNotas As String
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

        ' 1. Capturar datos actuales del Grid
        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index

        Dim tipoTextoGrid As String = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(0).Value.ToString().Trim().ToUpper()
        TxtNombre.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString()
        TxtDescripcion.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(2).Value.ToString()
        TxtNota.Text = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(3).Value.ToString()

        ' ======================================================================
        ' 2. SINCRONIZAR EL TEXTBOX Y DETERMINAR EL TIPO ORIGINAL (BD)
        ' ======================================================================
        Dim tipoOriginalBD As String
        Dim textoEspecialRecurso As String = resManager.GetString("Tipo_Especial").Trim().ToUpper()

        If tipoTextoGrid = textoEspecialRecurso Then
            TxtTipoConcepto.Text = tipoTextoGrid ' Muestra "SPECIAL", "SPÉCIAL", etc.
            tipoOriginalBD = "ESPECIAL"
        Else
            ' Respaldo de seguridad general
            TxtTipoConcepto.Text = tipoTextoGrid
            tipoOriginalBD = tipoTextoGrid
        End If

        ' ======================================================================
        ' 3. EVALUAR MODO (Editar o Eliminar) CON EXCEPCIÓN PARA "ESPECIAL"
        ' ======================================================================
        If vEditar = "SI" Then
            TxtNombre.Enabled = False
            ' Si es del sistema (ESPECIAL), bloqueamos su edición de raíz
            If tipoOriginalBD = "ESPECIAL" Then
                TxtDescripcion.Enabled = False
                TxtNota.Enabled = False
                BtnAceptar.Enabled = False
                BtnEliminar.Enabled = False

                ' Mostramos un aviso visual opcional en el formulario
                LblEditando.Text = rmse.GetString("Concepto_No_Editable")
                BtnCancelar.Select()
            Else
                ' Flujo normal para conceptos modificables del usuario
                TxtDescripcion.Select()
                BtnEliminar.Enabled = False
            End If
        Else
            ' MODO ELIMINAR (Muestra los datos bloqueados listos para pulsar Eliminar)
            LblEditando.Text = rmse.GetString("LblEliminando")
            TxtNombre.Enabled = False
            TxtDescripcion.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False

            ' Si el tipo es ESPECIAL, el botón de eliminar también se apaga por seguridad
            If tipoOriginalBD = "ESPECIAL" Then
                BtnEliminar.Enabled = False
                ' Mostramos un aviso visual opcional en el formulario
                LblEditando.Text = rmse.GetString("Concepto_No_Editable")
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
        vTxtNombre = TxtNombre.Text.Trim()
        vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
        vTxtNotas = TxtNota.Text

        ' --- TRADUCCIÓN INVERSA: Buscar el código original en la MDB ---
        Dim codigoOriginalMDB As String = vTxtNombre ' Por si está en español y no cambia

        ' Si el idioma actual no es español, buscamos la llave original en el ResX
        If My.Settings.CulturaUsuario <> "es" Then
            ' Obtenemos todos los registros guardados en el archivo de recursos
            Dim resSet As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentCulture, True, True)

            If resSet IsNot Nothing Then
                ' Recorremos todas las llaves del archivo de recursos
                For Each dict As System.Collections.DictionaryEntry In resSet
                    ' Si el valor traducido coincide con lo que hay en el TextBox...
                    If dict.Value.ToString().Trim().ToUpper() = vTxtNombre.ToUpper() Then
                        ' ¡Encontramos la llave original! (Ej: "RENT" -> su llave es "ALQUILER")
                        codigoOriginalMDB = dict.Key.ToString()
                        Exit For
                    End If
                Next
            End If
        End If

        ' Modificar Registro usando el código original recuperado
        ' *******************************************************
        vtipoSql = "UPDATE conceptos SET DescripcionCON = '" & vTxtDescripcion & "' , NotasCON = '" & vTxtNotas & "' "
        vtipoSql += " WHERE conceptos.CodigoCON = '" & codigoOriginalMDB & "' "
        cmdMdb1cr.CommandText = vtipoSql

        Try
            cmdMdb1cr.ExecuteNonQuery() ' Recomendado cambiar a ExecuteNonQuery
            Me.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorModificarRegistro"))
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Obtener el nombre del concepto visible que se pretende eliminar
        vTxtNombre = TxtNombre.Text.Trim()

        ' 2. REVERTIR EL IDIOMA (Búsqueda inversa usando tu variable de cultura original)
        Dim nombreOriginalBD As String = vTxtNombre

        ' Si la cultura no empieza por español, ejecutamos la traducción inversa precisa
        If Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
            Dim cultura As New System.Globalization.CultureInfo(My.Settings.CulturaUsuario)
            ' Nota: He usado 'resManager' que es tu recurso global de textos contables
            Dim recursos As System.Resources.ResourceSet = resManager.GetResourceSet(cultura, True, True)

            If recursos IsNot Nothing Then
                For Each elemento As System.Collections.DictionaryEntry In recursos
                    If elemento.Value.ToString().Trim().ToUpper() = vTxtNombre.ToUpper() Then
                        ' Restauramos el código original revirtiendo los guiones bajos por espacios
                        nombreOriginalBD = elemento.Key.ToString().Replace("_", " ")
                        Exit For
                    End If
                Next
            End If
        End If

        ' 3. VALIDACIÓN: Verificar si el concepto tiene el tipo original "ESPECIAL"
        Dim vSqlVerificarEspecial As String = "SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & nombreOriginalBD & "'"
        cmdMdb1cr.CommandText = vSqlVerificarEspecial

        Try
            Dim tipoOrigen As Object = cmdMdb1cr.ExecuteScalar()

            ' Si el resultado coincide con "ESPECIAL", bloqueamos la acción por completo
            If tipoOrigen IsNot Nothing AndAlso tipoOrigen.ToString().Trim().ToUpper() = "ESPECIAL" Then
                MsgBox(rmse.GetString("ConceptoSistemaNoBorrar"), vbExclamation, resManager.GetString("AccionCancelada"))
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorVerificarIntegridad") & ": " & ex.Message, vbCritical, rmse.GetString("$this.Text"))
            Exit Sub
        End Try

        ' 4. Mensaje de confirmación (Muestra vTxtNombre que es el texto traducido que el usuario entiende)
        respuesta = MsgBox(rmse.GetString("EliminarConcepto") & " " & vTxtNombre & " " & rmse.GetString("EliminarConcepto2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))

        If respuesta = vbYes Then

            ' 5. EJECUCIÓN EN CASCADA CORREGIDA: Se asigna Y SE EJECUTA cada tabla por separado

            ' A. Eliminar de la tabla 'conceptos' (¡Aquí se ejecutaba mal en tu código antiguo!)
            vtipoSql = "DELETE FROM conceptos WHERE CodigoCON = '" & nombreOriginalBD & "'"
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Error al eliminar el concepto base: " & ex.Message)
            End Try

            ' B. Eliminar Registros Apuntes
            vtipoSql = "DELETE FROM apuntes WHERE ConceptoAPU = '" & nombreOriginalBD & "'"
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntes"))
            Catch ex As Exception
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntesError") & vbNewLine & ex.Message)
            End Try

            ' C. Eliminar Registros Apuntes Periódicos
            vtipoSql = "DELETE FROM apuper"
            cmdMdb1cr.CommandText = vtipoSql
            ' (Mantenemos tu lógica original para mapear la query de apuper)
            vtipoSql = "DELETE FROM apuper WHERE apuper.ConceptoAPP = '" & nombreOriginalBD & "'"
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
            Catch ex As Exception
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicosError") & vbNewLine & ex.Message)
            End Try

            ' D. Eliminar Registros Presupuestos
            vtipoSql = "DELETE FROM presupuesto WHERE presupuesto.ConceptoPRE = '" & nombreOriginalBD & "'"
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmPresupuestos.rmse.GetString("EliminarPresupuestos"))
            Catch ex As Exception
                MsgBox(frmPresupuestos.rmse.GetString("EliminarPresupuestosError") & vbNewLine & ex.Message)
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