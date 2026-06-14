Imports System.Windows.Forms

Public Class EditarConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtNotas As String
    Public filaActual As Integer
    Public TL(2) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
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
        vTxtDescripcion = TxtDescripcion.Text.Trim() ' Ya no necesitas cambiar apóstrofes si usas parámetros
        vTxtNotas = TxtNota.Text.Trim()

        ' --- TRADUCCIÓN INVERSA OPTIMIZADA ---
        Dim codigoOriginalMDB As String = vTxtNombre
        Dim textoBuscar As String = vTxtNombre.ToUpper()

        ' Buscamos la llave original en el archivo de recursos de la interfaz actual
        Dim resSet As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)

        If resSet IsNot Nothing Then
            For Each dict As System.Collections.DictionaryEntry In resSet
                ' Evitamos nulos con el operador '?' y normalizamos a mayúsculas
                Dim valorTraducido As String = dict.Value?.ToString().Trim().ToUpper()

                If valorTraducido = textoBuscar Then
                    ' Encontramos la llave original de la base de datos (Ej: "ALQUILER")
                    codigoOriginalMDB = dict.Key.ToString()
                    Exit For
                End If
            Next
        End If

        ' --- CONTROL DE CONCEPTOS PROPIOS (Ej: en catalán o personalizados) ---
        ' Si tras buscar en el .resx encontramos una coincidencia, verificamos si es una de tus 
        ' llaves de sistema. Si no coincide con ninguna traducción real del sistema, 
        ' asumimos que es un concepto propio del usuario y mantenemos su texto original.
        If codigoOriginalMDB <> vTxtNombre Then
            ' Verificación secundaria: ¿La llave encontrada realmente existe en tu gestor de recursos?
            If resManager.GetString(codigoOriginalMDB) Is Nothing AndAlso resManager.GetString("Desc_" & codigoOriginalMDB) Is Nothing Then
                ' Si no es una llave del sistema, es un texto propio del usuario
                codigoOriginalMDB = vTxtNombre
            End If
        End If

        ' Modificar Registro usando parámetros seguros
        ' *******************************************************
        vtipoSql = "UPDATE conceptos SET DescripcionCON = ?, NotasCON = ? " &
           "WHERE conceptos.CodigoCON = ?"

        cmdMdb1cr.CommandText = vtipoSql

        ' ¡Recuerda! En Access/OleDb el orden de los parámetros debe ser EXACTO al del SQL
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionCON", vTxtDescripcion) ' Soporta apóstrofes nativos (')
        cmdMdb1cr.Parameters.AddWithValue("@NotasCON", vTxtNotas)             ' Soporta apóstrofes nativos (')
        cmdMdb1cr.Parameters.AddWithValue("@CodigoCON", codigoOriginalMDB)   ' Filtro WHERE usando la clave recuperada

        Try
            cmdMdb1cr.ExecuteNonQuery()
            Me.Close()
        Catch ex As Exception
            ' Mensaje de error traducido desde tu gestor de recursos
            MessageBox.Show(resManager.GetString("ErrorModificarRegistro"),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Obtener el nombre del concepto visible que se pretende eliminar
        vTxtNombre = TxtNombre.Text.Trim()

        ' 2. REVERTIR EL IDIOMA (Búsqueda inversa usando tu variable de cultura original)
        Dim nombreOriginalBD As String = vTxtNombre

        ' Si la cultura no empieza por español, ejecutamos la traducción inversa precisa
        ' --- TRADUCCIÓN INVERSA LIMPIA Y UNIFICADA ---
        Dim textoBuscar As String = vTxtNombre.ToUpper()

        ' .NET lee automáticamente el idioma visual que se activó en tus Preferencias
        Dim recursos As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)

        If recursos IsNot Nothing Then
            For Each elemento As System.Collections.DictionaryEntry In recursos
                Dim valorTraducido As String = elemento.Value?.ToString().Trim().ToUpper()

                If valorTraducido = textoBuscar Then
                    Dim llaveEncontrada As String = elemento.Key.ToString()

                    ' Si la llave corresponde a una descripción (Desc_), le quitamos el prefijo limpiamente sin dejar espacios
                    If llaveEncontrada.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then
                        nombreOriginalBD = llaveEncontrada.Substring(5)
                    Else
                        nombreOriginalBD = llaveEncontrada
                    End If
                    Exit For
                End If
            Next
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
        If vTxtNombre <> nombreOriginalBD Then
            nombreOriginalBD = vTxtNombre
        End If

        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("EliminarConcepto") & " " & nombreOriginalBD & " " & rmse.GetString("EliminarConcepto2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))

        If respuesta = vbYes Then
            ' 5. EJECUCIÓN EN CASCADA CORREGIDA: Se asigna Y SE EJECUTA cada tabla por separado
            vtipoSql = "DELETE FROM conceptos WHERE CodigoCON = '" & nombreOriginalBD & "'"
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmEditarConceptoContable.rmse.GetString("EliminarConcepto3"))
            Catch ex As Exception
                MsgBox(ex.Message)
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