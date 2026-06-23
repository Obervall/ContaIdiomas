Imports System.Windows.Forms

Public Class EditarConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtNotas As String
    Public filaActual As Integer
    Public TL(2) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

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
        ' 1. Capturamos los textos de los cuadros del formulario
        vTxtNombre = TxtNombre.Text.Trim()
        vTxtDescripcion = TxtDescripcion.Text.Trim()
        vTxtNotas = TxtNota.Text.Trim()

        ' 2. OBTENER EL ID NUMÉRICO REAL DESDE EL GRID DE LA PANTALLA ANTERIOR
        Dim idConceptoModificar As Integer

        Try
            Dim filaActual As Integer = frmConceptosContables.DgvConceptos.CurrentRow.Index
            ' Recuperamos el Id numérico de la fila seleccionada
            idConceptoModificar = Convert.ToInt32(frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(4).Value)
        Catch ex As Exception
            MessageBox.Show("Error al recuperar el identificador del registro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' 3. CONFIGURAR EL UPDATE USANDO EL ID NUMÉRICO (Inmune a problemas de guiones)
        ' Si además quieres permitir que se edite el NOMBRE del concepto y se guarde con guion, usa la línea de abajo comentada.
        ' De momento, modificamos solo Descripción y Notas:
        vtipoSql = "UPDATE conceptos SET DescripcionCON = ?, NotasCON = ? WHERE IdConceptoCON = ?"
        cmdMdb1cr.CommandText = vtipoSql

        ' En Access/OleDb el orden de los parámetros debe ser STABLE y EXACTO al del SQL
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionCON", vTxtDescripcion)
        cmdMdb1cr.Parameters.AddWithValue("@NotasCON", vTxtNotas)
        cmdMdb1cr.Parameters.AddWithValue("@IdConceptoCON", idConceptoModificar) ' Filtro WHERE

        Try
            Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()

            If filasAfectadas > 0 Then
                Me.Close() ' Guardado con éxito, cierra la ventana modal
            Else
                MessageBox.Show("No se encontró el registro para actualizar.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorModificarRegistro") & vbNewLine & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Capturamos el texto visual solo para los mensajes de la pantalla
        vTxtNombre = TxtNombre.Text.Trim()

        ' =========================================================================
        ' 2. OBTENER EL ID NUMÉRICO DIRECTO DESDE EL GRID DE LA PANTALLA ANTERIOR
        ' =========================================================================
        Dim idConcepto As Integer = 0
        Try
            Dim filaActual As Integer = frmConceptosContables.DgvConceptos.CurrentRow.Index
            ' Extraemos el ID numérico real guardado en la celda oculta (4)
            idConcepto = Convert.ToInt32(frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(4).Value)
        Catch ex As Exception
            MessageBox.Show("Error al recuperar el identificador para la eliminación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' =========================================================================
        ' 3. VALIDACIÓN DE PROTECCIÓN: CONCEPTOS "ESPECIAL"
        ' =========================================================================
        ' Consultamos el TipoCON directamente a través de su ID numérico estable
        Dim vSqlVerificarEspecial As String = "SELECT TipoCON FROM conceptos WHERE IdConceptoCON = ?"
        cmdMdb1cr.CommandText = vSqlVerificarEspecial
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("?", idConcepto)

        Dim tipoOrigen As String = ""
        Try
            Dim resultado As Object = cmdMdb1cr.ExecuteScalar()
            If resultado IsNot Nothing Then
                tipoOrigen = resultado.ToString().Trim().ToUpper()
            End If

            ' Bloqueo absoluto si es de fábrica/sistema
            If tipoOrigen = "ESPECIAL" Then
                MessageBox.Show(rmse.GetString("ConceptoSistemaNoBorrar"),
                resManager.GetString("AccionCancelada"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Exclamation)
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(rmse.GetString("ConceptoSistemaNoBorrar"),
                resManager.GetString("$this.Text"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Exit Sub
        End Try

        ' =========================================================================
        ' 4. MENSAJE DE CONFIRMACIÓN ÚNICO Y ADVERTENCIA EN CASCADA
        ' =========================================================================
        ' Construimos un mensaje claro que avise al usuario del impacto total en la base de datos
        Dim mensajeAlerta As String = rmse.GetString("EliminarConcepto") & " [" & vTxtNombre & "] " & rmse.GetString("EliminarConcepto2") & vbNewLine & vbNewLine &
                                  "⚠️ ADVERTENCIA: Esta acción eliminará de forma irreversible todos los apuntes históricos, " &
                                  "apuntes periódicos y presupuestos asociados a este concepto."

        Dim respuesta As MsgBoxResult = MsgBox(mensajeAlerta, vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))

        ' =========================================================================
        ' 5. EJECUCIÓN DEL BORRADO INTEGRAL POR ID NUMÉRICO
        ' =========================================================================
        If respuesta = vbYes Then
            Try
                ' A. Tabla: conceptos (¡CORREGIDO! Ahora filtra por ID numérico)
                cmdMdb1cr.CommandText = "DELETE FROM conceptos WHERE IdConceptoCON = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", idConcepto)
                cmdMdb1cr.ExecuteNonQuery()

                ' B. Tabla: apuntes
                cmdMdb1cr.CommandText = "DELETE FROM apuntes WHERE ConceptoAPU = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", idConcepto)
                cmdMdb1cr.ExecuteNonQuery()

                ' C. Tabla: apuper
                cmdMdb1cr.CommandText = "DELETE FROM apuper WHERE ConceptoAPP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", idConcepto)
                cmdMdb1cr.ExecuteNonQuery()

                ' D. Tabla: presupuesto
                cmdMdb1cr.CommandText = "DELETE FROM presupuesto WHERE ConceptoPRE = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", idConcepto)
                cmdMdb1cr.ExecuteNonQuery()

                ' Un único mensaje final de éxito limpio, sin interrumpir con cuatro ventanas seguidas
                MessageBox.Show(resManager.GetString("ConceptoYRegistrosEliminados"),
                resManager.GetString("$this.Text"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
                Me.Close() ' Cierra el formulario modal con éxito
            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorConceptoYRegistrosEliminados") & ": " & vbNewLine & ex.Message,
                resManager.GetString("Error"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            End Try
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

End Class