Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class EditarCuentaBancaria

    Public vtipoSql, vTxtNombre, vTxtNumero, vTxtTipo, vTxtNotas As String
    Public filaActual As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarCuentaBancaria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Dim TL(4) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoCuenta, frmCuentasBancarias.rmse.GetString("ToolTipTipoCuenta"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNumero, frmCuentasBancarias.rmse.GetString("ToolTipIBAN"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, frmCuentasBancarias.rmse.GetString("ToolTipNombre"))

        ' 1. LLENAR EL COMBO USANDO LA FUNCIÓN GLOBAL MULTIDIOMA
        ' ******************************************************
        CmbTipoCuenta.DropDownStyle = ComboBoxStyle.DropDownList
        CargarComboTipoCuentaGlobal(CmbTipoCuenta)
        ' 2. RECUPERAR LOS DATOS DE LA FILA SELECCIONADA
        ' ******************************************************
        filaActual = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
        ' El DataGridView muestra el tipo TRADUCIDO. Necesitamos buscar qué elemento 
        ' del ComboBox tiene ese mismo "TextoMostrar" para dejarlo preseleccionado.
        Dim tipoTraducidoEnGrid As String = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(0).Value.ToString
        For Each item As Object In CmbTipoCuenta.Items
            Dim elemento As ElementoCombo = CType(item, ElementoCombo)
            If elemento.TextoMostrar = tipoTraducidoEnGrid Then
                CmbTipoCuenta.SelectedItem = item
                Exit For
            End If
        Next

        TxtNombre.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(1).Value.ToString
        TxtNumero.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(2).Value.ToString
        TxtNota.Text = frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(4).Value.ToString
        ' 3. CONFIGURAR SEGÚN MODO EDITAR O MODO ELIMINAR
        ' ******************************************************
        If vEditar = "SI" Then
            'LblEditando.Text = Por defecto creado en el diseño del formulario, se le asigna el texto de Editando
            CmbTipoCuenta.Enabled = True
            TxtNombre.Enabled = False
            TxtNumero.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
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
            If BtnAceptar.Enabled Then BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vTxtNombre = TxtNombre.Text.Trim()
        vTxtNumero = TxtNumero.Text.Trim()
        vTxtNotas = TxtNota.Text.Trim()

        ' 2. OBTENER EL ID NUMÉRICO REAL DESDE EL GRID DE LA PANTALLA ANTERIOR
        Dim idCuentaModificar As Integer

        Try
            Dim filaActual As Integer = frmCuentasBancarias.DgvCuentas.CurrentRow.Index
            ' Recuperamos el Id numérico de la fila seleccionada
            idCuentaModificar = Convert.ToInt32(frmCuentasBancarias.DgvCuentas.Rows(filaActual).Cells(5).Value)
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorRecuperarID"), resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' EXTRAEMOS EL VALOR INTERNO ORIGINAL PARA LA BASE DE DATOS
        If CmbTipoCuenta.SelectedItem IsNot Nothing Then
            Dim itemSeleccionado As ElementoCombo = CType(CmbTipoCuenta.SelectedItem, ElementoCombo)
            vTxtTipo = itemSeleccionado.ValorInterno
        Else
            vTxtTipo = ""
        End If

        Dim vIdTipoCUE As Integer = 0
        If Not String.IsNullOrEmpty(vTxtTipo) Then

            ' 1. Normalizamos la cadena directamente en VB.NET (quitamos todos los espacios)
            Dim vTxtTipoLimpio As String = vTxtTipo.Replace(" ", "")

            ' 2. Simplificamos la consulta SQL eliminando la función 'Replace' que hace fallar a Access
            Dim cmdBuscarId As New OleDb.OleDbCommand("SELECT IdTipoCUE FROM tipocuentas WHERE CodigoTIP = ? OR CodigoTIP = ?", conexion1)

            ' 3. Pasamos los parámetros en el orden estricto en que aparecen en el SQL
            cmdBuscarId.Parameters.AddWithValue("?", vTxtTipo)
            cmdBuscarId.Parameters.AddWithValue("?", vTxtTipoLimpio)

            Try
                Dim resultado As Object = cmdBuscarId.ExecuteScalar() ' ExecuteScalar es ideal porque solo lee un número
                If resultado IsNot Nothing AndAlso Not IsDBNull(resultado) Then
                    ' 2. ¡EL TRUCO RAPIDO!: Buscamos el ID numérico directo en la base de datos
                    vIdTipoCUE = Convert.ToInt32(resultado)
                End If
            Catch ex As Exception
                MessageBox.Show($"{resManager.GetString("ErrorRecuperarID")}: {ex.Message}", resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        ' Modificar Registro
        '*******************
        ' 1. Limpias la consulta cambiando todas las comillas por "?"
        vtipoSql = "UPDATE cuentas SET NumeroCUE = ?, TipoCUE = ?, NotasCUE = ? WHERE IdCuentaCUE = ?"
        cmdMdb1cr.CommandText = vtipoSql

        ' 2. Limpias los parámetros para evitar acumulaciones
        cmdMdb1cr.Parameters.Clear()

        ' 3. Añades los 4 valores en el orden exacto de aparición
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNumero)
        cmdMdb1cr.Parameters.AddWithValue("?", vIdTipoCUE)
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNotas)
        cmdMdb1cr.Parameters.AddWithValue("?", idCuentaModificar) ' El WHERE va al final
        cmdMdb1cr.CommandText = vtipoSql

        Try
            Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()

            If filasAfectadas > 0 Then
                CargarCuentasBancarias()

                ' =====================================================================
                ' 🔍 RADAR POST-GUARDADO: AUDITORÍA EN TIEMPO REAL
                ' =====================================================================
                Try
                    ' Pasamos la consulta matemática para ver si siguen quedando IDs duplicados
                    cmdMdb1cr.CommandText =
                "SELECT COUNT(*) FROM (" &
                "  SELECT IdCuentaCUE FROM cuentas " &
                "  GROUP BY IdCuentaCUE " &
                "  HAVING COUNT(*) > 1" &
                ")"

                    Dim idsDuplicadosEncontrados As Integer = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())

                    If idsDuplicadosEncontrados = 0 Then
						' 🎉 ¡ÉXITO! Ya no hay duplicados: apagamos el modo emergencia de la sesión

						' Ocultamos la columna 5 (índice 4) automáticamente para dejar la tabla limpia
						If frmCuentasBancarias.DgvCuentas.Columns.Count >= 5 Then
							frmCuentasBancarias.DgvCuentas.Columns(5).Visible = False
						End If

						If vModoRepararDuplicados = "SI" Then
                            ' ⚠️ Si veníamos de un modo de reparación, avisamos al usuario que ya no hay duplicados
                            MessageBox.Show(resManager.GetString("EstructuraCuentasCorregida"),
                                resManager.GetString("AppDisplayName"),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                        End If

                        vModoRepararDuplicados = "NO"
                    Else
                        ' ⚠️ Si el usuario guardó pero SIGUE habiendo duplicados, nos aseguramos 
                        ' de mantener la variable en "SI" y la columna 5 bien abierta para que siga revisando
                        vModoRepararDuplicados = "SI"
                        If frmCuentasBancarias.DgvCuentas.Columns.Count >= 5 Then
                            frmCuentasBancarias.DgvCuentas.Columns(5).Visible = True
                        End If
                    End If

                Catch ex As Exception
                    ' Cortafuegos silencioso
                End Try

                Me.Close() ' Guardado con éxito, cierra la ventana modal
            Else
                MessageBox.Show(resManager.GetString("NoEncuentraRegistro"), resManager.GetString("Atencion"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorModificarRegistro") & vbNewLine & ex.Message,
                resManager.GetString("Error"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Aseguramos preventivamente que haya una fila seleccionada en el Grid
        If frmCuentasBancarias.DgvCuentas.CurrentRow Is Nothing Then Exit Sub

        ' Capturamos el ID de la cuenta que viaja seguro en la celda 5
        Dim vIdCuenta As Integer = Convert.ToInt32(frmCuentasBancarias.DgvCuentas.CurrentRow.Cells(5).Value)
        vTxtNombre = TxtNombre.Text

        Try
            ' =====================================================================
            ' 🛡️ ESCUDO QUIRÚRGICO CONTRA ID DUPLICADO
            ' =====================================================================
            ' Pasamos el radar para comprobar si este ID en concreto está duplicado en la base de datos
            cmdMdb1cr.CommandText = "SELECT COUNT(*) FROM cuentas WHERE IdCuentaCUE = " & vIdCuenta
            Dim copiasExistentes As Integer = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())

            Dim esDuplicado As Boolean = (copiasExistentes > 1)

            If esDuplicado Then
                ' Buscamos el nombre de la OTRA cuenta que comparte el mismo ID para alertar al usuario
                ' (Usamos parámetros para blindar la consulta de caracteres raros)
                cmdMdb1cr.CommandText = "SELECT NombreCUE FROM cuentas WHERE IdCuentaCUE = ? AND NombreCUE <> ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@id", OleDbType.Integer).Value = vIdCuenta
                cmdMdb1cr.Parameters.Add("@nom", OleDbType.VarWChar).Value = vTxtNombre.Trim()
                Dim otroNombre As String = Convert.ToString(cmdMdb1cr.ExecuteScalar())

                ' Lanzamos la advertencia explícita de duplicados
                Dim mensajePregunta As String =
                "¡Atención! El ID (" & vIdCuenta & ") está duplicado entre estas dos cuentas:" & vbCrLf &
                "1. - " & vTxtNombre.ToUpper() & " -" & vbCrLf &
                "2. - " & otroNombre.ToUpper() & " -" & vbCrLf & vbCrLf &
                "¿Está seguro de que desea eliminar ÚNICAMENTE la cuenta seleccionada: - " & vTxtNombre.ToUpper() & " - y conservar la otra con sus apuntes?"

                Dim respuestaDuplicado As MsgBoxResult = MsgBox(mensajePregunta, MsgBoxStyle.YesNo + MsgBoxStyle.Information, rmse.GetString("LblEliminando"))

                ' Si el usuario se arrepiente, abortamos la operación de forma segura
                If respuestaDuplicado = vbNo Then Exit Sub
            Else
                ' CASO NORMAL: Si el ID no está duplicado, hacemos tu pregunta de confirmación estándar de fábrica
                Dim respuestaNormal As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("EliminarCuenta") & " " & vTxtNombre & " " & rmse.GetString("EliminarCuenta2"), rmse.GetString("LblEliminando"))
                If respuestaNormal = vbNo Then Exit Sub
            End If

            ' =====================================================================
            ' 💾 PROCESO DE BORRADO SEGURO
            ' =====================================================================
            Dim filasAfectadas As Integer = 0

            ' --- 1. ELIMINAR REGISTROS EN APUNTES (Solo si NO está duplicado) ---
            ' 💡 Si está duplicado, NO borramos de la tabla apuntes porque destruiríamos los apuntes de la cuenta hermana.
            If Not esDuplicado Then
                vtipoSql = "DELETE FROM apuntes WHERE apuntes.CuentaAPU = ?"
                cmdMdb1cr.CommandText = vtipoSql
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@idCuenta", OleDbType.Integer).Value = vIdCuenta

                Try
                    filasAfectadas = cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    Dim msgError As String = "Error al eliminar los apuntes de la cuenta."
                    Try : msgError = frmApuntesContables.rmse.GetString("EliminarApuntesError") : Catch : End Try
                    MsgBox(msgError & vbNewLine & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If

            ' --- 2. ELIMINAR REGISTROS EN APUNTES PERIÓDICOS (Solo si NO está duplicado) ---
            If Not esDuplicado Then
                vtipoSql = "DELETE FROM apuper WHERE apuper.CuentaAPP = ?"
                cmdMdb1cr.CommandText = vtipoSql
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@idApuper", OleDbType.Integer).Value = vIdCuenta

                Try
                    filasAfectadas = cmdMdb1cr.ExecuteNonQuery()
                    If filasAfectadas > 0 Then
                        MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
                    End If
                Catch ex As Exception
                    MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicosError") & vbNewLine & ex.Message)
                End Try
            End If

            ' --- 3. ELIMINAR REGISTRO MAESTRO EN CUENTAS (Quirúrgico si hay duplicado) ---
            If esDuplicado Then
                ' 🌟 CLAVE MAESTRA: Si el ID está repetido, borramos filtrando estrictamente por ID Y NOMBRE a la vez.
                ' De esta forma la cuenta hermana sobrevive ilesa en la tabla.
                vtipoSql = "DELETE FROM cuentas WHERE cuentas.IdCuentaCUE = ? AND cuentas.NombreCUE = ?"
                cmdMdb1cr.CommandText = vtipoSql
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@idCuenta", OleDbType.Integer).Value = vIdCuenta
                cmdMdb1cr.Parameters.Add("@nombreCuenta", OleDbType.VarWChar).Value = vTxtNombre.Trim()
            Else
                ' Si es una cuenta normal sin duplicar, borramos de forma estándar por su ID
                vtipoSql = "DELETE FROM cuentas WHERE cuentas.IdCuentaCUE = ?"
                cmdMdb1cr.CommandText = vtipoSql
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.Add("@idCuenta", OleDbType.Integer).Value = vIdCuenta
            End If

            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("EliminarCuenta3"))
            Catch ex As Exception
                MsgBox(rmse.GetString("EliminarCuenta4") & vbNewLine & ex.Message)
                Exit Sub
            End Try

            ' Refrescamos la pantalla con tus funciones
            CargarCuentasBancarias()

            ' =====================================================================
            ' 🔍 RADAR POST-GUARDADO: AUDITORÍA EN TIEMPO REAL
            ' =====================================================================
            Try
                ' Pasamos la consulta matemática para ver si siguen quedando IDs duplicados
                cmdMdb1cr.CommandText =
                "SELECT COUNT(*) FROM (" &
                "  SELECT IdCuentaCUE FROM cuentas " &
                "  GROUP BY IdCuentaCUE " &
                "  HAVING COUNT(*) > 1" &
                ")"

                Dim idsDuplicadosEncontrados As Integer = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())

                If idsDuplicadosEncontrados = 0 Then
                    ' 🎉 ¡ÉXITO! Ya no hay duplicados: apagamos el modo emergencia de la sesión

                    ' Ocultamos la columna 5 (índice 4) automáticamente para dejar la tabla limpia
                    If frmCuentasBancarias.DgvCuentas.Columns.Count >= 5 Then
                        frmCuentasBancarias.DgvCuentas.Columns(5).Visible = False
                    End If

                    If vModoRepararDuplicados = "SI" Then
                        ' ⚠️ Si veníamos de un modo de reparación, avisamos al usuario que ya no hay duplicados
                        MessageBox.Show(resManager.GetString("EstructuraCuentasCorregida"),
                                resManager.GetString("AppDisplayName"),
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information)
                    End If

                    vModoRepararDuplicados = "NO"
                Else
                    ' ⚠️ Si el usuario guardó pero SIGUE habiendo duplicados, nos aseguramos 
                    ' de mantener la variable en "SI" y la columna 5 bien abierta para que siga revisando
                    vModoRepararDuplicados = "SI"
                    If frmCuentasBancarias.DgvCuentas.Columns.Count >= 5 Then
                        frmCuentasBancarias.DgvCuentas.Columns(5).Visible = True
                    End If
                End If

            Catch ex As Exception
                ' Cortafuegos silencioso
            End Try
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorGeneral") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = 3 Then
            e.Cancel = False
        End If
    End Sub
End Class