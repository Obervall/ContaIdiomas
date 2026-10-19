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
            TxtNombre.Enabled = True
            TxtNombre.Select()
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
            TxtNombre.Enabled = False
            TxtDescripcion.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtDescripcion.Select()
        End If
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Capturamos los textos limpios de los cuadros del formulario
        vTxtNombre = TxtNombre.Text.Trim().ToUpper()
        vTxtDescripcion = TxtDescripcion.Text.Trim()

        ' Cortafuegos básico
        If String.IsNullOrWhiteSpace(vTxtNombre) Then
            MsgBox(resManager.GetString("NombreTipoCuentaVacio"), MsgBoxStyle.Exclamation, resManager.GetString("Validacion"))
            Exit Sub
        End If

        ' =========================================================================
        ' 🛡️ PASO 2: LOGÍSTICA INVERSA MULTI-IDIOMA (Averiguar la clave base del .resx)
        ' =========================================================================
        ' Convertimos el texto de la pantalla a su equivalente original en Español de fábrica
        Dim nombreEnEspañolBase As String = vTxtNombre

        Dim resSet As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)
        If resSet IsNot Nothing Then
            For Each dict As System.Collections.DictionaryEntry In resSet
                Dim llaveKey As String = dict.Key.ToString()
                If llaveKey.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim valorTraducido As String = dict.Value?.ToString().Trim().ToUpper()

                ' Si lo que ha tecleado el usuario coincide con una traducción activa del .resx (ej: "COMPTE CORRENT")
                If valorTraducido = vTxtNombre Then
                    ' Recuperamos rígidamente la clave original de fábrica (ej: "CUENTA_CORRIENTE")
                    nombreEnEspañolBase = llaveKey.ToUpper()
                    Exit For
                End If
            Next
        End If

        ' Limpiamos espacios y guiones para comparar de forma estricta contra tu lista de la propiedad (ej: "CUENTACORRIENTE")
        Dim textoACompararSistema As String = nombreEnEspañolBase.Replace("_", "").Replace(" ", "").Trim().ToUpper()

        ' =========================================================================
        ' ❌ CORTAFUEGOS 1: PROHIBIDO USAR NOMBRES PROTEGIDOS DE SISTEMA
        ' =========================================================================
        ' Si el usuario intenta renombrar una cuenta libre y llamarla igual que un tipo de fábrica 
        ' (ya sea escribiéndolo en su idioma o en el código base), lo detenemos en seco.
        For Each tipoProtegido As String In TiposCuentaMuestraSistema
            Dim tipoProtegidoLimpio As String = tipoProtegido.Replace("_", "").Replace(" ", "").Trim().ToUpper()

            If textoACompararSistema = tipoProtegidoLimpio Then
                Dim msgBloqueoSystem As String = resManager.GetString("NoSePuedeUsarTipoCuentaPredeterminada")
                MsgBox(msgBloqueoSystem, MsgBoxStyle.Critical, resManager.GetString("AccionDenegada"))
                TxtNombre.Focus()
				TxtNombre.SelectAll()
				Exit Sub ' Frena el guardado inmediatamente
            End If
        Next

        ' =========================================================================
        ' 2. OBTENER EL ID NUMÉRICO REAL (Usamos tu variable global de control)
        ' =========================================================================
        ' Recuperamos el Id numérico de forma rígida a través de la variable global del taller 
        ' que seteamos al abrir esta pantalla (ej: vIdTipoModificar). 
        If idTipoModificar <= 0 Then
            MessageBox.Show(resManager.GetString("ErrorRecuperarID"), resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        ' =========================================================================
        ' 🔍 CORTAFUEGOS 2: RADAR ANTI-DUPLICADOS EN LA BASE DE DATOS
        ' =========================================================================

        ' Comprobamos tanto el nombre crudo introducido como su traducción inversa por seguridad
        cmdMdb1cr.CommandText = "SELECT COUNT(*) FROM tipocuentas WHERE (UCase(CodigoTIP) = ? Or UCase(CodigoTIP) = ?) And IdTipoCUE <> ?"
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@Nom1", vTxtNombre)
        cmdMdb1cr.Parameters.AddWithValue("@Nom2", nombreEnEspañolBase.Replace("_", " "))
        cmdMdb1cr.Parameters.AddWithValue("@Id", idTipoModificar)

        Try
            Dim existeDuplicado As Boolean = (Convert.ToInt32(cmdMdb1cr.ExecuteScalar()) > 0)
            If existeDuplicado Then
                MsgBox(resManager.GetString("YaExisteTipoCuenta"), MsgBoxStyle.Exclamation, resManager.GetString("RegistroDuplicado"))
                TxtNombre.Focus()
				TxtNombre.SelectAll()
				Exit Sub
            End If
        Catch : End Try

        ' =========================================================================
        ' 🚀 PASO 3: CONFIGURAR EL UPDATE SEGURO
        ' =========================================================================
        ' Guardamos siempre la clave en su formato plano limpio para no corromper las relaciones futuras
        vtipoSql = "UPDATE tipocuentas SET CodigoTIP = ?, DescripcionTIP = ? WHERE IdTipoCUE = ?"
        cmdMdb1cr.CommandText = vtipoSql

        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@CodigoTIP", vTxtNombre)
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionTIP", vTxtDescripcion)
        cmdMdb1cr.Parameters.AddWithValue("@IdTipoCUE", idTipoModificar)

        Try
            Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()
            If filasAfectadas > 0 Then
                MessageBox.Show(resManager.GetString("TipoCuentaActualizado"), resManager.GetString("Mantenimiento"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Close()
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