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
            ' Creamos un comando rápido para leer solo el ID de ese tipo de cuenta
            ' Usamos .Replace(" ", "") por si acaso el texto viene sin espacios
            Dim cmdBuscarId As New OleDb.OleDbCommand("SELECT IdTipoCUE FROM tipocuentas WHERE CodigoTIP = ? OR Replace(CodigoTIP, ' ', '') = ?", conexion1)
            cmdBuscarId.Parameters.AddWithValue("?", vTxtTipo)
            cmdBuscarId.Parameters.AddWithValue("?", vTxtTipo.Replace(" ", ""))

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
                Me.Close() ' Guardado con éxito, cierra la ventana modal
            Else
                MessageBox.Show(resManager.GetString("NoEncuentraRegistro"), resManager.GetString("Atencion"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorModificarRegistro") & vbNewLine & ex.Message,
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Aseguramos preventivamente que haya una fila seleccionada en el Grid
        If frmCuentasBancarias.DgvCuentas.CurrentRow Is Nothing Then Exit Sub

        ' Capturamos el ID de la cuenta que viaja seguro en la celda 5 (Oculta)
        Dim vIdCuenta As Integer = Convert.ToInt32(frmCuentasBancarias.DgvCuentas.CurrentRow.Cells(5).Value)
        vTxtNombre = TxtNombre.Text

        ' Preguntamos confirmación al usuario (Tu excelente aviso de fábrica)
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("EliminarCuenta") & " " & vTxtNombre & " " & rmse.GetString("EliminarCuenta2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))

        If respuesta = vbYes Then
            Dim filasAfectadas As Integer = 0

            ' --- 1. ELIMINAR REGISTROS EN APUNTES (Clave foránea - Se borra primero por integridad) ---
            ' Filtramos por el ID numérico que guarda la tabla apuntes
            vtipoSql = "DELETE FROM apuntes WHERE apuntes.CuentaAPU = ?"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            ' 🚀 CORRECCIÓN CLAVE: Le damos un nombre alfanumérico al parámetro en la RAM
            cmdMdb1cr.Parameters.Add("@idApu", OleDbType.Integer).Value = vIdCuenta

            Try
                filasAfectadas = cmdMdb1cr.ExecuteNonQuery()
                If filasAfectadas > 0 Then
                    MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntes"))
                End If
            Catch ex As Exception
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntesError") & vbNewLine & ex.Message)
            End Try

            ' --- 2. ELIMINAR REGISTROS EN APUNTES PERIÓDICOS ---
            ' Filtramos por el ID numérico que guarda la tabla apuper
            vtipoSql = "DELETE FROM apuper WHERE apuper.CuentaAPP = ?"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            ' 🚀 CORRECCIÓN CLAVE: Le damos un nombre alfanumérico al parámetro en la RAM
            cmdMdb1cr.Parameters.Add("@idApuper", OleDbType.Integer).Value = vIdCuenta

            Try
                filasAfectadas = cmdMdb1cr.ExecuteNonQuery()
                If filasAfectadas > 0 Then
                    MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
                End If
            Catch ex As Exception
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicosError") & vbNewLine & ex.Message)
            End Try

            ' --- 3. ELIMINAR REGISTRO MAESTRO EN CUENTAS (Se borra al final) ---
            ' Borramos por ID para evitar problemas si el usuario cambió el texto
            vtipoSql = "DELETE FROM cuentas WHERE cuentas.IdCuentaCUE = ?"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            ' 🚀 CORRECCIÓN CLAVE: Le damos un nombre alfanumérico al parámetro en la RAM
            cmdMdb1cr.Parameters.Add("@idCuenta", OleDbType.Integer).Value = vIdCuenta

            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("EliminarCuenta3"))
            Catch ex As Exception
                MsgBox(rmse.GetString("EliminarCuenta4") & vbNewLine & ex.Message)
                Exit Sub
            End Try

            CargarCuentasBancarias()
            ' Cerramos la ventana de edición/borrado
            Me.Close()
        End If
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