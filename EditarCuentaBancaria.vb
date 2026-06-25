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
        vTxtNombre = TxtNombre.Text
        vTxtNumero = TxtNumero.Text
        vTxtNotas = TxtNota.Text
        ' EXTRAEMOS EL VALOR INTERNO ORIGINAL PARA LA BASE DE DATOS
        If CmbTipoCuenta.SelectedItem IsNot Nothing Then
            Dim itemSeleccionado As ElementoCombo = CType(CmbTipoCuenta.SelectedItem, ElementoCombo)
            vTxtTipo = itemSeleccionado.ValorInterno
        Else
            vTxtTipo = ""
        End If

        ' Modificar Registro
        '*******************
        ' 1. Limpias la consulta cambiando todas las comillas por "?"
        vtipoSql = "UPDATE cuentas SET NumeroCUE = ?, TipoCUE = ?, NotasCUE = ? WHERE cuentas.NombreCUE = ?"
        cmdMdb1cr.CommandText = vtipoSql

        ' 2. Limpias los parámetros para evitar acumulaciones
        cmdMdb1cr.Parameters.Clear()

        ' 3. Añades los 4 valores en el orden exacto de aparición
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNumero)
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtTipo)
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNotas)
        cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre) ' El WHERE va al final
        cmdMdb1cr.CommandText = vtipoSql

        Try
            cmdMdb1cr.ExecuteNonQuery()
            Me.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorModificarRegistro"))
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        vTxtNombre = TxtNombre.Text
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("EliminarCuenta") & " " & vTxtNombre & " " & rmse.GetString("EliminarCuenta2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))
        If respuesta = vbYes Then

            ' Variable para medir si realmente se borró algo en los apuntes
            Dim filasAfectadas As Integer = 0

            ' --- 1. ELIMINAR REGISTRO CUENTAS ---
            vtipoSql = "DELETE FROM cuentas WHERE cuentas.NombreCUE = ?"
            cmdMdb1cr.CommandText = vtipoSql

            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre)
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("EliminarCuenta3"))
            Catch ex As Exception
                MsgBox(rmse.GetString("EliminarCuenta4") & vbNewLine & ex.Message)
                Exit Sub ' Si no se pudo eliminar la cuenta, no intentamos eliminar los apuntes relacionados
            End Try

            ' --- 2. ELIMINAR REGISTROS APUNTES ---
            vtipoSql = "DELETE FROM apuntes WHERE apuntes.CuentaAPU = ?"
            cmdMdb1cr.CommandText = vtipoSql

            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre)
            Try
                ' Capturamos cuántos apuntes reales se eliminan
                filasAfectadas = cmdMdb1cr.ExecuteNonQuery()

                ' FILTRO: Solo muestra el MsgBox si realmente existían y se borraron apuntes
                If filasAfectadas > 0 Then
                    MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntes"))
                End If
            Catch ex As Exception
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntesError") & vbNewLine & ex.Message)
            End Try

            ' --- 3. ELIMINAR REGISTROS APUNTES PERIÓDICOS ---
            vtipoSql = "DELETE FROM apuper WHERE apuper.CuentaAPP = ?"
            cmdMdb1cr.CommandText = vtipoSql

            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre)
            Try
                ' Capturamos cuántos apuntes periódicos se eliminan
                filasAfectadas = cmdMdb1cr.ExecuteNonQuery()

                ' FILTRO: Solo muestra el MsgBox si realmente existían y se borraron apuntes periódicos
                If filasAfectadas > 0 Then
                    MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
                End If
            Catch ex As Exception
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicosError") & vbNewLine & ex.Message)
            End Try
        End If
        Me.Close()
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