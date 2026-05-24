Imports System.Windows.Forms

Public Class EditarCuentaBancaria

    Public vtipoSql, vTxtNombre, vTxtNumero, vTxtTipo, vTxtNotas As String
    Public filaActual As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

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
        CargarComboTipoCuentaGlobal(CmbTipoCuenta, frmCuentasBancarias.rmse)
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
        vtipoSql = "UPDATE cuentas SET NumeroCUE = '" & vTxtNumero & "' , TipoCUE = '" & vTxtTipo & "' , NotasCUE = '" & vTxtNotas & "' "
        vtipoSql += " WHERE cuentas.NombreCUE = '" & vTxtNombre & "' "
        cmdMdb1cr.CommandText = vtipoSql

        Try
            ' CORRECCIÓN: Para comandos UPDATE se usa ExecuteNonQuery, no ExecuteReader
            cmdMdb1cr.ExecuteNonQuery()
            Me.Close()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorModificarRegistro"))
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        vTxtNombre = TxtNombre.Text
        respuesta = MsgBox(rmse.GetString("EliminarCuenta") & " " & vTxtNombre & " " & rmse.GetString("EliminarCuenta2"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("LblEliminando"))
        If respuesta = vbYes Then
            ' Nota: Tus consultas DELETE son seguras porque filtran por "NombreCUE",
            ' la cual es una cadena de texto propia del usuario y libre de traducciones.
            ' Eliminar Registro Cuentas
            vtipoSql = "DELETE FROM cuentas WHERE cuentas.NombreCUE = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("EliminarCuenta3"))
            Catch ex As Exception
                MsgBox(rmse.GetString("EliminarCuenta4") & vbNewLine & ex.Message)
                Exit Sub ' Si no se pudo eliminar la cuenta, no intentamos eliminar los apuntes relacionados
            End Try
            ' Eliminar Registros Apuntes
            vtipoSql = "DELETE FROM apuntes WHERE apuntes.CuentaAPU = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntes"))
            Catch ex As Exception
                MsgBox(frmApuntesContables.rmse.GetString("EliminarApuntesError") & vbNewLine & ex.Message)
            End Try

            ' Eliminar Registros Apuntes Periódicos
            vtipoSql = "DELETE FROM apuper WHERE apuper.CuentaAPP = '" & vTxtNombre & "' "
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(frmApuntesPeriodicos.rmse.GetString("EliminarApuntesPeriodicos"))
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