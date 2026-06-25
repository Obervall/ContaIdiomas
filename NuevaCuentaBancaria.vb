Imports System.Windows.Forms

Public Class NuevaCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtNumero, vTxtTipo, vTxtNotas As String
    Public TL(4) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub NuevaCuentaBancaria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        TL(4).SetToolTip(Me.TxtNombre, resManager.GetString("ToolTipNombre"))

        ' Aunque el combo se llame diferente en este formulario, funcionará igual:
        CargarComboTipoCuentaGlobal(Me.CmbTipoCuenta)
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        TxtNombre.Text = TxtNombre.Text.ToUpper
        TxtNombre.SelectionStart = Len(TxtNombre.Text)
        Dim vBusca As String
        vBusca = TxtNombre.Text.ToString
        DgvExistente.Visible = True

        ' Llenar Grid de Nombre/Código EXISTENTES en CUENTAS BANCARIAS
        '*************************************************************
        vtipoSql = "SELECT cuentas.NombreCUE "
        vtipoSql += "FROM cuentas WHERE cuentas.NombreCUE Like '" & vBusca & "%' ORDER BY cuentas.NombreCUE"
        vtipoGrid = "NOMBRESEXISTENTES2"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtNumero.Select()
        End If
    End Sub

    Private Sub TxtNumero_TextChanged(sender As Object, e As EventArgs) Handles TxtNumero.TextChanged
        TxtNumero.Text = TxtNumero.Text.ToUpper
        TxtNumero.SelectionStart = Len(TxtNumero.Text)
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtNumero.Select()
    End Sub

    Private Sub TxtNumero_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNumero.KeyPress
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
        ' 1. Capturamos los datos forzando MAYÚSCULAS en el nombre de la cuenta
        Dim nombreCuentaMayusculas As String = TxtNombre.Text.Trim().ToUpper()
        Dim numeroCuenta As String = TxtNumero.Text.Trim()
        Dim notasCuenta As String = TxtNota.Text.Trim()

        ' Como no hay cuadro de saldo en el diseño, el saldo inicial de fábrica siempre es 0
        Dim saldoInicial As Double = 0

        ' Validamos que el nombre de la cuenta no esté vacío
        If nombreCuentaMayusculas = "" Then
            Dim msgVacio As String = resManager.GetString("NoHayDatos") & ": " & resManager.GetString("Nombre")
            MessageBox.Show(msgVacio, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtNombre.Select()
            Exit Sub
        End If

        ' Validamos que haya seleccionado un Tipo de Cuenta en el ComboBox
        If CmbTipoCuenta.SelectedIndex = -1 Then
            MessageBox.Show("Por favor, seleccione un Tipo de Cuenta válido.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            CmbTipoCuenta.Select()
            Exit Sub
        End If

        ' =========================================================================
        ' 2. CONTROL DE DUPLICADOS EN MAYÚSCULAS (Inmune a fallos)
        ' =========================================================================
        vtipoSql = "SELECT COUNT(*) FROM cuentas WHERE NombreCUE = ?"
        cmdMdb1cr.CommandText = vtipoSql
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("?", nombreCuentaMayusculas)

        Try
            Dim existe As Integer = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())
            If existe > 0 Then
                Dim msgExiste As String = resManager.GetString("Nombre") & ": " & TxtNombre.Text.Trim() & ", " & resManager.GetString("Existe") & " " & rmse.GetString("$this.Text")
                MessageBox.Show(msgExiste, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                TxtNombre.Select()
                TxtNombre.SelectAll()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show("Error al verificar duplicados: " & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' =========================================================================
        ' 3. OBTENER EL ID NUMÉRICO DEL TIPO DE CUENTA SELECCIONADO
        ' =========================================================================
        ' Como la base de datos ahora guarda números enteros en TipoCUE, le sumamos 1 a la posición del ComboBox
        Dim idTipoCuentaMDB As Integer = CmbTipoCuenta.SelectedIndex + 1

        ' =========================================================================
        ' 4. CALCULAR EL SIGUIENTE ID DISPONIBLE PARA LA CUENTA (MAX + 1)
        ' =========================================================================
        Dim siguienteID As Integer = 1
        Try
            cmdMdb1cr.CommandText = "SELECT MAX(IdCuentaCUE) FROM cuentas"
            cmdMdb1cr.Parameters.Clear()
            Dim resultado As Object = cmdMdb1cr.ExecuteScalar()
            If resultado IsNot DBNull.Value AndAlso resultado IsNot Nothing Then
                siguienteID = Convert.ToInt32(resultado) + 1
            End If
        Catch ex As Exception
            siguienteID = 1
        End Try

        ' =========================================================================
        ' 5. INSERCIÓN TOTALMENTE PARAMETRIZADA (ID, TIPO, NOMBRE, NÚMERO, NOTAS, SALDO)
        ' =========================================================================
        vtipoSql = "INSERT INTO cuentas (IdCuentaCUE, TipoCUE, NombreCUE, NumeroCUE, NotasCUE, SaldoCUE) VALUES (?, ?, ?, ?, ?, ?)"
        cmdMdb1cr.CommandText = vtipoSql

        ' Limpiamos y asignamos los parámetros en el orden EXACTO del SQL para Access
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@IdCuentaCUE", siguienteID)
        cmdMdb1cr.Parameters.AddWithValue("@TipoCUE", idTipoCuentaMDB)
        cmdMdb1cr.Parameters.AddWithValue("@NombreCUE", nombreCuentaMayusculas)
        cmdMdb1cr.Parameters.AddWithValue("@NumeroCUE", numeroCuenta)
        cmdMdb1cr.Parameters.AddWithValue("@NotasCUE", notasCuenta)
        cmdMdb1cr.Parameters.AddWithValue("@SaldoCUE", saldoInicial) ' Pasamos el 0 fijo de forma segura

        Try
            cmdMdb1cr.ExecuteNonQuery()
            Me.Close() ' Registro grabado con éxito, cierra el subformulario modal
        Catch ex As Exception
            MessageBox.Show("Error al guardar la cuenta contable: " & vbNewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub TxtNombre_GotFocus(sender As Object, e As EventArgs) Handles TxtNombre.GotFocus
        PintaTxt()
    End Sub

    Public Sub PintaTxt()
        Dim Texto As TextBox
        Texto = Me.ActiveControl
        Texto.SelectionStart = 0
        Texto.SelectionLength = Texto.Text.Length
    End Sub

End Class