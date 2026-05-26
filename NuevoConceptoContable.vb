Imports System.Windows.Forms

Public Class NuevoConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtTipo, vTxtNotas As String
    Public TL(4) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub NuevoConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoConcepto, rmse.GetString("SeleccionarTipo"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtDescripcion, rmse.GetString("IntroducirDescripcion"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, rmse.GetString("MsgNombre"))

        ' 1. Configuramos primero el estilo del ComboBox
        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList

        ' 2. LLENAMOS EL COMBO PRIMERO (Así tendrá elementos antes de seleccionar el índice 0)
        ActualizarIdiomaComboConcepto(Me.CmbTipoConcepto, False)

        ' 3. Ahora que ya tiene filas, seleccionamos de forma segura la primera
        CmbTipoConcepto.SelectedIndex = 0
        CmbTipoConcepto.Select()
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        ' Controlar las mayúsculas sin perder la posición del cursor
        Dim posicionCursor As Integer = TxtNombre.SelectionStart
        Dim textoMayusculas As String = TxtNombre.Text.ToUpper()

        If TxtNombre.Text <> textoMayusculas Then
            TxtNombre.Text = textoMayusculas
            TxtNombre.SelectionStart = posicionCursor
        End If

        Dim vBusca As String = TxtNombre.Text.Trim()

        ' Si el cuadro está vacío, ocultamos el Grid para que no estorbe
        If vBusca = "" Then
            DgvExistente.Visible = False
            Exit Sub
        End If

        DgvExistente.Visible = True

        ' Busca exactamente lo que el usuario escribe (sea en español, inglés o francés)
        vtipoSql = "SELECT conceptos.CodigoCON "
        vtipoSql += "FROM conceptos WHERE conceptos.CodigoCON Like '" & vBusca & "%' ORDER BY conceptos.CodigoCON"
        vtipoGrid = "NOMBRESEXISTENTES"

        ' Tu procedimiento se encarga de llenar y traducir la cabecera del Grid
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    'Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
    '    TxtNombre.Text = TxtNombre.Text.ToUpper
    '    TxtNombre.SelectionStart = Len(TxtNombre.Text)
    '    Dim vBusca As String
    '    vBusca = TxtNombre.Text.ToString
    '    DgvExistente.Visible = True

    '    ' Llenar Grid de NOMBRES EXISTENTES en CONCEPTOS
    '    '***********************************************
    '    vtipoSql = "SELECT conceptos.CodigoCON "
    '    vtipoSql += "FROM conceptos WHERE conceptos.CodigoCON Like '" & vBusca & "%' ORDER BY conceptos.CodigoCON"
    '    vtipoGrid = "NOMBRESEXISTENTES"
    '    LlenarGrid(vtipoSql, vtipoGrid, "1")
    'End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtDescripcion.Select()
        End If
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtDescripcion.Select()
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
        ' 1. Guardar la palabra escrita en mayúsculas y sin espacios a los lados
        Dim nombreLimpio As String = TxtNombre.Text.Trim().ToUpper()

        If nombreLimpio <> "" Then

            ' Obtener de forma segura la traducción de "SALDO" del idioma actual (para los 6 idiomas)
            Dim saldoTraducido As String = ""
            Try
                saldoTraducido = rmse.GetString("PalabraSaldo").Trim().ToUpper()
            Catch ex As Exception
                saldoTraducido = "SALDO" ' Respaldo por si no se encuentra la clave en el recurso
            End Try

            ' 2. Validación de bloqueo: No permite "SALDO" en español ni su traducción internacional
            If nombreLimpio = "SALDO" OrElse (saldoTraducido <> "" AndAlso nombreLimpio = saldoTraducido) Then
                MsgBox(rmse.GetString("NoNombreSaldo"), vbCritical, rmse.GetString("$this.Text"))
                TxtNombre.Select()
                TxtNombre.SelectAll()
                Exit Sub ' Detiene el guardado inmediatamente
            End If

            ' Si pasa la validación, preparamos el resto de variables
            vTxtNombre = TxtNombre.Text.Trim()
            vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
            vTxtNotas = TxtNota.Text

            ' 3. Guardado multiidioma del tipo (Inamovible en la BD como GASTO o INGRESO)
            ' Posición 0 suele ser Gasto/Expense y Posición 1 es Ingreso/Income
            If CmbTipoConcepto.SelectedIndex = 0 Then
                vTxtTipo = "GASTO"
            ElseIf CmbTipoConcepto.SelectedIndex = 1 Then
                vTxtTipo = "INGRESO"
            Else
                vTxtTipo = CmbTipoConcepto.Text ' Respaldo en caso de que cambies el orden
            End If

            ' Verificar que no se repiten Nombres en Conceptos Contables
            '***********************************************************
            vtipoSql = "SELECT * FROM conceptos WHERE conceptos.CodigoCON = '" & vTxtNombre & "' "
            vtipoGrid = "NOMBRESEXISTENTES"
            cmdMdb1cr.CommandText = vtipoSql

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    drMdb1.Close()
                    MsgBox(resManager.GetString("Nombre") & ":  " & vTxtNombre & ", " & rmse.GetString("YaExisteConcepto"), vbOKOnly, rmse.GetString("$this.Text"))
                    TxtNombre.Select()
                Else
                    drMdb1.Close()
                    vtipoSql = "INSERT INTO conceptos "
                    vtipoSql += "(CodigoCON, DescripcionCON, TipoCON, NotasCON) "
                    vtipoSql += "VALUES ('" & vTxtNombre & "','" & vTxtDescripcion & "','" & vTxtTipo & "','" & vTxtNotas & "')"

                    cmdMdb1cr.CommandText = vtipoSql
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        Me.Close() ' Registro grabado con éxito, cierra el subformulario
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        Else
            ' Mensaje de error si el campo nombre está completamente vacío
            MsgBox(rmse.GetString("MsgDatosNombre"), vbCritical, rmse.GetString("$this.Text"))
            TxtNombre.Select()
        End If
    End Sub


    'Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
    '    If TxtNombre.Text <> "" Then
    '        If TxtNombre.Text = "SALDO" Then
    '            MsgBox(rmse.GetString("NoNombreSaldo"), vbCritical, rmse.GetString("$this.Text"))
    '            TxtNombre.Select()
    '            TxtNombre.SelectAll()
    '        Else
    '            vTxtNombre = TxtNombre.Text
    '            vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
    '            vTxtTipo = CmbTipoConcepto.Text
    '            vTxtNotas = TxtNota.Text

    '            ' Verificar que no se repiten Nombres en Conceptos Contables
    '            '***********************************************************
    '            vtipoSql = "SELECT * FROM conceptos WHERE conceptos.CodigoCON = '" & vTxtNombre & "' "
    '            vtipoGrid = "NOMBRESEXISTENTES"
    '            cmdMdb1cr.CommandText = vtipoSql

    '            Try
    '                drMdb1 = cmdMdb1cr.ExecuteReader()
    '                If drMdb1.HasRows Then
    '                    drMdb1.Close()
    '                    MsgBox(resManager.GetString("Nombre") & ":  " & vTxtNombre & ", " & rmse.GetString("YaExisteConcepto"), vbOKOnly, rmse.GetString("$this.Text"))
    '                    TxtNombre.Select()
    '                Else
    '                    drMdb1.Close()
    '                    vtipoSql = "INSERT INTO conceptos "
    '                    vtipoSql += "(CodigoCON, DescripcionCON, TipoCON, NotasCON) "
    '                    vtipoSql += "VALUES ('" & vTxtNombre & "','" & vTxtDescripcion & "','" & vTxtTipo & "','" & vTxtNotas & "')"

    '                    cmdMdb1cr.CommandText = vtipoSql
    '                    Try
    '                        cmdMdb1cr.ExecuteNonQuery()
    '                        'MsgBox("Registro, Grabado Correctamente")
    '                        Me.Close()
    '                    Catch ex As Exception
    '                        MsgBox(ex.ToString)
    '                    End Try
    '                End If
    '            Catch ex As Exception
    '                MsgBox(ex.ToString)
    '            End Try
    '        End If
    '    Else
    '        MsgBox(rmse.GetString("MsgDatosNombre"), vbCritical, rmse.GetString("$this.Text"))
    '        TxtNombre.Select()
    '    End If
    'End Sub

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