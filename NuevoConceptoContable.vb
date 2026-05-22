Imports System.Windows.Forms

Public Class NuevoConceptoContable

    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtTipo, vTxtNotas As String

    Private Sub NuevoConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(4) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, "Aceptar y Salir")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoConcepto, "Seleccionar el Tipo de Concepto Contable")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtDescripcion, "Introducir una descripción para el Concepto Contable")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, "Introducir un Código para identificar el Concepto Contable. Pueden ser Números (1,2,3,4..) o Letras (Mas recomendable... Ocio, Coche, Renta, Sueldo,...)")

        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbTipoConcepto.SelectedIndex = 0

        TxtNombre.Select()
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        TxtNombre.Text = TxtNombre.Text.ToUpper
        TxtNombre.SelectionStart = Len(TxtNombre.Text)
        Dim vBusca As String
        vBusca = TxtNombre.Text.ToString
        DgvExistente.Visible = True

        ' Llenar Grid de NOMBRES EXISTENTES en CONCEPTOS
        '***********************************************
        vtipoSql = "SELECT conceptos.CodigoCON "
        vtipoSql += "FROM conceptos WHERE conceptos.CodigoCON Like '" & vBusca & "%' ORDER BY conceptos.CodigoCON"
        vtipoGrid = "NOMBRESEXISTENTES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

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
        If TxtNombre.Text <> "" Then
            If TxtNombre.Text = "SALDO" Then
                MsgBox("No se puede crear un Concepto con el Nombre SALDO", vbCritical)
                TxtNombre.Select()
                TxtNombre.SelectAll()
            Else
                vTxtNombre = TxtNombre.Text
                vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
                vTxtTipo = CmbTipoConcepto.Text
                vTxtNotas = TxtNota.Text

                ' Verificar que no se repiten Nombres en Conceptos Contables
                '***********************************************************
                vtipoSql = "SELECT * FROM conceptos WHERE conceptos.CodigoCON = '" & vTxtNombre & "' "
                vtipoGrid = "NOMBRESEXISTENTES"
                cmdMdb1cr.CommandText = vtipoSql

                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    If drMdb1.HasRows Then
                        drMdb1.Close()
                        MsgBox("El Nombre: " & vTxtNombre & ", ya existe en Conceptos Contables", vbOKOnly, "Concepto Existente")
                        TxtNombre.Select()
                    Else
                        drMdb1.Close()
                        vtipoSql = "INSERT INTO conceptos "
                        vtipoSql += "(CodigoCON, DescripcionCON, TipoCON, NotasCON) "
                        vtipoSql += "VALUES ('" & vTxtNombre & "','" & vTxtDescripcion & "','" & vTxtTipo & "','" & vTxtNotas & "')"

                        cmdMdb1cr.CommandText = vtipoSql
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                            'MsgBox("Registro, Grabado Correctamente")
                            Me.Close()
                        Catch ex As Exception
                            MsgBox("Error al Grabar el Nuevo Concepto Contable")
                            MsgBox(ex.ToString)
                        End Try
                    End If
                Catch ex As Exception
                    MsgBox("Error al verificar que el Nombre no se repite en Conceptos Contables")
                    MsgBox(ex.ToString)
                End Try
            End If
        Else
            MsgBox("NO hay Datos en Nombre ...")
            TxtNombre.Select()
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