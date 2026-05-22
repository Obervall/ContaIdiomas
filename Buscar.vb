Imports System.Windows.Forms

Public Class Buscar

    Public vBuscar As String

    Private Sub Buscar_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CmbCampos.DropDownStyle = ComboBoxStyle.DropDownList
        CmbCampos.SelectedIndex = 0
        CmbTextoBuscar.Select()
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

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vBuscar = CmbTextoBuscar.Text
        CmbTextoBuscar.Items.Add(vBuscar)
        Me.Close()
    End Sub

    Private Sub CmbTextoBuscar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbTextoBuscar.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

End Class