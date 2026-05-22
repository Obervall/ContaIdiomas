Imports System.Windows.Forms

Public Class FiltroF5

    Private BtnAceptarPulsado As Integer

    Private Sub FiltroF5_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        ChkOtrosFiltros.Enabled = False
        ChkOtrosFiltros.Checked = False
        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Or frmApuntesContables.BtnFiltroConcepto.Enabled = False Or frmApuntesContables.BtnFiltroFecha.Enabled = False Or frmApuntesContables.BtnFiltroChekedList.Enabled = False Then
            ChkOtrosFiltros.Enabled = True
            ChkOtrosFiltros.Checked = True
        End If
        BtnAceptar.Select()
        BtnAceptarPulsado = 0
        CmbCampos.SelectedIndex = 0
        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        TxtFiltro.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
        TxtFiltro.Select()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        BtnAceptarPulsado = 0
        TxtFiltro.Text = ""
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es o no es el botón X.
        If BtnAceptarPulsado = 0 Then
            TxtFiltro.Text = ""
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        BtnAceptarPulsado = 1
        Me.Close()
    End Sub

    Private Sub TxtFiltro_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtFiltro.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub TxtFiltro_Click(sender As Object, e As EventArgs) Handles TxtFiltro.Click
        TxtFiltro.SelectAll()
    End Sub
End Class