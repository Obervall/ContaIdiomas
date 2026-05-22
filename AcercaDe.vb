Imports System.Windows.Forms

Public Class AcercaDe

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Me.Close()
    End Sub

    Private Sub AcercaDe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        If vHayNuevaVersion = "SI" Then
            LblVersion.Text = "Versión Instalada: " & My.Settings.Version & " - Versión disponible: " & vNuevaVersion
        Else
            LblVersion.Text = "Versión Instalada: " & My.Settings.Version
        End If
        LblAutoriza.Text = My.Settings.Autorizar
        LblCodigo.Text = My.Settings.Codigo
    End Sub
End Class