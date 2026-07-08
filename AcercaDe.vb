Public Class AcercaDe

    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Me.Close()
    End Sub

    Private Sub AcercaDe_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.KeyPreview = True
        ' Traducción de la versión
        If vHayNuevaVersion = "SI" Then
            LblVersion.Text = rmse.GetString("VersionInstalada") & ": " & My.Settings.Version & " - " & rmse.GetString("VersionDisponible") & ": " & vNuevaVersion
        Else
            LblVersion.Text = rmse.GetString("VersionInstalada") & ": " & My.Settings.Version
        End If

    End Sub
End Class