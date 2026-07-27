Imports System.Diagnostics

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

    Private Sub BtnPrivacidad_Click(sender As Object, e As EventArgs) Handles BtnPrivacidad.Click
        ' Ruta local hacia tus Documentos
        Dim carpetaDocumentos As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        Dim carpetaAppOficial As String = IO.Path.Combine(carpetaDocumentos, "ContaHogar3.0")
        Dim rutaCompletaPDF As String = IO.Path.Combine(carpetaAppOficial, "Politica de Privacidad.pdf")

        ' URL de respaldo en la nube
        Dim urlPrivacidad As String = "https://1drv.ms/b/c/1e195f066363b218/IQDyv4fHzLkjQZWFc7UmbvqJATQLSZyMXUcM-pzaiWvA1Ow?e=WhFxNG"
        Try
            Dim Proceso As New Process()

            If IO.File.Exists(rutaCompletaPDF) Then
                ' Opción 1: Abre el archivo PDF local de forma nativa
                Proceso.StartInfo.FileName = rutaCompletaPDF
                Proceso.StartInfo.Verb = "open"
                Proceso.Start()
            Else
                ' Opción 2: Si no existe el local, abre la URL en el navegador predeterminado
                Proceso.StartInfo.FileName = urlPrivacidad
                ' Crucial para .NET Core/.NET 5+ y la compatibilidad con el entorno MSIX
                Proceso.StartInfo.UseShellExecute = True
                Proceso.Start()
            End If

        Catch ex As Exception
            ' Si ambos métodos fallan (por ejemplo, sin internet y sin archivo local)
            Dim msgFalta As String = resManager.GetString("ErrorArchivoPrivacidadNoEncontrado")
            If String.IsNullOrEmpty(msgFalta) Then
                msgFalta = "No se pudo abrir la política de privacidad ni de forma local ni remota."
            End If
            MsgBox(msgFalta & vbCrLf & ex.Message, vbExclamation, resManager.GetString("Aviso"))
        End Try
    End Sub
End Class