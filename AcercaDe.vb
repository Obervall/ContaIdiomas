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

        ' --- TRADUCCIÓN SEGURA DE LBLAUTORIZA ---
        Dim autorizarOriginal As String = My.Settings.Autorizar
        Dim prefijoDemo As String = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo"
        Dim prefijoActivo As String = "Se autoriza el uso a: "

        If autorizarOriginal = prefijoDemo Then
            ' Si está en modo demo original, ponemos la traducción limpia
            LblAutoriza.Text = rmse.GetString("TextoAutorizaUso") & ": " & frmActivarSoftware.rmse.GetString("ModoDemo")
        ElseIf autorizarOriginal.StartsWith(prefijoActivo) Then
            ' Si está activado, extraemos el nombre del propietario (ej: "Juan Pérez")
            Dim nombrePropietario As String = autorizarOriginal.Substring(prefijoActivo.Length)
            ' Combinamos el prefijo traducido con su nombre real
            LblAutoriza.Text = rmse.GetString("TextoAutorizaUso") & ": " & nombrePropietario
        Else
            ' Por si acaso hay un formato antiguo o inesperado
            LblAutoriza.Text = autorizarOriginal
        End If

        ' --- TRADUCCIÓN SEGURA DE LBLCODIGO ---
        Dim codigoOriginal As String = My.Settings.Codigo
        Dim prefijoCodigo As String = "Codigo Activación: "

        If Not String.IsNullOrEmpty(codigoOriginal) AndAlso codigoOriginal.StartsWith(prefijoCodigo) Then
            ' Extraemos solo la parte derecha del texto
            Dim serial As String = codigoOriginal.Substring(prefijoCodigo.Length)
            If serial = "Sin Activar" Then
                serial = rmse.GetString("SinActivar")
            End If
            ' Combinamos con la traducción correspondiente
            LblCodigo.Text = rmse.GetString("TextoCodigoActivacion") & ": " & serial
        Else
            LblCodigo.Text = codigoOriginal
        End If

    End Sub
End Class