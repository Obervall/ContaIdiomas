Public Class AyudaApuntes

    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub AyudaApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Cargamos el bloque de texto traducido desde tu gestor de recursos
        ' (Donde "TextoManualCompleto" es la clave en tu archivo .resx que guarda los pasos)
        MiCuadroTextoAyuda.Text = rmse.GetString("TextoManualCompleto")
    End Sub

End Class