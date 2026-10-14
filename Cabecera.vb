Imports System.Data
Imports System.Windows.Forms

Public Class Cabecera
    Private Sub Cabecera_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub DgvUsuarios_DoubleClick(sender As Object, e As EventArgs) Handles DgvUsuarios.DoubleClick
        ' 1. Validamos que haya un usuario seleccionado de la lista
        If DgvUsuarios.CurrentRow IsNot Nothing AndAlso DgvUsuarios.Rows.Count > 0 Then

            ' Pescamos el nombre del archivo físico asociado (ej: "maria.mdb")
            ' [Asegúrate de darle el .Name = "ArchivoMdb" a esta columna al cargar la lista]
            Dim archivoUsuario As String = DgvUsuarios.CurrentRow.Cells("ArchivoMdb").Value.ToString().Trim()

            If Not String.IsNullOrEmpty(archivoUsuario) Then
                ' 2. CALCULAMOS LA RUTA BIOLÓGICA REAL DEL ARCHIVO DEL USUARIO
                ' Asumimos que los guardas dentro de una carpeta llamada "Datos" en la ruta del ejecutable
                Dim rutaBaseDatosUsuario As String = System.IO.Path.Combine(Application.StartupPath, "Datos", archivoUsuario)

                ' 3. 🎯 EL TRUCO MAESTRO: RE-CABLEAMOS LA CONEXIÓN GLOBAL DE TU APLICACIÓN
                ' Cerramos la conexión activa actual por seguridad si estuviera abierta
                If conexion1 IsNot Nothing AndAlso conexion1.State = ConnectionState.Open Then
                    conexion1.Close()
                End If

                ' Pisamos el ConnectionString de tu variable global fija 'conexion1' 
                ' apuntando al búnker específico de este usuario
                conexion1.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & rutaBaseDatosUsuario & ";Persist Security Info=False;"

                ' Abrimos la nueva conexión. A partir de este milisegundo, CUALQUIER consulta SQL 
                ' de tus formularios leerá y guardará datos ÚNICAMENTE en el archivo de este usuario
                conexion1.Open()

                ' 4. Abrimos el formulario principal del ContaHogar 3.0 Premium y ocultamos este selector
                frmPrincipal.Show()
                Me.Hide()
            End If
        End If
    End Sub

End Class