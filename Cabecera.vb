Imports System.Data
Imports System.Windows.Forms

Public Class Cabecera
    Private Sub Cabecera_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 1. Rotulamos el título dinámico con tu frase exacta (ej: "2 EMPRESAS EXISTENTES")
        LblTitulo.Text = $"{vEmpresasExistentes} {resManager.GetString("EmpresasExistentes")}"
        ' Si no tienes la clave traducida todavía, puedes poner el texto fijo:
        'LblTitulo.Text = vEmpresasExistentes & " EMPRESAS / USUARIOS"

        ' =========================================================================
        ' 🔒 CORTAFUEGOS ANTI-EDICIÓN Y ESTÍLO VISUAL SUAVE PREMIUM
        ' =========================================================================
        With DgvUsuarios
            .ReadOnly = True               ' Bloquea rígidamente que se pueda escribir
            .AllowUserToAddRows = False    ' Elimina la fila del asterisco (*) del final
            .AllowUserToDeleteRows = False ' Prohíbe borrar perfiles con la tecla Supr
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect ' Fila azul completa Premium

            ' Opcional: Centramos el texto del encabezado para que quede más simétrico
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            ' Opcional: Quitamos el borde tridimensional feo para que quede plano y moderno
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
        End With

        ' 2. Alimentamos el DataGridView devorando la tabla central
        ' Usamos la variable global 'carpetaDB' que ya tiene la ruta elástica de Documentos
        Dim rutaUsuariosMdb As String = System.IO.Path.Combine(carpetaDB, "usuarios.mdb")
        Dim dt As New DataTable()

        Using conn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & rutaUsuariosMdb & ";")
            Using cmd As New OleDb.OleDbCommand("SELECT IdUsuario, NombreUsuario, ArchivoMdb FROM usuarios ORDER BY IdUsuario ASC", conn)
                Try
                    conn.Open()
                    Dim dr As OleDb.OleDbDataReader = cmd.ExecuteReader()
                    dt.Load(dr)
                    conn.Close()
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorListarEmpresas") & ": " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End Using
        End Using

        ' =========================================================================
        ' 🎯 ASIGNACIÓN DIFERIDA ANTI-REENTRADA CON MEMORIA DE SELECCIÓN
        ' =========================================================================
        Me.BeginInvoke(Sub()
                           DgvUsuarios.DataSource = Nothing
                           DgvUsuarios.Columns.Clear()

                           ' Inyectamos los datos en calma total
                           DgvUsuarios.DataSource = dt

                           ' Configuración estética Premium estilo GesFácil (Lectura Estricta)
                           If DgvUsuarios.Columns.Count >= 3 Then
                               DgvUsuarios.Columns("IdUsuario").Visible = False
                               DgvUsuarios.Columns("ArchivoMdb").Visible = False

                               DgvUsuarios.Columns("NombreUsuario").HeaderText = resManager.GetString("EmpresaUsuario")
                               DgvUsuarios.Columns("NombreUsuario").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                           End If

                           ' =========================================================================
                           ' 🚀 ¡EL TRUCO MAESTRO DE MEMORIA GESFÁCIL!
                           ' =========================================================================
                           ' Rescatamos el nombre del último archivo activo desde tus settings (ej: "maria.mdb")
                           Dim ultimoArchivoUsado As String = System.IO.Path.GetFileName(My.Settings.RutaBD).ToLower().Trim()

                           If Not String.IsNullOrEmpty(ultimoArchivoUsado) AndAlso DgvUsuarios.Rows.Count > 0 Then
                               ' Recorremos la rejilla para buscar en qué fila está ese archivo
                               For i As Integer = 0 To DgvUsuarios.Rows.Count - 1
                                   Dim archivoFila As String = DgvUsuarios.Rows(i).Cells("ArchivoMdb").Value.ToString().ToLower().Trim()

                                   ' Si coincide, le clavamos la selección azul a esa fila entera
                                   If archivoFila = ultimoArchivoUsado Then
                                       ' Quitamos cualquier selección por defecto que ponga Windows
                                       DgvUsuarios.ClearSelection()

                                       ' Activamos la fila guardada como la celda corriente y la pintamos de azul
                                       DgvUsuarios.CurrentCell = DgvUsuarios.Rows(i).Cells("NombreUsuario")
                                       DgvUsuarios.Rows(i).Selected = True

                                       Exit For ' Salimos del bucle con éxito absoluto
                                   End If
                               Next
                           End If
                       End Sub)

    End Sub

    Private Sub DgvUsuarios_DoubleClick(sender As Object, e As EventArgs) Handles DgvUsuarios.DoubleClick
        ' 3. DOBLE CLIC SELECCIONA Y CONECTA AL INSTANTE (Tu motor inteligente)
        If DgvUsuarios.CurrentRow IsNot Nothing AndAlso DgvUsuarios.Rows.Count > 0 Then
            archivoSeleccionado = DgvUsuarios.CurrentRow.Cells("ArchivoMdb").Value.ToString().Trim()

            txtAvisoBdActiva = archivoSeleccionado
            ' Asignamos la ruta a tu variable global reina 'vRuta'
            vRuta = System.IO.Path.Combine(carpetaDB, archivoSeleccionado)
            archivoAConectar = vRuta
            My.Settings.RutaBD = vRuta
            My.Settings.Save()

            ' Cerramos la ventana de selección devolviendo el control de forma limpia
            Me.Close()
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        ' 1. CORTAFUEGOS DE SEGURIDAD BANCARIA Y DE ACCESS
        ' Nos aseguramos de cerrar de forma ordenada cualquier grifo de conexión abierto en la RAM
        Try
            If conexion1 IsNot Nothing AndAlso conexion1.State = ConnectionState.Open Then
                conexion1.Close()
            End If
        Catch
            ' Pasamos de largo si la conexión ya estaba cerrada
        End Try

        ' 2. ESTRANGULAMIENTO TOTAL DE LA APLICACIÓN
        ' Cerramos todas las ventanas abiertas en el monitor y descargamos los hilos de la CPU
        Application.Exit()

        ' 3. Línea de ejecución rígida de respaldo para obligar al sistema operativo a liberar el proceso
        End
    End Sub

End Class