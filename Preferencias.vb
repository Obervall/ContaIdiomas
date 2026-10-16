Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Windows.Forms

Public Class Preferencias
    Private estaCargado As Boolean = False
    Private TL(1) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub Preferencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtPathExportar.Text = My.Settings.PathExportar
        ChbPantallaCompleta.Checked = My.Settings.PantallaCompleta
        ChbPantallaCierre.Checked = My.Settings.PantallaCierre
        CheckBox1.Checked = My.Settings.Previsualizar
        CheckBox2.Checked = My.Settings.ElegirImpresora
        CheckBox3.Checked = My.Settings.DirectoImpresora
        TxtBaseDatos.Text = My.Settings.RutaBD

        estaCargado = False ' Se establece la variable estaCargado a False al iniciar la carga
        ' del formulario para evitar que los eventos de cambio de las preferencias se
        ' ejecuten antes de cargar los valores.

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.TxtBaseDatos, My.Settings.RutaBD)
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.TxtPathExportar, My.Settings.PathExportar)
        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox2.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados

        If My.Settings.CulturaUsuario = "es-ES" Then
            CmbElegirIdioma.Text = "Español"
        ElseIf My.Settings.CulturaUsuario = "ca" Then
            CmbElegirIdioma.Text = "Català"
        ElseIf My.Settings.CulturaUsuario = "en" Then
            CmbElegirIdioma.Text = "English"
        ElseIf My.Settings.CulturaUsuario = "fr" Then
            CmbElegirIdioma.Text = "Français"
        ElseIf My.Settings.CulturaUsuario = "de" Then
            CmbElegirIdioma.Text = "Deutsch"
        ElseIf My.Settings.CulturaUsuario = "pt" Then
            CmbElegirIdioma.Text = "Português"
        ElseIf My.Settings.CulturaUsuario = "it" Then
            CmbElegirIdioma.Text = "Italiano"
        End If

        ' Array con los símbolos individuales más comunes
        Dim simbolos() As String = {"€", "$", "CHF", "£", "S/.", "R$", "¥", "₣", "₹", "₩", "₽"}
        ' Limpia elementos previos y agrega los nuevos
        CmbMonedas.Items.Clear()
        CmbMonedas.Items.AddRange(simbolos)
        ' Busca la posición exacta del símbolo guardado
        Dim index As Integer = CmbMonedas.FindStringExact(My.Settings.Moneda)
        ' Si lo encuentra (índice diferente a -1), lo selecciona
        If index <> -1 Then
            CmbMonedas.SelectedIndex = index
        End If

        ' Una vez cargados los valores de las preferencias, se establece la variable estaCargado a True
        ' para permitir que los eventos de cambio de las preferencias se ejecuten correctamente.  
        estaCargado = True
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        ' Cerramos el formulario devolviendo el valor Cancel
        DialogResult = Windows.Forms.DialogResult.Cancel
        My.Settings.RutaBD = TxtBaseDatos.Text
        My.Settings.PathExportar = TxtPathExportar.Text
        My.Settings.Save()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
        DialogResult = Windows.Forms.DialogResult.Cancel
        My.Settings.RutaBD = TxtBaseDatos.Text
        My.Settings.PathExportar = TxtPathExportar.Text
        My.Settings.Save()
    End Sub

    Private Sub ChbPantallaCompleta_Click(sender As Object, e As EventArgs) Handles ChbPantallaCompleta.Click
        If ChbPantallaCompleta.Checked Then
            ChbPantallaCompleta.Checked = True
            ChbPantallaCierre.Checked = False
            My.Settings.PantallaCompleta = True
            My.Settings.PantallaCierre = False
            My.Settings.Save()
        Else
            ChbPantallaCompleta.Checked = False
            ChbPantallaCierre.Checked = True
            My.Settings.PantallaCompleta = False
            My.Settings.PantallaCierre = True
            My.Settings.Save()
        End If
    End Sub

    Private Sub ChbPantallaCierre_Click(sender As Object, e As EventArgs) Handles ChbPantallaCierre.Click
        If ChbPantallaCierre.Checked Then
            ChbPantallaCompleta.Checked = False
            ChbPantallaCierre.Checked = True
            My.Settings.PantallaCompleta = False
            My.Settings.PantallaCierre = True
            My.Settings.Save()
        Else
            ChbPantallaCompleta.Checked = True
            ChbPantallaCierre.Checked = False
            My.Settings.PantallaCompleta = True
            My.Settings.PantallaCierre = False
            My.Settings.Save()
        End If
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        CheckBox1.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        My.Settings.Previsualizar = True
        My.Settings.ElegirImpresora = False
        My.Settings.DirectoImpresora = False
        My.Settings.Save()
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = True
        CheckBox3.Checked = False
        My.Settings.Previsualizar = False
        My.Settings.ElegirImpresora = True
        My.Settings.DirectoImpresora = False
        My.Settings.Save()
    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        My.Settings.Previsualizar = False
        My.Settings.ElegirImpresora = False
        My.Settings.DirectoImpresora = True
        My.Settings.Save()
    End Sub

    Private Sub CmbElegirIdioma_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbElegirIdioma.SelectedIndexChanged

        If estaCargado Then
            Dim cultura As String = "es-ES"
            If CmbElegirIdioma.SelectedItem.ToString() = "Español" Then
                cultura = "es-ES"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "Català" Then
                cultura = "ca"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "English" Then
                cultura = "en"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "Français" Then
                cultura = "fr"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "Deutsch" Then
                cultura = "de"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "Português" Then
                cultura = "pt"
            ElseIf CmbElegirIdioma.SelectedItem.ToString() = "Italiano" Then
                cultura = "it"
            End If

            ' 1. Guardamos la nueva cultura elegida por el usuario
            My.Settings.CulturaUsuario = cultura
            My.Settings.Save()

            ' 2. Mostramos un aviso dócil informando del reinicio inmediato
            Dim txtAviso As String = If(resManager?.GetString("AppDisplayName"), "ContaHogar 3.0 Premium")
            Dim txtMensajeReinicio As String = If(resManager?.GetString("MsgReinicioIdioma"), "La aplicación se reiniciará para aplicar el nuevo idioma.")

            MsgBox(txtMensajeReinicio, MsgBoxStyle.Information, txtAviso)

            ' 3. Cerramos todos los hilos gráficos y volvemos a arrancar desde cero
            Application.Restart()
        End If
    End Sub

    Private Sub CmbMonedas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbMonedas.SelectedIndexChanged
        My.Settings.Moneda = CmbMonedas.SelectedItem.ToString()
        vMoneda = My.Settings.Moneda
        My.Settings.Save()
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        ' 🚀 REPARADO MODO ESCÁNER: Conectado al MouseMove del formulario para sortear el bloqueo del .Enabled = False

        ' Diccionario con tus controles deshabilitados y sus ToolTips correspondientes
        Dim controlesBloqueados As New Dictionary(Of Control, ToolTip) From {
            {Me.TxtBaseDatos, TL(0)},
            {Me.TxtPathExportar, TL(1)}
        }

        ' Capturamos la posición del ratón exacta respecto al Formulario Principal (Me)
        Dim posRatonRelativaAlForm As Point = Me.PointToClient(Cursor.Position)

        For Each par In controlesBloqueados
            Dim control As Control = par.Key
            Dim tool As ToolTip = par.Value

            If Not control.Enabled Then
                ' 🎯 LA JUGADA MAESTRA: Traducimos las coordenadas al contenedor donde vive el control gris
                Dim posRatonRelativaAlPadre As Point = control.Parent.PointToClient(Cursor.Position)

                ' Si las coordenadas del ratón caen dentro del rectángulo físico del control gris
                If control.Bounds.Contains(posRatonRelativaAlPadre) Then

                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso (My.Settings o textos fijos)
                    Dim textoCartelito As String = ""
                    If control Is Me.TxtBaseDatos Then
                        textoCartelito = My.Settings.RutaBD
                    ElseIf control Is Me.TxtPathExportar Then
                        textoCartelito = My.Settings.PathExportar ' 🚀 REPARADO: Pintamos su ruta de exportar correspondiente
                    End If

                    ' Hacemos brotar el globo flotante reluciente desplazado 15 píxeles para que no lo tape el cursor
                    tool.Show(textoCartelito, Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón se sale del perímetro de los cuadros grises, apagamos los carteles de inmediato
        TL(0).Hide(Me)
        TL(1).Hide(Me)
    End Sub

    Private Sub BtnCrearUsuario_Click(sender As Object, e As EventArgs) Handles BtnCrearUsuario.Click
        ' 1. Pedimos el nombre del nuevo usuario o empresa de forma limpia
        Dim nombreNuevo As String = InputBox(rmse.GetString("IntroduceNombreUsuario") & ": ", rmse.GetString("NuevoUsuario"))
        nombreNuevo = nombreNuevo.Trim().Replace("'", "").ToUpper ' Limpiamos comillas peligrosas

        ' Cortafuegos por si el usuario cancela o lo deja vacío
        If String.IsNullOrWhiteSpace(nombreNuevo) Then Exit Sub

        Try
            ' 2. Calculamos las rutas físicas en base a tu variable elástica 'carpetaDB' (Mis Documentos)
            Dim nombreArchivoNuevo As String = nombreNuevo.Replace(" ", "_") & ".mdb"
            Dim rutaNuevaMdb As String = System.IO.Path.Combine(carpetaDB, nombreArchivoNuevo)

            ' Comprobamos si ya existe un archivo físico con ese nombre para no machacarlo
            If System.IO.File.Exists(rutaNuevaMdb) Then
                MsgBox(rmse.GetString("YaExisteArchivoMismoNombre") & ": " & nombreArchivoNuevo, MsgBoxStyle.Exclamation)
                Exit Sub
            End If

			' 3. DUPLICAMOS LA BASE DE DATOS ACTIVA ACTUAL COMO PLANTILLA PERFECTA
			' Usamos tu base de datos actual 'vRuta' para que herede la estructura exacta de Access
			System.IO.File.Copy(vRuta, rutaNuevaMdb, True)

            ' 4. PREGUNTAMOS SI DESEA TRASPASAR LA ESTRUCTURA DE CONTROL (Toque GesFácil Premium)
            ' =========================================================================
            ' 🎯 4. SANEADO Y VACIADO SELECTIVO DE LA COPIA (Tu lógica correcta)
            ' =========================================================================
            Using cmdLimpiar As New OleDb.OleDbCommand()
                cmdLimpiar.Connection = conexion1
                If conexion1.State = ConnectionState.Closed Then conexion1.Open()

                ' =========================================================================
                ' 🎯 PREGUNTA DE TRASPASO BLINDADA MULTI-IDIOMA CON MARCADOR {0}
                ' =========================================================================
                ' 1. Pescamos el nombre limpio de la base de datos de origen (ej: "Admin" o "MARIA.mdb")
                Dim archivoOrigenLimpio As String = System.IO.Path.GetFileName(vRuta)
                If archivoOrigenLimpio.ToLower() = "contahogar.mdb" Then archivoOrigenLimpio = "Admin"

                ' 2. Recuperamos el molde de texto desde tu archivo de recursos
                Dim plantillaMensaje As String = rmse.GetString("PreguntaTraspasoEmpresa")

                ' Salvavidas por si acaso la clave no estuviera escrita todavía en el .resx
                If String.IsNullOrEmpty(plantillaMensaje) Then
                    plantillaMensaje = "La Nueva base de Datos se creará Vacía." & vbCrLf & vbCrLf &
                                       "¿Deseas Traspasar y Mantener los Conceptos y las Cuentas Bancarias de la actual: {0} al Nuevo usuario {1}?"
                End If

                ' 3. FUSIÓN DINÁMICA: Inyectamos de forma dócil el nombre en el marcador {0}
                Dim mensajeFinal As String = String.Format(plantillaMensaje, archivoOrigenLimpio.ToUpper, nombreNuevo.ToUpper)
                Dim tituloVentana As String = rmse.GetString("TraspasoDatosDinámico")

                ' Invocamos tu función reina unificada
                Dim respuestaTraspaso As MsgBoxResult = ConfirmarAccionTraducida(mensajeFinal, tituloVentana)

                ' Evaluamos la respuesta usando el estándar vbYes de tu solución
                If respuestaTraspaso = vbYes Then
                    ' ESCENARIO A: SÍ QUIERE TRASPASO
                    ' Dejamos Conceptos y Cuentas intactos en la copia. Solo vaciamos el historial vivo
                    Dim tablasA_VaciarA() As String = {"apuntes", "apuper", "ejercicios", "extracto", "presupuesto", "tempapu", "temppre", "tmpprint"}
                    For Each tabla As String In tablasA_VaciarA
                        cmdLimpiar.CommandText = $"DELETE * FROM {tabla} IN '{rutaNuevaMdb}'"
                        Try : cmdLimpiar.ExecuteNonQuery() : Catch : End Try
                    Next
                Else
                    ' ESCENARIO B: NO QUIERE TRASPASO (Lo vaciamos TODO en cadena)
                    Dim tablasA_VaciarB() As String = {"apuntes", "apuper", "conceptos", "cuentas", "ejercicios", "extracto", "presupuesto", "tempapu", "temppre", "tmpprint"}
                    For Each tabla As String In tablasA_VaciarB
                        cmdLimpiar.CommandText = $"DELETE * FROM {tabla} IN '{rutaNuevaMdb}'"
                        Try : cmdLimpiar.ExecuteNonQuery() : Catch : End Try
                    Next
                End If
            End Using

            ' =========================================================================
            ' 🎯 5. INSCRIPCIÓN EN EL BÚNKER CENTRAL (usuarios.mdb en Documentos)
            ' =========================================================================
            Dim rutaUsuariosMdb As String = System.IO.Path.Combine(carpetaDB, "usuarios.mdb")

            Using connCentral As New OleDb.OleDbConnection("Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & rutaUsuariosMdb & ";")
                Dim sqlInsertUsuario As String = "INSERT INTO usuarios (NombreUsuario, ArchivoMdb) VALUES (?, ?)"

                Using cmdCentral As New OleDb.OleDbCommand(sqlInsertUsuario, connCentral)
                    cmdCentral.Parameters.Add("@Nombre", OleDbType.VarChar).Value = nombreNuevo
                    cmdCentral.Parameters.Add("@Archivo", OleDbType.VarChar).Value = nombreArchivoNuevo

                    connCentral.Open()
                    cmdCentral.ExecuteNonQuery()
                    connCentral.Close()
                End Using
            End Using

            ' =========================================================================
            ' 🔄 REINICIO INTEGRAL CONTROLADO PREMIUM (Tu nueva regla de estabilidad)
            ' =========================================================================
            ' 1. Cerramos rígidamente la conexión activa para no corromper Access
            If conexion1 IsNot Nothing AndAlso conexion1.State = ConnectionState.Open Then
                conexion1.Close()
            End If

            ' 2. Avisamos al usuario del éxito rotundo
            MsgBox("¡Usuario '" & nombreNuevo & "' creado y registrado con éxito!" & vbCrLf &
                   "La aplicación se va a reiniciar de inmediato para cargar el nuevo perfil.", MsgBoxStyle.Information, "Nuevo Perfil Activado")

            ' 3. EL TRUCO DE REINICIO DE .NET: Levanta una nueva instancia del programa en background
            Application.Restart()

            ' 4. Estrangulamos de inmediato la instancia vieja para liberar el procesador
            End

        Catch ex As Exception
            MsgBox("Error al generar el nuevo usuario contable: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    '        ' =========================================================================
    '        ' 🎯 5. INSCRIPCIÓN EN EL BÚNKER CENTRAL (usuarios.mdb en Documentos)
    '        ' =========================================================================
    '        Dim rutaUsuariosMdb As String = System.IO.Path.Combine(carpetaDB, "usuarios.mdb")

    '        Using connCentral As New OleDb.OleDbConnection("Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & rutaUsuariosMdb & ";")
    '            Dim sqlInsertUsuario As String = "INSERT INTO usuarios (NombreUsuario, ArchivoMdb) VALUES (?, ?)"

    '            Using cmdCentral As New OleDb.OleDbCommand(sqlInsertUsuario, connCentral)
    '                cmdCentral.Parameters.Add("@Nombre", OleDbType.VarChar).Value = nombreNuevo
    '                cmdCentral.Parameters.Add("@Archivo", OleDbType.VarChar).Value = nombreArchivoNuevo

    '                connCentral.Open()
    '                cmdCentral.ExecuteNonQuery()
    '                connCentral.Close()
    '            End Using
    '        End Using

    '        ' =========================================================================
    '        ' 6. ASENTAMOS LOS CAMBIOS EN EL MOTOR GENERAL
    '        ' =========================================================================
    '        ' Actualizamos el cuadro de texto de la ubicación con la nueva ruta para que el usuario opere ya con ella
    '        TxtBaseDatos.Text = rutaNuevaMdb
    '        vRuta = rutaNuevaMdb
    '        My.Settings.RutaBD = TxtBaseDatos.Text
    '        My.Settings.Save()

    '        MsgBox("¡Usuario '" & nombreNuevo & "' creado y registrado con éxito! Se ha seleccionado como tu base de datos activa.", MsgBoxStyle.Information)

    '    Catch ex As Exception
    '        MsgBox("Error al generar el nuevo usuario contable: " & ex.Message, MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    Private Sub BtnCambiarEmpresa_Click(sender As Object, e As EventArgs) Handles BtnCambiarEmpresa.Click
        ' ESCUDO PREVENTIVO: Limpiamos el foco activo de Preferencias para evitar tirones en ordenadores lentos
        Me.ActiveControl = Nothing
        Application.DoEvents()

        Try
            ' 1. Abrimos de nuevo tu magnífico selector síncrono Cabecera.vb
            ' Bloqueará esta ventana hasta que el usuario haga doble clic en la lista
            frmCabecera.ShowDialog()

            ' 2. Cuando el usuario hace doble clic en la lista, Cabecera se cierra y la nueva ruta queda en vRuta.
            ' Actualizamos el cuadro de texto visual de la ubicación con el nuevo archivo elegido
            TxtBaseDatos.Text = vRuta

            ' 3. 🎯 EL RE-CABLEADO EN CALIENTE: Cerramos el grifo viejo y abrimos el nuevo
            If conexion1 IsNot Nothing AndAlso conexion1.State = ConnectionState.Open Then
                conexion1.Close()
            End If

            ' Invocamos tu rutina estrella de conexión pasándole tu variable vRuta ya actualizada
            My.Settings.RutaBD = vRuta
            My.Settings.Save()

            Conectarse("AccessMdb")

            ' =========================================================================
            ' 🎯 REPARADO: FORZAMOS EL REFRESCO EN CALIENTE DEL MENÚ PRINCIPAL
            ' =========================================================================
            ' Sincronizamos la variable global con el nuevo usuario activo de la RAM
            Dim nombreEmpresaNueva As String = System.IO.Path.GetFileNameWithoutExtension(vRuta)
            If nombreEmpresaNueva.ToLower() = "contahogar" Then nombreEmpresaNueva = "Admin"
            txtAvisoBdActiva = nombreEmpresaNueva.ToUpper() ' Asegúrate de actualizar la variable global aquí

            ' Invocamos tu función pasándole de forma rígida la instancia de tu Formulario Principal
            ' (Ajusta 'frmPrincipal' al nombre técnico o variable global de tu ventana madre)
            ActualizarTextosFormulario(frmPrincipal)
            ' ===============================================================

            MsgBox($"Conectado con éxito a la empresa: '{nombreEmpresaNueva}'. Los datos se han actualizado en pantalla.", MsgBoxStyle.Information, "Cambio de Empresa")

            ' 5. Cerramos el formulario de Preferencias automáticamente para que el usuario vea su nuevo panel
            Me.Close()

        Catch ex As Exception
            MsgBox("Error al cambiar de empresa en caliente: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Class