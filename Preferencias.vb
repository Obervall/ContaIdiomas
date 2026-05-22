Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Drawing

Public Class Preferencias
    Private estaCargado As Boolean = False
    Private TL(1) As ToolTip

    Private Sub Preferencias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        estaCargado = False ' Se establece la variable estaCargado a False al iniciar la carga
        ' del formulario para evitar que los eventos de cambio de las preferencias se ejecuten antes de cargar los valores.

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.TxtBaseDatos, My.Settings.RutaBD)
        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox2.MouseMove, AddressOf VerificarFiltrosDesactivados

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
        End If

        If My.Settings.Actualizar Then
            ChkBuscarActualizacion.Checked = True
            BtnBuscarActualizacion.Enabled = False
        Else
            ChkBuscarActualizacion.Checked = False
            BtnBuscarActualizacion.Enabled = True
        End If
        TxtPathExportar.Text = My.Settings.PathExportar
        ChbPantallaCompleta.Checked = My.Settings.PantallaCompleta
        ChbPantallaCierre.Checked = My.Settings.PantallaCierre
        CheckBox1.Checked = My.Settings.Previsualizar
        CheckBox2.Checked = My.Settings.ElegirImpresora
        CheckBox3.Checked = My.Settings.DirectoImpresora
        TxtBaseDatos.Text = My.Settings.RutaBD

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
        My.Settings.Actualizar = ChkBuscarActualizacion.Checked
        My.Settings.PathExportar = TxtPathExportar.Text
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub ChbPantallaCompleta_Click(sender As Object, e As EventArgs) Handles ChbPantallaCompleta.Click
        ChbPantallaCompleta.Checked = True
        ChbPantallaCierre.Checked = False
        My.Settings.PantallaCompleta = True
        My.Settings.PantallaCierre = False
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub ChbPantallaCierre_Click(sender As Object, e As EventArgs) Handles ChbPantallaCierre.Click
        ChbPantallaCompleta.Checked = False
        ChbPantallaCierre.Checked = True
        My.Settings.PantallaCompleta = False
        My.Settings.PantallaCierre = True
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        CheckBox1.Checked = True
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        My.Settings.Previsualizar = True
        My.Settings.ElegirImpresora = False
        My.Settings.DirectoImpresora = False
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = True
        CheckBox3.Checked = False
        My.Settings.Previsualizar = False
        My.Settings.ElegirImpresora = True
        My.Settings.DirectoImpresora = False
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        My.Settings.Previsualizar = False
        My.Settings.ElegirImpresora = False
        My.Settings.DirectoImpresora = True
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub BtnUbicacion_Click(sender As Object, e As EventArgs) Handles BtnUbicacion.Click
        restore.InitialDirectory = "C:\ContaHogar3.0\"
        restore.Title = "Ubicación Base de Datos"
        restore.CheckFileExists = False
        restore.CheckPathExists = False
        restore.DefaultExt = "mdb"
        restore.Filter = "Access (ContaHogar.mdb)|ContaHogar.mdb|All files (*.*)|*.*"
        restore.RestoreDirectory = True
        If restore.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                TxtBaseDatos.Text = restore.FileName
                My.Settings.RutaBD = TxtBaseDatos.Text
                My.Settings.Save()
                My.Settings.Reload()
            Catch ex As Exception
                MsgBox("Error al seleccionar la nueva ubicación de la Base de Datos, se mantendrá la ubicación actual.")
                MsgBox(ex.ToString)
            End Try
            MsgBox("Se cerrará el programa para iniciar con la nueva Ubicación.")
            End
        End If
    End Sub

    Private Sub ChkBuscarActualizacion_CheckedChanged(sender As Object, e As EventArgs) Handles ChkBuscarActualizacion.CheckedChanged
        If estaCargado Then
            If ChkBuscarActualizacion.Checked = True Then
                My.Settings.Actualizar = True
                BtnBuscarActualizacion.Enabled = False
                MsgBox("Se desactivará la búsqueda manual de actualizaciones, se buscará automáticamente al iniciar el programa.")
            Else
                My.Settings.Actualizar = False
                BtnBuscarActualizacion.Enabled = True
                MsgBox("Se activará la búsqueda manual de actualizaciones, se buscará al hacer click en el botón 'Buscar Actualización'.")
                BtnBuscarActualizacion.PerformClick()
            End If
            My.Settings.Save()
            My.Settings.Reload()
        End If
    End Sub

    Private Sub BtnBuscarActualizacion_Click(sender As Object, e As EventArgs) Handles BtnBuscarActualizacion.Click
        vActualizar = True
        BuscarActualizacion()
        If vHayNuevaVersion = "SI" Then
            MsgBox("Versión Instalada: " & My.Settings.Version & " - Versión disponible: " & vNuevaVersion)
        Else
            MsgBox("No hay nuevas versiones disponibles. Versión actual: " & My.Settings.Version)
        End If
    End Sub

    Private Sub BtnBuscarBackup_Click(sender As Object, e As EventArgs) Handles BtnBuscarBackup.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            ' List files in the folder.
            'ListFiles(FolderBrowserDialog1.SelectedPath)
            TxtPathExportar.Text = FolderBrowserDialog1.SelectedPath
            My.Settings.PathExportar = TxtPathExportar.Text
            My.Settings.Save()
            My.Settings.Reload()
        End If
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
            End If
            My.Settings.CulturaUsuario = cultura
            My.Settings.Save()
            My.Settings.Reload()

            ' Llamamos a la función del módulo
            My.Resources.Culture = New System.Globalization.CultureInfo(cultura)
            CambiarIdiomaGlobal(cultura)

            ' 1. Cambiar la cultura del hilo (para nuevos formularios)
            Dim nuevaCultura As New System.Globalization.CultureInfo(cultura)
            System.Threading.Thread.CurrentThread.CurrentUICulture = nuevaCultura
            System.Threading.Thread.CurrentThread.CurrentCulture = nuevaCultura
            System.Globalization.CultureInfo.DefaultThreadCurrentCulture = nuevaCultura
            System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = nuevaCultura

            ' 2. Aplicar a todos los formularios abiertos
            For Each f As Form In Application.OpenForms
                'MsgBox("Aplicando idioma a: " & f.Name)
                If TypeOf f Is Preferencias Then
                    ActualizarTextosFormulario(f)
                End If
                ' 3. Refrescar el formulario Principal
                If TypeOf f Is Principal Then
                    Dim frmPrincipal = DirectCast(f, Principal)
                    frmPrincipal.RefrescarTextos()
                    frmPrincipal.RefrescarMenus()
                End If
                ' 4. Aplicamos los recursos a cada control del formulario
                Dim resManager As New ComponentResourceManager(f.GetType())
                AplicarRecursosAControles(f, resManager)
                ' 5. Opcional: Actualiza el título del formulario
                If TypeOf f Is Preferencias Then
                    resManager.ApplyResources(f, "$this")
                End If
            Next
        End If
    End Sub

    Private Sub CmbMonedas_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbMonedas.SelectedIndexChanged
        My.Settings.Moneda = CmbMonedas.SelectedItem.ToString()
        vMoneda = My.Settings.Moneda
        My.Settings.Save()
        My.Settings.Reload()
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus controles deshabilitados y sus ToolTips correspondientes
        Dim controlesBloqueados As New Dictionary(Of Control, ToolTip) From {
            {Me.TxtBaseDatos, TL(0)}
        }

        For Each par In controlesBloqueados
            Dim control As Control = par.Key
            Dim tool As ToolTip = par.Value

            If Not control.Enabled Then
                ' Traducimos la posición del ratón al contenedor nativo del control (su GroupBox)
                Dim posRatonRelativaAlControl As Point = control.Parent.PointToClient(Cursor.Position)

                ' Si el ratón está sobre el control desactivado
                If control.Bounds.Contains(posRatonRelativaAlControl) Then
                    ' Calculamos la posición respecto al formulario para dibujar el cartelito en el lugar correcto
                    Dim posRatonRelativaAlForm As Point = Me.PointToClient(Cursor.Position)
                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso
                    tool.Show(My.Settings.RutaBD, Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón no está sobre ningún botón bloqueado, ocultamos los tres
        TL(0).Hide(Me)
    End Sub
End Class