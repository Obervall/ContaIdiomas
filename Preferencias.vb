Imports System.Collections.Generic
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
        My.Settings.Reload()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub ChbPantallaCompleta_Click(sender As Object, e As EventArgs) Handles ChbPantallaCompleta.Click
        If ChbPantallaCompleta.Checked Then
            ChbPantallaCompleta.Checked = True
            ChbPantallaCierre.Checked = False
            My.Settings.PantallaCompleta = True
            My.Settings.PantallaCierre = False
            My.Settings.Save()
            My.Settings.Reload()
        Else
            ChbPantallaCompleta.Checked = False
            ChbPantallaCierre.Checked = True
            My.Settings.PantallaCompleta = False
            My.Settings.PantallaCierre = True
            My.Settings.Save()
            My.Settings.Reload()
        End If
    End Sub

    Private Sub ChbPantallaCierre_Click(sender As Object, e As EventArgs) Handles ChbPantallaCierre.Click
        If ChbPantallaCierre.Checked Then
            ChbPantallaCompleta.Checked = False
            ChbPantallaCierre.Checked = True
            My.Settings.PantallaCompleta = False
            My.Settings.PantallaCierre = True
            My.Settings.Save()
            My.Settings.Reload()
        Else
            ChbPantallaCompleta.Checked = True
            ChbPantallaCierre.Checked = False
            My.Settings.PantallaCompleta = True
            My.Settings.PantallaCierre = False
            My.Settings.Save()
            My.Settings.Reload()
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
        My.Settings.Reload()
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

End Class