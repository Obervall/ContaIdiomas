Imports System.Diagnostics
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntes

    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, strText, vIntro, vLetras, vCombo, vDescripcion As String
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(12) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Label7.Text = vMoneda
        vIntro = "NO"
        vFecha1Enero = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptarOtro, rmse.GetString("BtnAceptarOtro.Text"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptarSalir, rmse.GetString("BtnAceptarSalir.Text"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, rmse.GetString("BtnCancelar.Text") & " " & rmse.GetString("$this.Text"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, rmse.GetString("SelecConcepto"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, rmse.GetString("SelecCuenta"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, rmse.GetString("SelecDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteAsiento"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, rmse.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, rmse.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, rmse.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnDescripcion, rmse.GetString("BtnDescripcion"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.TxtBuscarLetras, rmse.GetString("TxtABuscar"))

        ' Llenar el Combo Concepto
        '*************************
        LlenarConcepto()

        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()

        ' Llenar el Combo Cuenta
        '***********************
        cmdMdb1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuenta.Items.Add(drMdb1.GetValue(0))
                End While
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    CmbCuenta.Text = CmbCuenta.Items(frmApuntesContables.CmbCuenta.SelectedIndex)
                Else
                    CmbCuenta.Text = CmbCuenta.Items(0)
                End If
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        TxtImporte.Text = 0
    End Sub

    Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ' 1. Guardamos el texto que ha seleccionado el usuario con las flechas
            Dim textoSeleccionado As String = CmbConcepto.Text
            vConcepto = textoSeleccionado

            ' 2. Apagamos el buscador de arriba para que al borrarlo no active la base de datos
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Enabled = False

            ' 3. Restauramos el texto en el combo por si el borrado de arriba hizo amago de limpiarlo
            CmbConcepto.Text = textoSeleccionado

            ' =====================================================================
            ' 🛠️ NUEVO: CARGAR LA DESCRIPCIÓN POR DEFECTO DEL CONCEPTO SELECCIONADO
            ' =====================================================================
            If Not String.IsNullOrEmpty(vConcepto) Then
                Try
                    ' Cerramos el lector si estuviera abierto por seguridad
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

                    ' Consultamos la tabla de conceptos para traer su tipo y descripción asociada
                    cmdMdb1cr.CommandText = "SELECT * FROM conceptos WHERE CodigoCON = '" & vConcepto.Replace("'", "''") & "' ORDER BY CodigoCON ASC"
                    drMdb1 = cmdMdb1cr.ExecuteReader()

                    If drMdb1.Read() Then
                        ' Rellenamos el tipo de concepto
                        TxtTipoConcepto.Text = Convert.ToString(drMdb1.GetValue(2))

                        ' Rellenamos el combo de descripciones con la descripción por defecto de la base de datos
                        CmbDescripcion.Text = Convert.ToString(drMdb1.GetValue(1))
                        vDescripcion = CmbDescripcion.Text
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    ' Manejo silencioso o registro de error para no romper el flujo
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                End Try
            End If
            ' =====================================================================

            ' 4. Cerramos la lista y saltamos limpiamente al combo de descripción
            If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False
            CmbDescripcion.Select()
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        ' ¡EL TRUCO!: Si es un Intro, salimos inmediatamente sin borrar nada
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Exit Sub
        End If

        ' 1. Capturamos la letra pulsada en mayúsculas
        Dim letra As Char = e.KeyChar

        ' Permitimos únicamente caracteres válidos (letras, números o espacio) para evitar que salte con Backspace o Enter
        If Char.IsLetterOrDigit(letra) OrElse letra = " "c Then
            ' 2. Indicamos al sistema el modo de búsqueda
            vCombo = "concepto"

            ' 3. Preparamos el cuadro de búsqueda y le inyectamos la letra directamente
            TxtBuscarLetras.Enabled = True
            TxtBuscarLetras.Text += letra.ToString()

            ' 4. Pasamos el foco al cuadro de búsqueda de forma limpia
            TxtBuscarLetras.Focus()

            ' Colocamos el cursor al final del texto en el cuadro de búsqueda
            TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
            TxtBuscarLetras.SelectionLength = 0

            ' Guardamos la variable global
            vLetras = TxtBuscarLetras.Text

            ' =================================================================
            ' 5. LIMPIEZA DE COMBOS TOTALMENTE BLINDADA
            ' =================================================================
            ' Desconectamos el evento para que no intente ejecutar consultas SQL ni leer índices vacíos
            RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

            ' Forzamos el cierre de la persiana gráfica y reseteamos el índice para liberar el puntero de Windows
            CmbConcepto.DroppedDown = False
            CmbConcepto.SelectedIndex = -1

            ' Ahora el vaciado de la colección es 100% seguro
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.Items.Clear()
            End If

            ' Volvemos a conectar el evento una vez que el control está vacío y en un estado neutro
            AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

            ' Limpiamos los textos visuales
            CmbConcepto.Text = ""
            CmbDescripcion.Text = ""

            ' =================================================================
            ' 6. ¡EL SECRETO!: Cancelamos la pulsación en el ComboBox original
            e.Handled = True
        End If
    End Sub

    Private Sub CmbConcepto_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbConcepto.MouseClick
        TxtBuscarLetras.Enabled = False
        vIntro = "NO"
        ' Solo forzamos el despliegue automático si el usuario NO ha pulsado la flecha nativa
        ' (Nos aseguramos comprobando si la lista ya está abierta o abriéndola suavemente)
        If CmbConcepto.Items.Count <> 0 AndAlso Not CmbConcepto.DroppedDown Then
            CmbConcepto.DroppedDown = True
            CmbConcepto.SelectedIndex = 0
        End If
    End Sub

    Private Sub CmbConcepto_GotFocus(sender As Object, e As EventArgs) Handles CmbConcepto.GotFocus
        ' Preparamos las variables de control en un estado neutro
        TxtBuscarLetras.Enabled = False
        vIntro = "NO"
        vCombo = "concepto"

        ' Forzamos a que el cuadro de búsqueda convierta todo a MAYÚSCULAS automáticamente
        TxtBuscarLetras.CharacterCasing = CharacterCasing.Upper

        ' 🛠️ BLINDAJE: Vaciamos el texto visual para que el combo esté totalmente limpio
        ' y no tenga ningún concepto viejo seleccionado que interfiera con tu primera letra.
        RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
        CmbConcepto.SelectedIndex = -1
        CmbConcepto.Text = ""
        AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
    End Sub


    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        If vIntro = "NO" Then
            vConcepto = CmbConcepto.Text.ToString
            drMdb1.Close()
            cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' ORDER BY conceptos.CodigoCON ASC"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            drMdb1.Read()
            If drMdb1.HasRows Then
                TxtTipoConcepto.Text = drMdb1.GetValue(2)
                CmbDescripcion.Text = drMdb1.GetValue(1)
                drMdb1.Close()
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBuscarLetras.KeyDown
        ' Si el usuario pulsa FLECHA ABAJO desde el cuadro de búsqueda, saltamos al combo de forma inteligente
        If e.KeyCode = Keys.Down Then
            If vCombo = "concepto" AndAlso CmbConcepto.Items.Count > 0 Then
                e.Handled = True
                CmbConcepto.Focus()
                CmbConcepto.SelectedIndex = 0
            ElseIf vCombo = "descripcion" AndAlso CmbDescripcion.Items.Count > 0 Then
                e.Handled = True
                CmbDescripcion.Focus()
                CmbDescripcion.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtBuscarLetras.KeyPress
        ' Detectamos si se pulsa ENTER (Ascii 13)
        If Asc(e.KeyChar) = 13 Then
            vIntro = "SI"
            e.Handled = True ' Evita el pitido de Windows

            If vCombo = "concepto" Then
                MsgBox(rmse.GetString("MsgTxtBuscarConceptos"), MsgBoxStyle.Exclamation, rmse.GetString("$this.Text"))
                If CmbConcepto.Items.Count > 0 Then
                    CmbConcepto.DroppedDown = True
                    CmbConcepto.Focus()
                    CmbConcepto.SelectedIndex = 0
                End If
            Else
                MsgBox(rmse.GetString("MsgTxtBuscarDescripciones"), MsgBoxStyle.Exclamation, rmse.GetString("$this.Text"))
                If CmbDescripcion.Items.Count > 0 Then
                    CmbDescripcion.DroppedDown = True
                    CmbDescripcion.Focus()
                    CmbDescripcion.SelectedIndex = 0
                End If
            End If
        Else
            ' Si es cualquier otra letra, ponemos vIntro en NO y dejamos que pase limpia.
            vIntro = "NO"
        End If
    End Sub

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        ' 1. Sincronizamos la variable con el texto que el usuario está escribiendo real
        vLetras = TxtBuscarLetras.Text

        ' Si el usuario ha borrado el buscador por completo, limpiamos los combos y salimos
        If String.IsNullOrEmpty(vLetras) Then
            If vCombo = "concepto" Then CmbConcepto.Items.Clear()
            If vCombo = "descripcion" Then CmbDescripcion.Items.Clear()
            Exit Sub
        End If

        ' 2. Ajustamos los TabIndex de forma estática según el modo de trabajo,
        ' pero ELIMINAMOS las líneas que vaciaban el texto (CmbDescripcion.Text = "") 
        ' para que no corten la escritura a la cuarta letra.
        If vCombo = "concepto" Then
            CmbConcepto.TabIndex = 4
            CmbDescripcion.TabIndex = 5
        ElseIf vCombo = "descripcion" Then
            CmbConcepto.TabIndex = 2
            CmbDescripcion.TabIndex = 4
        End If

        ' 3. Lanzamos la búsqueda en la base de datos de manera limpia
        BuscarLetras(vCombo)
    End Sub

    Public Function BuscarLetras(combo As String) As String
        ' Cerramos cualquier lector abierto preventivamente
        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
            drMdb1.Close()
        End If

        ' =====================================================================
        ' MODO: CONCEPTO
        ' =====================================================================
        If combo = "concepto" Then
            ' 1. Desconectamos temporalmente el evento para evitar disparos accidentales
            RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

            ' 2. ¡EL SECRETO!: Forzamos el cierre de la persiana y reseteamos el índice.
            ' Esto destruye el puntero gráfico interno de Windows sobre la fila '0'.
            CmbConcepto.DroppedDown = False
            CmbConcepto.SelectedIndex = -1

            ' 3. Ahora el vaciado es totalmente seguro y no lanzará excepciones
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.Items.Clear()
            End If

            ' 4. Volvemos a conectar el evento de forma limpia
            AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

            ' Lanzamos la SQL limpia
            Dim letrasLimpias As String = vLetras.Replace("'", "''")
            cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos WHERE CodigoCON LIKE '%" & letrasLimpias & "%' ORDER BY CodigoCON ASC"

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        Dim valor As String = Convert.ToString(drMdb1.GetValue(0))
                        If valor <> "TRASPASO" Then CmbConcepto.Items.Add(valor)
                    End While
                    CmbConcepto.DroppedDown = True
                Else
                    CmbConcepto.DroppedDown = False
                End If
                drMdb1.Close()
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorBuscarLetrasConcepto") & ": " & ex.Message)
            End Try
        End If

        ' =====================================================================
        ' MODO: DESCRIPCIÓN
        ' =====================================================================
        If combo = "descripcion" AndAlso vLetras.Length > 2 Then
            ' 1. Desconectamos el evento por seguridad
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            ' 2. ¡EL SECRETO!: Cerramos la persiana y limpiamos el índice para proteger el borrado
            CmbDescripcion.DroppedDown = False
            CmbDescripcion.SelectedIndex = -1

            ' 3. Vaciamos de forma segura
            If CmbDescripcion.Items.Count > 0 Then
                CmbDescripcion.Items.Clear()
            End If

            ' 4. Volvemos a conectar el evento
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            Dim letrasLimpias As String = vLetras.Replace("'", "''")
            cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU LIKE '%" & letrasLimpias & "%' AND DescripcionAPU <> 'Saldo Inicial'"

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        Dim desc As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
                        If Not String.IsNullOrEmpty(desc) Then CmbDescripcion.Items.Add(desc)
                    End While
                    CmbDescripcion.DroppedDown = True
                    vCombo = "descripcion" ' Mantenemos la variable de control
                Else
                    ' =====================================================================
                    ' 🛠️ MODIFICACIÓN: CONTROL PARA AÑADIR NUEVA DESCRIPCIÓN
                    ' =====================================================================
                    ' Cerramos el lector inmediatamente antes del MsgBox para liberar la BD
                    drMdb1.Close()
                    CmbDescripcion.DroppedDown = False

                    Dim respuesta As MsgBoxResult = MsgBox("No existen Descripciones con: -" & vLetras.ToUpper() & "-" & vbCrLf & "¿Añadimos la descripción?", vbQuestion + vbYesNo + vbDefaultButton1, "Introducir Apunte")

                    If respuesta = vbYes Then
                        vIntro = "SI"

                        ' Fijamos el texto que ha redactado el usuario en el combo
                        CmbDescripcion.Text = vLetras
                        vDescripcion = vLetras

                        ' Apagamos el cuadro de búsqueda superior sin disparar eventos en cadena
                        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                        TxtBuscarLetras.Text = ""
                        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                        TxtBuscarLetras.Enabled = False

                        ' Saltamos de forma limpia al importe seleccionando su contenido
                        TxtImporte.Focus()
                        TxtImporte.SelectAll()
                    Else
                        ' Si dice que NO, retiramos con cuidado la última letra que provocó el fallo
                        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                        If TxtBuscarLetras.Text.Length > 0 Then
                            TxtBuscarLetras.Text = TxtBuscarLetras.Text.Substring(0, TxtBuscarLetras.Text.Length - 1)
                        End If
                        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

                        ' Reubicamos el cursor al final de las letras restantes en el buscador
                        TxtBuscarLetras.Focus()
                        TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
                        vLetras = TxtBuscarLetras.Text
                    End If
                    ' =====================================================================
                End If

                ' Doble comprobación de seguridad para asegurar el cierre del lector
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorBuscarLetrasDescripcion") & ": " & ex.Message)
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            End Try
        End If

        Return ""
    End Function

    Private Sub CmbDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbDescripcion.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ' 1. Guardamos la descripción seleccionada con las flechas
            Dim textoSeleccionado As String = CmbDescripcion.Text

            ' 2. Apagamos el buscador de arriba de forma controlada
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Enabled = False

            ' 3. Aseguramos el texto en el combo y en tu variable global
            CmbDescripcion.Text = textoSeleccionado
            vDescripcion = textoSeleccionado

            ' 4. Cerramos el desplegable y mandamos el cursor al Importe
            If CmbDescripcion.DroppedDown Then CmbDescripcion.DroppedDown = False

            ' 💡 REEMPLAZA "TxtImporte" por el nombre exacto de tu caja de texto del importe
            TxtImporte.Select()
        End If
    End Sub


    Private Sub CmbDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbDescripcion.KeyPress
        ' ¡EL TRUCO!: Si es un Intro, salimos inmediatamente sin borrar nada
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Exit Sub
        End If

        ' Convertimos la letra a mayúscula
        Dim letra As Char = e.KeyChar

        ' Solo actuamos si es una letra, número o espacio válido y no estamos confirmando datos (vIntro)
        If vIntro = "NO" AndAlso (Char.IsLetterOrDigit(letra) OrElse letra = " "c) Then
            ' Indicamos el modo de búsqueda
            vCombo = "descripcion"

            ' Preparamos el cuadro de búsqueda e inyectamos la letra
            TxtBuscarLetras.Enabled = True
            TxtBuscarLetras.Text += letra.ToString()

            ' Pasamos el foco de forma limpia utilizando las propiedades nativas de .NET
            TxtBuscarLetras.Focus()
            TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
            TxtBuscarLetras.SelectionLength = 0

            vLetras = TxtBuscarLetras.Text

            ' Desconectamos el evento para proteger el vaciado en Descripciones
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            CmbDescripcion.DroppedDown = False
            CmbDescripcion.SelectedIndex = -1

            If CmbDescripcion.Items.Count > 0 Then
                CmbDescripcion.Items.Clear()
            End If

            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            CmbDescripcion.Text = ""

            ' ¡EL SECRETO!: Cancelamos la pulsación en el combo para evitar parpadeos gráficos
            e.Handled = True
        End If
    End Sub

    Private Sub CmbDescripcion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDescripcion.SelectedIndexChanged
        ' OPTIMIZACIÓN CRÍTICA: Eliminamos la consulta SQL redundante a la base de datos.
        ' Cuando el usuario hace clic en un elemento de la lista, el combo ya adquiere ese texto de forma nativa.
        ' No necesitas abrir drMdb1, ni hacer un SELECT, ni volver a setear el .Text.

        If vIntro = "NO" Then
            ' Guardamos el valor seleccionado en tu variable global por si la usas en otro sitio
            vDescripcion = CmbDescripcion.Text.Trim()
        End If
    End Sub

    Private Sub CmbDescripcion_GotFocus(sender As Object, e As EventArgs) Handles CmbDescripcion.GotFocus
        ' Si venimos del combo de conceptos, nos aseguramos de dejar guardado su valor real
        If vCombo = "concepto" Then
            TxtBuscarLetras.Enabled = False
            CmbConcepto.DroppedDown = False
            CmbConcepto.Text = vConcepto
        End If

        ' Preparamos el entorno para la descripción
        vIntro = "NO"
        vCombo = "descripcion"

        ' Configuramos el cuadro de búsqueda superior en minúsculas
        TxtBuscarLetras.CharacterCasing = CharacterCasing.Lower

        ' =====================================================================
        ' 🛠️ CONTROL INTELIGENTE DE BORRADO
        ' =====================================================================
        ' SI EL COMBO YA TIENE TEXTO (porque el concepto le ha metido la descripción por defecto),
        ' NO lo borramos. Solo seleccionamos el texto para que el usuario pueda escribir encima si quiere.
        If String.IsNullOrEmpty(CmbDescripcion.Text) Then
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.SelectedIndex = -1
            CmbDescripcion.Text = ""
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        Else
            ' Si ya tiene texto, colocamos el cursor al final de la palabra de forma limpia
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
        End If
    End Sub


    Private Sub TxtImporte_GotFocus(sender As Object, e As EventArgs) Handles TxtImporte.GotFocus
        TxtBuscarLetras.Text = ""
        TxtBuscarLetras.Enabled = False
        vCombo = ""
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()

        ' Llenar el Combo Concepto al cerrar
        '***********************************
        If CmbConcepto.Items.Count <> 0 Then
            CmbConcepto.Items.Clear()
        End If
        LlenarConcepto()
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            CmbCuenta.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptarOtro.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox(rmse.GetString("EjercicioActual"), MsgBoxStyle.Information, rmse.GetString("Fecha31Diciembre"))
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptarSalir_Click(sender As Object, e As EventArgs) Handles BtnAceptarSalir.Click
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        Else

        End If
        If TxtImporte.Text <> "0" Then
            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    vDate1 = Format(frmApuntesContables.DateTimePicker1.Value, "yyyy/MM/dd")
                    vDate2 = Format(frmApuntesContables.DateTimePicker2.Value, "yyyy/MM/dd")
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                vFilaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
                If vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1 Then
                    MsgBox(resManager.GetString("MsgFila2"))
                Else
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            Else
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If frmApuntesContables.BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                    vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                        End If
                    Else
                        vtipoSql += " Or "
                        If frmApuntesContables.BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                Next
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
            Me.Close()
        Else
            MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
            TxtImporte.Select()
        End If
    End Sub

    Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        Else

        End If
        If TxtImporte.Text <> "0" Then
            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    vDate1 = Format(frmApuntesContables.DateTimePicker1.Value, "yyyy/MM/dd")
                    vDate2 = Format(frmApuntesContables.DateTimePicker2.Value, "yyyy/MM/dd")
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
            Else
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If frmApuntesContables.BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                    vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                        End If
                    Else
                        vtipoSql += " Or "
                        If frmApuntesContables.BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                Next
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
            TxtImporte.Text = 0
            DateTimePicker1.Select()
        Else
            MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
    End Sub

    '' =========================================================================
    '' 1. BOTÓN: ACEPTAR Y SALIR
    '' =========================================================================
    'Private Sub BtnAceptarSalir_Click(sender As Object, e As EventArgs) Handles BtnAceptarSalir.Click
    '    ' Ejecuta todo el proceso de guardado y, si tiene éxito, cierra la ventana
    '    If GuardarApunteEnBaseDatos() Then
    '        Me.Close()
    '    End If
    'End Sub

    '' =========================================================================
    '' 2. BOTÓN: ACEPTAR Y OTRO
    '' =========================================================================
    'Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
    '    ' Ejecuta todo el proceso de guardado y, si tiene éxito, limpia para el siguiente apunte
    '    If GuardarApunteEnBaseDatos() Then
    '        vIntro = "NO"
    '        CmbConcepto.Text = ""
    '        CmbDescripcion.Text = ""
    '        TxtImporte.Text = 0

    '        ' Forzamos a vaciar los combos para el siguiente registro de forma segura
    '        RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
    '        CmbConcepto.Items.Clear()
    '        AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

    '        RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
    '        CmbDescripcion.Items.Clear()
    '        AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

    '        ' Volvemos a llenar los catálogos iniciales
    '        LlenarConcepto()
    '        LlenarDescripcion()

    '        ' Mandamos el foco de vuelta al inicio (la fecha)
    '        DateTimePicker1.Select()
    '    End If
    'End Sub

    'Private Function GuardarApunteEnBaseDatos() As Boolean
    '    ' Cerramos el lector de datos por seguridad si se quedó abierto
    '    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

    '    ' Captura y formateo de variables de pantalla
    '    vFechaAPU = Format(DateTimePicker1.Value, "dd/MM/yyyy")
    '    vConceptoAPU = Trim(CmbConcepto.Text)
    '    vDescripcionAPU = Trim(CmbDescripcion.Text)
    '    vNotasAPU = Trim(TxtNotas.Text)
    '    vCuentaAPU = Trim(CmbCuenta.Text)

    '    ' CORRECCIÓN DE DECIMALES Y SIGNOS
    '    vImporteAPU = Val(TxtImporte.Text.Replace(",", "."))
    '    If TxtTipoConcepto.Text = "GASTO" Then
    '        vImporteAPU = -Math.Abs(vImporteAPU)
    '    Else
    '        vImporteAPU = Math.Abs(vImporteAPU)
    '    End If

    '    ' Construcción de la sentencia SQL de inserción
    '    vAñadirSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) VALUES "
    '    vAñadirSql += "('" & vFechaAPU & "', '" & vConceptoAPU & "', '" & vDescripcionAPU & "', " & Str(vImporteAPU) & ", " & CInt(vAñoEjercicio) & ", '" & vNotasAPU & "', '" & vCuentaAPU & "')"

    '    cmdMdb1cr.CommandText = vAñadirSql

    '    Try
    '        ' 1. Ejecutamos la inserción en la base de datos
    '        cmdMdb1cr.ExecuteNonQuery()

    '        ' 2. ¡EL CAMBIO CLAVE!: Definimos la consulta de lectura
    '        vtipoSql = "SELECT * FROM apuntes WHERE apuntes.EjercicioAPU = " & CInt(vAñoEjercicio) & " ORDER BY apuntes.FechaAPU ASC"

    '        ' 3. Usamos tu método nativo centralizado para refrescar el grid de la pantalla de atrás
    '        ' Esto aplica automáticamente las traducciones, alineaciones y anchos proporcionales
    '        LlenarGrid(vtipoSql, frmApuntesContables.DgvApuntes, "1")

    '        ' Si todo ha ido bien, devolvemos TRUE
    '        Return True

    '    Catch ex As Exception
    '        ' Si algo falla, avisamos y devolvemos FALSE
    '        MsgBox("Error al guardar el apunte: " & ex.Message, MsgBoxStyle.Critical, "Error de Grabación")
    '        Return False
    '    End Try
    'End Function

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs) Handles BtnCuenta.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub CmbCuenta_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbCuenta.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            TxtNota.Select()
        End If
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub BtnDescripcion_Click(sender As Object, e As EventArgs) Handles BtnDescripcion.Click
        ' Llenar el Combo Descripción
        '****************************
        'If CmbDescripcion.Items.Count <> 0 Then
        '    MsgBox("El Combo Descripción ya esta Lleno.", vbExclamation, "Combo Descripción")
        '    CmbDescripcion.Items.Clear()
        'End If
        LlenarDescripcion()
    End Sub

    Public Function LlenarConcepto() As String
        ' 1. BLINDAJE: Vaciamos el combo de forma segura desactivando eventos
        RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
        CmbConcepto.SelectedIndex = -1
        CmbConcepto.Items.Clear()
        AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

        ' Forzamos solo la columna que necesitamos
        cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON ASC"

        Try
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            drMdb1 = cmdMdb1cr.ExecuteReader()

            If drMdb1.HasRows Then
                While drMdb1.Read()
                    Dim valor As String = Convert.ToString(drMdb1.GetValue(0))
                    If valor <> "TRASPASO" Then
                        CmbConcepto.Items.Add(valor)
                    End If
                End While

                ' 2. Asignación del concepto inicial por defecto
                If CmbConcepto.Items.Count > 0 Then
                    If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                        If frmApuntesContables.ListBox1.SelectedItems.Count <> 0 Then
                            CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
                        Else
                            Dim index As Integer = frmApuntesContables.CmbConcepto.SelectedIndex - 1
                            If index >= 0 AndAlso index < CmbConcepto.Items.Count Then
                                CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(index))
                            Else
                                CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
                            End If
                        End If
                    Else
                        CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
                    End If
                End If
            End If
            drMdb1.Close()
        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(rmse.GetString("ErrorLlenarDesplegable") & rmse.GetString("Label2.Text") & " " & ex.Message, vbExclamation, rmse.GetString("$this.Text"))
        End Try
        Return ""
    End Function

    Public Function LlenarDescripcion() As String
        ' 1. BLINDAJE: Vaciamos el combo de forma segura desactivando eventos para evitar errores de índice
        RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        CmbDescripcion.SelectedIndex = -1
        CmbDescripcion.Items.Clear()
        AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

        ' 2. OPTIMIZACIÓN CRÍTICA: Usamos SELECT DISTINCT para que la base de datos filtre los duplicados al vuelo
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'Saldo Inicial' ORDER BY DescripcionAPU ASC"

        Try
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            drMdb1 = cmdMdb1cr.ExecuteReader()

            If drMdb1.HasRows Then
                While drMdb1.Read()
                    Dim desc As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
                    If Not String.IsNullOrEmpty(desc) Then
                        CmbDescripcion.Items.Add(desc)
                    End If
                End While
            End If
            drMdb1.Close()
        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(rmse.GetString("ErrorLlenarDesplegable") & rmse.GetString("Label3.Text") & " " & ex.Message, vbExclamation, rmse.GetString("$this.Text"))
        End Try
        Return ""
    End Function

    Private Sub DateTimePicker1_LostFocus(sender As Object, e As EventArgs) Handles DateTimePicker1.LostFocus
        BtnCalculadora.TabIndex = 0
        BtnConcepto.TabIndex = 0
        BtnDescripcion.TabIndex = 0
        BtnHoy.TabIndex = 0
        BtnCuenta.TabIndex = 0
    End Sub

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub

    Private Sub CmbDescripcion_Click(sender As Object, e As EventArgs) Handles CmbDescripcion.Click
        CmbDescripcion.DroppedDown = True
    End Sub

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            CmbConcepto.Select()
            ' Opcional: Ejecutar una búsqueda o guardar valor
            'BtnHoy.PerformClick()
        End If
    End Sub
End Class