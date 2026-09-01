Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class AprendizajeBancario

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, strText, vIntro, vLetras, vCombo, vDescripcion As String
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(11) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())
    Dim textoAutocompletadoEnAzul As String = ""
    ' Variable global del formulario para memorizar qué fila exacta estamos procesando en la Pasarela
    Private vIdExtractoActual As Integer = 0
    Public vValorPrimero As Integer = 1
    Public vValorTotal As Integer = 0
    Dim traduciendoComoMaestro As Boolean = False

    ''' <summary>
    ''' Succiona la primera fila de tipo 'TEMPORAL' de la tabla extracto y rellena 
    ''' automáticamente las casillas de la interfaz para que el usuario decida si importar o saltar.
    ''' </summary>
    Public Sub CargarPrimerConceptoBancario()
        Try
            ' 1. INTERROGATORIO AL BÚNKER: Buscamos por CodigoAPU en tu conexion1 las filas 'TEMPORAL'
            Using cmdTop As New OleDb.OleDbCommand()
                cmdTop.Connection = conexion1 ' 🔌 Tu conexión real homologada
                cmdTop.CommandText = "SELECT TOP 1 CodigoAPU, FechaAPU, DescripcionAPU, ImporteAPU, CuentaAPU FROM extracto WHERE NotasAPU = 'TEMPORAL' ORDER BY CodigoAPU ASC"

                Using dr As OleDb.OleDbDataReader = cmdTop.ExecuteReader()
                    If dr.Read() Then
                        ' 🔑 Guardamos el ID único en la memoria de la RAM para saber luego a quién procesar o borrar
                        vIdExtractoActual = Convert.ToInt32(dr("CodigoAPU"))

                        ' 2. RELLENADO DE CASILLAS BIOLÓGICO (Alineado con tus controles reales)
                        ' Pasamos la fecha como texto formateado corto a tu txtFecha de toda la vida
                        DateTimePicker1.Value = Convert.ToDateTime(dr("FechaAPU")).Date

                        ' Estampamos el chorizo de texto crudo de Openbank/BBVA directo en la Descripción 1
                        TxtDescripcion.Text = dr("DescripcionAPU").ToString().Trim()

                        ' =========================================================================
                        ' 🪓 EL SERRUCHO INDUSTRIAL DE DECIMALES (INMUNE AL 100% A WINDOWS ALEMÁN)
                        ' =========================================================================
                        ' 1. Convertimos el valor de la base de datos a decimal puro
                        Dim valorDecimalPuro As Decimal = Convert.ToDecimal(dr("ImporteAPU"))

                        ' 2. Forzamos el formateo rígido con punto decimal de toda la vida: "466.67"
                        Dim textoPuroPunto As String = valorDecimalPuro.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture)

                        ' =========================================================================
                        ' 🎯 EL CORTAFUEGOS DECIMAL DE GALA PARA CASILLAS (INMUNE A MONITOR ALEMÁN)
                        ' =========================================================================
                        ' Forzamos al formateador visual a usar la cultura de puntos y comas española,
                        ' de esta manera el "466,67" se pintará perfecto con su coma decimal en cualquier PC del planeta.
                        Dim culturaVisual As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("es-ES")
                        TxtImporte.Text = Convert.ToDecimal(dr("ImporteAPU")).ToString("N2", culturaVisual)
                        ' =========================================================================

                        ' Clavamos el combo de la Cuenta indexando el ID que capturó el radar
                        Dim idCuentaBanco As Integer = Convert.ToInt32(dr("CuentaAPU"))
                        CmbCuenta.SelectedValue = idCuentaBanco

                        ' 🎯 ACCIÓN PREMIUM DE FOCO: Ponemos por defecto el combo de Conceptos en el ID 1 (Varios)
                        ' y clavamos el cursor parpadeando allí dentro para que el usuario solo tenga que teclear
                        CmbConcepto.SelectedValue = 1
                        CmbConcepto.Focus()

                    Else
                        ' 🎉 ¡EL TRIUNFO TOTAL DEL ASISTENTE ARTESANAL! Si ya no quedan más filas temporales, cerramos el taller
                        MsgBox(rmse.GetString("ExtractoCompletado"), MsgBoxStyle.Information, resManager.GetString("AppDisplayName"))

                        vIdExtractoActual = 0

                        ' En forma Modal, es asi
                        'Me.Close() ' Bajamos la persiana automáticamente

                        'En forma NO modal, es asi
                        Me.Dispose()
                    End If
                    ' Concatenamos la traducción de "Label3.Text" con el avance: ej "Apunte 1/40"
                    Label3.Text = rmse.GetString("Label3.Text") & " " & vValorPrimero.ToString() & "/" & vValorTotal.ToString()

                End Using
            End Using

        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorEnExtracto") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub AprendizajeBancario_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' 🔒 DECAPITAMOS EL CAMINO DEL TABULADOR EN LAS CASILLAS DE SOLO LECTURA
        ' El cursor jamás se detendrá aquí dentro al pulsar el Tabulador
        TxtImporte.TabStop = False
        BtnConcepto.TabStop = False
        BtnDescripcion.TabStop = False
        CmbCuenta.TabStop = False ' Si el banco tampoco se toca, lo capamos también

        ' 🎯 EL FOCO DE ENTRADA: Obligamos al cursor a nacer parpadeando en el combo limpio
        CmbConcepto.TabStop = True
        CmbConcepto.Focus()


        Me.KeyPreview = True

        Label7.Text = vMoneda
        Label9.Text = vMoneda
        TxtSaldoFinal.Text = vSaldoFinal.ToString("N2")


        vIntro = "NO"
        ' 1. Convertimos el año base de forma segura a número entero
        Dim anio As Integer
        If Not Integer.TryParse(vAñoEjercicio, anio) Then
            ' Salvavidas: si falla o está vacío, usa el año actual
            anio = Date.Today.Year
        End If

        ' 2. Asignamos el año a tus variables por si las usas más adelante
        vFecha1Enero = anio
        vFecha31Diciembre = anio

        ' 3. Creamos las fechas límite de forma limpia
        Dim fechaInicio As New Date(anio, 1, 1)
        Dim fechaFin As New Date(anio, 12, 31)

        ' 4. Aplicamos los rangos al control
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker1.MaxDate = fechaFin

        ' 5. Aplicamos la lógica de asignación del valor según el año
        ' Comparamos de forma segura convirtiendo vAñoActual a número o texto de forma explícita
        If anio.ToString() <> vAñoActual.ToString() Then
            ' Si el año no es el actual, se inicializa en el último día de ese año contable
            DateTimePicker1.Value = fechaFin
        Else
            ' Si coincide con el año en curso, se inicializa con la fecha de hoy
            ' Nota: Asegúrate de que vfechaHoy contenga un objeto Date válido o usa Date.Today
            DateTimePicker1.Value = Convert.ToDateTime(vfechaHoy)
        End If

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptarOtro, rmse.GetString("BtnAceptarOtro.Text"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSaltarApuntes, rmse.GetString("BtnSaltarApuntes.Text"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnCancelar, rmse.GetString("BtnCancelar.Text") & " " & rmse.GetString("$this.Text"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.CmbConcepto, rmse.GetString("SelecConcepto"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbCuenta, rmse.GetString("SelecCuenta"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbDescripcion, rmse.GetString("SelecDescripcion"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteAsiento"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnDescripcion, rmse.GetString("BtnDescripcion"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.TxtBuscarLetras, rmse.GetString("TxtABuscar"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnSaltarApuntes, rmse.GetString("BtnSaltarApuntes.Text"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.TxtDescripcion, My.Settings.RutaBD)
        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox1.MouseMove, AddressOf VerificarFiltrosDesactivados

        CmbConcepto.DropDownStyle = ComboBoxStyle.DropDownList

        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()

        ' Llenar el Combo Concepto de forma segura y traducida (IntroApuntes)
        '******************************************************************
        Try
            ' 1. Encendemos tu escudo protector antes de rellenar los componentes
            cargandoFormulario = True

            ' 🌟 CABLE A: Cargamos el ComboBox de forma independiente ordenado de la A a la Z puro
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 🌟 CABLE B: Cargamos el combo de cuentas genérico de tu módulo (Si lo tienes en esta pantalla)
            LlenarComboCuentasGenerico(Me.CmbCuenta)

            ' 2. Apagamos el escudo tras la inyección exitosa en la memoria RAM
            cargandoFormulario = False

            ' =========================================================================
            ' 🌟 SINCRONIZACIÓN INTELIGENTE DE CONCEPTOS DE ATRÁS CONTRA IDs
            ' =========================================================================
            ' Verificamos si la pantalla de extractos de atrás existe en la RAM para heredar su filtro
            If frmApuntesContables IsNot Nothing AndAlso frmApuntesContables.IsHandleCreated Then

                ' Si en la pantalla principal el usuario ya tenía filtrado un concepto, 
                ' lo pre-seleccionamos automáticamente en esta ventana usando su ID numérico entero
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    CmbConcepto.SelectedValue = frmApuntesContables.CmbConcepto.SelectedValue
                Else
                    If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
                End If

            Else
                ' Plan B (Carga Aislada): Si entramos directo, marcamos el primer concepto de la lista
                If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorCargarCONyCUE") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
        End Try


        ' --- ÚLTIMAS LÍNEAS DE TU INTROAPUNTES_LOAD ---
        'TxtImporte.Text = "0"

        ' 🌟 LA CORRECCIÓN CLAVE: Apagamos el escudo protector AQUÍ, 
        ' justo antes de forzar la selección para que el evento deje pasar los datos
        cargandoFormulario = False

        ' SELECCIÓN PRIMER ELEMENTO: Forzamos el índice 0 (ej: ADESLAS)
        ' 🌟 TRUCO MAESTRO DE REINICIO DE ÍNDICE:
        ' Forzamos un vaivén de selección para obligar al evento a dispararse sí o sí
        If CmbConcepto.Items.Count > 1 Then
            CmbConcepto.SelectedIndex = -1 ' Lo bajamos a vacío primero
            CmbConcepto.SelectedIndex = 0  ' Lo subimos a la posición 1 para que rellene las descripciones
        ElseIf CmbConcepto.Items.Count > 0 Then
            CmbConcepto.SelectedIndex = -1
            CmbConcepto.SelectedIndex = 0
        End If

        ' =========================================================================
        ' 🎯 LA HERENCIA MAESTRA: Sincronizamos los combos con el filtro de fondo
        ' =========================================================================
        Try
            ' 1. Sincronizar el combo de Cuenta (Si en la pantalla de atrás hay filtro activo)
            If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaFiltro As Integer = Convert.ToInt32(frmApuntesContables.CmbCuenta.SelectedValue)
                If idCuentaFiltro > 0 Then
                    CmbCuenta.SelectedValue = idCuentaFiltro
                    ' Disparamos el SelectedIndexChanged manual por software si fuera necesario
                End If
            End If

            ' 2. Sincronizar el combo de Concepto (Si en la pantalla de atrás hay filtro activo)
            If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoFiltro As Integer = Convert.ToInt32(frmApuntesContables.CmbConcepto.SelectedValue)
                If idConceptoFiltro > 0 Then
                    CmbConcepto.SelectedValue = idConceptoFiltro

                    ' 🚀 ACTUALIZACIÓN AUTOMÁTICA EN LA RAM DE LA INTERFAZ
                    ' Forzamos a que pinte el tipo (Gasto/Ingreso) y la descripción por defecto
                    CmbConcepto_SelectedIndexChanged(CmbConcepto, EventArgs.Empty)
                End If
            End If
        Catch ex As Exception
            ' Cortafuegos silencioso de seguridad para el monitor
        End Try

    End Sub

    Private Sub TxtBuscarLetras_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBuscarLetras.KeyDown
        ' Si el usuario pulsa FLECHA ABAJO desde el cuadro de búsqueda, saltamos al combo de forma inteligente
        If e.KeyCode = Keys.Down Then
            If vCombo = "descripcion" AndAlso CmbDescripcion.Items.Count > 0 Then
                e.Handled = True
                CmbDescripcion.Focus()
                CmbDescripcion.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtBuscarLetras.KeyPress
        ' Detectamos si se pulsa ENTER (Ascii 13)
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True ' Evita el molesto pitido de Windows

            ' 🛠️ CORRECCIÓN DE ANIDACIÓN: Cambiado a un IF limpio y directo para corregir el teclado
            If vCombo = "descripcion_vacia" Then
                ' Bloque de alta de descripción nueva
                Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("NoExistenDescripciones") & ": -" & TxtBuscarLetras.Text.ToUpper() & "-" & vbCrLf & "¿" & rmse.GetString("AñadirDescripcion") & "?", rmse.GetString("$this.Text"))
                If respuesta = vbYes Then
                    vIntro = "SI"
                    vDescripcion = TxtBuscarLetras.Text

                    ' 🛡️ EL ESCUDO DE HANDLERS TOTAL: Desenchufamos los DOS cables a la vez para limpiar la RAM
                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

                    ' Forzamos los textos en frío sin disparar bucles fantasma en segundo plano
                    CmbDescripcion.Text = vDescripcion
                    TxtBuscarLetras.Text = ""
                    TxtBuscarLetras.Enabled = False
                    vLetras = ""

                    ' 🔌 VOLVEMOS A CONECTAR LOS DOS CABLES (Mesa de operaciones libre de corriente)
                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

                    ' 🎯 LA ESTOCADA DEL FOCO: Forzamos a Windows a procesar los gráficos y clavamos el cursor
                    Application.DoEvents()
                    CmbDescripcion.Enabled = True
                    CmbDescripcion.Focus()
                    CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
                    CmbDescripcion.SelectionLength = 0

                    ' 🔓 VOLVEMOS A DESPERTAR EL TEXTBOX ARRIBA (Pero el cursor ya se ha quedado anclado abajo)
                    Application.DoEvents()
                    TxtBuscarLetras.Enabled = True
                Else
                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    If TxtBuscarLetras.Text.Length > 0 Then
                        TxtBuscarLetras.Text = TxtBuscarLetras.Text.Substring(0, TxtBuscarLetras.Text.Length - 1)
                    End If
                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

                    TxtBuscarLetras.Focus()
                    TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
                    vLetras = TxtBuscarLetras.Text
                    vCombo = "descripcion"
                    BuscarLetras(vCombo)
                End If
            Else
                ' 🛠️ CORRECCIÓN: Si vCombo es "descripcion" normal (sí hay registros en el desplegable)
                ' Eliminamos el MsgBox molesto para que al pulsar Enter baje directamente al combo,
                ' despliegue las opciones y seleccione el primer elemento para moverte con las flechas.
                If CmbDescripcion.Items.Count > 0 Then
                    CmbDescripcion.Focus()
                    CmbDescripcion.DroppedDown = True
                    CmbDescripcion.SelectedIndex = 0
                End If
            End If
        Else
            ' Si es cualquier otra letra (Mayúscula, minúscula, espacio, etc.), 
            ' indicamos de forma limpia que estamos en proceso de escritura ordinaria
            vIntro = "NO"
        End If
    End Sub

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        ' Sincronizamos la variable global con lo que el usuario escribe o borra arriba
        vLetras = TxtBuscarLetras.Text

        ' Evitamos ejecutar la base de datos si el cuadro superior se limpia de forma automática
        If String.IsNullOrEmpty(vLetras) Then
            ' Desenchufamos el combo temporalmente para que no salte su SelectedIndexChanged al vaciarlo
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            CmbDescripcion.SelectedIndex = -1
            CmbDescripcion.Text = ""
            If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

            ' Volvemos a enchufar el cable del combo dócilmente
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            Exit Sub
        End If

        ' Forzamos a evaluar el bloque de consulta SQL de descripción siempre
        BuscarLetras("descripcion")
    End Sub
    Private Sub CmbDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbDescripcion.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ' 1. Guardamos la descripción seleccionada con las flechas
            Dim textoSeleccionado As String = CmbDescripcion.Text

            ' 2. Apagamos el buscador de arriba de forma controlada
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            'TxtBuscarLetras.Enabled = False

            ' 3. Aseguramos el texto en el combo y en tu variable global
            CmbDescripcion.Text = textoSeleccionado
            vDescripcion = textoSeleccionado

            ' 4. Cerramos el desplegable y mandamos el cursor al Importe
            If CmbDescripcion.DroppedDown Then CmbDescripcion.DroppedDown = False

            TxtNota.Select()
        End If
    End Sub

    Private Sub CmbDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbDescripcion.KeyPress
        ' Si es un Intro, salimos inmediatamente
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Exit Sub
        End If

        ' 🛠️ BLINDAJE ABSOLUTO MANTENIDO: Si vIntro es "SI", el usuario edita libremente
        If vIntro = "SI" Then
            ' Sincronizamos la variable de forma segura reteniendo mayúsculas/minúsculas nativas
            BeginInvoke(Sub() vDescripcion = CmbDescripcion.Text)
            Exit Sub
        End If

        ' 🛠️ NUEVA MEJORA EN CALIENTE: Si el usuario borra todo el texto a mano, 
        ' liberamos el combo para poder escribir una descripción nueva desde cero.
        If String.IsNullOrEmpty(CmbDescripcion.Text) Then
            vIntro = "SI"
            vDescripcion = ""
            Exit Sub
        End If

        Dim letra As Char = e.KeyChar

        ' Flujo por defecto si sólo está buscando palabras existentes
        If Char.IsLetterOrDigit(letra) OrElse letra = " "c Then
            vCombo = "descripcion"

            TxtBuscarLetras.Enabled = True
            TxtBuscarLetras.Text += letra.ToString()

            TxtBuscarLetras.Focus()
            TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
            TxtBuscarLetras.SelectionLength = 0

            vLetras = TxtBuscarLetras.Text

            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            ' =========================================================================
            ' 🎯 EL ESCUDO DE ACERTIJO: Vaciamos la lista de forma dócil sin provocar al motor Win32
            ' =========================================================================
            Try
                CmbDescripcion.DroppedDown = False
                CmbDescripcion.DataSource = Nothing ' Desvinculamos la caché relacional
                CmbDescripcion.SelectedIndex = -1
            Catch ex As Exception
                ' Cortafuegos silencioso
            End Try
            ' =========================================================================

            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.Text = ""
            e.Handled = True
        End If
    End Sub

    Private Sub CmbDescripcion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDescripcion.SelectedIndexChanged
        If vIntro = "NO" Then
            ' Guardamos el valor seleccionado en tu variable global por si la usas en otro sitio
            vDescripcion = CmbDescripcion.Text.Trim()
        End If
    End Sub

    Private Sub CmbDescripcion_GotFocus(sender As Object, e As EventArgs) Handles CmbDescripcion.GotFocus

        ' 🎯 EL ESCUDO INTERNACIONAL: Si estamos traduciendo el formulario, 
        ' salimos en paz sin dejar que WinForms altere el texto alemán.
        If traduciendoComoMaestro Then Exit Sub

        ' 🛡️ CORTAFUEGOS DE ALTA NUEVA (Tu escudo definitivo contra el robo del cursor)
        ' Si venimos de pulsar "SÍ" para añadir una descripción, saltamos la inicialización
        ' para evitar que los refrescos automáticos de Windows le devuelvan el foco al TextBox superior.
        If vIntro = "SI" Then
            vCombo = "descripcion"
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
            Exit Sub ' 🪓 Cortamos el hilo aquí y dejamos el cursor parpadeando inmóvil en el combo!
        End If

        ' Si venimos del combo de conceptos, nos aseguramos de dejar guardado su valor real
        If vCombo = "concepto" Then
            CmbConcepto.DroppedDown = False
            CmbConcepto.Text = vConcepto
        End If

        ' Preparamos el entorno para la descripción
        ' 🛠️ AJUSTE: Solo ponemos vIntro en "NO" si no estábamos ya editando una descripción nueva ("SI")
        'If vIntro <> "SI" Then vIntro = "NO"
        vIntro = "NO"
        vCombo = "descripcion"

        ' 🛠️ CORRECCIÓN ABSOLUTA: Cambiamos Lower por Normal para liberar las Mayúsculas/Minúsculas
        TxtBuscarLetras.CharacterCasing = CharacterCasing.Normal

        ' =====================================================================
        ' 🛠️ CONTROL INTELIGENTE DE BORRADO (MANTENIDO)
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

        ' 🛠️ MEJORA DE NAVEGACIÓN: Si estamos en modo de alta nueva, controlamos el cursor parpadeante
        If vIntro = "SI" Then
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
        End If
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        ' 1. Abrimos la pantalla de mantenimiento de conceptos del formulario principal
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()

        ' =========================================================================
        ' 🌟 RECARGA DE LA NUEVA ERA: CERO BUCLES WHILE Y 100% SEGURO CON IDs
        ' =========================================================================
        ' Encendemos el escudo protector para que los eventos de cambio no se vuelvan locos al recargar
        cargandoFormulario = True

        Try
            ' 2. Llamamos a nuestra rutina exclusiva que limpia, filtra especiales, 
            ' traduce e inyecta el DataTable con IDs numéricos en un milisegundo
            LlenarComboConceptosSueltosBD(Me.CmbConcepto) ' Llamada adicional para asegurar la sincronización con la base de datos

            ' 3. Apagamos el escudo protector para permitir la interacción del usuario
            cargandoFormulario = False

            ' 4. Volvemos a aplicar tu vaivén maestro de índices para forzar el relleno de descripciones
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1 ' Reseteamos a vacío primero
                CmbConcepto.SelectedIndex = 0  ' Seleccionamos el primer elemento de forma segura
            End If

        Catch ex As Exception
            cargandoFormulario = False
            MsgBox(resManager.GetString("ErrorRefrescarCON") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptarOtro.Select()
        End If
    End Sub

    Private Sub BtnSaltarApuntes_Click(sender As Object, e As EventArgs) Handles BtnSaltarApuntes.Click
        If vIdExtractoActual = 0 Then Exit Sub

        Try
            ' 🪓 Limpiamos el cromo de la pasarela temporal extracto para avanzar
            Using cmdDel As New OleDb.OleDbCommand("DELETE * FROM extracto WHERE CodigoAPU = " & vIdExtractoActual, conexion1)
                cmdDel.ExecuteNonQuery()
            End Using

            ' Succionamos la siguiente línea del banco dócilmente
            ' 🛗 EL ASCENSOR: Sumamos un cromo más al marcador visual antes de succionar la siguiente fila
            vValorPrimero += 1

            CargarPrimerConceptoBancario()

        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorSaltarApunte") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
        ' Cortafuegos preventivo
        If vIdExtractoActual = 0 Then Exit Sub

        Try
            ' =========================================================================
            ' 🎯 RETENCIÓN DE MEMORIA PREMIUM (Tu invento de IntroApuntes)
            ' =========================================================================
            ' Capturamos el ID del concepto seleccionado por el usuario ANTES de mutar la RAM
            Dim idConceptoAMemorizar As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            ' =========================================================================

            ' 1. PESCA DE VARIABLES PURAS DESDE TU INTERFAZ ELÁSTICA
            Dim fechaSQL As String = "#" & DateTimePicker1.Value.ToString("yyyy/MM/dd") & "#"
            Dim importeSQL As String = Convert.ToDecimal(TxtImporte.Text).ToString(System.Globalization.CultureInfo.InvariantCulture)
            Dim idConcepto As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            Dim idCuenta As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

            ' Guardamos en tus apuntes oficiales el texto pulido que tú hayas retocado o elegido en la Descripció 2
            Dim descripcionUsuario As String = CmbDescripcion.Text.Replace("'", "''").Trim()
            Dim notaSQL As String = TxtNota.Text.Replace("'", "''").Trim()

            ' 2. 📁 INYECCIÓN DIRECTA EN TU TABLA REAL DE APUNTES
            Using cmdIns As New OleDb.OleDbCommand()
                cmdIns.Connection = conexion1
                cmdIns.CommandText = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, CuentaAPU, NotasAPU) " &
                                     "VALUES (" & fechaSQL & ", " & idConcepto & ", '" & descripcionUsuario & "', " & importeSQL & ", " & vAñoEjercicio & ", " & idCuenta & ", '" & notaSQL & "')"
                cmdIns.ExecuteNonQuery()
            End Using

            ' 3. 🪓 ELIMINACIÓN DE LA PASARELA: Borramos este cromo de la tabla temporal extracto
            Using cmdDel As New OleDb.OleDbCommand("DELETE * FROM extracto WHERE CodigoAPU = " & vIdExtractoActual, conexion1)
                cmdDel.ExecuteNonQuery()
            End Using

            ' 4. 🚀 SIGUIENTE FILA: Succionamos el próximo apunte del banco en millonésimas de segundo
            ' 🛗 EL ASCENSOR: Sumamos un cromo más al marcador visual antes de succionar la siguiente fila
            vValorPrimero += 1

            CargarPrimerConceptoBancario()

            ' =========================================================================
            ' 🪓 EL TIRO DE GRACIA: Obligamos al Combo a recuperar el valor memorizado
            ' =========================================================================
            ' En lugar de resetearse al ID 1 (Varios), el combo se clava en tu última decisión
            If vIdExtractoActual > 0 Then
                CmbConcepto.SelectedValue = idConceptoAMemorizar
                CmbConcepto.Focus() ' Clavamos el cursor allí de nuevo para el Tabulador
            End If

        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorGuardarApunte") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' 🌟 SANEAMIENTO PREVENTIVO DE PARÁMETROS PARA EL REFRESCO
        cmdMdb1cr.Parameters.Clear()

        ' Consulta SQL Maestra de 11 celdas relacionales (Tu diseño perfecto)
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &
                           "conceptos.DescripcionCON As [ConceptoAPU], " &
                           "apuntes.DescripcionAPU As [DescripcionAPU], " &
                           "apuntes.ImporteAPU As [ImporteAPU], " &
                           "apuntes.ImporteAPU As [SaldoAPU], " &
                           "apuntes.NotasAPU As [NotasAPU], " &
                           "cuentas.NombreCUE As [CuentaAPU], " &
                           "apuntes.CodigoAPU As [CodigoAPU], " &
                           "conceptos.CodigoCON As [CodigoCON], " &
                           "apuntes.ConceptoAPU As [IdConceptoCON], " &
                           "apuntes.CuentaAPU As [IdCuentaCUE] " &
                           "FROM (apuntes " &
                           "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
                           "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString

        ' 🌟 CORRECCIÓN 1: Filtro por ID numérico de Cuenta (Sin comillas simples)
        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
            Dim idCuentaPrincipal As Integer = Convert.ToInt32(frmApuntesContables.CmbCuenta.SelectedValue)
            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaPrincipal} "
        End If

        ' 🌟 CORRECCIÓN 2: Filtro por ID numérico de Concepto (Sin comillas simples)
        If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
            Dim idConceptoPrincipal As Integer = Convert.ToInt32(frmApuntesContables.CmbConcepto.SelectedValue)
            vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoPrincipal} "
        End If

        ' 🌟 CORRECCIÓN 3: Sincronización estricta de parámetros de fechas
        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date

            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' Inyectamos los valores en el comando global en el orden de los signos '?'
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
        End If

        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(frmApuntesContables.DgvApuntes)

        vFilaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        If vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = frmApuntesContables.DgvApuntes.RowCount - 1
            frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
            frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
        End If


        ' =====================================================================
        ' 🛠️ REINICIO ASISTIDO DE PASOS PARA INTRODUCIR OTRO APUNTE
        ' =====================================================================
        ' 1. Apagamos el TextChanged superior para evitar llamadas falsas a la BD al limpiar
        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        TxtBuscarLetras.Text = ""
        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

        ' 2. Reestablecemos el estado inicial del buscador de descripciones
        TxtBuscarLetras.Enabled = True
        vLetras = ""
        vIntro = "NO"

        ' 1. Limpias los textos normales
        TxtNota.Text = ""

        ' 3. ¡La clave! Forzamos al formulario a procesar los cambios visuales antes de seguir
        Application.DoEvents()

        ' 4. Regresamos el cursor a la Fecha para arrancar de nuevo el flujo de introducción
        DateTimePicker1.Focus()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub BtnDescripcion_Click(sender As Object, e As EventArgs) Handles BtnDescripcion.Click
        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()
    End Sub

    Private Sub CmbDescripcion_Click(sender As Object, e As EventArgs) Handles CmbDescripcion.Click
        CmbDescripcion.DroppedDown = True
    End Sub

    Public Sub LlenarDescripcion()
        ' 1. SANEAMIENTO PREVENTIVO: Apagamos los eventos un instante para evitar disparos en falso
        RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        CmbDescripcion.DataSource = Nothing
        CmbDescripcion.Items.Clear()

        ' 2. LA SQL PERFECTA (Tu excelente consulta DISTINCT que delega el trabajo en Access)
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'Saldo Inicial' ORDER BY DescripcionAPU ASC"
        cmdMdb1cr.Parameters.Clear()

        Dim dtDescripciones As New DataTable()

        Try
            ' 🌟 EL ATAJO DE .NET: Cargamos los datos directos en el DataTable (Abre, lee y CIERRA el flujo al vuelo)
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                dtDescripciones.Load(dr)
            End Using

            ' 3. VINCULACIÓN DIRECTA EN UN MILISEGUNDO
            CmbDescripcion.ValueMember = "DescripcionAPU"
            CmbDescripcion.DisplayMember = "DescripcionAPU"
            CmbDescripcion.DataSource = dtDescripciones ' Al asignarlo de golpe, .NET dibuja el combo una sola vez

            CmbDescripcion.SelectedIndex = -1 ' Lo dejamos inicialmente vacío y limpio

        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorLlenarDesplegable") & " " & ex.Message, vbExclamation, rmse.GetString("$this.Text"))
        Finally
            ' Restauramos el escuchador de eventos de forma totalmente segura
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        End Try
    End Sub

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub

    Private Sub CmbConcepto_TextChanged(sender As Object, e As EventArgs) Handles CmbConcepto.TextChanged
        textoAutocompletadoEnAzul = CmbConcepto.Text.Trim()
    End Sub

    Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' 🎯 CORTAFUEGOS TOTAL: Bloqueamos el rebote elástico nativo de Windows hacia el "ICO..."
            e.SuppressKeyPress = True
            e.Handled = True

            ' 🚀 LA ESTOCADA ASÍNCRONA SIMÉTRICA: 
            ' Le damos un milisegundo de tregua a la interfaz para que asimile el cierre de la tecla
            ' y clame el cursor dentro del buscador de forma indestructible use la tecla que use.
            BeginInvoke(Sub()
                            Try
                                TxtBuscarLetras.Focus()
                                TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
                            Catch
                                ' Cortafuegos silencioso
                            End Try
                        End Sub)

            Dim idConceptoSel As Integer = 0
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 TU JUGADA MAESTRA: Si la variable guardó el texto en azul, buscamos esa palabra exacta
            Dim textoBuscar As String = textoAutocompletadoEnAzul
            If String.IsNullOrEmpty(textoBuscar) Then textoBuscar = CmbConcepto.Text.Trim()
            If Not String.IsNullOrEmpty(textoBuscar) Then
                Try
                    'MsgBox(rmse.GetString("ConceptoSeleccionado") & ": " & textoBuscar)

                    Dim dt As DataTable = CType(CmbConcepto.DataSource, DataTable)
                    If dt IsNot Nothing Then
                        ' Buscamos en la caché de la RAM la fila que coincide exactamente con el chivato
                        Dim filas() As DataRow = dt.Select("TextoCombo = '" & textoBuscar.Replace("'", "''") & "'")

                        If filas.Length = 0 Then
                            filas = dt.Select("TextoCombo LIKE '" & textoBuscar.Replace("'", "''") & "%'")
                        End If

                        If filas.Length > 0 Then
                            idConceptoSel = Convert.ToInt32(filas(0)("IdConceptoCON"))
                            codigoOriginal = filas(0)("CodigoCON").ToString().Trim()
                            descripcionOriginal = filas(0)("DescripcionCON").ToString().Trim()
                            If dt.Columns.Contains("TipoCON") Then tipoOriginal = filas(0)("TipoCON").ToString().Trim()

                            ' Forzamos al combo a quedarse rígido en la posición física correcta de la fila
                            CmbConcepto.SelectedIndex = dt.Rows.IndexOf(filas(0))
                        End If
                    End If
                Catch ex As Exception
                    ' Silencioso
                End Try
            End If

            ' Sincronizamos las variables globales con el ID numérico real hallado
            vConcepto = idConceptoSel.ToString()

            ' Apagamos el buscador de arriba para que al vaciarlo no active consultas
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

            TxtBuscarLetras.Enabled = True
            vCombo = "descripcion"
            If idConceptoSel > 0 Then CmbConcepto.SelectedValue = idConceptoSel

            ' =====================================================================
            ' 🌟 INYECCIÓN DIRECTA EN LA INTERFAZ DESDE LA MEMORIA CACHÉ (VERSION 3.2.9)
            ' =====================================================================
            If idConceptoSel > 0 Then
                Dim tradTipo As String = ""
                Select Case tipoOriginal.ToUpper()
                    Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                    Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                    Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                End Select
                If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
                TxtTipoConcepto.Text = tradTipo

                ' 🛡️ RADAR LIMPIADOR: Homogeneizamos el texto para evitar que las tildes rompan la búsqueda
                Dim codigoLimpio As String = codigoOriginal.ToUpper().Trim()
                codigoLimpio = codigoLimpio.Replace("É", "E").Replace("È", "E")
                codigoLimpio = codigoLimpio.Replace("Á", "A").Replace("À", "A")
                codigoLimpio = codigoLimpio.Replace("Í", "I").Replace("Ó", "O").Replace("Ú", "U")
                codigoLimpio = codigoLimpio.Replace(" ", "_")

                ' Construimos la llave limpia (Ej: "Desc_ESTETICA")
                Dim llaveDesc As String = "Desc_" & codigoLimpio
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Salvavidas específico: Si por la estructura de la frase se sigue resistiendo,
                ' forzamos el mapeo de tu clave maestra del ResXManager para la captura visual
                If codigoLimpio.Contains("ESTETICA") Then
                    Dim tradRescate As String = resManager.GetString("Desc_ESTETICA")
                    If Not String.IsNullOrEmpty(tradRescate) Then tradDesc = tradRescate
                End If

                ' Si no tiene traducción en el ResX, dejamos la descripción original de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                ' 1. Pintamos la descripción 1 perfecta
                If TypeOf TxtDescripcion Is TextBox Then TxtDescripcion.Text = tradDesc

                ' 🎯 2. ENCENDEMOS EL ESCUDO DE LA VIEJA ESCUELA
                ' Le prohibimos al GotFocus y a los motores de Windows Forms alterar el texto del combo 2
                traduciendoComoMaestro = True

                ' 3. Tu imbatible igualdad directa de la vieja escuela (Ahora con el texto alemán real)
                CmbDescripcion.Text = tradDesc
                vDescripcion = tradDesc
            End If
            ' =====================================================================

            ' Limpiamos el chivato para la próxima búsqueda, cerramos la lista y saltamos
            textoAutocompletadoEnAzul = ""
            If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False

            ' Mandamos el foco (Disparará el GotFocus, pero chocará con nuestro escudo y no romperá nada)
            CmbDescripcion.Select()

            ' Fijamos el texto del concepto
            CmbConcepto.Text = textoBuscar

            ' 🎯 4. APAGAMOS EL ESCUDO
            ' Una vez que toda la interfaz se ha asentado y dibujado en el monitor, liberamos el control
            traduciendoComoMaestro = False
        End If
    End Sub


    Private Sub CmbConcepto_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbConcepto.MouseClick
        'TxtBuscarLetras.Enabled = False
        vIntro = "NO"
        ' Solo forzamos el despliegue automático si el usuario NO ha pulsado la flecha nativa
        ' (Nos aseguramos comprobando si la lista ya está abierta o abriéndola suavemente)
        If CmbConcepto.Items.Count <> 0 AndAlso Not CmbConcepto.DroppedDown Then
            CmbConcepto.DroppedDown = True
            'CmbConcepto.SelectedIndex = 0
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 1. ESCUDO DE CARGA: Si el formulario se está iniciando o el combo está vacío, salimos inmediatamente
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        ' Se buscan Conceptos según lo seleccionado para mostrar su descripción y tipo en los cuadros de abajo
        '*****************************************************************************************************
        If vIntro = "NO" Then
            TxtBuscarLetras.Text = ""
            Try
                Dim codigoOriginal As String = ""
                Dim descripcionOriginal As String = ""
                Dim tipoOriginal As String = ""

                ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader)
                ' Como el combo está enlazado a un DataTable, convertimos el ítem actual en un DataRowView
                If CmbConcepto.SelectedItem IsNot Nothing Then
                    Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                    codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                    textoAutocompletadoEnAzul = codigoOriginal
                    descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()
                    ' Leemos el TipoCON de forma segura por si acaso
                    If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                        tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                    End If

                    ' =========================================================================
                    ' 🎯 SINCRONIZACIÓN ASÍNCRONA PREMIUM 3.2.6 (Inmune a DropDownList)
                    ' =========================================================================
                    ' Le damos un microsegundo de tregua a la CPU para que el motor cargue las descripciones
                    ' antes de forzar la selección visual en la pantalla.
                    Dim copiaDescripcion As String = descripcionOriginal
                    BeginInvoke(Sub()
                                    Try
                                        ' 1. Intentamos la vía dócil asignando el texto
                                        CmbDescripcion.Text = copiaDescripcion

                                        ' 2. 🛡️ EL SALVAVIDAS DE REDMOND: Si se quedó sordo por el DropDownList,
                                        ' obligamos al motor Win32 a buscar el texto exacto en su colección
                                        If CmbDescripcion.SelectedIndex = -1 Then
                                            CmbDescripcion.SelectedIndex = CmbDescripcion.FindStringExact(copiaDescripcion)
                                        End If
                                    Catch
                                        ' Cortafuegos silencioso
                                    End Try
                                End Sub)
                End If

                ' 3. Traducir y asignar los textos a la interfaz de forma segura
                If Not String.IsNullOrEmpty(codigoOriginal) Then
                    vConcepto = codigoOriginal ' Guardamos el código original en español para la BD

                    ' --- TRADUCIR EL TIPO (Gasto / Ingreso / Especial) ---
                    Dim tradTipo As String = ""
                    Select Case tipoOriginal.ToUpper()
                        Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                        Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                        Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                    End Select
                    If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
                    TxtTipoConcepto.Text = tradTipo

                    ' =========================================================================
                    ' 🎯 SINCRONIZACIÓN ASÍNCRONA DIRECTA (Inmune a problemas de refresco)
                    ' =========================================================================
                    ' 1. Construimos la clave uniendo el código que ya viene limpio (Ej: "Desc_ESTETICA")
                    Dim llaveDesc As String = "Desc_" & codigoOriginal.ToUpper().Trim()
                    Dim tradDesc As String = resManager.GetString(llaveDesc)

                    ' Si no tiene traducción en el ResX, dejamos la descripción original de la BD
                    If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                    ' La descripción 1 se pinta perfecta al instante
                    TxtDescripcion.Text = tradDesc

                    ' 2. LA TREGUA ASÍNCRONA: Le damos al formulario el mismo tiempo de respiro 
                    ' que le daba tu MsgBox, pero de forma invisible y elegante para el usuario.
                    Dim copiaTradDesc As String = tradDesc

                    BeginInvoke(Sub()
                                    Try
                                        ' Encendemos tu escudo protector de eventos
                                        traduciendoComoMaestro = True

                                        ' Vaciamos y rellenamos el combo con el texto traducido final
                                        CmbDescripcion.DataSource = Nothing
                                        CmbDescripcion.Items.Clear()
                                        CmbDescripcion.Items.Add(copiaTradDesc)
                                        CmbDescripcion.SelectedIndex = 0

                                        ' Tu imbatible igualdad de la vieja escuela
                                        CmbDescripcion.Text = TxtDescripcion.Text

                                        ' Forzamos el repintado gráfico en la pantalla
                                        CmbDescripcion.Refresh()

                                        ' Apagamos el escudo de forma segura
                                        traduciendoComoMaestro = False
                                    Catch
                                        ' Cortafuegos silencioso
                                    End Try
                                End Sub)
                End If
            Catch ex As Exception
                MsgBox(resManager.GetString("ErrorSincronizarCON") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
            End Try
        End If
    End Sub

    Private Sub CmbDescripcion_Enter(sender As Object, e As EventArgs) Handles CmbDescripcion.Enter
        ' 🛡️ EL ESCUDO ADUANERO: Si el foco entra al combo pero el buscador está vacío o listo,
        ' desviamos el cursor obligatoriamente a la caja de texto para unificar el criterio
        If TxtBuscarLetras.Focused = False Then
            TxtBuscarLetras.Focus()
            TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
        End If
    End Sub

    Public Function BuscarLetras(combo As String) As String
        ' Cerramos cualquier lector abierto preventivamente
        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
            drMdb1.Close()
        End If

        ' =====================================================================
        ' MODO: DESCRIPCIÓN
        ' =====================================================================
        If combo = "descripcion" Then
            ' 1. Desconectamos el evento por seguridad
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            ' 🚀 CONTROL DE TEXTO CORTO REPARADO: INMUNE AL BLOQUEO DE ÍNDICES Y DATASOURCE
            If vLetras.Length <= 2 Then
                ' 1. Rompemos el candado de datos y vaciamos primero para liberar la RAM
                CmbDescripcion.DataSource = Nothing
                If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

                ' 2. Saneamos los índices de selección de forma segura antes de tocar la persiana
                CmbDescripcion.SelectedIndex = -1

                ' 3. 🎯 LA CLAVE MAESTRA: Cerramos la persiana gráfica envuelta en un cortafuegos para evitar el rebote de Windows
                Try
                    CmbDescripcion.DroppedDown = False
                Catch
                    ' Absorbe cualquier micro-rebote de foco del teclado de Windows
                End Try

                vCombo = "descripcion"

                ' Volvemos a conectar el evento preventivamente antes de salir volando
                AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
                Return "" ' Salimos inmediatamente para no ejecutar la consulta SQL vacía
            End If

            Dim letrasLimpias As String = vLetras.Replace("'", "''")
            cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE UCase(DescripcionAPU) LIKE '%" & letrasLimpias.ToUpper() & "%' AND DescripcionAPU <> 'Saldo Inicial'"

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    ' Si hay registros, limpiamos y llenamos las sugerencias
                    CmbDescripcion.SelectedIndex = -1
                    If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

                    While drMdb1.Read()
                        Dim desc As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
                        If Not String.IsNullOrEmpty(desc) Then CmbDescripcion.Items.Add(desc)
                    End While

                    ' Desplegamos la persiana gráfica con los resultados
                    CmbDescripcion.DroppedDown = True
                    vCombo = "descripcion"
                Else
                    ' =====================================================================
                    ' 🛠️ CONTROL DE BÚSQUEDA VACÍA SANEADO
                    ' =====================================================================
                    drMdb1.Close()

                    ' Cerramos la persiana y vaciamos la lista vieja para que no muestre datos erróneos
                    CmbDescripcion.DroppedDown = False
                    CmbDescripcion.SelectedIndex = -1
                    If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

                    ' Guardamos el estado especial de que esta descripción es NUEVA
                    vCombo = "descripcion_vacia"
                End If

                ' Doble comprobación de seguridad para asegurar el cierre del lector
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorBuscarLetrasDescripcion") & ": " & ex.Message)
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            End Try

            ' Volvemos a conectar el evento limpiamente al finalizar
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        End If

        Return ""
    End Function

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        ' 🚀 REPARADO MODO ESCÁNER: Conectado al MouseMove del formulario para sortear el bloqueo del .Enabled = False

        ' Diccionario con tus controles deshabilitados y sus ToolTips correspondientes
        Dim controlesBloqueados As New Dictionary(Of Control, ToolTip) From {
            {Me.TxtDescripcion, TL(11)}
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
                    If control Is Me.TxtDescripcion Then
                        textoCartelito = TxtDescripcion.Text
                    End If

                    ' Hacemos brotar el globo flotante reluciente desplazado 15 píxeles para que no lo tape el cursor
                    tool.Show(textoCartelito, Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón se sale del perímetro de los cuadros grises, apagamos los carteles de inmediato
        TL(11).Hide(Me)
    End Sub

End Class