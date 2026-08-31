Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntes

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, strText, vIntro, vLetras, vCombo, vDescripcion As String
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(13) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())
    Dim textoAutocompletadoEnAzul As String = ""
    Public vAceptarSalir As String = "NO"

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Label7.Text = vMoneda
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
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, resManager.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnDescripcion, rmse.GetString("BtnDescripcion"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.TxtBuscarLetras, rmse.GetString("TxtABuscar"))
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnAyuda, rmse.GetString("BtnAyuda"))

        CmbConcepto.DropDownStyle = ComboBoxStyle.DropDown

        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()

        ' Llenar el Combo Concepto de forma segura y traducida (IntroApuntes)
        '********************************************************************
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

        TxtImporte.Text = "0"

        ' --- ÚLTIMAS LÍNEAS DE TU INTROAPUNTES_LOAD ---
        TxtImporte.Text = "0"

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

                    ' 🛡️ 1. DESCONECTAMOS LOS CABLES ELÉCTRICOS DE LA RAM
                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

                    ' 2. TRASPASO DE TEXTOS EN FRÍO
                    CmbDescripcion.Text = vDescripcion

                    ' 🪓 EL TRUCO DEL INGENIERO JEFE: Desactivamos la casilla superior provisionalmente
                    ' para que Windows Forms no pueda arrastrarle el cursor bajo ningún concepto.
                    TxtBuscarLetras.Enabled = False
                    TxtBuscarLetras.Text = ""
                    vLetras = ""

                    ' 🔌 3. VOLVEMOS A CONECTAR LOS CABLES
                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

                    ' 4. ENTREGAMOS EL CURSOR DE FORMA INALTERABLE ABAJO
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
    Private Function GuardarApunteEnBaseDatos() As Boolean
        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

        Dim txtConcepto As String = Trim(CmbConcepto.Text)

        ' TRADUCCIÓN INVERSA: Buscamos la clave neutra para guardarla limpia en la BD (ej: GAS_NATURAL)
        vConceptoAPU = ObtenerClaveNeutral(txtConcepto, resManager)
        If String.IsNullOrEmpty(vConceptoAPU) Then vConceptoAPU = txtConcepto

        vDescripcionAPU = Trim(CmbDescripcion.Text)
        vNotasAPU = Trim(TxtNota.Text)
        vCuentaAPU = Trim(CmbCuenta.Text)

        ' MULTIIDIOMA: Pasamos la fecha como objeto Date puro de .NET. Ya NO necesitas formatearla a String.
        Dim fechaAsiento As Date = DateTimePicker1.Value.Date

        ' Conversión segura multiidioma del cuadro de texto (¡Centralizado en tu módulo!)
        Dim importeDecimal As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)

        ' 4. Asignamos el valor limpio y exacto a tu variable aplicando el signo correcto
        If TxtTipoConcepto.Text = "GASTO" Then
            vImporteAPU = -Math.Abs(importeDecimal)
        Else
            vImporteAPU = Math.Abs(importeDecimal)
        End If

        ' 1. CONSTRUCCIÓN SQL PARA EL INSERT PARAMETRIZADO (Adiós a #, comillas simples y Str)
        vAñadirSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) " &
                     "VALUES (?, ?, ?, ?, ?, ?, ?)"

        cmdMdb1cr.CommandText = vAñadirSql

        ' Limpiamos y asignamos los parámetros en el orden EXACTO de los comodines '?'
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@FechaAPU", fechaAsiento)
        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vConceptoAPU)
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionAPU", vDescripcionAPU)

        ' Usamos Currency para que Access no se sature con la precisión decimal
        Dim paramImporte As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteAPU", OleDb.OleDbType.Currency)
        paramImporte.Value = Math.Round(vImporteAPU, 2)

        cmdMdb1cr.Parameters.AddWithValue("@EjercicioAPU", CInt(vAñoEjercicio))
        cmdMdb1cr.Parameters.AddWithValue("@NotasAPU", vNotasAPU)
        cmdMdb1cr.Parameters.AddWithValue("@CuentaAPU", vCuentaAPU)

        Try
            ' Ejecutamos la inserción de forma segura
            cmdMdb1cr.ExecuteNonQuery()

            ' 2. REFRESH DE LA GRILLA PARAMETRIZADO
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes " &
                       "WHERE apuntes.EjercicioAPU = ? " &
                       "ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"

            cmdMdb1cr.CommandText = vtipoSql

            ' Limpiamos y asignamos el parámetro del WHERE
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("@EjercicioAPU", CInt(vAñoEjercicio))

            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(frmApuntesContables.DgvApuntes)

            Return True
        Catch ex As Exception
            MessageBox.Show(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.Message,
                            "Error de Almacenamiento",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
            Return False
        End Try
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
            'TxtBuscarLetras.Enabled = False

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
        ' OPTIMIZACIÓN CRÍTICA: Eliminamos la consulta SQL redundante a la base de datos.
        ' Cuando el usuario hace clic en un elemento de la lista, el combo ya adquiere ese texto de forma nativa.
        ' No necesitas abrir drMdb1, ni hacer un SELECT, ni volver a setear el .Text.

        If vIntro = "NO" Then
            ' Guardamos el valor seleccionado en tu variable global por si la usas en otro sitio
            vDescripcion = CmbDescripcion.Text.Trim()
        End If
    End Sub

    Private Sub CmbDescripcion_GotFocus(sender As Object, e As EventArgs) Handles CmbDescripcion.GotFocus
        ' 🛡️ CORTAFUEGOS DE ALTA NUEVA (Tu escudo definitivo contra el robo del cursor)
        ' Si venimos de pulsar "SÍ" para añadir una descripción, saltamos la inicialización
        ' para evitar que los refrescos automáticos de Windows le devuelvan el foco al TextBox superior.
        If vIntro = "SI" Then
            vCombo = "descripcion"
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
            Exit Sub ' 🪓 Cortamos el hilo aquí y dejamos el cursor parpadeando inmóvil en el combo!
        End If

        ' =====================================================================
        ' 📁 COMPORTAMIENTO CLÁSICO DE TU FORMULARIO (Para cuando navegas normal)
        ' =====================================================================
        ' Si venimos del combo de conceptos, nos aseguramos de dejar guardado su valor real
        If vCombo = "concepto" Then
            CmbConcepto.DroppedDown = False
            CmbConcepto.Text = vConcepto
        End If

        ' Preparamos el entorno para la descripción normal
        vIntro = "NO"
        vCombo = "descripcion"

        ' Cambiamos Lower por Normal para liberar las Mayúsculas/Minúsculas
        TxtBuscarLetras.CharacterCasing = CharacterCasing.Normal

        ' CONTROL INTELIGENTE DE BORRADO (MANTENIDO)
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
        ' Limpiamos el texto del buscador de forma segura
        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        TxtBuscarLetras.Text = ""
        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

        ' Lo dejamos deshabilitado momentáneamente mientras se introduce el dinero
        'TxtBuscarLetras.Enabled = False
        vCombo = ""
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
            LlenarComboConceptosIntroApuntes(Me.CmbConcepto)

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

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        ' 1. 🛡️ EL ESCUDO UNIVERSAL ADMITE TODO: Números, borrar (Control), punto, coma o el Intro
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> ","c AndAlso e.KeyChar <> ChrW(Keys.Enter) Then
            e.Handled = True
            Exit Sub
        End If

        ' 2. 🎯 AL PULSAR INTRO: Pasamos el rodillo internacional e inyectamos en la variable de apuntes
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True

            ' Invocamos tu función global centralizada: Cero grasa digital en la RAM
            Dim importeFinal As Decimal = ParsearImporteUniversal(TxtImporte.Text)

            ' Guardamos de forma segura en tu variable global de doble precisión (vImporteAPU)
            vImporteAPU = Convert.ToDouble(importeFinal)

            ' Formateamos la caja visual con el estándar de dos decimales de gala
            TxtImporte.Text = importeFinal.ToString("N2")

            ' Mandamos el cursor directo al combo de la Cuenta de forma dócil
            CmbCuenta.Select()
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
        vAceptarSalir = "SI"
        GrabarYRefrescarGrid()
    End Sub

    Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
        vAceptarSalir = "NO"
        GrabarYRefrescarGrid()

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
        TxtImporte.Text = "0"
        TxtNota.Text = ""
        'TxtDescripcion.Text = ""

        ' 3. ¡La clave! Forzamos al formulario a procesar los cambios visuales antes de seguir
        Application.DoEvents()

        ' 4. Regresamos el cursor a la Fecha para arrancar de nuevo el flujo de introducción
        DateTimePicker1.Focus()
    End Sub

    Public Sub GrabarYRefrescarGrid()
		If TxtImporte.Text <> "0" Then
			' 1. Convertimos el texto de la caja a un número Decimal limpio y seguro
			Dim importeNumerico As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)

			' 🚀 JUGADA MAESTRA 1: Capturamos la descripción antes del salto de foco.
			' Si la caja manual TxtDescripcion tiene letras, priorizamos su texto. Si no, usamos el Combo.
			Dim descripcionDefinitiva As String = CmbDescripcion.Text.Trim()
			If TxtDescripcion.Visible = True AndAlso TxtDescripcion.Text.Trim() <> "" Then
				descripcionDefinitiva = TxtDescripcion.Text.Trim()
			End If

			' 2. Conseguimos el texto exacto que hay en la pantalla (pasado a MAYÚSCULAS)
			Dim tipoEnPantalla As String = TxtTipoConcepto.Text.Trim().ToUpper()

			' 3. Recuperamos la traducción oficial en inglés (o el idioma activo) usando tu KEY real: "Tipo_Gasto"
			Dim tipoTraducido As String = ""
			If resManager IsNot Nothing Then
				tipoTraducido = resManager.GetString("Tipo_Gasto")
			End If

			' 4. EVALUACIÓN DE IDIOMA SEGURA: ¿Es "GASTO" en español o coincide con la traducción?
			If tipoEnPantalla = "GASTO" OrElse (tipoTraducido <> "" AndAlso tipoEnPantalla = tipoTraducido.Trim().ToUpper()) Then
				' Si es un gasto y el usuario lo escribió en positivo, lo convertimos a negativo matemáticamente
				If importeNumerico > 0 Then
					importeNumerico = importeNumerico * -1
				End If
			End If

			' Asignamos el valor numérico final a tu variable global
			vImporteAPU = importeNumerico
			vNotasAPU = TxtNota.Text

			' --- RECUPERAR NOMBRE DE CUENTA EN ESPAÑOL SEGURO ---
			vCuentaAPU = ""
			If CmbCuenta.SelectedIndex >= 0 Then
				cmdMdb1cr.CommandText = "SELECT NombreCUE FROM cuentas ORDER BY NombreCUE ASC"
				Try
					Dim drCuentaGuardar As OleDbDataReader = cmdMdb1cr.ExecuteReader()
					Dim contCUE As Integer = 0
					While drCuentaGuardar.Read()
						If contCUE = CmbCuenta.SelectedIndex Then
							vCuentaAPU = drCuentaGuardar("NombreCUE").ToString()
							Exit While
						End If
						contCUE += 1
					End While
					drCuentaGuardar.Close()
				Catch ex As Exception
					' Si falla por cualquier motivo, dejamos el texto del combo como salvavidas
					vCuentaAPU = CmbCuenta.Text.ToString()
				End Try
			Else
				vCuentaAPU = CmbCuenta.Text.ToString()
			End If

			Dim idConceptoAsiento As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
			Dim idCuentaAsiento As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

			' 2. Construimos la SQL relacional con parámetros puros para evitar errores de comas o tipos
			vAñadir = "INSERT INTO apuntes " &
				"(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, NotasAPU, CuentaAPU, EjercicioAPU) " &
				"VALUES (?, ?, ?, ?, ?, ?, ?)"

			cmdMdb1cr.CommandText = vAñadir
			cmdMdb1cr.Parameters.Clear() ' Limpieza estricta de memoria RAM

			' 3. Inyectamos los valores en el orden exacto de los signos de interrogación '?'
			cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = DateTimePicker1.Value.Date
			cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoAsiento ' 🌟 Inyecta el ID del Concepto
			cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = CmbDescripcion.Text.Trim()
			cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = importeNumerico
			cmdMdb1cr.Parameters.Add("@Not", OleDb.OleDbType.VarWChar).Value = TxtNota.Text.Trim()
			cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = Convert.ToInt32(CmbCuenta.SelectedValue)
			'cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuentaAsiento    ' 🌟 Inyecta el ID de la Cuenta
			cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)

			''=========================================================================
			''🕵️ CHIVATO PARAMETROS CONTABLE: SIMULACIÓN DE LA SQL REAL QUE VA A IR A ACCESS
			''=========================================================================
			'Dim sqlSimulada As String = vAñadir
			'' Vamos reemplazando de izquierda a derecha cada signo '?' por su valor real formateado
			'' 1. Fecha
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, "#" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "#")
			'' 2. ID Concepto
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, idConceptoAsiento.ToString())
			'' 3. Descripción
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, "'" & TxtDescripcion.Text.Replace("'", "''") & "'")
			'' 4. Importe
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, TxtImporte.Text.Replace(",", "."))
			'' 5. Notas
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, "'" & TxtNota.Text.Replace("'", "''") & "'")
			'' 6. ID Cuenta
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, idCuentaAsiento.ToString())
			'' 7. Ejercicio
			'sqlSimulada = ReemplazarPrimerInterrogante(sqlSimulada, vAñoEjercicio)
			'' Lanzamos la ventana para inspeccionar el texto exacto con tus datos
			'MsgBox("SQL SIMULADA DE GRABACIÓN:" & vbCrLf & vbCrLf & sqlSimulada, MsgBoxStyle.Information, "Depuración de Parámetros")

			Try
				cmdMdb1cr.ExecuteNonQuery()
			Catch ex As Exception
				MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
			End Try


			If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
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
			Else
				' =========================================================================
				' 🌟 RAMA B: REFRESCO CUANDO EL LISTBOX LATERAL TIENE MULTISELECCIÓN
				' =========================================================================
				' 1. Saneamiento preventivo de parámetros en la memoria de la app
				cmdMdb1cr.Parameters.Clear()

				' 2. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
				Dim idConceptoSaldo As Integer = 1
				Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
					Dim resId = cmdBuscarId.ExecuteScalar()
					If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
				End Using

				' Guardamos si el filtro de fechas de la pantalla principal está activo
				Dim tieneFechasActivo As Boolean = (frmApuntesContables.BtnFiltroFecha.Enabled = False)

				' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS RELACIONALES (Tu diseño perfecto para nombres claros)
				Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"
				vtipoSql = sqlBase

				If frmApuntesContables.BtnFechasClick = "SI" Then
					vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
				Else
					vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
				End If

				' 🌟 RECOLECCIÓN DE IDs NUMÉRICOS DESDE EL LISTBOX DE LA PANTALLA PRINCIPAL
				Dim listaIdsConceptos As New List(Of Integer)
				Dim i As Integer

				For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
					Dim vConceptoFila As String = frmApuntesContables.ListBox1.SelectedItems(i).ToString()
					If vConceptoFila.StartsWith("**") Then Continue For

					' Buscamos el ID numérico original mapeando el texto del ListBox
					Dim idConceptoEncontrado As Integer = 0
					Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
						cmdId.Parameters.AddWithValue("?", vConceptoFila)
						Dim resId = cmdId.ExecuteScalar()
						If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoEncontrado = Convert.ToInt32(resId)
					End Using

					If idConceptoEncontrado > 0 Then listaIdsConceptos.Add(idConceptoEncontrado)
				Next

				' Si por un fallo la lista está vacía, le inyectamos un 0 de salvavidas
				If listaIdsConceptos.Count = 0 Then listaIdsConceptos.Add(0)

				' Inyectamos el filtro IN de enteros inmune a fallos de combinación
				vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIdsConceptos) & ") "

				' CORRECCIÓN: Filtro por ID numérico de Cuenta (Leyendo el SelectedValue de la pantalla principal)
				If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
					Dim idCuentaPrincipal As Integer = Convert.ToInt32(frmApuntesContables.CmbCuenta.SelectedValue)
					vtipoSql += $" And apuntes.CuentaAPU = {idCuentaPrincipal} "
				End If

				' 🌟 CRÍTICO: Las interrogaciones de fecha van SIEMPRE al final de las condiciones del WHERE
				If tieneFechasActivo Then
					vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
					vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
					vtipoSql += " And apuntes.FechaAPU >= ?"
					vtipoSql += " And apuntes.FechaAPU <= ?"

					cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
					cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
				End If

				vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
				vtipoGrid = "APUNTES_CONTABLES"

				LlenarGrid(vtipoSql, vtipoGrid, "1")
				TraducirGridApuntesBD(frmApuntesContables.DgvApuntes)

				If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
					vFila = frmApuntesContables.DgvApuntes.RowCount - 1
					frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
					frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
				End If
			End If
		Else
			MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
			vAceptarSalir = "NO"
			TxtImporte.Select()
		End If
		If vAceptarSalir = "SI" Then
			Me.Close()
		End If
	End Sub

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
        ' 1. Abrimos la pantalla de mantenimiento de cuentas del formulario principal
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()

        ' =========================================================================
        ' 🌟 RECARGA DE LA NUEVA ERA: ENLACE SIMÉTRICO DE CUENTAS BANCARIAS
        ' =========================================================================
        ' Encendemos el escudo protector para que los eventos de cambio no se vuelvan locos al recargar
        cargandoFormulario = True

        Try
            ' 2. Llamamos a tu rutina exclusiva para refrescar e inyectar las cuentas con sus IDs numéricos
            LlenarComboCuentasGenerico(Me.CmbCuenta)

            ' 3. Apagamos el escudo protector para permitir la interacción del usuario
            cargandoFormulario = False

            ' 4. Volvemos a aplicar tu vaivén maestro de índices para forzar el relleno en la rejilla
            If CmbCuenta.Items.Count > 0 Then
                CmbCuenta.SelectedIndex = -1 ' Reseteamos a vacío primero
                CmbCuenta.SelectedIndex = 0  ' Seleccionamos el primer elemento de forma segura
            End If

        Catch ex As Exception
            cargandoFormulario = False
            MsgBox(resManager.GetString("ErrorRefrecarCUE") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub BtnAyuda_Click(sender As Object, e As EventArgs) Handles BtnAyuda.Click
        ' 🛠️ CONTROL DE VENTANA FLOTANTE INDEPENDIENTE
        ' Comprobamos si la ventana de ayuda ya está abierta en pantalla
        Dim frmExistente As AyudaApuntes = Application.OpenForms.OfType(Of AyudaApuntes)().FirstOrDefault()

        If frmExistente IsNot Nothing Then
            ' Si ya estaba abierta, la cerramos (esto disparará automáticamente el evento FormClosed)
            frmExistente.Close()
        Else
            ' --- DESPLAZAR EL FORMULARIO ACTUAL A LA IZQUIERDA ---
            Dim pixelesDesplazamiento As Integer = 150
            Me.Left -= pixelesDesplazamiento

            ' Creamos la instancia de la ventana de ayuda
            Dim frmAyuda As New AyudaApuntes()

            ' --- DETECTAR EL CIERRE (BOTÓN O CRUZ X) ---
            ' Usamos AddHandler para ejecutar código cuando frmAyuda se cierre por cualquier motivo
            AddHandler frmAyuda.FormClosed, Sub(s, ev)
                                                ' Cuando la ayuda se cierra, devolvemos el formulario principal a la derecha
                                                Me.Left += pixelesDesplazamiento
                                            End Sub

            ' CÁLCULO DE POSICIÓN: Se calcula usando la NUEVA posición del formulario actual
            Dim x As Integer = Me.Location.X + Me.Width
            Dim y As Integer = Me.Location.Y

            frmAyuda.Location = New Point(x, y)

            ' La mostramos en modo "Show"
            frmAyuda.Show(Me)
        End If
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
        LlenarDescripcion()
    End Sub

    Private Sub DateTimePicker1_LostFocus(sender As Object, e As EventArgs) Handles DateTimePicker1.LostFocus
        BtnCalculadora.TabIndex = 0
        BtnConcepto.TabIndex = 0
        BtnDescripcion.TabIndex = 0
        BtnHoy.TabIndex = 0
        BtnCuenta.TabIndex = 0
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
            ' 🌟 INYECCIÓN DIRECTA EN LA INTERFAZ DESDE LA MEMORIA CACHÉ
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

                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                CmbDescripcion.Text = tradDesc
                vDescripcion = tradDesc

                If TypeOf TxtDescripcion Is TextBox Then TxtDescripcion.Text = tradDesc
            End If
            ' =====================================================================

            ' Limpiamos el chivato para la próxima búsqueda, cerramos la lista y saltamos
            textoAutocompletadoEnAzul = ""
            If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False
            CmbDescripcion.Select()
            CmbConcepto.Text = textoBuscar
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

                    ' --- TRADUCIR LAS DESCRIPCIONES (Desc_NOMBRE) ---
                    Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                    Dim tradDesc As String = resManager.GetString(llaveDesc)

                    ' Si no tiene traducción en el ResX, dejamos la descripción original de la BD
                    If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                    CmbDescripcion.Text = tradDesc
                    TxtDescripcion.Text = tradDesc

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

End Class