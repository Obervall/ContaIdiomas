Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class AprendizajeBancario

    Private cargandoFormulario As Boolean = True
    Public vCuentaAPU, vIntro, vLetras, vCombo, vDescripcion As String
    Private TL(10) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())
    Private vIdExtractoActual As Integer = 0
    Public vValorPrimero As Integer = 1
    Public vValorTotal As Integer = 0
    Private IsLimpiandoCombo As Boolean = False
    Private buscandoDescripcion As Boolean = False
    ' Esta bandera controlará cuándo permitimos que el TextChanged abra la tabla
    Private cargandoDatosBanco As Boolean = False


    Public Sub CargarPrimerConceptoBancario()
        Try
            ' 🌟 ACTIVAMOS EL INTERRUPTOR: Avisamos que el programa está rellenando casillas de forma automática
            cargandoDatosBanco = True
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

                        ' 1. Clavamos el combo de la Cuenta indexando el ID que capturó el radar
                        Dim idCuentaBanco As Integer = Convert.ToInt32(dr("CuentaAPU"))
                        CmbCuenta.SelectedValue = idCuentaBanco

                        ' 1. Seleccionamos el concepto por defecto
                        CmbConcepto.SelectedValue = 1

                        ' 2. Ejecutamos la rutina unificada (que rellena TxtBuscarLetras y guarda el texto en vDescripcion)
                        LlenarDescripcionConcepto()

                        ' 3. Nos aseguramos de que el DataGridView esté cerrado e invisible
                        DgvDescripcion.DataSource = Nothing
                        DgvDescripcion.Visible = False

                        ' 4. Ponemos el foco final en el concepto (o donde desees arrancar)
                        CmbConcepto.Focus()

                        ' 5. Volvemos a encender el evento. Ahora está listo para cuando el usuario teclee a mano
                        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                        ' =========================================================================

                        ' Clavamos el cursor parpadeando allí dentro para que el usuario solo tenga que teclear
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
        Finally
            ' 🌟 APAGAMOS EL INTERRUPTOR: La pantalla ya está dibujada y lista para el usuario
            cargandoDatosBanco = False
        End Try
    End Sub

	Private Sub AprendizajeBancario_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        GroupBox1.Text = rmse.GetString("$this.Text") & ": " & frmApuntesContables.CmbCuenta.Text

        ' 🔒 DECAPITAMOS EL CAMINO DEL TABULADOR EN LAS CASILLAS DE SOLO LECTURA
        ' El cursor jamás se detendrá aquí dentro al pulsar el Tabulador
        TxtImporte.TabStop = False
        BtnConcepto.TabStop = False
        CmbCuenta.TabStop = False ' Si el banco tampoco se toca, lo capamos también

        ' 🎯 EL FOCO DE ENTRADA: Obligamos al cursor a nacer parpadeando en el combo limpio
        CmbConcepto.TabStop = True
        CmbConcepto.Focus()

        Me.KeyPreview = True

        Label7.Text = vMoneda
        Label9.Text = vMoneda

        If vSaldoFinal = 0 Then
            TxtSaldoFinal.Text = "0,00"
        Else
            TxtSaldoFinal.Text = vSaldoFinal.ToString("N2")
        End If

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
        TL(5).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteAsiento"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtBuscarLetras, rmse.GetString("TxtABuscar"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnSaltarApuntes, rmse.GetString("BtnSaltarApuntes.Text"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.TxtDescripcion, My.Settings.RutaBD)

        ' Combo Conceptos: solo selección, sin escritura libre
        CmbConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbConcepto.AutoCompleteMode = AutoCompleteMode.None
        CmbConcepto.AutoCompleteSource = AutoCompleteSource.None

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

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        vLetras = TxtBuscarLetras.Text.Trim()

        ' 🌟 CORTAFUEGOS A LA ANTIGUA USANZA:
        ' Si el texto coincide exactamente con la descripción oficial del concepto (vDescripcion)
        ' o si es el primerísimo arranque de la pantalla, no permitimos que la tabla se abra sola.
        If Not String.IsNullOrEmpty(vDescripcion) AndAlso vLetras.ToUpper() = vDescripcion.Trim().ToUpper() Then
            DgvDescripcion.DataSource = Nothing
            DgvDescripcion.Visible = False
            Return
        End If

        ' Si hay menos de 3 letras → ocultamos y vaciamos el DataGridView
        If vLetras.Length < 3 Then
            DgvDescripcion.DataSource = Nothing
            DgvDescripcion.Visible = False
            Return
        End If

        ' La tabla SOLO se abrirá si el usuario interactúa físicamente escribiendo en el teclado
        DgvDescripcion.Visible = True
        BuscarLetras("descripcion")
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
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

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

    Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
        GrabarYRefrescarGrid()

        ' =====================================================================
        ' 🛠️ REINICIO ASISTIDO DE PASOS PARA INTRODUCIR OTRO APUNTE
        ' =====================================================================
        ' 1. Apagamos el TextChanged superior para evitar llamadas falsas a la BD al limpiar
        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        LlenarDescripcionConcepto()
        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

        ' 🌟 NUEVO: Limpiamos y ocultamos el DataGridView para el nuevo apunte
        DgvDescripcion.DataSource = Nothing
        DgvDescripcion.Visible = False

        ' 2. Reestablecemos el estado inicial del buscador de descripciones
        TxtBuscarLetras.Enabled = True
        vLetras = ""
        vIntro = "NO"

        ' Limpias los textos normales
        TxtNota.Text = ""

        Try
            ' 🌟 ACTIVAMOS EL INTERRUPTOR: Avisamos que el programa está rellenando casillas de forma automática
            cargandoDatosBanco = True

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
        Finally
            ' 🌟 APAGAMOS EL INTERRUPTOR: La pantalla ya está dibujada y lista para el usuario
            cargandoDatosBanco = False
        End Try

        ' 3. ¡La clave segura! Forzamos al formulario a repintarse limpiamente sin DoEvents
        Me.Refresh()

        ' 4. Regresamos el cursor a la Fecha para arrancar de nuevo el flujo de introducción
        CmbConcepto.Focus()
    End Sub

    Public Sub GrabarYRefrescarGrid()
        ' =====================================================================
        ' 🛑 VALIDACIÓN: PREVENIR DESCRIPCIÓN VACÍA
        ' =====================================================================
        ' Evaluamos tanto la variable global como el cuadro de texto por seguridad

        If String.IsNullOrWhiteSpace(vDescripcion) AndAlso String.IsNullOrWhiteSpace(TxtBuscarLetras.Text) Then

            ' Mostramos un mensaje de advertencia usando tus archivos de recursos de idioma
            MessageBox.Show(resManager.GetString("ErrorDescripcionVacia"),
                    resManager.GetString("Error"),
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)

            ' Devolvemos el foco al buscador para que el usuario escriba algo válido
            TxtBuscarLetras.Focus()

            ' Frenamos el método AQUÍ. Nada de lo de abajo (el INSERT) se ejecutará.
            Return
        End If

        'Si la variable está vacía pero el TextBox tiene texto, rescatamos el texto del cuadro
        If String.IsNullOrWhiteSpace(vDescripcion) Then
            vDescripcion = TxtBuscarLetras.Text.Trim()
        End If

        If TxtImporte.Text <> "0" Then
            ' 1. Convertimos el texto de la caja a un número Decimal limpio y seguro
            Dim importeNumerico As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)

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
                cmdMdb1cr.CommandText = "Select NombreCUE FROM cuentas ORDER BY NombreCUE ASC"
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

            Dim descripcionDefinitiva As String = vDescripcion.Replace("'", "''").Trim()

            ' 3. Inyectamos los valores en el orden exacto de los signos de interrogación '?'
            cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = DateTimePicker1.Value.Date
            cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoAsiento ' 🌟 Inyecta el ID del Concepto
            cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = descripcionDefinitiva
            cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = importeNumerico
            cmdMdb1cr.Parameters.Add("@Not", OleDb.OleDbType.VarWChar).Value = TxtNota.Text.Replace("'", "''").Trim()
            cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = Convert.ToInt32(CmbCuenta.SelectedValue)
            cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)

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
        End If
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

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub


    Private Sub TxtBuscarLetras_Leave(sender As Object, e As EventArgs) Handles TxtBuscarLetras.Leave
        Dim textoActual As String = TxtBuscarLetras.Text.Trim()

        ' Comparamos contra vDescripcion (el texto oficial del concepto)
        If DgvDescripcion.Visible = False AndAlso textoActual.ToUpper() <> vDescripcion.ToUpper() Then
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = vDescripcion
            TxtDescripcion.Text = vDescripcion ' Sincronizamos el chivato
            vIntro = "NO"
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        End If
    End Sub


    Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
        If e.KeyCode = Keys.Enter Then
            ' Anulamos el comportamiento por defecto del Intro en el ComboBox
            e.SuppressKeyPress = True
            e.Handled = True

            Dim idConceptoSel As Integer = 0
            Dim textoBuscar As String = CmbConcepto.Text.Trim()

            If Not String.IsNullOrEmpty(textoBuscar) Then
                Try
                    Dim dt As DataTable = CType(CmbConcepto.DataSource, DataTable)
                    If dt IsNot Nothing Then
                        ' Buscamos la fila que coincide con el texto escrito
                        Dim filas() As DataRow = dt.Select("TextoCombo = '" & textoBuscar.Replace("'", "''") & "'")

                        If filas.Length = 0 Then
                            filas = dt.Select("TextoCombo LIKE '" & textoBuscar.Replace("'", "''") & "%'")
                        End If

                        ' Si encontramos el concepto, forzamos la selección en el combo
                        If filas.Length > 0 Then
                            idConceptoSel = Convert.ToInt32(filas(0)("IdConceptoCON"))
                            CmbConcepto.SelectedIndex = dt.Rows.IndexOf(filas(0))
                        End If
                    End If
                Catch ex As Exception
                    ' Cortafuegos silencioso
                End Try
            End If

            ' Si encontramos un concepto válido, disparamos nuestra rutina unificada
            If idConceptoSel > 0 Then
                CmbConcepto.SelectedValue = idConceptoSel
                LlenarDescripcionConcepto() ' 🌟 Rellena textos, traduce y sincroniza variables de golpe
            End If

            ' Si el combo estaba desplegado, lo cerramos
            If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False
            CmbConcepto.Text = textoBuscar

            ' 🚀 LLEVAMOS EL CURSOR AL BUSCADOR DIRECTAMENTE Y SIN RUIDO
            ' El BeginInvoke sigue siendo una buena práctica aquí para asegurar que el foco viaje 
            ' limpiamente al TxtBuscarLetras tras procesar el Intro
            BeginInvoke(Sub()
                            Try
                                TxtBuscarLetras.Focus()
                                TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
                            Catch
                                ' Cortafuegos silencioso
                            End Try
                        End Sub)
        End If
    End Sub

    Private Sub CmbConcepto_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbConcepto.MouseClick
        ' Solo forzamos el despliegue automático si el usuario NO ha pulsado la flecha nativa
        ' (Nos aseguramos comprobando si la lista ya está abierta o abriéndola suavemente)
        If CmbConcepto.Items.Count <> 0 AndAlso Not CmbConcepto.DroppedDown Then
            CmbConcepto.DroppedDown = True
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Si estamos buscando descripciones → NO tocar el combo
        If buscandoDescripcion Then Exit Sub

        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        If vIntro = "NO" Then
            TxtBuscarLetras.Text = ""
            LlenarDescripcionConcepto()

        End If
    End Sub

    Public Function BuscarLetras(combo As String) As String

        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
            drMdb1.Close()
        End If

        If combo = "descripcion" Then
            Dim texto As String = If(vLetras, "").Trim()
            Dim letrasLimpias As String = texto.Replace("'", "''").ToUpper()

            cmdMdb1cr.CommandText =
                "Select DISTINCT DescripcionAPU " &
                "FROM apuntes " &
                "WHERE UCase(DescripcionAPU) Like '%" & letrasLimpias & "%' " &
                "AND DescripcionAPU <> 'Saldo Inicial'"

            Dim dt As New DataTable()

            Try
                Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                    dt.Load(dr)
                End Using

                DgvDescripcion.DataSource = dt

                If dt.Rows.Count > 0 Then
                    ' Si hay filas, configuramos columna y nos aseguramos de que se vea
                    If DgvDescripcion.Columns.Contains("DescripcionAPU") Then
                        DgvDescripcion.Columns("DescripcionAPU").HeaderText = resManager.GetString("Descripcion")
                        DgvDescripcion.Columns("DescripcionAPU").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    End If

                    If vValorPrimero = 1 Then
                        DgvDescripcion.Visible = False
                    Else
                        DgvDescripcion.Visible = True
                    End If
                Else
                    ' Si la búsqueda da 0 filas (como "impoo"), lo ocultamos
                    DgvDescripcion.Visible = False
                End If

            Catch ex As Exception
                DgvDescripcion.DataSource = Nothing
                DgvDescripcion.Visible = False
            End Try
        End If
        Return ""
    End Function

    Private Sub DgvDescripcion_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DgvDescripcion.CellDoubleClick
        ' Validamos que el usuario haya hecho clic en una fila de datos real y no en el encabezado
        If e.RowIndex >= 0 Then
            ProcesarSeleccionDescripcion(e.RowIndex)
        End If
    End Sub

    Private Sub DgvDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles DgvDescripcion.KeyDown
        ' Si el usuario presiona Enter sobre el DataGridView
        If e.KeyCode = Keys.Enter Then
            ' 🌟 ¡LA CLAVE! Cancelamos por completo la pulsación para que NO baje de registro
            e.SuppressKeyPress = True

            ' Validamos que exista una fila seleccionada y procesamos
            If DgvDescripcion.CurrentRow IsNot Nothing Then
                ProcesarSeleccionDescripcion(DgvDescripcion.CurrentRow.Index)
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBuscarLetras.KeyDown
        ' 1. Si el usuario presiona INTRO en el buscador
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True ' Evitamos el pitido de Windows

            Dim textoBuscado As String = TxtBuscarLetras.Text.Trim()

            ' CASO A: El campo está completamente vacío -> Mostramos toda la lista
            If String.IsNullOrWhiteSpace(textoBuscado) Then
				vLetras = ""
                DgvDescripcion.Visible = True
                BuscarLetras("descripcion")
                DgvDescripcion.Focus()
                Return
            End If

            ' 🌟 ¡LA CORRECCIÓN CLAVE ADAPTADA AL BANCO!
            ' Si la tabla está oculta, pero el texto del cuadro coincide exactamente con la descripción 
            ' oficial del concepto que tenemos guardada en memoria (vDescripcion), el usuario acepta la de por defecto.
            If DgvDescripcion.Visible = False AndAlso textoBuscado.ToUpper() = vDescripcion.Trim().ToUpper() Then
                ' Forzamos a TxtDescripcion a asimilar este texto para blindar también el evento Leave
                TxtDescripcion.Text = textoBuscado

                vIntro = "NO" ' No es una descripción nueva
                TxtNota.Focus() ' Saltamos directo a la nota (como tienes configurado aquí)
                Return
            End If

            ' CASO B: Hay texto, está oculto y NO coincide -> ¡AQUÍ SÍ ES UNA DESCRIPCIÓN TOTALMENTE NUEVA!
            If DgvDescripcion.Visible = False OrElse DgvDescripcion.Rows.Count = 0 Then

                Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(
                rmse.GetString("NoExistenDescripciones") & ": -" & textoBuscado.ToUpper() & "-" & vbCrLf & "¿" & rmse.GetString("AñadirDescripcion") & "?",
                rmse.GetString("$this.Text")
            )

                If respuesta = MsgBoxResult.Yes Then
                    vIntro = "SI"
                    vDescripcion = textoBuscado
                    ' Actualizamos el almacén de texto original con el nuevo texto autorizado
                    'TxtDescripcion.Text = textoBuscado
                    DgvDescripcion.Visible = False
                    TxtNota.Focus()
                Else
                    TxtBuscarLetras.SelectAll()
                    TxtBuscarLetras.Focus()
                End If

                Return
            End If

            ' CASO C: Si el DataGridView SÍ tiene datos en pantalla, mandamos el foco a la tabla
            If DgvDescripcion.Visible AndAlso DgvDescripcion.Rows.Count > 0 Then
                DgvDescripcion.Focus()
            End If
        End If

        ' 2. Si el usuario presiona FLECHA ABAJO
        If e.KeyCode = Keys.Down AndAlso DgvDescripcion.Visible Then
            DgvDescripcion.Focus()
            e.Handled = True
        End If
    End Sub

    Private Sub ProcesarSeleccionDescripcion(rowIndex As Integer)
        ' 1. Extraemos el valor de la columna de la fila seleccionada
        vDescripcion = DgvDescripcion.Rows(rowIndex).Cells("DescripcionAPU").Value.ToString()

        ' 2. Pasamos el texto seleccionado de vuelta a tu TextBox para que se vea la elección
        TxtBuscarLetras.Text = vDescripcion

        ' 3. Ocultamos el DataGridView puesto que ya finalizó la selección
        DgvDescripcion.Visible = False

        ' 4. ¡AUTOMÁTICO! Enviamos el foco directamente al campo del nota
        TxtNota.Focus()
    End Sub

    Private Sub LlenarDescripcionConcepto()
        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' Extraemos datos del concepto seleccionado
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            ' Guardamos código para la BD
            vConcepto = codigoOriginal

            ' --- TRADUCIR TIPO ---
            Dim tradTipo As String = ""
            Select Case tipoOriginal.ToUpper()
                Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
            End Select
            If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
            TxtTipoConcepto.Text = tradTipo

            ' --- TRADUCIR DESCRIPCIÓN ---
            Dim llaveDesc As String = "Desc_" & codigoOriginal.ToUpper().Trim()
            Dim tradDesc As String = resManager.GetString(llaveDesc)
            If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

            ' Pintamos la interfaz de forma limpia
            TxtBuscarLetras.Text = tradDesc
            vDescripcion = tradDesc ' 🌟 Sincronizamos de inmediato la variable global

        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorSincronizarCON") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub BtnSaltarApuntes_Click(sender As Object, e As EventArgs) Handles BtnSaltarApuntes.Click
        If vIdExtractoActual = 0 Then Exit Sub

        Try
            ' 🌟 ACTIVAMOS EL INTERRUPTOR: Avisamos que el programa está rellenando casillas de forma automática
            cargandoDatosBanco = True

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
        Finally
            ' 🌟 APAGAMOS EL INTERRUPTOR: La pantalla ya está dibujada y lista para el usuario
            cargandoDatosBanco = False
        End Try
    End Sub

End Class