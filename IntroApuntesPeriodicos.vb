Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntesPeriodicos

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid, vLetras, vIntro, vDescripcion As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, vAnexo, vbOK As String
    Public vNumeroPagos, vDate3Year As Integer
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Encendemos el escudo protector al inicio del método para congelar los eventos
        cargandoFormulario = True
        cmdMdb1cr.Parameters.Clear()

        Me.KeyPreview = True

        Label7.Text = vMoneda
        ' 1. Convertimos el año base de forma segura a número entero
        Dim anioBase As Integer
        If Not Integer.TryParse(vAñoEjercicio, anioBase) Then
            ' Salvavidas: si falla o está vacío, usa el año actual
            anioBase = Date.Today.Year
        End If

        ' 2. Asignamos los valores numéricos limpios a tus variables globales
        vFecha1Enero = anioBase
        vFecha31Diciembre = anioBase

        ' 3. Creamos los objetos de fecha límites de forma nativa
        Dim fechaInicio As New Date(anioBase, 1, 1)
        Dim fechaFin As New Date(anioBase, 12, 31)

        ' 4. Aplicamos los rangos al control
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker1.MaxDate = fechaFin

        ' 5. Evaluamos la condición lógica de forma limpia convirtiendo a texto explícito
        If anioBase.ToString() <> vAñoActual.ToString() Then
            ' Si el año de ejercicio es diferente al actual, calculamos el año siguiente de forma exacta
            Dim anioSiguiente As Integer = anioBase + 1

            ' Asignamos el valor al 31 de diciembre del año siguiente (lógica de apuntes periódicos)
            DateTimePicker1.Value = New Date(anioSiguiente, 12, 31)
        Else
            ' Si coincide con el año en curso, se inicializa con la fecha de hoy
            ' Nota: Aseguramos que la conversión a fecha sea limpia e independiente del idioma
            DateTimePicker1.Value = Convert.ToDateTime(vfechaHoy)
        End If


        Dim TL(12) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.DateTimePicker1, rmse.GetString("SelecFechaPrimer"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptarSalir, resManager.GetString("ToolTipAceptar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, frmIntroApuntes.rmse.GetString("SelecConcepto"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, frmIntroApuntes.rmse.GetString("SelecCuenta"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtBuscarLetras, resManager.GetString("TxtABuscar"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, frmEditarApuntes.rmse.GetString("ToolTipIngresarImporte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, resManager.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.TxtNumeroPagos, rmse.GetString("IntroNumeroPagos"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.CmbPeriocidad, rmse.GetString("SelecPeriodo"))

        ' Combo Conceptos: solo selección, sin escritura libre
        CmbConcepto.DropDownStyle = ComboBoxStyle.DropDownList
        CmbConcepto.AutoCompleteMode = AutoCompleteMode.None
        CmbConcepto.AutoCompleteSource = AutoCompleteSource.None

        ' =========================================================================
        ' 🌟 CARGA DE COMBOS DE LA NUEVA ERA COMPACTADA POR FUNCIONES (MSIX)
        ' =========================================================================

        ' 1. Llenar el Combo Concepto de forma aislada
        LlenarComboConceptosPeriodicos(Me.CmbConcepto)

        '' =========================================================================
        '' 🚀 TRAMO 2 REPARADO: CAPTURA DE LA PRIMERA DESCRIPCIÓN DE LA BASE DE DATOS
        '' =========================================================================
        '' 1. Ocultamos preventivamente la rejilla flotante azul de sugerencias
        'DgvDescripcion.Visible = False

        '' 2. Consulta ultra-rápida TOP 1: Capturamos la primera descripción real que exista
        'cmdMdb1cr.CommandText = "SELECT TOP 1 DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'SALDO INICIAL' And DescripcionAPU Is Not Null ORDER BY DescripcionAPU ASC"
        'cmdMdb1cr.Parameters.Clear()

        'Dim primeraDescripcion As String = ""

        'Try
        '    Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
        '        ' Al usar un IF en lugar de un WHILE, leemos estrictamente la primera fila y cerramos el grifo
        '        If dr.Read() Then
        '            primeraDescripcion = dr.GetValue(0).ToString().Trim()
        '        End If
        '    End Using
        'Catch ex As Exception
        '    MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        'End Try

        '' 3. 🎯 SINCRONIZACIÓN ABSOLUTA DEL NUEVO TXTBUSCARLETRAS
        '' Si la base de datos nos ha devuelto texto, lo inyectamos al céntimo en el cuadro
        'If Not String.IsNullOrEmpty(primeraDescripcion) Then
        '    vDescripcion = primeraDescripcion
        '    TxtBuscarLetras.Text = primeraDescripcion
        'Else
        '    vDescripcion = ""
        '    TxtBuscarLetras.Text = ""
        'End If

        '' Marcamos que de entrada no es una descripción nueva de paquete
        'vIntro = "NO"


        '' =========================================================================
        '' 🚀 TRAMO 2 REPARADO: LLENAR COMBO DESCRIPCIÓN INMUNE A BLOQUEOS (MSIX)
        '' =========================================================================
        '' Saneamos el CommandText a mayúsculas para evitar descalces con el rodillo 2.5
        'cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'SALDO INICIAL' And DescripcionAPU Is Not Null ORDER BY DescripcionAPU ASC"
        'cmdMdb1cr.Parameters.Clear()
        'CmbDescripcion.Items.Clear()

        'Try
        '    Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
        '        While dr.Read()
        '            Dim descLimpia As String = dr.GetValue(0).ToString().Trim()
        '            If Not String.IsNullOrEmpty(descLimpia) Then
        '                CmbDescripcion.Items.Add(descLimpia)
        '            End If
        '        End While
        '    End Using
        'Catch ex As Exception
        '    MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        'End Try

        ' 3. Llenar el Combo Cuenta de forma aislada
        LlenarComboCuentasPeriodicos(Me.CmbCuenta)

        ' 4. Valores e interfaz por defecto (Tus líneas originales impecables)
        TxtImporte.Text = "0"

        ' Apagamos el escudo protector de forma segura al finalizar la inyección
        cargandoFormulario = False

        CmbConcepto.Select()

        ' Ajustamos de forma segura la periodicidad por defecto
        If CmbPeriocidad.Items.Count > 3 Then
            CmbPeriocidad.Text = CmbPeriocidad.Items(3).ToString()
        End If
    End Sub

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        vLetras = TxtBuscarLetras.Text.Trim()

        '' 🌟 PROTECCIÓN: Si es la descripción por defecto, no hace falta buscar nada
        'If vLetras = TxtDescripcion.Text Then
        '    DgvDescripcion.Visible = False
        '    Return
        'End If

        ' Si hay menos de 3 letras → ocultamos y vaciamos el DataGridView
        If vLetras.Length < 3 Then
            DgvDescripcion.DataSource = Nothing
            DgvDescripcion.Visible = False ' <-- Esto lo "cierra" en pantalla
            Return
        End If

        ' Si hay 3 o más letras: mostramos la tabla y lanzamos la búsqueda
        DgvDescripcion.Visible = True
        BuscarLetras("descripcion")
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
                        ' =========================================================================
                        ' 🎯 ROTULACIÓN DINÁMICA DE LA CABECERA (Soporte F2 Premium)
                        ' =========================================================================
                        ' Intentamos rescatar la traducción de "F2: Forzar texto escrito" desde tu .resx
                        Dim textoCabeceraF2 As String = resManager.GetString("CabeceraF2Forzar")

                        ' Salvavidas por si acaso la clave no estuviera escrita todavía en el archivo de recursos
                        If String.IsNullOrEmpty(textoCabeceraF2) Then
                            textoCabeceraF2 = "F2: Forzar texto escrito"
                        End If

                        ' Machacamos el viejo literal "Descripción" por tu aviso de teclado inteligente
                        DgvDescripcion.Columns("DescripcionAPU").HeaderText = textoCabeceraF2
                        DgvDescripcion.Columns("DescripcionAPU").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    End If
                    DgvDescripcion.Visible = True
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

    Private Sub TxtBuscarLetras_Leave(sender As Object, e As EventArgs) Handles TxtBuscarLetras.Leave
        Dim textoActual As String = TxtBuscarLetras.Text.Trim()
        Dim textoOriginal As String = TxtBuscarLetras.Text.Trim()

        ' Si el DataGridView está invisible (lo que significa que el usuario no confirmó ninguna selección) 
        ' y además el texto actual no coincide con el original que cargó el concepto...
        If DgvDescripcion.Visible = False AndAlso textoActual.ToUpper() <> textoOriginal.ToUpper() Then

            ' 1. Apagamos momentáneamente el evento TextChanged para que no intente buscar mientras restauramos
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

            ' 2. Reestablecemos el texto original del concepto en el cuadro
            TxtBuscarLetras.Text = textoOriginal
            vDescripcion = textoOriginal ' Sincronizamos la variable global
            vIntro = "NO" ' Aseguramos que no se marque como descripción nueva

            ' 3. Volvemos a encender el evento de búsqueda
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        End If
    End Sub












    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' =========================================================================
        ' 🚀 REPARADO MODO INTEGRAL: COMPUERTAS DE SEGURIDAD ELÁSTICAS
        ' =========================================================================
        ' Solamente abortamos si la aplicación se está abriendo de fábrica por primera vez
        If cargandoFormulario AndAlso CmbConcepto.SelectedIndex < 0 Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        ' Se buscan Conceptos según lo seleccionado para mostrar su descripción y tipo
        ' **********************************************************************************
        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader)
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                ' Leemos el TipoCON de forma segura por si acaso
                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
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

                ' --- CAMBIO DINÁMICO DE ETIQUETAS SEGÚN EL SIGNO CONTABLE ---
                If tipoOriginal.ToUpper() = "GASTO" Then
                    LblNumeroPagosCobros.Text = resManager.GetString("NumeroPagos") & ":"
                    LblFechaPagoCobro.Text = resManager.GetString("1erPago") & ":"
                Else
                    LblNumeroPagosCobros.Text = resManager.GetString("NumeroCobros") & ":"
                    LblFechaPagoCobro.Text = resManager.GetString("1erCobro") & ":"
                End If

                ' --- TRADUCIR LAS DESCRIPCIONES (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción original de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                CmbDescripcion.Text = tradDesc.Trim()
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
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
            BtnAceptarSalir.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptarSalir_Click(sender As Object, e As EventArgs) Handles BtnAceptarSalir.Click
        If CmbDescripcion.Text.Trim() <> "" Then
            If TxtNumeroPagos.Text.Trim() <> "" Then
                If TxtImporte.Text.Trim() <> "" And TxtImporte.Text.Trim() <> "0" Then

                    ' 1. EXTRAEMOS LOS IDs NUMÉRICOS PUROS DESDE LOS COMBOS (Nueva era relacional)
                    Dim idConcepto As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                    Dim idCuenta As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

                    ' 2. Procesamos el número de vencimientos a generar
                    Dim numPagos As Integer = 0
                    If Not Integer.TryParse(TxtNumeroPagos.Text, numPagos) Then numPagos = 1

                    ' 3. Procesamos los importes contables en formato Decimal seguro
                    Dim importeBase As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)
                    ' Usamos ToUpper para que la validación del signo sea 100% inmune a mayúsculas/minúsculas
                    If TxtTipoConcepto.Text.ToUpper() = "GASTO" OrElse TxtTipoConcepto.Text = resManager.GetString("Tipo_Gasto") Then
                        importeBase = -Math.Abs(importeBase)
                    Else
                        importeBase = Math.Abs(importeBase)
                    End If

                    vNotasAPU = TxtNota.Text.Trim()
                    vbOK = "NO"

                    ' 🌟 EL BUCLE MAESTRO: Generamos cada uno de los vencimientos futuros
                    For i As Integer = 0 To numPagos - 1

                        ' 🌟 PROTECCIÓN DE IDIOMA: Evaluamos por SelectedIndex (posición fija) para que no falle 
                        ' si la palabra "Mensual" cambia a "Mensual" en catalán o "Monthly" en inglés
                        Select Case CmbPeriocidad.SelectedIndex
                            Case 0 ' Diaria
                                vDate3 = DateTimePicker1.Value.Date.AddDays(i)
                            Case 1 ' Semanal
                                vDate3 = DateTimePicker1.Value.Date.AddDays(i * 7)
                            Case 2 ' Quincenal
                                vDate3 = DateTimePicker1.Value.Date.AddDays(i * 15)
                            Case 3 ' Mensual
                                vDate3 = DateTimePicker1.Value.Date.AddMonths(i)
                            Case 4 ' Bimensual
                                vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 2)
                            Case 5 ' Trimestral
                                vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 3)
                            Case 6 ' Semestral
                                vDate3 = DateTimePicker1.Value.Date.AddMonths(i * 6)
                            Case 7 ' Anual
                                vDate3 = DateTimePicker1.Value.Date.AddYears(i)
                            Case Else
                                vDate3 = DateTimePicker1.Value.Date.AddMonths(i) ' Por defecto mensual
                        End Select

                        ' Forzamos a que el año del ejercicio sea el año real que le corresponde al vencimiento calculado
                        Dim anioEjercicioVencimiento As Integer = vDate3.Year

                        vAnexo = (i + 1).ToString()
                        vDescripcionAPU = CmbDescripcion.Text.Trim() & "  (" & vAnexo & " de " & numPagos.ToString() & ")"

                        ' 4. Diseñamos la estructura limpia para apuper usando comodines '?'
                        vAñadirSql = "INSERT INTO apuper (FechaAPP, ConceptoAPP, DescripcionAPP, ImporteAPP, EjercicioAPP, NotasAPP, CuentaAPP) " &
                                     "VALUES (?, ?, ?, ?, ?, ?, ?)"
                        cmdMdb1cr.CommandText = vAñadirSql
                        cmdMdb1cr.Parameters.Clear()

                        ' 5. Inyectamos los parámetros en el orden EXACTO de aparición del SQL
                        cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vDate3
                        cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConcepto       ' ID Numérico
                        cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = vDescripcionAPU
                        cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = Math.Round(importeBase, 2)
                        cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = anioEjercicioVencimiento
                        cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = vNotasAPU
                        cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuenta         ' ID Numérico

                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                            vbOK = "SI"
                        Catch ex As Exception
                            MsgBox(rmse.GetString("ErrorGrabarApuntePeriodico") & ": " & (i + 1).ToString() & vbCrLf & ex.Message, MsgBoxStyle.Critical)
                        End Try
                    Next

                    ' =========================================================================
                    ' 🌟 REFRESCAMOS LA REJILLA DE ATRÁS AUTOMÁTICAMENTE ANTES DE SALIR
                    ' =========================================================================
                    ' Reutilizamos tu rutina estrella que calcula los filtros con IDs y pinta el Grid relacional
                    If vbOK = "SI" Then
                        frmApuntesPeriodicos.RefrescarGridApuntesPeriodicos()
                    End If

                    Me.Close()
                Else
                    MsgBox(frmIntroApuntes.rmse.GetString("NoCantidadImporte"), vbExclamation)
                    TxtImporte.Select()
                End If
            Else
                MsgBox(rmse.GetString("MoHayPagosCobros"), vbExclamation)
                TxtNumeroPagos.Select()
            End If
        Else
            MsgBox(rmse.GetString("NoDescripcionVacia"), vbExclamation)
            CmbDescripcion.Select()
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
            LlenarComboConceptosPeriodicos(Me.CmbConcepto)

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
            LlenarComboCuentasPeriodicos(Me.CmbCuenta)

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

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub

    Private Sub CmbCuenta_Click(sender As Object, e As EventArgs) Handles CmbCuenta.Click
        CmbCuenta.DroppedDown = True
    End Sub

    ' =========================================================================
    ' 🚀 FUNCIÓN ACCESIBLE: LLENAR COMBO CONCEPTOS PERIODICOS
    ' =========================================================================
    Public Sub LlenarComboConceptosPeriodicos(ByVal combo As ComboBox)
        If combo Is Nothing Then Exit Sub

        ' 1. Saneamos el CommandText para que TipoCON <> 'ESPECIAL' e incluya DescripcionCON
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, DescripcionCON, TipoCON FROM conceptos WHERE TipoCON <> 'ESPECIAL'"

        Dim dtConceptos As New DataTable()
        Try
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                dtConceptos.Load(dr)
            End Using

            ' Creamos la columna virtual para el texto traducido
            dtConceptos.Columns.Add("TextoComboCON", GetType(String))

            For Each fila As DataRow In dtConceptos.Rows
                Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
                Dim textoFinal As String = codigoOriginal

                If resManager IsNot Nothing Then
                    Dim claveRecurso As String = codigoOriginal.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)
                    If Not String.IsNullOrEmpty(traduccion) Then textoFinal = traduccion
                End If
                fila("TextoComboCON") = textoFinal
            Next

            ' Ordenamos alfabéticamente por la traducción en la memoria RAM
            dtConceptos.DefaultView.Sort = "TextoComboCON ASC"

            combo.DataSource = Nothing
            combo.ValueMember = "IdConceptoCON"
            combo.DisplayMember = "TextoComboCON"
            combo.DataSource = dtConceptos.DefaultView

            ' Sincronización inteligente con la pantalla de atrás si venía filtrado
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                combo.SelectedValue = frmApuntesPeriodicos.CmbConcepto.SelectedValue
            Else
                If combo.Items.Count > 0 Then combo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' =========================================================================
    ' 🚀 FUNCIÓN ACCESIBLE: LLENAR COMBO CUENTAS PERIODICOS
    ' =========================================================================
    Public Sub LlenarComboCuentasPeriodicos(ByVal combo As ComboBox)
        If combo Is Nothing Then Exit Sub

        cmdMdb1cr.CommandText = "SELECT IdCuentaCUE, NombreCUE FROM cuentas"

        Dim dtCuentas As New DataTable()
        Try
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                dtCuentas.Load(dr)
            End Using

            dtCuentas.Columns.Add("TextoComboCUE", GetType(String))

            For Each fila As DataRow In dtCuentas.Rows
                Dim nombreOriginal As String = fila("NombreCUE").ToString().Trim()
                Dim textoFinal As String = nombreOriginal

                If resManager IsNot Nothing Then
                    Dim claveRecurso As String = nombreOriginal.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)
                    If Not String.IsNullOrEmpty(traduccion) Then textoFinal = traduccion
                End If
                fila("TextoComboCUE") = textoFinal
            Next

            ' Ordenamos alfabéticamente por la traducción en la memoria RAM
            dtCuentas.DefaultView.Sort = "TextoComboCUE ASC"

            combo.DataSource = Nothing
            combo.ValueMember = "IdCuentaCUE"
            combo.DisplayMember = "TextoComboCUE"
            combo.DataSource = dtCuentas.DefaultView

            ' Sincronización inteligente con la pantalla de atrás si venía filtrado
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                combo.SelectedValue = frmApuntesPeriodicos.CmbCuenta.SelectedValue
            Else
                If combo.Items.Count > 0 Then combo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub TxtBuscarLetras_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBuscarLetras.KeyDown
        ' =========================================================================
        ' 🌟 BOTÓN DE ESCAPE RÁPIDO F2: FORZAR DESCRIPCIÓN MANUAL (Tu nueva opción Premium)
        ' =========================================================================
        If e.KeyCode = Keys.F2 Then
            e.SuppressKeyPress = True ' Evitamos el molesto pitido de Windows
            e.Handled = True

            Dim textoBuscadoF2 As String = TxtBuscarLetras.Text.Trim()

            ' Cortafuegos básico: si el campo está vacío, pasamos de largo
            If String.IsNullOrWhiteSpace(textoBuscadoF2) Then Exit Sub

            ' Ocultamos rígidamente la tabla de sugerencias azul para simular que "No se encontraron registros"
            DgvDescripcion.Visible = False

            ' Lanzamos de forma directa exactamente la misma ventana de confirmación del CASO B
            Dim respuestaF2 As MsgBoxResult = ConfirmarAccionTraducida(
                resManager.GetString("NoExistenDescripciones") & ": -" & textoBuscadoF2.ToUpper() & "-" & vbCrLf & "¿" & resManager.GetString("AñadirDescripcion") & "?",
                rmse.GetString("$this.Text")
            )

            If respuestaF2 = MsgBoxResult.Yes Then
                vIntro = "SI"
                vDescripcion = textoBuscadoF2
                ' Actualizamos el almacén de texto original con el nuevo texto autorizado
                'TxtDescripcion.Text = textoBuscadoF2
                TxtNumeroPagos.Focus()
            Else
                TxtBuscarLetras.SelectAll()
                TxtBuscarLetras.Focus()
            End If

            Return ' Cortamos la ejecución aquí con éxito absoluto
        End If

        ' =========================================================================
        ' 1. SINTAXIS ORIGINAL DE SI EL USUARIO PRESIONA INTRO (Tu código intacto)
        ' =========================================================================
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

            ' Si la tabla está oculta, pero el texto coincide exactamente con la descripción oficial
            If DgvDescripcion.Visible = False AndAlso textoBuscado.ToUpper() = TxtBuscarLetras.Text.Trim().ToUpper() Then
                vDescripcion = textoBuscado ' Aseguramos la variable global
                vIntro = "NO" ' No es una descripción nueva
                TxtNumeroPagos.Focus() ' Saltamos directo al número de pagos
                Return
            End If

            ' CASO B: Hay texto, está oculto y NO coincide -> ¡DESCRIPCIÓN TOTALMENTE NUEVA!
            If DgvDescripcion.Visible = False OrElse DgvDescripcion.Rows.Count = 0 Then
                Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(
                    resManager.GetString("NoExistenDescripciones") & ": -" & textoBuscado.ToUpper() & "-" & vbCrLf & "¿" & resManager.GetString("AñadirDescripcion") & "?",
                    rmse.GetString("$this.Text")
                )

                If respuesta = MsgBoxResult.Yes Then
                    vIntro = "SI"
                    vDescripcion = textoBuscado
                    'TxtDescripcion.Text = textoBuscado
                    DgvDescripcion.Visible = False
                    TxtNumeroPagos.Focus()
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

        ' 2. Si el usuario presiona FLECHA ABAJO (Tu lógica original intacta)
        If e.KeyCode = Keys.Down AndAlso DgvDescripcion.Visible Then
            DgvDescripcion.Focus()
            e.Handled = True
        End If
    End Sub

End Class