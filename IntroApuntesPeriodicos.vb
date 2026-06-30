Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntesPeriodicos

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid As String
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
        TL(6).SetToolTip(Me.CmbDescripcion, frmEditarApuntes.rmse.GetString("ToolTipSeleccionarDescripcion"))
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

        ' =========================================================================
        ' 🌟 CARGA DE COMBOS DE LA NUEVA ERA (Con Idiomas, Orden A-Z e IDs Numéricos)
        ' =========================================================================

        ' 1. Llenar el Combo Concepto (Excluyendo 'TRASPASO', Traducido y Ordenado A-Z)
        ' -----------------------------------------------------------------------------
        ' Leemos el IdConceptoCON y el CodigoCON para amarrar la base de datos relacional
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON FROM conceptos WHERE CodigoCON <> 'TRASPASO'"

        Dim dtConceptos As New DataTable()
        Try
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                dtConceptos.Load(dr) ' Absorbe los datos en bloque y libera a Access al instante
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

            CmbConcepto.DataSource = Nothing
            CmbConcepto.ValueMember = "IdConceptoCON"   ' ID numérico oculto para guardar en la BD
            CmbConcepto.DisplayMember = "TextoComboCON"  ' Texto traducido visible en orden A-Z
            CmbConcepto.DataSource = dtConceptos.DefaultView

            ' Aplicamos tu sincronización inteligente con la pantalla de atrás si venía filtrado
            If frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Then
                CmbConcepto.SelectedValue = frmApuntesPeriodicos.CmbConcepto.SelectedValue
            Else
                If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' 2. Llenar el Combo Descripción (Optimizado a la velocidad del rayo con DISTINCT)
        ' -----------------------------------------------------------------------------
        ' Delegamos el descarte de duplicados en Access en un milisegundo, eliminando el bucle pesado
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'Saldo Inicial' And DescripcionAPU Is Not Null ORDER BY DescripcionAPU ASC"

        CmbDescripcion.Items.Clear()
        Try
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While dr.Read()
                    Dim descLimpia As String = dr("DescripcionAPU").ToString().Trim()
                    If Not String.IsNullOrEmpty(descLimpia) Then
                        CmbDescripcion.Items.Add(descLimpia)
                    End If
                End While
            End Using
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' 3. Llenar el Combo Cuenta (Traducido por resManager y Ordenado A-Z)
        ' -----------------------------------------------------------------------------
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

            CmbCuenta.DataSource = Nothing
            CmbCuenta.ValueMember = "IdCuentaCUE"    ' ID numérico oculto para guardar en la BD
            CmbCuenta.DisplayMember = "TextoComboCUE" ' Texto traducido visible en orden A-Z
            CmbCuenta.DataSource = dtCuentas.DefaultView

            ' Sincronización inteligente con la pantalla de atrás si venía filtrado
            If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Then
                CmbCuenta.SelectedValue = frmApuntesPeriodicos.CmbCuenta.SelectedValue
            Else
                If CmbCuenta.Items.Count > 0 Then CmbCuenta.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

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

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Si el formulario se está cargando o limpiando, salimos de inmediato
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access)
            ' Convertimos el ítem seleccionado en un DataRowView para leer sus columnas ocultas
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            If Not String.IsNullOrEmpty(codigoOriginal) Then
                ' Almacenamos el código en tu variable global para tus lógicas de fábrica
                vConcepto = codigoOriginal

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
                ' Evaluamos con el tipo original en mayúsculas para que sea 100% inmune al idioma activo
                If tipoOriginal.ToUpper() = "GASTO" Then
                    LblNumeroPagosCobros.Text = "Nº de Pagos:"
                    LblFechaPagoCobro.Text = "1er Pago:"
                Else
                    LblNumeroPagosCobros.Text = "Nº de Cobros:"
                    LblFechaPagoCobro.Text = "1er Cobro:"
                End If

                ' --- TRADUCIR LAS DESCRIPCIONES AUTOMÁTICAS (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción genérica de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal
                CmbDescripcion.Text = tradDesc
            End If

        Catch ex As Exception
            ' Evita cuelgues visuales si el combo parpadea en la carga
        End Try
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs)
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
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
        frmPrincipal.TsLabelFormulario.Text = frmConceptosContables.rmse.GetString("$this.Text")

        ' Comprobamos si existe un identificador asociado.
        If ((frmConceptosContables Is Nothing) OrElse (Not frmConceptosContables.IsHandleCreated)) Then
            frmConceptosContables = New ConceptosContables
        End If

        ' Llamamos al formulario de manera modal.
        frmConceptosContables.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmConceptosContables.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs) Handles BtnCuenta.Click
        frmPrincipal.TsLabelFormulario.Text = frmCuentasBancarias.rmse.GetString("$this.Text")

        ' Comprobamos si existe un identificador asociado.
        If ((frmCuentasBancarias Is Nothing) OrElse (Not frmCuentasBancarias.IsHandleCreated)) Then
            frmCuentasBancarias = New CuentasBancarias
        End If

        ' Llamamos al formulario de manera modal.
        frmCuentasBancarias.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmCuentasBancarias.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
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
End Class