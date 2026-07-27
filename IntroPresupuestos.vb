Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class IntroPresupuestos

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vFDesde, vBorrarPresu As String
    Public vMensual, vAnual, vEnero, vFebrero, vMarzo, vAbril, vMayo, vJunio, vSaldoAnualPresupuesto, vImporte As Double
    Public vJulio, vAgosto, vSeptiembre, vOctubre, vNoviembre, vDiciembre As Double
    Public TL(18) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroPresupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 🌟 PASO CRÍTICO 1: Encendemos el escudo de carga para congelar eventos automáticos
        cargandoFormulario = True
        Me.KeyPreview = True
        cmdMdb1cr.Parameters.Clear()

        ' Optimización: Bucle dinámico para las etiquetas dentro del GroupBox de meses (Tu lógica perfecta)
        For i As Integer = 16 To 28
            Dim lbl() As Control = Me.Controls.Find("Label" & i, True)
            If lbl.Length > 0 Then lbl(0).Text = vMoneda
        Next

        ' Inicialización centralizada de ToolTips (Mantenida tu excelente lógica .NET de fábrica)
        Dim controlesToolTip As Control() = {
            BtnConcepto, BtnAceptar, BtnCancelar, CmbConcepto, TxtAnual,
            TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
            TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre,
            RdbAnual, RdbMensual
        }

        Dim clavesToolTip As String() = {
            "AñadeEditaBorraCC", "BtnAceptar.Text", "BtnCancelar.Text", "SeleccionarConcepto", "TotalAnual",
            "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual",
            "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual",
            "RadioAnual", "RadioMensual"
        }

        For i As Integer = 0 To TL.Length - 1
            TL(i) = New ToolTip()
            TL(i).SetToolTip(controlesToolTip(i), rmse.GetString(clavesToolTip(i)))
        Next

        ' =========================================================================
        ' 🌟 CORTAFUEGOS DE CARGA: ABRIMOS LA COMPUERTA INTERNACIONAL DE LA RAM
        ' =========================================================================
        ' Apagamos el escudo aquí para que el vaivén del combo pueda disparar 
        ' el evento SelectedIndexChanged y traducir los textos en curso al vuelo.
        cargandoFormulario = False

        ' 🌟 LLENAR EL COMBO CONCEPTO (Internacionalizado, con IDs y Orden A-Z)
        Try
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' =========================================================================
            ' 🌟 SINCRO INTELIGENTE BLINDADA CONTRA NULLREFERENCE (Escudo Nueva Era)
            ' =========================================================================
            If frmPresupuestos IsNot Nothing AndAlso frmPresupuestos.IsHandleCreated Then

                ' Si la pantalla nodriza está filtrada, heredamos su selección exacta
                If frmPresupuestos.BtnFiltroConcepto.Enabled = False Then
                    ' 🚀 VAIVÉN INTERNACIONAL: Forzamos el -1 y disparamos la traducción por su ID
                    CmbConcepto.SelectedIndex = -1
                    CmbConcepto.SelectedValue = frmPresupuestos.CmbConcepto.SelectedValue
                Else
                    ' 🚀 VAIVÉN LOCAL: Pasamos por -1 y despertamos el índice 0 en vivo
                    CmbConcepto.SelectedIndex = -1
                    If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
                End If

            Else
                ' 🧰 PLAN B (Carga Aislada Segura): Vaivén limpio si la pantalla de atrás duerme
                CmbConcepto.SelectedIndex = -1
                If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorIniciarConceptos") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' =========================================================================
        ' COMPORTAMIENTO INICIAL DE LOS MANDOS (Tu excelente lógica de fábrica)
        ' =========================================================================
        If RdbAnual.Checked = True Then
            GBoxAnual.Enabled = True
            GBoxMensual.Enabled = False
            TxtAnual.Select()
            TxtAnual.SelectAll()
        End If

        If RdbMensual.Checked = True Then
            GBoxAnual.Enabled = False
            GBoxMensual.Enabled = True
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If

        ' Llamamos a tu macro que rellena los cuadros de texto mensuales heredados
        LlenarTextBox()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Si el formulario se está iniciando o limpiando, salimos de inmediato
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim idConceptoSel As Integer = 0
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access, cero bloqueos)
            ' Convertimos el ítem seleccionado en un DataRowView para leer sus columnas ocultas en la RAM
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                idConceptoSel = Convert.ToInt32(filaSeleccionada("IdConceptoCON"))
                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            If idConceptoSel > 0 Then
                ' Sincronizamos tus variables globales de fábrica
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

                ' --- TRADUCIR LAS DESCRIPCIONES AUTOMÁTICAS (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción genérica de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal
                TxtDescripcion.Text = tradDesc

                ' 2. Enfocamos la caja correspondiente según la selección (Tu lógica impecable)
                If RdbAnual.Checked = True Then
                    TxtAnual.Select()
                    TxtAnual.SelectAll()
                Else
                    TxtEnero.Select()
                    TxtEnero.SelectAll()
                End If

                ' 3. Rellenamos las 12 cajas mensuales con lo que haya en los presupuestos
                LlenarTextBox()
            End If

        Catch ex As Exception
            ' Evita cuelgues visuales si el combo parpadea en la carga
        End Try
    End Sub

    Public Sub LlenarTextBox()
        ' 1. Ponemos todas las cajas a cero por defecto usando un array para no repetir líneas (Tu lógica perfecta)
        Dim cajasMeses As TextBox() = {TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
                                       TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre}

        For Each txt In cajasMeses
            txt.Text = "0,00"
        Next
        TxtAnual.Text = "0,00"

        ' Array local para almacenar y comparar los 12 meses en memoria
        Dim importesMensuales(11) As Double
        Dim sumaAnual As Double = 0

        ' 🌟 CORRECCIÓN CRÍTICA: Capturamos el ID numérico real oculto del combo en vez de la variable de texto
        Dim idConceptoActual As Integer = 0
        If CmbConcepto.SelectedValue IsNot Nothing Then
            idConceptoActual = Convert.ToInt32(CmbConcepto.SelectedValue)
        End If

        ' Si no hay ningún concepto seleccionado de verdad, abortamos la carga para evitar fallos
        If idConceptoActual = 0 Then Exit Sub

        ' 2. Consulta SQL parametrizada sobre tu estructura relacional con comodines '?'
        vtipoSql = "SELECT ImportePRE, FDesdePRE FROM presupuesto WHERE EjercicioPRE = ? AND ConceptoPRE = ?"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using cmd As New OleDbCommand(vtipoSql, conexion)
                cmd.Parameters.Clear()

                ' Los parámetros de Access se asocian estrictamente por el orden de los comodines '?'
                cmd.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
                cmd.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoActual ' Inyectamos el ID entero puro

                Try
                    conexion.Open()
                    Using dr As OleDbDataReader = cmd.ExecuteReader()

                        ' Recorremos los registros que existan en tu MDB para este presupuesto
                        While dr.Read()
                            ' Forzamos una conversión limpia y segura inmune al signo contable
                            Dim importe As Double = Math.Abs(Convert.ToDouble(dr("ImportePRE")))
                            Dim fecha As Date = Convert.ToDateTime(dr("FDesdePRE"))
                            Dim mes As Integer = fecha.Month ' Extrae el número de mes (1 al 12)

                            sumaAnual += importe

                            ' Guardamos en el array y asignamos a la caja correspondiente de forma automática
                            If mes >= 1 AndAlso mes <= 12 Then
                                importesMensuales(mes - 1) = importe
                                cajasMeses(mes - 1).Text = importe.ToString("N2")
                            End If
                        End While

                        ' Mostramos la suma total acumulada en la casilla Anual
                        TxtAnual.Text = sumaAnual.ToString("N2")

                        ' 3. DETECTAR AUTOMÁTICAMENTE SI ERA REPARTO ANUAL O MENSUAL (Tu excelente lógica intacta)
                        Dim todosIguales As Boolean = True
                        Dim primerImporte As Double = importesMensuales(0)

                        For i As Integer = 1 To 11
                            If importesMensuales(i) <> primerImporte Then
                                todosIguales = False
                                Exit For
                            End If
                        Next

                        ' Desvinculamos temporalmente los eventos para que el cambio de RadioButton no limpie los TextBox
                        RemoveHandler RdbAnual.CheckedChanged, AddressOf RdbAnual_CheckedChanged
                        RemoveHandler RdbMensual.CheckedChanged, AddressOf RdbAnual_CheckedChanged

                        ' Si todos los meses son iguales y el presupuesto no está vacío, es Anual. Si no, Mensual.
                        If todosIguales AndAlso sumaAnual > 0 Then
                            RdbAnual.Checked = True
                            GBoxAnual.Enabled = True
                            GBoxMensual.Enabled = False
                        Else
                            RdbMensual.Checked = True
                            GBoxAnual.Enabled = False
                            GBoxMensual.Enabled = True
                        End If

                        ' Volvemos a activar los escuchadores de los RadioButtons
                        AddHandler RdbAnual.CheckedChanged, AddressOf RdbAnual_CheckedChanged
                        AddHandler RdbMensual.CheckedChanged, AddressOf RdbAnual_CheckedChanged

                    End Using
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorImportesMensuales") & ": " & ex.Message, MsgBoxStyle.Critical)
                End Try
            End Using
        End Using
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Validar que tengamos un concepto contable seleccionado de forma segura
        Dim idConcepto As Integer = 0
        If CmbConcepto.SelectedValue IsNot Nothing Then
            idConcepto = Convert.ToInt32(CmbConcepto.SelectedValue)
        End If

        If idConcepto = 0 Then
            MessageBox.Show(rmse.GetString("SeleccionarCC"), resManager.GetString("Aviso"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' 2. Mapeamos los valores de las cajas de texto a un array numérico en memoria (Tu excelente reparto de céntimos)
        Dim importesMensuales(11) As Double

        If RdbAnual.Checked Then
            Dim totalAnual As Double = 0
            Double.TryParse(TxtAnual.Text, totalAnual)
            Dim totalDecimal As Decimal = Convert.ToDecimal(totalAnual)
            Dim importeRepartido As Decimal = Math.Round(totalDecimal / 12D, 2)

            Dim acumuladoPrimerosMeses As Decimal = 0.0D
            For i As Integer = 0 To 10
                importesMensuales(i) = Convert.ToDouble(importeRepartido)
                acumuladoPrimerosMeses += importeRepartido
            Next

            Dim ultimoMesDecimal As Decimal = totalDecimal - acumuladoPrimerosMeses
            importesMensuales(11) = Convert.ToDouble(ultimoMesDecimal)
        Else
            ' Si es mensual, parseamos cada una de las 12 cajas de tu formulario de siempre
            Double.TryParse(TxtEnero.Text, importesMensuales(0))
            Double.TryParse(TxtFebrero.Text, importesMensuales(1))
            Double.TryParse(TxtMarzo.Text, importesMensuales(2))
            Double.TryParse(TxtAbril.Text, importesMensuales(3))
            Double.TryParse(TxtMayo.Text, importesMensuales(4))
            Double.TryParse(TxtJunio.Text, importesMensuales(5))
            Double.TryParse(TxtJulio.Text, importesMensuales(6))
            Double.TryParse(TxtAgosto.Text, importesMensuales(7))
            Double.TryParse(TxtSeptiembre.Text, importesMensuales(8))
            Double.TryParse(TxtOctubre.Text, importesMensuales(9))
            Double.TryParse(TxtNoviembre.Text, importesMensuales(10))
            Double.TryParse(TxtDiciembre.Text, importesMensuales(11))
        End If

        ' 3. GRABACIÓN SEGURA EN LA MDB ACTUAL DE LOS USUARIOS (Con Transacción Limpia)
        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Try
                conexion.Open()

                Using transaccion As OleDbTransaction = conexion.BeginTransaction()

                    ' 🌟 PLAN A: Limpiamos cualquier presupuesto anterior que tuviera este ID numérico de concepto en este año
                    ' Usamos los comodines '?' puros en el orden biológico correcto de Access
                    Dim sqlDelete As String = "DELETE FROM presupuesto WHERE ConceptoPRE = ? AND EjercicioPRE = ?"
                    Using cmdDelete As New OleDbCommand(sqlDelete, conexion, transaccion)
                        cmdDelete.Parameters.Clear()
                        cmdDelete.Parameters.Add("@con", OleDbType.Integer).Value = idConcepto
                        cmdDelete.Parameters.Add("@eje", OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
                        cmdDelete.ExecuteNonQuery()
                    End Using

                    ' 🌟 PLAN B: Inserción masiva de las 12 mensualidades relacionales por IDs
                    Dim sqlInsert As String = "INSERT INTO presupuesto (ConceptoPRE, ImportePRE, EjercicioPRE, FDesdePRE) VALUES (?, ?, ?, ?)"
                    Using cmdInsert As New OleDbCommand(sqlInsert, conexion, transaccion)

                        ' Declaramos parámetros con tipos estrictos fijos para el motor relacional
                        cmdInsert.Parameters.Clear()
                        cmdInsert.Parameters.Add("@con", OleDbType.Integer) ' 🌟 Cambiado a Integer para recibir el ID numérico
                        cmdInsert.Parameters.Add("@imp", OleDbType.Double)
                        cmdInsert.Parameters.Add("@eje", OleDbType.Integer)
                        cmdInsert.Parameters.Add("@fec", OleDbType.Date)

                        ' Ejecutamos el bucle para los 12 meses del año
                        For mes As Integer = 1 To 12
                            Dim fechaMes As New Date(Convert.ToInt32(vAñoEjercicio), mes, 1)

                            cmdInsert.Parameters(0).Value = idConcepto
                            cmdInsert.Parameters(1).Value = importesMensuales(mes - 1)
                            cmdInsert.Parameters(2).Value = Convert.ToInt32(vAñoEjercicio)
                            cmdInsert.Parameters(3).Value = fechaMes

                            cmdInsert.ExecuteNonQuery()
                        Next
                    End Using

                    ' Si todo ha ido bien sin errores, consolidamos los cambios en el archivo físico .mdb
                    transaccion.Commit()

                    ' =========================================================================
                    ' 🌟 REFRESCAMOS LA REJILLA DE ATRÁS AUTOMÁTICAMENTE ANTES DE SALIR
                    ' =========================================================================
                    ' Forzamos a la pantalla principal a ejecutar su macro-consulta relacional con INNER JOIN
                    ' para que los 12 meses aparezcan listados en unas mayúsculas perfectas de inmediato
                    If TypeOf frmPresupuestos Is Form Then
                        frmPresupuestos.BtnF6.PerformClick() ' 🚀 El atajo F6 limpia filtros y recarga todo el año relacional
                    End If

                    ' Mensaje de éxito original impecable
                    Dim msgExito As String = rmse.GetString("PresupuestoGuardado")
                    If String.IsNullOrEmpty(msgExito) Then msgExito = "Presupuesto guardado correctamente."
                    MessageBox.Show(msgExito, resManager.GetString("Exito"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Me.Close()
                End Using
            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorGrabarRegistro") & ": " & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub TxtAnual_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtAnual.Validating
        ' Solo actuamos si está seleccionada la opción de reparto Anual
        If RdbAnual.Checked Then
            Dim totalAnual As Double = 0
            ' Convertimos el texto a número de forma segura
            If Double.TryParse(TxtAnual.Text.Trim(), totalAnual) Then
                ' Dividimos entre 12 y redondeamos a 2 decimales
                ' 1. Convertimos la variable primero y luego dividimos por 12 en formato Decimal (12D)
                Dim importeMensual As Decimal = Math.Round(Convert.ToDecimal(totalAnual) / 12D, 2)
                ' 2. Formateamos a texto para la pantalla respetando los puntos y comas del usuario
                Dim textoFormateado As String = importeMensual.ToString("N2")

                ' Rellenamos las 12 cajas mensuales visualmente
                TxtEnero.Text = textoFormateado : TxtFebrero.Text = textoFormateado
                TxtMarzo.Text = textoFormateado : TxtAbril.Text = textoFormateado
                TxtMayo.Text = textoFormateado : TxtJunio.Text = textoFormateado
                TxtJulio.Text = textoFormateado : TxtAgosto.Text = textoFormateado
                TxtSeptiembre.Text = textoFormateado : TxtOctubre.Text = textoFormateado
                TxtNoviembre.Text = textoFormateado : TxtDiciembre.Text = textoFormateado

                ' Reajustamos el total anual por si el redondeo de decimales varió un céntimo
                TxtAnual.Text = (importeMensual * 12).ToString("N2")
            End If
        End If
    End Sub

    Private Sub CalcularSumaMensualidades()
        ' Solo actuamos si está seleccionada la opción de introducción Mensual
        If RdbMensual.Checked Then
            Dim sumaAcumulada As Decimal = 0

            ' 1. Creamos un array en caliente con tus 12 cajas mensuales del formulario
            Dim cajasMeses As TextBox() = {TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
                                           TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre}

            ' 2. El bucle maestro recorre las cajas y acumula los importes de forma 100% segura
            For Each txt In cajasMeses
                Dim valorCaja As Decimal = 0
                ' ConvertirDecimalSeguro limpia los puntos y comas según el idioma regional del Windows del usuario
                If Decimal.TryParse(txt.Text.Trim(), valorCaja) Then
                    sumaAcumulada += valorCaja
                End If
            Next

            ' 3. Mostramos el resultado totalizado con formato contable de dos decimales
            TxtAnual.Text = sumaAcumulada.ToString("N2")
        End If
    End Sub

    ' Función auxiliar rápida para acumular los valores
    Private Sub PointToSuma(ByRef total As Double, valor As Double)
        total += valor
    End Sub

    ' Enlazamos las 12 cajas al mismo evento para ahorrar código (Tu excelente arquitectura)
    Private Sub TxtMeses_Leave(sender As Object, e As EventArgs) Handles _
    TxtEnero.Leave, TxtFebrero.Leave, TxtMarzo.Leave, TxtAbril.Leave,
    TxtMayo.Leave, TxtJunio.Leave, TxtJulio.Leave, TxtAgosto.Leave,
    TxtSeptiembre.Leave, TxtOctubre.Leave, TxtNoviembre.Leave, TxtDiciembre.Leave

        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Si la pantalla está inyectando datos desde el Load, pasamos de largo
        If cargandoFormulario Then Exit Sub

        Dim txt As TextBox = CType(sender, TextBox)
        Dim valor As Decimal = 0

        ' 🌟 CORRECCIÓN DE PRECISIÓN: Damos formato contable exacto usando Decimal
        If Decimal.TryParse(txt.Text.Trim(), valor) Then
            txt.Text = valor.ToString("N2")
        Else
            txt.Text = "0,00"
        End If

        ' Recalculamos el total anual reflejado en la pantalla de forma legal y segura
        CalcularSumaMensualidades()
    End Sub

    Private Sub TxtMeses_Enter(sender As Object, e As EventArgs) Handles _
    TxtEnero.Enter, TxtFebrero.Enter, TxtMarzo.Enter, TxtAbril.Enter,
    TxtMayo.Enter, TxtJunio.Enter, TxtJulio.Enter, TxtAgosto.Enter,
    TxtSeptiembre.Enter, TxtOctubre.Enter, TxtNoviembre.Enter, TxtDiciembre.Enter

        If cargandoFormulario Then Exit Sub

        Dim txt As TextBox = CType(sender, TextBox)
        Dim valor As Decimal = 0

        ' Al entrar, quitamos los puntos de millar para facilitar la escritura manual sin perder precisión
        If Decimal.TryParse(txt.Text.Trim(), valor) Then
            If valor = 0 Then
                txt.Text = "" ' Si es cero, vaciamos la caja para que no tenga que borrar el "0,00"
            Else
                txt.Text = valor.ToString("F2") ' Formato limpio sin separador de miles (ej: 1250.00)
            End If
        End If
        txt.SelectAll()
    End Sub

    Private Sub RdbAnual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbAnual.CheckedChanged, RdbMensual.CheckedChanged
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Evita disparos en falso que vuelven loco al procesador durante el arranque
        If cargandoFormulario Then Exit Sub

        ' 1. Habilitamos o deshabilitamos los contenedores visuales según el RadioButton activo
        GBoxAnual.Enabled = RdbAnual.Checked
        GBoxMensual.Enabled = RdbMensual.Checked

        ' 2. Lógica específica al activar cada opción
        If RdbAnual.Checked Then
            ' Si pasa a ANUAL, ponemos el foco en el Total Anual para que defina la nueva cifra macro
            TxtAnual.Select()
            TxtAnual.SelectAll()
        ElseIf RdbMensual.Checked Then
            ' Si pasa a MENSUAL, calculamos la suma de lo que ya haya en las cajas mensuales 
            ' para que el Total Anual refleje la realidad actual de los meses inmediatamente.
            CalcularSumaMensualidades()

            ' Llevamos el foco al primer mes del año para que empiece a editar cómodamente
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtAnual_Click(sender As Object, e As EventArgs) Handles TxtAnual.Click
        TxtAnual.Select()
        TxtAnual.SelectAll()
    End Sub

    Private Sub TxtAnual_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAnual.KeyPress
        ' 1. 🛡️ EL ESCUDO UNIVERSAL ADMITE TODO: Números, borrar, punto, coma o el Intro
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> ","c AndAlso e.KeyChar <> ChrW(Keys.Enter) Then
            e.Handled = True
            Exit Sub
        End If

        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True

            ' 2. 🎯 EL PARSEO INVARIANTE INMUNE A IDIOMAS: Prosa limpia de fábrica
            Dim texto As String = TxtAnual.Text.Trim().Replace(",", ".")
            Dim importeAnualPure As Decimal = 0

            ' Formateo universal absoluto (Estilo Internacional de Redmond)
            Dim estilo As System.Globalization.NumberStyles = System.Globalization.NumberStyles.AllowDecimalPoint Or System.Globalization.NumberStyles.AllowThousands
            Decimal.TryParse(texto, estilo, System.Globalization.CultureInfo.InvariantCulture, importeAnualPure)

            vAnual = Convert.ToDouble(importeAnualPure)
            TxtAnual.Text = importeAnualPure.ToString("N2")

            ' (Tu bloque clásico intacto del reparto mensual)
            Dim importeRepartido As Decimal = Math.Round(importeAnualPure / 12D, 2)
            Dim acumuladoPrimerosMeses As Decimal = 0.0D
            Dim importesMensuales(11) As Double
            For i As Integer = 0 To 10
                importesMensuales(i) = Convert.ToDouble(importeRepartido)
                acumuladoPrimerosMeses += importeRepartido
            Next
            importesMensuales(11) = Convert.ToDouble(importeAnualPure - acumuladoPrimerosMeses)
            vEnero = importesMensuales(0) : vFebrero = importesMensuales(1) : vMarzo = importesMensuales(2) : vAbril = importesMensuales(3)
            vMayo = importesMensuales(4) : vJunio = importesMensuales(5) : vJulio = importesMensuales(6) : vAgosto = importesMensuales(7)
            vSeptiembre = importesMensuales(8) : vOctubre = importesMensuales(9) : vNoviembre = importesMensuales(10) : vDiciembre = importesMensuales(11)
            Dim cajasMeses As TextBox() = {TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio, TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre}
            For idx As Integer = 0 To 11
                cajasMeses(idx).Text = importesMensuales(idx).ToString("N2")
            Next
            RdbMensual.Select()
        End If
    End Sub

    Private Sub TxtMeses_KeyPress(sender As Object, e As KeyPressEventArgs) Handles _
    TxtEnero.KeyPress, TxtFebrero.KeyPress, TxtMarzo.KeyPress, TxtAbril.KeyPress,
    TxtMayo.KeyPress, TxtJunio.KeyPress, TxtJulio.KeyPress, TxtAgosto.KeyPress,
    TxtSeptiembre.KeyPress, TxtOctubre.KeyPress, TxtNoviembre.KeyPress, TxtDiciembre.KeyPress

        ' 1. 🛡️ EL ESCUDO UNIVERSAL MULTI-CAJA
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> ","c AndAlso e.KeyChar <> ChrW(Keys.Enter) Then
            e.Handled = True
            Exit Sub
        End If

        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True

            Dim txt As TextBox = CType(sender, TextBox)

            ' 2. 🎯 EL PARSEO INVARIANTE EN LOS MESES
            Dim texto As String = txt.Text.Trim().Replace(",", ".")
            Dim valorIngresado As Decimal = 0

            Dim estilo As System.Globalization.NumberStyles = System.Globalization.NumberStyles.AllowDecimalPoint Or System.Globalization.NumberStyles.AllowThousands
            Decimal.TryParse(texto, estilo, System.Globalization.CultureInfo.InvariantCulture, valorIngresado)

            txt.Text = valorIngresado.ToString("N2")

            Dim numDouble As Double = Convert.ToDouble(valorIngresado)
            Select Case txt.Name
                Case "TxtEnero" : vEnero = numDouble : TxtFebrero.Select()
                Case "TxtFebrero" : vFebrero = numDouble : TxtMarzo.Select()
                Case "TxtMarzo" : vMarzo = numDouble : TxtAbril.Select()
                Case "TxtAbril" : vAbril = numDouble : TxtMayo.Select()
                Case "TxtMayo" : vMayo = numDouble : TxtJunio.Select()
                Case "TxtJunio" : vJunio = numDouble : TxtJulio.Select()
                Case "TxtJulio" : vJulio = numDouble : TxtAgosto.Select()
                Case "TxtAgosto" : vAgosto = numDouble : TxtSeptiembre.Select()
                Case "TxtSeptiembre" : vSeptiembre = numDouble : TxtOctubre.Select()
                Case "TxtOctubre" : vOctubre = numDouble : TxtNoviembre.Select()
                Case "TxtNoviembre" : vNoviembre = numDouble : TxtDiciembre.Select()
                Case "TxtDiciembre" : vDiciembre = numDouble : BtnAceptar.Select()
            End Select

            Dim sumaDecimal As Decimal = Convert.ToDecimal(vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre)
            vAnual = Convert.ToDouble(sumaDecimal)
            TxtAnual.Text = sumaDecimal.ToString("N2")
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class