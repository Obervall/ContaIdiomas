Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class Presupuestos

    Private cargandoFormulario As Boolean = True
    Public vtipoSql, vtipoGrid, vConcepto, vAñadir, vAñadir2 As String
    Public vTmpprint As String
    Public PrintLine, Contador, FilaSelec As Integer
    Public vTipoConceptoActual As String = ""
    Public TL() As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub Presupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' 🌟 PASO CRÍTICO 1: Encendemos el escudo de carga para congelar eventos automáticos
        cargandoFormulario = True
        Me.KeyPreview = True

        ' Inicialización centralizada de ToolTips (Mantenida tu excelente lógica .NET)
        ' *========================================================================
        Dim controlesToolTip As Control() = {
            BtnGraficos2D, BtnSalir, BtnFiltroConcepto, BtnSinFiltroConcepto, BtnImprimir,
            BtnPrimero, BtnAnterior, BtnSiguiente, BtnUltimo, BtnEliminarRegistro, BtnGraficos3D,
            BtnEliminaSeleccion, BtnF6
        }

        Dim clavesToolTip As String() = {
            "ToolTipGraficos2D", "ToolTipSalir", "ToolTipAplicarFiltro", "ToolTipQuitarFiltro", "ToolTipImprimir",
            "ToolTipPrimero", "ToolTipAnterior", "ToolTipSiguiente", "ToolTipUltimo", "ToolTipEliminarPresupuesto", "ToolTipGraficos3D",
            "ToolTipEliminaSeleccion", "ToolTipF6"
        }

        ReDim TL(controlesToolTip.Length - 1)

        For i As Integer = 0 To controlesToolTip.Length - 1
            TL(i) = New ToolTip()
            TL(i).SetToolTip(controlesToolTip(i), resManager.GetString(clavesToolTip(i)))
        Next

        ' =========================================================================
        ' 🌟 1. LLENAR EL COMBO CONCEPTO (Internacionalizado, con IDs y Orden A-Z)
        ' =========================================================================
        ' Reutilizamos de forma magistral la función exclusiva para combos sueltos que creamos ayer
        Try
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' Selección inicial por defecto del primer elemento
            If CmbConcepto.Items.Count > 0 Then CmbConcepto.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' =========================================================================
        ' 🌟 2. LIMPIEZA RADICAL DE CONCEPTOS VACÍOS (Parametrizada y Segura)
        ' =========================================================================
        Dim sqlLimpieza As String = "DELETE FROM presupuesto WHERE EjercicioPRE = ? AND ImportePRE = 0"
        Using cmdLimpiar As New OleDbCommand(sqlLimpieza, conexion1)
            cmdLimpiar.Parameters.Clear()
            cmdLimpiar.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
            Try
                If conexion1.State <> ConnectionState.Open Then conexion1.Open()
                cmdLimpiar.ExecuteNonQuery()
            Catch ex As Exception
                ' Fallo silencioso por seguridad si la tabla está bloqueada
            End Try
        End Using

        ' =========================================================================
        ' 🌟 CONSULTA SQL DE LA NUEVA ERA REALINEADA A TU DISEÑO ORIGINAL
        ' =========================================================================
        'Posición 0 y 1: El concepto corto en mayúsculas (conceptos.CodigoCON).
        'Posición 2 y 3: El importe económico (presupuesto.ImportePRE).
        'Posición 4: La fecha real (presupuesto.FDesdePRE) que usa la macro para calcular el MonthName [1.1].
        'Posición 5(El comodín): Repetimos presupuesto.FDesdePRE para cubrir el hueco del Autonumérico ausente [1.1].
        'Posición 6: El Id numérico entero del concepto (presupuesto.ConceptoPRE), vital para que la macro calcule el Saldo Real parametrizado sin lanzar fallos de tipo [1.1].
        'Posición 7: El código original en castellano (conceptos.CodigoCON) para que funcione el motor de traducciones [1.1].
        vtipoSql = "SELECT conceptos.CodigoCON, " &
                    "conceptos.CodigoCON, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.ConceptoPRE, " &
                    "conceptos.CodigoCON " &
                    "FROM presupuesto " &
                    "INNER JOIN conceptos ON presupuesto.ConceptoPRE = conceptos.IdConceptoCON " &
                    "WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString & " "
        ' Si estás llamando a esta consulta desde el método EjecutarCalculoYDesviacion, le pegas el filtro aquí:
        ' vtipoSql += $"And presupuesto.ConceptoPRE = {idConceptoSel} "
        vtipoSql += "ORDER BY conceptos.CodigoCON ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' Volcamos los datos limpios y forzamos la traducción internacional
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridPresupuestos(Me.DgvPresupuestos)

        ' Ocultamos etiquetas de desviación originales impecables
        LblDesviacion.Visible = False
        Label2.Visible = False
        LblMontoDesviacion.Visible = False
        LblObjetivo.Visible = False

        ' 🌟 PASO CRÍTICO 2: Apagamos el escudo. El formulario ya está cargado en la RAM
        cargandoFormulario = False

        ' =========================================================================
        ' 🌟 LA CORRECCIÓN MAESTRA: SINCRO INICIAL DE LA DESCRIPCIÓN (.resx)
        ' =========================================================================
        ' Forzamos un vaivén en el índice del combo para obligar a .NET a despertar 
        ' el evento SelectedIndexChanged y que pinte la descripción en el arranque.
        If CmbConcepto.Items.Count > 0 Then
            CmbConcepto.SelectedIndex = -1 ' Lo movemos a vacío temporalmente
            CmbConcepto.SelectedIndex = 0  ' Lo devolvemos al primer concepto
        End If
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' 1. Restauramos el estado estético de la botonería
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 CONSULTA SQL MAESTRA RELACIONAL (Quitar Filtro con nombres limpios)
        ' =========================================================================
        ' Cruzamos presupuesto con conceptos para restaurar la vista totalitaria del año
        vtipoSql = "SELECT conceptos.CodigoCON, " &
                    "conceptos.CodigoCON, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.ConceptoPRE, " &
                    "conceptos.CodigoCON " &
                    "FROM presupuesto " &
                    "INNER JOIN conceptos ON presupuesto.ConceptoPRE = conceptos.IdConceptoCON " &
                    "WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString & " "
        ' Si estás llamando a esta consulta desde el método EjecutarCalculoYDesviacion, le pegas el filtro aquí:
        ' vtipoSql += $"And presupuesto.ConceptoPRE = {idConceptoSel} "
        vtipoSql += "ORDER BY conceptos.CodigoCON ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' Volcamos los datos limpios en la cuadrícula y aplicamos idiomas
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridPresupuestos(Me.DgvPresupuestos)

        ' =========================================================================
        ' TU LÓGICA DE ETIQUETAS CONTABLES ORIGINALES (Mantenida intacta)
        ' =========================================================================
        ' Evaluamos si corresponde "Parcial" o "Anual" según tus macros de cálculo
        ActualizarEtiquetaDesviacion()

        ' Al no haber filtro por concepto único, ocultamos la desviación macro por seguridad
        LblDesviacion.Enabled = False
        LblObjetivo.Visible = False
        LblMontoDesviacion.Text = ""

        ' =========================================================================
        ' 🌟 EL ESCUDO INDESTRUCTIBLE DE ENERO (Cortafuegos al quitar filtro)
        ' =========================================================================
        If Me.DgvPresupuestos.Rows.Count > 0 Then
            Dim vFechaFilaZero As Date
            If Me.DgvPresupuestos.Rows(0).Cells(4).Value IsNot Nothing AndAlso Date.TryParse(Me.DgvPresupuestos.Rows(0).Cells(4).Value.ToString(), vFechaFilaZero) Then
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(vFechaFilaZero.Month, False)
            Else
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(1, False)
            End If
        End If
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        EjecutarCalculoYDesviacion()
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

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access, cero reversión de textos)
            ' Como el combo está enlazado de forma relacional, convertimos el ítem actual en un DataRowView
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
                ' 🌟 SINCRO DE VARIABLES GLOBALES: Guardamos el código e ID numérico para tus grabaciones de fábrica
                vConcepto = codigoOriginal
                ' Guardamos el tipo real (GASTO/INGRESO) de la BD inmune al idioma del usuario
                vTipoConceptoActual = tipoOriginal.ToUpper()

                ' --- TRADUCIR LAS DESCRIPCIONES AUTOMÁTICAS (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción genérica de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal
                TxtConcepto.Text = tradDesc

                ' 3. Si el botón de aplicar filtro está activo (Enabled = False), recalculamos la pantalla
                If BtnFiltroConcepto.Enabled = False Then
                    EjecutarCalculoYDesviacion()
                End If
            End If

            ' =========================================================================
            ' 🌟 EL ESCUDO INDESTRUCTIBLE DE ENERO (Cortafuegos definitivo en el Combo)
            ' =========================================================================
            ' Como el chivato cantó que el sobreescrito ocurre aquí, este bloque lee la 
            ' fecha de la trastienda (Celda 4) y restaura "January / Enero" en un milisegundo.
            If Me.DgvPresupuestos.Rows.Count > 0 Then
                Dim vFechaFilaZero As Date
                If Me.DgvPresupuestos.Rows(0).Cells(4).Value IsNot Nothing AndAlso Date.TryParse(Me.DgvPresupuestos.Rows(0).Cells(4).Value.ToString(), vFechaFilaZero) Then
                    Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(vFechaFilaZero.Month, False)
                Else
                    ' Salvavidas regional de respaldo
                    Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(1, False)
                End If
            End If
        Catch ex As Exception
            ' Evita cuelgues visuales si el combo parpadea al interactuar
        End Try
    End Sub

    Private Sub EjecutarCalculoYDesviacion()
        BtnFiltroConcepto.Enabled = False
        BtnSinFiltroConcepto.Enabled = True
        cmdMdb1cr.Parameters.Clear()

        ' 🌟 EXTRAEMOS LOS DATA RELACIONALES DIRECTOS DESDE LA CACHÉ DE LA RAM
        Dim idConceptoSel As Integer = 0
        Dim tipoConcepto As String = vTipoConceptoActual.Trim().ToUpper()

        If CmbConcepto.SelectedItem IsNot Nothing Then
            Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)
            idConceptoSel = Convert.ToInt32(filaSeleccionada("IdConceptoCON"))
            ' Respaldamos el tipo directo desde la RAM por si la variable global falló
            If String.IsNullOrEmpty(tipoConcepto) AndAlso filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                tipoConcepto = filaSeleccionada("TipoCON").ToString().Trim().ToUpper()
            End If
        End If

        ' =========================================================================
        ' 🌟 CONSULTA SQL DE LA NUEVA ERA REALINEADA A TU DISEÑO ORIGINAL
        ' =========================================================================
        ' 0 = El Código Corto del Concepto (AESTHETICS, LUZ...) para tu columna 0.
        ' 1 = Duplicamos el campo para cumplir la estructura fija de 5 celdas de tu macro.
        ' 2 y 3 = Importe del presupuesto.
        ' 4 = Fecha de la trastienda para que la macro calcule el MonthName en la columna 1.
        ' 5, 6 y 7 = Los chivatos ocultos al final para que la macro no reviente por tipos.
        vtipoSql = "SELECT conceptos.CodigoCON, " &
                    "conceptos.CodigoCON, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.ConceptoPRE, " &
                    "conceptos.CodigoCON " &
                    "FROM presupuesto " &
                    "INNER JOIN conceptos ON presupuesto.ConceptoPRE = conceptos.IdConceptoCON " &
                    "WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString & " "
        ' Si estás llamando a esta consulta desde el método EjecutarCalculoYDesviacion, le pegas el filtro aquí:
        vtipoSql += $"And presupuesto.ConceptoPRE = {idConceptoSel} "
        vtipoSql += "ORDER BY conceptos.CodigoCON ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' Volcamos los datos limpios y forzamos la traducción internacional
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridPresupuestos(Me.DgvPresupuestos)

        ' 🚨 RESPALDO DE SEGURIDAD INMEDIATO: Guardamos los valores calculados en variables locales
        Dim miPresupuestoYTD As Double = vTotalPresupuestoYTD
        Dim miRealYTD As Double = vTotalRealYTD

        ' 3. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
        ActualizarEtiquetaDesviacion()

        ' 4. Pintamos el resultado basándonos en la mallas calculadas
        If Me.DgvPresupuestos.Rows.Count > 0 Then

            ' REFUERZO DE VISIBILIDAD: Forzamos que todos los Labels de resultados se muestren siempre
            LblDesviacion.Visible = True
            LblMontoDesviacion.Visible = True
            LblObjetivo.Visible = True
            LblDesviacion.Enabled = True

            ' EVALUACIÓN ESTÁNDAR UNIVERSAL: GASTO o INGRESO
            Dim esGasto As Boolean = (tipoConcepto = "GASTO")
            Dim desviacionFinal As Double

            ' 🔥 CORRECCIÓN MATEMÁTICA DE SIGNOS SEGÚN EL TIPO
            If esGasto Then
                ' Para Gastos: Desviación = Presupuesto - Real (Positivo es ahorro, objetivo logrado)
                desviacionFinal = miPresupuestoYTD - miRealYTD
            Else
                ' Para Ingresos: Pasamos el Real a positivo puro para compararlo correctamente
                Dim realIngresoPositivo As Double = Math.Abs(miRealYTD)
                ' Desviación = Real - Presupuesto (Si el Real es mayor, es un ingreso logrado)
                desviacionFinal = realIngresoPositivo - miPresupuestoYTD
            End If

            ' CONFIGURACIÓN VISUAL DEL TEXTO SEGÚN EL AÑO (Tus líneas originales impecables)
            Dim añoActualCalendario As Integer = Date.Today.Year

            If CInt(vAñoEjercicio) = añoActualCalendario Then
                ActualizarEtiquetaDesviacion()
            Else
                Dim textoAnual As String = rmse.GetString("LblDesviacion.Text")
                If String.IsNullOrEmpty(textoAnual) Then textoAnual = "Desviació Anual"
                LblDesviacion.Text = textoAnual & " " & vAñoEjercicio & " ="
            End If

            ' Mostramos la cifra final con formato "N2" para multiidioma
            LblMontoDesviacion.Text = desviacionFinal.ToString("N2") & " " & vMoneda

            ' CONTROL DE COLORES Y OBJETIVOS (Comportamiento financiero real de fábrica)
            If desviacionFinal >= 0 Then
                LblObjetivo.ForeColor = System.Drawing.Color.DarkGreen
                LblObjetivo.Text = rmse.GetString("LblObjetivo.Text")
                If String.IsNullOrEmpty(LblObjetivo.Text) Then LblObjetivo.Text = "Objectivo Logrado!"
                LblMontoDesviacion.ForeColor = System.Drawing.Color.DarkBlue
            Else
                LblObjetivo.ForeColor = System.Drawing.Color.DarkRed
                LblObjetivo.Text = rmse.GetString("NoLogrado")
                If String.IsNullOrEmpty(LblObjetivo.Text) Then LblObjetivo.Text = "Objectivo No Logrado"
                LblMontoDesviacion.ForeColor = System.Drawing.Color.Red
            End If
        Else
            ' Si el grid no tiene filas, limpiamos y ocultamos todo
            LblDesviacion.Enabled = False
            LblMontoDesviacion.Text = ""
            LblObjetivo.Visible = False
        End If

        ' =========================================================================
        ' 🌟 EL ESCUDO INDESTRUCTIBLE DE ENERO (Cortafuegos al quitar filtro)
        ' =========================================================================
        If Me.DgvPresupuestos.Rows.Count > 0 Then
            Dim vFechaFilaZero As Date
            If Me.DgvPresupuestos.Rows(0).Cells(4).Value IsNot Nothing AndAlso Date.TryParse(Me.DgvPresupuestos.Rows(0).Cells(4).Value.ToString(), vFechaFilaZero) Then
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(vFechaFilaZero.Month, False)
            Else
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(1, False)
            End If
        End If
    End Sub

    Private Sub BtnGraficos2D_Click(sender As Object, e As EventArgs) Handles BtnGraficos2D.Click
        ' 1. Comprobamos si existe un identificador asociado (Instancia segura)
        If (frmGraficosPresupuestos Is Nothing) OrElse (Not frmGraficosPresupuestos.IsHandleCreated) Then
            frmGraficosPresupuestos = New GraficosPresupuestos
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmGraficosPresupuestos)
        ' 2. TRUCO DE ORO: Le decimos que NO active el 3D (Gráfico plano)
        frmGraficosPresupuestos.EsGrafico3D = False

        ' 3. Llamamos al formulario de manera modal y lo destruimos al cerrar
        frmGraficosPresupuestos.ShowDialog()
        frmGraficosPresupuestos.Dispose()
    End Sub

    Private Sub BtnGraficos3D_Click(sender As Object, e As EventArgs) Handles BtnGraficos3D.Click
        ' 1. Comprobamos si existe un identificador asociado (Instancia segura)
        If (frmGraficosPresupuestos Is Nothing) OrElse (Not frmGraficosPresupuestos.IsHandleCreated) Then
            frmGraficosPresupuestos = New GraficosPresupuestos
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmGraficosPresupuestos)
        ' 2. TRUCO DE ORO: Le decimos que SÍ active el 3D (Gráfico con relieve)
        frmGraficosPresupuestos.EsGrafico3D = True

        ' 3. Llamamos al formulario de manera modal y lo destruimos al cerrar
        frmGraficosPresupuestos.ShowDialog()
        frmGraficosPresupuestos.Dispose()
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = 0
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
            Return
        Else
            vFila = vFilaActual - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
            Return
        Else
            vFila = vFilaActual + 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = DgvPresupuestos.RowCount - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnF6_Click(sender As Object, e As EventArgs) Handles BtnF6.Click
        ' Vuelve a Refrescar el DataGrid y dejar los Btn de los Filtros sin Filtrar
        ' ************************************************************************
        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 CONSULTA SQL DE LA NUEVA ERA REALINEADA A TU DISEÑO ORIGINAL
        ' =========================================================================
        ' 0 = El Código Corto del Concepto (AESTHETICS, LUZ...) para tu columna 0.
        ' 1 = Duplicamos el campo para cumplir la estructura fija de 5 celdas de tu macro.
        ' 2 y 3 = Importe del presupuesto.
        ' 4 = Fecha de la trastienda para que la macro calcule el MonthName en la columna 1.
        ' 5, 6 y 7 = Los chivatos ocultos al final para que la macro no reviente por tipos.
        vtipoSql = "SELECT conceptos.CodigoCON, " &
                    "conceptos.CodigoCON, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.ImportePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.FDesdePRE, " &
                    "presupuesto.ConceptoPRE, " &
                    "conceptos.CodigoCON " &
                    "FROM presupuesto " &
                    "INNER JOIN conceptos ON presupuesto.ConceptoPRE = conceptos.IdConceptoCON " &
                    "WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString & " "
        ' Si estás llamando a esta consulta desde el método EjecutarCalculoYDesviacion, le pegas el filtro aquí:
        ' vtipoSql += $"And presupuesto.ConceptoPRE = {idConceptoSel} "
        vtipoSql += "ORDER BY conceptos.CodigoCON ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' Volcamos los datos limpios y forzamos la traducción internacional
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridPresupuestos(Me.DgvPresupuestos)

        ' =========================================================================
        ' INTERFAZ ESTÉTICA Y RESETEO DE ETIQUETAS (Tu lógica original impecable)
        ' =========================================================================
        LblNumRegistros.Text = resManager.GetString("SinFiltrar")
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False

        ' Ocultamos las etiquetas de desviación al limpiar el filtro por concepto único
        LblDesviacion.Visible = False
        LblMontoDesviacion.Visible = False
        LblObjetivo.Visible = False

        ' 2. REPOSICIONAMIENTO SEGURO AL INICIO
        If DgvPresupuestos.RowCount > 0 Then
            DgvPresupuestos.Rows(0).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(0).Cells(0)
        End If

        ' =========================================================================
        ' 🌟 EL ESCUDO INDESTRUCTIBLE DE ENERO (Cortafuegos al quitar filtro)
        ' =========================================================================
        If Me.DgvPresupuestos.Rows.Count > 0 Then
            Dim vFechaFilaZero As Date
            If Me.DgvPresupuestos.Rows(0).Cells(4).Value IsNot Nothing AndAlso Date.TryParse(Me.DgvPresupuestos.Rows(0).Cells(4).Value.ToString(), vFechaFilaZero) Then
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(vFechaFilaZero.Month, False)
            Else
                Me.DgvPresupuestos.Rows(0).Cells(1).Value = MonthName(1, False)
            End If
        End If
    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        ' Elimina Visualmente las Filas Seleccionadas desde la memoria RAM
        ' *******************************************************************
        If DgvPresupuestos.SelectedRows.Count > 0 Then

            ' 1. Obtenemos el DataTable enlazado a tu DgvPresupuestos de forma legal para .NET
            Dim dt As DataTable = CType(DgvPresupuestos.DataSource, DataTable)

            ' 🌟 EXCLUSIVO PRESUPUESTOS: Quitamos la fila estática de "TOTAL" de la RAM antes de borrar
            If dt.Rows.Count > 0 Then
                Dim ultimaFila As DataRow = dt.Rows(dt.Rows.Count - 1)
                Dim textoTotalTraducido As String = resManager.GetString("TOTAL")
                If String.IsNullOrEmpty(textoTotalTraducido) Then textoTotalTraducido = "TOTAL"

                If Convert.ToString(ultimaFila(0)) = textoTotalTraducido Then
                    dt.Rows.Remove(ultimaFila)
                End If
            End If

            ' 🌟 TU MATRIZ DE AYER: Borramos del DataTable en la RAM, pero NO hace el DELETE en Access
            For i As Integer = DgvPresupuestos.SelectedRows.Count - 1 To 0 Step -1
                Dim fila As DataGridViewRow = DgvPresupuestos.SelectedRows(i)

                ' Nos saltamos la fila vacía del final por seguridad
                If fila.IsNewRow Then Continue For

                ' Extraemos el enlace de datos puro de la fila y lo eliminamos de la RAM
                If fila.DataBoundItem IsNot Nothing Then
                    Dim rowView As DataRowView = CType(fila.DataBoundItem, DataRowView)
                    rowView.Delete()
                End If
            Next

            ' =========================================================================
            ' 🌟 TU EXCELENTE LOGICA DE RECALCULO DE TOTALES (Mantenida e intacta)
            ' =========================================================================
            Dim totalCol2 As Decimal = 0
            Dim totalCol3 As Decimal = 0

            ' Recorremos las filas que han quedado vivas para acumular las sumas
            For Each fila As DataGridViewRow In DgvPresupuestos.Rows
                Dim valorCol2 As Decimal = 0
                Dim valorCol3 As Decimal = 0

                Decimal.TryParse(Convert.ToString(fila.Cells(2).Value), valorCol2)
                Decimal.TryParse(Convert.ToString(fila.Cells(3).Value), valorCol3)

                totalCol2 += valorCol2
                totalCol3 += valorCol3
            Next

            ' Creamos e inyectamos la nueva fila de totales reluciente en el DataTable
            Dim nuevaFila As DataRow = dt.NewRow()
            nuevaFila(0) = If(resManager?.GetString("TOTAL"), "TOTAL")
            nuevaFila(2) = totalCol2
            nuevaFila(3) = totalCol3
            dt.Rows.Add(nuevaFila)

            ' =========================================================================
            ' 2. REPOSICIONAMIENTO SEGURO AL FINAL (Tu idéntica estructura de ayer)
            ' =========================================================================
            If DgvPresupuestos.Rows.Count > 0 Then
                ' Limpiamos selecciones fantasma
                For idx = 0 To DgvPresupuestos.Rows.Count - 1
                    DgvPresupuestos.Rows(idx).Selected = False
                Next

                DgvPresupuestos.Select()

                ' Calculamos de forma dinámica el índice de la ÚLTIMA fila
                Dim ultimaFilaViva As Integer = DgvPresupuestos.Rows.Count - 1

                ' Validamos que el índice sea válido (mayor o igual a 0) para evitar desbordamientos
                If ultimaFilaViva >= 0 Then
                    DgvPresupuestos.Rows(ultimaFilaViva).Selected = True
                    DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(ultimaFilaViva).Cells(0)
                End If
                DgvPresupuestos.Refresh()
            End If
        End If
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        ' 1. Aseguramos de forma preventiva que haya una fila seleccionada en el Grid
        If DgvPresupuestos.CurrentRow Is Nothing Then Exit Sub
        cmdMdb1cr.Parameters.Clear()

        ' 2. CUADRO DE CONFIRMACIÓN INTERNACIONALIZADO
        Dim msgPregunta As String = rmse.GetString("PreguntaEliminarPresupuesto")
        Dim titPregunta As String = rmse.GetString("TituloEliminarPresupuesto")

        If ConfirmarAccionTraducida(msgPregunta, titPregunta) = MsgBoxResult.No Then
            Exit Sub
        End If

        filaActual = DgvPresupuestos.CurrentRow.Index

        ' =========================================================================
        ' 🌟 TU VACIADO CLÁSICO COMPLETO: BORRADO POR LA TUBERÍA MAESTRA (CONEXION1)
        ' =========================================================================
        If DgvPresupuestos.Rows(filaActual).Cells(6).Value IsNot Nothing AndAlso Not IsDBNull(DgvPresupuestos.Rows(filaActual).Cells(6).Value) Then
            Dim idConceptoBorrar As Integer = Convert.ToInt32(DgvPresupuestos.Rows(filaActual).Cells(6).Value)

            ' Sentencia parametrizada dirigida al corazón de presupuesto
            Dim sqlDelete As String = "DELETE FROM presupuesto WHERE ConceptoPRE = ? AND EjercicioPRE = ?"

            ' 🚀 LA JUGADA MAESTRA 1: Usamos directamente cmdMdb1cr sobre conexion1 para evitar el retraso de caché
            cmdMdb1cr.CommandText = sqlDelete
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoBorrar
            cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)

            Try
                cmdMdb1cr.ExecuteNonQuery()

                Dim msgExito As String = rmse.GetString("PresupuestosBorradosExito")
                If String.IsNullOrEmpty(msgExito) Then msgExito = "Registros en Presupuestos, Borrados !!!"
                MsgBox(msgExito, vbInformation, titPregunta)
            Catch ex As Exception
                Dim msgError As String = rmse.GetString("ErrorEliminarPresupuestos")
                If String.IsNullOrEmpty(msgError) Then
                    msgError = "No se han podido eliminar los registros en Presupuestos..."
                End If
                MsgBox(msgError & vbNewLine & ex.Message, vbCritical, resManager.GetString("Error"))
                Exit Sub
            End Try
        End If

        ' =========================================================================
        ' 4. RECARGA DEL GRID DE PRESUPUESTOS INMUNE A FANTASMAS (Estructura de 8 celdas)
        ' =========================================================================
        ' 🚀 LA JUGADA MAESTRA 2: Rompemos el dibujo viejo antes de volver a llamar al LlenarGrid
        DgvPresupuestos.DataSource = Nothing

        vtipoSql = "SELECT conceptos.CodigoCON, conceptos.CodigoCON, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE, presupuesto.FDesdePRE, presupuesto.ConceptoPRE, conceptos.CodigoCON FROM (presupuesto INNER JOIN conceptos ON presupuesto.ConceptoPRE = conceptos.IdConceptoCON) WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString & " ORDER BY conceptos.CodigoCON ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' Volcamos los datos limpios de la base en la cuadrícula
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridPresupuestos(Me.DgvPresupuestos)

        ' Evaluamos etiquetas de desviación (Tu lógica de fábrica impecable)
        ActualizarEtiquetaDesviacion()

        ' Forzamos la limpieza si la cuadrícula se quedó desierta
        If DgvPresupuestos.Rows.Count = 0 Then
            LblDesviacion.Enabled = False
            LblMontoDesviacion.Text = ""
            LblObjetivo.Visible = False
        Else
            ' Reposicionamos el foco en la fila inicial de forma dócil
            DgvPresupuestos.Rows(0).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(0).Cells(0)
        End If
    End Sub
    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub DgvPresupuestos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgvPresupuestos.CellFormatting
        ' Comprobamos que no sea la fila de cabecera
        If e.RowIndex >= 0 Then
            Dim dgv As DataGridView = CType(sender, DataGridView)

            ' =========================================================================
            ' 🌟 EL ATAJO INDEXADO FINANCIERO (Rendimiento ultra-rápido de la RAM)
            ' =========================================================================
            ' Forzamos el formateo estético en gris y negrita ÚNICAMENTE si la fila actual 
            ' es la última de la cuadrícula (la fila de totales agregada por el DataTable).
            If e.RowIndex = dgv.Rows.Count - 1 Then

                ' Aplicamos el fondo gris de cierre contable y el texto negro nítido
                e.CellStyle.BackColor = System.Drawing.Color.LightGray
                e.CellStyle.ForeColor = System.Drawing.Color.Black

                ' Heredamos la fuente nativa del Grid y le aplicamos negrita impecable
                e.CellStyle.Font = New Font(dgv.Font, FontStyle.Bold)

            Else
                ' 🚀 PLAN B: Si es una fila normal de mes (enero a diciembre), 
                ' nos aseguramos de restaurar el fondo blanco limpio y la letra normal de fábrica
                e.CellStyle.BackColor = System.Drawing.Color.White
                e.CellStyle.ForeColor = System.Drawing.Color.Black
                e.CellStyle.Font = New Font(dgv.Font, FontStyle.Regular)
            End If
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' 1. LIMPIEZA INICIAL DE LA TABLA TEMPORAL
        ' ***************************************
        vtipoSql = "DELETE * FROM tmpprint"
        Dim cmdMdb1 As New OleDbCommand(vtipoSql, conexion1)
        Try
            If conexion1.State <> ConnectionState.Open Then conexion1.Open()
            cmdMdb1.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorLimpiarTemporal") & ": " & ex.Message)
        End Try

        ' VARIABLES PARA CONTROL DE FECHAS (YTD)
        Dim añoActualCalendario As Integer = DateTime.Now.Year
        Dim mesActualCalendario As Integer = DateTime.Now.Month

        ' DICCIONARIO PARA AGRUPAR LOS TOTALES POR CONCEPTO ANTES DE IMPRIMIR
        Dim presupuestosAgrupados As New Dictionary(Of String, (Real As Double, Presu As Double, Tipo As String))

        ' 2. PRIMER PASO: RECORREMOS TU REJILLA PRINCIPAL FILA A FILA (MES A MES)
        ' *******************************************************************
        For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
            If fila.IsNewRow Then Continue For

            ' 🌟 CORTAFUEGOS INDESTRUCTIBLE MULTIIDIOMA POR BÚSQUEDA INVERSA
            ' Le pasamos el texto visual de la celda 0 a tu función del módulo.
            ' Si el escáner detecta que la llave de esa palabra es "TOTAL", saltamos la fila de golpe.
            If fila.Cells(0).Value IsNot Nothing Then
                Dim textoCeldaConcepto As String = fila.Cells(0).Value.ToString().Trim()
                If ObtenerClaveNeutral(textoCeldaConcepto, resManager) = "TOTAL" Then
                    Continue For
                End If
            End If


            Dim vNombreConcepto As String = fila.Cells(0).Value.ToString()

            ' Extraemos el mes de la celda de fecha (Columna 4) de forma segura
            Dim vFechaFila As Date
            Dim vMesInt As Integer = 1
            If fila.Cells(4).Value IsNot Nothing AndAlso Date.TryParse(fila.Cells(4).Value.ToString(), vFechaFila) Then
                vMesInt = vFechaFila.Month
            End If

            ' RECORTE QUIRÚRGICO YTD PARA EL PAPEL
            If CInt(vAñoEjercicio) = añoActualCalendario Then
                If vMesInt >= mesActualCalendario Then Continue For
            End If

            ' Extraemos los valores calculados de la fila actual de la rejilla
            Dim valRealFila As Double = 0
            Dim valPresuFila As Double = 0

            ' RECOLECCIÓN FIEL DE SIGNOS: Respetamos los signos de la pantalla (Negativos en Ingresos)
            If fila.Cells(2).Value IsNot Nothing Then Double.TryParse(fila.Cells(2).Value.ToString(), valRealFila)
            If fila.Cells(3).Value IsNot Nothing Then Double.TryParse(fila.Cells(3).Value.ToString(), valPresuFila)

            ' Si es la primera vez que vemos este concepto, buscamos su tipo (Gasto/Ingreso)
            If Not presupuestosAgrupados.ContainsKey(vNombreConcepto) Then
                Dim vTipoConceptoImprimir As String = "GASTO"
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Using cmd As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & vNombreConcepto.Replace("'", "''") & "'", con)
                        Try
                            con.Open()
                            Dim res As Object = cmd.ExecuteScalar()
                            If res IsNot Nothing Then vTipoConceptoImprimir = res.ToString().Trim().ToUpper()
                        Catch
                        End Try
                    End Using
                End Using

                Dim vCuentaTMP As String = If(vTipoConceptoImprimir = "INGRESO", "I", "G")
                presupuestosAgrupados(vNombreConcepto) = (0, 0, vCuentaTMP)
            End If

            ' Sumamos los valores respetando el signo algebraico puro
            Dim datosActuales = presupuestosAgrupados(vNombreConcepto)
            presupuestosAgrupados(vNombreConcepto) = (datosActuales.Real + valRealFila, datosActuales.Presu + valPresuFila, datosActuales.Tipo)
        Next

        ' 3. SEGUNDO PASO: ENVIAMOS LOS DATOS AGRUPADOS LIMPIOS A LA TABLA INTERMEDIA (tmpprint)
        ' **************************************************************************************
        For Each kvp In presupuestosAgrupados
            Dim concepto As String = kvp.Key
            Dim acumuladoReal As Double = kvp.Value.Real
            Dim presupuestoFinalGuardar As Double = kvp.Value.Presu
            Dim vCuentaTMP As String = kvp.Value.Tipo

            ' Si es de tipo "I" (Ingreso), nos aseguramos de que viajen a Access en NEGATIVO PURO para que la resta dé correcta
            If vCuentaTMP = "I" Then
                acumuladoReal = -Math.Abs(acumuladoReal)
                presupuestoFinalGuardar = -Math.Abs(presupuestoFinalGuardar)
            End If

            ' 🔥 CORREGIDO: Mapeo exacto con los nombres de tu tabla Access (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP)
            Dim vAñadir As String = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) "
            vAñadir += "VALUES (#2023-01-11#, ?, '', ?, '', ?, ?)"

            Using cmdMdb1cr As New OleDbCommand(vAñadir, conexion1)
                ' Los parámetros se asignan en orden estricto de aparición de los signos de interrogación (?)
                cmdMdb1cr.Parameters.AddWithValue("@ConceptoTMP", concepto)
                cmdMdb1cr.Parameters.AddWithValue("@CuentaTMP", vCuentaTMP)
                cmdMdb1cr.Parameters.AddWithValue("@ImporteTMP", acumuladoReal)
                cmdMdb1cr.Parameters.AddWithValue("@SaldoTMP", presupuestoFinalGuardar)

                Try
                    If conexion1.State <> ConnectionState.Open Then conexion1.Open()
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        Next

        ' 4. LLENADO FINAL Y APERTURA DEL REPORTE
        ' ***************************************
        vtipoSql = "SELECT * FROM tmpprint ORDER BY ConceptoTMP ASC"
        LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")

        ' Iniciamos Código para Imprimir
        ' ******************************
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString
        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        If My.Settings.Previsualizar = True Then
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        If My.Settings.ElegirImpresora = True Then
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        End If

        If My.Settings.DirectoImpresora = True Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Definición de Estilos Profesionales
        ' **********************************
        Dim FuenteTitulo As New Font("Segoe UI", 13, FontStyle.Bold)
        Dim FuenteSeccion As New Font("Segoe UI", 11, FontStyle.Bold)
        Dim FuenteNegrita As New Font("Segoe UI", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Segoe UI", 9, FontStyle.Regular)

        Dim sfDerecha As New StringFormat With {.Alignment = StringAlignment.Far}

        ' Encabezado del Reporte
        ' **********************
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteDetalles, Brushes.Gray, 750, 40, sfDerecha)

        Dim tituloReporte As String = If(BtnFiltroConcepto.Enabled = False, rmse.GetString("InformeFiltrado") & " " & vAñoEjercicio, rmse.GetString("InformeSinFiltrar") & " " & vAñoEjercicio)

        ' 🔥 CONTROL YTD: Si es el año actual, añadimos dinámicamente el mes traducido
        If CInt(vAñoEjercicio) = DateTime.Now.Year Then
            Dim mesAnteriorInt As Integer = DateTime.Now.Month - 1

            If mesAnteriorInt > 0 Then
                ' Generamos la clave para buscar en el .resx (ej: "Mes_1", "Mes_2"...)
                Dim claveMes As String = "Mes_" & mesAnteriorInt.ToString()

                ' Buscamos el nombre del mes en tu archivo de recursos
                Dim nombreMesAnterior As String = rmse.GetString(claveMes)

                ' Si no lo encuentra en el .resx, usamos MonthName como rueda de repuesto
                If String.IsNullOrEmpty(nombreMesAnterior) Then
                    nombreMesAnterior = MonthName(mesAnteriorInt, False)
                End If

                ' Buscamos la traducción de "Hasta el mes"
                Dim textoHastaMes As String = rmse.GetString("HastaElMes")
                If String.IsNullOrEmpty(textoHastaMes) Then textoHastaMes = "Hasta el mes de "

                ' Concatenamos el texto al título principal
                tituloReporte &= $" ({textoHastaMes} {nombreMesAnterior})"
            End If
        End If

        e.Graphics.DrawString(tituloReporte, FuenteTitulo, Brushes.DarkBlue, 50, 75)

        If frmImprimirForm.PictureBox1.Image IsNot Nothing Then
            e.Graphics.DrawImage(frmImprimirForm.PictureBox1.Image, 50, 30, 80, 40)
        End If

        ' Configuración de Columnas (Coordenadas X fijas)
        Dim colConceptoX As Integer = 50
        Dim colRealX As Integer = 380      ' Espacio optimizado para evitar solapamientos
        Dim colPresuX As Integer = 560
        Dim colDesvX As Integer = 750
        Dim posY As Integer = 140

        ' Variables de acumulación macro
        Dim totalIngresosReal As Double = 0 : Dim totalIngresosPresu As Double = 0
        Dim totalGastosReal As Double = 0 : Dim totalGastosPresu As Double = 0

        ' =========================================================================
        ' BLOQUE 1: INGRESOS
        ' =========================================================================
        e.Graphics.DrawString(rmse.GetString("EstrucIngresos"), FuenteSeccion, Brushes.DarkGreen, colConceptoX, posY)
        posY += 25

        e.Graphics.DrawString(rmse.GetString("Concepto") & ":", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString(rmse.GetString("Realidad") & " YTD:", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(rmse.GetString("Presupuesto") & ":", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        e.Graphics.DrawString(rmse.GetString("Desviacion") & ":", FuenteNegrita, Brushes.Black, colDesvX, posY, sfDerecha)
        posY += 20
        e.Graphics.DrawLine(Pens.Gray, colConceptoX, posY, colDesvX, posY)
        posY += 10

        Dim tieneIngresos As Boolean = False
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow OrElse fila.Cells("CuentaTMP").Value IsNot Nothing AndAlso fila.Cells("CuentaTMP").Value.ToString() <> "I" Then Continue For

            tieneIngresos = True
            Dim concepto As String = fila.Cells("ConceptoTMP").Value.ToString()

            Dim valReal As Double = 0 : Dim valPresu As Double = 0
            Double.TryParse(fila.Cells("ImporteTMP").Value?.ToString(), valReal)
            Double.TryParse(fila.Cells("SaldoTMP").Value?.ToString(), valPresu)

            ' Pasamos a positivo puro para razonar la lógica financiera estándar
            Dim valRealPositivo As Double = Math.Abs(valReal)
            Dim valPresuPositivo As Double = Math.Abs(valPresu)

            totalIngresosReal += valReal
            totalIngresosPresu += valPresu

            ' 🔥 LÓGICA DE INGRESOS REAL: Si el Real es MAYOR que el Presupuesto, la desviación es POSITIVA (Ganancia)
            Dim desv As Double = valRealPositivo - valPresuPositivo

            ' Determinamos el color: si la desviación es mayor o igual a 0, verde (¡Éxito!), si no, rojo
            Dim brushDesv As Brush = If(desv >= 0, Brushes.Green, Brushes.Red)

            ' Imprimimos en el papel (usamos Math.Abs para que no pinte guiones menos feos en los totales)
            e.Graphics.DrawString(valRealPositivo.ToString("N2") & " " & vMoneda, FuenteDetalles, Brushes.Black, colRealX, posY, sfDerecha)
            e.Graphics.DrawString(valPresuPositivo.ToString("N2") & " " & vMoneda, FuenteDetalles, Brushes.Black, colPresuX, posY, sfDerecha)

            ' Pintamos el signo más (+) si es positivo para que quede claro que es un beneficio extra
            e.Graphics.DrawString((If(desv >= 0, "+", "")) & desv.ToString("N2") & " " & vMoneda, FuenteDetalles, brushDesv, colDesvX, posY, sfDerecha)
            posY += 22

        Next

        If Not tieneIngresos Then
            e.Graphics.DrawString(rmse.GetString("NoHayIngresos"), FuenteDetalles, Brushes.Gray, colConceptoX + 20, posY)
            posY += 22
        End If

        e.Graphics.DrawLine(Pens.LightGray, colConceptoX, posY, colDesvX, posY)
        posY += 5
        e.Graphics.DrawString(rmse.GetString("TotalIngresos"), FuenteNegrita, Brushes.Black, colConceptoX, posY)

        ' 🔥 CALCULAMOS LA DESVIACIÓN DEL TOTAL EN POSITIVO (Real - Presupuesto)
        Dim desvIngTotal As Double = Math.Abs(totalIngresosReal) - Math.Abs(totalIngresosPresu)
        Dim brushTotalIng As Brush = If(desvIngTotal >= 0, Brushes.Green, Brushes.Red)

        ' 🔥 IMPRIMIMOS LOS TOTALES SIN SIGNOS NEGATIVOS MOLESTOS
        e.Graphics.DrawString(Math.Abs(totalIngresosReal).ToString("N2") & " " & vMoneda, FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(Math.Abs(totalIngresosPresu).ToString("N2") & " " & vMoneda, FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)

        ' 🔥 DIBUJAMOS EL RESULTADO DE LA DESVIACIÓN DEL TOTAL CON SU SIGNO CORRECTO
        e.Graphics.DrawString((If(desvIngTotal >= 0, "+", "")) & desvIngTotal.ToString("N2") & " " & vMoneda, FuenteNegrita, brushTotalIng, colDesvX, posY, sfDerecha)

        posY += 40

        ' =========================================================================
        ' BLOQUE 2: GASTOS
        ' =========================================================================
        e.Graphics.DrawString(rmse.GetString("EstrucGastos"), FuenteSeccion, Brushes.DarkRed, colConceptoX, posY)
        posY += 25

        e.Graphics.DrawString(rmse.GetString("Concepto") & ":", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString(rmse.GetString("Realidad") & " YTD:", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(rmse.GetString("Presupuesto") & ":", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        e.Graphics.DrawString(rmse.GetString("DesviacionAhorro") & ":", FuenteNegrita, Brushes.Black, colDesvX, posY, sfDerecha)
        posY += 20
        e.Graphics.DrawLine(Pens.Gray, colConceptoX, posY, colDesvX, posY)
        posY += 10

        Dim tieneGastos As Boolean = False
        For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
            If fila.IsNewRow OrElse fila.Cells("CuentaTMP").Value IsNot Nothing AndAlso fila.Cells("CuentaTMP").Value.ToString() <> "G" Then Continue For

            tieneGastos = True
            Dim concepto As String = fila.Cells("ConceptoTMP").Value.ToString()
            Dim valReal As Double = 0 : Dim valPresu As Double = 0
            Double.TryParse(fila.Cells("ImporteTMP").Value?.ToString(), valReal)
            Double.TryParse(fila.Cells("SaldoTMP").Value?.ToString(), valPresu)

            totalGastosReal += valReal
            totalGastosPresu += valPresu

            Dim desv As Double = valPresu - valReal
            Dim brushDesv As Brush = If(desv >= 0, Brushes.Green, Brushes.Red)

            e.Graphics.DrawString(concepto, FuenteDetalles, Brushes.Black, colConceptoX, posY)
            e.Graphics.DrawString(valReal.ToString("N2") & " " & vMoneda, FuenteDetalles, Brushes.Black, colRealX, posY, sfDerecha)
            e.Graphics.DrawString(valPresu.ToString("N2") & " " & vMoneda, FuenteDetalles, Brushes.Black, colPresuX, posY, sfDerecha)
            e.Graphics.DrawString((If(desv >= 0, "+", "")) & desv.ToString("N2") & " " & vMoneda, FuenteDetalles, brushDesv, colDesvX, posY, sfDerecha)
            posY += 22
        Next

        If Not tieneGastos Then
            e.Graphics.DrawString(rmse.GetString("NoHayGastos"), FuenteDetalles, Brushes.Gray, colConceptoX + 20, posY)
            posY += 22
        End If

        e.Graphics.DrawLine(Pens.LightGray, colConceptoX, posY, colDesvX, posY)
        posY += 5
        e.Graphics.DrawString(rmse.GetString("TotalGastos"), FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString(totalGastosReal.ToString("N2") & " " & vMoneda, FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(totalGastosPresu.ToString("N2") & " " & vMoneda, FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        Dim desvGasTotal As Double = totalGastosPresu - totalGastosReal
        e.Graphics.DrawString(desvGasTotal.ToString("N2") & " " & vMoneda, FuenteNegrita, If(desvGasTotal >= 0, Brushes.Green, Brushes.Red), colDesvX, posY, sfDerecha)

        posY += 50

        ' =========================================================================
        ' BLOQUE 3: CUADRO DE RESUMEN EJECUTIVO (Diseño vertical anti-solapamiento)
        ' =========================================================================
        ' Ampliamos el alto del recuadro a 105 para dar espacio al formato de lista vertical
        Dim fondoResumen As New Rectangle(colConceptoX, posY, colDesvX - colConceptoX, 105)
        e.Graphics.FillRectangle(Brushes.GhostWhite, fondoResumen)
        e.Graphics.DrawRectangle(Pens.SlateGray, fondoResumen)

        posY += 15
        e.Graphics.DrawString(rmse.GetString("ResultadoNeto"), FuenteSeccion, Brushes.Black, colConceptoX + 15, posY)

        ' 🔥 NEUTRALIZACIÓN DE SIGNOS ALGEBRAICOS (Pasamos ingresos y gastos a positivo puro)
        Dim ingresosRealPuro As Double = Math.Abs(totalIngresosReal)
        Dim ingresosPresuPuro As Double = Math.Abs(totalIngresosPresu)

        Dim gastosRealPuro As Double = Math.Abs(totalGastosReal)
        Dim gastosPresuPuro As Double = Math.Abs(totalGastosPresu)

        ' 🔥 1. CÁLCULO REAL NETO: Ingresos Puros - Gastos Puros
        Dim netoReal As Double = ingresosRealPuro - gastosRealPuro

        ' 🔥 2. CÁLCULO PRESUPUESTO NETO: Ingresos Puros - Gastos Puros
        Dim netoPresu As Double = ingresosPresuPuro - gastosPresuPuro

        ' 🔥 3. DESVIACIÓN GLOBAL REAL: Resultado Real Neto - Resultado Presupuestado Neto
        Dim desvNetalGlobal As Double = netoReal - netoPresu

        ' FILA 1: Resultado Real (Bajo su columna correspondiente)
        posY += 25
        e.Graphics.DrawString(rmse.GetString("ResultadoRealNeto") & ":", FuenteNegrita, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString(netoReal.ToString("N2") & " " & vMoneda, FuenteNegrita, If(netoReal >= 0, Brushes.DarkGreen, Brushes.DarkRed), colDesvX - 15, posY, sfDerecha)

        ' FILA 2: Resultado Presupuestado
        posY += 20
        e.Graphics.DrawString(rmse.GetString("ResultadoPresupuestadoNeto") & ":", FuenteDetalles, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString(netoPresu.ToString("N2") & " " & vMoneda, FuenteDetalles, If(netoPresu >= 0, Brushes.Black, Brushes.DarkRed), colDesvX - 15, posY, sfDerecha)

        ' FILA 3: Desviación Global Destacada
        posY += 22
        e.Graphics.DrawString(rmse.GetString("DesviacionGlobal") & ":", FuenteNegrita, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString((If(desvNetalGlobal >= 0, "+", "")) & desvNetalGlobal.ToString("N2") & " " & vMoneda, FuenteSeccion, If(desvNetalGlobal >= 0, Brushes.Green, Brushes.Red), colDesvX - 15, posY - 4, sfDerecha)


        ' =========================================================================
        ' PIE DE PÁGINA: NUMERACIÓN AUTOMÁTICA
        ' =========================================================================
        Dim nPagina As Integer = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        frmImprimirForm.LblNumeroPagina.Text = nPagina.ToString()

        Dim textoPagina As String = resManager.GetString("Pagina") & " " & nPagina.ToString()
        e.Graphics.DrawString(textoPagina, FuenteDetalles, Brushes.DimGray, 750, 1050, sfDerecha)

        e.HasMorePages = False
    End Sub

End Class