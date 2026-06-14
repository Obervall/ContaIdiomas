Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class Presupuestos

    Public vtipoSql, vtipoGrid, vConcepto, vAñadir, vAñadir2 As String
    Public vTmpprint As String
    Public PrintLine, Contador, FilaSelec As Integer
    Public vTipoConceptoActual As String = ""
    Public TL() As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub Presupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        ' Inicialización centralizada de ToolTips
        ' *=====================================
        Dim controlesToolTip As Control() = {
            BtnGraficos2D, BtnSalir, BtnFiltroConcepto, BtnSinFiltroConcepto, BtnImprimir,
            BtnPrimero, BtnAnterior, BtnSiguiente, BtnUltimo, BtnEliminarRegistro, BtnGraficos3D,
            BtnEliminaSeleccion, BtnF6 ' <-- Añadido aquí
        }

        Dim clavesToolTip As String() = {
            "ToolTipGraficos2D", "ToolTipSalir", "ToolTipAplicarFiltro", "ToolTipQuitarFiltro", "ToolTipImprimir",
            "ToolTipPrimero", "ToolTipAnterior", "ToolTipSiguiente", "ToolTipUltimo", "ToolTipEliminar", "ToolTipGraficos3D",
            "ToolTipEliminaSeleccion", "ToolTipF6" ' <-- Añadido aquí
        }

        ' TRUCO DE ORO: Redimensionamos la matriz TL automáticamente según el número de controles
        ' El (-1) es porque las matrices en .NET empiezan a contar desde el 0
        ReDim TL(controlesToolTip.Length - 1)

        ' El bucle ahora recorrerá los 11 elementos sin peligro de desbordamiento
        For i As Integer = 0 To controlesToolTip.Length - 1
            TL(i) = New ToolTip()
            TL(i).SetToolTip(controlesToolTip(i), resManager.GetString(clavesToolTip(i)))
        Next

        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbConcepto.Items.Add(drMdb1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' 🔥 LIMPIEZA RADICAL DE CONCEPTOS VACÍOS:
        ' Borramos de la base de datos cualquier registro de presupuesto de este año cuyo Importe sea 0
        Dim sqlLimpieza As String = "DELETE FROM presupuesto WHERE EjercicioPRE = " & vAñoEjercicio.ToString & " AND ImportePRE = 0"
        Using cmdLimpiar As New OleDbCommand(sqlLimpieza, conexion1)
            Try
                If conexion1.State <> ConnectionState.Open Then conexion1.Open()
                cmdLimpiar.ExecuteNonQuery()
            Catch ex As Exception
                ' Fallo silencioso por seguridad si la tabla está bloqueada temporalmente
            End Try
        End Using

        ' Llenar Grid de PRESUPUESTOS
        '****************************
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        LblDesviacion.Visible = False
        Label2.Visible = False
        LblMontoDesviacion.Visible = False
        LblObjetivo.Visible = False
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False

        ' 1. Lanzamos la consulta totalitaria del año. El módulo recalculará las variables globales.
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        ' 2. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
        ActualizarEtiquetaDesviacion()

        ' 2. Al no haber filtro por concepto único, lo estándar es ocultar la desviación macro
        LblDesviacion.Enabled = False
        LblObjetivo.Visible = False
        LblMontoDesviacion.Text = ""
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        EjecutarCalculoYDesviacion()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 1. Aseguramos que haya un texto seleccionado
        vConcepto = CmbConcepto.Text.ToString().Trim()
        If String.IsNullOrEmpty(vConcepto) Then Exit Sub

        ' 🔄 REVERTIR EL IDIOMA DE FORMA LIMPIA CON EL RESX MANAGER
        ' Buscamos la clave original en español usando el valor traducido visible
        Dim conceptoOriginalMDB As String = vConcepto

        ' Si la app no está en español, recuperamos de forma directa la cadena base
        If Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
            ' Buscamos en el diccionario de recursos la clave asociada al texto traducido
            Dim claveRecurso As String = resManager.GetString(vConcepto)
            If Not String.IsNullOrEmpty(claveRecurso) Then
                conceptoOriginalMDB = claveRecurso.Replace("_", " ")
            End If
        End If

        ' 2. Consulta combinada: Rescatamos la Descripción Y el Tipo original de golpe
        Dim sqlConcepto As String = "SELECT DescripcionCON, TipoCON FROM conceptos WHERE CodigoCON = ?"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using cmd As New OleDbCommand(sqlConcepto, conexion)
                cmd.Parameters.AddWithValue("@cod", conceptoOriginalMDB)
                Try
                    conexion.Open()
                    Using dr As OleDbDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            ' Asignamos la descripción al cuadro de texto en pantalla
                            TxtConcepto.Text = dr("DescripcionCON").ToString()
                            ' Guardamos el tipo real (GASTO/INGRESO) en nuestra variable global
                            vTipoConceptoActual = dr("TipoCON").ToString().Trim().ToUpper()
                        End If
                    End Using
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End Using

        ' 3. Si el botón de aplicar filtro está activo, recalculamos la pantalla
        If BtnFiltroConcepto.Enabled = False Then
            EjecutarCalculoYDesviacion()
        End If
    End Sub

    Private Sub EjecutarCalculoYDesviacion()
        BtnFiltroConcepto.Enabled = False
        BtnSinFiltroConcepto.Enabled = True

        ' 1. Generamos la consulta SQL de filtrado quirúrgico
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " And presupuesto.ConceptoPRE = '" & CmbConcepto.Text & "' "
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"

        ' 2. Llamamos al módulo. El módulo inyectará los valores en vTotalPresupuestoYTD y vTotalRealYTD
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' 🚨 RESPALDO DE SEGURIDAD INMEDIATO: Guardamos los valores calculados en variables locales
        ' para evitar que 'ActualizarEtiquetaDesviacion' u otro proceso del módulo los ponga a cero.
        Dim miPresupuestoYTD As Double = vTotalPresupuestoYTD
        Dim miRealYTD As Double = vTotalRealYTD

        ' 3. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
        ActualizarEtiquetaDesviacion()

        ' 4. Pintamos el resultado basándonos en la mallas calculadas
        If frmPresupuestos.DgvPresupuestos.Rows.Count > 0 Then

            ' REFUERZO DE VISIBILIDAD: Forzamos que todos los Labels de resultados se muestren siempre
            LblDesviacion.Visible = True
            LblMontoDesviacion.Visible = True
            LblObjetivo.Visible = True
            LblDesviacion.Enabled = True

            ' REFUERZO DE SEGURIDAD: Deducimos el tipo de concepto si la variable global falló
            Dim tipoConcepto As String = vTipoConceptoActual.Trim().ToUpper()

            If String.IsNullOrEmpty(tipoConcepto) Then
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Using cmd As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & CmbConcepto.Text.Replace("'", "''") & "'", con)
                        Try
                            con.Open()
                            Dim res As Object = cmd.ExecuteScalar()
                            If res IsNot Nothing Then tipoConcepto = res.ToString().Trim().ToUpper()
                        Catch
                        End Try
                    End Using
                End Using
            End If

            ' EVALUACIÓN ESTÁNDAR UNIVERSAL: GASTO o INGRESO
            Dim esGasto As Boolean = (tipoConcepto = "GASTO")
            Dim desviacionFinal As Double

            ' 🔥 CORRECCIÓN MATEMÁTICA DE SIGNOS SEGÚN EL TIPO
            If esGasto Then
                ' Para Gastos: Desviación = Presupuesto - Real (Positivo es ahorro, objetivo logrado)
                desviacionFinal = miPresupuestoYTD - miRealYTD
            Else
                ' Para Ingresos: En tu pantalla el real acumulado (miRealYTD) viene en NEGATIVO (ej: -2.956,90)
                ' Pasamos el Real a positivo puro para poder compararlo correctamente con el presupuesto positivo
                Dim realIngresoPositivo As Double = Math.Abs(miRealYTD)

                ' Desviación = Real - Presupuesto (Si el Real es mayor o igual al Presupuesto, es un ingreso logrado)
                desviacionFinal = realIngresoPositivo - miPresupuestoYTD
            End If

            ' CONFIGURACIÓN VISUAL DEL TEXTO SEGÚN EL AÑO
            Dim añoActualCalendario As Integer = DateTime.Now.Year

            If CInt(vAñoEjercicio) = añoActualCalendario Then
                ActualizarEtiquetaDesviacion()
            Else
                Dim textoAnual As String = rmse.GetString("LblDesviacion.Text")
                If String.IsNullOrEmpty(textoAnual) Then textoAnual = "Desviació Anual"
                LblDesviacion.Text = textoAnual & " " & vAñoEjercicio & " ="
            End If

            ' Mostramos la cifra final con formato "N2" para multiidioma
            LblMontoDesviacion.Text = desviacionFinal.ToString("N2") & " " & vMoneda

            ' CONTROL DE COLORES Y OBJETIVOS (Comportamiento financiero real)
            If desviacionFinal >= 0 Then
                ' GASTO: Gastaste menos de lo presupuestado (¡Logrado!)
                ' INGRESO: Ingresaste más de lo presupuestado (¡Logrado!)
                LblObjetivo.ForeColor = Color.DarkGreen
                LblObjetivo.Text = rmse.GetString("LblObjetivo.Text")
                If String.IsNullOrEmpty(LblObjetivo.Text) Then LblObjetivo.Text = "Objectiu Assolit!"
                LblMontoDesviacion.ForeColor = Color.DarkBlue
            Else
                ' GASTO: Te pasaste del presupuesto (No logrado)
                ' INGRESO: Ganaste menos de lo previsto (No logrado)
                LblObjetivo.ForeColor = Color.DarkRed
                LblObjetivo.Text = rmse.GetString("NoLogrado")
                If String.IsNullOrEmpty(LblObjetivo.Text) Then LblObjetivo.Text = "Objectiu No Assolit"
                LblMontoDesviacion.ForeColor = Color.Red
            End If
        Else
            ' Si el grid no tiene filas, limpiamos y ocultamos todo
            LblDesviacion.Enabled = False
            LblMontoDesviacion.Text = ""
            LblObjetivo.Visible = False
        End If
    End Sub

    Private Sub BtnGraficos2D_Click(sender As Object, e As EventArgs) Handles BtnGraficos2D.Click
        ' 1. Comprobamos si existe un identificador asociado (Instancia segura)
        If (frmGraficosPresupuestos Is Nothing) OrElse (Not frmGraficosPresupuestos.IsHandleCreated) Then
            frmGraficosPresupuestos = New GraficosPresupuestos
        End If

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
        'Vuelve a Refrecar el DataGrid y dejar los Btn de los Filtros sin Filtrar
        '************************************************************************
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        LblDesviacion.Visible = False
        LblMontoDesviacion.Visible = False
        LblObjetivo.Visible = False
        If DgvPresupuestos.RowCount - 1 >= 0 Then
            DgvPresupuestos.Rows(0).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(0).Cells(0)
        End If
    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        'Elimina las Filas Seleccionadas
        '*******************************
        For Each r As DataGridViewRow In DgvPresupuestos.SelectedRows
            If DgvPresupuestos.Rows.Count > 1 Then
                DgvPresupuestos.Rows.Remove(r)
            End If
        Next
        If DgvPresupuestos.Rows.Count > 1 Then
            FilaSelec = DgvPresupuestos.CurrentRow.Index
            For i = 0 To DgvPresupuestos.Rows.Count - 1
                DgvPresupuestos.Rows(i).Selected = False
            Next
            DgvPresupuestos.Select()
            DgvPresupuestos.CurrentRow.Selected = True
            DgvPresupuestos.Refresh()
        End If

        ' 1. Obtenemos el DataTable enlazado a tu DgvPresupuestos
        Dim dt As DataTable = CType(DgvPresupuestos.DataSource, DataTable)

        ' 2. Si hay filas, verificamos si la ÚLTIMA contiene el texto de "TOTAL" traducido
        If dt.Rows.Count > 0 Then
            Dim ultimaFila As DataRow = dt.Rows(dt.Rows.Count - 1)
            Dim textoTotalTraducido As String = resManager.GetString("TOTAL")

            ' Si la última fila en la columna 0 coincide con el recurso, la eliminamos
            If Convert.ToString(ultimaFila(0)) = textoTotalTraducido Then
                dt.Rows.Remove(ultimaFila)
            End If
        End If

        ' 3. Variables para acumular las sumas de las filas limpias
        Dim totalCol2 As Decimal = 0
        Dim totalCol3 As Decimal = 0

        ' 4. Recorremos las filas restantes para calcular los totales reales
        For Each fila As DataGridViewRow In DgvPresupuestos.Rows
            Dim valorCol2 As Decimal = 0
            Dim valorCol3 As Decimal = 0

            Decimal.TryParse(Convert.ToString(fila.Cells(2).Value), valorCol2)
            Decimal.TryParse(Convert.ToString(fila.Cells(3).Value), valorCol3)

            totalCol2 += valorCol2
            totalCol3 += valorCol3
        Next

        ' 5. Creamos la nueva fila de totales para el DataTable
        Dim nuevaFila As DataRow = dt.NewRow()

        ' 6. Asignamos el texto localizado mediante tu resManager y los totales calculados
        nuevaFila(0) = resManager.GetString("TOTAL")
        nuevaFila(2) = totalCol2
        nuevaFila(3) = totalCol3

        ' 7. Añadimos la fila final actualizada al DataTable
        dt.Rows.Add(nuevaFila)
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        ' 1. Aseguramos que haya una fila seleccionada en el Grid de forma segura
        If DgvPresupuestos.CurrentRow Is Nothing Then Exit Sub

        ' 2. Mensaje de confirmación traducible (Apunta a tus llaves globales del ResX)
        Dim msgPregunta As String = rmse.GetString("PreguntaEliminarPresupuesto")
        Dim titPregunta As String = rmse.GetString("TituloEliminarPresupuesto")
        Dim respuesta As MsgBoxResult = MsgBox(msgPregunta, vbQuestion + vbYesNo + vbDefaultButton2, titPregunta)

        If respuesta = vbYes Then
            filaActual = DgvPresupuestos.CurrentRow.Index
            Dim conceptoVisible As String = DgvPresupuestos.Rows(filaActual).Cells(0).Value.ToString().Trim()

            ' 🔄 REVERTIR EL IDIOMA DE FORMA DIRECTA Y LIMPIA
            Dim conceptoOriginalMDB As String = conceptoVisible

            If Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
                ' Buscamos la clave en el diccionario de recursos pasando el valor traducido
                Dim claveRecurso As String = resManager.GetString(conceptoVisible)
                If Not String.IsNullOrEmpty(claveRecurso) Then
                    conceptoOriginalMDB = claveRecurso.Replace("_", " ")
                End If
            End If

            ' 3. OPERACIÓN DE BORRADO SEGURA Y QUIRÚRGICA por Concepto y Año
            Dim sqlDelete As String = "DELETE FROM presupuesto WHERE ConceptoPRE = ? AND EjercicioPRE = ?"

            Using conexion As New OleDbConnection(conexion1.ConnectionString)
                Using cmd As New OleDbCommand(sqlDelete, conexion)
                    cmd.Parameters.AddWithValue("@con", conceptoOriginalMDB)
                    cmd.Parameters.AddWithValue("@eje", CInt(vAñoEjercicio)) ' Filtro crítico por año

                    Try
                        conexion.Open()
                        cmd.ExecuteNonQuery()
                        Dim msgBorrados As String = resManager.GetString("PresupuestosBorradosExito")
                        MsgBox(msgBorrados, vbInformation)
                    Catch ex As Exception
                        Dim msgError As String = resManager.GetString("ErrorEliminarPresupuestos")
                        MsgBox(msgError & vbNewLine & ex.Message, vbCritical)
                    End Try
                End Using
            End Using

            ' 4. RECARGA DEL GRID DE PRESUPUESTOS (Con tu norma exacta de columnas fija e intacta)
            vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
            vtipoSql += " WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
            vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
            vtipoGrid = "PRESUPUESTOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")

            ' Evaluamos si corresponde "Parcial" o "Anual" tras la recarga
            ActualizarEtiquetaDesviacion()

            ' Forzamos la limpieza de los cuadros de desviación si el Grid se quedó vacío
            If DgvPresupuestos.Rows.Count = 0 Then
                LblDesviacion.Enabled = False
                LblMontoDesviacion.Text = ""
                LblObjetivo.Visible = False
            End If
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub DgvPresupuestos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgvPresupuestos.CellFormatting
        ' Comprobamos que no sea la fila de cabecera
        If e.RowIndex >= 0 Then
            Dim dgv As DataGridView = CType(sender, DataGridView)

            ' 🔄 MULTIIDIOMA SEGURO: Recuperamos la palabra "TOTAL" traducida según el idioma actual
            Dim textoTotalTraducido As String = resManager.GetString("TOTAL")
            If String.IsNullOrEmpty(textoTotalTraducido) Then textoTotalTraducido = "TOTAL" ' Salvavidas por si acaso

            ' Comprobamos el valor de la columna 0 de forma segura
            If dgv.Rows(e.RowIndex).Cells(0).Value IsNot Nothing Then
                Dim valorCelda As String = dgv.Rows(e.RowIndex).Cells(0).Value.ToString().Trim().ToUpper()

                ' Comparamos de forma insensible a mayúsculas/minúsculas contra el término traducido y el base
                If valorCelda = textoTotalTraducido.ToUpper() OrElse valorCelda = "TOTAL" Then

                    ' Aplicamos el fondo gris, texto negro y negrita de forma persistente
                    e.CellStyle.BackColor = Color.LightGray
                    e.CellStyle.ForeColor = Color.Black
                    e.CellStyle.Font = New Font("Tahoma", 9, FontStyle.Bold)

                End If
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
            MsgBox(resManager.GetString("ErrorLimpiarTemporal") & ex.Message)
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

            ' Saltamos la fila de totales que genera la propia pantalla
            If fila.Cells(0).Value IsNot Nothing AndAlso fila.Cells(0).Value.ToString().Trim().ToUpper() = "TOTAL" Then
                Continue For
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