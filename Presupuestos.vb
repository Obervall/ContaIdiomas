Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class Presupuestos

    Public vtipoSql, vtipoGrid, vConcepto, vBorrarPresu, vAñadir, vAñadir2 As String
    Public vTmpprint As String
    Public vImporteConcepto, vImporteConcepto2, vExistenteImporteConcepto, vNewImporteConcepto As Double
    Public PrintLine, Contador As Integer
    Public vTipoConceptoActual As String = ""
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())


    Private Sub Presupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(9) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnGraficos, "Mostrar Gráficos")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSalir, "Salir de Presupuestos")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnFiltroConcepto, "Aplica el filtro a los Registros")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnSinFiltroConcepto, "Quitar el filtro a los Registros")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnImprimir, "Imprimir")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnPrimero, "Ir al Primer Registro")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnAnterior, "Ir al Anterior Registro")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnSiguiente, "Ir al Siguiente Registro")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnUltimo, "Ir al Ultimo Registro")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnEliminarRegistro, "Eliminar Concepto en Presupuestos")


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
            MsgBox("No se han podido cargar los Conceptos en el Combo, revise el código !!!")
            MsgBox(ex.ToString)
        End Try

        ' Llenar Grid de PRESUPUESTOS
        '****************************
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        ' 2. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
        ActualizarEtiquetaDesviacion()
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
        TxtDesviacion.Text = ""
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        EjecutarCalculoYDesviacion()
    End Sub

    Private Sub CmbTipoConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 1. Aseguramos que haya un texto seleccionado
        vConcepto = CmbConcepto.Text.ToString().Trim()
        If String.IsNullOrEmpty(vConcepto) Then Exit Sub

        ' 🔄 REVERTIR EL IDIOMA: Buscamos el nombre original en español para consultar la MDB
        Dim conceptoOriginalMDB As String = vConcepto
        If Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
            Dim cultura As New System.Globalization.CultureInfo(My.Settings.CulturaUsuario)
            Dim recursos As System.Resources.ResourceSet = resManager.GetResourceSet(cultura, True, True)
            If recursos IsNot Nothing Then
                For Each elemento As System.Collections.DictionaryEntry In recursos
                    If elemento.Value.ToString().Trim().ToUpper() = vConcepto.ToUpper() Then
                        conceptoOriginalMDB = elemento.Key.ToString().Replace("_", " ")
                        Exit For
                    End If
                Next
            End If
        End If

        ' 2. Consulta combinada: Rescatamos la Descripción Y el Tipo original del concepto de golpe
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
                            ' Guardamos el tipo real (GASTO/INGRESO/ESPECIAL) en nuestra variable global
                            vTipoConceptoActual = dr("TipoCON").ToString().Trim().ToUpper()
                        End If
                    End Using
                Catch ex As Exception
                    MsgBox("Error al cargar detalles del concepto: " & ex.Message)
                End Try
            End Using
        End Using

        ' 3. Si el botón de aplicar filtro está activo (el usuario ya estaba filtrando), actualizamos
        If BtnFiltroConcepto.Enabled = False Then
            EjecutarCalculoYDesviacion()
        End If
    End Sub


    ''' <summary>
    ''' MÉTODO CENTRALIZADO PROFESIONAL: Ejecuta la carga y pinta los resultados YTD calculados por el módulo
    ''' </summary>
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

        ' 2. Llamamos al módulo. Al procesar el Grid, el módulo inyectará los valores acumulados YTD 
        ' en las variables globales 'vSaldoAnualPresupuesto' y 'vSaldoAnualReal'.
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        ' 2. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
        ActualizarEtiquetaDesviacion()

        ' 3. Pintamos el resultado basándonos en la mallas calculadas por el módulo
        If frmPresupuestos.DgvPresupuestos.Rows.Count > 0 Then
            LblDesviacion.Enabled = True

            ' 🚨 REFUERZO DE SEGURIDAD: Si por algún motivo la variable quedó vacía, 
            ' leemos el tipo directamente de la primera fila del DataGrid (Celda 4 o donde tengas el Tipo si venía en la Query)
            Dim tipoConcepto As String = vTipoConceptoActual.Trim().ToUpper()

            ' Si la variable global falló, hacemos una deducción de emergencia basada en el nombre del concepto
            If String.IsNullOrEmpty(tipoConcepto) Then
                ' Buscamos de forma quirúrgica en la MDB el tipo real
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

            ' Evaluamos si es un gasto (Soporta Castellano, Catalán, Alemán o Inglés)
            Dim esGasto As Boolean = (tipoConcepto = "GASTO" OrElse
                                      tipoConcepto = "DESPESA" OrElse
                                      tipoConcepto = "AUSGABE" OrElse
                                      tipoConcepto = "EXPENSE" OrElse
                                      tipoConcepto.Contains("GAST"))

            Dim desviacionFinal As Double = 0

            ' B. APLICACIÓN ESTRICTA DE LA FÓRMULA CONTABLE
            If esGasto Then
                ' En GASTOS: Si el Presupuesto acumulado es MAYOR que el Gasto Real, hemos ahorrado (Desviación Positiva)
                desviacionFinal = vSaldoAnualPresupuesto - vSaldoAnualReal
            Else
                ' En INGRESOS: Si el Real es MAYOR que el Presupuesto, hemos ganado más (Desviación Positiva)
                desviacionFinal = Math.Abs(vSaldoAnualReal) - vSaldoAnualPresupuesto
            End If

            ' 🧪 MENSAJE DE DIAGNÓSTICO EN TIEMPO REAL (Para ver qué calcula el PC)
            ' Puedes comentar o borrar esta línea en cuanto veamos qué falla
            'MessageBox.Show("Concepto: " & CmbConcepto.Text & vbCrLf &
            '                "Tipo Detectado: " & tipoConcepto & " (¿Es Gasto?: " & esGasto.ToString() & ")" & vbCrLf &
            '                "Suma Presupuesto YTD: " & vSaldoAnualPresupuesto.ToString("F2") & vbCrLf &
            '                "Suma Real YTD: " & vSaldoAnualReal.ToString("F2") & vbCrLf &
            '                "Resta Final: " & desviacionFinal.ToString("F2"), "Diagnóstico Técnico")

            ' C. Mostramos la cifra en la casilla inferior
            LblDesviacion.Enabled = True
            TxtDesviacion.Text = Format(desviacionFinal, "###,##0.00")
            LblObjetivo.Visible = True

            ' D. EVALUACIÓN REAL DEL OBJETIVO
            ' Si la desviación es POSITIVA (>= 0), el presupuesto va bien
            If desviacionFinal >= 0 Then
                LblObjetivo.ForeColor = Color.DarkGreen
                LblObjetivo.Text = "Objetivo Logrado!"
            Else
                ' Si es NEGATIVA (< 0), nos hemos desviado del plan
                LblObjetivo.ForeColor = Color.DarkRed
                LblObjetivo.Text = "Objetivo NO Logrado!"
            End If
        Else
            LblDesviacion.Enabled = False
            TxtDesviacion.Text = ""
            LblObjetivo.Visible = False
        End If
    End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmGraficosPresupuestos Is Nothing) OrElse (Not frmGraficosPresupuestos.IsHandleCreated)) Then
            frmGraficosPresupuestos = New GraficosPresupuestos
        End If
        ' Llamamos al formulario de manera modal.
        frmGraficosPresupuestos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
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
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = 0
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = vFilaActual - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click

        ' Iniciamos Tabla Tmpprint
        ' ************************
        vTmpprint = "DELETE FROM tmpprint"
        cmdMdb1cr.CommandText = vTmpprint
        Try
            cmdMdb1cr.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("No se han podido eliminar los registros en tmpprint, revise el código !!!")
            MsgBox(ex.ToString)
        End Try

        ' Ordenamos la columna Concepto, antes de calcular los totales parciales.
        ' ***********************************************************************
        frmPresupuestos.DgvPresupuestos.Sort(frmPresupuestos.DgvPresupuestos.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

        ' Llenamos la tabla tmpprint con los Conceptos Agrupados desde DgvPresupuestos
        ' ****************************************************************************
        vNombreConcepto = ""
        For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
            ' 🚨 EXCLUSIÓN CRÍTICA: Si es la fila de totales que agregamos al Grid, la ignoramos por completo
            If fila.Cells(0).Value IsNot Nothing AndAlso fila.Cells(0).Value.ToString().ToUpper() = "TOTAL" Then Continue For
            If fila.IsNewRow Then Continue For

            vImporteConcepto = Convert.ToDouble(fila.Cells(3).Value)

            If vNombreConcepto <> fila.Cells(0).Value.ToString Then
                vNombreConcepto = fila.Cells(0).Value.ToString
                vImporteConcepto = Convert.ToDouble(fila.Cells(3).Value)

                ' Buscamos la suma total del Concepto
                ' ***********************************
                vtipoSql = "SELECT * FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                vtipoSql += " And apuntes.ConceptoAPU = '" & vNombreConcepto & "' "
                vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC"
                LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")  ' Agrupado Conceptos

                ' CORRECCIÓN DE TIPO: Inicializamos como Double numérico puro, nunca como "0,00" string
                Dim acumuladoReal As Double = 0
                For Each filas As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                    If filas.Cells(4).Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(filas.Cells(4).Value.ToString()) Then
                        Dim valorApunte As Double = 0
                        If Double.TryParse(filas.Cells(4).Value.ToString(), valorApunte) Then
                            acumuladoReal += Math.Abs(valorApunte)
                        End If
                    End If
                Next

                '' Guardamos usando formato de punto flotante estándar para SQL
                'Dim strReal As String = acumuladoReal.ToString("FFFF", System.Globalization.CultureInfo.InvariantCulture)
                'Dim strPresu As String = vImporteConcepto.ToString("FFFF", System.Globalization.CultureInfo.InvariantCulture)

                ' Determinamos si el concepto actual del Grid es un Gasto o un Ingreso
                Dim tipoConcepto As String = "G" ' Por defecto Gasto
                ' Si tu lógica determina que el concepto no es gasto (puedes adaptarlo según tus flags o el nombre del concepto)
                If vNombreConcepto.ToUpper().Contains("INGRESOS") OrElse vNombreConcepto.ToUpper().Contains("INGRESSOS") Then
                    tipoConcepto = "I"
                End If

                vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                          "VALUES (#2023-01-11#, ?, '', ?, '', ?, ?)"

                cmdMdb1cr.CommandText = vAñadir
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@ConceptoTMP", vNombreConcepto)
                cmdMdb1cr.Parameters.AddWithValue("@CuentaTMP", tipoConcepto) ' Guardamos "I" o "G"
                cmdMdb1cr.Parameters.AddWithValue("@ImporteTMP", acumuladoReal)
                cmdMdb1cr.Parameters.AddWithValue("@SaldoTMP", vImporteConcepto)


                '' Guardamos usando tipos numéricos puros (Double) mediante parámetros para evitar fallos de formato en la MDB
                'vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                '          "VALUES (#2023-01-11#, ?, '', '', '', ?, ?)"

                'cmdMdb1cr.CommandText = vAñadir
                'cmdMdb1cr.Parameters.Clear()

                '' En OleDb para Access, los parámetros se agregan ESTRICTAMENTE en el mismo orden en que aparecen los signos de interrogación '?'
                'cmdMdb1cr.Parameters.AddWithValue("@ConceptoTMP", vNombreConcepto)
                'cmdMdb1cr.Parameters.AddWithValue("@ImporteTMP", acumuladoReal)   ' Pasa el número puro (Double), .NET gestiona el signo/separador
                'cmdMdb1cr.Parameters.AddWithValue("@SaldoTMP", vImporteConcepto) ' Pasa el número puro (Double)

                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("Error al grabar en tmpprint de la MDB: " & vbCrLf & ex.Message)
                End Try

            Else
                Dim vExistenteImporteConcepto As Double = 0
                cmdMdb1cr.CommandType = CommandType.Text
                ' Usamos parámetro para el concepto en la lectura, evitando fallos con comillas o caracteres extra
                cmdMdb1cr.CommandText = "SELECT * FROM tmpprint WHERE tmpprint.ConceptoTMP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("@ConceptoBusqueda", vNombreConcepto)

                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            ' Leemos el valor numérico de forma nativa y segura
                            If Not IsDBNull(drMdb1("SaldoTMP")) Then
                                vExistenteImporteConcepto = Convert.ToDouble(drMdb1("SaldoTMP"))
                            End If
                        End While
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    MsgBox("No se ha podido leer el registro en tmpprint, revise el código !!!")
                    MsgBox(ex.ToString)
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                End Try

                ' Calculamos la suma del presupuesto acumulado
                Dim vNewImporteConcepto As Double = vImporteConcepto + vExistenteImporteConcepto

                ' CORRECCIÓN CON PARÁMETROS PARA EL UPDATE
                ' ***************************************
                vAñadir2 = "UPDATE tmpprint SET SaldoTMP = ? WHERE tmpprint.ConceptoTMP = ?"
                cmdMdb1cr.CommandText = vAñadir2
                cmdMdb1cr.Parameters.Clear()

                ' Recuerda que en OleDb el orden de los parámetros debe ser idéntico al de los signos '?'
                cmdMdb1cr.Parameters.AddWithValue("@NuevoSaldo", vNewImporteConcepto) ' Primer '?'
                cmdMdb1cr.Parameters.AddWithValue("@ConceptoFiltro", vNombreConcepto)  ' Segundo '?'

                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox("No se ha podido actualizar el registro en tmpprint, revise el código !!!")
                    MsgBox(ex.ToString)
                End Try
            End If
        Next

        vtipoSql = "SELECT * FROM tmpprint"
        vtipoSql += " ORDER BY tmpprint.ConceptoTMP ASC"
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
        Dim FuenteTitulo As New Font("Segoe UI", 16, FontStyle.Bold)
        Dim FuenteSeccion As New Font("Segoe UI", 11, FontStyle.Bold)
        Dim FuenteNegrita As New Font("Segoe UI", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Segoe UI", 9, FontStyle.Regular)

        Dim sfDerecha As New StringFormat With {.Alignment = StringAlignment.Far}

        ' Encabezado del Reporte
        ' **********************
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteDetalles, Brushes.Gray, 750, 40, sfDerecha)

        Dim tituloReporte As String = If(BtnFiltroConcepto.Enabled = False, "Informe de Presupuestos - Filtrado", "ESTADO DE RENDIMIENTO PRESUPUESTARIO")
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
        e.Graphics.DrawString("1. ESTRUCTURA DE INGRESOS", FuenteSeccion, Brushes.DarkGreen, colConceptoX, posY)
        posY += 25

        e.Graphics.DrawString("Concepto:", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString("Real YTD:", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString("Presupuesto:", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        e.Graphics.DrawString("Desviación:", FuenteNegrita, Brushes.Black, colDesvX, posY, sfDerecha)
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

            totalIngresosReal += valReal
            totalIngresosPresu += valPresu

            Dim desv As Double = valReal - valPresu
            Dim brushDesv As Brush = If(desv >= 0, Brushes.Green, Brushes.Red)

            e.Graphics.DrawString(concepto, FuenteDetalles, Brushes.Black, colConceptoX, posY)
            e.Graphics.DrawString(valReal.ToString("###,##0.00") & " €", FuenteDetalles, Brushes.Black, colRealX, posY, sfDerecha)
            e.Graphics.DrawString(valPresu.ToString("###,##0.00") & " €", FuenteDetalles, Brushes.Black, colPresuX, posY, sfDerecha)
            e.Graphics.DrawString((If(desv >= 0, "+", "")) & desv.ToString("###,##0.00") & " €", FuenteDetalles, brushDesv, colDesvX, posY, sfDerecha)
            posY += 22
        Next

        If Not tieneIngresos Then
            e.Graphics.DrawString("No se registraron movimientos de ingresos.", FuenteDetalles, Brushes.Gray, colConceptoX + 20, posY)
            posY += 22
        End If

        e.Graphics.DrawLine(Pens.LightGray, colConceptoX, posY, colDesvX, posY)
        posY += 5
        e.Graphics.DrawString("TOTAL INGRESOS", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString(totalIngresosReal.ToString("###,##0.00") & " €", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(totalIngresosPresu.ToString("###,##0.00") & " €", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        Dim desvIngTotal As Double = totalIngresosReal - totalIngresosPresu
        e.Graphics.DrawString(desvIngTotal.ToString("###,##0.00") & " €", FuenteNegrita, If(desvIngTotal >= 0, Brushes.Green, Brushes.Red), colDesvX, posY, sfDerecha)

        posY += 40

        ' =========================================================================
        ' BLOQUE 2: GASTOS
        ' =========================================================================
        e.Graphics.DrawString("2. ESTRUCTURA DE GASTOS", FuenteSeccion, Brushes.DarkRed, colConceptoX, posY)
        posY += 25

        e.Graphics.DrawString("Concepto:", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString("Real YTD:", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString("Presupuesto:", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        e.Graphics.DrawString("Desviación (Ahorro):", FuenteNegrita, Brushes.Black, colDesvX, posY, sfDerecha)
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
            e.Graphics.DrawString(valReal.ToString("###,##0.00") & " €", FuenteDetalles, Brushes.Black, colRealX, posY, sfDerecha)
            e.Graphics.DrawString(valPresu.ToString("###,##0.00") & " €", FuenteDetalles, Brushes.Black, colPresuX, posY, sfDerecha)
            e.Graphics.DrawString((If(desv >= 0, "+", "")) & desv.ToString("###,##0.00") & " €", FuenteDetalles, brushDesv, colDesvX, posY, sfDerecha)
            posY += 22
        Next

        If Not tieneGastos Then
            e.Graphics.DrawString("No se registraron movimientos de gastos.", FuenteDetalles, Brushes.Gray, colConceptoX + 20, posY)
            posY += 22
        End If

        e.Graphics.DrawLine(Pens.LightGray, colConceptoX, posY, colDesvX, posY)
        posY += 5
        e.Graphics.DrawString("TOTAL GASTOS", FuenteNegrita, Brushes.Black, colConceptoX, posY)
        e.Graphics.DrawString(totalGastosReal.ToString("###,##0.00") & " €", FuenteNegrita, Brushes.Black, colRealX, posY, sfDerecha)
        e.Graphics.DrawString(totalGastosPresu.ToString("###,##0.00") & " €", FuenteNegrita, Brushes.Black, colPresuX, posY, sfDerecha)
        Dim desvGasTotal As Double = totalGastosPresu - totalGastosReal
        e.Graphics.DrawString(desvGasTotal.ToString("###,##0.00") & " €", FuenteNegrita, If(desvGasTotal >= 0, Brushes.Green, Brushes.Red), colDesvX, posY, sfDerecha)

        posY += 50

        ' =========================================================================
        ' BLOQUE 3: CUADRO DE RESUMEN EJECUTIVO (Diseño vertical anti-solapamiento)
        ' =========================================================================
        ' Ampliamos el alto del recuadro a 105 para dar espacio al formato de lista vertical
        Dim fondoResumen As New Rectangle(colConceptoX, posY, colDesvX - colConceptoX, 105)
        e.Graphics.FillRectangle(Brushes.GhostWhite, fondoResumen)
        e.Graphics.DrawRectangle(Pens.SlateGray, fondoResumen)

        posY += 15
        e.Graphics.DrawString("RESULTADO NETO DEL PERIODO (Ingresos - Gastos)", FuenteSeccion, Brushes.Black, colConceptoX + 15, posY)

        ' Calculamos los importes netos
        Dim netoReal As Double = totalIngresosReal - totalGastosReal
        Dim netoPresu As Double = totalIngresosPresu - totalGastosPresu
        Dim desvNetalGlobal As Double = netoReal - netoPresu

        ' FILA 1: Resultado Real (Bajo su columna correspondiente)
        posY += 25
        e.Graphics.DrawString("Resultado Real Neto:", FuenteNegrita, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString(netoReal.ToString("###,##0.00") & " €", FuenteNegrita, If(netoReal >= 0, Brushes.DarkGreen, Brushes.DarkRed), colDesvX - 15, posY, sfDerecha)

        ' FILA 2: Resultado Presupuestado
        posY += 20
        e.Graphics.DrawString("Resultado Presupuestado Neto:", FuenteDetalles, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString(netoPresu.ToString("###,##0.00") & " €", FuenteDetalles, Brushes.Black, colDesvX - 15, posY, sfDerecha)

        ' FILA 3: Desviación Global Destacada
        posY += 22
        e.Graphics.DrawString("DESVIACIÓN GLOBAL COMPLETA:", FuenteNegrita, Brushes.Black, colConceptoX + 15, posY)
        e.Graphics.DrawString((If(desvNetalGlobal >= 0, "+", "")) & desvNetalGlobal.ToString("###,##0.00") & " €", FuenteSeccion, If(desvNetalGlobal >= 0, Brushes.Green, Brushes.Red), colDesvX - 15, posY - 4, sfDerecha)

        ' =========================================================================
        ' PIE DE PÁGINA: NUMERACIÓN AUTOMÁTICA
        ' =========================================================================
        Dim nPagina As Integer = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        frmImprimirForm.LblNumeroPagina.Text = nPagina.ToString()

        Dim textoPagina As String = If(resManager.GetString("Pagina"), "Pág.") & " " & nPagina.ToString()
        e.Graphics.DrawString(textoPagina, FuenteDetalles, Brushes.DimGray, 750, 1050, sfDerecha)

        e.HasMorePages = False
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        ' 1. Aseguramos que haya una fila seleccionada en el Grid de forma segura
        If DgvPresupuestos.CurrentRow Is Nothing Then Exit Sub

        ' 2. Mensaje de confirmación traducible (Apunta a tus llaves globales del ResX)
        Dim msgPregunta As String = resManager.GetString("PreguntaEliminarPresupuesto")
        If String.IsNullOrEmpty(msgPregunta) Then msgPregunta = "¿Está seguro de eliminar el presupuesto del concepto seleccionado para este ejercicio?"

        Dim titPregunta As String = resManager.GetString("TituloEliminarPresupuesto")
        If String.IsNullOrEmpty(titPregunta) Then titPregunta = "Eliminar presupuesto"

        respuesta = MsgBox(msgPregunta, vbQuestion + vbYesNo + vbDefaultButton2, titPregunta)

        If respuesta = vbYes Then
            filaActual = DgvPresupuestos.CurrentRow.Index
            Dim conceptoVisible As String = DgvPresupuestos.Rows(filaActual).Cells(0).Value.ToString().Trim()

            ' 🔄 REVERTIR EL IDIOMA: Buscamos el nombre original en español guardado en la MDB
            Dim conceptoOriginalMDB As String = conceptoVisible

            If Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
                Dim cultura As New System.Globalization.CultureInfo(My.Settings.CulturaUsuario)
                Dim recursos As System.Resources.ResourceSet = resManager.GetResourceSet(cultura, True, True)

                If recursos IsNot Nothing Then
                    For Each elemento As System.Collections.DictionaryEntry In recursos
                        If elemento.Value.ToString().Trim().ToUpper() = conceptoVisible.ToUpper() Then
                            ' Encontramos el código real (ej: cambia "RENT" por "ALQUILER")
                            conceptoOriginalMDB = elemento.Key.ToString().Replace("_", " ")
                            Exit For
                        End If
                    Next
                End If
            End If

            ' 3. OPERACIÓN DE BORRADO SEGURA Y QUIRÚRGICA por Concepto y Año
            ' Usamos un comando local parametrizado para evitar inyecciones y fallos de formato en Access
            Dim sqlDelete As String = "DELETE FROM presupuesto WHERE ConceptoPRE = ? AND EjercicioPRE = ?"

            Using conexion As New OleDbConnection(conexion1.ConnectionString)
                Using cmd As New OleDbCommand(sqlDelete, conexion)
                    cmd.Parameters.AddWithValue("@con", conceptoOriginalMDB)
                    cmd.Parameters.AddWithValue("@eje", CInt(vAñoEjercicio)) ' Filtro crítico por año

                    Try
                        conexion.Open()
                        cmd.ExecuteNonQuery()

                        Dim msgBorrados As String = resManager.GetString("PresupuestosBorradosExito")
                        If String.IsNullOrEmpty(msgBorrados) Then msgBorrados = "Registros en Presupuestos eliminados correctamente."
                        MsgBox(msgBorrados, vbInformation)

                    Catch ex As Exception
                        Dim msgError As String = resManager.GetString("ErrorEliminarPresupuestos")
                        If String.IsNullOrEmpty(msgError) Then msgError = "No se han podido eliminar los registros en Presupuestos."
                        MsgBox(msgError & vbNewLine & ex.Message, vbCritical)
                    End Try
                End Using
            End Using

            ' 4. RECARGA DEL GRID DE PRESUPUESTOS (Tu módulo lo procesará e indexará a la velocidad de la luz)
            vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
            vtipoSql += " WHERE "
            vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
            vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
            vtipoGrid = "PRESUPUESTOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            ' 2. ACTUAR SOBRE LA ETIQUETA: Evaluamos si corresponde "Parcial" o "Anual"
            ActualizarEtiquetaDesviacion()

            ' Forzamos la limpieza de los cuadros de desviación si el Grid se quedó vacío
            If DgvPresupuestos.Rows.Count = 0 Then
                LblDesviacion.Enabled = False
                TxtDesviacion.Text = ""
                LblObjetivo.Visible = False
            End If
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = vFilaActual + 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = DgvPresupuestos.RowCount - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub DgvPresupuestos_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DgvPresupuestos.CellFormatting
        ' Comprobamos que no sea la fila de cabecera
        If e.RowIndex >= 0 Then
            Dim dgv As DataGridView = CType(sender, DataGridView)

            ' Si la celda de la columna 0 contiene la palabra "TOTAL", pintamos toda la fila
            If dgv.Rows(e.RowIndex).Cells(0).Value IsNot Nothing AndAlso
           dgv.Rows(e.RowIndex).Cells(0).Value.ToString() = "TOTAL" Then

                ' Aplicamos el fondo gris, texto negro y negrita de forma persistente
                e.CellStyle.BackColor = Color.LightGray
                e.CellStyle.ForeColor = Color.Black
                e.CellStyle.Font = New Font("Tahoma", 9, FontStyle.Bold)

            End If
        End If
    End Sub

    Private Sub ActualizarEtiquetaDesviacion()
        Dim añoActualCalendario As Integer = DateTime.Now.Year

        ' Comprobamos si el ejercicio consultado es el año en curso
        If CInt(vAñoEjercicio) = añoActualCalendario Then
            ' Si es el año actual, la desviación es parcial (YTD)
            LblDesviacion.Text = If(resManager.GetString("DesviacionParcial"), "Desviación Parcial")
        Else
            ' Si es un año pasado o diferente, es la desviación de todo el año
            LblDesviacion.Text = If(resManager.GetString("DesviacionAnual"), "Desviación Anual")
        End If
    End Sub

End Class