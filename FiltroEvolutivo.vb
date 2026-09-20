Imports System.Data
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms


Public Class FiltroEvolutivo

    Public Property TipoInforme As String ' "CUENTAS" o "CONCEPTOS"
    Private dtDatosInforme As System.Data.DataTable
    Private cargandoFormulario As Boolean = True
    Private vTxtNombreElementoReporte As String
    Private PrintLine, Contador As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Function ObtenerDatosEvolutivos(idElemento As Integer, mes As Integer, dia As Integer) As DataTable
        Dim dtFinal As New DataTable()
        dtFinal.Columns.Add("Anio", GetType(Integer))

        ' 1. Seleccionamos el nombre de la columna según el tipo de informe
        If Me.TipoInforme = "CUENTAS" Then
            dtFinal.Columns.Add("SaldoAFecha", GetType(Decimal))
        Else
            dtFinal.Columns.Add("TotalAcumulado", GetType(Decimal))
        End If

        ' 2. Traemos los ejercicios reales de la BD ordenados desde el más antiguo al más reciente
        Dim dtEjercicios As New DataTable()
        Using cmdEje As New OleDbCommand("SELECT EjercicioEJE FROM ejercicios WHERE EjercicioEJE > 0 AND EjercicioEJE <= " & vAñoEjercicio & " ORDER BY EjercicioEJE ASC", conexion1)
            Using adapterEje As New OleDbDataAdapter(cmdEje)
                Try
                    If conexion1.State = ConnectionState.Closed Then conexion1.Open()
                    adapterEje.Fill(dtEjercicios)
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorLeerEjercicios") & ": " & ex.Message, MsgBoxStyle.Critical)
                    Return dtFinal
                End Try
            End Using
        End Using

        ' =========================================================================
        ' 📊 MUNDO A: PROCESAMIENTO PARA REPORTE DE CUENTAS (Tu motor de simulación)
        ' =========================================================================
        If Me.TipoInforme = "CUENTAS" Then
            Dim saldoArrastradoPasado As Decimal = 0
            Dim esPrimerAnio As Boolean = True

            For Each rowEje As DataRow In dtEjercicios.Rows
                Dim anioEvaluar As Integer = Convert.ToInt32(rowEje("EjercicioEJE"))
                Dim movimientosHastaFecha As Decimal = 0
                Dim movimientosTodoElAnio As Decimal = 0

                ' A) Movimientos reales de este año hasta la fecha de corte
                Dim sqlCorte As String = "SELECT SUM(ImporteAPU) FROM apuntes WHERE CuentaAPU = ? AND EjercicioAPU = ? AND FechaAPU < DateSerial(?, ?, ?) + 1"
                Using cmdCorte As New OleDbCommand(sqlCorte, conexion1)
                    cmdCorte.Parameters.AddWithValue("@Cta", idElemento)
                    cmdCorte.Parameters.AddWithValue("@Eje", anioEvaluar)
                    cmdCorte.Parameters.AddWithValue("@EjeM", anioEvaluar)
                    cmdCorte.Parameters.AddWithValue("@Mes", mes)
                    cmdCorte.Parameters.AddWithValue("@Dia", dia)
                    Dim res = cmdCorte.ExecuteScalar()
                    If res IsNot DBNull.Value AndAlso res IsNot Nothing Then movimientosHastaFecha = Convert.ToDecimal(res)
                End Using

                ' B) Movimientos reales de todo el año completo para el arrastre
                Dim sqlTotalAnio As String = "SELECT SUM(ImporteAPU) FROM apuntes WHERE CuentaAPU = ? AND EjercicioAPU = ?"
                Using cmdTotal As New OleDbCommand(sqlTotalAnio, conexion1)
                    cmdTotal.Parameters.AddWithValue("@Cta", idElemento)
                    cmdTotal.Parameters.AddWithValue("@Eje", anioEvaluar)
                    Dim res = cmdTotal.ExecuteScalar()
                    If res IsNot DBNull.Value AndAlso res IsNot Nothing Then movimientosTodoElAnio = Convert.ToDecimal(res)
                End Using

                ' C) Arrastre continuo cronológico
                Dim saldoAFechaFinal As Decimal = 0
                If esPrimerAnio Then
                    saldoAFechaFinal = movimientosHastaFecha
                    saldoArrastradoPasado = movimientosTodoElAnio
                    esPrimerAnio = False
                Else
                    saldoAFechaFinal = saldoArrastradoPasado + movimientosHastaFecha
                    saldoArrastradoPasado = saldoArrastradoPasado + movimientosTodoElAnio
                End If

                dtFinal.Rows.Add(anioEvaluar, saldoAFechaFinal)
            Next

            ' D) Parche contable del ejercicio activo en curso (Con tu apunte físico ID 89)
            If dtFinal.Rows.Count > 0 Then
                Try
                    Dim saldoRealAnioActual As Decimal = 0
                    Dim sqlAnioActual As String = "SELECT SUM(ImporteAPU) FROM apuntes WHERE CuentaAPU = ? AND EjercicioAPU = ? AND FechaAPU < DateSerial(?, ?, ?) + 1"
                    Using cmdActual As New OleDbCommand(sqlAnioActual, conexion1)
                        cmdActual.Parameters.AddWithValue("@Cuenta", idElemento)
                        cmdActual.Parameters.AddWithValue("@Eje1", vAñoEjercicio)
                        cmdActual.Parameters.AddWithValue("@Eje2", vAñoEjercicio)
                        cmdActual.Parameters.AddWithValue("@Mes", mes)
                        cmdActual.Parameters.AddWithValue("@Dia", dia)
                        Dim resultado As Object = cmdActual.ExecuteScalar()
                        If resultado IsNot DBNull.Value AndAlso resultado IsNot Nothing Then saldoRealAnioActual = Convert.ToDecimal(resultado)
                    End Using
                    dtFinal.Rows(dtFinal.Rows.Count - 1)("SaldoAFecha") = saldoRealAnioActual
                Catch : End Try
            End If

            ' =========================================================================
            ' 🏷️ MUNDO B: PROCESAMIENTO PARA REPORTE DE CONCEPTOS (Año natural estricto)
            ' =========================================================================
        Else
            For Each rowEje As DataRow In dtEjercicios.Rows
                Dim anioEvaluar As Integer = Convert.ToInt32(rowEje("EjercicioEJE"))
                Dim totalConceptoAnio As Decimal = 0

                ' Suma pura de lo gastado/ingresado en este concepto únicamente en este año hasta la fecha
                Dim sqlConcepto As String = "SELECT SUM(ImporteAPU) FROM apuntes WHERE ConceptoAPU = ? AND EjercicioAPU = ? AND FechaAPU < DateSerial(?, ?, ?) + 1"
                Using cmdConcepto As New OleDbCommand(sqlConcepto, conexion1)
                    cmdConcepto.Parameters.AddWithValue("@Concepto", idElemento)
                    cmdConcepto.Parameters.AddWithValue("@Eje", anioEvaluar)
                    cmdConcepto.Parameters.AddWithValue("@EjeM", anioEvaluar)
                    cmdConcepto.Parameters.AddWithValue("@Mes", mes)
                    cmdConcepto.Parameters.AddWithValue("@Dia", dia)
                    Dim res = cmdConcepto.ExecuteScalar()
                    If res IsNot DBNull.Value AndAlso res IsNot Nothing Then totalConceptoAnio = Convert.ToDecimal(res)
                End Using

                dtFinal.Rows.Add(anioEvaluar, totalConceptoAnio)
            Next
        End If

        dtDatosInforme = dtFinal
        Return dtFinal
    End Function

    Private Sub FiltroEvolutivo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Limitamos el control de los Días (de 1 a 31)
        NudDia.Minimum = 1
        NudDia.Maximum = 31

        ' Limitamos el control de los Meses (de 1 a 12)
        NudMes.Minimum = 1
        NudMes.Maximum = 12

        If Me.TipoInforme = "CUENTAS" Then
            Label4.Text = resManager.GetString("TextLabel4Cuenta") & ":"
        Else
            Label4.Text = resManager.GetString("TextLabel4Concepto") & ":"
        End If


        ' 1. Encendemos tu escudo protector antes de rellenar los componentes
        cargandoFormulario = True

        ' 2. Inicializar los controles numéricos al día y mes actuales por comodidad
        NudDia.Value = Date.Now.Day
        NudMes.Value = Date.Now.Month

        ' 🌟 CABLE A: Cargamos el ComboBox de forma independiente ordenado de la A a la Z puro
        LlenarComboConceptosSueltosBD(Me.CmbConcepto)

        ' 🌟 CABLE B: Cargamos el combo de cuentas genérico de tu módulo
        LlenarComboCuentasGenerico(Me.CmbCuenta)

        ' 3. Control de visibilidad e interfaz según lo elegido en el menú principal
        If Me.TipoInforme = "CUENTAS" Then
            CmbCuenta.Visible = True
            CmbConcepto.Visible = False
            Me.Text = rmse.GetString("EvolucionSaldosPorCuenta")
        Else
            CmbCuenta.Visible = False
            CmbConcepto.Visible = True
            Me.Text = rmse.GetString("EvolucionSaldosPorConcepto")
        End If

        ' 4. Apagamos el escudo tras la inyección exitosa en la memoria RAM
        cargandoFormulario = False
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        Dim idSeleccionado As Integer = 0
        Dim nombreSeleccionado As String = ""

        ' 1. Determinar qué ID y Nombre capturar según el informe activo
        If Me.TipoInforme = "CUENTAS" Then
            If CmbCuenta.SelectedValue IsNot Nothing Then
                idSeleccionado = Convert.ToInt32(CmbCuenta.SelectedValue)
                nombreSeleccionado = CmbCuenta.Text
            End If
        Else
            If CmbConcepto.SelectedValue IsNot Nothing Then
                idSeleccionado = Convert.ToInt32(CmbConcepto.SelectedValue)
                nombreSeleccionado = CmbConcepto.Text
            End If
        End If

        ' Validación de seguridad por si no hay nada seleccionado
        If idSeleccionado <= 0 Then
            MsgBox(rmse.GetString("SeleccionarElemento"), MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        ' 2. Llamar a la función que creamos antes para rellenar el DataTable en memoria
        dtDatosInforme = ObtenerDatosEvolutivos(idSeleccionado, CInt(NudMes.Value), CInt(NudDia.Value))

        ' 🌟 INICIO DE TU MACRO DE REPORTE PERSONALIZADA
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString()

        ' Guardamos variables globales para usarlas en el motor de dibujo (PrintPage)
        vTxtNombreElementoReporte = nombreSeleccionado ' Para pintar el nombre del banco o concepto en el folio
        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        ' 3. Configuración de flujos según las preferencias de tus usuarios (My.Settings)
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

        Me.Close()
    End Sub

    Private Sub NudMes_ValueChanged(sender As Object, e As EventArgs) Handles NudMes.ValueChanged
        ' Averiguamos cuántos días tiene el mes seleccionado en el año actual de trabajo
        Dim diasDelMes As Integer = DateTime.DaysInMonth(vAñoEjercicio, CInt(NudMes.Value))

        ' Ajustamos el máximo dinámicamente (puede ser 28, 29, 30 o 31)
        NudDia.Maximum = diasDelMes
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' 1. Definimos los tipos de letras a utilizar en el reporte (Tu estructura original)
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)

        ' Definimos el formato alineado a la derecha para importes económicos
        Dim sfFar As New StringFormat With {.Alignment = StringAlignment.Far}

        ' Configuramos el título del informe en la plantilla
        vTituloInforme = If(Me.TipoInforme = "CUENTAS", rmse.GetString("EvolucionSaldosPorCuentaMay"), rmse.GetString("EvolucionSaldosPorConceptoMay"))
        frmImprimirForm.LblTitulo.Text = vTituloInforme

        ' Definimos el rango o fecha de corte para el LblEntreFechas
        frmImprimirForm.LblEntreFechas.Text = rmse.GetString("FechaCorte") & ": " & NudDia.Value.ToString("00") & " / " & NudMes.Value.ToString("00")

        ' 2. Imprimimos el encabezado gráfico usando tu plantilla frmImprimirForm
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)

        If frmImprimirForm.PictureBox1.Image IsNot Nothing Then
            Dim newImage As System.Drawing.Image = frmImprimirForm.PictureBox1.Image
            e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        End If

        e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Left + 200, frmImprimirForm.LblEntreFechas.Top)

        ' 3. Pintamos los subtítulos o nombres de las columnas en el folio
        ' Reutilizamos Punto1 para el Año y Punto2 para el Nombre del Elemento. El Importe va en Punto5 alineado a la derecha.
        e.Graphics.DrawString(rmse.GetString("Año") & " / " & resManager.GetString("Ejercicio"), FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)

        Dim encabezadoElemento As String = If(Me.TipoInforme = "CUENTAS", rmse.GetString("Cuenta"), rmse.GetString("Concepto"))
        e.Graphics.DrawString(encabezadoElemento & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)

        e.Graphics.DrawString(resManager.GetString("Importe") & " (" & vMoneda & "):", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30, sfFar)

        ' Imprimimos tu línea divisoria debajo de los encabezados
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        ' 4. Recorrido del DataTable en memoria (dtDatosInforme) en lugar del DataGridView
        Dim startY As Integer = frmImprimirForm.Punto1.Top

        ' Recorremos las filas usando la variable global PrintLine para soportar múltiples páginas de forma nativa
        Do While PrintLine < dtDatosInforme.Rows.Count
            ' Control de salto de página automático asistido por el alto de tu plantilla
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' Extracción limpia de datos de la fila de la memoria RAM
            Dim filaActual As DataRow = dtDatosInforme.Rows(PrintLine)

            Dim textoAnio As String = filaActual("Anio").ToString()
            Dim textoElemento As String = vTxtNombreElementoReporte ' Capturado en el botón Aceptar (ej: "BBVA")

            ' El importe viene de la columna 1 del SQL, lo formateamos limpiamente como moneda
            Dim valorImporte As Decimal = Convert.ToDecimal(filaActual(1))
            Dim textoImporte As String = valorImporte.ToString("N2") ' Formato numérico estándar con 2 decimales

            ' 🌟 EL CORTAFUEGOS TRADUCTOR MULTIIDIOMA PARA EL ELEMENTO (Concepto o Cuenta)
            If resManager IsNot Nothing Then
                Dim culturaActivaEnVivo As System.Globalization.CultureInfo = Threading.Thread.CurrentThread.CurrentUICulture

                ' Escaneamos de forma inversa por RAM para cazar la Key neutral (ej: "EFECTIVO", "BANCO", "LUZ")
                Dim claveElementoNeutral As String = ObtenerClaveNeutral(textoElemento, resManager)

                If Not String.IsNullOrEmpty(claveElementoNeutral) Then
                    ' Si encontramos la clave en el diccionario, inyectamos la traducción del idioma activo
                    Dim tradElemento As String = resManager.GetString(claveElementoNeutral, culturaActivaEnVivo)
                    If Not String.IsNullOrEmpty(tradElemento) Then textoElemento = tradElemento.Trim().ToUpper()
                Else
                    ' Plan B de respaldo: reemplazo de espacios por guiones bajos por si acaso
                    Dim tradDirecta As String = resManager.GetString(textoElemento.Replace(" ", "_"), culturaActivaEnVivo)
                    If Not String.IsNullOrEmpty(tradDirecta) Then textoElemento = tradDirecta.Trim().ToUpper()
                End If
            End If

            ' 5. Dibujo de la fila en el lienzo gráfico (Folio)
            e.Graphics.DrawString(textoAnio, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(textoElemento, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)

            ' Pintamos el importe alineado a la derecha usando el punto de destino final
            e.Graphics.DrawString(textoImporte, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Left, startY, sfFar)

            ' Avanzamos la posición vertical para la siguiente fila
            startY += 25
            PrintLine += 1
        Loop

        ' =========================================================================
        ' 🌟 BLOQUE DE PIE DE PÁGINA RECUPERADO (¡Tu consistencia intacta!)
        ' =========================================================================
        ' 5. LÍNEA DE PIE DE PÁGINA (Solo se dibuja si ya se han impreso todas las filas)
        If PrintLine >= dtDatosInforme.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
        End If

        ' 6. CONTADOR DE PÁGINAS DINÁMICO
        frmImprimirForm.LblNumeroPagina.Text = (CInt(frmImprimirForm.LblNumeroPagina.Text) + 1).ToString()

        e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
        ' =========================================================================

        ' Si terminó de dibujar todo el reporte, reiniciamos el contador de páginas para próximas ejecuciones
        If PrintLine >= dtDatosInforme.Rows.Count Then
            e.HasMorePages = False
            PrintLine = 0
        Else
            e.HasMorePages = True
        End If
    End Sub

    Private Sub NumericUpDown_Enter(sender As Object, e As EventArgs) Handles NudDia.Enter, NudMes.Enter
        Dim num As NumericUpDown = CType(sender, NumericUpDown)

        ' Ejecuta la selección justo un instante después de recibir el foco
        BeginInvoke(Sub()
                        num.Select(0, num.Text.Length)
                    End Sub)
    End Sub

End Class