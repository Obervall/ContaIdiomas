Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Reflection
Imports System.Resources
Imports System.Threading
Imports System.Windows.Forms
Imports MessageBox = ContaHogar.MsgBoxTraductorGlobal

Module Funciones

    Public Declare Function GetPrivateProfileStringKey Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    Public frmPrincipal As New Principal
    Public frmSeleccionEjercicio As New SeleccionEjercicio
    Public frmCuentasBancarias As New CuentasBancarias
    Public frmConceptosContables As New ConceptosContables
    Public frmApuntesContables As New ApuntesContables
    Public frmApuntesPeriodicos As New ApuntesPeriodicos
    Public frmIntroApuntes As New IntroApuntes
    Public frmIntroApuntesPeriodicos As New IntroApuntesPeriodicos
    Public frmEditarApuntes As New EditarApuntes
    Public frmEditarApuntesPeriodicos As New EditarApuntesPeriodicos
    Public frmTraspasoCuentas As New TraspasoCuentas
    Public frmIntroPresupuestos As New IntroPresupuestos
    Public frmPresupuestos As New Presupuestos
    Public frmAcercaDe As New AcercaDe
    Public frmPreferencias As New Preferencias
    Public frmNuevoConceptoContable As New NuevoConceptoContable
    Public frmNuevaCuentaBancaria As New NuevaCuentaBancaria
    Public frmNuevoTipoCuentaBancaria As New NuevoTipoCuentaBancaria
    Public frmTipoCuentaBancaria As New TipoCuentaBancaria
    Public frmEditarConceptoContable As New EditarConceptoContable
    Public frmEditarCuentaBancaria As New EditarCuentaBancaria
    Public frmEditarTipoCuentaBancaria As New EditarTipoCuentaBancaria
    Public frmBuscar As New Buscar
    Public frmFiltroF5 As New FiltroF5
    Public frmImprimirForm As New ImprimirForm
    Public frmSeleccionFechas As New SeleccionFechas
    Public frmTipoInformeApuntes As New TipoInformeApuntes
    Public frmTipoInformeApuntesPeriodicos As New TipoInformeApuntesPeriodicos
    Public frmTipoGrafico As New TipoGrafico
    Public frmTipoGraficoPeriodico As New TipoGraficoPeriodico
    Public frmGraficosConceptos As New GraficosConceptos
    Public frmGraficosCuentas As New GraficosCuentas
    Public frmGraficosFechas As New GraficosFechas
    Public frmGraficosMeses As New GraficosMeses
    Public frmGraficosPresupuestos As New GraficosPresupuestos
    Public frmSeleccionarDatosIngresos As New SeleccionDatosIngresos
    Public frmSeleccionarDatosGastos As New SeleccionDatosGastos
    Public frmGraficosSoloConceptos As New GraficosSoloConceptos
    Public frmActivarSoftware As New ActivarSoftware
    Public frmAportacionBizum As New AportacionBizum
    'Public frmAyudaApuntes As New AyudaApuntes

    Public backup As New SaveFileDialog
    Public restore As New OpenFileDialog

    Public conexion1 As New OleDbConnection()
    Public cmdMdb1cr As New OleDbCommand
    Public drMdb1 As OleDbDataReader

    Public vgrid, linSql, opcion, vTipoEstados, vNombreCuenta, vNombreConcepto, vFecha, vFechaMes As String
    Public vtipoSql, vAñadirSql, vtipoGrid, vMes, vEditar, vBuscar, vTituloInforme, vtipoSqlChk As String
    Public vValor, vIngresos, vGastos, vSaldo, vSaldoCuentas, vSaldoMes As Double
    Public i, vFila1, vFila2, vFila, vFilaActual, filaActual, vregData1, vAñoActual, vAñoEjercicio As Integer
    Public vCerrar, vGrafico, vGraficoSolo, vLetras, vNumeros, vNotas, vPathExportar, vConcepto As String
    Public vDescripcionAPU, vImporteAPU, vNotasAPU, vConceptoAPU As String
    Public vActualizar, vActivado, vAviso As Boolean

    Public vOrdenadoPorFechasAPU, vOrdenadoPorConceptosAPU, vOrdenadoPorImportesAPU As Integer
    Public vSoloIngresosAPU, vSoloGastosAPU, centroX, AnchoFrmPrincipal, posX, posY As Integer
    Public vOrdenadoPorFechasAPP, vOrdenadoPorConceptosAPP, vOrdenadoPorImportesAPP As Integer
    Public vSoloIngresosAPP, vSoloGastosAPP, vFecha1Enero, vFecha31Diciembre, vCantAños As Integer
    Public vRuta, vVersion, vNewVersion, vHayNuevaVersion, vNuevaVersion, vMoneda As String
    Public resManager As New ResourceManager("Contahogar.Recursos", Assembly.GetExecutingAssembly())
    ' Para copiar en en el Classe de cada Form la línea:
    'Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())
    Public vAñadir, vAñadir2, vImporteConcepto, vNewImporteConcepto, vExistenteImporteConcepto As String
    Public vDate1 As DateTime
    Public vDate2 As DateTime
    Public vDate3 As DateTime
    Public vFechaTemp As DateTime
    Public vFechaTemp2 As DateTime
    Public vfechaHoy As DateTime = DateTime.Today
    Public vTotalPresupuestoYTD As Double = 0
    Public vTotalRealYTD As Double = 0

    Public Structure ElementoCombo
        Public Property TextoMostrar As String  ' Lo que ve el usuario (ej: "Ausgaben")
        Public Property ValorInterno As String  ' Lo que va a la BD (ej: "GASTO")

        ' Esto es vital para que el ComboBox sepa qué texto pintar en pantalla
        Public Overrides Function ToString() As String
            Return TextoMostrar
        End Function
    End Structure

    Public Sub CambiarIdiomaGlobal(ByVal codIdioma As String)
        ' Configuramos el idioma (ej: "es", "ca", "en")
        Dim cultura As New CultureInfo(codIdioma)
        Thread.CurrentThread.CurrentCulture = cultura
        Thread.CurrentThread.CurrentUICulture = cultura

        ' Assignem els botons traduits a la nostra classe personalitzada
        ' (Revisa que "BotoAceptar", etc. coincideixin amb les Keys del teu fitxer .resx)
        MsgBoxTraductorGlobal.TextBotoOk = resManager.GetString("BotonAceptar")
        MsgBoxTraductorGlobal.TextBotoCancel = resManager.GetString("ToolTipCancelar")
        MsgBoxTraductorGlobal.TextBotoYes = resManager.GetString("BotonSi")
        MsgBoxTraductorGlobal.TextBotoNo = resManager.GetString("BotonNo")

        ' Refrescamos todos los formularios que estén abiertos en este momento
        For Each f As Form In Application.OpenForms
            ActualizarTextosFormulario(f)
        Next
    End Sub

    Public Sub ActualizarTextosFormulario(ByVal f As Form)
        ' 1. Crear el ComponentResourceManager
        Dim rmse As New System.ComponentModel.ComponentResourceManager(f.GetType())

        ' 2. Traducir el formulario y dejar que escale el tamaño libremente
        rmse.ApplyResources(f, "$this")

        ' 3. Traducir todos los controles de forma recursiva
        AplicarRecursosControles(f.Controls, rmse)

        ' 4. TRADUCCIÓ DINÀMICA DEL TÍTOL (Tu paso 4 intacto)
        If f.Name = "Principal" Then
            Dim txtTitol As String = resManager.GetString("TitolApp")
            Dim txtVersio As String = resManager.GetString("Versio")
            Dim txtExercici As String = resManager.GetString("Ejercicio")

            If txtTitol IsNot Nothing AndAlso txtVersio IsNot Nothing AndAlso txtExercici IsNot Nothing Then
                f.Text = String.Format("{0}  -  {1}:  {2}  -  {3}: {4}",
                                            txtTitol,
                                            txtVersio,
                                            My.Settings.Version,
                                            txtExercici,
                                            vAñoEjercicio.ToString())
            End If
        End If
    End Sub

    ' Método intermedio recursivo indispensable para limpiar tu antiguo bucle con "HasChildren"
    Private Sub AplicarRecursosControles(ByVal controles As Control.ControlCollection, ByVal rmse As System.ComponentModel.ComponentResourceManager)
        For Each ctrl As Control In controles
            ' Aplica automáticamente las propiedades (Text, etc.) buscando por el nombre del control
            rmse.ApplyResources(ctrl, ctrl.Name)

            ' Si el control tiene hijos (Panels, GroupBox, TabPages, etc.), se llama a sí mismo sin importar los niveles
            If ctrl.HasChildren Then
                AplicarRecursosControles(ctrl.Controls, rmse)
            End If
        Next
    End Sub

    Public Sub Conectarse(ByRef tipoDsn As String)
        Dim strSql As String
        strSql = tipoDsn
        vregData1 = 0
        If strSql = "AccessMdb" Then
            If conexion1.State = 0 Then
                conexion1.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0; Data Source=" & vRuta & "; Persist Security Info=False;"
                ':::Utilizamos el try para capturar posibles errores
                Try
                    ':::Abrimos la conexión
                    conexion1.Open()
                    If conexion1.State <> 0 Then
                        vregData1 = 1
                    End If
                    ':::Si se estableció conexión correctamente dirá "Conectado"
                    'MessageBox.Show("Se ha conectado a " & conexion1.ConnectionString)
                Catch ex As Exception
                    ':::Si no se conecta nos mostrara el posible fallo en la conexión
                    MsgBox(resManager.GetString("NoConectoPor") & ": " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Function IniciarSaldosIniciales(vAny As String) As Boolean
        Dim vAñoEjercicio As Integer = CInt(vAny)
        Dim vAñoAnterior As Integer = vAñoEjercicio - 1
        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Try
                conexion.Open()

                ' =========================================================================
                ' PASO 0: RECUPERAR EL ID NUMÉRICO DEL CONCEPTO "SALDO"
                ' =========================================================================
                Dim idConceptoSaldo As Integer = 1
                Using conexionId As New OleDbConnection(conexion1.ConnectionString)
                    Try
                        conexionId.Open()
                        Using cmdBuscarId As New OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexionId)
                            Dim resId = cmdBuscarId.ExecuteScalar()
                            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then
                                idConceptoSaldo = Convert.ToInt32(resId)
                            End If
                        End Using
                    Catch
                        idConceptoSaldo = 1 ' Valor por defecto de seguridad
                    End Try
                End Using

                ' =========================================================================
                ' PASO 1: BORRADO (Filtrando por el ID numérico que acabamos de rescatar)
                ' =========================================================================
                ' Modificamos el DELETE para que busque por el ID del concepto y el año CInt
                vtipoSql = "DELETE FROM apuntes WHERE apuntes.ConceptoAPU = ? And apuntes.EjercicioAPU = ?"
                Using conexionDel As New OleDbConnection(conexion1.ConnectionString)
                    Using cmdDelete As New OleDbCommand(vtipoSql, conexionDel)
                        ' Al ser Access, añadimos los parámetros en el orden exacto de los signos '?'
                        cmdDelete.Parameters.AddWithValue("?", idConceptoSaldo)
                        cmdDelete.Parameters.AddWithValue("?", CInt(vAñoEjercicio))
                        Try
                            conexionDel.Open()
                            Dim filasBorradas As Integer = cmdDelete.ExecuteNonQuery()
                            ' MsgBox("Se han limpiado correctamente los saldos anteriores: " & filasBorradas.ToString())
                        Catch ex As Exception
                            MsgBox("Error al ejecutar el borrado inicial con ID: " & ex.Message)
                            Return False
                        End Try
                    End Using
                End Using

                ' =========================================================================
                ' 🌟 RETORNO A LA LÓGICA DE ARRASTRE HISTÓRICO COMPLETO (Suma todo el pasado)
                ' =========================================================================
                ' Volvemos a tu consulta original: sumamos absolutamente todo lo anterior al año seleccionado
                Dim consulta As String =
                "SELECT A.EjercicioAPU, A.CuentaAPU, SUM(A.ImporteAPU) AS SumaAño " &
                "FROM (Ejercicios AS E INNER JOIN Apuntes AS A ON E.EjercicioEJE = A.EjercicioAPU) " &
                "WHERE E.EjercicioEJE < ? " &
                "GROUP BY A.EjercicioAPU, A.CuentaAPU " &
                "ORDER BY A.EjercicioAPU ASC"
                Dim dtMovimientos As New DataTable()
                Using conexion2 As New OleDbConnection(conexion1.ConnectionString)
                    Using comando As New OleDbCommand(consulta, conexion2)
                        ' Pasamos el año de la apertura (ej: 2026, buscará todo lo menor a 2026)
                        comando.Parameters.AddWithValue("@AñoSeleccionado", CInt(vAñoEjercicio))
                        Using adaptador As New OleDbDataAdapter(comando)
                            Try
                                conexion2.Open()
                                adaptador.Fill(dtMovimientos)
                            Catch ex As Exception
                                MessageBox.Show("Error al leer históricos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                Return False
                            End Try
                        End Using
                    End Using
                End Using
                ' Processing en memoria por ID de cuenta (Integer)
                Dim saldosAcumulados As New Dictionary(Of Integer, Decimal)()
                For Each fila As DataRow In dtMovimientos.Rows
                    Dim cuentaId As Integer = Convert.ToInt32(fila("CuentaAPU"))
                    Dim importeAño As Decimal = Convert.ToDecimal(fila("SumaAño"))

                    If saldosAcumulados.ContainsKey(cuentaId) Then
                        saldosAcumulados(cuentaId) += importeAño
                    Else
                        saldosAcumulados.Add(cuentaId, importeAño)
                    End If
                Next
                If saldosAcumulados.Count = 0 Then
                    vAviso = True
                    Return False
                Else
                    vAviso = False
                End If

                '' 🕵️ CHIVATO DE CONTROL: Ver los datos que se han calculado en la memoria RAM
                'Dim resumenChivato As String = $"AÑO CALCULADO: {vAñoAnterior}" & vbCrLf &
                '                       $"Total de cuentas encontradas: {saldosAcumulados.Count}" & vbCrLf & vbCrLf &
                '                       "Detalle de Saldos por ID de Cuenta:" & vbCrLf &
                '                       "------------------------------------" & vbCrLf

                'For Each par In saldosAcumulados
                '    resumenChivato &= $"ID Cuenta: {par.Key} -> Saldo Acumulado: {par.Value:N2} €" & vbCrLf
                'Next

                '' Muestra la ventana en pantalla antes de seguir
                'MsgBox(resumenChivato, MsgBoxStyle.Information, "Chivato de Control de Saldos")

                ' =========================================================================
                ' PASO 4: INSERCIÓN CON TRADUCCIÓN DINÁMICA DE LA DESCRIPCIÓN
                ' =========================================================================
                Dim fechaSaldoInicial As New Date(CInt(vAñoEjercicio), 1, 1)

                ' 🌟 TRUCO DE IDIOMA: Buscamos la traducción de "Saldo Inicial" en tu .resx
                Dim descripcionTraducida As String = "Saldo Inicial" ' Valor por defecto en castellano

                If resManager IsNot Nothing Then
                    ' Buscamos la Key (por ejemplo: "SaldoInicial") en tu ResX Manager
                    Dim trad As String = resManager.GetString("SaldoInicial")
                    If Not String.IsNullOrEmpty(trad) Then
                        descripcionTraducida = trad ' Si existe, usamos la traducción (Ej: "Balance Initial" o "Saldo Inicial")
                    End If
                End If

                Dim consultaInsert As String =
                "INSERT INTO Apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, CuentaAPU) " &
                "VALUES (?, ?, ?, ?, ?, ?)"

                ' ... (tu código de configuración de caché de Access se queda igual) ...

                Using comandoInsert As New OleDbCommand(consultaInsert, conexion)
                    comandoInsert.Parameters.Clear()
                    comandoInsert.Parameters.Add("@Fecha", OleDbType.Date)
                    comandoInsert.Parameters.Add("@Concepto", OleDbType.Integer)
                    comandoInsert.Parameters.Add("@Descripcion", OleDbType.VarWChar)
                    comandoInsert.Parameters.Add("@Importe", OleDbType.Currency)
                    comandoInsert.Parameters.Add("@Ejercicio", OleDbType.Integer)
                    comandoInsert.Parameters.Add("@Cuenta", OleDbType.Integer)

                    Try
                        Using transaccion As OleDbTransaction = conexion.BeginTransaction()
                            comandoInsert.Transaction = transaccion

                            For Each par In saldosAcumulados
                                Dim cuentaId As Integer = par.Key
                                Dim saldoFinalPasado As Decimal = par.Value

                                If saldoFinalPasado <> 0 Then
                                    comandoInsert.Parameters("@Fecha").Value = fechaSaldoInicial
                                    comandoInsert.Parameters("@Concepto").Value = idConceptoSaldo
                                    ' 🌟 INYECTAMOS LA DESCRIPCIÓN TRADUCIDA SEGÚN EL IDIOMA ACTUAL
                                    comandoInsert.Parameters("@Descripcion").Value = descripcionTraducida
                                    comandoInsert.Parameters("@Importe").Value = Math.Round(saldoFinalPasado, 2)
                                    comandoInsert.Parameters("@Ejercicio").Value = CInt(vAñoEjercicio)
                                    comandoInsert.Parameters("@Cuenta").Value = cuentaId

                                    comandoInsert.ExecuteNonQuery()
                                End If
                            Next

                            transaccion.Commit()
                        End Using
                        Return True
                    Catch ex As Exception
                        MessageBox.Show("Error al generar saldos iniciales: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End Try
                End Using
            Catch ex As Exception
                MessageBox.Show("Error al generar saldos iniciales: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Using
    End Function

    Public Sub LlenarGrid(ByRef tipoSql As String, tipoGrid As String, tipoopc As String)
        linSql = tipoSql.ToString
        vgrid = tipoGrid.ToString
        opcion = tipoopc
        If vgrid = "APUNTES_CONTABLES" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
                End If
                Dim Tabla As New DataTable
                adp.Fill(Tabla)
                frmApuntesContables.DgvApuntes.DataSource = Nothing
                frmApuntesContables.DgvApuntes.DataSource = Tabla
            End Using
            With frmApuntesContables.DgvApuntes
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue
                .ScrollBars = ScrollBars.Both
                .AllowUserToResizeColumns = True
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                'Liberamos las columnas del auto-ajuste estricto para permitir el Scroll manual posterior
                'For Each columna As DataGridViewColumn In .Columns
                '    columna.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                'Next

                ' arreglamos columnas
                '********************
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(4).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(5).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(3).DefaultCellStyle.Format = "N2"
                .Columns(4).DefaultCellStyle.Format = "N2"
                .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.Format = "dd/MM/yyyy"
                .Columns(0).Width = 100
                .Columns(0).HeaderText = resManager.GetString("Fecha") ' "Fecha"
                .Columns(1).Width = 200
                .Columns(1).HeaderText = resManager.GetString("Concepto") ' "Concepto"
                .Columns(2).Width = 200
                .Columns(2).HeaderText = resManager.GetString("Descripcion") ' "Descripción"
                .Columns(3).Width = 100
                .Columns(3).HeaderText = resManager.GetString("Importe") & " " & vMoneda
                .Columns(4).Width = 90
                .Columns(4).HeaderText = resManager.GetString("Saldo") & " " & vMoneda
                .Columns(5).Width = 140
                .Columns(5).HeaderText = resManager.GetString("Notas") ' "Notas"
                .Columns(6).Width = 140
                .Columns(6).HeaderText = resManager.GetString("Cuenta") ' "Cuenta"
                .Columns(7).Width = 0
                .Columns(7).HeaderText = resManager.GetString("Codigo") ' "Código"
                ' Ocultamos por completo las columnas técnicas que usa el traductor por debajo
                If .ColumnCount >= 11 Then
                    .Columns(7).Visible = False  ' [CodigoAPU] (Mejor que Width = 0 por seguridad)
                    .Columns(8).Visible = False  ' IdConceptoCON
                    .Columns(9).Visible = False ' DescripcionCON
                    .Columns(10).Visible = False ' IdCuentaCUE
                End If
            End With
            'Llama a la función
            DgvApuntesContables(3, 4)

            ' Para insertar alguna columna
            'Dim columna As New DataGridViewTextBoxColumn With {
            '.HeaderText = "Notas",
            '.Width = 350
            '}
            'frmApuntesContables.DgvApuntes.Columns.Insert(5, columna)

        ElseIf vgrid = "PRINT_APUNTES_CONTABLES" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
                End If
                Dim Tabla As New DataTable
                adp.Fill(Tabla)
                frmImprimirForm.DgvApuntes.DataSource = Nothing
                frmImprimirForm.DgvApuntes.DataSource = Tabla
            End Using

        ElseIf vgrid = "PRINT_INFORME_APUNTES" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
                Dim Tabla As New DataTable
                adp.Fill(Tabla)
                frmImprimirForm.DgvApuntes.DataSource = ""
                frmImprimirForm.DgvApuntes.DataSource = Tabla
            End Using
            vValor = 0
            frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ": 0,00 " & vMoneda
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(3).Value
                frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)
            Next

        ElseIf vgrid = "PRINT_TEMP_APUNTES" Or vgrid = "PRINT_TEMP_APUNTES_FECHAS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla

        ElseIf vgrid = "APUNTES_PERIODICOS" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
                End If
                Dim Tabla As New DataTable
                adp.Fill(Tabla)
                frmApuntesPeriodicos.DgvApuper.DataSource = ""
                frmApuntesPeriodicos.DgvApuper.DataSource = Tabla
            End Using
            With frmApuntesPeriodicos.DgvApuper
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue
                ' arreglamos columnas
                '********************
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(4).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(5).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(3).DefaultCellStyle.Format = "N2"
                .Columns(4).DefaultCellStyle.Format = "N2"
                .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.Format = "dd/MM/yyyy"
                .Columns(0).Width = 100
                .Columns(0).HeaderText = "Fecha"
                .Columns(1).Width = 200
                .Columns(1).HeaderText = "Concepto"
                .Columns(2).Width = 250
                .Columns(2).HeaderText = "Descripción"
                .Columns(3).Width = 100
                .Columns(3).HeaderText = "Importe(" & vMoneda & ")"
                .Columns(4).Width = 90
                .Columns(4).HeaderText = "Saldo(" & vMoneda & ")"
                .Columns(5).Width = 145
                .Columns(5).HeaderText = "Notas"
                .Columns(6).Width = 140
                .Columns(6).HeaderText = "Cuenta"
                .Columns(7).Width = 0
                .Columns(7).HeaderText = "Código"
            End With
            'Llama a la función
            DgvApuntesPeriodicos()

        ElseIf vgrid = "CONCEPTOS_CONTABLES" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmConceptosContables.DgvConceptos.DataSource = ""
            frmConceptosContables.DgvConceptos.DataSource = Tabla
            With frmConceptosContables.DgvConceptos
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue
                ' arreglamos columnas
                '********************
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(3).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(0).Width = 100
                .Columns(0).HeaderText = resManager.GetString("Tipo") ' My.Resources.Recursos.Tipo
                .Columns(1).Width = 200
                .Columns(1).HeaderText = resManager.GetString("Codigo") ' My.Resources.Recursos.Codigo
                .Columns(2).Width = 225
                .Columns(2).HeaderText = resManager.GetString("Descripcion") ' My.Resources.Recursos.Descripcion
                ' --- NUEVO: Hacemos que la columna 3 rellene el espacio restante del Grid ---
                .Columns(3).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(3).HeaderText = resManager.GetString("Notas")
                .Columns(4).Visible = False ' Oculta IdConceptoCON por completo
                Dim vNumRegistros As String = frmConceptosContables.DgvConceptos.Rows.Count.ToString
                frmConceptosContables.TxtNumRegistros.Text = vNumRegistros
                If frmConceptosContables.BtnFiltroTipoConcepto.Enabled = False Then
                    frmConceptosContables.LblNumRegistros.Text = resManager.GetString("Filtrado") ' My.Resources.Recursos.Filtrado
                Else
                    frmConceptosContables.LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
                End If
            End With

        ElseIf vgrid = "PRINT_CONCEPTOS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla

        ElseIf vgrid = "CUENTAS_BANCARIAS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)

            ' --- RECORRIDO DE FILAS CON TRADUCCIÓN DE TIPOS Y NOMBRES SISTEMA ---
            For Each filaData As DataRow In Tabla.Rows

                ' =========================================================================
                ' 3. CÁLCULO DEL SALDO POR ID (Lo que ya funcionaba impecable)
                ' =========================================================================
                Dim vIdCuenta As Integer = Convert.ToInt32(filaData("IdCuentaCUE"))

                cmdMdb1cr.CommandText = "SELECT apuntes.ImporteAPU FROM apuntes"
                cmdMdb1cr.CommandText += " WHERE apuntes.CuentaAPU = " & vIdCuenta
                cmdMdb1cr.CommandText += " And apuntes.EjercicioAPU = " & vAñoEjercicio

                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    vSaldoCuentas = 0
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            vSaldoCuentas += Convert.ToDecimal(drMdb1.GetValue(0))
                        End While
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                    MsgBox(resManager.GetString("ErrorAlEjecutar") & ": " & cmdMdb1cr.CommandText & vbCrLf & ex.Message)
                End Try

                filaData(3) = Math.Round(Convert.ToDecimal(vSaldoCuentas), 2)
            Next

            ' --- AHORA SÍ: Enlazamos la tabla ya calculada al Grid ---
            frmCuentasBancarias.DgvCuentas.DataSource = Nothing
            frmCuentasBancarias.DgvCuentas.DataSource = Tabla

            With frmCuentasBancarias.DgvCuentas
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue

                ' Configuración de alineaciones y colores
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).DefaultCellStyle.ForeColor = Color.DarkBlue

                ' TRUCO MAESTRO: Forzamos el formato N2 ahora que la columna contiene números puros
                .Columns(3).DefaultCellStyle.Format = "N2"

                ' Dimensiones y encabezados traducidos
                .Columns(0).Width = 150
                .Columns(0).HeaderText = resManager.GetString("Tipo")
                .Columns(1).Width = 200
                .Columns(1).HeaderText = resManager.GetString("Nombre")
                .Columns(2).Width = 200
                .Columns(2).HeaderText = resManager.GetString("Numero")
                .Columns(3).Width = 125
                .Columns(3).HeaderText = resManager.GetString("Saldo") & "(" & vMoneda & ")"
                .Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(4).HeaderText = resManager.GetString("Notas")
                .Columns(5).Visible = False
                Dim vNumRegistros As String = .Rows.Count.ToString
                frmCuentasBancarias.TxtNumRegistros.Text = vNumRegistros
                If frmCuentasBancarias.BtnFiltroTipoCuenta.Enabled = False Then
                    frmCuentasBancarias.LblNumRegistros.Text = resManager.GetString("Filtrado")
                Else
                    frmCuentasBancarias.LblNumRegistros.Text = resManager.GetString("SinFiltrar")
                End If
            End With
            DgvCuentasBancarias()

        ElseIf vgrid = "PRINT_CUENTAS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla
            frmImprimirForm.DgvApuntes.Columns(3).DefaultCellStyle.Format = "N2"
            vValor = 0
            frmImprimirForm.LblTotal.Text = resManager.GetString("TOTAL") & ": 0,00 " & vMoneda
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vNombreCuenta = fila.Cells(1).Value
                ' Buscar el Saldo de cada Cuenta Bancaria en Apuntes
                '***************************************************
                cmdMdb1cr.CommandText = "SELECT apuntes.ImporteAPU FROM apuntes"
                cmdMdb1cr.CommandText += " WHERE "
                cmdMdb1cr.CommandText += "apuntes.CuentaAPU = '" & vNombreCuenta & "' "
                cmdMdb1cr.CommandText += "And apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    vSaldoCuentas = 0
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            vSaldoCuentas += drMdb1.GetValue(0)
                        End While
                    Else
                        'MsgBox("No existen registros en " & tipoSql)
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorAlEjecutar") & ":  " & cmdMdb1cr.CommandText & " : " & ex.Message)
                End Try
                fila.Cells(3).Value = Math.Round(Convert.ToDecimal(vSaldoCuentas), 2)
                vValor += vSaldoCuentas
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)

        ElseIf vgrid = "PRINT_CUENTAS_PERIODICAS" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                If frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                    adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
                End If
                Dim Tabla As New DataTable
                adp.Fill(Tabla)
                frmImprimirForm.DgvApuntes.DataSource = ""
                frmImprimirForm.DgvApuntes.DataSource = Tabla
            End Using

        ElseIf vgrid = "PRESUPUESTOS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)

            ' Asignamos la tabla al Grid para que se generen las filas
            frmPresupuestos.DgvPresupuestos.DataSource = Tabla

            With frmPresupuestos.DgvPresupuestos
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue

                ' Configuramos las cabeceras fijas en castellano tal cual tu código base
                .Columns(0).Width = 175
                .Columns(0).HeaderText = resManager.GetString("Concepto")
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkBlue

                .Columns(1).Width = 100
                .Columns(1).HeaderText = frmPresupuestos.rmse.GetString("Mes")
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue

                .Columns(2).Width = 97
                .Columns(2).HeaderText = resManager.GetString("Realidad")
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.Format = "N2"

                .Columns(3).Width = 97
                .Columns(3).HeaderText = frmPresupuestos.rmse.GetString("Presupuesto")
                .Columns(3).DefaultCellStyle.Format = "N2"

                .Columns(4).Width = 0
                .Columns(4).HeaderText = resManager.GetString("Fecha")

                ' Contador de registros
                frmPresupuestos.TxtNumRegistros.Text = .Rows.Count.ToString()
                If frmPresupuestos.BtnFiltroConcepto.Enabled = False Then
                    frmPresupuestos.LblNumRegistros.Text = resManager.GetString("Filtrado")
                Else
                    frmPresupuestos.LblNumRegistros.Text = resManager.GetString("SinFiltrar")
                End If

                ' Averiguamos el tipo de concepto activo leyendo la primera fila (INGRESO o GASTO)
                Dim vTipoConceptoActual As String = "GASTO"
                If .Rows.Count > 0 AndAlso .Rows(0).Cells(0).Value IsNot Nothing Then
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & .Rows(0).Cells(0).Value.ToString().Replace("'", "''") & "'", con)
                            Try
                                con.Open()
                                Dim res As Object = cmd.ExecuteScalar()
                                If res IsNot Nothing Then vTipoConceptoActual = res.ToString().Trim().ToUpper()
                            Catch
                            End Try
                        End Using
                    End Using
                End If

                ' NUEVAS VARIABLES: Para acumular las sumas de las columnas
                Dim vSumaColumnaRealCompleta As Double = 0
                Dim vSumaColumnaPresuCompleta As Double = 0

                ' VARIABLES PARA EL CONTROL DE DESVIACIÓN CONTROLADA (YTD)
                vTotalPresupuestoYTD = 0
                vTotalRealYTD = 0
                Dim mesActualCalendario As Integer = DateTime.Now.Month
                Dim añoActualCalendario As Integer = DateTime.Now.Year

                ' BUCLE PRINCIPAL MODIFICADO FILA A FILA
                For Each fila As DataGridViewRow In .Rows
                    If fila.IsNewRow Then Continue For

                    Dim vFecha As Date
                    Dim vMes As Integer = 1 ' Valor por defecto por si falla

                    If fila.Cells(4).Value IsNot Nothing AndAlso Date.TryParse(fila.Cells(4).Value.ToString(), vFecha) Then
                        vMes = vFecha.Month
                    End If

                    ' Ponemos el nombre del mes en la columna 1
                    fila.Cells(1).Value = MonthName(vMes, False)

                    Dim vNombreConcepto As String = fila.Cells(0).Value.ToString()

                    ' 🔥 CONSULTA INTERNA POR FILA: Averiguamos el tipo de este concepto específico (INGRESO o GASTO)
                    Dim vTipoConceptoFila As String = "GASTO"
                    Using con As New OleDbConnection(conexion1.ConnectionString)
                        Using cmd As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE CodigoCON = '" & vNombreConcepto.Replace("'", "''") & "'", con)
                            Try
                                con.Open()
                                Dim res As Object = cmd.ExecuteScalar()
                                If res IsNot Nothing Then vTipoConceptoFila = res.ToString().Trim().ToUpper()
                            Catch
                            End Try
                        End Using
                    End Using

                    ' Buscamos el saldo real de este concepto en este mes
                    Dim cmdMySql1cr As New OleDbCommand()
                    cmdMySql1cr.Connection = conexion1
                    cmdMySql1cr.CommandText = "SELECT FechaAPU, ConceptoAPU, ImporteAPU FROM apuntes"
                    cmdMySql1cr.CommandText += " WHERE EjercicioAPU = " & vAñoEjercicio.ToString
                    cmdMySql1cr.CommandText += " And ConceptoAPU = '" & vNombreConcepto.Replace("'", "''") & "' "

                    Dim vSaldoMes As Double = 0
                    Try
                        If conexion1.State <> ConnectionState.Open Then conexion1.Open()
                        Using drMySql1 As OleDbDataReader = cmdMySql1cr.ExecuteReader()
                            If drMySql1.HasRows Then
                                While drMySql1.Read()
                                    Dim vFechaMes As Date
                                    If Date.TryParse(drMySql1.GetValue(0).ToString(), vFechaMes) Then
                                        If vFechaMes.Month = vMes Then
                                            vSaldoMes += Convert.ToDouble(drMySql1.GetValue(2))
                                        End If
                                    End If
                                End While
                            End If
                        End Using
                    Catch ex As Exception
                        MsgBox(resManager.GetString("ErrorEjecutarSaldo") & ex.Message)
                    End Try

                    ' Asignamos el valor real final
                    fila.Cells(2).Value = -vSaldoMes
                    vSumaColumnaRealCompleta += -vSaldoMes

                    ' Conversión limpia y segura del valor del presupuesto
                    Dim importePresuFila As Double = 0
                    If fila.Cells(3).Value IsNot Nothing Then
                        Double.TryParse(fila.Cells(3).Value.ToString(), importePresuFila)

                        ' 🔥 SI ES INGRESO, lo forzamos a NEGATIVO tanto en la celda como en la variable de cálculo
                        If vTipoConceptoFila = "INGRESO" Then
                            importePresuFila = -Math.Abs(importePresuFila)
                            fila.Cells(3).Value = importePresuFila
                        End If

                        vSumaColumnaPresuCompleta += importePresuFila
                    End If

                    ' Acumulados controlados para el YTD financiero
                    If CInt(vAñoEjercicio) < añoActualCalendario Then
                        vTotalPresupuestoYTD += importePresuFila
                        vTotalRealYTD += (-vSaldoMes)
                    ElseIf CInt(vAñoEjercicio) = añoActualCalendario Then
                        Dim vMesInt As Integer = vMes
                        If vMesInt < mesActualCalendario Then
                            vTotalPresupuestoYTD += importePresuFila
                            vTotalRealYTD += (-vSaldoMes)
                        End If
                    End If
                Next
                ' Sincronizamos los totales de las etiquetas mediante tu resta limpia estándar
                Dim vDiferenciaDesviacion As Double = vTotalPresupuestoYTD - vTotalRealYTD

                ' Volcamos el resultado exacto en la caja de texto
                frmPresupuestos.LblMontoDesviacion.Text = vDiferenciaDesviacion.ToString("N2")

                ' Cambiamos los colores de la etiqueta según si el resultado es positivo (ganancia/ahorro) o negativo
                If vDiferenciaDesviacion >= 0 Then
                    frmPresupuestos.LblObjetivo.ForeColor = Color.DarkGreen
                    frmPresupuestos.LblObjetivo.Text = frmPresupuestos.rmse.GetString("LblObjetivo.Text")
                    If String.IsNullOrEmpty(frmPresupuestos.LblObjetivo.Text) Then frmPresupuestos.LblObjetivo.Text = "Objectiu Assolit!"
                    frmPresupuestos.LblMontoDesviacion.ForeColor = Color.DarkBlue
                Else
                    frmPresupuestos.LblObjetivo.ForeColor = Color.DarkRed
                    frmPresupuestos.LblObjetivo.Text = frmPresupuestos.rmse.GetString("NoLogrado")
                    If String.IsNullOrEmpty(frmPresupuestos.LblObjetivo.Text) Then frmPresupuestos.LblObjetivo.Text = "Objectiu No Assolit"
                    frmPresupuestos.LblMontoDesviacion.ForeColor = Color.Red
                End If

                ActualizarEtiquetaDesviacion()

                ' INSERCIÓN DE LA FILA DE TOTALES EN LA REJILLA
                Try
                    Dim filaTotales As DataRow = Tabla.NewRow()
                    filaTotales(0) = resManager.GetString("TOTAL")
                    filaTotales(1) = ""
                    filaTotales(2) = vSumaColumnaRealCompleta
                    filaTotales(3) = vSumaColumnaPresuCompleta
                    filaTotales(4) = DBNull.Value

                    Tabla.Rows.Add(filaTotales)
                    Tabla.AcceptChanges()
                Catch ex As Exception
                End Try
            End With

        ElseIf vgrid = "TIPO_CUENTAS_BANCARIAS" Then    'Tipo Cuentas Bancarias
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmTipoCuentaBancaria.DgvTipoCuentasBancarias.DataSource = Tabla
            With frmTipoCuentaBancaria.DgvTipoCuentasBancarias
                .DefaultCellStyle.Font = New Font("Tahoma", 10)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.Beige
                .DefaultCellStyle.SelectionForeColor = Color.Yellow
                .DefaultCellStyle.SelectionBackColor = Color.Black
                'arreglamos columnas
                '*******************
                .Columns(0).HeaderText = resManager.GetString("Codigo")
                .Columns(0).Width = 230
                ' --- NUEVO: Hacemos que la columna 4 rellene el espacio restante del Grid ---
                .Columns(1).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(1).HeaderText = resManager.GetString("Descripcion")
                .Columns(2).Visible = False
                Dim vNumRegistros As String = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows.Count.ToString
                frmTipoCuentaBancaria.TxtNumRegistros.Text = vNumRegistros
            End With

        ElseIf vgrid = "PRINT_TIPO_CUENTAS" Then    'Tipo Cuentas Bancarias
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = Tabla

        ElseIf vgrid = "NOMBRESEXISTENTES" Then    'Conceptos Contables
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmNuevoConceptoContable.DgvExistente.DataSource = Tabla
            With frmNuevoConceptoContable.DgvExistente
                .DefaultCellStyle.Font = New Font("Tahoma", 10)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.Beige
                .DefaultCellStyle.SelectionForeColor = Color.Yellow
                .DefaultCellStyle.SelectionBackColor = Color.Black
                'arreglamos columnas
                '*******************
                .Columns(0).HeaderText = resManager.GetString("Nombre") ' My.Resources.Recursos.NombresExistentes
                .Columns(0).Width = 230
            End With

        ElseIf vgrid = "NOMBRESEXISTENTES2" Then  'Cuentas Bancarias
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmNuevaCuentaBancaria.DgvExistente.DataSource = Tabla
            With frmNuevaCuentaBancaria.DgvExistente
                .DefaultCellStyle.Font = New Font("Tahoma", 10)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.Beige
                .DefaultCellStyle.SelectionForeColor = Color.Yellow
                .DefaultCellStyle.SelectionBackColor = Color.Black
                'arreglamos columnas
                '*******************
                .Columns(0).HeaderText = resManager.GetString("Nombre") ' My.Resources.Recursos.NombresExistentes
                .Columns(0).Width = 230
            End With

        ElseIf vgrid = "NOMBRESEXISTENTES3" Then  'Cuentas Bancarias
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmNuevoTipoCuentaBancaria.DgvExistente.DataSource = Tabla
            With frmNuevoTipoCuentaBancaria.DgvExistente
                .DefaultCellStyle.Font = New Font("Tahoma", 10)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.Beige
                .DefaultCellStyle.SelectionForeColor = Color.Yellow
                .DefaultCellStyle.SelectionBackColor = Color.Black
                'arreglamos columnas
                '*******************
                .Columns(0).HeaderText = resManager.GetString("Nombre") ' My.Resources.Recursos.NombresExistentes
                .Columns(0).Width = 230
            End With
        End If
    End Sub

    Public Sub ActualizarEtiquetaDesviacion()
        Dim añoActualCalendario As Integer = DateTime.Now.Year

        ' Comprobamos si el ejercicio consultado es el año en curso
        If CInt(vAñoEjercicio) = añoActualCalendario Then

            ' Si filtramos por concepto, ocultamos los campos de desviación
            If frmPresupuestos.BtnFiltroConcepto.Enabled = True Then ' Sin Filtrar Concepto
                frmPresupuestos.LblDesviacion.Visible = False
                frmPresupuestos.LblMontoDesviacion.Visible = False
            Else
                frmPresupuestos.LblDesviacion.Visible = True
                frmPresupuestos.LblMontoDesviacion.Visible = True

                ' Obtenemos la fecha del mes anterior restando 1 mes a la fecha de hoy
                Dim fechaMesAnterior As Date = DateTime.Now.AddMonths(-1)

                ' Obtenemos el nombre de ese mes en el idioma del sistema (ej: "mayo" si estamos en junio)
                Dim nombreMesAnterior As String = StrConv(fechaMesAnterior.ToString("MMMM"), VbStrConv.ProperCase)

                ' "Desviación Parcial Hasta: Mayo" (Traído desde tus recursos)
                frmPresupuestos.LblDesviacion.Text = frmPresupuestos.rmse.GetString("DesviacionParcial") & " " & nombreMesAnterior & " ="
            End If

        ElseIf CInt(vAñoEjercicio) < añoActualCalendario Then
            ' Si es un año pasado, el ejercicio ya está cerrado: Desviación Anual
            frmPresupuestos.LblDesviacion.Visible = True
            frmPresupuestos.LblMontoDesviacion.Visible = True
            Dim textoAnual As String = frmPresupuestos.rmse.GetString("LblDesviacion.Text")
            If String.IsNullOrEmpty(textoAnual) Then textoAnual = "Desviació Anual"
            frmPresupuestos.LblDesviacion.Text = textoAnual & " " & vAñoEjercicio & "= "
        Else
            ' Si es un año futuro, podrías querer ocultarlo o gestionarlo
            frmPresupuestos.LblDesviacion.Visible = False
            frmPresupuestos.LblMontoDesviacion.Visible = False
        End If
    End Sub

    Public Function DgvCuentasBancarias()
        frmCuentasBancarias.TxtIngresos.Text = ""
        frmCuentasBancarias.TxtGastos.Text = ""
        frmCuentasBancarias.TxtSaldo.Text = ""
        Dim vNumRegistros As String = frmCuentasBancarias.DgvCuentas.Rows.Count.ToString
        frmCuentasBancarias.TxtNumRegistros.Text = vNumRegistros
        If frmCuentasBancarias.BtnFiltroTipoCuenta.Enabled = False Then
            frmCuentasBancarias.LblNumRegistros.Text = resManager.GetString("Filtrado")
        Else
            frmCuentasBancarias.LblNumRegistros.Text = resManager.GetString("SinFiltrar")
        End If
        vIngresos = 0
        vGastos = 0
        vValor = 0
        For Each fila As DataGridViewRow In frmCuentasBancarias.DgvCuentas.Rows
            If fila.Cells(3).Value >= 0 Then
                vIngresos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.DarkBlue
                frmCuentasBancarias.TxtIngresos.Text = Format(Math.Abs(vIngresos).ToString("N2"))
            Else
                vGastos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.IndianRed
                frmCuentasBancarias.TxtGastos.Text = Format(Math.Abs(vGastos).ToString("N2"))
            End If
        Next
        vSaldo = vIngresos + vGastos
        frmCuentasBancarias.TxtSaldo.Text = Format(Math.Abs(vSaldo).ToString("N2"))
        Return vValor
    End Function

    Public Function DgvApuntesContables(vFila1, vFila2)
        ' En esta función se calcula el Saldo de cada Apunte y el Saldo Total, además de los Totales de Ingresos y Gastos.
        frmApuntesContables.TxtIngresos.Text = ""
        frmApuntesContables.TxtGastos.Text = ""
        frmApuntesContables.TxtSaldo.Text = ""
        Dim vNumRegistros As String = frmApuntesContables.DgvApuntes.Rows.Count.ToString
        frmApuntesContables.TxtNumRegistros.Text = vNumRegistros
        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Or frmApuntesContables.BtnFiltroConcepto.Enabled = False Or frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            frmApuntesContables.LblNumRegistros.Text = resManager.GetString("Filtrado")
        Else
            frmApuntesContables.LblNumRegistros.Text = resManager.GetString("SinFiltrar")
        End If
        vIngresos = 0
        vGastos = 0
        vValor = 0
        For Each fila As DataGridViewRow In frmApuntesContables.DgvApuntes.Rows
            vSaldo = fila.Cells(vFila1).Value + vValor
            fila.Cells(vFila2).Value = vSaldo
            vValor = fila.Cells(4).Value
            If fila.Cells(vFila1).Value >= 0 Then
                vIngresos += fila.Cells(vFila1).Value
                fila.Cells(vFila1).Style.ForeColor = Color.DarkBlue
                frmApuntesContables.TxtIngresos.Text = Format(Math.Abs(vIngresos).ToString("N2"))
            Else
                vGastos += fila.Cells(vFila1).Value
                fila.Cells(vFila1).Style.ForeColor = Color.IndianRed
                frmApuntesContables.TxtGastos.Text = Format(Math.Abs(vGastos).ToString("N2"))
            End If
            If fila.Cells(vFila2).Value >= 0 Then
                fila.Cells(vFila2).Style.ForeColor = Color.DarkBlue
            Else
                fila.Cells(vFila2).Style.ForeColor = Color.IndianRed
            End If
        Next
        frmApuntesContables.TxtSaldo.Text = Format(Math.Abs(vValor).ToString("N2"))
        Return vValor
    End Function

    Public Function DgvApuntesPeriodicos()
        ' En esta función se calcula el Saldo de cada Apunte y el Saldo Total del Periodo, además de los Totales de Ingresos y Gastos.
        frmApuntesPeriodicos.TxtIngresos.Text = ""
        frmApuntesPeriodicos.TxtGastos.Text = ""
        frmApuntesPeriodicos.TxtSaldo.Text = ""
        Dim vNumRegistros As String = frmApuntesPeriodicos.DgvApuper.Rows.Count.ToString
        frmApuntesPeriodicos.TxtNumRegistros.Text = vNumRegistros
        If frmApuntesPeriodicos.BtnFiltroCuenta.Enabled = False Or frmApuntesPeriodicos.BtnFiltroConcepto.Enabled = False Or frmApuntesPeriodicos.BtnFiltroFecha.Enabled = False Then
            frmApuntesPeriodicos.LblNumRegistros.Text = resManager.GetString("Filtrado")
        Else
            frmApuntesPeriodicos.LblNumRegistros.Text = resManager.GetString("SinFiltrar")
        End If
        vIngresos = 0
        vGastos = 0
        vValor = 0
        For Each fila As DataGridViewRow In frmApuntesPeriodicos.DgvApuper.Rows
            vSaldo = fila.Cells(3).Value + vValor
            fila.Cells(4).Value = vSaldo
            vValor = fila.Cells(4).Value
            If fila.Cells(3).Value >= 0 Then
                vIngresos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.DarkBlue
                frmApuntesPeriodicos.TxtIngresos.Text = Format(Math.Abs(vIngresos).ToString("N2"))
            Else
                vGastos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.IndianRed
                frmApuntesPeriodicos.TxtGastos.Text = Format(Math.Abs(vGastos).ToString("N2"))
            End If
            If fila.Cells(4).Value >= 0 Then
                fila.Cells(4).Style.ForeColor = Color.DarkBlue
            Else
                fila.Cells(4).Style.ForeColor = Color.IndianRed
            End If
        Next
        frmApuntesPeriodicos.TxtSaldo.Text = Format(Math.Abs(vValor).ToString("N2"))
        Return vValor
    End Function

    'Funcion para que solo permite el ingreso de caracteres tipo numerico y punto
    '****************************************************************************
    Public Sub SoloNumerosConPunto(ByRef e As System.Windows.Forms.KeyPressEventArgs)
        If Char.IsDigit(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf e.KeyChar = "." Then
            e.Handled = False
        Else
            e.Handled = True
            MsgBox(resManager.GetString("SoloAdmitePunto"),
            MsgBoxStyle.Exclamation, resManager.GetString("SeparadorDecimal"))
        End If
    End Sub
    Public Function ApostrofePorAcentoAgudo(ByVal sNombreCampo As String) As String
        Dim newNombreCampo As String = ""
        Try
            Dim a As Integer = InStr(1, sNombreCampo, "'", vbBinaryCompare)
            If a <> 0 Then
                newNombreCampo = Replace(sNombreCampo, "'", Convert.ToChar(180))
            Else
                newNombreCampo = sNombreCampo
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & " " & Err.Number & NL & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Informacion"))
        End Try
        Return newNombreCampo
    End Function

    Public Function BuscarActualizacion()
        Dim conectado As New Devices.Computer
        If vActualizar = True Then
            If conectado.Network.IsAvailable = True Then
                'MsgBox("Estas conectado a una red")
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                Try
                    Dim MyUrl As String = "https://filedn.eu/ljfTvwyEW2tVj4PWYI9927f/ContaHogar/Hogar.txt"
                    Dim MyHttpWebRequest As HttpWebRequest = CType(WebRequest.Create(MyUrl), HttpWebRequest)
                    MyHttpWebRequest.Credentials = CredentialCache.DefaultCredentials
                    Dim MyHttpWebResponse As HttpWebResponse = CType(MyHttpWebRequest.GetResponse(), HttpWebResponse)
                    Dim MyStream As Stream = MyHttpWebResponse.GetResponseStream
                    Dim MyStreamReader As New StreamReader(MyStream)
                    Dim MyHtml As String = MyStreamReader.ReadToEnd
                    Dim MyHtmlEnLineas() As String = MyHtml.Split(vbNewLine)
                    Dim vNewVersion As String = MyHtmlEnLineas(3)
                    vNewVersion = Mid(vNewVersion, 10)
                    vNewVersion = Trim(vNewVersion)

                    If My.Settings.Version < vNewVersion Then
                        MsgBox(resManager.GetString("VersionInstalada") & " " & My.Settings.Version & vbNewLine & resManager.GetString("VersionDisponible") & " " & vNewVersion, MsgBoxStyle.Information, resManager.GetString("Informacion"))
                        vNuevaVersion = vNewVersion
                        vHayNuevaVersion = "SI"
                        Dim respuesta As MsgBoxResult = MsgBox(resManager.GetString("ActualizarVersion") & ": " & vNewVersion & " ?", vbQuestion + vbYesNo + vbDefaultButton1, resManager.GetString("Versio") & " ContaHogar 3.0")
                        If respuesta = vbYes Then
                            Dim respuesta2 As MsgBoxResult = MsgBox(resManager.GetString("GuardarCopiaSeguridad"), vbQuestion + vbYesNo + vbDefaultButton1, resManager.GetString("Versio") & " ContaHogar 3.0")
                            If respuesta2 = vbYes Then
                                ' Si no existe la carpeta de BackUp la creamos.
                                Dim path As String = "C:\ContaHogar3.0\Backup"
                                If Directory.Exists(path) Then
                                    'MsgBox("Ya existe la Ruta C:\ConatHogar\Backup.")
                                Else
                                    Directory.CreateDirectory(path)
                                    'MsgBox("Ruta C:\ContaHogar3.0\Backup, Creada.")
                                End If
                                Dim NombreBaseDatos As String = $"ContaHogar3.0[{Now:ddMMyyyy}][{Now:HHmmss}].mdb"
                                'Dim NombreBaseDatos As String = "ContaHogar3.0" & "[" & Format(Now.ToString("ddMMyyyy")) & "]" & "[" & Format(Now.ToString("HHmmss")) & "]" & ".mdb"
                                Dim DataBaseFile As String = vRuta
                                Dim FileDestino As String = "C:\ContaHogar3.0\Backup\" & NombreBaseDatos
                                backup.InitialDirectory = "C:\ContaHogar3.0\Backup\"
                                backup.Title = "Backup Base de Datos Access"
                                backup.CheckFileExists = False
                                backup.CheckPathExists = False
                                backup.DefaultExt = "mdb"
                                backup.FileName = NombreBaseDatos
                                backup.Filter = "Access (ContaHogar*.mdb)|ContaHogar*.mdb|All files (*.*)|*.*"
                                backup.RestoreDirectory = True
                                If backup.ShowDialog = Windows.Forms.DialogResult.OK Then
                                    Try
                                        FileCopy(DataBaseFile, FileDestino)
                                        MessageBox.Show(resManager.GetString("BackupSatisfactorioBuscarActualizacio"), resManager.GetString("Backup"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    Catch ex As Exception
                                        MsgBox(resManager.GetString("ErrorBackup") & ": " & ex.ToString)
                                    End Try
                                End If
                            End If
                            MessageBox.Show(resManager.GetString("AhoraDescargaActualizacion"), resManager.GetString("ActualizarSoftware"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Try
                                Dim descargar As New Devices.Computer
                                With descargar
                                    .Network.DownloadFile("https://filedn.eu/ljfTvwyEW2tVj4PWYI9927f/ContaHogar/Actualizar/" & vNewVersion & "/InstaladorContaHogar3.0.msi", "C:\ContaHogar3.0\InstaladorContaHogar3.0.msi", "", "", True, 100, True, 3)
                                End With
                                Dim Proceso As New Process()
                                MsgBox(resManager.GetString("DescargadaActualizacion") & ":  " & vNewVersion & ", " & resManager.GetString("ProcederaInstalar"), MsgBoxStyle.Information, resManager.GetString("ActualizarSoftware"))
                                Proceso.StartInfo.FileName = "C:\ContaHogar3.0\InstaladorContaHogar3.0.msi"
                                Proceso.StartInfo.Arguments = ""
                                Proceso.Start()
                            Catch ex As Exception
                                MsgBox(resManager.GetString("ErrorDescargarActualizacion") & ":  " & vNewVersion & ", " & resManager.GetString("ErrorDescargarActualizacion2"))
                                MsgBox(ex.ToString)
                            End Try
                            Application.Exit()
                        End If
                    Else
                        vHayNuevaVersion = "NO"
                        'MsgBox(My.Settings.Version & " = " & vNewVersion)
                    End If
                Catch ex As Exception
                    MsgBox(resManager.GetString("ErrorDescargarActualizacion2"))
                    MsgBox(ex.ToString)
                End Try
            Else
                MsgBox(resManager.GetString("NoConectadoRed"), MsgBoxStyle.Information, resManager.GetString("Informacion"))
            End If
            Return vHayNuevaVersion
        Else
            Return "NO"
        End If
    End Function

    ' Función recursiva para recorrer todos los controles (paneles, groupbox, etc.)
    Public Function AplicarRecursosAControles(parent As Control, res As ComponentResourceManager)
        For Each c As Control In parent.Controls
            res.ApplyResources(c, c.Name)
            If c.HasChildren Then
                AplicarRecursosAControles(c, res)
            End If
        Next
        Return True
    End Function

    'En el MsgBox han observado & NL &, bien, en un módulo con funciones públicas para toda la solución va esto
    Friend NL As String = Environment.NewLine '(Me hace un salto de línea, es muy práctico).

    ''' <summary>
    ''' Traduce las columnas de cualquier Grid usando el traductor específico de cada formulario
    ''' </summary>
    ''' <param name="grid">El DataGridView a procesar</param>
    Public Sub TraducirColumnasGridCuentas(ByVal grid As DataGridView)
        Try
            If grid IsNot Nothing AndAlso grid.Rows.Count > 0 Then

                For Each fila As DataGridViewRow In grid.Rows
                    If Not fila.IsNewRow Then

                        ' --- COLUMNA (0): TipoCUE (Mixto - CORREGIDO VISUAL) ---
                        If grid.Columns.Count > 0 AndAlso fila.Cells(0).Value IsNot Nothing Then
                            Dim valorTipo As String = fila.Cells(0).Value.ToString().Trim()

                            ' Normalizamos a guion bajo para buscar de forma segura en las Keys del .resx
                            Dim llaveBase As String = valorTipo.Replace(" ", "_")
                            Dim tradTipo As String = resManager.GetString(llaveBase)

                            If Not String.IsNullOrEmpty(tradTipo) Then
                                fila.Cells(0).Value = tradTipo
                            Else
                                ' ¡EL PARCHE VISUAL SEGURO!: Si no hay traducción o estamos en español,
                                ' le quitamos los guiones bajos para que el usuario lo vea impecable
                                fila.Cells(0).Value = valorTipo.Replace("_", " ")
                            End If
                        End If

                        ' --- COLUMNA (1): NombreCUE (Mayúsculas - SE MANTIENE INTACTO) ---
                        If grid.Columns.Count > 1 AndAlso fila.Cells(1).Value IsNot Nothing Then
                            Dim valorNombre As String = fila.Cells(1).Value.ToString().Trim().ToUpper()

                            ' Normalizamos a guion bajo por si acaso tuviera espacios
                            Dim llaveNombre As String = valorNombre.Replace(" ", "_")
                            Dim tradNombre As String = resManager.GetString(llaveNombre)

                            If Not String.IsNullOrEmpty(tradNombre) Then
                                fila.Cells(1).Value = tradNombre
                            Else
                                ' Respaldo si está en español o es propio: quitamos guiones
                                fila.Cells(1).Value = valorNombre.Replace("_", " ")
                            End If
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorAlEjecutar") & ex.Message, MsgBoxStyle.Exclamation, resManager.GetString("Error"))
        End Try
    End Sub

    ''' <summary>
    ''' Rellena de forma híbrida y multidioma cualquier ComboBox con los tipos de cuenta desde Access
    ''' </summary>
    ''' <param name="combo">El control ComboBox que se quiere rellenar</param>
    Public Sub CargarComboTipoCuentaGlobal(ByVal combo As ComboBox)
        Dim textoTraducido As String = ""
        cmdMdb1cr.CommandText = "SELECT tipocuentas.CodigoTIP FROM tipocuentas ORDER BY tipocuentas.CodigoTIP ASC"
        Try
            Dim indiceSeleccionado As Integer = combo.SelectedIndex
            Dim historialSeguimiento As String = "--- HISTORIAL DE TRADUCCIONES ---" & vbNewLine
            combo.Items.Clear()
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    Dim valorBD As String = drMdb1.GetValue(0).ToString().Trim()

                    textoTraducido = resManager.GetString(valorBD)
                    If String.IsNullOrEmpty(textoTraducido) Then
                        textoTraducido = valorBD
                    End If

                    ' ===================================================================
                    ' EL ÚNICO CAMBIO: En vez de combo.Items.Add(textoTraducido)
                    ' Guardamos el objeto híbrido con su valor original de Access
                    ' ===================================================================
                    Dim elemento As New ElementoCombo With {
                    .TextoMostrar = textoTraducido,
                    .ValorInterno = valorBD
                }
                    combo.Items.Add(elemento)
                    ' ===================================================================

                    historialSeguimiento &= $"BD:  {valorBD} -> Trad: {textoTraducido}" & vbNewLine
                End While

                'MsgBox(historialSeguimiento, MsgBoxStyle.Information, "Resumen de Carga")

                If indiceSeleccionado >= 0 AndAlso indiceSeleccionado < combo.Items.Count Then
                    combo.SelectedIndex = indiceSeleccionado
                ElseIf combo.Items.Count > 0 Then
                    combo.SelectedIndex = 0
                End If
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorAlEjecutar") & ex.Message)
        Finally
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
                drMdb1.Close()
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Traduce las celdas dinámicas de Tipo y Descripción en el Grid de Cuentas Bancarias
    ''' </summary>
    ''' <param name="grid">El DataGridView a procesar</param>
    Public Sub TraducirContenidoGridTiposCuenta(ByVal grid As DataGridView)
        Try
            ' Validamos que el grid tenga filas y al menos las 2 columnas necesarias (Tipo y Descripción)
            If grid IsNot Nothing AndAlso grid.Rows.Count > 0 AndAlso grid.Columns.Count > 1 Then

                For Each fila As DataGridViewRow In grid.Rows
                    If Not fila.IsNewRow Then

                        ' 1. Obtener los valores originales de la base de datos de esta fila
                        Dim tipoOriginal As String = If(fila.Cells(0).Value?.ToString().Trim(), "")
                        Dim descOriginal As String = If(fila.Cells(1).Value?.ToString().Trim(), "")

                        ' Si la celda clave está vacía, saltamos a la siguiente fila
                        If String.IsNullOrEmpty(tipoOriginal) Then Continue For

                        ' 2. Formatear la clave base reemplazando espacios por guiones bajos (ej: "Cuenta Corriente" -> "Cuenta_Corriente")
                        Dim llaveBase As String = tipoOriginal.Replace(" ", "_")

                        ' --- TRADUCIR COLUMNA (0): Tipo de Cuenta ---
                        Dim tradTipo As String = resManager.GetString(llaveBase)
                        If Not String.IsNullOrEmpty(tradTipo) Then
                            fila.Cells(0).Value = tradTipo
                        End If

                        ' --- TRADUCIR COLUMNA (1): Descripción del Tipo ---
                        ' Buscamos usando el prefijo "Desc_" combinado con la llave del tipo
                        Dim llaveDesc As String = "Desc_" & llaveBase
                        Dim tradDesc As String = resManager.GetString(llaveDesc)

                        If Not String.IsNullOrEmpty(tradDesc) Then
                            fila.Cells(1).Value = tradDesc
                        Else
                            ' Si el usuario creó un tipo personalizado, no existirá en ResX. 
                            ' Nos aseguramos de mantener el texto original que venía de la BD.
                            fila.Cells(1).Value = descOriginal
                        End If

                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorAlEjecutar") & ": " & ex.Message, MsgBoxStyle.Exclamation, resManager.GetString("Error"))
        End Try
    End Sub

    ''' <summary>
    ''' Traduce los elementos fijos del ComboBox al idioma activo
    ''' </summary>
    Public Sub ActualizarIdiomaComboConcepto(ByVal combo As ComboBox, ByVal conespecial As Boolean)
        Try
            ' Guardamos la posición que tenía seleccionada el usuario
            Dim posicionActual As Integer = combo.SelectedIndex

            ' Limpiamos los elementos e insertamos las traducciones oficiales desde ResX
            combo.Items.Clear()
            combo.Items.Add(resManager.GetString("Tipo_Gasto")) ' Posición 0
            combo.Items.Add(resManager.GetString("Tipo_Ingreso")) ' Posición 1
            If conespecial Then
                combo.Items.Add(resManager.GetString("Tipo_Especial")) ' Posición 2
            End If
            ' Restauramos la posición del usuario de forma segura
            If posicionActual >= 0 AndAlso posicionActual < combo.Items.Count Then
                combo.SelectedIndex = posicionActual
            Else
                combo.SelectedIndex = 0
            End If
        Catch ex As Exception
            ' Evita errores si las Keys aún no están dadas de alta en el diseño del formulario
        End Try
    End Sub

    Public Sub LlenarYTraducirControlesConceptosBD(ByVal combo As ComboBox, ByVal lista As ListBox, ByVal dr As OleDb.OleDbDataReader)
        Try
            Dim posicionActual As Integer = combo.SelectedIndex

            ' 1. Cargamos los datos en memoria de forma segura SIN que rompa el flujo de Access
            Dim dt As New DataTable()
            dt.Load(dr)

            dt.Columns.Add("TextoTraducido", GetType(String))
            lista.Items.Clear()

            Dim vTipoConcepto As String = ""

            For Each fila As DataRow In dt.Rows
                Dim idConcepto As Integer = Convert.ToInt32(fila("IdConceptoCON"))
                Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
                Dim descOriginal As String = fila("DescripcionCON").ToString()
                Dim tipoOriginal As String = ""

                If dt.Columns.Contains("TipoCON") Then
                    tipoOriginal = fila("TipoCON").ToString().Trim().ToUpper()
                End If

                Dim codigoTraducido As String = ""
                If resManager IsNot Nothing Then
                    Dim claveRecurso As String = codigoOriginal.Trim().Replace(" ", "_")
                    codigoTraducido = resManager.GetString(claveRecurso)
                End If
                ' 🌟 CORRECCIÓN AQUÍ: Si no tiene traducción en el .resx, dejamos el CodigoCON
                ' corto original de la BD, NO la descripción
                If String.IsNullOrEmpty(codigoTraducido) Then codigoTraducido = codigoOriginal

                If codigoOriginal.ToUpper() = "TRASPASO" Then
                    Dim tradTraspaso As String = ""
                    If resManager IsNot Nothing Then tradTraspaso = resManager.GetString("TRASPASO")
                    If Not String.IsNullOrEmpty(tradTraspaso) Then codigoTraducido = tradTraspaso
                End If

                fila("TextoTraducido") = codigoTraducido

                ' --- AGREGAR CABECERAS DE GRUPO EN EL LISTBOX ---
                If tipoOriginal <> "" AndAlso vTipoConcepto <> tipoOriginal Then
                    vTipoConcepto = tipoOriginal

                    Select Case vTipoConcepto
                        Case "GASTO"
                            lista.Items.Add("** " & resManager.GetString("Tipo_Gasto") & " **")
                        Case "INGRESO"
                            lista.Items.Add("** " & resManager.GetString("Tipo_Ingreso") & " **")
                        Case "ESPECIAL"
                            If resManager IsNot Nothing Then
                                lista.Items.Add("** " & resManager.GetString("Tipo_Especial") & " **")
                            End If
                    End Select
                End If

                ' 🌟 CORRECCIÓN LISTBOX: Metemos objetos ElementoCombo para que guarde el ID numérico oculto
                ' Al igual que haces con las cuentas, el ListBox ahora recordará el número de ID real de cada concepto
                lista.Items.Add(codigoTraducido)
            Next

            ' 2. VINCULAMOS AL COMBOBOX
            combo.ValueMember = "IdConceptoCON"
            combo.DisplayMember = "TextoTraducido"
            combo.DataSource = dt

            If posicionActual >= 0 AndAlso posicionActual < combo.Items.Count Then
                combo.SelectedIndex = posicionActual
            ElseIf combo.Items.Count > 0 Then
                combo.SelectedIndex = 0
            End If

        Catch ex As Exception
            MsgBox("Error al sincronizar controles: " & ex.Message, MsgBoxStyle.Critical)
        Finally
            ' 🌟 EL ESCUDO DEFINITIVO: Obligamos al Reader a cerrarse pase lo que pase en el volcado
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
        End Try
    End Sub

    Public Sub LlenarYTraducirComboConceptosBD(ByVal combo As ComboBox, ByVal dr As OleDbDataReader, ByVal res As System.Resources.ResourceManager)
        Try
            Dim posicionActual As Integer = combo.SelectedIndex
            combo.Items.Clear()

            ' 🌟 CORRECCIÓN: Para evitar colgar hilos de Access, volcamos a DataTable primero
            Dim dtAux As New DataTable()
            dtAux.Load(dr)

            For Each fila As DataRow In dtAux.Rows
                Dim idConcepto As Integer = Convert.ToInt32(fila("IdConceptoCON"))
                Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
                Dim llaveBase As String = codigoOriginal.Replace(" ", "_")

                Dim codigoTraducido As String = res.GetString(llaveBase)
                If String.IsNullOrEmpty(codigoTraducido) Then codigoTraducido = codigoOriginal

                If codigoOriginal.ToUpper() = "TRASPASO" Then
                    Dim tradTraspaso As String = res.GetString("TRASPASO")
                    If Not String.IsNullOrEmpty(tradTraspaso) Then codigoTraducido = tradTraspaso
                End If

                ' Guardamos con ID numérico en el combo
                combo.Items.Add(codigoTraducido)
            Next

            If posicionActual >= 0 AndAlso posicionActual < combo.Items.Count Then
                combo.SelectedIndex = posicionActual
            ElseIf combo.Items.Count > 0 Then
                combo.SelectedIndex = 0
            End If
        Catch ex As Exception
            ' Evita cuelgues
        Finally
            ' 🌟 Liberación de seguridad inmediata
            If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
        End Try
    End Sub

    Public Sub TraducirGridApuntesBD(ByVal dgv As DataGridView)
        Try
            If dgv.Rows.Count = 0 Then Exit Sub

            For Each row As DataGridViewRow In dgv.Rows
                If row.IsNewRow Then Continue For

                Dim filaData As DataRowView = CType(row.DataBoundItem, DataRowView)

                If filaData IsNot Nothing Then

                    ' =========================================================================
                    ' 1. TRADUCCIÓN DE CONCEPTO (Celda 1 visible - CORREGIDO)
                    ' =========================================================================
                    ' Leemos el código alfanumérico estable de la base de datos (Ej: "REGULARITZACIO_1")
                    Dim codigoCON As String = filaData("CodigoCON").ToString().Trim()

                    ' Por defecto, el texto visual será el título original en mayúsculas sin guiones
                    Dim conceptoVisual As String = codigoCON.Replace("_", " ").ToUpper()

                    If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(codigoCON) Then
                        Dim claveRecurso As String = codigoCON.Replace(" ", "_")
                        Dim traduccion As String = resManager.GetString(claveRecurso)

                        ' Si el resManager encuentra el título traducido (Ej: "PHONE"), lo usamos
                        If Not String.IsNullOrEmpty(traduccion) Then
                            conceptoVisual = traduccion.Trim().ToUpper()
                        End If
                    End If

                    ' IMPACTAR: Forzamos el título en la Celda 1 (Columna de Concepto)
                    row.Cells(1).Value = conceptoVisual

                    ' =============================================================
                    ' 3. TRADUCCIÓN DE CUENTA (Celda 6)
                    ' =============================================================
                    Dim nombreCUE As String = filaData("CuentaAPU").ToString().Trim()
                    Dim cuentaVisual As String = nombreCUE

                    If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(nombreCUE) Then
                        Dim claveBase As String = nombreCUE.Replace(" ", "_")
                        Dim tradCuenta As String = resManager.GetString("Desc_" & claveBase)

                        If String.IsNullOrEmpty(tradCuenta) Then
                            tradCuenta = resManager.GetString(claveBase)
                        End If

                        If Not String.IsNullOrEmpty(tradCuenta) Then
                            cuentaVisual = tradCuenta
                        End If
                    End If

                    row.Cells(6).Value = cuentaVisual

                End If
            Next
            ' 🌟 TRUCO MAESTRO: Corrección visual automática de los IDs de Saldo
            ' --- DENTRO DE TU BUCLE DE FILAS EXISTENTE ---
            For Each fila As DataGridViewRow In dgv.Rows
                If Not fila.IsNewRow Then

                    ' Traducimos la celda del concepto (Celda 1)
                    If fila.Cells(1).Value IsNot Nothing Then
                        Dim codigoConceptoOriginal As String = fila.Cells(1).Value.ToString().Trim()

                        ' Buscamos la Key en el .resx (Pasamos a formato "ProperCase" o título si tu key es "Saldo")
                        ' Si tu clave en el .resx está guardada exactamente como "Saldo", usamos "Saldo"
                        Dim conceptoTraducido As String = resManager.GetString("Saldo")

                        ' Si el registro de la celda es cualquier otro concepto (ej: DECESOS), busca su propia clave
                        If codigoConceptoOriginal <> "SALDO" Then
                            conceptoTraducido = resManager.GetString(codigoConceptoOriginal)
                        End If

                        If Not String.IsNullOrEmpty(conceptoTraducido) Then
                            ' 🌟 EL TRUCO: Forzamos a que se guarde en la pantalla convertido a MAYÚSCULAS
                            fila.Cells(1).Value = conceptoTraducido.ToUpper()
                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            ' Previene errores visuales durante el rediseño del grid
        End Try
    End Sub

    Public Function ObtenerClaveNeutral(textoTraducido As String, rm As System.Resources.ResourceManager) As String
        ' 1. Evitamos buscar si el texto viene vacío o nulo
        If String.IsNullOrEmpty(textoTraducido) OrElse rm Is Nothing Then Return ""

        Try
            ' 2. Obtenemos el conjunto de recursos para el idioma/cultura activo en este momento
            Dim recursosActuales As System.Resources.ResourceSet =
            rm.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)

            If recursosActuales IsNot Nothing Then
                ' 3. Recorremos todos los elementos guardados en el archivo de recursos
                Dim de As System.Collections.DictionaryEntry
                For Each de In recursosActuales
                    ' 4. Comparamos el valor traducido de forma limpia (sin importar espacios ni mayúsculas)
                    If Convert.ToString(de.Value).Trim().ToUpper() = textoTraducido.Trim().ToUpper() Then
                        Return Convert.ToString(de.Key) ' ¡Éxito! Devolvemos el nombre de la clave original
                    End If
                Next
            End If
        Catch ex As Exception
            ' Si ocurre algún error en la lectura, devolvemos un texto vacío para no colgar la app
            Return ""
        End Try

        ' Si recorre todo el archivo y no encuentra coincidencia, devuelve vacío
        Return ""
    End Function

    Public Sub LlenarTempApuConceptos(Dgv As String)
        Dim filas As DataGridViewRowCollection = Nothing
        If Dgv = "CONCEPTOS_APUNTES_PERIODICOS" Then
            filas = frmApuntesPeriodicos.DgvApuper.Rows
        ElseIf Dgv = "CONCEPTOS_APUNTES_CONTABLES" Then
            filas = frmApuntesContables.DgvApuntes.Rows
        End If
        If filas IsNot Nothing Then
            For Each fila As DataGridViewRow In filas
                vImporteConcepto = fila.Cells(3).Value
                If vNombreConcepto <> fila.Cells(1).Value.ToString Then
                    vNombreConcepto = fila.Cells(1).Value.ToString
                    vImporteConcepto = ""
                    vImporteConcepto = fila.Cells(3).Value
                    ' 1. Diseñamos la estructura parametrizada limpia para la tabla temporal
                    vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                    cmdMdb1cr.CommandText = vAñadir

                    ' 2. Inyectamos los parámetros en el orden exacto de los comodines '?'
                    cmdMdb1cr.Parameters.Clear()

                    ' El concepto se limpia de apóstrofes automáticamente de forma nativa
                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                    ' Importe blindado en formato Moneda nativo de Access usando tu función global
                    Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                    paramImpTemp.Value = Math.Round(ConvertirDecimalSeguro(vImporteConcepto), 2)
                    cmdMdb1cr.CommandText = vAñadir
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro1, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                        MsgBox(ex.ToString)
                    End Try
                Else
                    cmdMdb1cr.CommandType = CommandType.Text
                    cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vExistenteImporteConcepto = drMdb1.GetValue(1)
                            End While
                        Else
                            'MsgBox("No existen registros en " + cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                        MsgBox(ex.ToString)
                    End Try
                    ' 1. Convertimos los dos importes a Decimal de forma segura (multiidioma)
                    Dim importeConcepto As Decimal = 0.0D
                    Dim existenteImporte As Decimal = 0.0D

                    ' Convertimos el primer importe (vImporteConcepto) de forma segura
                    importeConcepto = ConvertirDecimalSeguro(vImporteConcepto)

                    ' Convertimos el segundo importe (vExistenteImporteConcepto) de forma segura
                    existenteImporte = ConvertirDecimalSeguro(vExistenteImporteConcepto)

                    ' 2. Realizamos la suma matemática exacta
                    Dim sumaFinal As Decimal = Math.Round(importeConcepto + existenteImporte, 2)
                    vNewImporteConcepto = sumaFinal
                    ' 3. Preparamos la consulta SQL para Access
                    Dim vAñadir2 As String = "UPDATE tempapu SET SumaImporteAPU = ? WHERE ConceptoAPU = ?"
                    cmdMdb1cr.CommandText = vAñadir2
                    cmdMdb1cr.Parameters.Clear()
                    ' 4. CORRECCIÓN: Definimos los tipos de parámetros exactos para evitar conflictos de precisión con Access
                    ' Primero el Importe (SumaImporteAPU) indicando que es un Double/Decimal de base de datos
                    Dim pSuma As New OleDbParameter("@SumaImporte", OleDbType.Double)
                    pSuma.Value = Convert.ToDouble(sumaFinal)
                    cmdMdb1cr.Parameters.Add(pSuma)
                    ' Segundo el Concepto (ConceptoAPU)
                    Dim pConcepto As New OleDbParameter("@Concepto", OleDbType.VarChar)
                    pConcepto.Value = vNombreConcepto.ToString()
                    cmdMdb1cr.Parameters.Add(pConcepto)
                    ' 5. CORRECCIÓN CRÍTICA: Los UPDATE se ejecutan con ExecuteNonQuery, NO con ExecuteReader
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro2, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                    End Try
                End If
            Next
        End If
    End Sub

    Public Sub LlenarTempApuCuentas(Dgv As String)
        Dim filas As DataGridViewRowCollection = Nothing
        If Dgv = "CUENTAS_APUNTES_PERIODICOS" Then
            filas = frmApuntesPeriodicos.DgvApuper.Rows
        ElseIf Dgv = "CUENTAS_APUNTES_CONTABLES" Then
            filas = frmApuntesContables.DgvApuntes.Rows
        End If

        If filas IsNot Nothing Then
            For Each fila As DataGridViewRow In filas
                If fila.Cells(3).Value <> 0 Then
                    vImporteConcepto = fila.Cells(3).Value
                    If vNombreConcepto <> fila.Cells(6).Value.ToString Then
                        vNombreConcepto = fila.Cells(6).Value.ToString
                        vImporteConcepto = ""
                        vImporteConcepto = fila.Cells(3).Value
                        ' 1. Diseñamos la estructura parametrizada limpia para la tabla temporal
                        vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                        cmdMdb1cr.CommandText = vAñadir

                        ' 2. Inyectamos los parámetros en el orden exacto de los comodines '?'
                        cmdMdb1cr.Parameters.Clear()

                        ' El concepto se limpia de apóstrofes automáticamente de forma nativa
                        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                        ' Importe blindado en formato Moneda nativo de Access usando tu función global
                        Dim paramImpTemp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                        paramImpTemp.Value = Math.Round(ConvertirDecimalSeguro(vImporteConcepto), 2)
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                            MsgBox(ex.ToString)
                        End Try
                        ' 1. Diseñamos la estructura parametrizada limpia para la fila espejo
                        vAñadir = "INSERT INTO tempapu (ConceptoAPU, SumaImporteAPU) VALUES (?, 0)"
                        cmdMdb1cr.CommandText = vAñadir

                        ' 2. Inyectamos los parámetros en el orden exacto (solo el concepto, el 0 va fijo en el SQL)
                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                            MsgBox(ex.ToString)
                        End Try
                    Else ' Si el Concepto existe y hay importe diferente a cero, si es positivo o negativo se suma
                        cmdMdb1cr.CommandType = CommandType.Text

                        ' 1. Convertimos el importe a Decimal de forma segura (multiidioma)
                        Dim importeDecimal As Decimal = 0.0D

                        importeDecimal = ConvertirDecimalSeguro(vImporteConcepto)

                        ' 2. Limpiamos los parámetros previos del comando
                        cmdMdb1cr.Parameters.Clear()

                        ' 3. Evaluamos de forma exacta usando el número decimal puro
                        If importeDecimal > 0 Then
                            ' Consulta usando parámetros (?) para evitar fallos por comillas o caracteres raros
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU > 0"

                            ' Añadimos el parámetro que sustituye al primer "?"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto.ToString())

                        ElseIf importeDecimal < 0 Then
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU < 0"

                            ' Añadimos el parámetro que sustituye al primer "?"
                            cmdMdb1cr.Parameters.AddWithValue("@Concepto", vNombreConcepto.ToString())
                        End If

                        Try
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then 'Significa que existe con las condiciones
                                While drMdb1.Read()
                                    vExistenteImporteConcepto = drMdb1.GetValue(1)
                                End While
                                drMdb1.Close()
                                ' 1. Convertimos ambos importes a variables decimales exactas
                                Dim importe1 As Decimal = 0.0D
                                Dim importe2 As Decimal = 0.0D

                                ' Conversión segura del primer importe
                                importe1 = ConvertirDecimalSeguro(vImporteConcepto)

                                ' Conversión segura del segundo importe
                                importe2 = ConvertirDecimalSeguro(vExistenteImporteConcepto)

                                ' 2. Sumamos los números reales de forma exacta
                                vNewImporteConcepto = importe1 + importe2

                                If importe1 > 0 Then
                                    ' 1. Estructura parametrizada para Ingresos (> 0)
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU > 0"
                                    cmdMdb1cr.CommandText = vAñadir2

                                    ' 2. Inyectamos los parámetros en el orden exacto de los '?'
                                    cmdMdb1cr.Parameters.Clear()
                                    Dim paramSuma1 As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                                    paramSuma1.Value = Math.Round(ConvertirDecimalSeguro(vNewImporteConcepto), 2)
                                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                                ElseIf importe1 < 0 Then
                                    ' 3. Estructura parametrizada para Gastos (< 0)
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU < 0"
                                    cmdMdb1cr.CommandText = vAñadir2

                                    ' 4. Inyectamos los parámetros en el orden exacto de los '?'
                                    cmdMdb1cr.Parameters.Clear()
                                    Dim paramSuma2 As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                                    paramSuma2.Value = Math.Round(ConvertirDecimalSeguro(vNewImporteConcepto), 2)
                                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
                                End If
                                cmdMdb1cr.CommandText = vAñadir2
                                Try
                                    cmdMdb1cr.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                                    MsgBox(ex.ToString)
                                End Try
                            Else   'NO existe, lo añadimos al cero
                                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                cmdMdb1cr.CommandText += "And tempapu.SumaImporteAPU = 0 "
                                drMdb1.Close() ' Nos aseguramos de cerrar el DataReader antes de abrir uno nuevo
                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then 'Significa que existe con las condiciones
                                    While drMdb1.Read()
                                        vExistenteImporteConcepto = drMdb1.GetValue(1)
                                    End While
                                    drMdb1.Close()

                                    ' 1. Convertimos ambos importes a variables decimales exactas usando tu función
                                    Dim importe1 As Decimal = ConvertirDecimalSeguro(vImporteConcepto)
                                    Dim importe2 As Decimal = ConvertirDecimalSeguro(vExistenteImporteConcepto)

                                    ' 2. Sumamos los números reales de forma matemática y exacta
                                    Dim vNewImporteConceptoDecimal As Decimal = importe1 + importe2

                                    ' 3. Consulta parametrizada blindada contra idiomas y comillas
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? " &
                                               "WHERE tempapu.ConceptoAPU = ? AND tempapu.SumaImporteAPU = 0"

                                    cmdMdb1cr.CommandText = vAñadir2

                                    ' Limpiamos y asignamos parámetros en el orden EXACTO del SQL
                                    cmdMdb1cr.Parameters.Clear()
                                    cmdMdb1cr.Parameters.AddWithValue("@SumaImporteAPU", vNewImporteConceptoDecimal) ' Envía el Decimal puro
                                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)              ' Soporta comillas/acentos

                                    Try
                                        cmdMdb1cr.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MessageBox.Show(resManager.GetString("ErrorGrabarTemporal"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        MessageBox.Show(ex.Message, "Detalle Técnico", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                    End Try
                                End If
                                drMdb1.Close()
                            End If
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                            MsgBox(ex.ToString)
                        End Try
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub LlenarTempApuFechas(Dgv As String)
        Dim filas As DataGridViewRowCollection = Nothing
        Dim vFechaConcepto As DateTime = DateTime.MinValue ' Inicialización segura

        If Dgv = "FECHAS_APUNTES_PERIODICOS" Then
            filas = frmApuntesPeriodicos.DgvApuper.Rows
        ElseIf Dgv = "FECHAS_APUNTES_CONTABLES" Then
            filas = frmApuntesContables.DgvApuntes.Rows
        End If

        If filas IsNot Nothing Then
            For Each fila As DataGridViewRow In filas
                ' Omitir la fila nueva en blanco automática del DataGridView para evitar nulos
                If fila.IsNewRow Then Continue For

                ' Validación de que la celda del importe no esté vacía
                If fila.Cells(3).Value IsNot Nothing AndAlso IsNumeric(fila.Cells(3).Value) AndAlso Convert.ToDecimal(fila.Cells(3).Value) <> 0 Then

                    ' Conversión segura del importe de la celda
                    Dim importeFila As Decimal = Convert.ToDecimal(fila.Cells(3).Value)
                    vImporteConcepto = importeFila

                    ' Conversión segura de la fecha de la celda
                    Dim fechaFila As DateTime = Convert.ToDateTime(fila.Cells(0).Value)

                    ' Formateamos la fecha al estándar universal que Access entiende siempre (#aaaa-mm-dd#)
                    'Dim fechaFormatoAccess As String = "#" & fechaFila.ToString("yyyy-MM-dd") & "#"

                    If vFechaConcepto <> fechaFila Then
                        vFechaConcepto = fechaFila
                        vImporteConcepto = importeFila

                        ' Primer INSERT
                        ' 1. Diseñamos la estructura parametrizada limpia para la tabla de impresión
                        vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                              "VALUES (?, '', '', '', '', ?, 0)"
                        cmdMdb1cr.CommandText = vAñadir

                        ' 2. Inyectamos los parámetros en el orden exacto de los comodines '?'
                        cmdMdb1cr.Parameters.Clear()

                        ' Fecha pura en binario (Inmune a cualquier idioma de Windows)
                        ' NOTA: Pásale aquí tu variable de tipo Date real (ej: miView(x)("Fecha") o vDate) en vez de fechaFormatoAccess
                        cmdMdb1cr.Parameters.AddWithValue("@FechaTMP", fechaFila)

                        ' Importe blindado en formato Moneda nativo de Access usando tu función global
                        Dim paramImpPrint As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteTMP", OleDb.OleDbType.Currency)
                        paramImpPrint.Value = Math.Round(ConvertirDecimalSeguro(vImporteConcepto), 2)
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                        End Try

                        ' Segundo INSERT
                        ' 1. Diseñamos la estructura parametrizada limpia para la inicialización
                        vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                              "VALUES (?, '', '', '', '', 0, 0)"
                        cmdMdb1cr.CommandText = vAñadir

                        ' 2. Inyectamos únicamente el parámetro de la fecha (los ceros van fijos en el SQL)
                        cmdMdb1cr.Parameters.Clear()

                        ' Pasamos el objeto Date puro (asegúrate de que fechaFila sea tu variable Date de esa fila)
                        cmdMdb1cr.Parameters.AddWithValue("@FechaTMP", fechaFila)
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                        End Try

                    Else ' Si la fecha ya existe, sumamos o acumulamos el importe
                        cmdMdb1cr.CommandType = CommandType.Text

                        ' Construimos el SELECT filtrando por signo
                        cmdMdb1cr.CommandText = "SELECT ImporteTMP FROM tmpprint WHERE FechaTMP = ?"
                        If vImporteConcepto > 0 Then
                            cmdMdb1cr.CommandText += " AND ImporteTMP > 0"
                        Else
                            cmdMdb1cr.CommandText += " AND ImporteTMP < 0"
                        End If

                        Try
                            Dim existeRegistro As Boolean = False
                            Dim importeExistente As Decimal = 0

                            ' Usamos el Reader para capturar el valor
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then
                                existeRegistro = True
                                If drMdb1.Read() Then
                                    ' Conversión segura a Decimal del valor existente
                                    Decimal.TryParse(drMdb1.GetValue(0).ToString(), importeExistente)
                                End If
                            End If
                            drMdb1.Close() ' Es vital cerrarlo inmediatamente aquí

                            If existeRegistro Then
                                ' Calculamos el nuevo importe sumando ambos valores numéricos
                                Dim vNewImporteConcepto As Decimal = vImporteConcepto + importeExistente

                                ' Preparamos el UPDATE
                                ' 1. Construimos la consulta base usando comodines '?'
                                vAñadir2 = "UPDATE tmpprint SET ImporteTMP = ? WHERE FechaTMP = ?"

                                ' 2. Añadimos la condición dinámica al texto SQL sin mezclar variables
                                If vImporteConcepto > 0 Then
                                    vAñadir2 += " AND ImporteTMP > 0"
                                Else
                                    vAñadir2 += " AND ImporteTMP < 0"
                                End If
                                cmdMdb1cr.CommandText = vAñadir2

                                ' 3. Inyectamos los parámetros en el orden SECUENCIAL EXACTO de los '?'
                                cmdMdb1cr.Parameters.Clear()

                                ' Primer '?': El importe de la actualización (SET) blindado como Currency
                                Dim paramImpPrint As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteTMP", OleDb.OleDbType.Currency)
                                paramImpPrint.Value = Math.Round(ConvertirDecimalSeguro(vNewImporteConcepto), 2)

                                ' Segundo '?': La fecha del filtro (WHERE). 
                                ' NOTA: Pásale aquí tu variable de tipo Date real (ej: fechaFila o DateTimePicker) en lugar de fechaFormatoAccess
                                cmdMdb1cr.Parameters.AddWithValue("@FechaTMP", fechaFila)
                                cmdMdb1cr.CommandText = vAñadir2
                                Try
                                    cmdMdb1cr.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                                End Try
                            Else ' NO existe registro con ese signo, buscamos el que tiene importe = 0
                                cmdMdb1cr.CommandText = "SELECT ImporteTMP FROM tmpprint WHERE FechaTMP = ? AND ImporteTMP = 0"

                                ' 2. ¡OBLIGATORIO!: Limpiamos parámetros previos e inyectamos el nuevo
                                cmdMdb1cr.Parameters.Clear()
                                cmdMdb1cr.Parameters.AddWithValue("@FechaTMP", fechaFila) ' Usamos la fecha de la fila actual

                                Dim existeCero As Boolean = False
                                Dim importeCeroExistente As Decimal = 0

                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then
                                    existeCero = True
                                    If drMdb1.Read() Then
                                        Decimal.TryParse(drMdb1.GetValue(0).ToString(), importeCeroExistente)
                                    End If
                                End If
                                drMdb1.Close()

                                If existeCero Then
                                    Dim vNewImporteConcepto As Decimal = vImporteConcepto + importeCeroExistente

                                    ' 1. Diseñamos la estructura parametrizada limpia para la actualización
                                    vAñadir2 = "UPDATE tmpprint SET ImporteTMP = ? WHERE FechaTMP = ? AND ImporteTMP = 0"
                                    cmdMdb1cr.CommandText = vAñadir2

                                    ' 2. Inyectamos los parámetros en el orden SECUENCIAL EXACTO de los comodines '?'
                                    cmdMdb1cr.Parameters.Clear()

                                    ' Primer '?': El importe de la actualización (SET) blindado como Currency
                                    Dim paramImpPrint As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteTMP", OleDb.OleDbType.Currency)
                                    paramImpPrint.Value = Math.Round(ConvertirDecimalSeguro(vNewImporteConcepto), 2)

                                    ' Segundo '?': La fecha del filtro (WHERE) como objeto Date puro
                                    ' (Recuerda apuntar a tu variable Date real de ese bucle, por ejemplo, fechaFila)
                                    cmdMdb1cr.Parameters.AddWithValue("@FechaTMP", fechaFila)
                                    cmdMdb1cr.CommandText = vAñadir2
                                    Try
                                        cmdMdb1cr.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                                    End Try
                                End If
                            End If
                        Catch ex As Exception
                            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                            MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                        End Try
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub LlenarTempApuMeses(Dgv As String)
        Dim filas As DataGridViewRowCollection = Nothing

        If Dgv = "MESES_APUNTES_PERIODICOS" Then
            filas = frmApuntesPeriodicos.DgvApuper.Rows
        ElseIf Dgv = "MESES_APUNTES_CONTABLES" Then
            filas = frmApuntesContables.DgvApuntes.Rows
        End If

        If filas IsNot Nothing Then
            For Each fila As DataGridViewRow In filas
                ' Evitamos procesar la fila nueva vacía automática de .NET
                If fila.IsNewRow Then Continue For

                ' Forzamos la lectura limpia del importe de la celda actual usando tu función del módulo
                Dim importeCeldaSeguro As Decimal = ConvertirDecimalSeguro(fila.Cells(3).Value)

                ' Si el importe es 0, no hay nada que acumular en este mes
                If importeCeldaSeguro <> 0 Then
                    vImporteConcepto = importeCeldaSeguro

                    ' Extraemos de forma genérica el Año-Mes de la celda de fecha
                    Dim fechaReal As DateTime = Convert.ToDateTime(fila.Cells(0).Value)
                    Dim claveMesAño As String = fechaReal.ToString("yy") & "-" & fechaReal.Month.ToString("D2")

                    If vNombreConcepto <> claveMesAño Then
                        vNombreConcepto = claveMesAño

                        ' Inserción 1: Del importe real (Parametrizada y segura)
                        vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES (?, ?)"
                        cmdMdb1cr.CommandText = vAñadir
                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
                        Dim paramInsert As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                        paramInsert.Value = Math.Round(ConvertirDecimalSeguro(vImporteConcepto), 2)
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Error Insert Real", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try

                        ' Inserción 2: De la fila espejo a cero (Parametrizada y segura)
                        vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES (?, 0)"
                        cmdMdb1cr.CommandText = vAñadir
                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MessageBox.Show(ex.Message, "Error Insert Cero", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End Try
                    Else
                        ' Si ya existe el registro del mes actual en tempapu, actualizamos acumulando el importe
                        cmdMdb1cr.CommandType = CommandType.Text

                        ' Evaluamos la cantidad real con céntimos incluidos y preparamos la query exacta
                        If importeCeldaSeguro > 0 Then
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU > 0"
                        Else
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU < 0"
                        End If

                        cmdMdb1cr.Parameters.Clear()
                        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                        Try
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then
                                While drMdb1.Read()
                                    vExistenteImporteConcepto = drMdb1.GetValue(1)
                                End While
                                drMdb1.Close() ' Cerramos inmediatamente el reader

                                ' 1. Conversión y suma matemática exacta con tipos Decimal
                                Dim imp1 As Decimal = ConvertirDecimalSeguro(vImporteConcepto)
                                Dim imp2 As Decimal = ConvertirDecimalSeguro(vExistenteImporteConcepto)
                                ' Sumamos y redondeamos estrictamente a 2 decimales para que quepa en el campo de Access
                                Dim vNewImporteConceptoDecimal As Decimal = Math.Round(imp1 + imp2, 2)

                                ' 2. Construimos la query parametrizada limpia sin mezclas estáticas
                                vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ? "
                                vAñadir2 += If(imp1 > 0, "And tempapu.SumaImporteAPU > 0", "And tempapu.SumaImporteAPU < 0")

                                cmdMdb1cr.CommandText = vAñadir2

                                ' ¡LIMPIEZA RADICAL! Vaciamos por completo el comando antes de asignar los nuevos parámetros
                                cmdMdb1cr.Parameters.Clear()
                                ' Obligamos al parámetro a comportarse como Moneda pura para que Access no se sature con la precisión
                                Dim paramSuma1 As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                                paramSuma1.Value = Math.Round(vNewImporteConceptoDecimal, 2)
                                cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                                Try
                                    cmdMdb1cr.ExecuteNonQuery()
                                Catch ex As Exception
                                    MessageBox.Show(ex.Message, "Error Update 1 (Interno)", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End Try
                            Else
                                drMdb1.Close()

                                ' Si no existe, acumulamos sobre el registro que se creó a cero (Consulta parametrizada)
                                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU = 0"
                                cmdMdb1cr.Parameters.Clear()
                                cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then
                                    While drMdb1.Read()
                                        vExistenteImporteConcepto = drMdb1.GetValue(1)
                                    End While
                                    drMdb1.Close()

                                    ' Conversión y suma exacta para el segundo caso
                                    Dim impActual As Decimal = ConvertirDecimalSeguro(vImporteConcepto)
                                    Dim impExistente As Decimal = ConvertirDecimalSeguro(vExistenteImporteConcepto)
                                    ' Hacemos lo mismo en el segundo acumulador por seguridad
                                    Dim vNewImporteConceptoDecimal2 As Decimal = Math.Round(impActual + impExistente, 2)

                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = ? WHERE tempapu.ConceptoAPU = ? And tempapu.SumaImporteAPU = 0"
                                    cmdMdb1cr.CommandText = vAñadir2

                                    cmdMdb1cr.Parameters.Clear()
                                    ' Aplicamos el mismo blindaje en el segundo acumulador
                                    Dim paramSuma2 As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@SumaImporteAPU", OleDb.OleDbType.Currency)
                                    paramSuma2.Value = Math.Round(vNewImporteConceptoDecimal2, 2)
                                    cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vNombreConcepto)

                                    Try
                                        cmdMdb1cr.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MessageBox.Show(ex.Message, "Error Update 2", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                    End Try
                                Else
                                    drMdb1.Close()
                                End If
                            End If
                        Catch ex As Exception
                            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                            MessageBox.Show(ex.Message, "Error General en Rama Else", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End Try
                    End If
                End If
            Next
        End If
    End Sub

    Public Sub LimpiarTempApu()
        Dim vBorrar As String = "DELETE FROM tempapu"
        cmdMdb1cr.CommandText = vBorrar
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Tempapu, Limpiada Correctamente")
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorLimpiarTemporal"))
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Sub LimpiarTempPrint()
        Dim vtmpprint As String = "DELETE FROM tmpprint"
        cmdMdb1cr.CommandText = vtmpprint
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Registros tmpprint, Borrados !!!")
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorLimpiarTemporal"))
            MsgBox(ex.ToString)
        End Try
    End Sub

    Public Function ConvertirDecimalSeguro(vValor As Object) As Decimal
        Dim importeResultado As Decimal = 0.0D
        If vValor IsNot Nothing AndAlso vValor IsNot DBNull.Value Then
            Dim textoImporteRaw As String = vValor.ToString().Trim()

            ' 1. Intento con cultura local de Windows
            If Not Decimal.TryParse(textoImporteRaw,
                                System.Globalization.NumberStyles.Number,
                                System.Globalization.CultureInfo.CurrentCulture,
                                importeResultado) Then

                ' 2. Plan B con cultura invariante (punto universal)
                Decimal.TryParse(textoImporteRaw,
                             System.Globalization.NumberStyles.Number,
                             System.Globalization.CultureInfo.InvariantCulture,
                             importeResultado)
            End If
        End If
        Return importeResultado
    End Function

    Public Sub LlenarComboConceptoExclusivoTraspaso(ByVal combo As ComboBox)
        combo.DataSource = Nothing
        combo.Items.Clear()

        ' 1. Buscamos de forma directa el registro exacto de 'TRASPASO' en tu Access
        Dim sql As String = "SELECT IdConceptoCON, CodigoCON, DescripcionCON FROM conceptos WHERE CodigoCON = 'TRASPASO'"

        Dim dtConceptos As New DataTable()

        Using cmd As New OleDbCommand(sql, conexion1)
            Dim dr As OleDbDataReader = Nothing
            Try
                dr = cmd.ExecuteReader()
                dtConceptos.Load(dr)
            Catch ex As Exception
                MsgBox("Error al localizar el concepto TRASPASO: " & ex.Message, MsgBoxStyle.Critical)
            Finally
                If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            End Try
        End Using

        ' 2. Si existe el registro, le aplicamos tu traducción oficial (.resx)
        If dtConceptos.Rows.Count > 0 Then
            dtConceptos.Columns.Add("TextoCombo", GetType(String))

            Dim fila As DataRow = dtConceptos.Rows(0)
            Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
            Dim textoFinal As String = codigoOriginal ' Mantiene "TRASPASO" por defecto

            ' Si tu archivo de recursos tiene traducida la palabra, la hereda
            If resManager IsNot Nothing Then
                Dim trad As String = resManager.GetString("TRASPASO")
                If Not String.IsNullOrEmpty(trad) Then textoFinal = trad
            End If

            fila("TextoCombo") = textoFinal

            ' 3. Vinculamos de golpe el combo con el ID numérico correspondiente
            combo.ValueMember = "IdConceptoCON"       ' Mantiene el ID numérico real de la BD
            combo.DisplayMember = "TextoCombo"        ' Enseña la palabra limpia (TRASPASO)
            combo.DataSource = dtConceptos

            ' Forzamos la selección del único elemento y bloqueamos el combo por seguridad
            combo.SelectedIndex = 0
            combo.Enabled = False ' 🌟 Opcional: Bloquea el combo para que sea meramente informativo
        Else
            MsgBox("Atención: No se ha encontrado el concepto obligatorio 'TRASPASO' en la base de datos.", MsgBoxStyle.Exclamation)
        End If
    End Sub


    Public Sub LlenarComboConceptosGenerico(ByVal combo As ComboBox)
        ' 1. Sincronizados los nombres de las columnas.
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, DescripcionCON FROM conceptos ORDER BY TipoCON ASC, IdConceptoCON ASC"

        Dim dtConceptos As New DataTable()

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            dtConceptos.Load(drMdb1)
            drMdb1.Close()

            ' Creamos la columna virtual para el texto traducido que verá el usuario
            dtConceptos.Columns.Add("TextoTraducido", GetType(String))

            ' Recorremos las filas para traducir cada concepto
            For Each fila As DataRow In dtConceptos.Rows
                Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
                Dim descOriginal As String = fila("DescripcionCON").ToString()
                Dim textoFinal As String = descOriginal ' Salvavidas por defecto

                ' Si tenemos el gestor de recursos, buscamos la traducción del concepto
                If resManager IsNot Nothing Then
                    ' 1. EL TRUCO CONTRA ESPACIOS DOBLES: Reemplazamos bucles de espacios dobles por uno solo
                    Dim textoLimpio As String = codigoOriginal.Trim()
                    'While textoLimpio.Contains("  ")
                    '    textoLimpio = textoLimpio.Replace("  ", " ")
                    'End While
                    ' 2. Ahora que seguro solo hay UN espacio, hacemos el cambio a guion bajo
                    Dim claveRecurso As String = textoLimpio.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)

                    If Not String.IsNullOrEmpty(traduccion) Then
                        textoFinal = traduccion
                    End If
                End If

                fila("TextoTraducido") = textoFinal
            Next

            ' 2. VINCULAMOS AL COMBOBOX PASADO POR PARÁMETRO (Con el orden óptimo de Windows Forms)
            combo.ValueMember = "IdConceptoCON"       ' El número oculto (1, 2, 3...)
            combo.DisplayMember = "TextoTraducido"   ' Lo que VE el usuario
            combo.DataSource = dtConceptos            ' Al final enlazamos los datos

        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox("Error al cargar los conceptos: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub LlenarComboCuentasGenerico(ByVal combo As ComboBox)
        ' 1. SQL adaptado: traemos el IdCUE y el NombreCUE de tu nueva tabla cuentas
        cmdMdb1cr.CommandText = "SELECT IdCuentaCUE, NombreCUE FROM cuentas ORDER BY NombreCUE ASC"

        Dim dtCuentas As New DataTable()

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            dtCuentas.Load(drMdb1)
            drMdb1.Close()

            ' Creamos la columna virtual para el texto traducido que verá el usuario
            dtCuentas.Columns.Add("NombreTraducido", GetType(String))

            ' Recorremos las filas para traducir cada cuenta de forma directa e invariant
            For Each fila As DataRow In dtCuentas.Rows
                Dim nombreOriginal As String = fila("NombreCUE").ToString().Trim()
                Dim textoFinal As String = nombreOriginal ' Salvavidas por defecto

                ' Si tenemos el gestor de recursos, buscamos la traducción de la cuenta
                If resManager IsNot Nothing Then
                    Dim textoLimpio As String = nombreOriginal.Trim()
                    'While textoLimpio.Contains("  ")
                    '    textoLimpio = textoLimpio.Replace("  ", " ")
                    'End While
                    ' 2. Ahora que seguro solo hay UN espacio, hacemos el cambio a guion bajo
                    Dim claveRecurso As String = textoLimpio.Replace(" ", "_")
                    Dim traduccion As String = resManager.GetString(claveRecurso)

                    If Not String.IsNullOrEmpty(traduccion) Then
                        textoFinal = traduccion
                    End If
                End If

                fila("NombreTraducido") = textoFinal
            Next

            ' 2. VINCULAMOS AL COMBOBOX (Con el orden óptimo e idéntico de Windows Forms)
            combo.ValueMember = "IdCuentaCUE"             ' El número oculto (1, 2, 3...) de tu Access
            combo.DisplayMember = "NombreTraducido" ' Lo que VE el usuario en la pantalla
            combo.DataSource = dtCuentas            ' Al final de todo enlazamos los datos

        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox("Error al cargar las cuentas desde el módulo: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    ' === LISTA FIJA DE PROTECCIÓN (TUS CONCEPTOS DE MUESTRA ORIGINALES) LISTA CONCEPTOS Y LISTA CUENTAS=== 
    ' Escribe aquí en mayúsculas los 33 códigos exactos que metes de fábrica en la mdb
    Public ReadOnly Property ConceptosMuestraSistema As New List(Of String)(New String() {
    "AGUA", "ALIMENTACION", "CANAL+", "CASA", "CLIENTE00", "COMUNIDAD", "DECESOS", "EL CORTE INGLES", "ESTETICA", "FARMACIA", "GASNATURAL", "GASOLINA", "GASTOS BANCARIOS",
    "HACIENDA", "IMPUESTO 1", "IMPUESTO 2", "IMPUESTO 3", "IMPUESTO 4", "IMPUESTO 5", "INTERESES", "JARDIN", "LUZ",
    "OCIO", "PENSION", "REGULARIZACION 1", "REGULARIZACION 2", "ROPA", "SEGURO CASA", "SEGURO COCHE", "SEGURO MOTO", "TELEFONO", "VARIOS", "VEHICULOS"
})

    Public ReadOnly Property TiposCuentaMuestraSistema As New List(Of String)(New String() {
    "CUENTACORRIENTE", "CUENTAVIVIENDA", "EFECTIVO", "FONDODEINVERSION", "PLANDEPENSIONES", "TARJETADECREDITO"
})

    Public ReadOnly Property CuentasMuestraSistema As New List(Of String)(New String() {
    "BBVA", "CAJAEFECTIVO", "OPENBANK", "PLANPENSIONES"
})

    Public Sub VerificarYActualizarEstructuraBD()
        Dim necesitaActualizar As Boolean = False

        ' =========================================================================
        ' PASO 1: DETECCIÓN AUTOMÁTICA DE LA VERSIÓN DE LA BASE DE DATOS
        ' =========================================================================
        Try
            cmdMdb1cr.Connection = conexion1
            cmdMdb1cr.CommandText = "SELECT TOP 1 ConceptoAPU FROM apuntes"
            cmdMdb1cr.Parameters.Clear()

            Using adapter As New OleDb.OleDbDataAdapter(cmdMdb1cr)
                Dim dtPrueba As New DataTable()
                adapter.Fill(dtPrueba)
                ' Si la columna ConceptoAPU sigue siendo de tipo Texto/String, hay que migrar
                If dtPrueba.Columns("ConceptoAPU").DataType = GetType(String) Then
                    necesitaActualizar = True
                End If
            End Using
        Catch ex As Exception
            Exit Sub ' Si la tabla no responde, cancelamos para evitar cuelgues
        End Try

        ' Si ya es una base de datos moderna con IDs numéricos, salimos sin hacer nada
        If Not necesitaActualizar Then Exit Sub

        ' =========================================================================
        ' 🌟 PASO 1.5: LIMPIEZA RADICAL ANTES DE CAMBIAR LOS CAMPOS (BD ANTIGUA DETECTADA)
        ' =========================================================================
        ' Como todavía es de tipo Texto, borramos de golpe los registros "SALDO" viejos
        cmdMdb1cr.CommandText = "DELETE FROM apuntes WHERE ConceptoAPU = 'SALDO'"
        cmdMdb1cr.Parameters.Clear()
        Try
            cmdMdb1cr.ExecuteNonQuery()
        Catch ex As Exception
            ' Evita cuelgues si la tabla está vacía en algún entorno de pruebas
        End Try

        ' =========================================================================
        ' PASO 2: MIGRACIÓN ESTRUCTURAL COMPLETA Y CONSERVACIÓN DE DATOS
        ' =========================================================================
        Try
            ' ---------------------------------------------------------------------
            ' A. CREACIÓN DE COLUMNAS TEMPORALES (Fase de preparación)
            ' ---------------------------------------------------------------------
            ' Auxiliares para Apuntes
            Try : cmdMdb1cr.CommandText = "ALTER TABLE apuntes ADD COLUMN ConceptoAPU_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try
            Try : cmdMdb1cr.CommandText = "ALTER TABLE apuntes ADD COLUMN CuentaAPU_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try

            Try : cmdMdb1cr.CommandText = "ALTER TABLE conceptos ADD COLUMN IdConceptoCON INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try
            Try : cmdMdb1cr.CommandText = "ALTER TABLE cuentas ADD COLUMN IdCuentaCUE INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try


            ' Auxiliares para Apuntes Periódicos (apuper)
            Try : cmdMdb1cr.CommandText = "ALTER TABLE apuper ADD COLUMN ConceptoAPP_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try
            Try : cmdMdb1cr.CommandText = "ALTER TABLE apuper ADD COLUMN CuentaAPP_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try

            ' Auxiliar para Presupuestos
            Try : cmdMdb1cr.CommandText = "ALTER TABLE presupuesto ADD COLUMN ConceptoPRE_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try

            ' Auxiliar para Cuentas (Enlace a Tipo de Cuenta)
            Try : cmdMdb1cr.CommandText = "ALTER TABLE cuentas ADD COLUMN TipoCUE_NEW INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try

            ' Campo ID para la tabla maestra de tipos de cuentas
            Try : cmdMdb1cr.CommandText = "ALTER TABLE tipocuentas ADD COLUMN IdTipoCUE INTEGER" : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : End Try


            ' =========================================================================
            ' B. CARGA DE EQUIVALENCIAS Y ASIGNACIÓN DE IDS EN MEMORIA
            ' =========================================================================

            ' ---------------------------------------------------------------------
            ' 1. PRIMERO UNIFICAMOS LOS TEXTOS EN LA BD FÍSICA (Antes de leer a memoria)
            ' ---------------------------------------------------------------------
            Try
                ' Normalizamos los nombres antiguos con espacios o tildes al nuevo formato oficial con guion bajo
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = 'CUENTA_CORRIENTE' WHERE CodigoTIP = 'CUENTA CORRIENTE'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = 'CUENTA_VIVIENDA' WHERE CodigoTIP = 'CUENTA VIVIENDA'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = 'FONDO_DE_INVERSION' WHERE CodigoTIP = 'FONDO DE INVERSION' OR CodigoTIP = 'FONDO DE INVERSIÓN'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = 'PLAN_DE_PENSIONES' WHERE CodigoTIP = 'PLAN DE PENSIONES'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = 'TARJETA_DE_CREDITO' WHERE CodigoTIP = 'TARJETA DE CREDITO' OR CodigoTIP = 'TARJETA DE CRÉDITO'" : cmdMdb1cr.ExecuteNonQuery()

                ' Sincronizamos el campo TipoCUE de la tabla cuentas para que use los mismos guiones bajos
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = 'CUENTA_CORRIENTE' WHERE TipoCUE = 'CUENTA CORRIENTE'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = 'CUENTA_VIVIENDA' WHERE TipoCUE = 'CUENTA VIVIENDA'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = 'FONDO_DE_INVERSION' WHERE TipoCUE = 'FONDO DE INVERSION' OR TipoCUE = 'FONDO DE INVERSIÓN'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = 'PLAN_DE_PENSIONES' WHERE TipoCUE = 'PLAN DE PENSIONES'" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = 'TARJETA_DE_CREDITO' WHERE TipoCUE = 'TARJETA DE CREDITO' OR TipoCUE = 'TARJETA DE CRÉDITO'" : cmdMdb1cr.ExecuteNonQuery()

                ' Forzamos la conversión a MAYÚSCULAS puras en todo el esquema de la MDB
                cmdMdb1cr.CommandText = "UPDATE conceptos SET CodigoCON = UCASE(CodigoCON)" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET CodigoTIP = UCASE(CodigoTIP)" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = UCASE(TipoCUE)" : cmdMdb1cr.ExecuteNonQuery()
                cmdMdb1cr.CommandText = "UPDATE cuentas SET NombreCUE = UCASE(NombreCUE)" : cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                ' Evita bloqueos si algún registro no existiera en la BD de pruebas
            End Try

            ' ---------------------------------------------------------------------
            ' 2. LEEMOS LAS TABLAS YA NORMALIZADAS PARA CREAR LAS LISTAS EN MEMORIA
            ' ---------------------------------------------------------------------

            ' Mapeo de Conceptos
            cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON"
            cmdMdb1cr.Parameters.Clear()
            Dim listaConceptos As New List(Of KeyValuePair(Of Integer, String))()
            Dim contadorConcepto As Integer = 1
            Using reader As OleDb.OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While reader.Read()
                    listaConceptos.Add(New KeyValuePair(Of Integer, String)(contadorConcepto, reader("CodigoCON").ToString().Trim()))
                    contadorConcepto += 1
                End While
            End Using

            ' Mapeo de Cuentas (Aquí leerá "CAJA EFECTIVO" o "BBVA" en mayúsculas fijas)
            cmdMdb1cr.CommandText = "SELECT NombreCUE FROM cuentas ORDER BY NombreCUE"
            cmdMdb1cr.Parameters.Clear()
            Dim listaCuentas As New List(Of KeyValuePair(Of Integer, String))()
            Dim contadorCuenta As Integer = 1
            Using reader As OleDb.OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While reader.Read()
                    listaCuentas.Add(New KeyValuePair(Of Integer, String)(contadorCuenta, reader("NombreCUE").ToString().Trim()))
                    contadorCuenta += 1
                End While
            End Using

            ' Mapeo de Tipos de Cuentas (Leerá "EFECTIVO", "CUENTA_CORRIENTE", etc.)
            cmdMdb1cr.CommandText = "SELECT CodigoTIP FROM tipocuentas ORDER BY CodigoTIP"
            cmdMdb1cr.Parameters.Clear()
            Dim listaTipos As New List(Of KeyValuePair(Of Integer, String))()
            Dim contadorTipo As Integer = 1
            Using reader As OleDb.OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While reader.Read()
                    listaTipos.Add(New KeyValuePair(Of Integer, String)(contadorTipo, reader("CodigoTIP").ToString().Trim()))
                    contadorTipo += 1
                End While
            End Using


            ' ---------------------------------------------------------------------
            ' C. VOLCADO Y ACTUALIZACIÓN CRUZADA DE DATOS (Fase de inyección aislada)
            ' ---------------------------------------------------------------------

            ' --- 1. Inyectar números en la tabla maestra: conceptos ---
            For Each item In listaConceptos
                cmdMdb1cr.CommandText = "UPDATE conceptos SET IdConceptoCON = ? WHERE CodigoCON = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' --- 2. Inyectar números en la tabla maestra: cuentas ---
            For Each item In listaCuentas
                cmdMdb1cr.CommandText = "UPDATE cuentas SET IdCuentaCUE = ? WHERE NombreCUE = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' --- 3. Inyectar números en la tabla maestra: tipocuentas ---
            For Each item In listaTipos
                cmdMdb1cr.CommandText = "UPDATE tipocuentas SET IdTipoCUE = ? WHERE CodigoTIP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' --- 4. Actualizar tablas de movimientos y enlaces históricos ---
            ' Movimientos de Conceptos
            For Each item In listaConceptos
                cmdMdb1cr.CommandText = "UPDATE apuntes SET ConceptoAPU_NEW = ? WHERE ConceptoAPU = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()

                cmdMdb1cr.CommandText = "UPDATE apuper SET ConceptoAPP_NEW = ? WHERE ConceptoAPP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()

                cmdMdb1cr.CommandText = "UPDATE presupuesto SET ConceptoPRE_NEW = ? WHERE ConceptoPRE = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' Movimientos de Cuentas
            For Each item In listaCuentas
                cmdMdb1cr.CommandText = "UPDATE apuntes SET CuentaAPU_NEW = ? WHERE CuentaAPU = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()

                cmdMdb1cr.CommandText = "UPDATE apuper SET CuentaAPP_NEW = ? WHERE CuentaAPP = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' Enlace de Tipo de Cuenta dentro de la tabla Cuentas
            For Each item In listaTipos
                cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE_NEW = ? WHERE TipoCUE = ?"
                cmdMdb1cr.Parameters.Clear()
                cmdMdb1cr.Parameters.AddWithValue("?", item.Key) : cmdMdb1cr.Parameters.AddWithValue("?", item.Value) : cmdMdb1cr.ExecuteNonQuery()
            Next

            ' ---------------------------------------------------------------------
            ' D. REESTRUCTURACIÓN FINAL Y LIMPIEZA DE TABLAS (Fase de consolidación)
            ' ---------------------------------------------------------------------

            ' 1. Reestructurar Tabla: apuntes
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes DROP COLUMN ConceptoAPU" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes DROP COLUMN CuentaAPU" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes ADD COLUMN ConceptoAPU INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes ADD COLUMN CuentaAPU INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "UPDATE apuntes SET ConceptoAPU = ConceptoAPU_NEW, CuentaAPU = CuentaAPU_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes DROP COLUMN ConceptoAPU_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuntes DROP COLUMN CuentaAPU_NEW" : cmdMdb1cr.ExecuteNonQuery()

            ' 2. Reestructurar Tabla: apuper
            cmdMdb1cr.CommandText = "ALTER TABLE apuper DROP COLUMN ConceptoAPP" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuper DROP COLUMN CuentaAPP" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuper ADD COLUMN ConceptoAPP INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuper ADD COLUMN CuentaAPP INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "UPDATE apuper SET ConceptoAPP = ConceptoAPP_NEW, CuentaAPP = CuentaAPP_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuper DROP COLUMN ConceptoAPP_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE apuper DROP COLUMN CuentaAPP_NEW" : cmdMdb1cr.ExecuteNonQuery()

            ' 3. Reestructurar Tabla: presupuesto
            cmdMdb1cr.CommandText = "ALTER TABLE presupuesto DROP COLUMN ConceptoPRE" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE presupuesto ADD COLUMN ConceptoPRE INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "UPDATE presupuesto SET ConceptoPRE = ConceptoPRE_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE presupuesto DROP COLUMN ConceptoPRE_NEW" : cmdMdb1cr.ExecuteNonQuery()

            ' 4. Reestructurar Tabla: cuentas (Enlace de tipos)
            cmdMdb1cr.CommandText = "ALTER TABLE cuentas DROP COLUMN TipoCUE" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE cuentas ADD COLUMN TipoCUE INTEGER" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "UPDATE cuentas SET TipoCUE = TipoCUE_NEW" : cmdMdb1cr.ExecuteNonQuery()
            cmdMdb1cr.CommandText = "ALTER TABLE cuentas DROP COLUMN TipoCUE_NEW" : cmdMdb1cr.ExecuteNonQuery()

            MsgBox("La estructura de la base de datos se ha actualizado de manera integral a la nueva versión.", vbInformation, "Actualización Completada")
        Catch ex As Exception
            MsgBox("Error crítico durante la reestructuración completa: " & vbNewLine & ex.Message, vbCritical, "Migración Interrumpida")
        End Try

        ' =========================================================================
        ' 🌟 PASO FINAL: CREACIÓN Y BLINDAJE DEL CONCEPTO MAESTRO "SALDO" (ID MANUAL)
        ' =========================================================================
        ' Una vez que los campos ya cambiaron a número, creamos el concepto de fábrica
        Dim siguienteIdCON As Integer = 1
        cmdMdb1cr.CommandText = "SELECT MAX(IdConceptoCON) FROM conceptos"
        cmdMdb1cr.Parameters.Clear()
        Try
            Dim resMax = cmdMdb1cr.ExecuteScalar()
            If resMax IsNot Nothing AndAlso Not IsDBNull(resMax) Then
                siguienteIdCON = Convert.ToInt32(resMax) + 1
            End If
        Catch ex As Exception
            siguienteIdCON = 1
        End Try

        ' Insertamos el concepto parametrizado con los 4 atributos exactos
        cmdMdb1cr.CommandText = "INSERT INTO conceptos (IdConceptoCON, CodigoCON, DescripcionCON, TipoCON) VALUES (?, ?, ?, ?)"
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("?", siguienteIdCON)
        cmdMdb1cr.Parameters.AddWithValue("?", "SALDO")
        cmdMdb1cr.Parameters.AddWithValue("?", "Saldo Inicial")
        cmdMdb1cr.Parameters.AddWithValue("?", "ESPECIAL") ' Blindaje para que no lo editen ni eliminen
        Try
            cmdMdb1cr.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Error al inyectar el concepto maestro 'SALDO': " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Public Sub LlenarComboConceptosIntroApuntes(ByVal combo As ComboBox)
        combo.DataSource = Nothing
        combo.Items.Clear()

        ' FILTRADO QUIRÚRGICO: Excluimos 'ESPECIAL' de la pantalla de introducción
        Dim sql As String = "SELECT IdConceptoCON, CodigoCON, DescripcionCON FROM conceptos " &
                            "WHERE TipoCON <> 'ESPECIAL' " &
                            "ORDER BY TipoCON ASC, IdConceptoCON ASC"

        Dim dtConceptos As New DataTable()

        Using cmd As New OleDbCommand(sql, conexion1)
            Dim dr As OleDbDataReader = Nothing
            Try
                dr = cmd.ExecuteReader()
                dtConceptos.Load(dr)
            Catch ex As Exception
                MsgBox("Error al leer conceptos para introducción: " & ex.Message, MsgBoxStyle.Critical)
            Finally
                If dr IsNot Nothing AndAlso Not dr.IsClosed Then dr.Close()
            End Try
        End Using

        ' Creamos la columna virtual para el texto del combo
        dtConceptos.Columns.Add("TextoCombo", GetType(String))

        For Each fila As DataRow In dtConceptos.Rows
            Dim codigoOriginal As String = fila("CodigoCON").ToString().Trim()
            Dim textoFinal As String = codigoOriginal ' 🌟 SALVAVIDAS ORIGINAL: Mantiene siempre el CodigoCON corto

            ' Si existe traducción en el .resx para el código corto, la aplicamos
            If resManager IsNot Nothing Then
                Dim claveRecurso As String = codigoOriginal.Replace(" ", "_")
                Dim traduccion As String = resManager.GetString(claveRecurso)

                If Not String.IsNullOrEmpty(traduccion) Then
                    textoFinal = traduccion
                End If
            End If

            ' Caso especial para el Traspaso del sistema
            If codigoOriginal.ToUpper() = "TRASPASO" AndAlso resManager IsNot Nothing Then
                Dim tradTraspaso As String = resManager.GetString("TRASPASO")
                If Not String.IsNullOrEmpty(tradTraspaso) Then textoFinal = tradTraspaso
            End If

            fila("TextoCombo") = textoFinal
        Next

        ' VINCULAMOS AL COMBOBOX CON ID NUMÉRICOS
        combo.ValueMember = "IdConceptoCON"       ' ID numérico oculto para guardar en la BD
        combo.DisplayMember = "TextoCombo"        ' Muestra el CodigoCON corto (o traducido)
        combo.DataSource = dtConceptos
    End Sub

    Public Function ReemplazarPrimerInterrogante(ByVal textoOriginal As String, ByVal valorReemplazo As String) As String
        Dim posicion As Integer = textoOriginal.IndexOf("?")
        If posicion >= 0 Then
            ' Cortamos la cadena en el signo '?' e inyectamos el valor real en medio
            Return textoOriginal.Substring(0, posicion) & valorReemplazo & textoOriginal.Substring(posicion + 1)
        End If
        Return textoOriginal
    End Function

    Public Sub RefrescarGridApuntesContables()
        ' 🌟 SANEAMIENTO PREVENTIVO: Limpiamos la memoria de consultas anteriores
        cmdMdb1cr.Parameters.Clear()

        ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' Guardamos en booleanos el estado de los filtros para no liar la SQL
        Dim filtroCuentaActivo As Boolean = (frmApuntesContables.BtnFiltroCuenta.Enabled = False)
        Dim filtroConceptoActivo As Boolean = (frmApuntesContables.BtnFiltroConcepto.Enabled = False)
        Dim filtroFechaActivo As Boolean = (frmApuntesContables.BtnFiltroFecha.Enabled = False)

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS RELACIONALES (Nombres traducidos y legibles)
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        ' Condición base del año contable o descarte de saldos
        If frmApuntesContables.BtnFechasClick = "SI" Then
            vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If

        ' 2. CONTROL DE FILTRADO DEL LISTBOX DE CONCEPTOS (MULTISELECCIÓN)
        If frmApuntesContables.ListBox1.SelectedItems.Count > 0 Then
            Dim listaIds As New List(Of Integer)
            For i As Integer = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                Dim vConceptoFila As String = frmApuntesContables.ListBox1.SelectedItems(i).ToString()
                If vConceptoFila.StartsWith("**") Then Continue For

                Dim idEncontrado As Integer = 0
                Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                    cmdId.Parameters.AddWithValue("?", vConceptoFila)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idEncontrado = Convert.ToInt32(resId)
                End Using
                If idEncontrado > 0 Then listaIds.Add(idEncontrado)
            Next
            If listaIds.Count = 0 Then listaIds.Add(0)
            vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIds) & ") "
        Else
            ' Filtro individual por combo de concepto si la lista lateral está vacía
            If filtroConceptoActivo Then
                Dim idConceptoSel As Integer = Convert.ToInt32(frmApuntesContables.CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If
        End If

        ' 3. FILTRO POR ID NUMÉRICO DE CUENTA
        If filtroCuentaActivo Then
            Dim idCuentaSel As Integer = Convert.ToInt32(frmApuntesContables.CmbCuenta.SelectedValue)
            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
        End If

        ' 4. FILTRO DE FECHAS PARÁMETRIZADO AL FINAL DE LA SQL
        If filtroFechaActivo Then
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

        ' Foco automático seguro en la última fila del Grid
        If frmApuntesContables.DgvApuntes.RowCount > 0 Then
            Dim ultimaFila As Integer = frmApuntesContables.DgvApuntes.RowCount - 1
            frmApuntesContables.DgvApuntes.Rows(ultimaFila).Selected = True
            frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(ultimaFila).Cells(0)
        End If
    End Sub

    ''' <summary>
    ''' Muestra un cuadro de confirmación Sí/No adaptado al 100% al idioma del resManager
    ''' </summary>
    Public Function ConfirmarAccionTraducida(ByVal mensaje As String, ByVal titulo As String) As MsgBoxResult
        ' Forzamos la traducción de las palabras clave usando tu gestor de recursos
        Dim textoSi As String = If(resManager?.GetString("SI"), "Sí")
        Dim textoNo As String = If(resManager?.GetString("NO"), "No")

        ' Si estás en catalán y tus archivos .resx tienen "Sí" y "No", las heredará.
        ' Si no, puedes forzar un desvío rápido por código si detectas el idioma:
        ' If vIdiomaActivo = "ca" Then textoSi = "Sí" : textoNo = "No"

        ' Creamos un formulario temporal ligero sobre la marcha (Cero archivos de diseño .vb)
        Dim frm As New Form()
        Dim lbl As New Label()
        Dim btnSi As New Button()
        Dim btnNo As New Button()

        frm.Text = titulo
        lbl.Text = mensaje
        btnSi.Text = textoSi
        btnNo.Text = textoNo

        ' Configuramos las respuestas lógicas de .NET
        btnSi.DialogResult = DialogResult.Yes
        btnNo.DialogResult = DialogResult.No

        ' --- Estética rápida y limpia ---
        frm.Size = New Size(400, 160)
        frm.FormBorderStyle = FormBorderStyle.FixedDialog
        frm.MaximizeBox = False
        frm.MinimizeBox = False
        frm.StartPosition = FormStartPosition.CenterScreen

        lbl.SetBounds(20, 20, 350, 40)
        btnSi.SetBounds(180, 80, 90, 30)
        btnNo.SetBounds(280, 80, 90, 30)

        frm.Controls.AddRange(New Control() {lbl, btnSi, btnNo})
        frm.AcceptButton = btnSi
        frm.CancelButton = btnNo

        ' Mostramos la ventana y capturamos la respuesta del usuario
        Dim resultado As DialogResult = frm.ShowDialog()
        frm.Dispose()

        If resultado = DialogResult.Yes Then
            Return MsgBoxResult.Yes
        Else
            Return MsgBoxResult.No
        End If
    End Function

End Module
