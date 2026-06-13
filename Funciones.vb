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
                f.Text = String.Format("{0}  -  {1}: {2}  -  {3}: {4}",
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
                    MsgBox("No se conectó por: " & ex.Message)
                End Try
            End If
        End If
    End Sub

    Public Function IniciarSaldosIniciales(vAny As String) As Boolean
        vAñoEjercicio = vAny
        vConceptoAPU = "SALDO"

        ' Creamos una única estructura de conexión para todo el procedimiento
        Using conexion As New OleDbConnection(conexion1.ConnectionString) '[1]
            Try
                conexion.Open()

                ' =================================================================
                ' PASO 1: BORRADO (Usando la misma conexión limpia)
                ' =================================================================
                Dim sqlDelete As String = "DELETE FROM apuntes WHERE ConceptoAPU = ? And EjercicioAPU = ?"
                Using cmdDelete As New OleDbCommand(sqlDelete, conexion)
                    cmdDelete.Parameters.AddWithValue("@concepto", vConceptoAPU)
                    cmdDelete.Parameters.AddWithValue("@ejercicio", CInt(vAñoEjercicio))
                    cmdDelete.ExecuteNonQuery()
                End Using

                ' =================================================================
                ' PASO 2: CARGA DE DATOS HISTÓRICOS
                ' =================================================================
                Dim sqlSelect As String =
                "SELECT A.EjercicioAPU, A.CuentaAPU, SUM(A.ImporteAPU) AS SumaAño " &
                "FROM (Ejercicios AS E INNER JOIN Apuntes AS A ON E.EjercicioEJE = A.EjercicioAPU) " &
                "WHERE E.EjercicioEJE < ? AND A.ConceptoAPU <> 'SALDO' " &
                "GROUP BY A.EjercicioAPU, A.CuentaAPU " &
                "ORDER BY A.EjercicioAPU ASC"

                Dim dtMovimientos As New DataTable()
                Using cmdSelect As New OleDbCommand(sqlSelect, conexion)
                    cmdSelect.Parameters.AddWithValue("@AñoSeleccionado", vAñoEjercicio)
                    Using adaptador As New OleDbDataAdapter(cmdSelect)
                        adaptador.Fill(dtMovimientos)
                    End Using
                End Using

                ' =================================================================
                ' PASO 3: PROCESAMIENTO EN MEMORIA
                ' =================================================================
                Dim saldosAcumulados As New Dictionary(Of String, Decimal)()

                For Each fila As DataRow In dtMovimientos.Rows
                    Dim cuenta As String = fila("CuentaAPU").ToString()
                    Dim importeAño As Decimal = Convert.ToDecimal(fila("SumaAño"))

                    If saldosAcumulados.ContainsKey(cuenta) Then
                        saldosAcumulados(cuenta) += importeAño
                    Else
                        saldosAcumulados.Add(cuenta, importeAño)
                    End If
                Next

                ' Si no hay saldos, salimos de la función cerrando la conexión automáticamente [1]
                If saldosAcumulados.Count = 0 Then
                    vAviso = True
                    Return False
                Else
                    vAviso = False
                End If

                ' =================================================================
                ' PASO 4: INSERCIÓN DE LOS NUEVOS SALDOS INICIALES
                ' =================================================================
                Dim fechaSaldoInicial As New Date(CInt(vAñoEjercicio), 1, 1)
                Dim sqlInsert As String =
                "INSERT INTO Apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, CuentaAPU) " &
                "VALUES (?, ?, ?, ?, ?, ?)"

                ' Forzamos confirmación síncrona en el motor Access para evitar búferes retrasados
                Using comandoConfig As New OleDbCommand("SET FORCED COMMIT TRUE", conexion)
                    Try : comandoConfig.ExecuteNonQuery() : Catch : End Try
                End Using

                Using cmdInsert As New OleDbCommand(sqlInsert, conexion)
                    ' Parámetros configurados con tipos explícitos en orden estricto
                    cmdInsert.Parameters.Add("@Fecha", OleDbType.Date)
                    cmdInsert.Parameters.Add("@Concepto", OleDbType.VarWChar)
                    cmdInsert.Parameters.Add("@Descripcion", OleDbType.VarWChar)
                    cmdInsert.Parameters.Add("@Importe", OleDbType.Currency)
                    cmdInsert.Parameters.Add("@Ejercicio", OleDbType.Integer)
                    cmdInsert.Parameters.Add("@Cuenta", OleDbType.VarWChar)

                    ' Ejecutamos todo bajo una transacción atómica segura
                    Using transaccion As OleDbTransaction = conexion.BeginTransaction()
                        cmdInsert.Transaction = transaccion

                        For Each par In saldosAcumulados
                            Dim cuenta As String = par.Key
                            Dim saldoFinalPasado As Decimal = par.Value

                            If saldoFinalPasado <> 0 Then
                                cmdInsert.Parameters("@Fecha").Value = fechaSaldoInicial
                                cmdInsert.Parameters("@Concepto").Value = "SALDO"
                                cmdInsert.Parameters("@Descripcion").Value = "Saldo Inicial"
                                cmdInsert.Parameters("@Importe").Value = saldoFinalPasado
                                cmdInsert.Parameters("@Ejercicio").Value = CInt(vAñoEjercicio)
                                cmdInsert.Parameters("@Cuenta").Value = cuenta
                                cmdInsert.ExecuteNonQuery()
                            End If
                        Next

                        ' Confirmamos y volcamos inmediatamente los datos al archivo físico (.mdb)
                        transaccion.Commit()
                    End Using
                End Using

                ' El proceso se completó de forma totalmente síncrona y real
                Return True

            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorInsertarSaldos") & " " & ex.Message,
                            resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Using ' [1] <-- Aquí se cierra de golpe cualquier rastro de la conexión, liberando el archivo físico
    End Function

    Public Sub LlenarGrid(ByRef tipoSql As String, tipoGrid As String, tipoopc As String)
        linSql = tipoSql.ToString
        vgrid = tipoGrid.ToString
        opcion = tipoopc
        If vgrid = "APUNTES_CONTABLES" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha1", OleDbType.Date)).Value = vDate1
                adp.SelectCommand.Parameters.Add(New OleDbParameter("@fecha2", OleDbType.Date)).Value = vDate2
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
            End With
            If frmApuntesContables.DgvApuntes.ColumnCount > 0 Then
                frmApuntesContables.DgvApuntes.Columns(frmApuntesContables.DgvApuntes.ColumnCount - 1).Visible = False
            End If
            'Llama a la función
            DgvApuntesContables(3, 4)

            ' Para insertar alguna columna
            'Dim columna As New DataGridViewTextBoxColumn With {
            '.HeaderText = "Notas",
            '.Width = 350
            '}
            'frmApuntesContables.DgvApuntes.Columns.Insert(5, columna)

        ElseIf vgrid = "PRINT_APUNTES_CONTABLES" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla

        ElseIf vgrid = "PRINT_INFORME_APUNTES" Then
            Using adp As New OleDbDataAdapter(linSql, conexion1)
                adp.SelectCommand.Parameters.Add(New OleDbParameter("fecha", OleDbType.Date)).Value = vDate1
                adp.SelectCommand.Parameters.Add(New OleDbParameter("fecha", OleDbType.Date)).Value = vDate2
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
                adp.SelectCommand.Parameters.Add(New OleDbParameter("fecha", OleDbType.Date)).Value = vDate1
                adp.SelectCommand.Parameters.Add(New OleDbParameter("fecha", OleDbType.Date)).Value = vDate2
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

            ' --- SOLUCIÓN: Recorremos las filas de la TABLA de datos antes de enlazar al Grid ---
            For Each filaData As DataRow In Tabla.Rows
                vNombreCuenta = filaData(1).ToString() ' Celda 1 (Nombre)

                ' Buscar el Saldo de cada Cuenta Bancaria en Apuntes
                cmdMdb1cr.CommandText = "SELECT apuntes.ImporteAPU FROM apuntes"
                cmdMdb1cr.CommandText += " WHERE apuntes.CuentaAPU = '" & vNombreCuenta & "' "
                cmdMdb1cr.CommandText += "And apuntes.EjercicioAPU = " & vAñoEjercicio.ToString

                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    vSaldoCuentas = 0
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            ' Sumamos de forma limpia convirtiendo a Decimal
                            vSaldoCuentas += Convert.ToDecimal(drMdb1.GetValue(0))
                        End While
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                    MsgBox(resManager.GetString("ErrorAlEjecutar") & ": " & cmdMdb1cr.CommandText & ex.Message)
                End Try

                ' Guardamos el número exacto directamente en el registro de la tabla
                ' Al calcularlo aquí, .NET sabrá que es un número real
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
                .Columns(0).Width = 135
                .Columns(0).HeaderText = resManager.GetString("Tipo")
                .Columns(1).Width = 200
                .Columns(1).HeaderText = resManager.GetString("Nombre")
                .Columns(2).Width = 200
                .Columns(2).HeaderText = resManager.GetString("Numero")
                .Columns(3).Width = 125
                .Columns(3).HeaderText = resManager.GetString("Saldo") & "(" & vMoneda & ")"

                .Columns(4).AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .Columns(4).HeaderText = resManager.GetString("Notas")

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
                    MsgBox("Error al ejecutar: " & cmdMdb1cr.CommandText & " por: " & ex.Message)
                End Try
                fila.Cells(3).Value = Math.Round(Convert.ToDecimal(vSaldoCuentas), 2)
                vValor += vSaldoCuentas
            Next
            frmImprimirForm.LblTotal.Text = String.Format("{0}: {1} {2}", resManager.GetString("TOTAL"), vValor.ToString("N2"), vMoneda)

        ElseIf vgrid = "PRINT_CUENTAS_PERIODICAS" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla


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
            MsgBox("Solo admite el . Punto como separador decimal",
            MsgBoxStyle.Exclamation, "Separador decimal")
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
            MsgBox("Error N° " & Err.Number & NL & ex.Message, MsgBoxStyle.Critical, "Información")
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
                        MsgBox("Versión Instalada: " & My.Settings.Version & vbNewLine & "Versión Disponible: " & vNewVersion, MsgBoxStyle.Information, "Comprobar Nueva Versión")
                        vNuevaVersion = vNewVersion
                        vHayNuevaVersion = "SI"
                        Dim respuesta As MsgBoxResult = MsgBox("¿Quieres actualizar a la Versión: " & vNewVersion & " ?", vbQuestion + vbYesNo + vbDefaultButton1, "Versión ContaHogar 3.0")
                        If respuesta = vbYes Then
                            Dim respuesta2 As MsgBoxResult = MsgBox("Quieres guardar una Copia de Seguridad de la Base de Datos.", vbQuestion + vbYesNo + vbDefaultButton1, "Actualizar Software")
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
                                        MessageBox.Show("Backup realizado satisfactoriamente. Ahora, se descargará la actualización.", "BACKUP", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                    Catch ex As Exception
                                        MsgBox("Error al realizar el Backup de la Base de Datos, revise que no exista otro Backup con el mismo nombre o que el archivo no esté abierto.")
                                        MsgBox(ex.ToString)
                                    End Try
                                End If
                            End If
                            MessageBox.Show("Ahora, se descargará la actualización.", "Actualizar Software", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            Try
                                Dim descargar As New Devices.Computer
                                With descargar
                                    .Network.DownloadFile("https://filedn.eu/ljfTvwyEW2tVj4PWYI9927f/ContaHogar/Actualizar/" & vNewVersion & "/InstaladorContaHogar3.0.msi", "C:\ContaHogar3.0\InstaladorContaHogar3.0.msi", "", "", True, 100, True, 3)
                                End With
                                Dim Proceso As New Process()
                                MsgBox("Se ha descargado la actualización: " & vNewVersion & ", se procederá a instalarla. Pulsa Aceptar para continuar.", MsgBoxStyle.Information, "Actualizar Software")
                                Proceso.StartInfo.FileName = "C:\ContaHogar3.0\InstaladorContaHogar3.0.msi"
                                Proceso.StartInfo.Arguments = ""
                                Proceso.Start()
                            Catch ex As Exception
                                MsgBox("Error al descargar la actualización: " & vNewVersion & ", revise su conexión a Internet o que el archivo no esté abierto.")
                                MsgBox(ex.ToString)
                            End Try
                            Application.Exit()
                        End If
                    Else
                        vHayNuevaVersion = "NO"
                        'MsgBox(My.Settings.Version & " = " & vNewVersion)
                    End If
                Catch ex As Exception
                    MsgBox("Error al comprobar la nueva versión, revise su conexión a Internet.")
                    MsgBox(ex.ToString)
                End Try
            Else
                MsgBox("No estas conectado a una red para comprobar nueva Versión.", MsgBoxStyle.Information, "Sin Conexión a Internet")
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
    ''' <param name="manejadorRecursos">El objeto rmse propio del formulario</param>
    Public Sub TraducirColumnasGridCuentas(ByVal grid As DataGridView, ByVal manejadorRecursos As System.Resources.ResourceManager)
        Try
            If grid IsNot Nothing AndAlso grid.Rows.Count > 0 Then

                For Each fila As DataGridViewRow In grid.Rows
                    If Not fila.IsNewRow Then

                        ' --- COLUMNA (0): TipoCUE (Mixto) ---
                        If grid.Columns.Count > 0 AndAlso fila.Cells(0).Value IsNot Nothing Then
                            Dim valorTipo As String = fila.Cells(0).Value.ToString().Trim()
                            ' Usamos el parámetro interno, que será el 'rmse' que envíes
                            Dim tradTipo As String = manejadorRecursos.GetString(valorTipo)

                            If Not String.IsNullOrEmpty(tradTipo) Then
                                fila.Cells(0).Value = tradTipo
                            End If
                        End If

                        ' --- COLUMNA (1): NombreCUE (Mayúsculas) ---
                        If grid.Columns.Count > 1 AndAlso fila.Cells(1).Value IsNot Nothing Then
                            Dim valorNombre As String = fila.Cells(1).Value.ToString().Trim().ToUpper()
                            ' Usamos el parámetro interno
                            Dim tradNombre As String = manejadorRecursos.GetString(valorNombre)

                            If Not String.IsNullOrEmpty(tradNombre) Then
                                fila.Cells(1).Value = tradNombre
                            End If
                        End If

                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorAlEjecutar") & ex.Message, MsgBoxStyle.Exclamation, manejadorRecursos.GetString("$this.Text"))
        End Try
    End Sub

    ''' <summary>
    ''' Rellena de forma híbrida y multidioma cualquier ComboBox con los tipos de cuenta desde Access
    ''' </summary>
    ''' <param name="combo">El control ComboBox que se quiere rellenar</param>
    ''' <param name="rm">El administrador de recursos (resManager o rmse) del formulario que llama</param>
    Public Sub CargarComboTipoCuentaGlobal(ByVal combo As ComboBox, ByVal rm As System.ComponentModel.ComponentResourceManager)
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

                    textoTraducido = rm.GetString(valorBD)
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

                    historialSeguimiento &= $"BD: {valorBD} -> Trad: {textoTraducido}" & vbNewLine
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
    ''' <param name="manejadorRecursos">El objeto ResourceManager (rmse) propio del formulario</param>
    Public Sub TraducirContenidoGridTiposCuenta(ByVal grid As DataGridView, ByVal manejadorRecursos As System.Resources.ResourceManager)
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
                        Dim tradTipo As String = manejadorRecursos.GetString(llaveBase)
                        If Not String.IsNullOrEmpty(tradTipo) Then
                            fila.Cells(0).Value = tradTipo
                        End If

                        ' --- TRADUCIR COLUMNA (1): Descripción del Tipo ---
                        ' Buscamos usando el prefijo "Desc_" combinado con la llave del tipo
                        Dim llaveDesc As String = "Desc_" & llaveBase
                        Dim tradDesc As String = manejadorRecursos.GetString(llaveDesc)

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
            ' Muestra el mensaje de error usando los recursos para internacionalizar el aviso
            Dim tituloError As String = If(manejadorRecursos.GetString("$this.Text"), "Error")
            Dim msgError As String = "Error: " ' Valor por defecto por si falla resManager global
            Try
                msgError = resManager.GetString("ErrorAlEjecutar")
            Catch
            End Try
            MsgBox(msgError & " " & ex.Message, MsgBoxStyle.Exclamation, tituloError)
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
                    vAñadir = "INSERT INTO tempapu"
                    vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                    vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
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
                    ' Convertimos el primer importe (vImporteConcepto)
                    If vImporteConcepto IsNot Nothing Then
                        Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            importeConcepto)
                    End If
                    ' Convertimos el segundo importe (vExistenteImporteConcepto)
                    If vExistenteImporteConcepto IsNot Nothing Then
                        Decimal.TryParse(vExistenteImporteConcepto.ToString().Replace(",", "."),
                            System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture,
                            existenteImporte)
                    End If
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
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                            MsgBox(ex.ToString)
                        End Try
                        vAñadir = "INSERT INTO tempapu"
                        vAñadir += "(ConceptoAPU, SumaImporteAPU) "
                        vAñadir += "VALUES ('" & vNombreConcepto & "',' 0 ')"
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
                        If vImporteConcepto IsNot Nothing Then
                            Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."),
                     System.Globalization.NumberStyles.Any,
                     System.Globalization.CultureInfo.InvariantCulture,
                     importeDecimal)
                        End If

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
                                'vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                ' 1. Convertimos ambos importes a variables decimales exactas
                                Dim importe1 As Decimal = 0.0D
                                Dim importe2 As Decimal = 0.0D

                                ' Conversión segura del primer importe
                                If vImporteConcepto IsNot Nothing Then
                                    Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe1)
                                End If

                                ' Conversión segura del segundo importe
                                If vExistenteImporteConcepto IsNot Nothing Then
                                    Decimal.TryParse(vExistenteImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe2)
                                End If

                                ' 2. Sumamos los números reales de forma exacta
                                vNewImporteConcepto = importe1 + importe2

                                If importe1 > 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU > 0 "
                                ElseIf importe1 < 0 Then
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU < 0 "
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
                                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto).ToString
                                    ' 1. Convertimos ambos importes a variables decimales exactas
                                    Dim importe1 As Decimal = 0.0D
                                    Dim importe2 As Decimal = 0.0D

                                    ' Conversión segura del primer importe
                                    If vImporteConcepto IsNot Nothing Then
                                        Decimal.TryParse(vImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe1)
                                    End If

                                    ' Conversión segura del segundo importe
                                    If vExistenteImporteConcepto IsNot Nothing Then
                                        Decimal.TryParse(vExistenteImporteConcepto.ToString().Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, importe2)
                                    End If

                                    ' 2. Sumamos los números reales de forma exacta
                                    vNewImporteConcepto = importe1 + importe2

                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' "
                                    vAñadir2 += " WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                    vAñadir2 += "And tempapu.SumaImporteAPU = 0 "
                                    cmdMdb1cr.CommandText = vAñadir2
                                    Try
                                        cmdMdb1cr.ExecuteNonQuery()
                                    Catch ex As Exception
                                        MsgBox(resManager.GetString("ErrorGrabarTemporal"))
                                        MsgBox(ex.ToString)
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
                    Dim fechaFormatoAccess As String = "#" & fechaFila.ToString("yyyy-MM-dd") & "#"

                    If vFechaConcepto <> fechaFila Then
                        vFechaConcepto = fechaFila
                        vImporteConcepto = importeFila

                        ' Primer INSERT
                        vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                              "VALUES (" & fechaFormatoAccess & ", '', '', '', '', '" & vImporteConcepto & "', '0')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                        End Try

                        ' Segundo INSERT
                        vAñadir = "INSERT INTO tmpprint (FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) " &
                              "VALUES (" & fechaFormatoAccess & ", '', '', '', '', '0', '0')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try
                            cmdMdb1cr.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                        End Try

                    Else ' Si la fecha ya existe, sumamos o acumulamos el importe
                        cmdMdb1cr.CommandType = CommandType.Text

                        ' Construimos el SELECT filtrando por signo
                        cmdMdb1cr.CommandText = "SELECT ImporteTMP FROM tmpprint WHERE FechaTMP = " & fechaFormatoAccess
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
                                vAñadir2 = "UPDATE tmpprint SET ImporteTMP = '" & vNewImporteConcepto & "' WHERE FechaTMP = " & fechaFormatoAccess
                                If vImporteConcepto > 0 Then
                                    vAñadir2 += " AND ImporteTMP > 0"
                                Else
                                    vAñadir2 += " AND ImporteTMP < 0"
                                End If
                                cmdMdb1cr.CommandText = vAñadir2
                                Try
                                    cmdMdb1cr.ExecuteNonQuery()
                                Catch ex As Exception
                                    MsgBox(resManager.GetString("ErrorGrabarTemporal") & vbCrLf & ex.Message)
                                End Try
                            Else ' NO existe registro con ese signo, buscamos el que tiene importe = 0
                                cmdMdb1cr.CommandText = "SELECT ImporteTMP FROM tmpprint WHERE FechaTMP = " & fechaFormatoAccess & " AND ImporteTMP = 0"

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

                                    vAñadir2 = "UPDATE tmpprint SET ImporteTMP = '" & vNewImporteConcepto & "' WHERE FechaTMP = " & fechaFormatoAccess & " AND ImporteTMP = 0"
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

                If fila.Cells(3).Value <> 0 Then
                    vImporteConcepto = fila.Cells(3).Value

                    ' Extraemos de forma genérica el Año-Mes de la celda de fecha
                    Dim fechaReal As DateTime = Convert.ToDateTime(fila.Cells(0).Value)
                    Dim claveMesAño As String = fechaReal.ToString("yy") & "-" & fechaReal.Month.ToString("D2")
                    ' "D2" fuerza a que el mes salga como "01" en lugar de "1" para mantener el orden en la base de datos

                    If vNombreConcepto <> claveMesAño Then
                        vNombreConcepto = claveMesAño
                        vImporteConcepto = fila.Cells(3).Value

                        ' Inserción del importe real
                        vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES ('" & vNombreConcepto & "','" & vImporteConcepto & "')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : MsgBox(ex.ToString) : End Try

                        ' Inserción de la fila espejo a cero
                        vAñadir = "INSERT INTO tempapu(ConceptoAPU, SumaImporteAPU) VALUES ('" & vNombreConcepto & "',' 0 ')"
                        cmdMdb1cr.CommandText = vAñadir
                        Try : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : MsgBox(ex.ToString) : End Try
                    Else
                        ' Si ya existe el registro del mes actual en tempapu, actualizamos acumulando el importe
                        cmdMdb1cr.CommandType = CommandType.Text
                        If Val(vImporteConcepto) > 0 Then
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' And tempapu.SumaImporteAPU > 0 "
                        ElseIf Val(vImporteConcepto) < 0 Then
                            cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' And tempapu.SumaImporteAPU < 0 "
                        End If

                        Try
                            drMdb1 = cmdMdb1cr.ExecuteReader()
                            If drMdb1.HasRows Then
                                While drMdb1.Read() : vExistenteImporteConcepto = drMdb1.GetValue(1) : End While
                                drMdb1.Close() ' Importante cerrar el reader antes del Update

                                vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto)
                                vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' "
                                vAñadir2 += If(Val(vImporteConcepto) > 0, "And tempapu.SumaImporteAPU > 0 ", "And tempapu.SumaImporteAPU < 0 ")

                                cmdMdb1cr.CommandText = vAñadir2
                                Try : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : MsgBox(ex.ToString) : End Try
                            Else
                                drMdb1.Close()
                                ' Si no existe, acumulamos sobre el registro que se creó a cero
                                cmdMdb1cr.CommandText = "SELECT * FROM tempapu WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' And tempapu.SumaImporteAPU = 0 "
                                drMdb1 = cmdMdb1cr.ExecuteReader()
                                If drMdb1.HasRows Then
                                    While drMdb1.Read() : vExistenteImporteConcepto = drMdb1.GetValue(1) : End While
                                    drMdb1.Close()

                                    vNewImporteConcepto = Val(vImporteConcepto) + Val(vExistenteImporteConcepto)
                                    vAñadir2 = "UPDATE tempapu SET SumaImporteAPU = '" & vNewImporteConcepto & "' WHERE tempapu.ConceptoAPU = '" & vNombreConcepto & "' And tempapu.SumaImporteAPU = 0 "
                                    cmdMdb1cr.CommandText = vAñadir2
                                    Try : cmdMdb1cr.ExecuteNonQuery() : Catch ex As Exception : MsgBox(ex.ToString) : End Try
                                End If
                                drMdb1.Close()
                            End If
                        Catch ex As Exception
                            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
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

End Module
