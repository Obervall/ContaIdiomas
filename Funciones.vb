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
    Public frmGraficosConceptos3D As New GraficosConceptos3D
    Public frmGraficosCuentas As New GraficosCuentas
    Public frmGraficosCuentas3D As New GraficosCuentas3D
    Public frmGraficosFechas As New GraficosFechas
    Public frmGraficosFechas3D As New GraficosFechas3D
    Public frmGraficosMeses As New GraficosMeses
    Public frmGraficosMeses3D As New GraficosMeses3D
    Public frmGraficosPresupuestos As New GraficosPresupuestos
    Public frmSeleccionarDatosIngresos As New SeleccionDatosIngresos
    Public frmSeleccionarDatosGastos As New SeleccionDatosGastos
    Public frmGraficosSoloConceptos As New GraficosSoloConceptos
    Public frmGraficosSoloConceptos3D As New GraficosSoloConceptos3D
    Public frmActivarSoftware As New ActivarSoftware
    Public frmAportacionBizum As New AportacionBizum

    Public backup As New SaveFileDialog
    Public restore As New OpenFileDialog

    Public conexion1 As New OleDbConnection()
    Public cmdMdb1cr As New OleDbCommand
    Public drMdb1 As OleDbDataReader

    Public vgrid, linSql, opcion, vTipoEstados, vNombreCuenta, vNombreConcepto, vFecha, vFechaMes As String
    Public vtipoSql, vAñadirSql, vtipoGrid, vMes, vEditar, respuesta, vBuscar, vTituloInforme, vtipoSqlChk As String
    Public vValor, vIngresos, vGastos, vSaldo, vSaldoCuentas, vSaldoMes, vSaldoAnualReal As Double
    Public i, vFila1, vFila2, vFila, vFilaActual, filaActual, vregData1, vAñoActual, vAñoEjercicio As Integer
    Public vCerrar, vGrafico, vLetras, vNumeros, vNotas, vPathExportar, vConcepto As String
    Public vDescripcionAPU, vImporteAPU, vNotasAPU, vConceptoAPU As String
    Public vActualizar, vActivado As Boolean

    Public vOrdenadoPorFechasAPU, vOrdenadoPorConceptosAPU, vOrdenadoPorImportesAPU As Integer
    Public vSoloIngresosAPU, vSoloGastosAPU, centroX, AnchoFrmPrincipal, posX, posY As Integer
    Public vOrdenadoPorFechasAPP, vOrdenadoPorConceptosAPP, vOrdenadoPorImportesAPP As Integer
    Public vSoloIngresosAPP, vSoloGastosAPP, vFecha1Enero, vFecha31Diciembre, vCantAños As Integer
    Public vRuta, vVersion, vNewVersion, vHayNuevaVersion, vNuevaVersion, vMoneda As String
    Public resManager As New ResourceManager("Contahogar.Recursos", Assembly.GetExecutingAssembly())
    ' Para copiar en en el Classe de cada Form la línea:
    'Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())
    Public vDate1, vDate2, vDate3, vFechaTemp, vFechaTemp2 As DateTime
    Public vfechaHoy As DateTime = DateTime.Today
    'Public resManager As ResourceManager

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
        ' 1. Traducir el Título del Formulario
        Dim titol As String = resManager.GetString("$this")
        If titol IsNot Nothing Then
            f.Text = titol
        End If
        ' 2. Traducir los controles (Labels, Buttons, etc.)
        ' Este bucle busca en el .resx una llave que se llame igual que el NAME del control
        For Each ctrl As Control In f.Controls
            Dim texto As String = resManager.GetString(ctrl.Name)
            If texto IsNot Nothing Then
                ctrl.Text = texto
            End If
            ' Si usas Panels o GroupBox, hay que mirar dentro de ellos también:
            If ctrl.HasChildren Then
                For Each child As Control In ctrl.Controls
                    Dim textoChild As String = resManager.GetString(child.Name)
                    If textoChild IsNot Nothing Then
                        child.Text = textoChild
                    End If
                Next
            End If
        Next
        ' 3. TRADUCCIÓ DINÀMICA DEL TÍTOL (Aquí és on va la línia!)
        ' Comprovem si és el formulari principal per aplicar el format especial
        If f.Name = "Principal" Then ' Substitueix "FormPrincipal" pel nom real del teu formulari
            ' Busquem les paraules traduïdes dins del fitxer de recursos
            ' NOTA: "recursos" aquí és el ComponentResourceManager que ja has creat a dalt
            Dim txtTitol As String = resManager.GetString("TitolApp") ' My.Resources.Recursos.TitolApp
            Dim txtVersio As String = resManager.GetString("Versio") ' My.Resources.Recursos.Versio
            Dim txtExercici As String = resManager.GetString("Exercici") ' My.Resources.Recursos.Exercici
            'MsgBox("Títol: " & txtTitol & vbNewLine & "Versió: " & txtVersio & vbNewLine & "Exercici: " & txtExercici, MsgBoxStyle.Information, "Comprovar Traducció Títol Principal")
            f.Text = String.Format("{0}  -  {1}: {2}  -  {3}: {4}",
                                            txtTitol,
                                            txtVersio,
                                            My.Settings.Version,
                                            txtExercici,
                                            vAñoEjercicio.ToString())
        End If
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

    Public Sub IniciarSaldosIniciales(vAny As String)
        vAñoEjercicio = vAny
        'Quitamos el Concepto SALDO del Ejercicio marcado
        vConceptoAPU = "SALDO"
        vtipoSql = "DELETE FROM apuntes"
        vtipoSql += " WHERE apuntes.ConceptoAPU = '" & vConceptoAPU & "' "
        cmdMdb1cr.CommandText = vtipoSql
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Registros SALDO, Borrados !!!")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' 1. Consulta SQL: Agrupamos y sumamos SOLO los años ESTRICTAMENTE MENORES al seleccionado
        Dim consulta As String =
        "SELECT A.EjercicioAPU, A.CuentaAPU, SUM(A.ImporteAPU) AS SumaAño " &
        "FROM (Ejercicios AS E INNER JOIN Apuntes AS A ON E.EjercicioEJE = A.EjercicioAPU) " &
        "WHERE E.EjercicioEJE < ? " &
        "GROUP BY A.EjercicioAPU, A.CuentaAPU " &
        "ORDER BY A.EjercicioAPU ASC"

        Dim dtMovimientos As New DataTable()

        ' 2. Carga de datos históricos en memoria
        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using comando As New OleDbCommand(consulta, conexion)
                comando.Parameters.AddWithValue("@AñoSeleccionado", vAñoEjercicio)
                Using adaptador As New OleDbDataAdapter(comando)
                    Try
                        conexion.Open()
                        adaptador.Fill(dtMovimientos)
                        'MsgBox("Datos históricos cargados en memoria: " & dtMovimientos.Rows.Count.ToString & " registros encontrados.")
                    Catch ex As Exception
                        MessageBox.Show("Error al leer históricos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                End Using
            End Using
        End Using

        ' 3. Procesamiento en memoria para obtener el saldo final acumulado por cuenta
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

        ' Si no hay saldos que arrastrar del pasado, salimos del proceso de inserción
        If saldosAcumulados.Count = 0 Then
            MessageBox.Show("No hay datos históricos en años anteriores para generar saldos iniciales.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        ' 4. Inserción de los Saldos Iniciales en la tabla Apuntes
        ' Definimos la fecha de inserción fija: 01/01/vAñoEjercicio
        Dim fechaSaldoInicial As New Date(vAñoEjercicio, 1, 1)

        Dim consultaInsert As String =
        "INSERT INTO Apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, CuentaAPU) " &
        "VALUES (?, ?, ?, ?, ?, ?)"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using comandoInsert As New OleDbCommand(consultaInsert, conexion)

                ' Declaramos los parámetros en el orden exacto de los signos de interrogación '?'
                comandoInsert.Parameters.Add("@Fecha", OleDbType.Date)
                comandoInsert.Parameters.Add("@Concepto", OleDbType.VarWChar)
                comandoInsert.Parameters.Add("@Descripcion", OleDbType.VarWChar)
                comandoInsert.Parameters.Add("@Importe", OleDbType.Currency)
                comandoInsert.Parameters.Add("@Ejercicio", OleDbType.Integer)
                comandoInsert.Parameters.Add("@Cuenta", OleDbType.VarWChar)

                Try
                    conexion.Open()

                    ' 1. CONFIGURACIÓN CRÍTICA: Desactivar la escritura diferida en la caché de Access
                    ' El valor 1 obliga a que cada Commit escriba inmediatamente en el disco duro (.mdb)
                    Using comandoConfig As New OleDbCommand("SET FORCED COMMIT TRUE", conexion)
                        Try
                            comandoConfig.ExecuteNonQuery()
                        Catch
                            ' Si el motor de Access rechaza la instrucción SQL directa, configuramos la propiedad ADO.NET:
                            Try
                                conexion.GetType().GetProperty("Providers")?.SetValue(conexion, "Jet OLEDB:Transaction Commit Mode=1")
                            Catch
                                ' Si falla, pasamos al plan B automático del paso 3
                            End Try
                        End Try
                    End Using

                    ' Usamos una transacción para asegurar que se guarden todos los saldos o ninguno
                    Using transaccion As OleDbTransaction = conexion.BeginTransaction()
                        comandoInsert.Transaction = transaccion

                        ' Recorremos cada cuenta calculada e inyectamos su saldo acumulado
                        For Each par In saldosAcumulados
                            Dim cuenta As String = par.Key
                            Dim saldoFinalPasado As Decimal = par.Value

                            ' Omitimos cuentas con saldo acumulado cero si no deseas registros vacíos
                            If saldoFinalPasado <> 0 Then
                                comandoInsert.Parameters("@Fecha").Value = fechaSaldoInicial
                                comandoInsert.Parameters("@Concepto").Value = "SALDO"
                                comandoInsert.Parameters("@Descripcion").Value = "Saldo Inicial"
                                comandoInsert.Parameters("@Importe").Value = saldoFinalPasado
                                comandoInsert.Parameters("@Ejercicio").Value = vAñoEjercicio
                                comandoInsert.Parameters("@Cuenta").Value = cuenta ' Guarda el identificador/nombre de la cuenta
                                comandoInsert.ExecuteNonQuery()
                            End If
                        Next

                        ' Confirmamos los cambios en la base de datos
                        transaccion.Commit()
                        'MessageBox.Show("Saldos iniciales del año " & vAñoEjercicio & " generados e insertados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Using

                    ' 3. PLAN B AUTOMÁTICO (En caso de que el sistema siga reteniendo hilos de caché):
                    ' Hacemos una micropausa de código limpia (no visual) para asegurar el vaciado del búfer del SO.
                    System.Threading.Thread.Sleep(500)
                Catch ex As Exception
                    MessageBox.Show("Error al insertar los saldos iniciales: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using
    End Sub

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
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'Liberamos las columnas del auto-ajuste estricto para permitir el Scroll manual posterior
                For Each columna As DataGridViewColumn In .Columns
                    columna.AutoSizeMode = DataGridViewAutoSizeColumnMode.None
                Next

                ' arreglamos columnas
                '********************
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(4).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(5).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(3).DefaultCellStyle.Format = "###,##0.00"
                .Columns(4).DefaultCellStyle.Format = "###,##0.00"
                .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.Format = "dd/MM/yyyy"
                .Columns(0).Width = 100
                .Columns(0).HeaderText = "Fecha"
                .Columns(1).Width = 200
                .Columns(1).HeaderText = "Concepto"
                .Columns(2).Width = 200
                .Columns(2).HeaderText = "Descripción"
                .Columns(3).Width = 100
                .Columns(3).HeaderText = "Importe(" & vMoneda & ")"
                .Columns(4).Width = 90
                .Columns(4).HeaderText = "Saldo(" & vMoneda & ")"
                .Columns(5).Width = 140
                .Columns(5).HeaderText = "Notas"
                .Columns(6).Width = 140
                .Columns(6).HeaderText = "Cuenta"
                .Columns(7).Width = 0
                .Columns(7).HeaderText = "Código"
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
            frmImprimirForm.LblTotal.Text = "Total: 0,00 " & vMoneda
            For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                vValor += fila.Cells(3).Value
                frmImprimirForm.LblTotal.Text = "Total:  " & Format(vValor, "###,##0.00 ").ToString & vMoneda
            Next

        ElseIf vgrid = "PRINT_TEMP_APUNTES" Then
            Dim adp As New OleDbDataAdapter(linSql, conexion1)
            Dim Tabla As New DataTable
            adp.Fill(Tabla)
            frmImprimirForm.DgvApuntes.DataSource = ""
            frmImprimirForm.DgvApuntes.DataSource = Tabla

        ElseIf vgrid = "PRINT_TEMP_APUNTES_FECHAS" Then
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
                .Columns(3).DefaultCellStyle.Format = "###,##0.00"
                .Columns(4).DefaultCellStyle.Format = "###,##0.00"
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
                .Columns(0).Width = 95
                .Columns(0).HeaderText = resManager.GetString("Tipo") ' My.Resources.Recursos.Tipo
                .Columns(1).Width = 150
                .Columns(1).HeaderText = resManager.GetString("Codigo") ' My.Resources.Recursos.Codigo
                .Columns(2).Width = 150
                .Columns(2).HeaderText = resManager.GetString("Descripcion") ' My.Resources.Recursos.Descripcion
                .Columns(3).Width = 180
                .Columns(3).HeaderText = resManager.GetString("Notas") ' My.Resources.Recursos.Notas
                Dim vNumRegistros As String = frmConceptosContables.DgvConceptos.Rows.Count.ToString
                frmConceptosContables.TxtNumRegistros.Text = vNumRegistros
                If frmConceptosContables.BtnFiltroTipoConcepto.Enabled = False Then
                    frmConceptosContables.LblNumRegistros.Text = resManager.GetString("Filtrado") ' My.Resources.Recursos.Filtrado
                Else
                    frmConceptosContables.LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
                End If
                For Each fila As DataGridViewRow In frmConceptosContables.DgvConceptos.Rows
                    If fila.Cells(0).Value = "GASTO" Then
                        fila.Cells(0).Style.ForeColor = Color.DarkRed
                    ElseIf fila.Cells(0).Value = "INGRESO" Then
                        fila.Cells(0).Style.ForeColor = Color.DarkBlue
                    ElseIf fila.Cells(0).Value = "ESPECIAL" Then
                        fila.Cells(0).Style.ForeColor = Color.DarkGreen
                    End If
                Next
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
            frmCuentasBancarias.DgvCuentas.DataSource = ""
            frmCuentasBancarias.DgvCuentas.DataSource = Tabla
            With frmCuentasBancarias.DgvCuentas
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue
                ' arreglamos columnas
                '********************
                .Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkGreen
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                .Columns(3).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(0).Width = 135
                .Columns(0).HeaderText = resManager.GetString("Tipo") ' My.Resources.Recursos.Tipo
                .Columns(1).Width = 135
                .Columns(1).HeaderText = resManager.GetString("Nombre") ' My.Resources.Recursos.Nombre
                .Columns(2).Width = 200
                .Columns(2).HeaderText = resManager.GetString("Numero") ' My.Resources.Recursos.Numero
                .Columns(3).Width = 90
                .Columns(3).HeaderText = resManager.GetString("Saldo") & "(" & vMoneda & ")" ' My.Resources.Recursos.Saldo
                .Columns(4).Width = 155
                .Columns(4).HeaderText = resManager.GetString("Notas") ' My.Resources.Recursos.Notas

                '' Para insertar alguna columna
                'Dim columna As New DataGridViewTextBoxColumn With {
                '.HeaderText = "Saldo(" & vMoneda & ")",
                '.Width = 350
                '}
                '.Columns.Insert(3, columna)

                For Each fila As DataGridViewRow In frmCuentasBancarias.DgvCuentas.Rows
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
                        MsgBox(resManager.GetString("ErrorAlEjecutar") & ": " & cmdMdb1cr.CommandText & ex.Message)
                    End Try
                    fila.Cells(3).Value = Format(vSaldoCuentas, "###,##0.00")
                Next

                Dim vNumRegistros As String = frmCuentasBancarias.DgvCuentas.Rows.Count.ToString
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
            vValor = 0
            frmImprimirForm.LblTotal.Text = "Total: 0,00 " & vMoneda
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
                fila.Cells(3).Value = Format(vSaldoCuentas, "###,##0.00")
                vValor += vSaldoCuentas
            Next
            frmImprimirForm.LblTotal.Text = "Total: " & Format(vValor, "###,##0.00 ").ToString & vMoneda

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
            frmPresupuestos.DgvPresupuestos.DataSource = Tabla
            With frmPresupuestos.DgvPresupuestos
                .DefaultCellStyle.Font = New Font("Tahoma", 9)
                .DefaultCellStyle.ForeColor = Color.Black
                .DefaultCellStyle.BackColor = Color.White
                .DefaultCellStyle.SelectionForeColor = Color.White
                .DefaultCellStyle.SelectionBackColor = Color.Blue
                ' arreglamos columnas
                '********************
                .Columns(0).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(1).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.ForeColor = Color.DarkBlue
                .Columns(2).DefaultCellStyle.Format = "###,##0.00"
                .Columns(3).DefaultCellStyle.Format = "###,##0.00"
                .Columns(0).Width = 175
                .Columns(0).HeaderText = "Concepto"
                .Columns(1).Width = 100
                .Columns(1).HeaderText = "Mes"
                .Columns(2).Width = 97
                .Columns(2).HeaderText = "Real"
                .Columns(3).Width = 97
                .Columns(3).HeaderText = "Presupuesto"
                .Columns(4).Width = 0
                .Columns(4).HeaderText = "Fecha"
                Dim vNumRegistros As String = frmPresupuestos.DgvPresupuestos.Rows.Count.ToString
                frmPresupuestos.TxtNumRegistros.Text = vNumRegistros
                If frmPresupuestos.BtnFiltroConcepto.Enabled = False Then
                    frmPresupuestos.LblNumRegistros.Text = resManager.GetString("Filtrado") ' My.Resources.Recursos.Filtrado
                Else
                    frmPresupuestos.LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
                End If
                For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
                    vFecha = fila.Cells(4).Value.ToString
                    vMes = Mid(vFecha, 4, 2).ToString
                    ' Con TRUE retorna el mes abreviado.
                    fila.Cells(1).Value = MonthName(vMes, False)
                    vNombreConcepto = fila.Cells(0).Value
                    ' Buscar el Saldo Total de cada Concepto por meses
                    '*************************************************
                    cmdMdb1cr.CommandText = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.ImporteAPU FROM apuntes"
                    cmdMdb1cr.CommandText += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    cmdMdb1cr.CommandText += "And apuntes.ConceptoAPU = '" & vNombreConcepto & "' "
                    Try
                        drMdb1 = cmdMdb1cr.ExecuteReader()
                        vSaldoMes = 0
                        vSaldoAnualReal = 0
                        If drMdb1.HasRows Then
                            While drMdb1.Read()
                                vSaldoAnualReal += -drMdb1.GetValue(2)
                                vFechaMes = drMdb1.GetValue(0).ToString
                                If Mid(vFechaMes, 4, 2) = vMes Then
                                    vSaldoMes += drMdb1.GetValue(2).ToString
                                End If
                            End While
                        Else
                            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                        End If
                        drMdb1.Close()
                    Catch ex As Exception
                        MsgBox("Error al ejecutar: " & cmdMdb1cr.CommandText & " por: " & ex.Message)
                    End Try
                    fila.Cells(2).Value = -vSaldoMes
                Next
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
                .Columns(0).HeaderText = "Código"
                .Columns(0).Width = 230
                .Columns(1).HeaderText = "Descripción"
                .Columns(1).Width = 230
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
                .Columns(0).HeaderText = "Nombres Existentes"
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
                .Columns(0).HeaderText = "Nombres Existentes"
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
                .Columns(0).HeaderText = "Nombres Existentes"
                .Columns(0).Width = 230
            End With
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
                frmCuentasBancarias.TxtIngresos.Text = Format(vIngresos, "###,##0.00").ToString
            Else
                vGastos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.IndianRed
                frmCuentasBancarias.TxtGastos.Text = Format(vGastos, "###,##0.00").ToString
            End If
        Next
        vSaldo = vIngresos + vGastos
        frmCuentasBancarias.TxtSaldo.Text = Format(vSaldo, "###,##0.00").ToString
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
                frmApuntesContables.TxtIngresos.Text = Format(vIngresos, "###,##0.00").ToString
            Else
                vGastos += fila.Cells(vFila1).Value
                fila.Cells(vFila1).Style.ForeColor = Color.IndianRed
                frmApuntesContables.TxtGastos.Text = Format(vGastos, "###,##0.00").ToString
            End If
            If fila.Cells(vFila2).Value >= 0 Then
                fila.Cells(vFila2).Style.ForeColor = Color.DarkBlue
            Else
                fila.Cells(vFila2).Style.ForeColor = Color.IndianRed
            End If
        Next
        frmApuntesContables.TxtSaldo.Text = Format(vValor, "###,##0.00").ToString
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
                frmApuntesPeriodicos.TxtIngresos.Text = Format(vIngresos, "###,##0.00").ToString
            Else
                vGastos += fila.Cells(3).Value
                fila.Cells(3).Style.ForeColor = Color.IndianRed
                frmApuntesPeriodicos.TxtGastos.Text = Format(vGastos, "###,##0.00").ToString
            End If
            If fila.Cells(4).Value >= 0 Then
                fila.Cells(4).Style.ForeColor = Color.DarkBlue
            Else
                fila.Cells(4).Style.ForeColor = Color.IndianRed
            End If
        Next
        frmApuntesPeriodicos.TxtSaldo.Text = Format(vValor, "###,##0.00").ToString
        Return vValor
    End Function

    ''Función para cambiar la fecha de orden a yyyy/MM/dd
    ''***************************************************
    'Public Function CambiarFecha(ByVal tipoFecha As String)
    '    Dim NuevaFecha As String
    '    NuevaFecha = Mid(tipoFecha, 7, 4) & "/" & Mid(tipoFecha, 4, 2) & "/" & Mid(tipoFecha, 1, 2).ToString
    '    Return NuevaFecha
    'End Function

    ''Función para cambiar la fecha de orden a MM/dd/yyyy
    ''***************************************************
    'Public Function CambiarFecha2(ByVal tipoFecha As String)
    '    Dim NuevaFecha2 As String
    '    NuevaFecha2 = Mid(tipoFecha, 4, 2) & "/" & Mid(tipoFecha, 1, 2) & "/" & Mid(tipoFecha, 7, 4).ToString
    '    Return NuevaFecha2
    'End Function

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

                    'vCalculoVersion1 = Val(Mid(vNewVersion, 5, 1)) + 1
                    'If vCalculoVersion1 < 10 Then
                    '    vNumeroVersion = Mid(vNewVersion, 1, 4) & vCalculoVersion1.ToString
                    'Else
                    '    vCalculoVersion2 = Val(Mid(vNewVersion, 3, 1)) + 1
                    '    If vCalculoVersion2 < 10 Then
                    '        vNumeroVersion = Mid(vNewVersion, 1, 2) & vCalculoVersion2.ToString & ".0"
                    '    Else
                    '        vCalculoVersion3 = Val(Mid(vNewVersion, 1, 1)) + 1
                    '        If vCalculoVersion3 < 10 Then
                    '            vNumeroVersion = vCalculoVersion3 & ".0.0"
                    '        Else
                    '            MsgBox("Error en el cálculo de la nueva versión, revise el formato de la versión")
                    '        End If
                    '    End If
                    'End If
                    'MsgBox("Versión Instalada: " & My.Settings.Version & vbNewLine & "Versión Disponible: " & vNumeroVersion, MsgBoxStyle.Information, "Comprobar Nueva Versión")

                    If My.Settings.Version < vNewVersion Then
                        MsgBox("Versión Instalada: " & My.Settings.Version & vbNewLine & "Versión Disponible: " & vNewVersion, MsgBoxStyle.Information, "Comprobar Nueva Versión")
                        vNuevaVersion = vNewVersion
                        vHayNuevaVersion = "SI"
                        respuesta = MsgBox("¿Quieres actualizar a la Versión: " & vNewVersion & " ?", vbQuestion + vbYesNo + vbDefaultButton1, "Versión ContaHogar 3.0")
                        If respuesta = vbYes Then
                            respuesta = MsgBox("Quieres guardar una Copia de Seguridad de la Base de Datos.", vbQuestion + vbYesNo + vbDefaultButton1, "Actualizar Software")
                            If respuesta = vbYes Then
                                ' Si no existe la carpeta de BackUp la creamos.
                                Dim path As String = "C:\ContaHogar3.0\Backup"
                                If Directory.Exists(path) Then
                                    'MsgBox("Ya existe la Ruta C:\ConatHogar\Backup.")
                                Else
                                    Directory.CreateDirectory(path)
                                    'MsgBox("Ruta C:\ContaHogar3.0\Backup, Creada.")
                                End If
                                Dim NombreBaseDatos As String = "ContaHogar3.0" & "[" & Format(Now.ToString("ddMMyyyy")) & "]" & "[" & Format(Now.ToString("HHmmss")) & "]" & ".mdb"
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
            ' 1. Crea una variable para acumular el texto arriba de tu bucle While
            Dim historialSeguimiento As String = "--- HISTORIAL DE TRADUCCIONES ---" & vbNewLine
            combo.Items.Clear()
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    Dim valorBD As String = drMdb1.GetValue(0).ToString().Trim()
                    ' BUSQUEDA DIRECTA: Ya no importa si es global o local, usa el que le pases
                    textoTraducido = rm.GetString(valorBD)
                    If String.IsNullOrEmpty(textoTraducido) Then
                        textoTraducido = valorBD
                    End If
                    combo.Items.Add(textoTraducido)
                    ' 2. En lugar de MsgBox, vas acumulando las líneas aquí
                    historialSeguimiento &= $"BD: {valorBD} -> Trad: {textoTraducido}" & vbNewLine
                End While
                ' 3. Muestras un ÚNICO MsgBox con todo el resumen al terminar el bucle
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



    'Public Function ReadINIkey(file As String, section As String, key As String) As String
    '    Dim lret As Long, i As Long
    '    Dim ret As String, newstr As String = "", c As Char
    '    If file = "" Then
    '        Return ""
    '        Exit Function
    '    End If
    '    ret = New String(CChar(" "), 255)
    '    lret = GetPrivateProfileStringKey(section.Trim, key.Trim, "", ret, Len(ret), file.Trim)
    '    If InStr(ret, Chr(0)) Then
    '        ret = Left$(ret, Len(ret) - 1)
    '    End If
    '    ret = ret.Trim
    '    ret = ret.Replace(" ", "|")
    '    For i = 0 To ret.Length - 1
    '        c = ret.Substring(i, 1).Trim
    '        If Char.IsControl(c) = False Then newstr = newstr.Trim & c.ToString().Trim
    '    Next
    '    ret = newstr.Trim
    '    ret = ret.Replace("|", " ")
    '    If ret = "" Then
    '        Return ""
    '    Else
    '        Return ret.Trim
    '    End If
    'End Function

    'Public Function SaveINIkey(file As String, section As String, key As String, value As String) As Boolean
    '    Dim lret As Long
    '    Dim ret As String = ""
    '    Try
    '        lret = WritePrivateProfileString(section.Trim, key.Trim, value.Trim, file.Trim)
    '        ret = lret.ToString().Trim().ToLower
    '    Catch ex As Exception
    '        MessageBox.Show("No se ha escrito nada en INI")
    '    End Try
    '    If ret = "0" Then
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function

End Module
