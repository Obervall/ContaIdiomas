Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class ApuntesContables

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vTxtNombre, filaActual, vPosicion As String
    Public vTxtDescripcion, BtnFechasClick, vTipoConcepto, vCodigo, carpetaPdf As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, PrintLine, Contador, filaSelec As Integer
    Public fechaformatomin, fechaformatomax As Date
    Public x, y, z As Integer
    Public TL(30) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    ' Método recursivo para actualizar la fuente de todos los controles
    Private Sub CambiarTamañoFuente(ByVal controles As Control.ControlCollection, ByVal nuevoTamaño As Single)
        For Each ctrl As Control In controles
            ' Aplicar el nuevo tamaño manteniendo el tipo de letra y estilo (negrita, cursiva, etc.)
            ctrl.Font = New Font(ctrl.Font.FontFamily, nuevoTamaño, ctrl.Font.Style)

            ' Si el control contiene otros controles (como un Panel o GroupBox), se llama a sí mismo
            If ctrl.HasChildren Then
                CambiarTamañoFuente(ctrl.Controls, nuevoTamaño)
            End If
        Next
    End Sub

    Private Sub ApuntesContables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        ' 1. Convertimos el año de texto a un número entero de forma segura
        Dim anio As Integer
        If Not Integer.TryParse(vAñoEjercicio, anio) Then
            ' Si por alguna razón vAñoEjercicio no es un número válido, 
            ' le asignamos el año actual por defecto para que no rompa el programa
            anio = Date.Today.Year
        End If

        ' 2. Asignamos a tus variables el año obtenido (opcional, por si las usas en otra parte)
        vFecha1Enero = anio
        vFecha31Diciembre = anio

        ' 3. Configuramos los DateTimePicker de forma limpia
        Dim fechaInicio As New Date(anio, 1, 1)
        Dim fechaFin As New Date(anio, 12, 31)

        ' Aplicamos los límites y valores del primer control
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker1.MaxDate = fechaFin
        DateTimePicker1.Value = fechaInicio

        ' Aplicamos los límites y valores del segundo control
        DateTimePicker2.MinDate = fechaInicio
        DateTimePicker2.MaxDate = fechaFin
        DateTimePicker2.Value = fechaFin

        ' 4. El resto de tu lógica original
        BtnFechasClick = "NO"
        BtnFechasFondo.Visible = False

        ' Ejemplo de uso del ResourceManager para obtener una cadena traducida
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnFiltroCuenta, resManager.GetString("ToolTipAplicarFiltro"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSinFiltroCuenta, resManager.GetString("ToolTipQuitarFiltro"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnFiltroConcepto, resManager.GetString("ToolTipAplicarFiltro"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnSinFiltroConcepto, resManager.GetString("ToolTipQuitarFiltro"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnFiltroFecha, resManager.GetString("ToolTipAplicarFiltro"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnSinFiltroFecha, resManager.GetString("ToolTipQuitarFiltro"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnAñadirRegistro, resManager.GetString("ToolTipAñadir"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnEditarRegistro, resManager.GetString("ToolTipEditar"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnEliminarRegistro, resManager.GetString("ToolTipEliminar"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnTraspasarRegistro, rmse.GetString("ToolTipTraspasar"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnGraficos, resManager.GetString("ToolTipGraficos"))
        TL(14) = New ToolTip
        TL(14).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(15) = New ToolTip
        TL(15).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(16) = New ToolTip
        TL(16).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(17) = New ToolTip
        TL(17).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(18) = New ToolTip
        TL(18).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(19) = New ToolTip
        TL(19).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))
        TL(20) = New ToolTip
        TL(20).SetToolTip(Me.BtnEliminaSeleccion, resManager.GetString("ToolTipEliminaSeleccion"))
        TL(21) = New ToolTip
        TL(21).SetToolTip(Me.BtnFiltroChekedList, rmse.GetString("ToolTipChekedList"))
        TL(22) = New ToolTip
        TL(22).SetToolTip(Me.LblApuntes, rmse.GetString("ToolTipLabelApuntes"))
        TL(23) = New ToolTip
        TL(23).SetToolTip(Me.ListBox1, rmse.GetString("ToolTipListBox"))
        TL(24) = New ToolTip
        TL(24).SetToolTip(Me.BtnFechas, rmse.GetString("ToolTipFechas"))
        TL(25) = New ToolTip
        TL(25).SetToolTip(Me.BtnFiltroF5, rmse.GetString("ToolTipF5"))
        TL(26) = New ToolTip
        TL(26).SetToolTip(Me.BtnF6, resManager.GetString("ToolTipF6"))
        TL(27) = New ToolTip
        TL(27).SetToolTip(Me.BtnExcel, rmse.GetString("ToolTipExcel"))
        TL(28) = New ToolTip
        TL(28).SetToolTip(Me.BtnAumentar, rmse.GetString("ToolTipAumentar"))
        TL(29) = New ToolTip
        TL(29).SetToolTip(Me.BtnNormal, rmse.GetString("ToolTipNormal"))
        TL(30) = New ToolTip
        TL(30).SetToolTip(Me.BtnImportarBanco, rmse.GetString("ToolTipImportarBanco"))

        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox4.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox5.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox6.MouseMove, AddressOf VerificarFiltrosDesactivados

        ' También vigilamos el fondo del formulario por si el usuario saca el ratón rápido
        AddHandler Me.MouseMove, AddressOf VerificarFiltrosDesactivados


        ' Llenar Grid de APUNTES al cargra el programa
        '**********************************************
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &  'Celda 0
           "conceptos.DescripcionCON As [ConceptoAPU], " & ' Celda 1 (Texto visible)
           "apuntes.DescripcionAPU As [DescripcionAPU], " & ' Celda 2
           "apuntes.ImporteAPU As [ImporteAPU], " &       ' Celda 3
           "apuntes.ImporteAPU As [SaldoAPU], " &         ' Celda 4
           "apuntes.NotasAPU As [NotasAPU], " &           ' Celda 5
           "cuentas.NombreCUE As [CuentaAPU], " &         ' Celda 6 (Texto visible)
           "apuntes.CodigoAPU As [CodigoAPU], " &         ' Celda 7
           "conceptos.CodigoCON As [CodigoCON], " &       ' Celda 8 (¡CORREGIDO! Clave estable para resManager)
           "apuntes.ConceptoAPU As [IdConceptoCON], " &   ' Celda 9 (ID numérico concepto para guardar)
           "apuntes.CuentaAPU As [IdCuentaCUE] " &        ' Celda 10 (ID numérico cuenta para guardar)
           "FROM (apuntes " &
           "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
           "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        ' =========================================================================
        ' 🌟 RECARGA INDEPENDIENTE DE CONTROLES (La propuesta maestra)
        ' =========================================================================
        Try
            ' 1. Encendemos tu escudo protector antes de rellenar los componentes
            cargandoFormulario = True

            ' 🌟 CABLE A: Cargamos el ComboBox de forma independiente ordenado de la A a la Z puro
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 🌟 CABLE B: Cargamos el ListBox1 manteniendo tus cabeceras estéticas por grupos
            cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, DescripcionCON, TipoCON FROM conceptos"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            LlenarYTraducirListBoxConceptosBD(Me.ListBox1, drMdb1)

            ' 2. Apagamos el escudo tras la inyección exitosa en memoria RAM
            cargandoFormulario = False

            ' FORZAMOS la selección inicial dócil del primer concepto por defecto de fábrica
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 0
            End If

        Catch ex As Exception
            cargandoFormulario = False
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(resManager.GetString("ErrorIniciarComponentes") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' Llenar el Combo Cuenta de forma segura y traducida
        '***************************************************
        Try
            ' 1. Encendemos tu escudo protector antes de rellenar para evitar eventos prematuros
            cargandoFormulario = True
            ' 2. Llamamos a la función genérica de tu módulo pasándole TU combo de la pantalla
            LlenarComboCuentasGenerico(Me.CmbCuenta)
            ' 3. Apagamos tu escudo protector
            cargandoFormulario = False
            ' 4. SELECCIÓN SEGURA: Si el combo tiene datos, seleccionamos el primero; si no, lo vaciamos
            If CmbCuenta.Items.Count > 0 Then
                CmbCuenta.SelectedIndex = 0
            Else
                CmbCuenta.SelectedIndex = -1
                CmbCuenta.Text = ""
            End If
        Catch ex As Exception
            ' En caso de un error general en el formulario, apagamos el escudo como salvavidas
            cargandoFormulario = False
            MsgBox(resManager.GetString("ErrorLlenarCuentas") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))

        For Each columna As DataGridViewColumn In DgvApuntes.Columns
            If columna.Name <> "ImporteAPU" And columna.Name <> "SaldoAPU" And columna.Name <> "CuentaAPU" And columna.Name <> "CodigoAPU" And columna.Name <> "CodigoCON" And columna.Name <> "IdConceptoCON" And columna.Name <> "IdCuentaCUE" Then
                frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
            End If
        Next
        ' Al final del todo, avisamos que la carga terminó y los eventos ya pueden actuar
        cargandoFormulario = False
    End Sub

    Private Sub BtnFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnFiltroCuenta.Click
        ' Limpiamos la memoria de consultas anteriores al entrar
        cmdMdb1cr.Parameters.Clear()

        If ListBox1.SelectedItems.Count <> 0 Then
            MessageBox.Show(rmse.GetString("MsgAviso1"), rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

            Dim idConceptoSaldo As Integer = 1
            Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
                Dim resId = cmdBuscarId.ExecuteScalar()
                If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
            End Using

            ' Escudo de seguridad ligero
            Dim totalRegistros As Integer = 0
            Using cmdContar As New OleDbCommand("SELECT COUNT(*) FROM apuntes WHERE CuentaAPU = ? And EjercicioAPU = ?", conexion1)
                cmdContar.Parameters.AddWithValue("?", idCuentaSel)
                cmdContar.Parameters.AddWithValue("?", CInt(vAñoEjercicio))
                Try : totalRegistros = Convert.ToInt32(cmdContar.ExecuteScalar()) : Catch : End Try
            End Using

            ' Si la cuenta está vacía, avisamos y frenamos
            If totalRegistros = 0 Then
                cmdMdb1cr.Parameters.Clear()
                MessageBox.Show(resManager.GetString("AvisoNoHayRegistrosEjercicio"))
                Exit Sub
            End If

            ' Si tiene registros, activamos los estados de los botones
            BtnFiltroCuenta.Enabled = False
            BtnSinFiltroCuenta.Enabled = True

            ' Consulta SQL Maestra de 11 celdas
            vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "

            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

            If BtnFiltroFecha.Enabled = False Then
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuntes.FechaAPU >= ?"
                vtipoSql += " And apuntes.FechaAPU <= ?"
                cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
            End If

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"

            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnFiltroChekedList_Click(sender As Object, e As EventArgs) Handles BtnFiltroChekedList.Click
        ListBox1.Visible = True
        CmbConcepto.Enabled = False
        BtnFiltroChekedList.Enabled = False
    End Sub

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = 27 Then  'Tecla Esc
            'Quitar todos los checked
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, False)
            Next
            ListBox1.Visible = False
            CmbConcepto.Enabled = True
            BtnFiltroChekedList.Enabled = True
        End If
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        ' 🌟 SANEAMIENTO INICIAL: Vaciamos la memoria de parámetros anteriores de la app
        cmdMdb1cr.Parameters.Clear()

        ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Inner Joins para nombres claros)
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE WHERE "

        ' 2. Aplicamos la condición base del Ejercicio/Fechas usando el ID numérico de SALDO
        If BtnFechasClick = "SI" Then
            vtipoSql += $" apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " apuntes.EjercicioAPU = " & vAñoEjercicio.ToString & " "
        End If

        ' 3. GESTIÓN DE CONCEPTOS (Multiselección o Simple)
        If ListBox1.SelectedItems.Count <> 0 Then ' Listbox con multiselección activo
            TxtConcepto.Text = rmse.GetString("MsgText3")

            ' 🌟 LA CORRECCIÓN CLAVE: Rompemos el enlace del DataSource para poder manipular los Items a mano
            CmbConcepto.DataSource = Nothing
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.SelectedIndex = 0

            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True

            ' Creamos una lista para almacenar los IDs reales numéricos
            Dim listaIdsConceptos As New List(Of Integer)

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim vConceptoReal As String = Convert.ToString(ListBox1.SelectedItems(i))

                ' Ignoramos las cabeceras estéticas (** GASTO **, etc.)
                If vConceptoReal.StartsWith("**") Then Continue For

                ' 🌟 TRUCO MAESTRO: Buscamos el ID numérico de este concepto seleccionado de forma aislada
                Dim idConceptoFila As Integer = 0
                Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                    cmdId.Parameters.AddWithValue("?", vConceptoReal)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                End Using

                If idConceptoFila > 0 Then
                    listaIdsConceptos.Add(idConceptoFila)
                End If
            Next

            ' El resultado final será numérico puro e inmune a errores: And apuntes.ConceptoAPU IN (4, 12, 8)
            If listaIdsConceptos.Count > 0 Then
                vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIdsConceptos) & ") "
            End If

        Else ' No hay selección múltiple, usamos el ComboBox individual y su SelectedValue (ID numérico)
            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True

            Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
        End If

        ' 4. FILTROS ADICIONALES SANEADOS CON IDs NUMÉRICOS
        If BtnFiltroCuenta.Enabled = False Then
            Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
        End If

        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' Sincronizamos las interrogaciones al final
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
        End If

        ' 5. Cierre de la consulta y ordenación
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"

        BtnFiltroChekedList.Enabled = False
        ListBox1.Visible = False
        vtipoGrid = "APUNTES_CONTABLES"

        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)

        If DgvApuntes.RowCount > 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnFiltroFecha.Click
        ' =========================================================================
        ' 🚀 ENCENDEMOS LA REDONDA QUE GIRA: Cambiamos el cursor al modo Espera
        ' =========================================================================
        ' Esto le dice visualmente al usuario: "Estoy procesando tus 7.000 apuntes, espera un segundo..."
        Me.Cursor = Cursors.WaitCursor
        frmApuntesContables.Cursor = Cursors.WaitCursor ' Si usas la grilla del formulario de impresión

        ' 🌟 SANEAMIENTO INICIAL: Vaciamos cualquier rastro de parámetros anteriores en la app
        cmdMdb1cr.Parameters.Clear()

        If ListBox1.SelectedItems.Count <> 0 Then
            MessageBox.Show(rmse.GetString("MsgAviso1"), rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            BtnFiltroFecha.Enabled = False
            BtnSinFiltroFecha.Enabled = True

            ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma aislada
            Dim idConceptoSaldo As Integer = 1
            Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
                Dim resId = cmdBuscarId.ExecuteScalar()
                If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
            End Using

            ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Garantiza ver nombres legibles en la rejilla)
            vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

            ' Filtro base del ejercicio
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' 🌟 PASO CRÍTICO 1: Agregamos primero las condiciones fijas de IDs numéricos (Sin comillas)
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

            ' 🌟 PASO CRÍTICO 2: Las interrogaciones de las fechas van SIEMPRE al final de las condiciones
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' 🌟 PASO CRÍTICO 3: Sincronizamos las variables en el orden exacto de los signos '?'
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"

            vtipoGrid = "APUNTES_CONTABLES"
            Try
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)

            Catch ex As Exception
                MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
            Finally
                ' =========================================================================
                ' 🌟 APAGAMOS EL GIRO VISUAL: Devolvemos el ratón a su estado ordinario
                ' =========================================================================
                ' Ponemos este bloque dentro del 'Finally' para garantizar que el ratón 
                ' recupere su libertad pase lo que pase, incluso si la base de datos da un error.
                Me.Cursor = Cursors.Default
                frmApuntesContables.Cursor = Cursors.Default
            End Try

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnSinFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroCuenta.Click
        ' 🌟 SANEAMIENTO INICIAL: Vaciamos cualquier rastro de parámetros anteriores en la app
        cmdMdb1cr.Parameters.Clear()

        ' 1. Activamos los estados de los botones de cuenta originales
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False

        ' =========================================================================
        ' 🌟 REINICIO TOTAL AUTOMÁTICO DE LOS CONCEPTOS (Evita el texto congelado)
        ' =========================================================================
        ' Delegamos la limpieza en el botón de concepto que ya limpia el ListBox,
        ' apaga los estados visuales y vuelve a inyectar el DataSource real en el combo
        BtnSinFiltroConcepto.PerformClick()

        ' =========================================================================
        ' 🌟 ADEMÁS REINICIAMOS LAS FECHAS AL AÑO COMPLETO (Enero a Diciembre)
        ' =========================================================================
        ' Encendemos el escudo temporal para que los calendarios no hagan consultas basura
        cargandoFormulario = True

        If BtnFechasClick = "SI" Then
            BtnFechasClick = "NO"
            BtnFechasFondo.Visible = False
        End If

        Dim anio As Integer
        If Not Integer.TryParse(vAñoEjercicio, anio) Then anio = Date.Today.Year
        vFecha1Enero = anio
        vFecha31Diciembre = anio

        DateTimePicker1.Value = New Date(anio, 1, 1)
        DateTimePicker2.Value = New Date(anio, 12, 31)

        ' Restablecemos los estados de los botones de fecha
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False

        ' =========================================================================
        ' 2. EJECUTAMOS LA CONSULTA SQL MAESTRA DE 11 COLUMNAS ANUAL COMPLETA
        ' =========================================================================
        ' Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' Consulta relacional limpia de fábrica (Sin filtros restrictivos secundarios cruzados)
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &
                   "conceptos.DescripcionCON As [ConceptoAPU], " &
                   "apuntes.DescripcionAPU As [DescripcionAPU], " &
                   "apuntes.ImporteAPU As [ImporteAPU], " &
                   "apuntes.ImporteAPU As [SaldoAPU], " &
                   "apuntes.NotasAPU As [NotasAPU], " &
                   "cuentas.NombreCUE As [CuentaAPU], " &
                   "apuntes.CodigoAPU As [CodigoAPU], " &
                   "conceptos.CodigoCON As [CodigoCON], " &
                   "apuntes.ConceptoAPU As [IdConceptoCON], " &
                   "apuntes.CuentaAPU As [IdCuentaCUE] " &
                   "FROM (apuntes " &
                   "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
                   "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        ' Aplicamos el filtro base del ejercicio contable activo
        If BtnFechasClick = "SI" Then
            vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If

        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"

        ' Forzamos vaciado final de parámetros de la app antes de recargar
        cmdMdb1cr.Parameters.Clear()

        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)

        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        ' Apagamos el escudo de los calendarios al terminar con éxito
        cargandoFormulario = False
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' =========================================================================
        ' 🚀 PURGA RELACIONAL INDESTRUCTIBLE (¡Adiós 'DataReader Open' para siempre!)
        ' =========================================================================
        Try
            If drMdb1 IsNot Nothing Then drMdb1.Close()
            cmdMdb1cr.Cancel()
        Catch
        End Try

        ' Refresco biológico del canal de Access en la memoria RAM
        Try
            If conexion1.State = ConnectionState.Open Then
                conexion1.Close()
                conexion1.Open()
            End If
        Catch
        End Try

        ' 🌟 SANEAMIENTO INICIAL: Vaciamos cualquier rastro de parámetros anteriores en la app
        cmdMdb1cr.Parameters.Clear()

        ' 1. Activamos el escudo temporal para que el combo no dispare eventos a destiempo
        cargandoFormulario = True

        ' 2. LLENAR Y TRADUCIR CON NUESTRA FUNCIÓN MODULAR (Añadido el IdConceptoCON obligatorio)
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, TipoCON, DescripcionCON FROM conceptos ORDER BY TipoCON ASC, CodigoCON ASC"

        ' 1. Activamos el escudo temporal para que el combo no dispare eventos a destiempo
        cargandoFormulario = True

        ' =========================================================================
        ' 🚀 REPARADO: ELIMINACIÓN DE CONSULTAS COJAS (¡Adiós DataReader Open!)
        ' =========================================================================
        Try
            ' 1. Cargamos el ComboBox de forma pura de la A a la Z mediante su DataSource nativo
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 2. Cargamos el ListBox1. Usamos un comando aislado para que no choque con la variable global
            Using cmdList As New OleDb.OleDbCommand("SELECT IdConceptoCON, CodigoCON, DescripcionCON, TipoCON FROM conceptos", conexion1)
                Using drList As OleDbDataReader = cmdList.ExecuteReader()
                    ' Rellenamos el ListBox con su rutina dedicada traducida
                    LlenarYTraducirListBoxConceptosBD(Me.ListBox1, drList)
                End Using ' 🔴 ¡ÉXITO!: El lector local drList se cierra automáticamente aquí
            End Using

            ' FORZAMOS la selección inicial dócil del primer concepto por defecto de fábrica
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 0
            End If

            ' 4. Extraemos la descripción al vuelo desde el DataSource sin abrir lectores redundantes
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)
                TxtConcepto.Text = filaSeleccionada("DescripcionCON").ToString()
            Else
                TxtConcepto.Text = ""
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        Finally
            ' Aseguramos que el lector global quede purgado pase lo que pase
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
        End Try


        ' 5. Apagamos el escudo: el combo ya está listo
        cargandoFormulario = False

        ' Tus estados de informes originales
        frmTipoInformeApuntes.RadioButton1.Enabled = True
        frmTipoInformeApuntes.RadioButton2.Enabled = True
        frmTipoInformeApuntes.RadioButton5.Enabled = False
        BtnFiltroChekedList.Enabled = True
        BtnFiltroConcepto.Enabled = True
        CmbConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False

        ' Quitar todos los checked
        For i = 0 To ListBox1.Items.Count - 1
            ListBox1.SetSelected(i, False)
        Next

        ' 6. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Inner Joins para nombres claros)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql = sqlBase
        If BtnFechasClick = "SI" Then
            vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If

        ' 🌟 CORRECCIÓN CLAVE: Cambiado de texto a numérico (> 0)
        vtipoSql += " And apuntes.ConceptoAPU > 0 "

        ' Saneamos el filtro secundario de Cuenta inyectando su ID puro
        If BtnFiltroCuenta.Enabled = False Then
            Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
        End If

        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' Sincronizamos las interrogaciones al final
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
        End If

        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"

        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)

        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSinFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroFecha.Click
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Congelamos los calendarios para que no saboteen el proceso
        cargandoFormulario = True

        ' Saneamiento inicial de parámetros residuales en la memoria de la app
        cmdMdb1cr.Parameters.Clear()

        If BtnFechasClick = "SI" Then
            BtnFechasClick = "NO"
            BtnFechasFondo.Visible = False
        End If

        ' 1. Convertimos el año a número entero de forma segura
        Dim anio As Integer
        If Not Integer.TryParse(vAñoEjercicio, anio) Then
            anio = Date.Today.Year
        End If

        ' 2. Asignamos el año a tus variables originales
        vFecha1Enero = anio
        vFecha31Diciembre = anio

        ' 3. Reseteamos los calendarios al año completo de enero a diciembre
        ' (Gracias a cargandoFormulario = True, estas líneas no dispararán eventos basura)
        DateTimePicker1.Value = New Date(anio, 1, 1)
        DateTimePicker2.Value = New Date(anio, 12, 31)

        ' 4. Tu lógica de botones original de fechas
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False

        ' =========================================================================
        ' 🌟 TRUCO MAESTRO: DELEGAMOS EN TU BOTÓN DE CONCEPTO SI HAY MULTISELECCIÓN
        ' =========================================================================
        If ListBox1.SelectedItems.Count <> 0 Then
            ' Cedemos el control absoluto al botón de quitar conceptos que ya es estable
            BtnSinFiltroConcepto.PerformClick()

            ' Apagamos el escudo protector y salimos limpiamente de la rutina
            cargandoFormulario = False
            Exit Sub
        End If

        ' =========================================================================
        ' LÓGICA DEL ELSE (CUANDO EL USUARIO NO USABA EL LISTBOX DE CONCEPTOS)
        ' =========================================================================
        ' Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' Consulta SQL Maestra de 11 celdas (Saneada para LlenarGrid)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql = sqlBase
        If BtnFechasClick = "SI" Then
            vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If

        ' Si el filtro de cuenta sigue activo, mantenemos su ID numérico
        If BtnFiltroCuenta.Enabled = False Then
            Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
        End If

        ' Si el filtro de concepto ordinario sigue activo, mantenemos su ID numérico
        If BtnFiltroConcepto.Enabled = False Then
            Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
        End If

        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"

        ' Doble limpieza preventiva para LlenarGrid
        cmdMdb1cr.Parameters.Clear()

        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)

        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        ' 🌟 APAGAMOS EL ESCUDO: Todo el proceso ha terminado con éxito rotundo
        cargandoFormulario = False
    End Sub

    Private Sub CmbCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCuenta.SelectedIndexChanged
        ' ESCUDO: Si el formulario se está cargando o limpiando, salimos inmediatamente
        If cargandoFormulario Then Exit Sub
        If CmbCuenta.SelectedIndex < 0 Then Exit Sub

        ' =========================================================================
        ' 🌟 ADAPTACIÓN A IDs NUMÉRICOS: CERO SELECTIONS COMPLEMENTARIAS
        ' =========================================================================
        ' Extraemos los IDs numéricos directos de los combos sin abrir ningún lector (drMdb1)
        Dim idCuentaOriginal As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

        ' 🌟 ESCUDO EN EL COMBO: Si la cuenta elegida manualmente no tiene datos, avisamos de inmediato
        Dim registrosEnCombo As Integer = 0
        Using cmdContarCbo As New OleDb.OleDbCommand("SELECT COUNT(*) FROM apuntes WHERE CuentaAPU = ? And EjercicioAPU = ?", conexion1)
            cmdContarCbo.Parameters.AddWithValue("?", idCuentaOriginal)
            cmdContarCbo.Parameters.AddWithValue("?", CInt(vAñoEjercicio))
            Try : registrosEnCombo = Convert.ToInt32(cmdContarCbo.ExecuteScalar()) : Catch : End Try
        End Using

        If registrosEnCombo = 0 Then
            ' 🌟 CORRECCIÓN DEFINITIVA: Usamos la misma estructura de 11 columnas del Load,
            ' pero forzamos un filtro imposible (EjercicioAPU = 0) para vaciar el Grid limpiamente
            vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &
                       "conceptos.DescripcionCON As [ConceptoAPU], " &
                       "apuntes.DescripcionAPU As [DescripcionAPU], " &
                       "apuntes.ImporteAPU As [ImporteAPU], " &
                       "apuntes.ImporteAPU As [SaldoAPU], " &
                       "apuntes.NotasAPU As [NotasAPU], " &
                       "cuentas.NombreCUE As [CuentaAPU], " &
                       "apuntes.CodigoAPU As [CodigoAPU], " &
                       "conceptos.CodigoCON As [CodigoCON], " &
                       "apuntes.ConceptoAPU As [IdConceptoCON], " &
                       "apuntes.CuentaAPU As [IdCuentaCUE] " &
                       "FROM (apuntes " &
                       "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
                       "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE " &
                       "WHERE apuntes.EjercicioAPU = 0" ' ⬅️ Truco: Año 0 para vaciar sin romper tipos

            cmdMdb1cr.Parameters.Clear()
            LlenarGrid(vtipoSql, "APUNTES_CONTABLES", "1")
            MessageBox.Show(resManager.GetString("AvisoNoHayRegistrosEjercicio"))
            Exit Sub
        End If

        Dim idConceptoComboOriginal As Integer = 0
        If CmbConcepto.SelectedIndex >= 0 Then
            idConceptoComboOriginal = Convert.ToInt32(CmbConcepto.SelectedValue)
        End If

        ' Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' --- PROCESO DE FILTRADO DEL GRID ---
        If ListBox1.SelectedItems.Count = 0 Then
            If BtnFiltroCuenta.Enabled = False Then

                ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Con los INNER JOIN obligatorios para ver textos)
                vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &
                           "conceptos.DescripcionCON As [ConceptoAPU], " &
                           "apuntes.DescripcionAPU As [DescripcionAPU], " &
                           "apuntes.ImporteAPU As [ImporteAPU], " &
                           "apuntes.ImporteAPU As [SaldoAPU], " &
                           "apuntes.NotasAPU As [NotasAPU], " &
                           "cuentas.NombreCUE As [CuentaAPU], " &
                           "apuntes.CodigoAPU As [CodigoAPU], " &
                           "conceptos.CodigoCON As [CodigoCON], " &
                           "apuntes.ConceptoAPU As [IdConceptoCON], " &
                           "apuntes.CuentaAPU As [IdCuentaCUE] " &
                           "FROM (apuntes " &
                           "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
                           "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

                ' Filtro numérico puro para descartar el concepto SALDO
                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                ' CORRECCIÓN: Filtro numérico por ID de cuenta sin comillas
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaOriginal} "

                ' CORRECCIÓN: Filtro numérico por ID de concepto sin comillas
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoComboOriginal} "
                End If

                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"

                    ' Aseguramos los parámetros de fecha en el comando general de la app
                    cmdMdb1cr.Parameters.Clear()
                    cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                    cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
                Else
                    cmdMdb1cr.Parameters.Clear()
                End If

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)

                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        Else
            ' =========================================================================
            ' 🌟 FILTRADO MÚLTIPLE CON EL LISTBOX1 SELECCIONADO (ADAPTADO A IDs)
            ' =========================================================================
            If BtnFiltroCuenta.Enabled = False Then
                Dim i As Integer

                ' Consulta SQL Maestra con INNER JOINs (Igual a la del Load para ver nombres claros)
                Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"
                vtipoSql = sqlBase

                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                ' Preparamos la limpieza de fechas para el comando global
                cmdMdb1cr.Parameters.Clear()
                Dim tieneFechasActivo As Boolean = (BtnFiltroFecha.Enabled = False)

                For i = 0 To ListBox1.SelectedItems.Count - 1
                    Dim conceptoTraducido As String = ListBox1.SelectedItems(i).ToString()

                    ' Ignoramos las cabeceras estéticas (** GASTO **, etc.)
                    If conceptoTraducido.StartsWith("**") Then Continue For

                    ' 🌟 TRUCO MAESTRO: Buscamos el ID numérico del concepto seleccionado de forma aislada
                    ' Eliminamos el bucle While drMdb1.Read() pesado de antes
                    Dim idConceptoFila As Integer = 0
                    Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                        cmdId.Parameters.AddWithValue("?", conceptoTraducido)
                        Dim resId = cmdId.ExecuteScalar()
                        If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                    End Using

                    ' Si no se encuentra traducción o ID, le asignamos 0 o pasamos por seguridad
                    If idConceptoFila = 0 Then Continue For

                    If i = 0 Then
                        ' CORRECCIÓN: Filtros numéricos puros por ID sin comillas
                        vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "
                        vtipoSql += $" And apuntes.CuentaAPU = {idCuentaOriginal} "

                        If tieneFechasActivo Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += $" apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += " apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If

                        ' CORRECCIÓN: Filtros numéricos por ID en la parte del OR
                        vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "
                        vtipoSql += $" And apuntes.CuentaAPU = {idCuentaOriginal} "

                        If tieneFechasActivo Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
                        End If
                    End If
                Next

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"

                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)

                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        End If
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ' 🌟 ESCUDO PROTECTOR: Si estamos cargando o limpiando filtros por código, salimos de inmediato
        If cargandoFormulario Then Exit Sub

        ' Si el filtro de fechas no está activo en la pantalla, no hace falta recalcular nada
        If BtnFiltroFecha.Enabled = True Then Exit Sub

        ' Saneamiento inicial de parámetros del comando general
        cmdMdb1cr.Parameters.Clear()

        ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Inner Joins para nombres claros)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        If ListBox1.SelectedItems.Count = 0 Then
            ' =========================================================================
            ' LÓGICA SIN MULTISELECCIÓN EN LISTBOX
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' Filtros de IDs secundarios
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

            ' Las interrogaciones de fecha van SIEMPRE al final
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' Sincronizamos parámetros secuenciales ? de OleDb
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        Else
            ' =========================================================================
            ' LÓGICA CON MULTISELECCIÓN (Convertida al IN numérico indestructible)
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' Recolectamos los IDs de conceptos seleccionados
            Dim listaIdsConceptos As New List(Of Integer)
            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim vConceptoReal As String = ListBox1.SelectedItems(i).ToString()
                If vConceptoReal.StartsWith("**") Then Continue For

                Dim idConceptoFila As Integer = 0
                Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                    cmdId.Parameters.AddWithValue("?", vConceptoReal)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                End Using

                If idConceptoFila > 0 Then listaIdsConceptos.Add(idConceptoFila)
            Next

            If listaIdsConceptos.Count = 0 Then listaIdsConceptos.Add(0)
            vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIdsConceptos) & ") "

            ' Filtro de cuenta por ID numérico
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            ' Agregamos las interrogaciones de fecha al final de todo el bloque
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        ' 🌟 ESCUDO PROTECTOR: Si estamos cargando o limpiando filtros por código, salimos de inmediato
        If cargandoFormulario Then Exit Sub

        ' Si el filtro de fechas no está activo en la pantalla, no hace falta recalcular nada
        If BtnFiltroFecha.Enabled = True Then Exit Sub

        ' Saneamiento inicial de parámetros del comando general
        cmdMdb1cr.Parameters.Clear()

        ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Inner Joins para nombres claros)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        If ListBox1.SelectedItems.Count = 0 Then
            ' =========================================================================
            ' LÓGICA SIN MULTISELECCIÓN EN LISTBOX
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' Filtros de IDs secundarios
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

            ' Las interrogaciones de fecha van SIEMPRE al final
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            ' Sincronizamos parámetros secuenciales ? de OleDb
            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        Else
            ' =========================================================================
            ' LÓGICA CON MULTISELECCIÓN (Convertida al IN numérico indestructible)
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' Recolectamos los IDs de conceptos seleccionados
            Dim listaIdsConceptos As New List(Of Integer)
            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim vConceptoReal As String = ListBox1.SelectedItems(i).ToString()
                If vConceptoReal.StartsWith("**") Then Continue For

                Dim idConceptoFila As Integer = 0
                Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                    cmdId.Parameters.AddWithValue("?", vConceptoReal)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                End Using

                If idConceptoFila > 0 Then listaIdsConceptos.Add(idConceptoFila)
            Next

            If listaIdsConceptos.Count = 0 Then listaIdsConceptos.Add(0)
            vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIdsConceptos) & ") "

            ' Filtro de cuenta por ID numérico
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            ' Agregamos las interrogaciones de fecha al final de todo el bloque
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"

            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        frmPrincipal.TsLabelFormulario.Text = rmse.GetString("MsgText2")
        ' 1. Controlar la instancia física del formulario de forma tradicional
        If ((frmIntroApuntes Is Nothing) OrElse (Not frmIntroApuntes.IsHandleCreated)) Then
            frmIntroApuntes = New IntroApuntes
        End If
        ' 2. Forzar la traducción antes de mostrarlo (por consistencia con tus otros formularios)
        ' Nota: Si tienes el método ActualizarTextosFormulario accesible, puedes llamarlo aquí:
        ActualizarTextosFormulario(frmIntroApuntes)
        ' 3. ¡EL TRUCO!: Decimos que se centre respecto a su contenedor "padre"
        frmIntroApuntes.StartPosition = FormStartPosition.CenterParent
        ' ¡LA PROTECCIÓN CRÍTICA!: Procesamos todos los mensajes visuales pendientes en Windows
        Application.DoEvents()
        ' 4. Abrimos el formulario modal pasando "Me" (este segundo formulario) como dueño
        frmIntroApuntes.ShowDialog(Me)
        ' 5. Destrucción explícita al cerrar
        frmIntroApuntes.Dispose()
        ' 6. IMPORTANTE: Limpiar la variable manual para evitar el error de objeto destruido
        frmIntroApuntes = Nothing
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        ' 1. Validamos que haya al menos una fila seleccionada de verdad en la rejilla
        If DgvApuntes.SelectedRows.Count = 0 Then
            MsgBox(resManager.GetString("MsgSeleccionarFilasEliminar"), MsgBoxStyle.Information)
            Exit Sub
        End If

        ' 2. CUADRO DE CONFIRMACIÓN: Preguntamos antes de extirpar de la BD el lote completo
        ' =========================================================================
        ' 🌟 ADAPTACIÓN INTERNACIONAL DE MENSAJES MEDIANTE PLANTILLAS (resManager)
        ' =========================================================================
        Dim mensajeConfirmacion As String = ""
        Dim totalFilas As Integer = DgvApuntes.SelectedRows.Count
        If totalFilas = 1 Then
            ' Recuperamos el molde singular de fábrica (No necesita inyectar números)
            Dim plantillaSingular As String = resManager.GetString("MsgConfirmarBorradoSingular")
            ' Salvavidas por si la Key no estuviera escrita en el .resx
            If String.IsNullOrEmpty(plantillaSingular) Then
                plantillaSingular = "¿Está completamente seguro de que desea eliminar FÍSICAMENTE de la Base de Datos el apunte seleccionado?"
            End If
            mensajeConfirmacion = plantillaSingular
        Else
            ' Recuperamos el molde plural que contiene el comodín {0}
            Dim plantillaPlural As String = resManager.GetString("MsgConfirmarBorradoPlural")
            ' Salvavidas por si la Key no estuviera escrita en el .resx
            If String.IsNullOrEmpty(plantillaPlural) Then
                plantillaPlural = "¿Está completamente seguro de que desea eliminar FÍSICAMENTE de la Base de Datos los {0} apuntes seleccionados?"
            End If
            ' 🌟 EL TRUCO .NET: String.Format inyecta el número 'totalFilas' en el lugar del '{0}'
            mensajeConfirmacion = String.Format(plantillaPlural, totalFilas)
        End If
        ' 🌟 LA CORRECCIÓN CLAVE: Llamamos al confirmador inmune al idioma de Windows para que no salga YES en vez de SI
        Dim tituloVentana As String = If(resManager?.GetString("ConfirmarBorrado"), "Confirmar Borrado Múltiple")
        If ConfirmarAccionTraducida(mensajeConfirmacion, tituloVentana) = MsgBoxResult.No Then
            Exit Sub
        End If

        ' 3. Buscamos el ID numérico del concepto "SALDO" para blindarlo contra borrados accidentales
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        Dim contadorBorrados As Integer = 0

        ' =========================================================================
        ' 🌟 PUNTO 3 REPARADO: BUCLE PARA RECOLECTAR TODOS LOS IDs SELECCIONADOS
        ' =========================================================================
        For Each fila As DataGridViewRow In DgvApuntes.SelectedRows
            ' Saltamos la fila vacía del final del grid por seguridad
            If fila.IsNewRow Then Continue For

            ' ESCUDO DE SEGURIDAD 1: Validamos que la fila tenga conceptos reales y que no sea SALDO
            If fila.Cells(9).Value IsNot Nothing AndAlso Not IsDBNull(fila.Cells(9).Value) Then
                Dim idConceptoFila As Integer = Convert.ToInt32(fila.Cells(9).Value)
                If idConceptoFila = idConceptoSaldo Then Continue For ' Si incluye el SALDO inicial, pasa de largo y lo salva
            End If

            ' ESCUDO DE SEGURIDAD 2: Rescatamos el ID físico único de esta fila (Celda 7)
            If fila.Cells(7).Value IsNot Nothing AndAlso Not IsDBNull(fila.Cells(7).Value) Then
                Dim idRegistroFisico As Integer = Convert.ToInt32(fila.Cells(7).Value)

                ' Lanzamos la sentencia DELETE individual para cada ID del lote de forma ultra rápida
                Using cmdDelete As New OleDb.OleDbCommand("DELETE FROM apuntes WHERE CodigoAPU = ?", conexion1)
                    cmdDelete.Parameters.Clear()
                    cmdDelete.Parameters.Add("@id", OleDb.OleDbType.Integer).Value = idRegistroFisico

                    Try
                        cmdDelete.ExecuteNonQuery()
                        contadorBorrados += 1
                    Catch ex As Exception
                        ' 🌟 Recuperamos el molde del .resx de forma segura
                        Dim plantillaError As String = resManager.GetString("ErrorBorrarFilaID")

                        ' Salvavidas por si acaso la Key no estuviera escrita todavía
                        If String.IsNullOrEmpty(plantillaError) Then
                            plantillaError = "Error al borrar la fila con ID {0}: "
                        End If

                        ' Fusionamos el ID con el texto traducido y le pegamos el mensaje nativo del sistema (ex.Message)
                        Dim mensajeFinal As String = String.Format(plantillaError, idRegistroFisico) & ex.Message

                        MsgBox(mensajeFinal, MsgBoxStyle.Critical, resManager.GetString("Error"))
                    End Try
                End Using
            End If
        Next

        ' =========================================================================
        ' 4. REFRESCAMOS Y RECALCULAMOS LA INTERFAZ DE FORMA AUTOMÁTICA
        ' =========================================================================
        ' Refrescamos la rejilla con tu rutina estrella de las 11 celdas relacionales
        RefrescarGridApuntesContables()

        ' Volvemos a calcular la columna de saldos acumulados de la pantalla al céntimo
        DgvApuntesContables(3, 4)

        ' Avisamos del resultado final al usuario
        ' =========================================================================
        ' 🌟 MENSAJE DE ÉXITO FINAL INTERNACIONALIZADO (resManager)
        ' =========================================================================
        ' 1. Recuperamos la plantilla del mensaje del .resx
        Dim plantillaExito As String = resManager.GetString("MsgLoteEliminadoExito")
        If String.IsNullOrEmpty(plantillaExito) Then
            plantillaExito = "Operación completada. Se han eliminado {0} registros físicos de la Base de Datos."
        End If

        ' 2. Recuperamos el título de la ventana del .resx
        Dim tituloExito As String = resManager.GetString("TituloBorradoFinalizado")
        If String.IsNullOrEmpty(tituloExito) Then
            tituloExito = "Borrado Finalizado"
        End If

        ' 3. Fusionamos el número de borrados con la plantilla correspondiente
        Dim mensajeFinalExito As String = String.Format(plantillaExito, contadorBorrados)

        ' Mostramos el aviso final dócil y traducido
        MsgBox(mensajeFinalExito, MsgBoxStyle.Information, tituloExito)
    End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoGrafico Is Nothing) OrElse (Not frmTipoGrafico.IsHandleCreated)) Then
            frmTipoGrafico = New TipoGrafico
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmTipoGrafico)
        ' Llamamos al formulario de manera modal.
        frmTipoGrafico.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoGrafico.Dispose()
        ' Llenar Grid de APUNTES al cargra el programa
        '**********************************************
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &  'Celda 0
           "conceptos.DescripcionCON As [ConceptoAPU], " & ' Celda 1 (Texto visible)
           "apuntes.DescripcionAPU As [DescripcionAPU], " & ' Celda 2
           "apuntes.ImporteAPU As [ImporteAPU], " &       ' Celda 3
           "apuntes.ImporteAPU As [SaldoAPU], " &         ' Celda 4
           "apuntes.NotasAPU As [NotasAPU], " &           ' Celda 5
           "cuentas.NombreCUE As [CuentaAPU], " &         ' Celda 6 (Texto visible)
           "apuntes.CodigoAPU As [CodigoAPU], " &         ' Celda 7
           "conceptos.CodigoCON As [CodigoCON], " &       ' Celda 8 (¡CORREGIDO! Clave estable para resManager)
           "apuntes.ConceptoAPU As [IdConceptoCON], " &   ' Celda 9 (ID numérico concepto para guardar)
           "apuntes.CuentaAPU As [IdCuentaCUE] " &        ' Celda 10 (ID numérico cuenta para guardar)
           "FROM (apuntes " &
           "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
           "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        BtnSinFiltroFecha.PerformClick()
        BtnSinFiltroConcepto.PerformClick()
        BtnSinFiltroCuenta.PerformClick()

        ' =========================================================================
        ' 🌟 RECARGA INDEPENDIENTE DE CONTROLES (La propuesta maestra)
        ' =========================================================================
        Try
            ' 1. Encendemos tu escudo protector antes de rellenar los componentes
            cargandoFormulario = True

            ' 🌟 CABLE A: Cargamos el ComboBox de forma independiente ordenado de la A a la Z puro
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 🌟 CABLE B: Cargamos el ListBox1 manteniendo tus cabeceras estéticas por grupos
            cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, DescripcionCON, TipoCON FROM conceptos"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            LlenarYTraducirListBoxConceptosBD(Me.ListBox1, drMdb1)

            ' 2. Apagamos el escudo tras la inyección exitosa en memoria RAM
            cargandoFormulario = False

            ' FORZAMOS la selección inicial dócil del primer concepto por defecto de fábrica
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 0
            End If

        Catch ex As Exception
            cargandoFormulario = False
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(resManager.GetString("ErrorIniciarComponentes") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        ' Elimina Visualmente las Filas Seleccionadas desde la memoria RAM
        ' *******************************************************************
        If DgvApuntes.SelectedRows.Count > 0 Then

            ' 🌟 CAMBIO DE LA NUEVA ERA: Borramos de la memoria RAM (DataTable) y no de la rejilla visual
            ' Recorremos al revés o de forma directa las filas seleccionadas
            For i As Integer = DgvApuntes.SelectedRows.Count - 1 To 0 Step -1
                Dim fila As DataGridViewRow = DgvApuntes.SelectedRows(i)

                ' Nos saltamos la fila vacía del final por seguridad
                If fila.IsNewRow Then Continue For

                ' Extraemos el enlace de datos puro de la fila y lo eliminamos de la RAM
                If fila.DataBoundItem IsNot Nothing Then
                    Dim rowView As DataRowView = CType(fila.DataBoundItem, DataRowView)
                    rowView.Delete() ' Se borra del DataTable en la RAM, pero NO hace el DELETE en Access
                End If
            Next

            ' 2. Tu excelente bloque contable de recalculo se queda intacto y dócil
            If DgvApuntes.Rows.Count > 0 Then
                ' Limpiamos selecciones fantasma
                For i = 0 To DgvApuntes.Rows.Count - 1
                    DgvApuntes.Rows(i).Selected = False
                Next

                ' Vuelve a calcular la columna de saldos acumulados con lo que queda vivo en la pantalla
                DgvApuntesContables(3, 4)

                DgvApuntes.Select()
                ' 🌟 LA CORRECCIÓN CLAVE: Calculamos de forma dinámica el índice de la ÚLTIMA fila
                Dim ultimaFilaViva As Integer = DgvApuntes.Rows.Count - 1

                ' Validamos que el índice sea válido (mayor o igual a 0) para evitar desbordamientos
                If ultimaFilaViva >= 0 Then
                    DgvApuntes.Rows(ultimaFilaViva).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(ultimaFilaViva).Cells(0)
                End If
                DgvApuntes.Refresh()
            End If
        End If
    End Sub

    Private Sub DgvApuntes_KeyDown(sender As Object, e As KeyEventArgs) Handles DgvApuntes.KeyDown
        ' 🌟 LA CORRECCIÓN CLAVE: Detectamos si el usuario pulsa la tecla SUPR (Delete)
        If e.KeyCode = Keys.Delete OrElse e.KeyCode = 46 Then
            ' Evitamos que Windows Forms aplique su borrado visual interno basura
            e.SuppressKeyPress = True

            ' 🚀 TRUCO MAESTRO DE REUTILIZACIÓN: Disparamos tu botón unificado de borrado físico masivo
            ' que ya pide confirmación traducida, salva el SALDO inicial y recalcula los céntimos
            BtnEliminarRegistro.PerformClick()
        End If
    End Sub

    Private Sub BtnFechas_Click(sender As Object, e As EventArgs) Handles BtnFechas.Click
        If BtnFechasClick = "NO" Then
            BtnFechasClick = "SI"
            BtnFechasFondo.Visible = True

            ' 🌟 CLON DE TU BAÚL: Capturamos el año más ANTIGUO de la historia
            cmdMdb1cr.CommandText = "SELECT TOP 1 EjercicioEJE FROM ejercicios ORDER BY EjercicioEJE ASC"
            cmdMdb1cr.Parameters.Clear()

            Try
                Dim resAnio = cmdMdb1cr.ExecuteScalar()
                If resAnio IsNot Nothing AndAlso Not IsDBNull(resAnio) Then
                    ' 🚀 REPARADO: Guardamos el año puro (Integer) en tu variable de siempre
                    vFecha1Enero = Convert.ToInt32(resAnio)
                Else
                    vFecha1Enero = Date.Today.Year
                End If
            Catch ex As Exception
                vFecha1Enero = Date.Today.Year
                MsgBox(resManager.GetString("ErrorLeerHistoricos") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try

            ' Capturamos el año del ejercicio contable activo en curso (ej: 2025 o 2026)
            Dim anioCurso As Integer
            If Not Integer.TryParse(vAñoEjercicio, anioCurso) Then
                anioCurso = Date.Today.Year
            End If

            ' =========================================================================
            ' 🌟 CORTAFUEGOS INDESTRUCTIBLE DE LÍMITES CRONOLÓGICOS (Evita colisiones en la RAM)
            ' =========================================================================
            DateTimePicker1.MinDate = New Date(1753, 1, 1)
            DateTimePicker1.MaxDate = New Date(9998, 12, 31)
            DateTimePicker2.MinDate = New Date(1753, 1, 1)
            DateTimePicker2.MaxDate = New Date(9998, 12, 31)

            ' =========================================================================
            ' 🌟 CONFIGURACIÓN CRONOLÓGICA (Conversión segura en caliente)
            ' =========================================================================
            ' Creamos las variables de fecha locales para que el calendario las trague sin quejas
            Dim fechaInicioHistorico As New Date(Convert.ToInt32(vFecha1Enero), 1, 1)
            Dim fechaFinCurso As New Date(anioCurso, 12, 31)

            ' Calendario 1: Se clava en el 1 de Enero del año más antiguo encontrado
            DateTimePicker1.MinDate = fechaInicioHistorico
            DateTimePicker1.Value = fechaInicioHistorico

            ' Calendario 2: Se estira de forma elástica hasta el 31 de Diciembre del ejercicio activo
            DateTimePicker2.MinDate = fechaInicioHistorico
            DateTimePicker2.Value = fechaFinCurso
        Else
            ' =========================================================================
            ' TU LÓGICA ORIGINAL PERFECTA DEL ELSE (Sincronizada con el año del ejercicio activo)
            ' =========================================================================
            Dim anio As Integer
            If Not Integer.TryParse(vAñoEjercicio, anio) Then
                anio = Date.Today.Year
            End If

            Dim fechaInicio As New Date(anio, 1, 1)
            Dim fechaFin As New Date(anio, 12, 31)

            ' Abrimos las compuertas preventivas antes de encajonar las fechas de nuevo
            DateTimePicker1.MinDate = New Date(1753, 1, 1)
            DateTimePicker1.MaxDate = New Date(9998, 12, 31)
            DateTimePicker2.MinDate = New Date(1753, 1, 1)
            DateTimePicker2.MaxDate = New Date(9998, 12, 31)

            DateTimePicker1.MinDate = fechaInicio
            DateTimePicker1.MaxDate = fechaFin
            DateTimePicker1.Value = fechaInicio

            DateTimePicker2.MinDate = fechaInicio
            DateTimePicker2.MaxDate = fechaFin
            DateTimePicker2.Value = fechaFin

            BtnFechasClick = "NO"
            BtnFechasFondo.Visible = False
        End If
    End Sub

    Private Sub DgvApuntes_DoubleClick(sender As Object, e As EventArgs) Handles DgvApuntes.DoubleClick
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        Dim fechaFila As DateTime
        If DateTime.TryParse(frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(0).Value.ToString(), fechaFila) Then
            If fechaFila.Year <> vAñoEjercicio Then
                MessageBox.Show(resManager.GetString("NoEditarFueraEjercicio"), resManager.GetString("Validación"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
                vTxtNombre = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString

                ' Rescatamos el ID único del apunte desde la celda oculta de tu nueva estructura (Celda 7 u 8 según tu Grid)
                ' Si vCodigo sigue funcionando como buscador único, lo mantenemos intacto
                vCodigo = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(7).Value.ToString

                ' Comparación segura respetando las traducciones del sistema
                Dim palabraSaldoMayusculas As String = resManager.GetString("Saldo")?.ToUpper()
                If String.IsNullOrEmpty(palabraSaldoMayusculas) Then palabraSaldoMayusculas = "SALDO"

                If vTxtNombre.ToUpper() = palabraSaldoMayusculas Then
                    MsgBox(resManager.GetString("MsgSaldos2"))
                Else
                    ' Comprobamos si existe un identificador asociado.
                    If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
                        frmEditarApuntes = New EditarApuntes
                    End If
                    ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
                    ActualizarTextosFormulario(frmEditarApuntes)
                    ' Llamamos al formulario de manera modal en modo edición
                    vEditar = "SI"
                    frmEditarApuntes.ShowDialog()
                    frmEditarApuntes.Dispose()

                    ' =========================================================================
                    ' 🌟 EXCELENTE OPTIMIZACIÓN: ADIÓS AL CÓDIGO REDUNDANTE
                    ' =========================================================================
                    ' Disparamos tu rutina pública para refrescar el Grid con IDs y traducciones
                    RefrescarGridApuntesContables()

                    ' Reposicionamos el foco en el registro que el usuario acaba de editar
                    If DgvApuntes.Rows.Count > 0 Then
                        Dim filaEncontrada As Integer = 0
                        For Each row As DataGridViewRow In DgvApuntes.Rows
                            ' Sincronizamos contra el código identificador único
                            If CStr(row.Cells(7).Value) = vCodigo Then
                                filaEncontrada = row.Index
                                Exit For
                            End If
                        Next

                        DgvApuntes.Rows(filaEncontrada).Selected = True
                        DgvApuntes.CurrentCell = DgvApuntes.Rows(filaEncontrada).Cells(0)
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoInformeApuntes Is Nothing) OrElse (Not frmTipoInformeApuntes.IsHandleCreated)) Then
            frmTipoInformeApuntes = New TipoInformeApuntes
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmTipoInformeApuntes)
        ' Llamamos al formulario de manera modal.
        frmTipoInformeApuntes.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoInformeApuntes.Dispose()
    End Sub

    Private Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles BtnExcel.Click
        ' =========================================================================
        ' 🚀 MOTOR DE EXPORTACIÓN UNIVERSAL (COMPATIBLE CON LIBREOFFICE Y MS EXCEL)
        ' =========================================================================
        Using sfd As New SaveFileDialog()
            sfd.Filter = rmse.GetString("ArchivosDeExcel") & " (*.xlsx)|*.xlsx|" & rmse.GetString("TodosLosArchivos") & "(* .*)|*.*"
            sfd.FileName = vAñoEjercicio & "_" & rmse.GetString("LblApuntes.Text") & ".xlsx"

            ' 🎯 LA RECTIFICACIÓN MAESTRA: Leemos la ruta que tienes guardada en tu panel
            Dim rutaGuardada As String = My.Settings.PathExportar

            ' Si la variable tiene una ruta real y esa carpeta existe físicamente en el disco duro, la abrimos.
            ' Si estuviera vacía (primer arranque), usamos "Mis Documentos" como salvavidas preventivo.
            If Not String.IsNullOrEmpty(rutaGuardada) AndAlso System.IO.Directory.Exists(rutaGuardada) Then
                sfd.InitialDirectory = rutaGuardada
            Else
                sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            End If

            If sfd.ShowDialog() = DialogResult.OK Then
                Dim strFileName As String = sfd.FileName
                My.Settings.PathExportar = Path.GetDirectoryName(strFileName)
                My.Settings.Save()
                My.Settings.Reload()

                Try
                    If ((DgvApuntes.Columns.Count = 0) Or (DgvApuntes.Rows.Count = 0)) Then
                        Exit Sub
                    End If

                    Dim vNumRegistros As Integer = DgvApuntes.Rows.Count
                    PrbExport.Visible = True
                    PrbExport.Minimum = 0
                    PrbExport.Maximum = vNumRegistros
                    PrbExport.Value = 0

                    ' 1. CREACIÓN DEL DATASET EN MEMORIA
                    Dim dset As New DataSet
                    dset.Tables.Add()

                    ' Agregar Columnas especificando tipos reales para evitar el error de Texto
                    ' Restamos columnas según tu lógica original (ej. ColumnCount - 5)
                    Dim totalColumnasDataset As Integer = DgvApuntes.ColumnCount - 5
                    For i As Integer = 0 To totalColumnasDataset - 1
                        Dim col As New DataColumn(DgvApuntes.Columns(i).HeaderText)
                        If i = 3 OrElse i = 4 Then
                            col.DataType = GetType(Double) ' Tipo numérico real para Importe y Saldo
                        Else
                            col.DataType = GetType(String)
                        End If
                        dset.Tables(0).Columns.Add(col)
                    Next

                    Dim dr1 As DataRow
                    Dim vSuma As Double = 0

                    ' 2. LLENADO DE FILAS BLINDADO CONTRA CELDAS VACÍAS (DBNull)
                    If DgvApuntes.SelectedRows.Count > 1 Then
                        ' --- CASO A: SOLO FILAS SELECCIONADAS ---
                        For i As Integer = 0 To DgvApuntes.RowCount - 1
                            If DgvApuntes.Rows(i).Selected Then
                                dr1 = dset.Tables(0).NewRow
                                For j As Integer = 0 To totalColumnasDataset - 1
                                    Dim celdaValor As Object = DgvApuntes.Rows(i).Cells(j).Value

                                    If celdaValor Is Nothing OrElse IsDBNull(celdaValor) Then
                                        dr1(j) = If(j = 3 OrElse j = 4, 0.0, "")
                                        Continue For
                                    End If

                                    If j = 0 Then
                                        dr1(j) = Convert.ToDateTime(celdaValor).ToString("dd/MM/yyyy")
                                    ElseIf j = 1 Or j = 2 Or j = 6 Then
                                        dr1(j) = Trim(celdaValor.ToString())
                                    ElseIf j = 3 Then
                                        Dim valNum As Double = Convert.ToDouble(celdaValor)
                                        dr1(j) = valNum
                                        vSuma += valNum
                                    ElseIf j = 4 Then
                                        dr1(j) = vSuma
                                    ElseIf j = 5 Then
                                        Dim vNotasTmp As String = Trim(celdaValor.ToString())
                                        dr1(j) = If(String.Compare(vNotasTmp, vLetras) > 0 Or String.Compare(vNotasTmp, vNumeros) > 0, "*" & vNotasTmp, "")
                                    End If
                                Next
                                dset.Tables(0).Rows.Add(dr1)
                            End If
                        Next
                    Else
                        ' --- CASO B: TODAS LAS FILAS ---
                        For i As Integer = 0 To DgvApuntes.RowCount - 1
                            dr1 = dset.Tables(0).NewRow
                            For j As Integer = 0 To totalColumnasDataset - 1
                                Dim celdaValor As Object = DgvApuntes.Rows(i).Cells(j).Value

                                If celdaValor Is Nothing OrElse IsDBNull(celdaValor) Then
                                    dr1(j) = If(j = 3 OrElse j = 4, 0.0, "")
                                    Continue For
                                End If

                                If j = 0 Then
                                    dr1(j) = Convert.ToDateTime(celdaValor).ToString("dd/MM/yyyy")
                                ElseIf j = 1 Or j = 2 Or j = 6 Then
                                    dr1(j) = Trim(celdaValor.ToString())
                                ElseIf j = 3 Or j = 4 Then
                                    dr1(j) = Convert.ToDouble(celdaValor)
                                ElseIf j = 5 Then
                                    Dim vNotasTmp As String = Trim(celdaValor.ToString())
                                    dr1(j) = If(String.Compare(vNotasTmp, vLetras) > 0 Or String.Compare(vNotasTmp, vNumeros) > 0, "*" & vNotasTmp, "")
                                End If
                            Next
                            dset.Tables(0).Rows.Add(dr1)
                            PrbExport.Value = i
                        Next
                    End If

                    ' =========================================================================
                    ' 3. GENERACIÓN DEL ARCHIVO CON CLOSEDXML (SIN DEPENDER DE OFFICE)
                    ' =========================================================================
                    Using workbook As New ClosedXML.Excel.XLWorkbook()
                        Dim wSheet = workbook.Worksheets.Add(rmse.GetString("LblApuntes.Text"))
                        Dim dt As System.Data.DataTable = dset.Tables(0)

                        ' Volcar tabla completa de golpe (Rápido y eficiente)
                        wSheet.Cell(1, 1).InsertTable(dt)

                        ' El formato contable exacto: Positivos estándar, Negativos en ROJO
                        ' El número de columna en ClosedXML empieza en 1 (Columna D = 4, Columna E = 5)
                        Dim formatoMonedaRojo As String = "#,##0.00 €;[Red]-#,##0.00 €"
                        wSheet.Column(4).Style.NumberFormat.Format = formatoMonedaRojo
                        wSheet.Column(5).Style.NumberFormat.Format = formatoMonedaRojo

                        ' Estética de cabecera y autoajuste
                        wSheet.Row(1).Style.Font.Bold = True
                        wSheet.Columns().AdjustToContents()

                        ' Guardar físicamente el archivo binario nativo (.xlsx)
                        workbook.SaveAs(strFileName)
                    End Using

                    PrbExport.Visible = False

                    ' 🚀 EL DISPARADOR PREMIUM: Preguntamos al usuario de forma dócil
                    Dim msgOpen As String = rmse.GetString("AbrirArchivo")
                    Dim tituloExportacion As String = rmse.GetString("ExportacionCompletada")
                    If ConfirmarAccionTraducida(msgOpen, tituloExportacion) = MsgBoxResult.Yes Then
                        Try
                            ' Despierta el visor nativo de Excel de Windows usando la ruta real (.FileName)
                            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(strFileName) With {.UseShellExecute = True})
                        Catch ex As Exception
                            MsgBox(rmse.GetString("ErrorAbrirArchivo") & ": " & ex.Message, MsgBoxStyle.Critical)
                        End Try
                    End If
                Catch ex As Exception
                    PrbExport.Visible = False
                    MessageBox.Show(rmse.GetString("ErrorExportacion") & ": " & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Public Sub BtnTraspasarRegistro_Click(sender As Object, e As EventArgs) Handles BtnTraspasarRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTraspasoCuentas Is Nothing) OrElse (Not frmTraspasoCuentas.IsHandleCreated)) Then
            frmTraspasoCuentas = New TraspasoCuentas
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmTraspasoCuentas)

        ' 🌟 LA CORRECCIÓN CLAVE: Capturamos la respuesta de la ventana al cerrarse
        Dim respuesta As DialogResult = frmTraspasoCuentas.ShowDialog()

        ' Destruimos el formulario de la memoria RAM limpiamente
        frmTraspasoCuentas.Dispose()

        ' 🌟 SÓLO SI EL USUARIO GRABÓ (OK): Ejecutamos el refresco y movemos el foco
        ' Si el usuario pulsó Cancelar o cerró la x, pasamos de largo en silencio y la pantalla no se enfada
        If respuesta = DialogResult.OK Then
            RefrescarGridApuntesContables()

            If frmApuntesContables.DgvApuntes.RowCount > 0 Then
                vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1
                frmApuntesContables.DgvApuntes.Rows(vFilaActual).Selected = True
                frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFilaActual).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' 1. Llamamos al formulario de manera modal capturando la respuesta del usuario
        Dim respuesta As DialogResult = frmBuscar.ShowDialog()

        ' 🌟 CASO A: EL USUARIO SÍ BUSCÓ ALGO (Pulsó Buscar / OK)
        If respuesta = DialogResult.OK Then
            BtnSeguirBuscando.Enabled = True
            vBuscar = frmBuscar.CmbTextoBuscar.Text
            vCampo = frmBuscar.CmbCampos.SelectedIndex
            vRow = 0 ' Mantiene tu lógica de arrancar la búsqueda desde la primera fila
            If DgvApuntes.Rows.Count = 0 Then Exit Sub
            For Each row As DataGridViewRow In DgvApuntes.Rows
                If frmBuscar.ChkPrimerRegistro.Checked = True Then 'Desde el primer registro
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 1 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 2 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 3 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    ElseIf vCampo = 4 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            ' 🌟 LA CORRECCIÓN CLAVE: Usamos .ToString() y controlamos el nulo de forma segura
                            Dim textoNotas As String = ""
                            If row.Cells(5).Value IsNot Nothing AndAlso Not IsDBNull(row.Cells(5).Value) Then
                                textoNotas = row.Cells(5).Value.ToString().Trim()
                            End If
                            ' Ahora la comparación es 100% inmune a celdas vacías
                            If textoNotas.ToLower().Contains(vBuscar.ToLower()) Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        Else
                            If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRow = row.Index
                                Exit For
                            Else
                                vRow = -1
                            End If
                        End If
                    End If
                Else ' desde donde está la fila seleccionada
                    vRow = frmApuntesContables.DgvApuntes.CurrentRow.Index
                    If row.Index > vRow Then
                        If vCampo = 0 Then
                            If frmBuscar.ChkExacta.Checked = False Then
                                If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            Else
                                If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            End If
                        ElseIf vCampo = 1 Then
                            If frmBuscar.ChkExacta.Checked = False Then
                                If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            Else
                                If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            End If
                        ElseIf vCampo = 2 Then
                            If frmBuscar.ChkExacta.Checked = False Then
                                If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            Else
                                If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            End If
                        ElseIf vCampo = 3 Then
                            If frmBuscar.ChkExacta.Checked = False Then
                                If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            Else
                                If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            End If
                        ElseIf vCampo = 4 Then
                            If frmBuscar.ChkExacta.Checked = False Then
                                If CStr(row.Cells(5).Value).ToLower.Contains(vBuscar.ToLower) Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            Else
                                If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                                    row.Selected = True
                                    vRow = row.Index
                                    Exit For
                                Else
                                    vRow = -1
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            If vRow = -1 Then
                MsgBox(resManager.GetString("MsgDatos1"))
                BtnSeguirBuscando.Enabled = False
            Else
                If vCampo = 0 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(1)
                ElseIf vCampo = 3 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(2)
                ElseIf vCampo = 4 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(5)
                End If
            End If
            ' 🌟 CASO B: EL USUARIO CANCELÓ O CERRÓ LA VENTANA
        Else
            ' Desactivamos el botón de seguir buscando porque no hay una búsqueda activa
            BtnSeguirBuscando.Enabled = False
            ' =========================================================================
            ' REUBICACIÓN AL FINAL ABSOLUTO TRAS CANCELAR (Tu deseo contable)
            ' =========================================================================
            DgvApuntes.Select()
            ' Calculamos de forma dinámica el índice de la ÚLTIMA fila disponible en tu Grid
            Dim ultimaFilaViva As Integer = DgvApuntes.Rows.Count - 1
            ' El escudo de seguridad evita errores si la rejilla estuviera vacía (debe ser >= 0)
            If ultimaFilaViva >= 0 Then
                DgvApuntes.Rows(ultimaFilaViva).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(ultimaFilaViva).Cells(0)
            End If
            DgvApuntes.Refresh()
        End If
    End Sub

    Private Sub BtnAumentar_Click(sender As Object, e As EventArgs) Handles BtnAumentar.Click
        ' Cambia a tamaño 14 (puedes usar una variable para ir sumando de 2 en 2)
        CambiarTamañoFuente(Me.Controls, 14.0F)
    End Sub

    Private Sub BtnNormal_Click(sender As Object, e As EventArgs) Handles BtnNormal.Click
        CambiarTamañoFuente(Me.Controls, 10.0F)
    End Sub

    Private Sub BtnFiltroF5_Click(sender As Object, e As EventArgs) Handles BtnFiltroF5.Click
        ' Filtra Apuntes por la Descripción Seleccionada
        ' **********************************************
        If DgvApuntes.Rows.Count > 1 Then
            filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
            vTxtDescripcion = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString

            ' Llamamos al formulario de manera modal
            frmFiltroF5.ShowDialog()

            If frmFiltroF5.TxtFiltro.Text <> "" Then
                ' 🌟 SANEAMIENTO PREVENTIVO: Limpiamos la memoria de consultas anteriores
                cmdMdb1cr.Parameters.Clear()

                vTxtDescripcion = frmFiltroF5.TxtFiltro.Text

                ' Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
                Dim idConceptoSaldo As Integer = 1
                Using cmdBuscarId As New OleDb.OleDbCommand("Select IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
                    Dim resId = cmdBuscarId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
                End Using

                ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Inner Joins para nombres legibles y traducidos)
                Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"
                vtipoSql = sqlBase

                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                ' Aplicamos la búsqueda por texto mediante el LIKE (Duplicando comillas por seguridad)
                vtipoSql += " And apuntes.DescripcionAPU LIKE '%" & vTxtDescripcion.Replace("'", "''") & "%' "

                ' Lógica de filtros secundarios cruzados saneados a IDs numéricos
                Dim tieneFechasActivo As Boolean = False

                If frmFiltroF5.ChkOtrosFiltros.Checked = True And frmFiltroF5.ChkOtrosFiltros.Enabled = True Then
                    ' CORRECCIÓN: Filtro por ID de cuenta directo sin comillas
                    If BtnFiltroCuenta.Enabled = False Then
                        Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                        vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
                    End If

                    ' CORRECCIÓN: Filtro por ID de concepto directo sin comillas
                    If BtnFiltroConcepto.Enabled = False Then
                        Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                        vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
                    End If

                    ' Si las fechas están encendidas, guardamos el booleano para inyectar al final de la SQL
                    If BtnFiltroFecha.Enabled = False Then
                        tieneFechasActivo = True
                    End If
                End If

                ' 🌟 CRÍTICO: Las interrogaciones de fecha van siempre al final absoluto del WHERE lineal de OleDb
                If tieneFechasActivo Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"

                    cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                    cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
                End If

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"

                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)
            End If
        End If
    End Sub

    Private Sub BtnImportarBanco_Click(sender As Object, e As EventArgs) Handles BtnImportarBanco.Click

        ' 🎯 EL ABREPUERTAS BANCARIO: Seleccionamos el archivo Excel del BBVA u Openbank
        Try
            Using ofd As New OpenFileDialog()
                ' 1. Filtramos rígidamente para que el usuario solo pueda elegir matrices Excel limpias
                ofd.Filter = rmse.GetString("ArchivosDeExcel") & " (*.xlsx;*.xls)|*.xlsx;*.xls|" & rmse.GetString("TodosLosArchivos") & " (*.*)|*.*"
                ofd.Title = rmse.GetString("SeleccionaExtracto")

                ' =========================================================================
                ' 🎯 DIRECTO A DESCARGAS (VERSIÓN 3.2.8.0 Premium)
                ' =========================================================================
                ' 1. Calculamos la ruta biológica de la carpeta de descargas de este PC
                Dim rutaDescargasWindows As String = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")

                ' 2. Forzamos al explorador a brotar clavado allí de forma obligatoria y rígida
                If System.IO.Directory.Exists(rutaDescargasWindows) Then
                    ofd.InitialDirectory = rutaDescargasWindows
                Else
                    ' Salvavidas ultra-remoto por si el sistema operativo no tuviera la carpeta
                    ofd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                End If
                ' =========================================================================

                ' 3. Desplegamos la persiana gráfica en el monitor del usuario
                If ofd.ShowDialog() = DialogResult.OK Then
                    Dim rutaArchivoExcelBanco As String = ofd.FileName
                    ' 🚀 ADAPTACIÓN PREMIUM INTERNACIONAL UNIFICADA (Inmune a grasa digital)
                    ' Pescamos los textos traducidos desde el búnker de recursos .resx
                    Dim textoMensaje As String = rmse.GetString("ProcederaAnalizarArchivo") & ": " & System.IO.Path.GetFileName(rutaArchivoExcelBanco)
                    Dim textoTitulo As String = rmse.GetString("ImportacionBancaria")

                    ' Invocamos tu función reina para que pinte la interfaz simétrica en pantalla
                    If ConfirmarAccionTraducida(textoMensaje, textoTitulo) = MsgBoxResult.Yes Then

                        ' =========================================================================
                        ' 🚀 EXTRACCIÓN SEGURO POR SQL INDESTRUCTIBLE (VERSIÓN 3.2.8.0)
                        ' =========================================================================
                        If Not String.IsNullOrEmpty(CmbCuenta.Text) Then
                            Try
                                ' 1. Pesca del texto que el usuario ve real en su monitor (ej: "BBVA")
                                Dim nombreCuentaABuscar As String = CmbCuenta.Text.Replace("'", "''").Trim()

                                Dim textoNotas As String = ""
                                Dim idBanco As Integer = 0

                                ' 2. INTERROGATORIO AL BÚNKER DE ACCESS: Rescatamos las NotasCUE y el ID de un viaje
                                Using cmdMdb As New OleDb.OleDbCommand()
                                    cmdMdb.Connection = conexion1 ' [Ajusta a tu variable global de conexión activa]

                                    ' Buscamos la fila exacta que coincide con el texto seleccionado por el usuario
                                    cmdMdb.CommandText = "SELECT IdCuentaCUE, NotasCUE FROM cuentas WHERE NombreCUE = '" & nombreCuentaABuscar & "'"

                                    Using dr As OleDb.OleDbDataReader = cmdMdb.ExecuteReader()
                                        If dr.Read() Then
                                            textoNotas = dr("NotasCUE").ToString().Trim()
                                            idBanco = Convert.ToInt32(dr("IdCuentaCUE"))
                                        End If
                                    End Using
                                End Using

                                ' 3. 🛡️ EL ESCUDO ADUANERO: Desmenuzamos el paréntesis (6, 3, 4, 6) si todo ha ido OK
                                If idBanco > 0 AndAlso textoNotas.Contains("(") AndAlso textoNotas.Contains(")") Then

                                    ' Quitamos los paréntesis de la RAM
                                    Dim textoLimpio As String = textoNotas.Replace("(", "").Replace(")", "").Trim()
                                    Dim coordenadas() As String = textoLimpio.Split(","c)

                                    ' Convertimos los trozos en enteros numéricos puros en millonésimas de segundo
                                    Dim filaInicio As Integer = Convert.ToInt32(coordenadas(0).Trim())
                                    Dim colFecha As Integer = Convert.ToInt32(coordenadas(1).Trim())
                                    Dim colConcepto As Integer = Convert.ToInt32(coordenadas(2).Trim())
                                    Dim colImporte As Integer = Convert.ToInt32(coordenadas(3).Trim())

                                    ' 🚀 LA ESTOCADA PERFECTA: Invocamos tu función pasándole las 4 coordenadas + tu ID de Cuenta real final
                                    ProcesarMatrizBancariaManual(rutaArchivoExcelBanco, filaInicio, colFecha, colConcepto, colImporte, idBanco)
                                    'ProcesarMatrizBancariaManual(rutaArchivoExcelBanco, 5, 3, 4, 6, 1) 'BBVA c/c
                                    'ProcesarMatrizBancariaManual(rutaArchivoExcelBanco, 5, 2, 4, 5, 1) 'BBVA VISA
                                    'ProcesarMatrizBancariaManual(rutaArchivoExcelBanco, 11, 4, 6, 8, 11) 'OPENBANK
                                Else
                                    MsgBox("Aquest compte no té assignades les coordenades o el format a 'NotasCUE' és incorrecte.", MsgBoxStyle.Information, "ContaHogar")
                                End If
                            Catch ex As Exception
                                MsgBox("Error al buscar o desmenuzar las Notas de la cuenta: " & ex.Message, MsgBoxStyle.Critical)
                            End Try
                        End If
                    End If
                End If
            End Using
        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorAbrirExcelBanco") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub


    Public Sub BtnF6_Click(sender As Object, e As EventArgs) Handles BtnF6.Click
        'Vuelve a Refrecar el DataGrid y dejar los Btn de los Filtros sin Filtrar
        '************************************************************************
        vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], " &  'Celda 0
           "conceptos.DescripcionCON As [ConceptoAPU], " & ' Celda 1 (Texto visible)
           "apuntes.DescripcionAPU As [DescripcionAPU], " & ' Celda 2
           "apuntes.ImporteAPU As [ImporteAPU], " &       ' Celda 3
           "apuntes.ImporteAPU As [SaldoAPU], " &         ' Celda 4
           "apuntes.NotasAPU As [NotasAPU], " &           ' Celda 5
           "cuentas.NombreCUE As [CuentaAPU], " &         ' Celda 6 (Texto visible)
           "apuntes.CodigoAPU As [CodigoAPU], " &         ' Celda 7
           "conceptos.CodigoCON As [CodigoCON], " &       ' Celda 8 (¡CORREGIDO! Clave estable para resManager)
           "apuntes.ConceptoAPU As [IdConceptoCON], " &   ' Celda 9 (ID numérico concepto para guardar)
           "apuntes.CuentaAPU As [IdCuentaCUE] " &        ' Celda 10 (ID numérico cuenta para guardar)
           "FROM (apuntes " &
           "INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) " &
           "INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuntes)
        LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False
        BtnFiltroChekedList.Enabled = True
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        SeguirF3()
    End Sub

    Private Sub ApuntesContables_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        ' =========================================================================
        ' 🌟 ATAJO GENERAL DE TECLADO UNIFICADO (REUTILIZACIÓN TOTAL)
        ' =========================================================================

        ' 1. CONTROL DE LA TECLA F3: Continuar la búsqueda parcial de texto
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                SeguirF3()
            End If
        End If

        ' 2. CONTROL DE LA TECLA F5 (116): Filtro rápido por descripción
        ' Delegamos directamente en el botón que ya busca por IDs, LIKE y parámetros limpios
        If (e.KeyCode = Keys.F5 OrElse e.KeyCode = 116) AndAlso DgvApuntes.RowCount > 0 Then
            e.SuppressKeyPress = True ' Evitamos que Windows haga ruidos raros de sistema

            If BtnFiltroF5.Enabled Then
                BtnFiltroF5.PerformClick()
            End If
        End If

        ' 3. CONTROL DE LA TECLA F6 (117): Limpieza total y refresco general anual
        ' Delegamos en el botón que quita los filtros, apaga escudos y repinta el ejercicio
        If e.KeyCode = Keys.F6 OrElse e.KeyCode = 117 Then
            e.SuppressKeyPress = True

            If BtnF6.Enabled Then
                BtnF6.PerformClick()
            End If

            '' Cambia 'BtnSinFiltroFecha' por el botón físico que uses en tu pantalla para resetear
            'If BtnSinFiltroFecha.Enabled Then
            '    BtnSinFiltroFecha.PerformClick()
            'End If
        End If
    End Sub

    Private Sub SeguirF3()
        vCantidadFilas = DgvApuntes.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox(resManager.GetString("MsgDatos2"))
            BtnSeguirBuscando.Enabled = False
        Else
            vContador = -1
            If DgvApuntes.Rows.Count = 0 Then Exit Sub
            For Each row As DataGridViewRow In DgvApuntes.Rows
                vContador += 1
                If vContador > vRow Then
                    If vCampo = 0 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Or CStr(row.Cells(4).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Or CStr(row.Cells(4).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 1 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(0).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(0).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 2 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(1).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(1).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 3 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(2).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(2).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    ElseIf vCampo = 4 Then
                        If frmBuscar.ChkExacta.Checked = False Then
                            If CStr(row.Cells(5).Value).ToLower.Contains(vBuscar.ToLower) Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        Else
                            If CStr(row.Cells(5).Value).ToLower = vBuscar.ToLower Then
                                row.Selected = True
                                vRowSeguir = row.Index
                                Exit For
                            Else
                                vRowSeguir = -1
                            End If
                        End If
                    End If
                End If
            Next
            If vRowSeguir = -1 Then
                MsgBox(resManager.GetString("MsgDatos2"))
                BtnSeguirBuscando.Enabled = False
            Else
                vRow = vRowSeguir
                If vCampo = 0 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(1)
                ElseIf vCampo = 3 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(2)
                ElseIf vCampo = 4 Then
                    DgvApuntes.Rows(vRow).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(5)
                End If
                vRowSeguir = 0
            End If
        End If
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvApuntes.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = 0
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvApuntes.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = vFilaActual - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvApuntes.CurrentRow.Index
        If vFilaActual = DgvApuntes.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = vFilaActual + 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvApuntes.CurrentRow.Index
        If vFilaActual = DgvApuntes.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' ESCUDO: Si el formulario se está cargando, salimos inmediatamente y no hacemos nada
        If cargandoFormulario Then Exit Sub

        ' Validación de seguridad por si el combo se queda sin filas (base de datos vacía)
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        If ListBox1.SelectedItems.Count = 0 Then

            ' =========================================================================
            ' 🌟 EXTRACCIÓN ULTRA RÁPIDA DESDE MEMORIA (Cero consultas DataReader)
            ' =========================================================================
            Dim idConceptoSel As Integer = 0
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""

            If CmbConcepto.SelectedItem IsNot Nothing Then
                ' Como está enlazado a un DataTable, convertimos el ítem a DataRowView
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                idConceptoSel = Convert.ToInt32(filaSeleccionada("IdConceptoCON"))
                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString()
            End If

            ' --- TRADUCCIÓN DE LA DESCRIPCIÓN (Tu lógica original de recursos) ---
            Dim descripcionTraducida As String = descripcionOriginal

            If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(codigoOriginal) Then
                Dim traduccionDesc As String = resManager.GetString("Desc_" & codigoOriginal.Replace(" ", "_"))
                If Not String.IsNullOrEmpty(traduccionDesc) Then
                    descripcionTraducida = traduccionDesc
                End If
            End If

            ' Asignamos la descripción al cuadro de texto de al lado
            TxtConcepto.Text = descripcionTraducida

            ' =========================================================================
            ' 🌟 FILTRADO DEL GRID ADAPTADO A IDs NUMÉRICOS
            ' =========================================================================
            If BtnFiltroConcepto.Enabled = False Then
                ' Saneamos la memoria de parámetros del comando global de la app
                cmdMdb1cr.Parameters.Clear()

                ' Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
                Dim idConceptoSaldo As Integer = 1
                Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
                    Dim resId = cmdBuscarId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
                End Using

                ' Consulta SQL Maestra de 11 celdas (Garantiza ver nombres legibles en la rejilla)
                vtipoSql = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                ' CORRECCIÓN: Filtramos por el ID numérico puro del concepto, sin comillas simples
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "

                ' CORRECCIÓN: Si el filtro de cuenta está encendido, inyectamos su ID numérico real
                If BtnFiltroCuenta.Enabled = False Then
                    Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                    vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
                End If

                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"

                    ' Inyectamos los parámetros de fecha en el orden estricto de las '?'
                    cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
                    cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
                End If

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"

                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)

                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        End If
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub DgvApuntes_ColumnHeaderMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DgvApuntes.ColumnHeaderMouseClick
        ' Para que no ordene la columna(4) de "Saldo"
        frmApuntesContables.DgvApuntes.Columns.Item(4).SortMode = DataGridViewColumnSortMode.NotSortable
        DgvApuntesContables(3, 4)
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False
        BtnFiltroChekedList.Enabled = True
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            BtnFiltroCuenta.Enabled = True
            BtnSinFiltroCuenta.Enabled = False
            BtnFiltroConcepto.Enabled = True
            BtnSinFiltroConcepto.Enabled = False
            BtnFiltroFecha.Enabled = True
            BtnSinFiltroFecha.Enabled = False
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnSinFiltroCuenta, TL(1)},
            {Me.BtnSinFiltroConcepto, TL(3)},
            {Me.BtnSinFiltroFecha, TL(5)},
            {Me.BtnSeguirBuscando, TL(11)}
        }

        For Each par In botonesBloqueados
            Dim boton As Button = par.Key
            Dim tool As ToolTip = par.Value

            If Not boton.Enabled Then
                ' Traducimos la posición del ratón al contenedor nativo del botón (su GroupBox)
                Dim posRatonRelativaAlBoton As Point = boton.Parent.PointToClient(Cursor.Position)

                ' Si el ratón está sobre el botón desactivado
                If boton.Bounds.Contains(posRatonRelativaAlBoton) Then
                    ' Calculamos la posición respecto al formulario para dibujar el cartelito en el lugar correcto
                    Dim posRatonRelativaAlForm As Point = Me.PointToClient(Cursor.Position)
                    'tool.Show(resManager.GetString("ToolTipQuitarFiltro"), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso
                    Dim textoKey As String = If(boton Is Me.BtnSeguirBuscando, "ToolTipSeguirBuscando", "ToolTipQuitarFiltro")
                    tool.Show(resManager.GetString(textoKey), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón no está sobre ningún botón bloqueado, ocultamos los tres
        TL(1).Hide(Me)
        TL(3).Hide(Me)
        TL(5).Hide(Me)
        TL(11).Hide(Me)
    End Sub

    Private Sub DgvApuntes_Sorted(sender As Object, e As EventArgs) Handles DgvApuntes.Sorted
        Try
            ' 🚀 LA ESTOCADA: Cuando el usuario reordene las columnas, obligamos a la app
            ' a limpiar el acumulador viejo de la RAM y recalcular los saldos de arriba a abajo.
            ' Llamamos directamente a tu función global del módulo de forma transparente.
            DgvApuntesContables(3, 4)
        Catch ex As Exception
            ' Evita cualquier parpadeo visual en el hilo principal del formulario
        End Try
    End Sub


End Class