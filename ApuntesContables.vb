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
    Public TL(29) As ToolTip
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
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        End If
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        ' Llenar el Combo Concepto y ListBox1 utilizando la función sincronizada
        '*******************************************************************************
        ' 1. SQL adaptado para traer el IdConceptoCON y el CodigoCON que necesita la nueva lógica
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, DescripcionCON, TipoCON FROM conceptos ORDER BY TipoCON ASC, IdConceptoCON ASC"

        Try
            ' 2. Abrimos el lector original de siempre
            drMdb1 = cmdMdb1cr.ExecuteReader()
            ' 3. Encendemos tu escudo protector antes de rellenar
            cargandoFormulario = True
            ' 4. ¡CORREGIDO!: Le pasamos el drMdb1 (Reader) que la función de tu módulo espera recibir
            LlenarYTraducirControlesConceptosBD(Me.CmbConcepto, Me.ListBox1, drMdb1)
            ' 5. Apagamos el escudo tras la carga y cerramos el lector
            cargandoFormulario = False
            drMdb1.Close()
            ' FORZAMOS a que lea el tercer concepto seleccionado por defecto al abrir la pantalla
            If CmbConcepto.Items.Count > 2 Then ' Validamos que al menos haya 2 elementos
                ' Nos aseguramos de provocar el cambio de índice moviéndolo a vacío (-1) primero
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 2
            End If
        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox("Error al procesar conceptos: " & ex.Message, MsgBoxStyle.Critical)
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
            MsgBox("Error al cargar las cuentas en la pantalla: " & ex.Message, MsgBoxStyle.Critical, "Error")
        End Try

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvApuntes.Columns
            If columna.Name <> "ImporteAPU" And columna.Name <> "Expr1003" And columna.Name <> "CuentaAPU" And columna.Name <> "CodigoAPU" Then
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
                'MessageBox.Show("Aviso: No existen apuntes registrados para la cuenta seleccionada en este ejercicio.",
                'rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
        'MsgBox(e.KeyCode)
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
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnSinFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroCuenta.Click
        ' 🌟 SANEAMIENTO DE ENTRADA: Limpiamos la memoria de consultas previas
        cmdMdb1cr.Parameters.Clear()

        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False

        ' 1. Buscamos el ID numérico real del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' Guardamos un booleano para saber si las fechas están activas
        Dim tieneFechasActivo As Boolean = (BtnFiltroFecha.Enabled = False)

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Con los INNER JOIN obligatorios para ver nombres claros)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As [FechaAPU], conceptos.DescripcionCON As [ConceptoAPU], apuntes.DescripcionAPU As [DescripcionAPU], apuntes.ImporteAPU As [ImporteAPU], apuntes.ImporteAPU As [SaldoAPU], apuntes.NotasAPU As [NotasAPU], cuentas.NombreCUE As [CuentaAPU], apuntes.CodigoAPU As [CodigoAPU], conceptos.CodigoCON As [CodigoCON], apuntes.ConceptoAPU As [IdConceptoCON], apuntes.CuentaAPU As [IdCuentaCUE] FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        If ListBox1.SelectedItems.Count <> 0 Then
            TxtConcepto.Text = rmse.GetString("MsgText3")
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.Text = CmbConcepto.Items(0)
            Dim i As Integer
            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True

            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString

                ' Ignoramos las cabeceras estéticas (** GASTO **, etc.)
                If vConcepto.StartsWith("**") Then Continue For

                ' Buscamos el ID numérico del concepto del ListBox de forma aislada
                Dim idConceptoFila As Integer = 0
                Using cmdId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                    cmdId.Parameters.AddWithValue("?", vConcepto)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                End Using

                If idConceptoFila = 0 Then Continue For

                If i = 0 Then
                    ' CORRECCIÓN: Filtro por ID numérico sin comillas
                    vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "

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

                    vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "

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
            BtnFiltroChekedList.Enabled = False
            ListBox1.Visible = False
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
            ' LÓGICA DEL ELSE (CUANDO EL LISTBOX NO TIENE SELECCIONES)
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' 🌟 CORRECCIÓN CLAVE: Quitamos apuntes.CuentaAPU <> '' porque ahora es Numérico. 
            ' En su lugar, indicamos que traiga cualquier ID de cuenta válido (mayor que 0)
            vtipoSql += " And apuntes.CuentaAPU > 0 "

            ' CORRECCIÓN: Si el filtro de concepto está encendido, extraemos su ID numérico real
            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

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

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' 🌟 SANEAMIENTO INICIAL: Vaciamos cualquier rastro de parámetros anteriores en la app
        cmdMdb1cr.Parameters.Clear()

        ' 1. Activamos el escudo temporal para que el combo no dispare eventos a destiempo
        cargandoFormulario = True

        ' 2. LLENAR Y TRADUCIR CON NUESTRA FUNCIÓN MODULAR (Añadido el IdConceptoCON obligatorio)
        cmdMdb1cr.CommandText = "SELECT IdConceptoCON, CodigoCON, TipoCON, DescripcionCON FROM conceptos ORDER BY TipoCON ASC, CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()

            ' Se encarga de limpiar, rellenar, agrupar y traducir CmbConcepto y ListBox1
            LlenarYTraducirControlesConceptosBD(Me.CmbConcepto, Me.ListBox1, drMdb1)

            ' El lector ya se cierra de forma obligatoria y segura dentro de LlenarYTraducirControlesConceptosBD

            ' 3. Selección segura del tercer elemento (Índice 2) o del primero como tenías planeado
            If CmbConcepto.Items.Count > 2 Then
                CmbConcepto.SelectedIndex = 2
                vConcepto = CmbConcepto.Text.ToString()
            ElseIf CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = 0
                vConcepto = CmbConcepto.Text.ToString()
            Else
                vConcepto = ""
            End If

            ' 4. NUEVO Y ULTRA RÁPIDO: Rellenar el TxtConcepto sin abrir un segundo lector repetido
            ' Como CmbConcepto está enlazado a un DataTable gracias al DataSource, extraemos la descripción al vuelo
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)
                TxtConcepto.Text = filaSeleccionada("DescripcionCON").ToString()
            Else
                TxtConcepto.Text = ""
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
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
        ' 🌟 SANEAMIENTO INICIAL: Vaciamos cualquier rastro de parámetros anteriores
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

        ' 3. Asignamos las fechas por defecto a los DateTimePicker
        DateTimePicker1.Value = New Date(anio, 1, 1)
        DateTimePicker2.Value = New Date(anio, 12, 31)

        ' 4. Tu lógica de botones original
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False

        ' 5. Buscamos el ID numérico del concepto "SALDO" de forma segura y aislada
        Dim idConceptoSaldo As Integer = 1
        Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
            Dim resId = cmdBuscarId.ExecuteScalar()
            If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
        End Using

        ' 🌟 CONSULTA SQL MAESTRA DE 11 CELDAS (Saneada con alias limpios para que LlenarGrid los reconozca al vuelo)
        Dim sqlBase As String = "SELECT apuntes.FechaAPU As FechaAPU, conceptos.DescripcionCON As ConceptoAPU, apuntes.DescripcionAPU As DescripcionAPU, apuntes.ImporteAPU As ImporteAPU, apuntes.ImporteAPU As SaldoAPU, apuntes.NotasAPU As NotasAPU, cuentas.NombreCUE As CuentaAPU, apuntes.CodigoAPU As CodigoAPU, conceptos.CodigoCON As CodigoCON, apuntes.ConceptoAPU As IdConceptoCON, apuntes.CuentaAPU As IdCuentaCUE FROM (apuntes INNER JOIN conceptos ON apuntes.ConceptoAPU = conceptos.IdConceptoCON) INNER JOIN cuentas ON apuntes.CuentaAPU = cuentas.IdCuentaCUE"

        If ListBox1.SelectedItems.Count <> 0 Then
            TxtConcepto.Text = rmse.GetString("MsgText3")
            CmbConcepto.DataSource = Nothing
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.Text = CmbConcepto.Items(0)

            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True

            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            ' 🌟 RECOLECCIÓN DE IDs INMUNE AL IDIOMA SELECCIONADO
            Dim listaIdsConceptos As New List(Of Integer)

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim vConceptoReal As String = ListBox1.SelectedItems(i).ToString().Trim()
                If vConceptoReal.StartsWith("**") Then Continue For

                Dim idConceptoFila As Integer = 0

                ' 🌟 ESCUDO DE IDIOMA: Buscamos el ID ya sea emparejando por Código corto o por la Descripción larga de la BD
                Dim sqlBuscarConcepto As String = "SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ? Or DescripcionCON = ?"
                Using cmdId As New OleDb.OleDbCommand(sqlBuscarConcepto, conexion1)
                    cmdId.Parameters.Clear()
                    cmdId.Parameters.AddWithValue("?", vConceptoReal)
                    cmdId.Parameters.AddWithValue("?", vConceptoReal)
                    Dim resId = cmdId.ExecuteScalar()
                    If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoFila = Convert.ToInt32(resId)
                End Using

                ' Plan B: Si tu rutina de traducción alteró mucho el texto (ej: BALANCE), buscamos en el resManager la inversa
                If idConceptoFila = 0 AndAlso resManager IsNot Nothing Then
                    ' Recorremos la tabla conceptos en memoria rápido para emparejar la traducción actual
                    Using cmdAux As New OleDb.OleDbCommand("SELECT IdConceptoCON, CodigoCON, DescripcionCON FROM conceptos", conexion1)
                        Using drAux As OleDb.OleDbDataReader = cmdAux.ExecuteReader()
                            While drAux.Read()
                                Dim codOriginal As String = drAux("CodigoCON").ToString().Trim()
                                Dim clave As String = codOriginal.Replace(" ", "_")
                                Dim trad As String = resManager.GetString(clave)

                                If (Not String.IsNullOrEmpty(trad) AndAlso trad.ToUpper() = vConceptoReal.ToUpper()) OrElse codOriginal.ToUpper() = vConceptoReal.ToUpper() Then
                                    idConceptoFila = Convert.ToInt32(drAux("IdConceptoCON"))
                                    Exit While
                                End If
                            End While
                        End Using
                    End Using
                End If

                If idConceptoFila > 0 Then
                    listaIdsConceptos.Add(idConceptoFila)
                End If
            Next

            ' 🌟 INYECCIÓN DE SEGURIDAD CONTRA ARRASTRES VACÍOS: 
            ' Si por un fallo de traducción la lista está vacía, le inyectamos un ID 0 para que no rompa la sintaxis del IN()
            If listaIdsConceptos.Count = 0 Then listaIdsConceptos.Add(0)

            ' Inyectamos la lista en el filtro IN numérico perfecto
            vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaIdsConceptos) & ") "

            ' Si el filtro de cuenta está encendido, añadimos su ID numérico real
            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"

            ' Vaciamos parámetros para que LlenarGrid reciba el comando limpio
            cmdMdb1cr.Parameters.Clear()

            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        Else
            ' =========================================================================
            ' LÓGICA DEL ELSE (SIN LISTBOX FILTRADO)
            ' =========================================================================
            vtipoSql = sqlBase
            If BtnFechasClick = "SI" Then
                vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If

            If BtnFiltroCuenta.Enabled = False Then
                Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
                vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSel} "
            End If
            If BtnFiltroConcepto.Enabled = False Then
                Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoSel} "
            End If

            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"

            cmdMdb1cr.Parameters.Clear()

            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)

            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
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

            'MessageBox.Show("Aviso: No existen apuntes registrados para la cuenta seleccionada en este ejercicio.",
            'rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
        If ListBox1.SelectedItems.Count = 0 Then
            If BtnFiltroFecha.Enabled = False Then
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuntes.FechaAPU >= ?"
                vtipoSql += " And apuntes.FechaAPU <= ?"
                If BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
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
            Else
                'MsgBox("BtnFiltroFecha.Enabled = " & BtnFiltroFecha.Enabled.ToString)
            End If
        Else
            If BtnFiltroFecha.Enabled = False Then
                Dim i As Integer
                vtipoSql = "Select apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To ListBox1.SelectedItems.Count - 1
                    vConcepto = ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
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

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        If ListBox1.SelectedItems.Count = 0 Then
            If BtnFiltroFecha.Enabled = False Then
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuntes.FechaAPU >= ?"
                vtipoSql += " And apuntes.FechaAPU <= ?"
                If BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                'MsgBox(vtipoSql)
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)
                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        Else
            If BtnFiltroFecha.Enabled = False Then
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To ListBox1.SelectedItems.Count - 1
                    vConcepto = ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
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

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        frmPrincipal.TsLabelFormulario.Text = rmse.GetString("MsgText2")
        ' 1. Controlar la instancia física del formulario de forma tradicional
        If ((frmIntroApuntes Is Nothing) OrElse (Not frmIntroApuntes.IsHandleCreated)) Then
            frmIntroApuntes = New IntroApuntes
        End If
        ' 2. Forzar la traducción antes de mostrarlo (por consistencia con tus otros formularios)
        ' Nota: Si tienes el método ActualizarTextosFormulario accesible, puedes llamarlo aquí:
        ' ActualizarTextosFormulario(frmIntroApuntes)
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
        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        vTxtNombre = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString

        ' Buscamos la traducción y la pasamos a Mayúsculas
        Dim palabraSaldoMayusculas As String = resManager.GetString("Saldo")?.ToUpper()
        If String.IsNullOrEmpty(palabraSaldoMayusculas) Then palabraSaldoMayusculas = "SALDO"

        ' Comparación totalmente segura
        If vTxtNombre.ToUpper() = palabraSaldoMayusculas Then
            MsgBox(rmse.GetString("MsgSaldos1"))
        Else
            ' Comprobamos si existe un identificador asociado.
            If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
                frmEditarApuntes = New EditarApuntes
            End If
            ' Llamamos al formulario de manera modal.
            vEditar = "NO"  ' Eliminar
            frmEditarApuntes.ShowDialog()
            frmEditarApuntes.Dispose()

            ' =========================================================================
            ' 🌟 CORRECCIÓN CRÍTICA: BUSQUEDAS EN LÍNEA INDEPENDIENTES (Cierran el DataReader)
            ' =========================================================================
            Dim idConceptoSaldo As Integer = 1
            ' Usamos una estructura limpia con un comando local exclusivo
            Using cmdBuscarId As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = 'SALDO'", conexion1)
                Dim resId = cmdBuscarId.ExecuteScalar() ' ExecuteScalar lee el dato y cierra el flujo al instante
                If resId IsNot Nothing AndAlso Not IsDBNull(resId) Then idConceptoSaldo = Convert.ToInt32(resId)
            End Using

            Dim idCuentaSeleccionada As Integer = 0
            If frmApuntesContables.BtnFiltroCuenta.Enabled = False AndAlso frmApuntesContables.CmbCuenta.SelectedItem IsNot Nothing Then
                Using cmdCta As New OleDb.OleDbCommand("SELECT IdCuentaCUE FROM cuentas WHERE NombreCUE = ?", conexion1)
                    cmdCta.Parameters.AddWithValue("?", frmApuntesContables.CmbCuenta.Text)
                    Dim resCta = cmdCta.ExecuteScalar()
                    If resCta IsNot Nothing AndAlso Not IsDBNull(resCta) Then idCuentaSeleccionada = Convert.ToInt32(resCta)
                End Using
            End If

            ' Reutilizamos tu lógica base de la consulta de refresco
            Dim sqlBase As String = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"

            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
                vtipoSql = sqlBase
                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSeleccionada} "
                End If

                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
                End If

                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    vDate1 = frmApuntesContables.DateTimePicker1.Value
                    vDate2 = frmApuntesContables.DateTimePicker2.Value
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
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
            Else
                Dim i As Integer
                vtipoSql = sqlBase
                If BtnFechasClick = "SI" Then
                    vtipoSql += $" WHERE apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                    vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString

                    Dim idConceptoFila As Integer = 0
                    ' Usamos otro comando local aislado para el bucle del ListBox
                    Using cmdCpt As New OleDb.OleDbCommand("SELECT IdConceptoCON FROM conceptos WHERE CodigoCON = ?", conexion1)
                        cmdCpt.Parameters.AddWithValue("?", vConcepto)
                        Dim resCpt = cmdCpt.ExecuteScalar()
                        If resCpt IsNot Nothing AndAlso Not IsDBNull(resCpt) Then idConceptoFila = Convert.ToInt32(resCpt)
                    End Using

                    If i = 0 Then
                        vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSeleccionada} "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value
                            vDate2 = frmApuntesContables.DateTimePicker2.Value
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += $" apuntes.ConceptoAPU <> {idConceptoSaldo} And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += " apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If

                        vtipoSql += $" And apuntes.ConceptoAPU = {idConceptoFila} "

                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += $" And apuntes.CuentaAPU = {idCuentaSeleccionada} "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value
                            vDate2 = frmApuntesContables.DateTimePicker2.Value
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                Next

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)

                If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        End If
    End Sub


    'Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
    '    filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
    '    vTxtNombre = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString

    '    Buscamos la traducción y la pasamos a Mayúsculas
    '    Dim palabraSaldoMayusculas As String = resManager.GetString("Saldo")?.ToUpper()
    '    If String.IsNullOrEmpty(palabraSaldoMayusculas) Then palabraSaldoMayusculas = "SALDO"

    '    Comparación totalmente segura
    '    If vTxtNombre.ToUpper() = palabraSaldoMayusculas Then
    '        MsgBox(rmse.GetString("MsgSaldos1"))
    '    Else
    '        Comprobamos si existe un identificador asociado.
    '        If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
    '            frmEditarApuntes = New EditarApuntes
    '        End If
    '        Llamamos al formulario de manera modal.
    '        vEditar = "NO"  ' Eliminar
    '        frmEditarApuntes.ShowDialog()
    '        MessageBox.Show("Se ha cerrado el formulario.")
    '        Destruimos el formulario.
    '        frmEditarApuntes.Dispose()
    '        Stop
    '        If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
    '            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
    '            If BtnFechasClick = "SI" Then
    '                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
    '            Else
    '                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '            End If
    '            If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
    '                vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
    '            End If
    '            If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
    '                vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
    '            End If
    '            If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
    '                vDate1 = frmApuntesContables.DateTimePicker1.Value
    '                vDate2 = frmApuntesContables.DateTimePicker2.Value
    '                vtipoSql += " And apuntes.FechaAPU >= ?"
    '                vtipoSql += " And apuntes.FechaAPU <= ?"
    '            End If
    '            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
    '            vtipoGrid = "APUNTES_CONTABLES"
    '            LlenarGrid(vtipoSql, vtipoGrid, "1")
    '            TraducirGridApuntesBD(Me.DgvApuntes)
    '            If DgvApuntes.RowCount - 1 >= 0 Then
    '                vFila = DgvApuntes.RowCount - 1
    '                DgvApuntes.Rows(vFila).Selected = True
    '                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
    '            End If
    '        Else
    '            Dim i As Integer
    '            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
    '            If BtnFechasClick = "SI" Then
    '                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
    '            Else
    '                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '            End If
    '            For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
    '                vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString
    '                If i = 0 Then
    '                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
    '                    If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
    '                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
    '                    End If
    '                    If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
    '                        vDate1 = frmApuntesContables.DateTimePicker1.Value
    '                        vDate2 = frmApuntesContables.DateTimePicker2.Value
    '                        vtipoSql += " And apuntes.FechaAPU >= ?"
    '                        vtipoSql += " And apuntes.FechaAPU <= ?"
    '                    End If
    '                Else
    '                    vtipoSql += " Or "
    '                    If BtnFechasClick = "SI" Then
    '                        vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
    '                    Else
    '                        vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '                    End If
    '                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
    '                    If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
    '                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
    '                    End If
    '                    If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
    '                        vDate1 = frmApuntesContables.DateTimePicker1.Value
    '                        vDate2 = frmApuntesContables.DateTimePicker2.Value
    '                        vtipoSql += " And apuntes.FechaAPU >= ?"
    '                        vtipoSql += " And apuntes.FechaAPU <= ?"
    '                    End If
    '                End If
    '            Next
    '            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
    '            vtipoGrid = "APUNTES_CONTABLES"
    '            LlenarGrid(vtipoSql, vtipoGrid, "1")
    '            TraducirGridApuntesBD(Me.DgvApuntes)
    '            If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
    '                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
    '                frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
    '                frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
    '            End If
    '        End If
    '    End If
    'End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoGrafico Is Nothing) OrElse (Not frmTipoGrafico.IsHandleCreated)) Then
            frmTipoGrafico = New TipoGrafico
        End If
        ' Llamamos al formulario de manera modal.
        frmTipoGrafico.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoGrafico.Dispose()
    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        'Elimina las Filas Seleccionadas
        '*******************************
        For Each r As DataGridViewRow In DgvApuntes.SelectedRows
            If DgvApuntes.Rows.Count > 1 Then
                DgvApuntes.Rows.Remove(r)
            End If
        Next
        If DgvApuntes.Rows.Count > 1 Then
            filaSelec = DgvApuntes.CurrentRow.Index
            For i = 0 To DgvApuntes.Rows.Count - 1
                DgvApuntes.Rows(i).Selected = False
            Next
            'Variable que guardara el valor
            'Dim iTotal As Integer = Me.DgvApuntes.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
            'Definimos la variable i para controlar el ciclo for
            'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
            DgvApuntesContables(3, 4)
            DgvApuntes.Select()
            DgvApuntes.CurrentRow.Selected = True
            DgvApuntes.Refresh()
        End If
    End Sub

    Private Sub DgvApuntes_KeyDown(sender As Object, e As KeyEventArgs) Handles DgvApuntes.KeyDown
        'MsgBox(e.KeyCode)
        'Elimina las Filas Seleccionadas
        '*******************************
        If e.KeyCode = 46 Then  'Tecla Supr
            For Each r As DataGridViewRow In DgvApuntes.SelectedRows
                If DgvApuntes.Rows.Count > 1 Then
                    DgvApuntes.Rows.Remove(r)
                End If
            Next
            If DgvApuntes.Rows.Count > 1 Then
                filaSelec = DgvApuntes.CurrentRow.Index
                For i = 0 To DgvApuntes.Rows.Count - 1
                    DgvApuntes.Rows(i).Selected = False
                Next
                'Variable que guardara el valor
                'Dim iTotal As Integer = Me.DgvApuntes.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
                'Definimos la variable i para controlar el ciclo for
                'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
                DgvApuntesContables(3, 4)
                DgvApuntes.Select()
                DgvApuntes.CurrentRow.Selected = True
                DgvApuntes.Refresh()
            End If
        End If
    End Sub

    Private Sub BtnFechas_Click(sender As Object, e As EventArgs) Handles BtnFechas.Click
        If BtnFechasClick = "NO" Then
            BtnFechasClick = "SI"
            BtnFechasFondo.Visible = True
            cmdMdb1cr.CommandText = "SELECT * FROM ejercicios ORDER BY ejercicios.EjercicioEJE DESC"
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        ' 1. Verificamos que el valor de la base de datos no sea NULL
                        If Not drMdb1.IsDBNull(0) Then
                            Dim valorDb As String = drMdb1.GetValue(0).ToString()
                            Dim anio As Integer

                            ' 2. Convertimos el texto a número entero de forma segura
                            If Integer.TryParse(valorDb, anio) Then
                                vFecha1Enero = anio
                            Else
                                ' Si no es un número válido (ej. texto), asignamos año actual por seguridad
                                vFecha1Enero = Date.Today.Year
                            End If
                        Else
                            ' Si el campo en la base de datos es NULL, asignamos año actual
                            vFecha1Enero = Date.Today.Year
                        End If
                    End While
                Else
                    'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                End If
                drMdb1.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
            DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
            DateTimePicker2.MinDate = New Date(vFecha1Enero, 1, 1)
        Else
            ' 1. Convertimos el año a número entero de forma segura
            Dim anio As Integer
            If Not Integer.TryParse(vAñoEjercicio, anio) Then
                ' Salvavidas: si falla o está vacío, usa el año actual
                anio = Date.Today.Year
            End If

            ' 2. Asignamos el año a tus variables del formulario
            vFecha1Enero = anio
            vFecha31Diciembre = anio

            ' 3. Creamos los objetos de fecha una sola vez para mejorar el rendimiento
            Dim fechaInicio As New Date(anio, 1, 1)
            Dim fechaFin As New Date(anio, 12, 31)

            ' 4. Configuramos el primer DateTimePicker
            DateTimePicker1.MinDate = fechaInicio
            DateTimePicker1.MaxDate = fechaFin
            DateTimePicker1.Value = fechaInicio

            ' 5. Configuramos el segundo DateTimePicker
            DateTimePicker2.MinDate = fechaInicio
            DateTimePicker2.MaxDate = fechaFin
            DateTimePicker2.Value = fechaFin

            ' 6. Tu lógica original de la interfaz
            BtnFechasClick = "NO"
            BtnFechasFondo.Visible = False
        End If
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        vTxtNombre = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString
        vCodigo = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(7).Value.ToString

        If vTxtNombre = "SALDO" Then
            MsgBox(rmse.GetString("MsgSaldos2"))
        Else
            ' Comprobamos si existe un identificador asociado.
            If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
                frmEditarApuntes = New EditarApuntes
            End If
            ' Llamamos al formulario de manera modal.
            vEditar = "SI"
            frmEditarApuntes.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmEditarApuntes.Dispose()
            If ListBox1.SelectedItems.Count <> 0 Then
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To ListBox1.SelectedItems.Count - 1
                    vConcepto = ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                Next
            Else
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                If BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)
            If DgvApuntes.Rows.Count <> 0 Then
                For Each row As DataGridViewRow In DgvApuntes.Rows
                    If CStr(row.Cells(7).Value) = vCodigo Then
                        vRow = row.Index
                        Exit For
                    Else
                        vRow = row.Index
                    End If
                Next
                DgvApuntes.Rows(vRow).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
            End If
        End If
    End Sub

    Private Sub DgvApuntes_DoubleClick(sender As Object, e As EventArgs) Handles DgvApuntes.DoubleClick
        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        vTxtNombre = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString
        vCodigo = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(7).Value.ToString

        If vTxtNombre = "SALDO" Then
            MsgBox(rmse.GetString("MsgSaldos2"))
        Else
            ' Comprobamos si existe un identificador asociado.
            If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
                frmEditarApuntes = New EditarApuntes
            End If
            ' Llamamos al formulario de manera modal.
            vEditar = "SI"
            frmEditarApuntes.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmEditarApuntes.Dispose()
            If ListBox1.SelectedItems.Count <> 0 Then
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To ListBox1.SelectedItems.Count - 1
                    vConcepto = ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    Else
                        vtipoSql += " Or "
                        If BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vDate1 = DateTimePicker1.Value.Date
                            vDate2 = DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                Next
            Else
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                If BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
                End If
                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)
            If DgvApuntes.Rows.Count <> 0 Then
                For Each row As DataGridViewRow In DgvApuntes.Rows
                    If CStr(row.Cells(7).Value) = vCodigo Then
                        vRow = row.Index
                        Exit For
                    Else
                        vRow = row.Index
                    End If
                Next
                DgvApuntes.Rows(vRow).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vRow).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoInformeApuntes Is Nothing) OrElse (Not frmTipoInformeApuntes.IsHandleCreated)) Then
            frmTipoInformeApuntes = New TipoInformeApuntes
        End If
        ' Llamamos al formulario de manera modal.
        frmTipoInformeApuntes.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoInformeApuntes.Dispose()
    End Sub

    Private Sub BtnExcel_Click(sender As Object, e As EventArgs) Handles BtnExcel.Click
        ' 1. SINCRONIZACIÓN DE LA RUTA
        ' Si la variable viene vacía o apunta a la raíz por defecto, forzamos la ruta estándar de la app
        If vPathExportar = "" OrElse vPathExportar = "C:\" OrElse vPathExportar Is Nothing Then
            vPathExportar = "C:\ContaHogar3.0\Excel"
        End If

        ' Guardamos la ruta definitiva (sea la estándar o la que el usuario cambió en Preferencias)
        My.Settings.PathExportar = vPathExportar
        My.Settings.Save()
        My.Settings.Reload()

        ' 2. VERIFICACIÓN Y CREACIÓN FÍSICA EN EL DISCO
        Try
            ' Si la carpeta (personalizada o por defecto) no existe en el disco, la creamos
            If Not Directory.Exists(My.Settings.PathExportar) Then
                Directory.CreateDirectory(My.Settings.PathExportar)
                ' Solo avisamos si la carpeta es nueva
                MsgBox(rmse.GetString("RutaExcelCreada"), MsgBoxStyle.Information, rmse.GetString("$this.Text"))
            End If
        Catch ex As Exception
            ' Si la ruta de Preferencias apunta a un pendrive desconectado o carpeta sin permisos, detenemos el proceso
            MessageBox.Show(rmse.GetString("ErrorCrearRuta") & " " & ex.Message, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        Try
            If ((DgvApuntes.Columns.Count = 0) Or (DgvApuntes.Rows.Count = 0)) Then
                Exit Sub
            End If
            Dim vNumRegistros As Integer = DgvApuntes.Rows.Count
            PrbExport.Visible = True
            PrbExport.Minimum = 0
            PrbExport.Maximum = vNumRegistros
            PrbExport.Value = 0

            'Creando Dataset para Exportar
            Dim dset As New DataSet
            'Agregar tabla al Dataset
            dset.Tables.Add()

            ' AGregar Columna a la tabla especificando tipos de datos reales
            For i As Integer = 0 To DgvApuntes.ColumnCount - 2
                Dim col As New DataColumn(DgvApuntes.Columns(i).HeaderText)
                ' Si es la columna de Importe (3) o Saldo (4), configuramos su tipo como Double
                If i = 3 OrElse i = 4 Then
                    col.DataType = GetType(Double)
                Else
                    col.DataType = GetType(String)
                End If
                dset.Tables(0).Columns.Add(col)
            Next

            'Agregar filas a la tabla
            Dim dr1 As DataRow
            Dim vSuma As Double = 0

            If DgvApuntes.SelectedRows.Count > 1 Then 'Si hay filas seleccionadas, se exportan solo las filas seleccionadas
                For i As Integer = 0 To DgvApuntes.RowCount - 1
                    If DgvApuntes.Rows(i).Selected Then
                        dr1 = dset.Tables(0).NewRow
                        For j As Integer = 0 To DgvApuntes.Columns.Count - 1
                            If j = 0 Then
                                Dim fechaCelda As DateTime = Convert.ToDateTime(DgvApuntes.Rows(i).Cells(j).Value)
                                dr1(j) = fechaCelda.ToString("yyyy'/'MM'/'dd")
                            ElseIf j = 1 Or j = 2 Or j = 6 Then
                                dr1(j) = Trim(Convert.ToString(DgvApuntes.Rows(i).Cells(j).Value))
                            ElseIf j = 3 Then
                                ' Guardamos el valor numérico puro
                                Dim valNum As Double = Convert.ToDouble(DgvApuntes.Rows(i).Cells(j).Value)
                                dr1(j) = valNum
                                vSuma = vSuma + valNum
                            ElseIf j = 4 Then
                                ' Guardamos el acumulado numérico puro
                                dr1(j) = vSuma
                            ElseIf j = 5 Then
                                vNotas = Trim(Convert.ToString(DgvApuntes.Rows(i).Cells(j).Value))
                                If Not String.IsNullOrEmpty(vNotas) Then
                                    dr1(j) = "*" & vNotas
                                Else
                                    dr1(j) = ""
                                End If
                            Else
                            End If
                        Next
                        dset.Tables(0).Rows.Add(dr1)
                    End If
                Next
            Else 'Si no hay filas seleccionadas, se exportan todas las filas
                For i As Integer = 0 To DgvApuntes.RowCount - 1
                    dr1 = dset.Tables(0).NewRow
                    For j As Integer = 0 To DgvApuntes.Columns.Count - 1
                        If j = 0 Then
                            Dim fechaCelda As DateTime = Convert.ToDateTime(DgvApuntes.Rows(i).Cells(j).Value)
                            dr1(j) = fechaCelda.ToString("yyyy'/'MM'/'dd")
                        ElseIf j = 1 Or j = 2 Or j = 6 Then
                            dr1(j) = Trim(Convert.ToString(DgvApuntes.Rows(i).Cells(j).Value))
                        ElseIf j = 3 Or j = 4 Then
                            ' Guardamos el valor numérico puro de la celda
                            dr1(j) = Convert.ToDouble(DgvApuntes.Rows(i).Cells(j).Value)
                        ElseIf j = 5 Then
                            vNotas = Trim(Convert.ToString(DgvApuntes.Rows(i).Cells(j).Value))
                            If Not String.IsNullOrEmpty(vNotas) Then
                                dr1(j) = "*" & vNotas
                            Else
                                dr1(j) = ""
                            End If
                        Else
                        End If
                    Next
                    dset.Tables(0).Rows.Add(dr1)
                Next
            End If

            Dim aplicacion As New Microsoft.Office.Interop.Excel.Application
            Dim wBook As Microsoft.Office.Interop.Excel.Workbook
            Dim wSheet As Microsoft.Office.Interop.Excel.Worksheet

            wBook = aplicacion.Workbooks.Add()
            wSheet = wBook.ActiveSheet()

            Dim dt As System.Data.DataTable = dset.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                aplicacion.Cells(1, colIndex) = dc.ColumnName
            Next

            For Each dr In dt.Rows
                PrbExport.Value = rowIndex
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    aplicacion.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next

            ' Configurar con negrilla la cabecera y tenga autofit
            wSheet.Rows.Item(1).Font.Bold = 1
            wSheet.Columns.AutoFit()

            ' 3. APLICACIÓN DEL FORMATO CONTABLE CON NEGATIVOS EN ROJO DIRECTAMENTE EN EXCEL
            ' Seleccionamos los rangos desde la fila 2 hasta la última fila escrita para las columnas D (3) y E (4)
            ' El formato "#,##0.00 €;[Red]-#,##0.00 €" define: Positivos estándar; Negativos en ROJO con signo menos.
            Dim formatoMonedaRojo As String = "#,##0.00 €;[Red]-#,##0.00 " & vMoneda
            wSheet.Range("D2", "D" & (rowIndex + 1)).NumberFormat = formatoMonedaRojo
            wSheet.Range("E2", "E" & (rowIndex + 1)).NumberFormat = formatoMonedaRojo

            Dim strFileName As String = My.Settings.PathExportar & "\" & vAñoEjercicio & "_" & rmse.GetString("LblApuntes.Text") & ".xlsx"
            Dim blnFileOpen As Boolean = False
            Try
                Dim fileTemp As System.IO.FileStream = System.IO.File.OpenWrite(strFileName)
                fileTemp.Close()
            Catch ex As Exception
                blnFileOpen = False
            End Try

            If System.IO.File.Exists(strFileName) Then
                System.IO.File.Delete(strFileName)
            End If

            wBook.SaveAs(strFileName)
            aplicacion.Workbooks.Open(strFileName)
            aplicacion.Visible = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            MsgBox(ex.ToString)
        End Try
        PrbExport.Visible = False
    End Sub

    Private Sub BtnTraspasarRegistro_Click(sender As Object, e As EventArgs) Handles BtnTraspasarRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTraspasoCuentas Is Nothing) OrElse (Not frmTraspasoCuentas.IsHandleCreated)) Then
            frmTraspasoCuentas = New TraspasoCuentas
        End If
        ' Llamamos al formulario de manera modal.
        frmTraspasoCuentas.ShowDialog()
        vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTraspasoCuentas.Dispose()
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
        If BtnFechasClick = "SI" Then
            vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If
        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
        End If
        If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
            vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
        End If
        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
            vDate1 = frmApuntesContables.DateTimePicker1.Value
            vDate2 = frmApuntesContables.DateTimePicker2.Value
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"
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

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        vRow = 0
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
    End Sub

    Private Sub BtnAumentar_Click(sender As Object, e As EventArgs) Handles BtnAumentar.Click
        ' Cambia a tamaño 14 (puedes usar una variable para ir sumando de 2 en 2)
        CambiarTamañoFuente(Me.Controls, 14.0F)
    End Sub

    Private Sub BtnNormal_Click(sender As Object, e As EventArgs) Handles BtnNormal.Click
        CambiarTamañoFuente(Me.Controls, 10.0F)
    End Sub

    Private Sub BtnFiltroF5_Click(sender As Object, e As EventArgs) Handles BtnFiltroF5.Click
        'Filtra Apuntes por la Descripción Seleccionada
        '**********************************************
        If DgvApuntes.Rows.Count > 1 Then
            filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
            vTxtDescripcion = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
            ' Llamamos al formulario de manera modal.
            frmFiltroF5.ShowDialog()
            If frmFiltroF5.TxtFiltro.Text <> "" Then
                vTxtDescripcion = frmFiltroF5.TxtFiltro.Text
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                vtipoSql += " And apuntes.DescripcionAPU LIKE '%" & vTxtDescripcion & "%' "
                If frmFiltroF5.ChkOtrosFiltros.Checked = True And frmFiltroF5.ChkOtrosFiltros.Enabled = True Then
                    If BtnFiltroCuenta.Enabled = False Then
                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                    End If
                    If BtnFiltroConcepto.Enabled = False Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
                    End If
                    If BtnFiltroFecha.Enabled = False Then
                        vtipoSql += " And apuntes.FechaAPU >= ?"
                        vtipoSql += " And apuntes.FechaAPU <= ?"
                    End If
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirGridApuntesBD(Me.DgvApuntes)
            End If
        End If
    End Sub

    Private Sub BtnF6_Click(sender As Object, e As EventArgs) Handles BtnF6.Click
        'Vuelve a Refrecar el DataGrid y dejar los Btn de los Filtros sin Filtrar
        '************************************************************************
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
        If BtnFechasClick = "SI" Then
            vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If
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
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                SeguirF3()
            End If
        End If

        If e.KeyCode = 116 And frmApuntesContables.DgvApuntes.RowCount > 0 Then 'Tecla F5 y con Filas Existentes
            'Filtra Apuntes por la Descripción Seleccionada
            '**********************************************
            If DgvApuntes.Rows.Count > 1 Then
                filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
                vTxtDescripcion = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
                ' Llamamos al formulario de manera modal.
                frmFiltroF5.ShowDialog()
                If frmFiltroF5.TxtFiltro.Text <> "" Then
                    vTxtDescripcion = frmFiltroF5.TxtFiltro.Text
                    vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                    If BtnFechasClick = "SI" Then
                        vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                    Else
                        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    End If
                    vtipoSql += " And apuntes.DescripcionAPU LIKE '%" & vTxtDescripcion & "%' "
                    If frmFiltroF5.ChkOtrosFiltros.Checked = True And frmFiltroF5.ChkOtrosFiltros.Enabled = True Then
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroConcepto.Enabled = False Then
                            vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vtipoSql += " And apuntes.FechaAPU >= ?"
                            vtipoSql += " And apuntes.FechaAPU <= ?"
                        End If
                    End If
                    vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                    vtipoGrid = "APUNTES_CONTABLES"
                    LlenarGrid(vtipoSql, vtipoGrid, "1")
                    TraducirGridApuntesBD(Me.DgvApuntes)
                End If
            End If
        End If

        If e.KeyCode = 117 Then 'Tecla F6
            'Vuelve a Refrecar el DataGrid y dejar los Btn de los Filtros sin Filtrar
            '************************************************************************
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            If BtnFechasClick = "SI" Then
                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            TraducirGridApuntesBD(Me.DgvApuntes)
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
        End If
    End Sub

    Private Sub SeguirF3()
        vCantidadFilas = DgvApuntes.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox(rmse.GetString("MsgDatos2"))
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
                MsgBox(rmse.GetString("MsgDatos2"))
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
End Class