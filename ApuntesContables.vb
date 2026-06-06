Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class ApuntesContables

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

        vFecha1Enero = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker2.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)
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
        TL(26).SetToolTip(Me.BtnF6, rmse.GetString("ToolTipF6"))
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
        vtipoSql = "Select apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
        vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        End If
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If

        ' Llenar el Combo Concepto y ChekedListBox1
        '******************************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.TipoCON ASC, conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                vTipoConcepto = ""
                vFila = 0
                While drMdb1.Read()
                    If vTipoConcepto <> drMdb1.GetValue(2) Then
                        If vFila = 0 Then
                            If drMdb1.GetValue(0) = "TRASPASO" Then
                                ListBox1.Items.Add(resManager.GetString("TRASPASO"))
                                CmbConcepto.Items.Add(resManager.GetString("TRASPASO"))
                            End If
                            'vTipoConcepto = drMdb1.GetValue(2)
                            vFila += 1
                        Else
                            vTipoConcepto = drMdb1.GetValue(2)
                            If vTipoConcepto = "GASTO" Then
                                ListBox1.Items.Add("** " & rmse.GetString("Label8.Text") & " **")
                            Else
                                ListBox1.Items.Add("** " & rmse.GetString("Label7.Text") & " **")
                            End If
                            CmbConcepto.Items.Add(drMdb1.GetValue(0))
                            ListBox1.Items.Add(drMdb1.GetValue(0))
                        End If
                    Else
                        vFila += 1
                        CmbConcepto.Items.Add(drMdb1.GetValue(0))
                        ListBox1.Items.Add(drMdb1.GetValue(0))
                    End If
                End While
                CmbConcepto.Text = CmbConcepto.Items(1)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMdb1cr.CommandText = "Select * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuenta.Items.Add(drMdb1.GetValue(0))
                End While
                CmbCuenta.Text = CmbCuenta.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
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
    End Sub

    Private Sub BtnFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnFiltroCuenta.Click
        If ListBox1.SelectedItems.Count <> 0 Then
            MessageBox.Show(rmse.GetString("MsgAviso1"), rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            BtnFiltroCuenta.Enabled = False
            BtnSinFiltroCuenta.Enabled = True
            vtipoSql = "Select apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            If BtnFechasClick = "SI" Then
                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If
            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
            End If
            If BtnFiltroFecha.Enabled = False Then
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuntes.FechaAPU >= ?"
                vtipoSql += " And apuntes.FechaAPU <= ?"
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
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

        ' 1. Inicialización de la consulta base
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes WHERE "

        ' 2. Aplicamos la condición base del Ejercicio/Fechas
        If BtnFechasClick = "SI" Then
            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString & " "
        End If

        ' 3. GESTIÓN DE CONCEPTOS (Multiselección o Simple)
        If ListBox1.SelectedItems.Count <> 0 Then ' Listbox con multiselección activo
            TxtConcepto.Text = rmse.GetString("MsgText3")
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.Text = CmbConcepto.Items(0)

            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True

            ' Creamos una lista para almacenar los conceptos procesados
            Dim listaConceptos As New List(Of String)

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim textoOriginal As String = Convert.ToString(ListBox1.SelectedItems(i))
                Dim claveObtenida As String = ObtenerClaveNeutral(textoOriginal, resManager)

                If claveObtenida.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then
                    claveObtenida = ""
                End If

                If Not String.IsNullOrEmpty(claveObtenida) Then
                    vConcepto = claveObtenida
                Else
                    vConcepto = textoOriginal
                End If

                ' Duplicamos las comillas simples por si el concepto del usuario contiene alguna (ej: Foster's)
                listaConceptos.Add("'" & vConcepto.Replace("'", "''") & "'")
            Next

            ' El resultado será: And apuntes.ConceptoAPU IN ('ADESLAS', 'ADESLAS 2', 'LUZ')
            vtipoSql += " And apuntes.ConceptoAPU IN (" & String.Join(",", listaConceptos) & ") "

        Else ' No hay selección múltiple, usamos el ComboBox individual
            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True
            vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text.Replace("'", "''") & "' "
        End If

        ' 4. FILTROS ADICIONALES (Se aplican UNA SOLA VEZ para toda la consulta)
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text.Replace("'", "''") & "' "
        End If

        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"
        End If

        ' 5. Cierre de la consulta y ejecución
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"

        BtnFiltroChekedList.Enabled = False
        ListBox1.Visible = False
        vtipoGrid = "APUNTES_CONTABLES"

        ' Ejecución del llenado del Grid
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Selección automática de la última fila
        If DgvApuntes.RowCount > 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    'Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
    '    '' Ejemplo de prueba para revertir o inversa :
    '    'MsgBox(CmbConcepto.Items(0))
    '    'Dim claveOriginal As String = ObtenerClaveNeutral(CmbConcepto.Items(0), resManager)

    '    '' Muestra en pantalla el nombre de la clave interna (ej: "BtnEliminar.Text" o "Eliminar")
    '    'MsgBox("La clave neutra en la base de recursos es: " & claveOriginal)

    '    If ListBox1.SelectedItems.Count <> 0 Then 'listbox con multiselección
    '        TxtConcepto.Text = rmse.GetString("MsgText3")
    '        CmbConcepto.Items.Clear()
    '        CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
    '        CmbConcepto.Text = CmbConcepto.Items(0)
    '        Dim i As Integer
    '        BtnFiltroConcepto.Enabled = False
    '        BtnSinFiltroConcepto.Enabled = True
    '        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
    '        If BtnFechasClick = "SI" Then
    '            vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
    '        Else
    '            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '        End If
    '        For i = 0 To ListBox1.SelectedItems.Count - 1
    '            'Dim textoConcepto As String = Convert.ToString(ListBox1.SelectedItems(i))
    '            '' 1. Buscamos la clave neutral en el resManager
    '            'Dim claveObtenida As String = ObtenerClaveNeutral(textoConcepto, resManager)
    '            '' 2. ¡EL FILTRO INTELIGENTE!: Si por error la función nos devuelve la clave de la descripción 
    '            '' (porque empieza por "Desc_"), limpiamos la variable para quedarnos solo con el concepto puro.
    '            'If claveObtenida.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then
    '            '    vConcepto = ""
    '            'Else
    '            '    vConcepto = claveObtenida
    '            'End If
    '            '' 3. Si hemos obtenido un concepto válido en mayúsculas, mostramos el aviso
    '            'If Not String.IsNullOrEmpty(vConcepto) Then
    '            '    'MsgBox("Concepto neutro de origen detectado: " & vConcepto)

    '            '    ' 💡 Aquí puedes meter tu lógica para insertar el registro traducido en la base de datos local del usuario...
    '            '    ' Ejemplo: InsertarConceptoEnDB(vConcepto, textoConcepto)


    '            ' 1. Capturamos el texto tal cual sale del ListBox de la Base de Datos
    '            Dim textoOriginal As String = Convert.ToString(ListBox1.SelectedItems(i))

    '            ' 2. Intentamos buscar si es uno de los 34 conceptos de fábrica
    '            Dim claveObtenida As String = ObtenerClaveNeutral(textoOriginal, resManager)

    '            ' 3. Filtramos para asegurarnos de que no se confunda con una descripción ("Desc_")
    '            If claveObtenida.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then
    '                claveObtenida = ""
    '            End If

    '            ' 4. ¡EL CAMBIO CLAVE (LOGICA INVERSA)!: 
    '            ' Si la clave obtenida NO está vacía, es de fábrica -> Usamos la clave neutra.
    '            ' Si está vacía, es un concepto nuevo del usuario -> Usamos su texto original en su idioma.
    '            If Not String.IsNullOrEmpty(claveObtenida) Then
    '                vConcepto = claveObtenida
    '            Else
    '                vConcepto = textoOriginal
    '            End If

    '            If i = 0 Then
    '                vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
    '                If BtnFiltroCuenta.Enabled = False Then
    '                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
    '                End If
    '                If BtnFiltroFecha.Enabled = False Then
    '                    vDate1 = DateTimePicker1.Value.Date
    '                    vDate2 = DateTimePicker2.Value.Date
    '                    vtipoSql += " And apuntes.FechaAPU >= ?"
    '                    vtipoSql += " And apuntes.FechaAPU <= ?"
    '                End If
    '            Else
    '                vtipoSql += " Or "
    '                If BtnFechasClick = "SI" Then
    '                    vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
    '                Else
    '                    vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '                End If
    '                vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
    '                If BtnFiltroCuenta.Enabled = False Then
    '                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
    '                End If
    '                If BtnFiltroFecha.Enabled = False Then
    '                    vDate1 = DateTimePicker1.Value.Date
    '                    vDate2 = DateTimePicker2.Value.Date
    '                    vtipoSql += " And apuntes.FechaAPU >= ?"
    '                    vtipoSql += " And apuntes.FechaAPU <= ?"
    '                End If
    '            End If
    '        Next
    '    Else
    '        BtnFiltroConcepto.Enabled = False
    '        BtnSinFiltroConcepto.Enabled = True
    '        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
    '        If BtnFechasClick = "SI" Then
    '            vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' Andapuntes.EjercicioAPU <> 0 "
    '        Else
    '            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
    '        End If
    '        vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
    '        If BtnFiltroCuenta.Enabled = False Then
    '            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
    '        End If
    '        If BtnFiltroFecha.Enabled = False Then
    '            vDate1 = DateTimePicker1.Value.Date
    '            vDate2 = DateTimePicker2.Value.Date
    '            vtipoSql += " And apuntes.FechaAPU >= ?"
    '            vtipoSql += " And apuntes.FechaAPU <= ?"
    '        End If
    '    End If
    '    vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
    '    BtnFiltroChekedList.Enabled = False
    '    ListBox1.Visible = False
    '    vtipoGrid = "APUNTES_CONTABLES"
    '    LlenarGrid(vtipoSql, vtipoGrid, "1")
    '    If DgvApuntes.RowCount - 1 >= 0 Then
    '        vFila = DgvApuntes.RowCount - 1
    '        DgvApuntes.Rows(vFila).Selected = True
    '        DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
    '    End If
    'End Sub

    Private Sub BtnFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnFiltroFecha.Click
        If ListBox1.SelectedItems.Count <> 0 Then
            'MsgBox("Desactive primero el filtro de Concepto para poder aplicar el filtro de Fecha")
            MessageBox.Show(rmse.GetString("MsgAviso1"), rmse.GetString("MsgText1"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            BtnFiltroFecha.Enabled = False
            BtnSinFiltroFecha.Enabled = True
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
                vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
            End If
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"

            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnSinFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroCuenta.Click
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False
        If ListBox1.SelectedItems.Count <> 0 Then
            TxtConcepto.Text = rmse.GetString("MsgText3")
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.Text = CmbConcepto.Items(0)
            Dim i As Integer
            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            If BtnFechasClick = "SI" Then
                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' Andapuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If
            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString
                If i = 0 Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
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
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                    If BtnFiltroFecha.Enabled = False Then
                        vDate1 = DateTimePicker1.Value.Date
                        vDate2 = DateTimePicker2.Value.Date
                        vtipoSql += " And apuntes.FechaAPU >= ?"
                        vtipoSql += " And apuntes.FechaAPU <= ?"
                    End If
                End If
            Next
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            BtnFiltroChekedList.Enabled = False
            ListBox1.Visible = False
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        Else
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            If BtnFechasClick = "SI" Then
                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If
            vtipoSql += " And apuntes.CuentaAPU <> '' "
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
            End If
            If BtnFiltroFecha.Enabled = False Then
                vDate1 = DateTimePicker1.Value.Date
                vDate2 = DateTimePicker2.Value.Date
                vtipoSql += " And apuntes.FechaAPU >= ?"
                vtipoSql += " And apuntes.FechaAPU <= ?"
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' Llenar el Combo Concepto y ChekedListBox1
        '******************************************
        CmbConcepto.Items.Clear()
        ListBox1.Items.Clear()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.TipoCON ASC, conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                vTipoConcepto = ""
                vFila = 0
                While drMdb1.Read()
                    If vTipoConcepto <> drMdb1.GetValue(2) Then
                        If vFila = 0 Then
                            If drMdb1.GetValue(0) = "TRASPASO" Then
                                ListBox1.Items.Add(resManager.GetString("TRASPASO"))
                                CmbConcepto.Items.Add(resManager.GetString("TRASPASO"))
                            End If
                            'vTipoConcepto = drMdb1.GetValue(2)
                            vFila += 1
                        Else
                            vTipoConcepto = drMdb1.GetValue(2)
                            If vTipoConcepto = "GASTO" Then
                                ListBox1.Items.Add("** " & rmse.GetString("Label8.Text") & " **")
                            Else
                                ListBox1.Items.Add("** " & rmse.GetString("Label7.Text") & " **")
                            End If

                            CmbConcepto.Items.Add(drMdb1.GetValue(0))
                            ListBox1.Items.Add(drMdb1.GetValue(0))
                        End If
                    Else
                        vFila += 1
                        CmbConcepto.Items.Add(drMdb1.GetValue(0))
                        ListBox1.Items.Add(drMdb1.GetValue(0))
                    End If
                End While
                CmbConcepto.Text = CmbConcepto.Items(1)
                vConcepto = CmbConcepto.Text.ToString
                drMdb1.Close()
                cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
                drMdb1 = cmdMdb1cr.ExecuteReader()
                drMdb1.Read()
                TxtConcepto.Text = drMdb1.GetValue(1)
                drMdb1.Close()
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        frmTipoInformeApuntes.RadioButton1.Enabled = True
        frmTipoInformeApuntes.RadioButton2.Enabled = True
        frmTipoInformeApuntes.RadioButton5.Enabled = False
        BtnFiltroChekedList.Enabled = True
        BtnFiltroConcepto.Enabled = True
        CmbConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        'Quitar todos los checked
        For i = 0 To ListBox1.Items.Count - 1
            ListBox1.SetSelected(i, False)
        Next
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
        If BtnFechasClick = "SI" Then
            vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
        Else
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
        End If
        vtipoSql += " And apuntes.ConceptoAPU <> '' "
        If BtnFiltroCuenta.Enabled = False Then
            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
        End If
        If BtnFiltroFecha.Enabled = False Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"
        End If
        vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
        vtipoGrid = "APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        If DgvApuntes.RowCount - 1 >= 0 Then
            vFila = DgvApuntes.RowCount - 1
            DgvApuntes.Rows(vFila).Selected = True
            DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSinFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroFecha.Click
        If BtnFechasClick = "SI" Then
            BtnFechasClick = "NO"
            BtnFechasFondo.Visible = False
        End If
        vFecha1Enero = Val(vAñoEjercicio)
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False
        If ListBox1.SelectedItems.Count <> 0 Then
            TxtConcepto.Text = rmse.GetString("MsgText3")
            CmbConcepto.Items.Clear()
            CmbConcepto.Items.Add(rmse.GetString("MsgText4"))
            CmbConcepto.Text = CmbConcepto.Items(0)
            Dim i As Integer
            BtnFiltroConcepto.Enabled = False
            BtnSinFiltroConcepto.Enabled = True
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            If BtnFechasClick = "SI" Then
                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
            Else
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            End If
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString
                If i = 0 Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                    If BtnFiltroFecha.Enabled = False Then
                        vDate1 = DateTimePicker1.Value.Date
                        vDate2 = DateTimePicker2.Value.Date
                        vtipoSql += " And apuntes.FechaAPU >= ?"
                        vtipoSql += " And apuntes.FechaAPU <= ?"
                    End If
                    If BtnFiltroCuenta.Enabled = False Then
                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                    End If
                Else
                    vtipoSql += " Or "
                    If BtnFechasClick = "SI" Then
                        vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                    Else
                        vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    End If
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                    If BtnFiltroFecha.Enabled = False Then
                        vDate1 = DateTimePicker1.Value.Date
                        vDate2 = DateTimePicker2.Value.Date
                        vtipoSql += " And apuntes.FechaAPU >= ?"
                        vtipoSql += " And apuntes.FechaAPU <= ?"
                    End If
                    If BtnFiltroCuenta.Enabled = False Then
                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                    End If
                End If
            Next
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            BtnFiltroChekedList.Enabled = False
            ListBox1.Visible = False
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        Else
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
                vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
            End If
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
            End If
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            If DgvApuntes.RowCount - 1 >= 0 Then
                vFila = DgvApuntes.RowCount - 1
                DgvApuntes.Rows(vFila).Selected = True
                DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
            End If
        End If
    End Sub

    Private Sub CmbCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCuenta.SelectedIndexChanged
        If ListBox1.SelectedItems.Count = 0 Then
            If BtnFiltroCuenta.Enabled = False Then
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                End If
                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        Else
            If BtnFiltroCuenta.Enabled = False Then
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM   apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To ListBox1.SelectedItems.Count - 1
                    vConcepto = ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                'MsgBox(vtipoSql)
                LlenarGrid(vtipoSql, vtipoGrid, "1")
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                'MsgBox(vtipoSql)
                LlenarGrid(vtipoSql, vtipoGrid, "1")
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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

        If vTxtNombre = "SALDO" Then
            MsgBox(rmse.GetString("MsgSaldos1"))
        Else
            ' Comprobamos si existe un identificador asociado.
            If ((frmEditarApuntes Is Nothing) OrElse (Not frmEditarApuntes.IsHandleCreated)) Then
                frmEditarApuntes = New EditarApuntes
            End If
            ' Llamamos al formulario de manera modal.
            vEditar = "NO"  ' Eliminar
            frmEditarApuntes.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmEditarApuntes.Dispose()
            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
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
                If DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = DgvApuntes.RowCount - 1
                    DgvApuntes.Rows(vFila).Selected = True
                    DgvApuntes.CurrentCell = DgvApuntes.Rows(vFila).Cells(0)
                End If
            Else
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                    vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If
        End If
    End Sub

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
                        vFecha1Enero = Val(drMdb1.GetValue(0))
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
            vFecha1Enero = Val(vAñoEjercicio)
            vFecha31Diciembre = Val(vAñoEjercicio)
            DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
            DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
            DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
            DateTimePicker2.MinDate = New Date(vFecha1Enero, 1, 1)
            DateTimePicker2.MaxDate = New Date(vFecha31Diciembre, 12, 31)
            DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
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
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                End If
                If BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
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
                                dr1(j) = Format(DgvApuntes.Rows(i).Cells(j).Value, "yyyy/MM/dd")
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
                            dr1(j) = Format(DgvApuntes.Rows(i).Cells(j).Value, "yyyy/MM/dd")
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
            Dim formatoMonedaRojo As String = "#,##0.00 €;[Red]-#,##0.00 €"
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
                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                    End If
                    If BtnFiltroConcepto.Enabled = False Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                    End If
                    If BtnFiltroFecha.Enabled = False Then
                        vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "# "
                        vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "# "
                    End If
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
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
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If BtnFiltroConcepto.Enabled = False Then
                            vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                        End If
                        If BtnFiltroFecha.Enabled = False Then
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "# "
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "# "
                        End If
                    End If
                    vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                    vtipoGrid = "APUNTES_CONTABLES"
                    LlenarGrid(vtipoSql, vtipoGrid, "1")
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
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        If ListBox1.SelectedItems.Count = 0 Then
            vConcepto = CmbConcepto.Text.ToString
            drMdb1.Close()
            cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
            drMdb1 = cmdMdb1cr.ExecuteReader()
            drMdb1.Read()
            TxtConcepto.Text = drMdb1.GetValue(1)
            drMdb1.Close()
            If BtnFiltroConcepto.Enabled = False Then
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If
                vtipoSql += " And apuntes.ConceptoAPU = '" & CmbConcepto.Text & "' "
                If BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                End If
                If BtnFiltroFecha.Enabled = False Then
                    vDate1 = DateTimePicker1.Value.Date
                    vDate2 = DateTimePicker2.Value.Date
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
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