Imports System.Collections.Generic
Imports System.Drawing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class ConceptosContables

    Public vtipoSql, vtipoGrid, vTxtNombre, filaActual As String
    Public vRow, vCampo As Integer
    Public PrintLine, Contador As Integer
    Public TL(12) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub ConceptosContables_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Me.KeyPreview = True
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnFiltroTipoConcepto, resManager.GetString("ToolTipAplicarFiltro"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSinFiltroTipoConcepto, resManager.GetString("ToolTipQuitarFiltro"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAñadirRegistro, resManager.GetString("ToolTipAñadir"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnEditarRegistro, resManager.GetString("ToolTipEditar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnEliminarRegistro, resManager.GetString("ToolTipEliminar"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))

        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox4.MouseMove, AddressOf VerificarFiltrosDesactivados

        ' 1. Configuramos primero el estilo del ComboBox
        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList

        ' 2. LLENAMOS EL COMBO PRIMERO (Así tendrá elementos antes de seleccionar el índice 0)
        ActualizarIdiomaComboConcepto(Me.CmbTipoConcepto, True)

        ' 3. Ahora que ya tiene filas, seleccionamos de forma segura la primera
        CmbTipoConcepto.SelectedIndex = 0

        ' 4. Cargar los datos puros de la BD y traducirlos inmediatamente
        CargarYTraducirGridCompleto()

        ' 5. Llenar el Combo Campos de búsqueda usando los títulos traducidos
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        For Each columna As DataGridViewColumn In DgvConceptos.Columns
            frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
        Next
    End Sub

    ' --- MÉTODOS DE CONSULTA A BASE DE DATOS Y TRADUCCIÓN ---

    Private Sub CargarYTraducirGridCompleto()
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"

        ' Carga los datos de la BD en el Grid
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Traduce los textos de las celdas
        TraducirCeldasDelGrid()
    End Sub

    Private Sub FiltrarYTraducirGrid()
        Dim tipoParaDB As String = ""
        Select Case CmbTipoConcepto.SelectedIndex
            Case 0 : tipoParaDB = "GASTO"
            Case 1 : tipoParaDB = "INGRESO"
            Case 2 : tipoParaDB = "ESPECIAL"
        End Select

        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        vtipoSql += " WHERE conceptos.TipoCON = '" & tipoParaDB & "' "
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"

        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirCeldasDelGrid()
    End Sub

    ''' <summary>
    ''' RECORRE Y TRADUCE LAS CELDAS: Lee los valores y aplica las traducciones de ResX Manager
    ''' </summary>
    Private Sub TraducirCeldasDelGrid()
        Try
            If DgvConceptos IsNot Nothing AndAlso DgvConceptos.Rows.Count > 0 Then
                For Each fila As DataGridViewRow In DgvConceptos.Rows
                    If Not fila.IsNewRow Then

                        ' Verificamos que las celdas críticas tengan valor
                        If fila.Cells(0).Value IsNot Nothing AndAlso fila.Cells(1).Value IsNot Nothing Then

                            Dim tipoOriginal As String = fila.Cells(0).Value.ToString().Trim().ToUpper()
                            Dim codigoOriginal As String = fila.Cells(1).Value.ToString().Trim()
                            Dim llaveBase As String = codigoOriginal.Replace(" ", "_")

                            ' --- TRADUCIR COLUMNA (0): TipoCON ---
                            Dim tradTipo As String = ""
                            Select Case tipoOriginal
                                Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                                Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                                Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                            End Select
                            If Not String.IsNullOrEmpty(tradTipo) Then fila.Cells(0).Value = tradTipo

                            ' --- TRADUCIR COLUMNA (1): CodigoCON ---
                            Dim tradCodigo As String = resManager.GetString(llaveBase)
                            If Not String.IsNullOrEmpty(tradCodigo) Then fila.Cells(1).Value = tradCodigo

                            ' --- TRADUCIR COLUMNA (2): DescripcionCON ---
                            Dim llaveDesc As String = "Desc_" & llaveBase
                            Dim tradDesc As String = resManager.GetString(llaveDesc)
                            If Not String.IsNullOrEmpty(tradDesc) Then fila.Cells(2).Value = tradDesc


                            ' --- TRADUCIR COLUMNA (3): NotasCON (Solo si el origen es ESPECIAL) ---
                            If tipoOriginal = "ESPECIAL" AndAlso fila.Cells(3).Value IsNot Nothing Then
                                Dim llaveNota As String = "Nota_" & llaveBase

                                ' Buscamos primero en el global, si no, en el local
                                Dim tradNota As String = rmse.GetString(llaveNota)
                                If String.IsNullOrEmpty(tradNota) Then tradNota = rmse.GetString(llaveNota)

                                If Not String.IsNullOrEmpty(tradNota) Then fila.Cells(3).Value = tradNota
                            End If
                        End If
                    End If
                Next
                ' Ordena la columna (1): CodigoCON de forma Ascendente utilizando la ordenación automática del Grid
                DgvConceptos.Sort(DgvConceptos.Columns(1), System.ComponentModel.ListSortDirection.Ascending)

                ' Aplica el formato de color a la columna (0) según el tipo, comparando tanto con el texto original como con el traducido
                For Each fila As DataGridViewRow In frmConceptosContables.DgvConceptos.Rows
                    If fila.Cells(0).Value IsNot Nothing Then
                        Dim valorCelda As String = fila.Cells(0).Value.ToString().Trim()
                        ' Comparamos con el texto en español OR con el texto traducido actual
                        If valorCelda = "GASTO" OrElse valorCelda = resManager.GetString("Tipo_Gasto") Then
                            fila.Cells(0).Style.ForeColor = Color.DarkRed

                        ElseIf valorCelda = "INGRESO" OrElse valorCelda = resManager.GetString("Tipo_Ingreso") Then
                            fila.Cells(0).Style.ForeColor = Color.DarkBlue

                        ElseIf valorCelda = "ESPECIAL" OrElse valorCelda = resManager.GetString("Tipo_Especial") Then
                            fila.Cells(0).Style.ForeColor = Color.DarkGreen
                        End If
                    End If
                Next

            End If
        Catch ex As Exception
            ' Evita cuelgues visuales si el volcado está incompleto
        End Try
    End Sub

    ' --- ACCIONES DE FILTROS Y BOTONES ---

    Private Sub BtnSinFiltroTipoConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroTipoConcepto.Click
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
        CargarYTraducirGridCompleto()
    End Sub

    Private Sub BtnFiltroTipoConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroTipoConcepto.Click
        BtnFiltroTipoConcepto.Enabled = False
        BtnSinFiltroTipoConcepto.Enabled = True
        FiltrarYTraducirGrid()
    End Sub

    Private Sub CmbTipoConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbTipoConcepto.SelectedIndexChanged
        If BtnFiltroTipoConcepto.Enabled = False Then
            FiltrarYTraducirGrid()
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If e.CloseReason = 3 Then e.Cancel = False
        BtnFiltroTipoConcepto.Enabled = True
        BtnSinFiltroTipoConcepto.Enabled = False
    End Sub

    ' --- MOTOR DE BÚSQUEDA DEL FORMULARIO ---

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True
        EjecutarBusquedaConceptos(forzarDesdeInicio:=True)
    End Sub

    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        EjecutarBusquedaConceptos(forzarDesdeInicio:=False)
    End Sub

    Private Sub ConceptosContables_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled AndAlso e.KeyCode = Keys.F3 Then
            EjecutarBusquedaConceptos(forzarDesdeInicio:=False)
        End If
    End Sub

    Private Sub EjecutarBusquedaConceptos(ByVal forzarDesdeInicio As Boolean)
        Dim buscarTexto As String = frmBuscar.CmbTextoBuscar.Text.ToLower().Trim()
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        Dim exacta As Boolean = frmBuscar.ChkExacta.Checked
        Dim desdePrimerRegistro As Boolean = frmBuscar.ChkPrimerRegistro.Checked

        Dim filaInicio As Integer = 0
        If Not forzarDesdeInicio OrElse Not desdePrimerRegistro Then
            If vRow >= 0 AndAlso vRow < DgvConceptos.Rows.Count Then
                filaInicio = vRow + 1
            ElseIf DgvConceptos.CurrentRow IsNot Nothing Then
                filaInicio = DgvConceptos.CurrentRow.Index + 1
            End If
        End If

        Dim coincidenciaEncontrada As Boolean = False

        For i As Integer = filaInicio To DgvConceptos.Rows.Count - 1
            Dim fila As DataGridViewRow = DgvConceptos.Rows(i)
            If fila.IsNewRow Then Continue For

            Dim celdasAEvaluar As New List(Of Integer)()
            If vCampo = 0 Then
                For c As Integer = 0 To fila.Cells.Count - 1 : celdasAEvaluar.Add(c) : Next
            Else
                celdasAEvaluar.Add(vCampo - 1)
            End If

            For Each idx As Integer In celdasAEvaluar
                If idx < fila.Cells.Count AndAlso fila.Cells(idx).Value IsNot Nothing Then
                    Dim contenidoCelda As String = fila.Cells(idx).Value.ToString().ToLower().Trim()
                    If (exacta AndAlso contenidoCelda = buscarTexto) OrElse (Not exacta AndAlso contenidoCelda.Contains(buscarTexto)) Then
                        vRow = i
                        DgvConceptos.CurrentCell = DgvConceptos.Rows(i).Cells(0)
                        coincidenciaEncontrada = True
                        Exit For
                    End If
                End If
            Next
            If coincidenciaEncontrada Then Exit For
        Next
        If Not coincidenciaEncontrada Then
            MsgBox(resManager.GetString("MsgDatos1"), vbInformation, resManager.GetString("ToolTipBuscar"))
            vRow = -1
        End If
    End Sub

    ''' <summary>
    ''' Controla de forma visual la activación o desactivación de los botones de filtro
    ''' </summary>
    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnSinFiltroTipoConcepto, TL(1)},
            {Me.BtnSeguirBuscando, TL(6)}
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
        TL(6).Hide(Me)
    End Sub

    ' --- BOTONES DE DESPLAZAMIENTO / NAVEGACIÓN ---

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = 0
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If

        'If DgvConceptos.Rows.Count > 0 Then
        '    DgvConceptos.CurrentCell = DgvConceptos.Rows(0).Cells(0)
        'End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = vFilaActual - 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If

        'If DgvConceptos.CurrentRow IsNot Nothing Then
        '    Dim filaIndex As Integer = DgvConceptos.CurrentRow.Index
        '    If filaIndex > 0 Then
        '        DgvConceptos.CurrentCell = DgvConceptos.Rows(filaIndex - 1).Cells(0)
        '    End If
        'End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = DgvConceptos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = vFilaActual + 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If

        'If DgvConceptos.CurrentRow IsNot Nothing Then
        '    Dim filaIndex As Integer = DgvConceptos.CurrentRow.Index
        '    ' Restamos 1 o 2 dependiendo de si el Grid tiene activa la fila en blanco de inserción final
        '    Dim limite As Integer = If(DgvConceptos.AllowUserToAddRows, DgvConceptos.Rows.Count - 2, DgvConceptos.Rows.Count - 1)

        '    If filaIndex < limite Then
        '        DgvConceptos.CurrentCell = DgvConceptos.Rows(filaIndex + 1).Cells(0)
        '    End If
        'End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvConceptos.CurrentRow.Index
        If vFilaActual = DgvConceptos.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = DgvConceptos.RowCount - 1
            DgvConceptos.Rows(vFila).Selected = True
            DgvConceptos.CurrentCell = DgvConceptos.Rows(vFila).Cells(0)
        End If

        'If DgvConceptos.Rows.Count > 0 Then
        '    Dim limite As Integer = If(DgvConceptos.AllowUserToAddRows, DgvConceptos.Rows.Count - 2, DgvConceptos.Rows.Count - 1)
        '    If limite >= 0 Then
        '        DgvConceptos.CurrentCell = DgvConceptos.Rows(limite).Cells(0)
        '    End If
        'End If
    End Sub

    ' --- ACCIONES PRINCIPALES DEL MANTENIMIENTO ---

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmNuevoConceptoContable Is Nothing) OrElse (Not frmNuevoConceptoContable.IsHandleCreated)) Then
            frmNuevoConceptoContable = New NuevoConceptoContable
        End If
        ' Llamamos al formulario de manera modal.
        frmNuevoConceptoContable.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmNuevoConceptoContable.Dispose()
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        If BtnFiltroTipoConcepto.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & CmbTipoConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' Traduce los textos de las celdas
        TraducirCeldasDelGrid()

        '' Configuramos la variable para indicar que es un registro NUEVO
        'vEditar = "NO"

        '' Abrimos el formulario de edición en modo modal
        'frmEditarConceptoContable.ShowDialog()

        '' Al regresar, recargamos el Grid completo para ver el nuevo registro traducido
        'CargarYTraducirGridCompleto()
    End Sub

    Private Sub DgvConceptos_DoubleClick(sender As Object, e As EventArgs) Handles DgvConceptos.DoubleClick
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index
        vTxtNombre = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString

        ' --- VALIDACIÓN: PROTEGER CONCEPTOS DE MUESTRA (CORREGIDO) ---
        Dim esConceptoDeMuestra As Boolean = False
        Dim textoCelda As String = vTxtNombre.Trim().ToUpper()

        ' Si el idioma es español, la llave coincide directamente con el texto
        Dim llaveBase As String = textoCelda.Replace(" ", "_")
        If resManager.GetString(llaveBase) IsNot Nothing OrElse resManager.GetString("Desc_" & llaveBase) IsNot Nothing Then
            esConceptoDeMuestra = True
        ElseIf My.Settings.CulturaUsuario <> "es" Then
            ' Si está en otro idioma, buscamos de forma inversa en los valores del ResX
            Dim resSet As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentCulture, True, True)
            If resSet IsNot Nothing Then
                For Each dict As System.Collections.DictionaryEntry In resSet
                    ' Comparamos el valor traducido en el ResX con el texto visible de la celda
                    If dict.Value.ToString().Trim().ToUpper() = textoCelda Then
                        esConceptoDeMuestra = True
                        Exit For
                    End If
                Next
            End If
        End If

        If esConceptoDeMuestra Then
            Dim msgAviso As String = resManager.GetString("AvisoConceptoProtegido")
            If String.IsNullOrEmpty(msgAviso) Then msgAviso = "Los conceptos predeterminados del sistema están protegidos contra modificaciones."

            MessageBox.Show(msgAviso, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        ' -------------------------------------------------------------
        ' -------------------------------------------------

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarConceptoContable Is Nothing) OrElse (Not frmEditarConceptoContable.IsHandleCreated)) Then
            frmEditarConceptoContable = New EditarConceptoContable
        End If
        ' Llamamos al formulario de manera modal.
        If vEditar = "NO" Then
            vEditar = "SI"  ' Editar
        Else
            vEditar = "SI"
        End If
        frmEditarConceptoContable.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarConceptoContable.Dispose()
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        If BtnFiltroTipoConcepto.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & CmbTipoConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirCeldasDelGrid()
        DgvConceptos.CurrentCell = DgvConceptos.Rows(filaActual).Cells(0)
        DgvConceptos.Rows(filaActual).Selected = True
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        filaActual = frmConceptosContables.DgvConceptos.CurrentRow.Index
        vTxtNombre = frmConceptosContables.DgvConceptos.Rows(filaActual).Cells(1).Value.ToString

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarConceptoContable Is Nothing) OrElse (Not frmEditarConceptoContable.IsHandleCreated)) Then
            frmEditarConceptoContable = New EditarConceptoContable
        End If
        ' Llamamos al formulario de manera modal.
        If vEditar = "NO" Then
            vEditar = "NO"  ' Eliminar
        Else
            vEditar = "NO"
        End If
        frmEditarConceptoContable.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmEditarConceptoContable.Dispose()
        vtipoSql = "SELECT conceptos.TipoCON, conceptos.CodigoCON, conceptos.DescripcionCON, conceptos.NotasCON FROM conceptos"
        If BtnFiltroTipoConcepto.Enabled = False Then
            vtipoSql += " WHERE "
            vtipoSql += "conceptos.TipoCON = '" & CmbTipoConcepto.Text & "' "
        End If
        vtipoSql += " ORDER BY conceptos.CodigoCON ASC"
        vtipoGrid = "CONCEPTOS_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirCeldasDelGrid()
        DgvConceptos.CurrentCell = DgvConceptos.Rows(filaActual).Cells(0)
        DgvConceptos.Rows(filaActual).Selected = True

    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' 1. Reiniciamos las variables de control de páginas antes de lanzar la impresión
        PrintLine = 0
        Contador = 0

        ' ✨ CORRECCIÓN CRÍTICA: Recuperamos la carga de datos en la plantilla 'frmImprimirForm'
        ' Generamos la consulta SQL idéntica a la que usas en tu pantalla principal para ver los conceptos
        Dim sqlConceptos As String = "SELECT TipoCON, CodigoCON, DescripcionCON, NotasCON FROM conceptos ORDER BY CodigoCON ASC"

        Try
            ' Aquí llamas a tu método del módulo (pon el nombre real de tu Sub, por ejemplo: LlenarGrid)
            ' Le pasas la consulta SQL, el identificador del Grid "PRINT_CONCEPTOS" y tus parámetros habituales
            LlenarGrid(sqlConceptos, "PRINT_CONCEPTOS", "1")

            ' Traducimos las celdas del Grid de la plantilla "al vuelo" antes de imprimir
            ' para que salgan en el papel en el idioma que el usuario tiene activo AHORA MISMO
            If My.Settings.CulturaUsuario IsNot Nothing AndAlso Not My.Settings.CulturaUsuario.StartsWith("es", StringComparison.OrdinalIgnoreCase) Then
                For Each fila As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                    If Not fila.IsNewRow AndAlso fila.Cells(1).Value IsNot Nothing Then
                        Dim codigoOriginal As String = fila.Cells(1).Value.ToString().Trim()
                        Dim llaveBase As String = codigoOriginal.Replace(" ", "_")

                        ' Traducir Código (Celda 1)
                        Dim tradCodigo As String = resManager.GetString(llaveBase)
                        If Not String.IsNullOrEmpty(tradCodigo) Then fila.Cells(1).Value = tradCodigo

                        ' Traducir Descripción (Celda 2)
                        Dim tradDesc As String = resManager.GetString("Desc_" & llaveBase)
                        If Not String.IsNullOrEmpty(tradDesc) Then fila.Cells(2).Value = tradDesc

                        ' Traducir Tipo (Celda 0)
                        Dim tipoOriginal As String = fila.Cells(0).Value.ToString().Trim().ToUpper()
                        Dim tradTipo As String = ""
                        Select Case tipoOriginal
                            Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                            Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                            Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                        End Select
                        If Not String.IsNullOrEmpty(tradTipo) Then fila.Cells(0).Value = tradTipo
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox("Error al preparar los datos para la impresión: " & ex.Message)
            Exit Sub
        End Try


        ' 2. Lógica de lanzamiento de la impresión (Se mantiene igual a tu código original)
        Dim seHaLanzado As Boolean = False

        ' En tu BtnImprimir_Click, antes de "If My.Settings.Previsualizar = True Then"
        PrintDocument1.DefaultPageSettings = New System.Drawing.Printing.PageSettings(PrintDocument1.PrinterSettings)
        Application.DoEvents() ' Fuerza a Windows a vaciar la caché visual y aplicar el idioma actual


        If My.Settings.Previsualizar = True Then
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
            seHaLanzado = True
        End If

        If My.Settings.ElegirImpresora = True AndAlso Not seHaLanzado Then
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
                seHaLanzado = True
            End If
        End If

        If My.Settings.DirectoImpresora = True AndAlso Not seHaLanzado Then
            PrintDocument1.Print()
            seHaLanzado = True
        End If

        If Not seHaLanzado Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' 1. CONFIGURACIÓN DE FUENTES, FORMATOS Y CULTURA EN CALIENTE
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)

        Dim sfDerecha As New StringFormat With {.Alignment = StringAlignment.Far}

        ' Capturamos de forma estricta la cultura guardada por el usuario para el hilo actual
        Dim cultura As New System.Globalization.CultureInfo(My.Settings.CulturaUsuario)
        System.Globalization.CultureInfo.CurrentCulture = cultura
        System.Globalization.CultureInfo.CurrentUICulture = cultura

        ' 2. DETERMINAR TÍTULO Y FECHA DESDE EL RESX
        Dim textoTituloFinal As String = resManager.GetString("TituloReporteConceptos", cultura)
        If String.IsNullOrEmpty(textoTituloFinal) Then textoTituloFinal = "Listado de Conceptos Contables"
        frmImprimirForm.LblTitulo.Text = textoTituloFinal

        ' Generamos la fecha larga con el formato regional del idioma activo
        Dim textoFecha As String = DateTime.Now.ToString("D", cultura)

        ' 3. DIBUJAR ENCABEZADO ESTRUCTURAL DE LA PLANTILLA
        ' Imprimimos la fecha larga perfectamente pegada al margen derecho de la hoja
        e.Graphics.DrawString(textoFecha, FuenteNegrita, Brushes.Black, e.MarginBounds.Right, frmImprimirForm.LblFecha.Top, sfDerecha)

        ' Imprimimos el título y la imagen usando los Left de tu plantilla preferida
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)

        If frmImprimirForm.PictureBox1.Image IsNot Nothing Then
            Dim newImage As Image = frmImprimirForm.PictureBox1.Image
            e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        End If

        ' Imprimimos los títulos de columnas dinámicos usando las posiciones de la plantilla (Punto1, Punto2, Punto3)
        e.Graphics.DrawString(resManager.GetString("Tipo", cultura) & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Codigo", cultura) & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Descripcion", cultura) & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)

        ' Línea divisoria superior
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        Dim startX As Integer = frmImprimirForm.Punto1.Left
        Dim startY As Integer = frmImprimirForm.Punto1.Top

        ' Formato para que la descripción se corte limpiamente si excede los márgenes
        Dim formatoCortado As New StringFormat()
        formatoCortado.Trimming = StringTrimming.EllipsisCharacter
        formatoCortado.FormatFlags = StringFormatFlags.NoWrap

        ' 4. BUCLE DE IMPRESIÓN DE FILAS (RECORRIDO DEL GRID DE LA PLANTILLA CONCEPTOS)
        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            ' Evitamos procesar la fila vacía automática si existe al final
            If frmImprimirForm.DgvApuntes.Rows(PrintLine).IsNewRow Then
                PrintLine += 1
                Contador += 1
                Continue Do
            End If

            ' Control de salto de página automático
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' Extraemos los valores ya traducidos en caliente del grid de la plantilla
            Dim tipoActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString().Trim(), "")
            Dim codigoActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString().Trim(), "")
            Dim descActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value?.ToString().Trim(), "")
            Dim notaActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value?.ToString().Trim(), "")

            ' Manejo especial para notas de sistema en conceptos ESPECIALES
            Dim textoCelda2 As String = descActual
            Dim tipoUpper As String = tipoActual.ToUpper()
            If tipoUpper = "ESPECIAL" OrElse (resManager.GetString("Tipo_Especial", cultura) IsNot Nothing AndAlso tipoUpper = resManager.GetString("Tipo_Especial", cultura).ToUpper()) Then
                Dim llaveNota As String = "Nota_" & codigoActual.Replace(" ", "_")
                Dim tradNota As String = resManager.GetString(llaveNota, cultura)
                textoCelda2 = If(Not String.IsNullOrEmpty(tradNota), tradNota, notaActual)
            End If

            ' DIBUJAR LOS DATOS USANDO LOS PUNTOS DE TU PLANTILLA PREFERIDA
            e.Graphics.DrawString(tipoActual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(codigoActual, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)

            ' Calculamos el ancho de la columna descripción basándonos en el margen derecho de la hoja
            Dim anchoDisponibleCol2 As Integer = e.MarginBounds.Right - frmImprimirForm.Punto3.Left
            Dim rectanguloCelda2 As New RectangleF(frmImprimirForm.Punto3.Left, startY, anchoDisponibleCol2, frmImprimirForm.Punto1.Height)
            e.Graphics.DrawString(textoCelda2, FuenteDetalles, Brushes.Black, rectanguloCelda2, formatoCortado)

            ' Avanzamos espacio vertical para la siguiente fila
            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop

        ' 5. LÍNEA DE PIE DE PÁGINA (Idéntica a la del encabezado)
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
        End If

        ' 6. CONTADOR DE PÁGINAS DINÁMICO
        frmImprimirForm.LblNumeroPagina.Text = (CInt(frmImprimirForm.LblNumeroPagina.Text) + 1).ToString()
        e.Graphics.DrawString(resManager.GetString("Pagina", cultura), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

    ' --- MÓDULO DE IMPRESIÓN DEL REPORTE (PRINT DOCUMENT) ---
    'Private Sub PrintDocument1_PrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
    '    ' 1. CONFIGURACIÓN DE FUENTES Y CULTURA EN CALIENTE
    '    Dim FuenteTitulo As New Font("Microsoft Sans Serif", 14, FontStyle.Bold)
    '    Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
    '    Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
    '    Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Or FontStyle.Bold)

    '    ' Capturamos de forma estricta la cultura guardada por el usuario para el hilo actual
    '    Dim cultura As New System.Globalization.CultureInfo(My.Settings.CulturaUsuario)
    '    System.Globalization.CultureInfo.CurrentCulture = cultura
    '    System.Globalization.CultureInfo.CurrentUICulture = cultura

    '    ' 2. DETERMINAR IDIOMA PARA ENCABEZADOS Y TÍTULOS DESDE EL RESX
    '    ' Leemos el título limpio directamente desde tu nueva llave del archivo de recursos
    '    Dim textoTituloFinal As String = resManager.GetString("TituloReporteConceptos", cultura)
    '    If String.IsNullOrEmpty(textoTituloFinal) Then textoTituloFinal = "Listado de Conceptos Contables"

    '    ' --- SOLUCIÓN A LA FECHA LARGA (ALINEADA A LA DERECHA) ---
    '    ' 1. Creamos un formato para obligar a que el texto se alinee a la derecha (Far)
    '    Dim sfFecha As New StringFormat With {.Alignment = StringAlignment.Far}

    '    ' 2. Generamos la fecha larga con la "D" mayúscula
    '    Dim textoFecha As String = DateTime.Now.ToString("D", cultura)

    '    ' 3. Dibujamos la fecha pegada al margen derecho exacto de la hoja (e.MarginBounds.Right)
    '    e.Graphics.DrawString(textoFecha, FuenteNegrita, Brushes.Black, e.MarginBounds.Right, e.MarginBounds.Top, sfFecha)

    '    ' B. El Logo en su margen superior seguro
    '    If frmImprimirForm.PictureBox1.Image IsNot Nothing Then
    '        e.Graphics.DrawImage(frmImprimirForm.PictureBox1.Image, e.MarginBounds.Left, e.MarginBounds.Top + 10, 40, 40)
    '    End If

    '    ' C. El Título Grande bajado a Top + 35 y SIN añadidos después del guión
    '    e.Graphics.DrawString(textoTituloFinal, FuenteTitulo, Brushes.Black, e.MarginBounds.Left + 55, e.MarginBounds.Top + 35)


    '    ' 4. REDISEÑO DE COLUMNAS FIJAS Y ALINEACIÓN DE ENCABEZADOS
    '    Dim encTipo As String = resManager.GetString("Tipo", cultura)
    '    If String.IsNullOrEmpty(encTipo) Then encTipo = "Tipo"

    '    Dim encCodigo As String = resManager.GetString("Codigo", cultura)
    '    If String.IsNullOrEmpty(encCodigo) Then encCodigo = "Código"

    '    Dim encDesc As String = resManager.GetString("Descripcion", cultura)
    '    If String.IsNullOrEmpty(encDesc) Then encDesc = "Descripción"

    '    Dim colTipoLeft As Integer = e.MarginBounds.Left
    '    Dim colCodigoLeft As Integer = e.MarginBounds.Left + 110
    '    Dim colDescripcionLeft As Integer = e.MarginBounds.Left + 320

    '    ' Dibujamos las etiquetas de cabecera de las 3 columnas
    '    e.Graphics.DrawString(encTipo & ":", FuenteSubrayada, Brushes.Black, colTipoLeft, e.MarginBounds.Top + 95)
    '    e.Graphics.DrawString(encCodigo & ":", FuenteSubrayada, Brushes.Black, colCodigoLeft, e.MarginBounds.Top + 95)
    '    e.Graphics.DrawString(encDesc & ":", FuenteSubrayada, Brushes.Black, colDescripcionLeft, e.MarginBounds.Top + 95)

    '    ' Línea divisoria debajo de los encabezados
    '    e.Graphics.DrawString(New String("-"c, 110), FuenteDetalles, Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top + 115)

    '    ' Posición vertical donde empiezan las filas de los apuntes
    '    Dim startY As Integer = e.MarginBounds.Top + 135
    '    Dim altoFila As Integer = 20

    '    Dim formatoCortado As New StringFormat()
    '    formatoCortado.Trimming = StringTrimming.EllipsisCharacter
    '    formatoCortado.FormatFlags = StringFormatFlags.NoWrap

    '    ' 5. BUCLE DE IMPRESIÓN DE FILAS (RECORRIDO DEL GRID DE LA PLANTILLA)
    '    If frmImprimirForm.DgvApuntes IsNot Nothing AndAlso frmImprimirForm.DgvApuntes.Rows.Count > 0 Then
    '        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
    '            ' Control de salto de página automático
    '            If startY + altoFila > e.MarginBounds.Bottom Then
    '                e.HasMorePages = True
    '                Exit Do
    '            End If

    '            ' Extraemos los valores que el módulo y el botón ya tradujeron en el DataGrid de la plantilla
    '            Dim tipoActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString().Trim(), "")
    '            Dim codigoActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString().Trim(), "")
    '            Dim descActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value?.ToString().Trim(), "")
    '            Dim notaActual As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value?.ToString().Trim(), "")

    '            ' Manejo especial para notas de sistema si corresponde
    '            Dim textoCelda2 As String = descActual
    '            Dim tipoUpper As String = tipoActual.ToUpper()

    '            ' Si el tipo es Especial (o su equivalente traducido), buscamos su nota larga en el recurso
    '            If tipoUpper = "ESPECIAL" OrElse tipoUpper = resManager.GetString("Tipo_Especial", cultura).ToUpper() Then
    '                ' Generamos la llave directa usando el código limpio (ej: "Nota_ALQUILER")
    '                Dim llaveNota As String = "Nota_" & codigoActual.Replace(" ", "_")
    '                Dim tradNota As String = resManager.GetString(llaveNota, cultura)
    '                textoCelda2 = If(Not String.IsNullOrEmpty(tradNota), tradNota, notaActual)
    '            End If

    '            ' Acortamos el texto del código si excede de 30 caracteres para que no pise la descripción
    '            If codigoActual.Length > 30 Then codigoActual = codigoActual.Substring(0, 30)

    '            ' DIBUJAR LOS DATOS ALINEADOS POR COLUMNAS
    '            e.Graphics.DrawString(tipoActual, FuenteDetalles, Brushes.Black, colTipoLeft, startY)
    '            e.Graphics.DrawString(codigoActual, FuenteDetalles, Brushes.Black, colCodigoLeft, startY)

    '            ' Rectángulo de la descripción para evitar que sobresalga de la página
    '            Dim anchoDisponibleDesc As Integer = e.MarginBounds.Right - colDescripcionLeft
    '            Dim rectanguloDesc As New RectangleF(colDescripcionLeft, startY, anchoDisponibleDesc, altoFila)
    '            e.Graphics.DrawString(textoCelda2, FuenteDetalles, Brushes.Black, rectanguloDesc, formatoCortado)

    '            startY += altoFila
    '            PrintLine += 1
    '            Contador += 1
    '        Loop
    '    End If

    '    ' Línea de fondo al terminar el reporte (Corregida con la misma lógica que el encabezado)
    '    If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
    '        e.Graphics.DrawString(New String("-"c, 110), FuenteDetalles, Brushes.Black, e.MarginBounds.Left, startY + 5)
    '    End If

    '    ' 6. CONTADOR DE PÁGINAS DINÁMICO (CORREGIDO)
    '    Static NumeroPaginaImpresion As Integer = 0
    '    Dim textoPaginaVisible As String = resManager.GetString("Pagina", cultura)
    '    If String.IsNullOrEmpty(textoPaginaVisible) Then textoPaginaVisible = "Página"

    '    ' Incrementamos el número de página
    '    NumeroPaginaImpresion += 1

    '    ' Imprimimos el número de página abajo al centro de la hoja usando la variable correcta
    '    e.Graphics.DrawString(textoPaginaVisible & " " & NumeroPaginaImpresion, FuenteDetalles, Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width / 2) - 30, e.MarginBounds.Bottom + 20)

    '    ' Si terminó el reporte completo, reiniciamos el contador para la próxima vez
    '    If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
    '        NumeroPaginaImpresion = 0
    '    End If
    'End Sub
End Class
