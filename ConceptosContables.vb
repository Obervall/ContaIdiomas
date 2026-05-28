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
            MsgBox(resManager.GetString("NoSeEncontraronCoincidencias"), vbInformation, resManager.GetString("ToolTipBuscar"))
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

        '' Validamos si hay una fila seleccionada
        'If DgvConceptos.CurrentRow Is Nothing Then
        '    MsgBox(resManager.GetString("SeleccionarTipo"), vbExclamation, resManager.GetString("AccionCancelada"))
        '    Exit Sub
        'End If

        '' Indicamos que vamos a EDITAR un registro existente
        'vEditar = "SI"
        'frmEditarConceptoContable.ShowDialog()

        '' Refrescamos la pantalla al cerrar el diálogo
        'CargarYTraducirGridCompleto()
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

        '' 1. Validar fila seleccionada
        'If DgvConceptos.CurrentRow Is Nothing Then
        '    MsgBox(resManager.GetString("SeleccionarTipo"), vbExclamation, resManager.GetString("AccionCancelada"))
        '    Exit Sub
        'End If

        '' 2. Capturamos los datos de la fila actual del Grid
        'filaActual = DgvConceptos.CurrentRow.Index
        'vTxtNombre = DgvConceptos.Rows(filaActual).Cells(1).Value.ToString().Trim() ' Código del concepto (ej: TRASPASO)

        '' 3. EXCEPCIÓN CRÍTICA: Bloquear si intentan borrar el concepto de sistema "ESPECIAL"
        '' Realizamos una búsqueda inversa rápida por si el texto de la celda 0 viene traducido
        'Dim tipoCelda As String = DgvConceptos.Rows(filaActual).Cells(0).Value.ToString().Trim().ToUpper()
        'Dim esEspecial As Boolean = (tipoCelda = "ESPECIAL" OrElse tipoCelda = resManager.GetString("Tipo_Especial").ToUpper())

        'If esEspecial Then
        '    MsgBox(resManager.GetString("ConceptoSistemaNoBorrar"), vbExclamation, resManager.GetString("AccionCancelada"))
        '    Exit Sub
        'End If

        '' 4. Mensaje de confirmación en el idioma correcto
        'Dim mensaje As String = resManager.GetString("EliminarConcepto") & " " & vTxtNombre & " " & resManager.GetString("EliminarConcepto2")
        'If MsgBox(mensaje, vbQuestion + vbYesNo + vbDefaultButton2, resManager.GetString("LblEliminando")) = vbYes Then

        '    ' 5. Lanzamos el borrado físico usando el código crudo de la BD
        '    cmdMdb1cr.Parameters.Clear()
        '    vtipoSql = "DELETE FROM conceptos WHERE conceptos.CodigoCON = '" & vTxtNombre & "'"
        '    cmdMdb1cr.CommandText = vtipoSql

        '    Try
        '        cmdMdb1cr.ExecuteNonQuery()
        '        MsgBox(resManager.GetString("EliminarConcepto3"), vbInformation, resManager.GetString("RegistroBorrado"))

        '        ' Recarga completa tras borrar
        '        CargarYTraducirGridCompleto()
        '    Catch ex As Exception
        '        MsgBox(resManager.GetString("EliminarConcepto4") & vbNewLine & ex.Message, vbCritical, resManager.GetString("Error"))
        '    End Try
        'End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' Reiniciamos las variables de control de páginas antes de lanzar la impresión
        PrintLine = 0
        Contador = 0

        'Para ver la plantilla de impresión
        '    'frmImprimirForm.Show()

        ' Suponiendo que tienes un control PrintDocument en el formulario
        If My.Settings.Previsualizar = True Then
            'Te deja ver un preview del reporte antes de imprimir
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        If My.Settings.ElegirImpresora = True Then
            'Te deja elegir la impresora
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        End If

        If My.Settings.DirectoImpresora = True Then
            'Imprime en la impresora por defecto
            PrintDocument1.Print()
        End If
    End Sub
    ' --- MÓDULO DE IMPRESIÓN DEL REPORTE (PRINT DOCUMENT) ---

    Private Sub PrintDocument1_PrintPage(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        ' Definimos los tipos de letras a utilizar en el reporte
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)

        ' Configuramos el título unificando el recurso de texto
        frmImprimirForm.LblTitulo.Text = resManager.GetString("TituloReporte") & " " & frmTipoCuentaBancaria.Text

        ' Imprimimos el encabezado y los datos fijos del formulario
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)

        ' Imagen/Logo del encabezado
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image
        e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        ' Imprimimos las etiquetas de cabecera de las columnas
        Dim textoEncabezado0 As String = resManager.GetString("Label3.Text") & ":"
        If textoEncabezado0.Length > 30 Then textoEncabezado0 = textoEncabezado0.Substring(0, 30)
        e.Graphics.DrawString(textoEncabezado0, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)

        ' Línea divisoria debajo de los encabezados
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        ' Posicionamiento inicial de las filas
        Dim startX As Integer = frmImprimirForm.Punto1.Left
        Dim startY As Integer = frmImprimirForm.Punto1.Top

        ' Formato para que la descripción y notas se corten con "..." antes del margen derecho de la hoja
        Dim formatoCortado As New StringFormat()
        formatoCortado.Trimming = StringTrimming.EllipsisCharacter
        formatoCortado.FormatFlags = StringFormatFlags.NoWrap

        ' --- BUCLE DE IMPRESIÓN DE FILAS (RECORRIDO DEL GRID) ---
        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            ' Control de salto de página automático si se supera el borde inferior
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' 1. Extraemos los valores crudos de la fila del Grid de forma segura
            Dim tipoOriginal As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString().Trim().ToUpper(), "")
            Dim codigoOriginal As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString().Trim(), "")
            Dim descOriginal As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value?.ToString().Trim(), "")
            Dim notaOriginal As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value?.ToString().Trim(), "")

            ' 2. TRADUCCIÓN HÍBRIDA EN CALIENTE PARA EL REPORTE IMREPSO
            ' --- Columna Tipo ---
            Dim textoCelda0 As String = ""
            Select Case tipoOriginal
                Case "GASTO" : textoCelda0 = resManager.GetString("Tipo_Gasto")
                Case "INGRESO" : textoCelda0 = resManager.GetString("Tipo_Ingreso")
                Case "ESPECIAL" : textoCelda0 = resManager.GetString("Tipo_Especial")
                Case Else : textoCelda0 = tipoOriginal
            End Select
            If String.IsNullOrEmpty(textoCelda0) Then textoCelda0 = tipoOriginal

            ' --- Columna Código / Nombre del Concepto ---
            Dim llaveConcepto As String = codigoOriginal.Replace(" ", "_")
            Dim tradConcepto As String = resManager.GetString(llaveConcepto)
            Dim textoCelda1 As String = If(Not String.IsNullOrEmpty(tradConcepto), tradConcepto, codigoOriginal)

            ' Aplicamos el límite estricto de máximo 30 caracteres solicitado para el concepto
            If textoCelda1.Length > 30 Then textoCelda1 = textoCelda1.Substring(0, 30)

            ' --- Columna Descripción / Notas (Si es ESPECIAL traduce la nota de sistema) ---
            Dim textoCelda2 As String = descOriginal
            If tipoOriginal = "ESPECIAL" Then
                Dim llaveNota As String = "Nota_" & llaveConcepto
                Dim tradNota As String = resManager.GetString(llaveNota)
                textoCelda2 = If(Not String.IsNullOrEmpty(tradNota), tradNota, notaOriginal)
            End If

            ' 3. DIBUJAR LOS TEXTOS TRADUCIDOS EN LA HOJA
            ' Imprimimos el Tipo en el Punto1
            e.Graphics.DrawString(textoCelda0, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)

            ' Imprimimos el Nombre del Concepto (Celda 1) desplazado levemente
            e.Graphics.DrawString(textoCelda1, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left + 80, startY)

            ' Imprimimos la Descripción o Nota en el Punto2 limitando el ancho dinámicamente hasta el borde derecho de la hoja
            Dim anchoDisponibleCol1 As Integer = e.MarginBounds.Right - frmImprimirForm.Punto2.Left
            Dim rectanguloCelda2 As New RectangleF(frmImprimirForm.Punto2.Left, startY, anchoDisponibleCol1, frmImprimirForm.Punto1.Height)
            e.Graphics.DrawString(textoCelda2, FuenteDetalles, Brushes.Black, rectanguloCelda2, formatoCortado)

            ' Desplazamiento vertical y contadores de control
            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop

        ' Imprimimos el pie de página o cierres al terminar todos los registros
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
        End If

        ' Contador dinámico de número de páginas del reporte impreso
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

End Class
