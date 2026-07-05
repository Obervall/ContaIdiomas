Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class TipoCuentaBancaria

    Public vtipoSql, vtipoGrid, vTxtNombre, filaActual As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, PrintLine, Contador As Integer
    Public TL(10) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub TipoCuentasBancarias_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAñadirRegistro, resManager.GetString("ToolTipAñadir"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEditarRegistro, resManager.GetString("ToolTipEditar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnEliminarRegistro, resManager.GetString("ToolTipEliminar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))

        ' Añade una línea por cada GroupBox donde tengas estos botones:
        AddHandler Me.GroupBox2.MouseMove, AddressOf VerificarFiltrosDesactivados
        AddHandler Me.GroupBox3.MouseMove, AddressOf VerificarFiltrosDesactivados

        ' Llenar Grid de TIPO CUENTAS BANCARIAS
        '**************************************
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP, tipocuentas.IdTipoCUE FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirContenidoGridTiposCuenta(DgvTipoCuentasBancarias)

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))
        ' Cambiamos a bucle For desde la columna 0 hasta la penúltima (Count - 2)
        For i As Integer = 0 To DgvTipoCuentasBancarias.Columns.Count - 2
            frmBuscar.CmbCampos.Items.Add(DgvTipoCuentasBancarias.Columns(i).HeaderText)
        Next
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    ''' <summary>
    ''' Función auxiliar que centraliza las reglas lógicas de comparación de texto.
    ''' </summary>
    Private Function EvaluarCoincidencia(celda0 As String, celda1 As String, buscar As String, campo As Integer, exacta As Boolean) As Boolean
        Select Case campo
            Case 0 ' Todos los campos (Celda 0 o Celda 1)
                If exacta Then
                    Return celda0 = buscar OrElse celda1 = buscar
                Else
                    Return celda0.Contains(buscar) OrElse celda1.Contains(buscar)
                End If

            Case 1 ' Solo Nombre / Código (Celda 0)
                If exacta Then Return celda0 = buscar Else Return celda0.Contains(buscar)

            Case 2 ' Solo Descripción (Celda 1)
                If exacta Then Return celda1 = buscar Else Return celda1.Contains(buscar)

            Case Else
                Return False
        End Select
    End Function

    ' BOTÓN BUSCAR: Abre la ventana de parámetros y busca desde el principio si el Check lo pide
    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        ' Llamamos al motor pasándole True para que respete el estado inicial del formulario de búsqueda
        EjecutarBusquedaTipos(forzarDesdeInicio:=True)
    End Sub

    ' BOTÓN SEGUIR BUSCANDO: No abre ventana, busca directamente la siguiente coincidencia
    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        ' Llamamos al motor pasándole False para obligarle a saltar a la siguiente fila
        EjecutarBusquedaTipos(forzarDesdeInicio:=False)
    End Sub

    ' DETECTAR TECLA F3 EN EL FORMULARIO
    Private Sub frmTipoCuentaBancaria_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                EjecutarBusquedaTipos(forzarDesdeInicio:=False)
            End If
        End If
    End Sub

    Private Sub EjecutarBusquedaTipos(ByVal forzarDesdeInicio As Boolean)
        ' Protegemos el código capturando el texto de búsqueda de forma segura
        vBuscar = If(frmBuscar.CmbTextoBuscar.Text, "").ToLower().Trim()
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        Dim buscarExacto As Boolean = frmBuscar.ChkExacta.Checked
        Dim desdePrimerRegistro As Boolean = frmBuscar.ChkPrimerRegistro.Checked

        ' Si no hay nada que buscar, cancelamos
        If String.IsNullOrEmpty(vBuscar) Then Exit Sub

        ' Determinamos el punto de inicio real en la tabla
        Dim filaInicio As Integer = 0

        ' Si se pulsa "Seguir Buscando" O si el usuario NO marcó empezar desde el primer registro
        If (Not forzarDesdeInicio OrElse Not desdePrimerRegistro) AndAlso DgvTipoCuentasBancarias.CurrentRow IsNot Nothing Then
            filaInicio = DgvTipoCuentasBancarias.CurrentRow.Index + 1
        End If

        ' Si al intentar avanzar nos salimos del límite del Grid, avisamos y salimos
        If filaInicio >= DgvTipoCuentasBancarias.Rows.Count Then
            MsgBox(resManager.GetString("MsgDatos2"), MsgBoxStyle.Information, Me.Text)
            BtnSeguirBuscando.Enabled = False
            Exit Sub
        End If

        ' Mapeamos los índices de las columnas para Tipo de Cuentas (0 = Nombre/Código, 1 = Descripción)
        Dim columnasEvaluar As Integer()
        Select Case vCampo
            Case 0 : columnasEvaluar = {0, 1} ' Ambos campos
            Case 1 : columnasEvaluar = {0}    ' Solo el Nombre/Código del Tipo
            Case 2 : columnasEvaluar = {1}    ' Solo la Descripción del Tipo
            Case Else : columnasEvaluar = {}
        End Select

        vRow = -1

        ' Bucle de búsqueda en el DataGridView de tipos de cuentas
        For i As Integer = filaInicio To DgvTipoCuentasBancarias.Rows.Count - 1
            Dim row As DataGridViewRow = DgvTipoCuentasBancarias.Rows(i)

            ' Evitamos evaluar la fila nueva vacía automática del final
            If row.IsNewRow Then Continue For

            Dim coincide As Boolean = False

            For Each colIdx As Integer In columnasEvaluar
                ' Protección contra celdas vacías (Nothing)
                Dim valorCelda As String = ""
                If row.Cells(colIdx).Value IsNot Nothing Then
                    valorCelda = row.Cells(colIdx).Value.ToString().ToLower().Trim()
                End If

                ' Evaluamos según el checkbox de coincidencia exacta
                If buscarExacto Then
                    coincide = (valorCelda = vBuscar)
                Else
                    coincide = valorCelda.Contains(vBuscar)
                End If

                If coincide Then Exit For
            Next

            ' Si encontramos la coincidencia, movemos el foco visual de la pantalla
            If coincide Then
                DgvTipoCuentasBancarias.ClearSelection()
                row.Selected = True

                Try
                    ' Selecciona la celda del campo por el que buscó para forzar el scroll visual
                    DgvTipoCuentasBancarias.CurrentCell = row.Cells(columnasEvaluar(0))
                Catch
                    DgvTipoCuentasBancarias.CurrentCell = row.Cells(0)
                End Try

                ' Sincronizamos el scroll automático
                DgvTipoCuentasBancarias.FirstDisplayedScrollingRowIndex = row.Index

                vRow = row.Index
                Exit For
            End If
        Next

        ' Avisar al usuario si no encontró absolutamente nada en todo el recorrido
        If vRow = -1 Then
            MsgBox(resManager.GetString("MsgDatos1"), MsgBoxStyle.Information, Me.Text)
            BtnSeguirBuscando.Enabled = False
        End If
    End Sub

    Private Function TraducirDinamico(textoOriginal As String, esDescripcion As Boolean) As String
        ' 1. Validamos que no venga vacío
        If String.IsNullOrEmpty(textoOriginal) Then Return ""

        ' 2. Formateamos el texto para que coincida con la Key de ResX Manager (ej: "Cuenta Corriente" -> "Cuenta_Corriente")
        Dim llave As String = textoOriginal.Trim().Replace(" ", "_")
        If esDescripcion Then llave = "Desc_" & llave

        Try
            ' 3. Buscamos en el gestor de recursos (reemplaza 'resManager' por tu objeto ResourceManager activo)
            Dim textoTraducido As String = rmse.GetString(llave)

            ' 4. Si existe traducción en el .resx la devolvemos; si no, devolvemos el texto original de la BD
            If Not String.IsNullOrEmpty(textoTraducido) Then
                Return textoTraducido
            Else
                Return textoOriginal
            End If
        Catch ex As Exception
            ' En caso de cualquier error imprevisto de lectura, no rompemos la app, devolvemos el dato original
            Return textoOriginal
        End Try
    End Function

    Private Sub DgvTipoCuentasBancarias_DoubleClick(sender As Object, e As EventArgs) Handles DgvTipoCuentasBancarias.DoubleClick
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        ' 1. Obtener la fila seleccionada de la cuadrícula de tipos de cuenta
        filaActual = DgvTipoCuentasBancarias.CurrentRow.Index

        ' 2. LEER EL CÓDIGO REAL: La columna del código es la Celda 0 de tu Grid (ej: SPARBUCH / SAVINGS)
        Dim codigoCelda As String = DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString().Trim().ToUpper()

        ' =========================================================================
        ' 3. REVERTIR EL IDIOMA PARA COMPARAR CON TU LISTA BASE EN LA RAM
        ' =========================================================================
        Dim codigoEnEspañol As String = codigoCelda ' Por defecto asumimos que es ese

        Dim resSet As System.Resources.ResourceSet = resManager.GetResourceSet(System.Globalization.CultureInfo.CurrentUICulture, True, True)
        If resSet IsNot Nothing Then
            For Each dict As System.Collections.DictionaryEntry In resSet
                Dim llaveKey As String = dict.Key.ToString()

                ' ¡EL FILTRO FILTRADO MAESTRO!: Si la llave empieza por "Desc_", la ignoramos
                ' por completo para no confundir el tipo con su descripción larga
                If llaveKey.StartsWith("Desc_", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim valorTraducido As String = dict.Value?.ToString().Trim().ToUpper()

                ' Si el texto de la celda coincide con la traducción activa del .resx
                If valorTraducido = codigoCelda Then
                    ' Recuperamos la llave original limpia de la base de datos (Ej: "AHORRO" o "CORRIENTE")
                    codigoEnEspañol = llaveKey.Replace("_", " ").ToUpper()
                    Exit For
                End If
            Next
        End If

        ' =========================================================================
        ' 4. VALIDACIÓN DE BLOQUEO DE EDICIÓN USANDO LA LISTA GLOBAL DE TIPOS
        ' =========================================================================
        ' Apuntamos directamente a tu lista protectora del módulo (TiposCuentaMuestraSistema)
        If TiposCuentaMuestraSistema.Contains(codigoEnEspañol) Then
            Dim msgAviso As String = resManager.GetString("AvisoTipoCuentaProtegido")
            If String.IsNullOrEmpty(msgAviso) Then msgAviso = "Los tipos de cuentas predeterminados del sistema están protegidos contra modificaciones, si no se va a usar se puede Eliminar."
            MessageBox.Show(msgAviso, resManager.GetString("Aviso"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub ' Se frena en seco: bloquea por completo la edición
        End If

        ' =========================================================================
        ' 5. VALIDACIÓN DE BLOQUEO DE EDICIÓN (BLINDADA CONTRA ESPACIOS)
        ' =========================================================================
        ' Eliminamos cualquier guion bajo y espacio para comparar cadenas limpias (Ej: "CUENTACORRIENTE")
        Dim textoValidarLimpio As String = codigoEnEspañol.Replace("_", "").Replace(" ", "").Trim().ToUpper()
        If TiposCuentaMuestraSistema.Contains(textoValidarLimpio) Then
            Dim msgAviso As String = resManager.GetString("AvisoTipoCuentaProtegido")
            If String.IsNullOrEmpty(msgAviso) Then msgAviso = "Los tipos de cuentas predeterminados del sistema están protegidos contra modificaciones, si no se va a usar se puede Eliminar."
            MessageBox.Show(msgAviso, resManager.GetString("Aviso"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' =========================================================================
        ' 6. ABRIR FORMULARIO DE EDICIÓN MODAL (Si es libre, le deja pasar)
        ' =========================================================================
        vTxtNombre = codigoCelda

        If ((frmEditarTipoCuentaBancaria Is Nothing) OrElse (Not frmEditarTipoCuentaBancaria.IsHandleCreated)) Then
            frmEditarTipoCuentaBancaria = New EditarTipoCuentaBancaria
        End If

        ' Forzamos la traducción y la interfaz antes de medir la ventana modal
        ActualizarTextosFormulario(frmEditarTipoCuentaBancaria)
        vEditar = "SI"
        frmEditarTipoCuentaBancaria.ShowDialog()
        frmEditarTipoCuentaBancaria.Dispose()

        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP, tipocuentas.IdTipoCUE FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirContenidoGridTiposCuenta(DgvTipoCuentasBancarias)
        DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(filaActual).Cells(0)
        DgvTipoCuentasBancarias.Rows(filaActual).Selected = True
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click

        ' 1. Verificar si hay alguna fila seleccionada en el Grid
        If frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow Is Nothing Then
            MessageBox.Show(rmse.GetString("SeleccionarTipo"), rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exit Sub
        End If

        ' 2. Obtener la fila actual y capturar el ID y el texto visual
        Dim filaActual As Integer = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.CurrentRow.Index
        Dim textoTraducido As String = frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(0).Value.ToString().Trim()

        ' =========================================================================
        ' ✨ ADIÓS TRADUCCIÓN INVERSA: Leemos el ID numérico directo de la Celda 2
        ' =========================================================================
        Dim idTipoCUE As Integer = 0
        Try
            idTipoCUE = Convert.ToInt32(frmTipoCuentaBancaria.DgvTipoCuentasBancarias.Rows(filaActual).Cells(2).Value)
        Catch ex As Exception
            MessageBox.Show(resManager.GetString("ErrorRecuperarID"), resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try

        ' =========================================================================
        ' 3. VALIDACIÓN DE INTEGRIDAD REAL USANDO EL ID NUMÉRICO
        ' =========================================================================
        ' Filtramos la tabla cuentas buscando el ID numérico en TipoCUE
        Dim vSqlVerificar As String = "SELECT COUNT(*) FROM cuentas WHERE TipoCUE = ?"
        cmdMdb1cr.CommandText = vSqlVerificar
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("?", idTipoCUE)

        Dim cuentasAsociadas As Integer
        Try
            cuentasAsociadas = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())
        Catch ex As Exception
            MessageBox.Show(rmse.GetString("ErrorVerificarIntegridad") & ": " & ex.Message, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
        ' 4. Bloquear el borrado si está asignado a alguna cuenta activa
        If cuentasAsociadas > 0 Then
            Dim msgBloqueo As String = rmse.GetString("NoSePuedeEliminar") & " [" & textoTraducido & "] " &
                                  rmse.GetString("PorqueAsignado") & " " & cuentasAsociadas & " " &
                                  rmse.GetString("CuentaBancaria")
            MessageBox.Show(msgBloqueo, rmse.GetString("AccionCancelada"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' =========================================================================
        ' 5. CONFIRMACIÓN Y EJECUCIÓN DEL BORRADO POR ID
        ' =========================================================================
        Dim msgConfirmar As String = rmse.GetString("SeguroEliminar") & ": " & textoTraducido & "?"
        Dim respuesta As DialogResult = MessageBox.Show(msgConfirmar, rmse.GetString("ConfirmarBorrado"), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2)

        If respuesta = DialogResult.Yes Then
            ' Diseñamos la consulta limpia apuntando a la clave inalterable IdTipoCUE
            Dim vtipoSql As String = "DELETE FROM tipocuentas WHERE IdTipoCUE = ?"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("?", idTipoCUE)

            Try
                Dim filasAfectadas As Integer = cmdMdb1cr.ExecuteNonQuery()

                If filasAfectadas > 0 Then
                    MessageBox.Show(resManager.GetString("RegistroBorrado"), resManager.GetString("ToolTipEliminar"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show(resManager.GetString("RegistroNoBorrado") & " " & textoTraducido, resManager.GetString("ToolTipEliminar"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If

                ' 6. RECARGA AUTOMÁTICA DEL GRID CON FILTROS INTEGRADOS
                vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP, tipocuentas.IdTipoCUE FROM tipocuentas"
                vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
                vtipoGrid = "TIPO_CUENTAS_BANCARIAS"

                LlenarGrid(vtipoSql, vtipoGrid, "1")
                TraducirContenidoGridTiposCuenta(frmTipoCuentaBancaria.DgvTipoCuentasBancarias)

            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorEliminarRegistro") & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            MessageBox.Show(resManager.GetString("AccionCancelada"), rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmNuevoTipoCuentaBancaria Is Nothing) OrElse (Not frmNuevoTipoCuentaBancaria.IsHandleCreated)) Then
            frmNuevoTipoCuentaBancaria = New NuevoTipoCuentaBancaria
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmNuevoTipoCuentaBancaria)
        ' Llamamos al formulario de manera modal.
        frmNuevoTipoCuentaBancaria.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmNuevoTipoCuentaBancaria.Dispose()
        vtipoSql = "SELECT tipocuentas.CodigoTIP, tipocuentas.DescripcionTIP, tipocuentas.IdTipoCUE FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"
        vtipoGrid = "TIPO_CUENTAS_BANCARIAS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirContenidoGridTiposCuenta(DgvTipoCuentasBancarias)
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = 0
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"), vbInformation)
        Else
            vFila = vFilaActual - 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = DgvTipoCuentasBancarias.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = vFilaActual + 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvTipoCuentasBancarias.CurrentRow.Index
        If vFilaActual = DgvTipoCuentasBancarias.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"), vbInformation)
        Else
            vFila = DgvTipoCuentasBancarias.RowCount - 1
            DgvTipoCuentasBancarias.Rows(vFila).Selected = True
            DgvTipoCuentasBancarias.CurrentCell = DgvTipoCuentasBancarias.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub VerificarFiltrosDesactivados(sender As Object, e As MouseEventArgs)
        ' Diccionario con tus botones deshabilitados y sus ToolTips correspondientes
        Dim botonesBloqueados As New Dictionary(Of Button, ToolTip) From {
            {Me.BtnEliminarRegistro, TL(2)},
            {Me.BtnSeguirBuscando, TL(4)}
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
                    'tool.Show(resManager.GetString("ToolTipEliminar"), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    ' Cargamos dinámicamente su texto correspondiente desde tu recurso
                    Dim textoKey As String = If(boton Is Me.BtnSeguirBuscando, "ToolTipSeguirBuscando", "ToolTipEliminar")
                    tool.Show(resManager.GetString(textoKey), Me, posRatonRelativaAlForm.X + 15, posRatonRelativaAlForm.Y + 15)
                    Exit Sub
                End If
            End If
        Next

        ' Si el ratón no está sobre ningún botón bloqueado, ocultamos los tres
        TL(2).Hide(Me)
        TL(4).Hide(Me)
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' 1. Consulta limpia a la base de datos de Access
        vtipoSql = "SELECT * FROM tipocuentas"
        vtipoSql += " ORDER BY tipocuentas.CodigoTIP ASC"

        ' 2. Llenamos el Grid oculto de la trastienda con los datos crudos
        LlenarGrid(vtipoSql, "PRINT_TIPO_CUENTAS", 1)
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString()

        ' 3. Reiniciamos los punteros de página y renglones de la RAM
        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        'Para ver la plantilla de impresión
        'frmImprimirForm.Show()

        ' 2. Lanza el proceso de impresión (esto activa automáticamente el evento PrintPage)
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

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Cualquier variable que desees que conserve su valor debes declararla fuera del Printdocument
        'Todas las variable declaradas dentro de printdocument pierden su valor al cambiar de pagina
        'Definimos los tipos de letras a utilizar en el reporte
        '******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)
        frmImprimirForm.LblTitulo.Text = rmse.GetString("TituloReporte")

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        ' Encabezado Columna 0: Tomamos el texto y lo recortamos si supera los 30 caracteres
        Dim textoEncabezado0 As String = resManager.GetString("Tipo") & ":"
        If textoEncabezado0.Length > 30 Then textoEncabezado0 = textoEncabezado0.Substring(0, 30)
        e.Graphics.DrawString(textoEncabezado0, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)

        ' Encabezado Columna 1: Se queda igual en su posición fija
        e.Graphics.DrawString(resManager.GetString("Descripcion") & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left + 50, frmImprimirForm.Punto2.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        ' Imprimimos los detalles del reporte, es decir el listado de Apuntes
        ' *******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left ' Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top  ' Tomamos la posicion vertical de la letra 'Punto1'

        Dim alturaFila As Integer = CInt(FuenteDetalles.Height * 1.5)

        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If frmImprimirForm.DgvApuntes.Rows(PrintLine).IsNewRow Then
                PrintLine += 1
                Continue Do
            End If

            If startY + alturaFila > e.MarginBounds.Bottom Then
                e.HasMorePages = True
                Exit Do
            End If

            ' =========================================================================
            ' 🌟 COLUMNA 0: TIPO DE CUENTA (Traducción Relacional Directa)
            ' =========================================================================
            Dim valorBD0 As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value?.ToString(), "").Trim()
            Dim textoCelda0 As String = valorBD0

            ' 🚀 TRUCO MAESTRO: Buscamos la traducción en el .resx limpiando espacios con guiones
            If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(valorBD0) Then
                Dim claveRecurso As String = valorBD0.Replace(" ", "_")
                Dim traduccion As String = resManager.GetString(claveRecurso)
                If Not String.IsNullOrEmpty(traduccion) Then textoCelda0 = traduccion.Trim()
            End If
            textoCelda0 = textoCelda0.Replace("_", " ").ToUpper()

            If textoCelda0.Length > 30 Then
                textoCelda0 = textoCelda0.Substring(0, 30)
            End If
            e.Graphics.DrawString(textoCelda0, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)

            ' =========================================================================
            ' 🌟 COLUMNA 1: DESCRIPCIÓN DEL TIPO
            ' =========================================================================
            Dim valorBD1 As String = If(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value?.ToString(), "").Trim()
            Dim textoCelda1 As String = valorBD1

            ' Buscamos en el ResX la combinación "Desc_" & "Nombre_Tipo" (ej: "Desc_AHORRO")
            If resManager IsNot Nothing AndAlso Not String.IsNullOrEmpty(valorBD0) Then
                Dim claveDesc As String = "Desc_" & valorBD0.Replace(" ", "_")
                Dim traduccionDesc As String = resManager.GetString(claveDesc)
                If Not String.IsNullOrEmpty(traduccionDesc) Then textoCelda1 = traduccionDesc.Trim()
            End If

            ' Dibujamos en la hoja controlando el ancho disponible con puntos suspensivos (...)
            Dim anchoDisponibleCol1 As Integer = e.MarginBounds.Right - (frmImprimirForm.Punto2.Left + 50)
            Dim formatoCortado As New StringFormat With {
                .Trimming = StringTrimming.EllipsisCharacter,
                .FormatFlags = StringFormatFlags.NoWrap
            }

            Dim rectanguloCelda1 As New RectangleF(frmImprimirForm.Punto2.Left + 50, startY, anchoDisponibleCol1, FuenteDetalles.Height)
            e.Graphics.DrawString(textoCelda1, FuenteDetalles, Brushes.Black, rectanguloCelda1, formatoCortado)

            ' 🚀 SEPARACIÓN ELEGANTE: Multiplicamos por 2 para que las filas no se aplasten
            startY += (alturaFila * 2)
            PrintLine += 1
            Contador += 1
        Loop

        'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
        'Imprimimos los valores que salen despues del datagridview al final del reporte
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(resManager.GetString("Pagina"), FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

End Class