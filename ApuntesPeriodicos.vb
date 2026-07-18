Imports System.Data
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class ApuntesPeriodicos

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid, vTxtNombre As String
    Public vRow, vRowSeguir, vCampo, vContador, vCantidadFilas, filaSelec As Integer
    Public TL(20) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub ApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargandoFormulario = True

        Me.KeyPreview = True
        ' 1. Convertimos el año base de forma segura
        Dim anioBase As Integer
        If Not Integer.TryParse(vAñoEjercicio, anioBase) Then
            ' Si falla, usamos el año actual como salvavidas
            anioBase = Date.Today.Year
        End If

        ' 2. Calculamos los dos años que necesitas
        Dim anioInicio As Integer = anioBase
        Dim anioFin As Integer = anioBase + 20 ' Sumamos los 20 años de margen para los periódicos

        ' 3. Guardamos los valores en tus variables globales por si las usas luego
        vFecha1Enero = anioInicio
        vFecha31Diciembre = anioFin

        ' 4. Creamos las fechas exactas de inicio y fin
        Dim fechaInicio As New Date(anioInicio, 1, 1)
        Dim fechaFin As New Date(anioFin, 12, 31)

        ' 5. Configuramos los DateTimePicker con los rangos correctos
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker2.MinDate = fechaInicio

        DateTimePicker1.MaxDate = fechaFin
        DateTimePicker2.MaxDate = fechaFin

        ' 6. Asignamos los valores iniciales por defecto
        DateTimePicker1.Value = fechaInicio
        DateTimePicker2.Value = fechaFin

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
        TL(9).SetToolTip(Me.BtnBuscarRegistro, resManager.GetString("ToolTipBuscar"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnSeguirBuscando, resManager.GetString("ToolTipSeguirBuscando"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnImprimir, resManager.GetString("ToolTipImprimir"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.BtnGraficos, resManager.GetString("ToolTipGraficos"))
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(14) = New ToolTip
        TL(14).SetToolTip(Me.BtnSalir, resManager.GetString("ToolTipSalir"))
        TL(15) = New ToolTip
        TL(15).SetToolTip(Me.BtnPrimero, resManager.GetString("ToolTipPrimero"))
        TL(16) = New ToolTip
        TL(16).SetToolTip(Me.BtnAnterior, resManager.GetString("ToolTipAnterior"))
        TL(17) = New ToolTip
        TL(17).SetToolTip(Me.BtnSiguiente, resManager.GetString("ToolTipSiguiente"))
        TL(18) = New ToolTip
        TL(18).SetToolTip(Me.BtnUltimo, resManager.GetString("ToolTipUltimo"))
        TL(19) = New ToolTip
        TL(19).SetToolTip(Me.BtnEliminaSeleccion, resManager.GetString("ToolTipEliminaSeleccion"))
        TL(20) = New ToolTip
        TL(20).SetToolTip(Me.BtnF6, resManager.GetString("ToolTipF6"))


        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 1. CONSULTA SQL MAESTRA RELACIONAL ALINEADA CON CORTAFUEGOS DE AÑO
        ' =========================================================================
        vtipoSql = "SELECT apuper.FechaAPP As [FechaAPP], " &
                   "conceptos.CodigoCON As [ConceptoAPP], " &
                   "apuper.DescripcionAPP As [DescripcionAPP], " &
                   "apuper.ImporteAPP As [ImporteAPP], " &
                   "apuper.ImporteAPP As [SaldoAPP], " &
                   "apuper.NotasAPP As [NotasAPP], " &
                   "cuentas.NombreCUE As [CuentaAPP], " &
                   "apuper.CodigoAPP As [CodigoAPP], " &
                   "conceptos.CodigoCON As [CodigoCON], " &
                   "apuper.ConceptoAPP As [IdConceptoCON], " &
                   "apuper.CuentaAPP As [IdCuentaCUE], " &
                   "conceptos.TipoCON As [TipoCON] " &
                   "FROM (apuper " &
                   "INNER JOIN conceptos ON apuper.ConceptoAPP = conceptos.IdConceptoCON) " &
                   "INNER JOIN cuentas ON apuper.CuentaAPP = cuentas.IdCuentaCUE"

        ' 🎯 LA ESTOCADA: Filtramos estrictamente por el ejercicio de trabajo actual
        vtipoSql += " WHERE apuper.EjercicioAPP = ?"
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"

        vtipoGrid = "APUNTES_PERIODICOS"

        ' Volcamos los datos relacionales traducidos en tu DataGridView
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuper)

        ' =========================================================================
        ' 🌟 RECARGA DE COMBOS DE LA NUEVA ERA (Inmune a NullReference y Ordenado A-Z)
        ' =========================================================================
        Try
            ' 2. LLAMADA SEGURA: Usamos la nueva función exclusiva para combos sin ListBox
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 3. Llamamos a la función genérica de tu módulo para las cuentas
            LlenarComboCuentasGenerico(Me.CmbCuenta)

            ' 4. Apagamos el escudo tras la carga exitosa en memoria RAM
            cargandoFormulario = False

            ' 5. SELECCIÓN INICIAL SEGURA: Forzamos el vaivén para sincronizar descripciones
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 0
            End If
            If CmbCuenta.Items.Count > 0 Then
                CmbCuenta.SelectedIndex = -1
                CmbCuenta.SelectedIndex = 0
            End If

        Catch ex As Exception
            cargandoFormulario = False
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' Llenar el Combo Campos
        '***********************
        frmBuscar.CmbCampos.Items.Clear()
        frmBuscar.CmbCampos.Items.Add(resManager.GetString("Todos_Los_Campos"))

        For Each columna As DataGridViewColumn In DgvApuper.Columns
            If columna.Name <> "ImporteAPP" And columna.Name <> "SaldoAPP" And columna.Name <> "CuentaAPP" And columna.Name <> "CodigoAPP" And columna.Name <> "CodigoCON" And columna.Name <> "IdConceptoCON" And columna.Name <> "IdCuentaCUE" Then
                frmBuscar.CmbCampos.Items.Add(columna.HeaderText)
            End If
        Next

    End Sub

    Private Sub BtnFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnFiltroCuenta.Click
        ' 1. Modificamos el estado estético de los botones de la pantalla
        BtnFiltroCuenta.Enabled = False
        BtnSinFiltroCuenta.Enabled = True

        ' 2. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        ' 1. Modificamos el estado estético de los botones de la pantalla
        BtnFiltroConcepto.Enabled = False
        BtnSinFiltroConcepto.Enabled = True

        ' 2. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub BtnFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnFiltroFecha.Click
        ' 1. Modificamos el estado estético de los botones de la pantalla
        BtnFiltroFecha.Enabled = False
        BtnSinFiltroFecha.Enabled = True

        ' 2. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub BtnSinFiltroCuenta_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroCuenta.Click
        ' 1. Restauramos el estado estético de la botonería
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False

        ' 2. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        ' 1. Restauramos el estado estético de la botonería
        ' 🌟 ¡CORREGIDO!: Eliminamos el bucle While destructivo que borraba los IDs del combo
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False

        ' 2. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub BtnSinFiltroFecha_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroFecha.Click
        ' 1. Devolvemos los calendarios a sus límites anuales por defecto
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)

        ' 2. Restauramos el estado estético de la botonería
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False

        ' 3. 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional a IDs
        RefrescarGridApuntesPeriodicos()
    End Sub

    Private Sub CmbCuenta_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbCuenta.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR: Bloquea ejecuciones prematuras en el Load
        If cargandoFormulario Then Exit Sub
        If CmbCuenta.SelectedIndex < 0 Then Exit Sub

        ' Si el botón de filtro de cuenta está activo, refrescamos la rejilla
        If BtnFiltroCuenta.Enabled = False Then
            RefrescarGridApuntesPeriodicos()
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR: Bloquea ejecuciones prematuras en el Load
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            ' Extraemos los datos de la fila seleccionada directamente desde la RAM
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                ' Rellenamos el cuadro de texto del concepto de forma limpia
                TxtConcepto.Text = filaSeleccionada("CodigoCON").ToString().Trim()
            End If

            ' Si el botón de filtro de concepto está activo, refrescamos la rejilla
            If BtnFiltroConcepto.Enabled = False Then
                RefrescarGridApuntesPeriodicos()
            End If

        Catch ex As Exception
            ' Manejo silencioso de interfaz
        End Try
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        ' 🌟 ESCUDO PROTECTOR: Si el formulario está cargando o reseteando fechas, salimos
        If cargandoFormulario Then Exit Sub

        ' Si el botón de filtro de fecha está activo, refrescamos la rejilla con IDs
        If BtnFiltroFecha.Enabled = False Then
            ' 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional
            RefrescarGridApuntesPeriodicos()
        End If
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        ' 🌟 ESCUDO PROTECTOR: Si el formulario está cargando o reseteando fechas, salimos
        If cargandoFormulario Then Exit Sub

        ' Si el botón de filtro de fecha está activo, refrescamos la rejilla con IDs
        If BtnFiltroFecha.Enabled = False Then
            ' 🚀 TRUCO MAESTRO: Delegamos todo en la rutina unificada relacional
            RefrescarGridApuntesPeriodicos()
        End If
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        ' 1. Validamos que haya al menos una fila seleccionada de verdad en la rejilla de periódicos
        If DgvApuper.SelectedRows.Count = 0 Then
            MsgBox(resManager.GetString("MsgSeleccionarFilasEliminar"), MsgBoxStyle.Information)
            Exit Sub
        End If

        ' 2. CUADRO DE CONFIRMACIÓN: Preguntamos antes de extirpar de la BD el lote completo
        ' =========================================================================
        Dim mensajeConfirmacion As String = ""
        Dim totalFilas As Integer = DgvApuper.SelectedRows.Count

        If totalFilas = 1 Then
            Dim plantillaSingular As String = resManager.GetString("MsgConfirmarBorradoSingular")
            If String.IsNullOrEmpty(plantillaSingular) Then
                plantillaSingular = "¿Está completamente seguro de que desea eliminar FÍSICAMENTE de la Base de Datos el apunte periódico seleccionado?"
            End If
            mensajeConfirmacion = plantillaSingular
        Else
            Dim plantillaPlural As String = resManager.GetString("MsgConfirmarBorradoPlural")
            If String.IsNullOrEmpty(plantillaPlural) Then
                plantillaPlural = "¿Está completamente seguro de que desea eliminar FÍSICAMENTE de la Base de Datos los {0} apuntes periódicos seleccionados?"
            End If
            mensajeConfirmacion = String.Format(plantillaPlural, totalFilas)
        End If

        Dim tituloVentana As String = If(resManager?.GetString("ConfirmarBorrado"), "Confirmar Borrado Múltiple")
        If ConfirmarAccionTraducida(mensajeConfirmacion, tituloVentana) = MsgBoxResult.No Then
            Exit Sub
        End If

        Dim contadorBorrados As Integer = 0

        ' =========================================================================
        ' 🌟 REPARADO MODO INTEGRAL: BUCLE DE EXTIRPACIÓN EN LA TABLA APUPER
        ' =========================================================================
        For Each fila As DataGridViewRow In DgvApuper.SelectedRows
            ' Saltamos la fila vacía del final del grid por seguridad si existiera
            If fila.IsNewRow Then Continue For

            ' 🎯 LA CLAVE: Rescatamos el ID físico único (CodigoAPP) que viaja en la celda 7 (Oculta)
            ' Asegúrate de que el ID relacional de tu tabla apuper esté en la columna 7 de tu DataTable
            If fila.Cells(7).Value IsNot Nothing AndAlso Not IsDBNull(fila.Cells(7).Value) Then
                Dim idRegistroFisico As Integer = Convert.ToInt32(fila.Cells(7).Value)

                ' Lanzamos la sentencia DELETE individual dirigida al corazón de apuper
                Using cmdDelete As New OleDb.OleDbCommand("DELETE FROM apuper WHERE CodigoAPP = ?", conexion1)
                    cmdDelete.Parameters.Clear()
                    cmdDelete.Parameters.Add("@id", OleDb.OleDbType.Integer).Value = idRegistroFisico

                    Try
                        cmdDelete.ExecuteNonQuery()
                        contadorBorrados += 1
                    Catch ex As Exception
                        Dim plantillaError As String = resManager.GetString("ErrorBorrarFilaID")
                        If String.IsNullOrEmpty(plantillaError) Then
                            plantillaError = "Error al borrar el apunte periódico con ID {0}: "
                        End If
                        Dim mensajeFinal As String = String.Format(plantillaError, idRegistroFisico) & ex.Message
                        MsgBox(mensajeFinal, MsgBoxStyle.Critical, resManager.GetString("Error"))
                    End Try
                End Using
            End If
        Next

        RefrescarGridApuntesPeriodicos()

        ' Volvemos a pasar el rodillo matemático que calcula saldos e ingresos/gastos de forma limpia
        DgvApuntesPeriodicos()

        ' Avisamos del resultado final al usuario
        ' =========================================================================
        Dim plantillaExito As String = resManager.GetString("MsgLoteEliminadoExito")
        If String.IsNullOrEmpty(plantillaExito) Then
            plantillaExito = "Operación completada. Se han eliminado {0} apuntes periódicos de la Base de Datos."
        End If

        Dim tituloExito As String = resManager.GetString("TituloBorradoFinalizado")
        If String.IsNullOrEmpty(tituloExito) Then tituloExito = "Borrado Finalizado"

        Dim mensajeFinalExito As String = String.Format(plantillaExito, contadorBorrados)
        MsgBox(mensajeFinalExito, MsgBoxStyle.Information, tituloExito)
    End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoGraficoPeriodico Is Nothing) OrElse (Not frmTipoGraficoPeriodico.IsHandleCreated)) Then
            frmTipoGraficoPeriodico = New TipoGraficoPeriodico
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmTipoGraficoPeriodico)
        ' Llamamos al formulario de manera modal.
        frmTipoGraficoPeriodico.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoGraficoPeriodico.Dispose()
        ' =========================================================================
        ' 🌟 1. CONSULTA SQL MAESTRA RELACIONAL ALINEADA (¡Corregido!)
        ' =========================================================================
        ' Traemos las 11 celdas biológicas en su orden real simétrico.
        ' 🚀 LA CORRECCIÓN: Cambiamos conceptos.DescripcionCON por conceptos.CodigoCON en la segunda columna.
        vtipoSql = "SELECT apuper.FechaAPP As [FechaAPP], " &
                   "conceptos.CodigoCON As [ConceptoAPP], " &
                   "apuper.DescripcionAPP As [DescripcionAPP], " &
                   "apuper.ImporteAPP As [ImporteAPP], " &
                   "apuper.ImporteAPP As [SaldoAPP], " &
                   "apuper.NotasAPP As [NotasAPP], " &
                   "cuentas.NombreCUE As [CuentaAPP], " &
                   "apuper.CodigoAPP As [CodigoAPP], " &
                   "conceptos.CodigoCON As [CodigoCON], " &
                   "apuper.ConceptoAPP As [IdConceptoCON], " &
                   "apuper.CuentaAPP As [IdCuentaCUE], " &
                   "conceptos.TipoCON As [TipoCON] " & ' 🚀 LA CLAVE: Inyectamos el Tipo real en la posición 11
                   "FROM (apuper " &
                   "INNER JOIN conceptos ON apuper.ConceptoAPP = conceptos.IdConceptoCON) " &
                   "INNER JOIN cuentas ON apuper.CuentaAPP = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"

        vtipoGrid = "APUNTES_PERIODICOS"

        ' Volcamos los datos relacionales traducidos en tu DataGridView
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuper)

        ' =========================================================================
        ' 🌟 RECARGA DE COMBOS DE LA NUEVA ERA (Inmune a NullReference y Ordenado A-Z)
        ' =========================================================================
        Try
            ' 2. LLAMADA SEGURA: Usamos la nueva función exclusiva para combos sin ListBox
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)

            ' 3. Llamamos a la función genérica de tu módulo para las cuentas
            LlenarComboCuentasGenerico(Me.CmbCuenta)

            ' 4. Apagamos el escudo tras la carga exitosa en memoria RAM
            cargandoFormulario = False

            ' 5. SELECCIÓN INICIAL SEGURA: Forzamos el vaivén para sincronizar descripciones
            If CmbConcepto.Items.Count > 0 Then
                CmbConcepto.SelectedIndex = -1
                CmbConcepto.SelectedIndex = 0
            End If
            If CmbCuenta.Items.Count > 0 Then
                CmbCuenta.SelectedIndex = -1
                CmbCuenta.SelectedIndex = 0
            End If

        Catch ex As Exception
            cargandoFormulario = False
            MsgBox(resManager.GetString("Error") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub BtnEliminaSeleccion_Click(sender As Object, e As EventArgs) Handles BtnEliminaSeleccion.Click
        'Elimina las Filas Seleccionadas
        '*******************************
        For Each r As DataGridViewRow In DgvApuper.SelectedRows
            If DgvApuper.Rows.Count > 1 Then
                DgvApuper.Rows.Remove(r)
            End If
        Next
        filaSelec = DgvApuper.CurrentRow.Index
        For i = 0 To DgvApuper.Rows.Count - 1
            DgvApuper.Rows(i).Selected = False
        Next
        'Variable que guardara el valor
        'Dim iTotal As Integer = Me.DgvApuper.Rows.Count 'ITotal toma el valor del numero de registros que tiene la tabla
        'Definimos la variable i para controlar el ciclo for
        'Definimos del ciclo que va desde que i vale cero hasta que i valga itotal menos uno, osea el penultimo regsitro de la tabla
        DgvApuntesContables(3, 4)
        DgvApuper.Select()
        DgvApuper.CurrentRow.Selected = True
        DgvApuper.Refresh()
    End Sub

    Private Sub BtnEditarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEditarRegistro.Click
        ' 1. Validamos de forma preventiva que haya una fila seleccionada en la rejilla
        If frmApuntesPeriodicos.DgvApuper.CurrentRow Is Nothing Then Exit Sub

        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
        vTxtNombre = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(1).Value.ToString()

        ' Comprobamos si existe un identificador asociado.
        If ((frmEditarApuntesPeriodicos Is Nothing) OrElse (Not frmEditarApuntesPeriodicos.IsHandleCreated)) Then
            frmEditarApuntesPeriodicos = New EditarApuntesPeriodicos
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmEditarApuntesPeriodicos)

        ' Llamamos al formulario de manera modal en modo edición
        vEditar = "SI"
        frmEditarApuntesPeriodicos.ShowDialog()
        frmEditarApuntesPeriodicos.Dispose()

        RefrescarGridApuntesPeriodicos()

        ' Volvemos a pasar el rodillo matemático que calcula saldos e ingresos/gastos de forma limpia
        DgvApuntesPeriodicos()

        ' 2. REPOSICIONAMIENTO SEGURO: Volvemos a colocar el cursor en la fila editada
        ' Validamos que la fila siga existiendo tras el refresco para evitar desbordamientos
        If DgvApuper.Rows.Count > 0 AndAlso filaActual < DgvApuper.Rows.Count Then
            DgvApuper.Rows(filaActual).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(filaActual).Cells(0)
        End If
    End Sub

    Private Sub BtnBuscarRegistro_Click(sender As Object, e As EventArgs) Handles BtnBuscarRegistro.Click
        ' Llamamos al formulario de manera modal.
        frmBuscar.ShowDialog()
        BtnSeguirBuscando.Enabled = True

        vBuscar = frmBuscar.CmbTextoBuscar.Text
        vCampo = frmBuscar.CmbCampos.SelectedIndex
        vRow = 0
        For Each row As DataGridViewRow In DgvApuper.Rows
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
                vRow = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index
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
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
            ElseIf vCampo = 1 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
            ElseIf vCampo = 2 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(1)
            ElseIf vCampo = 3 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(2)
            ElseIf vCampo = 4 Then
                DgvApuper.Rows(vRow).Selected = True
                DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(5)
            End If
        End If
    End Sub

    Private Sub BtnSeguirBuscando_Click(sender As Object, e As EventArgs) Handles BtnSeguirBuscando.Click
        SeguirF3()
    End Sub

    Private Sub ApuntesPeriodicos_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        If BtnSeguirBuscando.Enabled = True Then
            If e.KeyCode = Keys.F3 Then
                SeguirF3()
            End If
        End If
    End Sub

    Private Sub SeguirF3()
        vCantidadFilas = DgvApuper.RowCount
        If vRow + 1 = vCantidadFilas Then
            MsgBox(resManager.GetString("MsgDatos2"))
            BtnSeguirBuscando.Enabled = False
        Else
            vContador = -1
            For Each row As DataGridViewRow In DgvApuper.Rows
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
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
                ElseIf vCampo = 1 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(0)
                ElseIf vCampo = 2 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(1)
                ElseIf vCampo = 3 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(2)
                ElseIf vCampo = 4 Then
                    DgvApuper.Rows(vRow).Selected = True
                    DgvApuper.CurrentCell = DgvApuper.Rows(vRow).Cells(5)
                End If
                vRowSeguir = 0
            End If
        End If
    End Sub

    Private Sub BtnAñadirRegistro_Click(sender As Object, e As EventArgs) Handles BtnAñadirRegistro.Click
        frmPrincipal.TsLabelFormulario.Text = rmse.GetString("IntroApuPer")
        ' Comprobamos si existe un identificador asociado.
        If ((frmIntroApuntesPeriodicos Is Nothing) OrElse (Not frmIntroApuntesPeriodicos.IsHandleCreated)) Then
            frmIntroApuntesPeriodicos = New IntroApuntesPeriodicos
        End If
        ' Forzar la traducción y el tamaño correcto antes de mostrar el formulario
        ActualizarTextosFormulario(frmIntroApuntesPeriodicos)
        ' Llamamos al formulario de manera modal.
        frmIntroApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmIntroApuntesPeriodicos.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("MsgFila1"))
        Else
            vFila = 0
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox(resManager.GetString("ToolTipAnterior"))
        Else
            vFila = vFilaActual - 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnSalir_Click_1(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = DgvApuper.RowCount - 1 Then
            MsgBox(resManager.GetString("ToolTipSiguiente"))
        Else
            vFila = vFilaActual + 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvApuper.CurrentRow.Index
        If vFilaActual = DgvApuper.RowCount - 1 Then
            MsgBox(resManager.GetString("MsgFila2"))
        Else
            vFila = DgvApuper.RowCount - 1
            DgvApuper.Rows(vFila).Selected = True
            DgvApuper.CurrentCell = DgvApuper.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnF6_Click(sender As Object, e As EventArgs) Handles BtnF6.Click
        ' =========================================================================
        ' 🌟 1. CONSULTA SQL MAESTRA RELACIONAL ALINEADA (¡Corregido!)
        ' =========================================================================
        ' Traemos las 11 celdas biológicas en su orden real simétrico.
        ' 🚀 LA CORRECCIÓN: Cambiamos conceptos.DescripcionCON por conceptos.CodigoCON en la segunda columna.
        vtipoSql = "SELECT apuper.FechaAPP As [FechaAPP], " &
                   "conceptos.CodigoCON As [ConceptoAPP], " &
                   "apuper.DescripcionAPP As [DescripcionAPP], " &
                   "apuper.ImporteAPP As [ImporteAPP], " &
                   "apuper.ImporteAPP As [SaldoAPP], " &
                   "apuper.NotasAPP As [NotasAPP], " &
                   "cuentas.NombreCUE As [CuentaAPP], " &
                   "apuper.CodigoAPP As [CodigoAPP], " &
                   "conceptos.CodigoCON As [CodigoCON], " &
                   "apuper.ConceptoAPP As [IdConceptoCON], " &
                   "apuper.CuentaAPP As [IdCuentaCUE], " &
                   "conceptos.TipoCON As [TipoCON] " & ' 🚀 LA CLAVE: Inyectamos el Tipo real en la posición 11
                   "FROM (apuper " &
                   "INNER JOIN conceptos ON apuper.ConceptoAPP = conceptos.IdConceptoCON) " &
                   "INNER JOIN cuentas ON apuper.CuentaAPP = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "
        vtipoSql += " ORDER BY apuper.FechaAPP ASC"

        vtipoGrid = "APUNTES_PERIODICOS"

        ' Volcamos los datos relacionales traducidos en tu DataGridView
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuper)

        LblNumRegistros.Text = resManager.GetString("SinFiltrar") ' My.Resources.Recursos.SinFiltrar
        BtnFiltroCuenta.Enabled = True
        BtnSinFiltroCuenta.Enabled = False
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        BtnFiltroFecha.Enabled = True
        BtnSinFiltroFecha.Enabled = False
        ' 1. Convertimos el año base de forma segura
        Dim anioBase As Integer
        If Not Integer.TryParse(vAñoEjercicio, anioBase) Then
            ' Si falla, usamos el año actual como salvavidas
            anioBase = Date.Today.Year
        End If

        ' 2. Calculamos los dos años que necesitas
        Dim anioInicio As Integer = anioBase
        Dim anioFin As Integer = anioBase + 20 ' Sumamos los 20 años de margen para los periódicos

        ' 3. Guardamos los valores en tus variables globales por si las usas luego
        vFecha1Enero = anioInicio
        vFecha31Diciembre = anioFin

        ' 4. Creamos las fechas exactas de inicio y fin
        Dim fechaInicio As New Date(anioInicio, 1, 1)
        Dim fechaFin As New Date(anioFin, 12, 31)

        ' 5. Configuramos los DateTimePicker con los rangos correctos
        DateTimePicker1.MinDate = fechaInicio
        DateTimePicker2.MinDate = fechaInicio

        DateTimePicker1.MaxDate = fechaFin
        DateTimePicker2.MaxDate = fechaFin

        ' 6. Asignamos los valores iniciales por defecto
        DateTimePicker1.Value = fechaInicio
        DateTimePicker2.Value = fechaFin

    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub DgvApuper_DoubleClick(sender As Object, e As EventArgs) Handles DgvApuper.DoubleClick
        BtnEditarRegistro.PerformClick()
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmTipoInformeApuntesPeriodicos Is Nothing) OrElse (Not frmTipoInformeApuntesPeriodicos.IsHandleCreated)) Then
            frmTipoInformeApuntesPeriodicos = New TipoInformeApuntesPeriodicos
        End If
        ' Forzar la traducción y el tamaño correcto antes de mostrar el formulario
        ActualizarTextosFormulario(frmTipoInformeApuntesPeriodicos)
        ' Llamamos al formulario de manera modal.
        frmTipoInformeApuntesPeriodicos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmTipoInformeApuntesPeriodicos.Dispose()
    End Sub

    Public Sub RefrescarGridApuntesPeriodicos()
        ' 🌟 SANEAMIENTO PREVENTIVO: Limpiamos la memoria de parámetros de consultas anteriores
        cmdMdb1cr.Parameters.Clear()

        ' Guardamos en booleanos el estado de tus botones de filtro de la pantalla
        Dim filtroCuentaActivo As Boolean = (BtnFiltroCuenta.Enabled = False)
        Dim filtroConceptoActivo As Boolean = (BtnFiltroConcepto.Enabled = False)
        Dim filtroFechaActivo As Boolean = (BtnFiltroFecha.Enabled = False)

        ' =========================================================================
        ' 🌟 CONSULTA SQL MAESTRA UNIFICADA DE 12 COLUMNAS (Inmune a Descalces)
        ' =========================================================================
        ' 🎯 LA CLAVE: Usamos CodigoCON en la celda 1 e inyectamos TipoCON en la celda 11
        vtipoSql = "SELECT apuper.FechaAPP As [FechaAPP], " &
                   "conceptos.CodigoCON As [ConceptoAPP], " &
                   "apuper.DescripcionAPP As [DescripcionAPP], " &
                   "apuper.ImporteAPP As [ImporteAPP], " &
                   "apuper.ImporteAPP As [SaldoAPP], " &
                   "apuper.NotasAPP As [NotasAPP], " &
                   "cuentas.NombreCUE As [CuentaAPP], " &
                   "apuper.CodigoAPP As [CodigoAPP], " &
                   "conceptos.CodigoCON As [CodigoCON], " &
                   "apuper.ConceptoAPP As [IdConceptoCON], " &
                   "apuper.CuentaAPP As [IdCuentaCUE], " &
                   "conceptos.TipoCON As [TipoCON] " &
                   "FROM (apuper " &
                   "INNER JOIN conceptos ON apuper.ConceptoAPP = conceptos.IdConceptoCON) " &
                   "INNER JOIN cuentas ON apuper.CuentaAPP = cuentas.IdCuentaCUE"

        vtipoSql += " WHERE apuper.EjercicioAPP <> 0 "

        ' 1. FILTRO POR ID NUMÉRICO DE CUENTA (SelectedValue puro de la RAM)
        If filtroCuentaActivo AndAlso CmbCuenta.SelectedValue IsNot Nothing Then
            Dim idCuentaSel As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)
            vtipoSql += $" And apuper.CuentaAPP = {idCuentaSel} "
        End If

        ' 2. FILTRO POR ID NUMÉRICO DE CONCEPTO (SelectedValue puro de la RAM)
        If filtroConceptoActivo AndAlso CmbConcepto.SelectedValue IsNot Nothing Then
            Dim idConceptoSel As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            vtipoSql += $" And apuper.ConceptoAPP = {idConceptoSel} "
        End If

        ' 3. FILTRO DE FECHAS PARÁMETRIZADO AL FINAL DE LA SQL
        If filtroFechaActivo Then
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"

            cmdMdb1cr.Parameters.AddWithValue("?", vDate1)
            cmdMdb1cr.Parameters.AddWithValue("?", vDate2)
        End If

        vtipoSql += " ORDER BY apuper.FechaAPP ASC"
        vtipoGrid = "APUNTES_PERIODICOS"

        ' Volcamos los datos relacionales traducidos en tu DataGridView
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        TraducirGridApuntesBD(Me.DgvApuper)
    End Sub

    ' =========================================================================
    ' 🌟 INTERCEPTOR DE ORDENACIÓN DINÁMICA DE REJILLA (¡Saldo recalculado en vivo!)
    ' =========================================================================
    Private Sub DgvApuper_Sorted(sender As Object, e As EventArgs) Handles DgvApuper.Sorted
        Try
            ' 🚀 LA ESTOCADA: Cuando el usuario reordene las columnas, obligamos a la app
            ' a limpiar el acumulador viejo de la RAM y recalcular los saldos de arriba a abajo.
            ' Llamamos directamente a tu función global del módulo de forma transparente.
            DgvApuntesPeriodicos()
        Catch ex As Exception
            ' Evita cualquier parpadeo visual en el hilo principal del formulario
        End Try
    End Sub

End Class