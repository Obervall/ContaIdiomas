Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class EditarApuntesPeriodicos

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU As String
    Public vCodigoAPU As Integer
    Public vimporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarApuntesPeriodicos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Label7.Text = vMoneda
        Dim TL(8) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEliminar, resManager.GetString("ToolTipEliminar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, frmEditarApuntes.rmse.GetString("ToolTipSeleccionarConcepto"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, frmEditarApuntes.rmse.GetString("ToolTipSeleccionarCuenta"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, frmEditarApuntes.rmse.GetString("ToolTipSeleccionarDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, frmEditarApuntes.rmse.GetString("ToolTipIngresarImporte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))

        ' =========================================================================
        ' 🌟 CARGA DE COMBOS DE LA NUEVA ERA (Con Idiomas, Orden A-Z e IDs Numéricos)
        ' =========================================================================
        ' 1. Encendemos tu escudo protector antes de rellenar los componentes
        cargandoFormulario = True
        cmdMdb1cr.Parameters.Clear()

        Try
            ' 2. LLAMADAS SEGURAS: Usamos tus funciones de módulo que cargan DataTables e IDs en microsegundos
            LlenarComboConceptosSueltosBD(Me.CmbConcepto)
            LlenarComboCuentasGenerico(Me.CmbCuenta)
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorIniciaDesplegables") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' 3. Llenar el Combo Descripción (Optimizado a la velocidad del rayo con DISTINCT directo de Access)
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'Saldo Inicial' And DescripcionAPU Is Not Null ORDER BY DescripcionAPU ASC"
        CmbDescripcion.Items.Clear()
        Try
            Using dr As OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While dr.Read()
                    Dim descLimpia As String = dr("DescripcionAPU").ToString().Trim()
                    If Not String.IsNullOrEmpty(descLimpia) Then
                        CmbDescripcion.Items.Add(descLimpia)
                    End If
                End While
            End Using
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorCargarDescripciones") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' =========================================================================
        ' 🌟 SINCRONIZACIÓN BIOLÓGICA CON LA REJILLA DE ATRÁS POR IDs NUMÉRICOS
        ' =========================================================================
        filaActual = frmApuntesPeriodicos.DgvApuper.CurrentRow.Index

        ' Volcamos los textos directos y la fecha limpia
        DateTimePicker1.Value = Convert.ToDateTime(frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(0).Value)
        CmbDescripcion.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(2).Value.ToString()
        TxtNota.Text = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(5).Value.ToString()

        ' Rescatamos el ID Autonumérico único de este apunte periódico (Celda 7)
        vCodigoAPU = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(7).Value

        ' 🌟 CORRECCIÓN MAESTRA: Seleccionamos en los combos por su ID numérico oculto
        ' usando SelectedValue. Así viajan emparejados de forma indestructible.
        Dim idConceptoFila As Integer = Convert.ToInt32(frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(9).Value)
        Dim idCuentaFila As Integer = Convert.ToInt32(frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(10).Value)

        CmbConcepto.SelectedValue = idConceptoFila
        CmbCuenta.SelectedValue = idCuentaFila

        ' 🚀 REPARADO 1: Forzamos la actualización manual del TxtTipoConcepto en el Load
        ' Vamos a buscar el tipo (Gasto/Ingreso) al maestro de conceptos usando el ID relacional
        Using con As New OleDbConnection(conexion1.ConnectionString)
            Using cmd As New OleDbCommand("SELECT TipoCON FROM conceptos WHERE IdConceptoCON = ?", con)
                cmd.Parameters.Add("@id", OleDbType.Integer).Value = idConceptoFila
                Try
                    con.Open()
                    Dim res = cmd.ExecuteScalar()
                    If res IsNot Nothing Then
                        ' Traducimos e inyectamos el tipo ("G" o "I") en tu TxtTipoConcepto local
                        TxtTipoConcepto.Text = res.ToString().Trim().ToUpper()
                    End If
                Catch
                End Try
            End Using
        End Using

        ' 🚀 REPARADO 2: Control estricto del signo decimal (¡Adiós Math.Abs transgresor!)
        vimporteAPU = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(3).Value.ToString()
        Dim importeDecimal As Decimal = ConvertirDecimalSeguro(vimporteAPU)

        ' Conservamos el signo menos si es un gasto menor que cero para que no mute a positivo
        TxtImporte.Text = importeDecimal.ToString("N2")

        ' 4. Apagamos el escudo protector de forma segura tras la asignación
        cargandoFormulario = False

        ' =========================================================================
        ' INTERFAZ ESTÉTICA SEGÚN MODO: EDICIÓN O ELIMINACIÓN (Lógica original perfecta)
        ' =========================================================================
        If vEditar = "SI" Then
            LblEditando.Text = rmse.GetString("LblEditando.Text")
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
            DateTimePicker1.Enabled = False
            CmbConcepto.Enabled = False
            CmbDescripcion.Enabled = False
            TxtImporte.Enabled = False
            CmbCuenta.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' 1. Cuadro de confirmación original traducido desde tus recursos (.resx)
        Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("SeguroEliminarRegistro"), rmse.GetString("$this.Text"))
        If respuesta = vbYes Then
            ' 🌟 BORRADO FÍSICO PARAMETRIZADO INDESTRUCTIBLE (Nueva Era Relacional)
            ' Usamos el comodín '?' para inyectar el ID de forma nativa en el motor de Access
            cmdMdb1cr.CommandText = "DELETE FROM apuper WHERE CodigoAPP = ?"
            cmdMdb1cr.Parameters.Clear()

            ' Pasamos el ID Autonumérico único de este apunte periódico (vCodigoAPU)
            cmdMdb1cr.Parameters.Add("@id", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vCodigoAPU)

            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox(rmse.GetString("RegistroApuntePeriódicoBorrado"))
            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorEliminarApuntePeriodico") & ": " & ex.Message, MsgBoxStyle.Critical)
            End Try
        Else
            ' Si cancela, devolvemos el foco a su sitio original en la rejilla de atrás
            If frmApuntesPeriodicos.DgvApuper.Rows.Count > 0 AndAlso filaActual < frmApuntesPeriodicos.DgvApuper.Rows.Count Then
                frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Selected = True
                frmApuntesPeriodicos.DgvApuper.CurrentCell = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(0)
            End If
        End If

        Me.Close()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Si el formulario está cargando o limpiando, salimos en microsegundos
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access)
            ' Convertimos el ítem seleccionado en un DataRowView para leer sus columnas ocultas en la RAM
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            If Not String.IsNullOrEmpty(codigoOriginal) Then
                ' Almacenamos el código en tu variable global para tus lógicas de fábrica
                vConcepto = codigoOriginal

                ' --- TRADUCIR EL TIPO (Gasto / Ingreso / Especial) ---
                Dim tradTipo As String = ""
                Select Case tipoOriginal.ToUpper()
                    Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                    Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                    Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                End Select
                If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
                TxtTipoConcepto.Text = tradTipo

                ' --- TRADUCIR LAS DESCRIPCIONES AUTOMÁTICAS (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción genérica de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal
                CmbDescripcion.Text = tradDesc
            End If

        Catch ex As Exception
            ' Evita cuelgues visuales si el combo parpadea en la carga
        End Try
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            CmbCuenta.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
        If TxtImporte.Text = "" Then
            TxtImporte.Text = 0
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Normalización inicial del cuadro de texto
        If String.IsNullOrWhiteSpace(TxtImporte.Text) Then TxtImporte.Text = "0"

        ' 2. Extraemos el importe de forma segura usando tu función del módulo
        Dim importeDecimal As Decimal = ConvertirDecimalSeguro(TxtImporte.Text)

        If importeDecimal <> 0 Then
            ' =========================================================================
            ' 🌟 CORTAFUEGOS DE SIGNO CONTABLE: COMPARACIÓN RELACIONAL INDESTRUCTIBLE
            ' =========================================================================
            ' 🚀 REPARADO: Añadimos el control por la letra "G" pura que inyecta Access en la RAM
            Dim tipoConceptoTxt As String = TxtTipoConcepto.Text.Trim().ToUpper()
            Dim traduccionGasto As String = If(resManager?.GetString("Tipo_Gasto"), "GASTO").Trim().ToUpper()

            If tipoConceptoTxt = "G" OrElse tipoConceptoTxt = "GASTO" OrElse tipoConceptoTxt = traduccionGasto Then
                vimporteAPU = -Math.Abs(importeDecimal) ' Forzamos signo negativo para Gastos
            Else
                vimporteAPU = Math.Abs(importeDecimal)  ' Forzamos signo positivo para Ingresos
            End If

            vDate3 = DateTimePicker1.Value.Date

            ' 🌟 EXTRAEMOS LOS IDs NUMÉRICOS PUROS DESDE LOS COMBOS (Nueva era relacional)
            Dim idConceptoEditado As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
            Dim idCuentaEditada As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

            ' =========================================================================
            ' 🌟 FASE 1: EJECUCIÓN DEL UPDATE PARAMETRIZADO CON IDs
            ' =========================================================================
            vtipoSql = "UPDATE apuper SET FechaAPP = ?, ConceptoAPP = ?, DescripcionAPP = ?, ImporteAPP = ?, CuentaAPP = ?, NotasAPP = ? " &
                       "WHERE CodigoAPP = ?"
            cmdMdb1cr.CommandText = vtipoSql

            ' Los parámetros de Access se asocian estrictamente por el orden de los comodines '?'
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vDate3
            cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoEditado ' ID Numérico
            cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = CmbDescripcion.Text.Trim()

            ' Forzamos formato Currency para evitar conflictos de precisión decimal en Access
            Dim paramImp As OleDb.OleDbParameter = cmdMdb1cr.Parameters.Add("@ImporteAPP", OleDb.OleDbType.Currency)
            paramImp.Value = Math.Round(vimporteAPU, 2)

            cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuentaEditada     ' ID Numérico
            cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = TxtNota.Text.Trim()
            cmdMdb1cr.Parameters.Add("@id", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vCodigoAPU)

            Try
                cmdMdb1cr.ExecuteNonQuery()
            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorActualizarApunte") & ": " & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub ' Si falla el guardado, detenemos el flujo para no corromper la grilla
            End Try

            ' =========================================================================
            ' 🌟 FASE 2: REUTILIZACIÓN TOTAL (Refresco de la rejilla nodriza)
            ' =========================================================================
            frmApuntesPeriodicos.RefrescarGridApuntesPeriodicos()

            ' Reposicionamos la fila seleccionada por el usuario de forma 100% segura
            If frmApuntesPeriodicos.DgvApuper.Rows.Count > 0 AndAlso filaActual < frmApuntesPeriodicos.DgvApuper.Rows.Count Then
                frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Selected = True
                frmApuntesPeriodicos.DgvApuper.CurrentCell = frmApuntesPeriodicos.DgvApuper.Rows(filaActual).Cells(0)
            End If

            ' Cerramos la ventana modal de edición con éxito
            Me.Close()
        Else
            MessageBox.Show(frmIntroApuntes.rmse.GetString("NoQuantityAmount"), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            TxtImporte.Select()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCalculadora_Click(sender As Object, e As EventArgs) Handles BtnCalculadora.Click
        Dim Proceso As New Process()
        Proceso.StartInfo.FileName = "calc.exe"
        Proceso.StartInfo.Arguments = ""
        Proceso.Start()
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs)
        frmPrincipal.TsLabelFormulario.Text = frmConceptosContables.rmse.GetString("$this.Text")

        ' Comprobamos si existe un identificador asociado.
        If ((frmConceptosContables Is Nothing) OrElse (Not frmConceptosContables.IsHandleCreated)) Then
            frmConceptosContables = New ConceptosContables
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmNuevoConceptoContable)
        ' Llamamos al formulario de manera modal.
        frmConceptosContables.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmConceptosContables.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs)
        frmPrincipal.TsLabelFormulario.Text = frmCuentasBancarias.rmse.GetString("$this.Text")

        ' Comprobamos si existe un identificador asociado.
        If ((frmCuentasBancarias Is Nothing) OrElse (Not frmCuentasBancarias.IsHandleCreated)) Then
            frmCuentasBancarias = New CuentasBancarias
        End If
        ' 3. Forzar la traducción y el tamaño correcto antes de medir la ventana
        ActualizarTextosFormulario(frmCuentasBancarias)
        ' Llamamos al formulario de manera modal.
        frmCuentasBancarias.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmCuentasBancarias.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class