Imports System.Data
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class EditarApuntes

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU As String
    Public vCodigoAPU As Integer
    Public vimporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(8) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargandoFormulario = True
        Me.KeyPreview = True

        Label7.Text = vMoneda
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEliminar, resManager.GetString("ToolTipEliminar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, rmse.GetString("ToolTipSeleccionarConcepto"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, rmse.GetString("ToolTipSeleccionarCuenta"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, rmse.GetString("ToolTipSeleccionarDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ToolTipModificaImporte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))

        ' Establece el rango usando tu variable vAñoEjercicio
        DateTimePicker1.MinDate = New DateTime(vAñoEjercicio, 1, 1)
        DateTimePicker1.MaxDate = New DateTime(vAñoEjercicio, 12, 31)

        ' Llenar los Combo
        '*****************
        Try
            ' Usamos la función exclusiva que no carga los 'ESPECIALES' si es para introducir/editar ordinarios
            ' (O la que uses en este formulario, pero asegurando que use DataTable)
            LlenarComboConceptosIntroApuntes(Me.CmbConcepto)
            LlenarComboCuentasGenerico(Me.CmbCuenta)
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorCargarCONyCUE") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        ' Llenar el Combo Descripción (Optimizado a la velocidad del rayo)
        '*****************************************************************
        ' 🌟 LA CORRECCIÓN MAESTRA: Usamos DISTINCT para que Access elimine los duplicados 
        ' en milisegundos y filtramos directamente para excluir "Saldo Inicial" desde la BD
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes " &
                                "WHERE DescripcionAPU <> 'Saldo Inicial' " &
                                "And DescripcionAPU Is Not Null " &
                                "ORDER BY DescripcionAPU ASC"

        CmbDescripcion.Items.Clear()
        cmdMdb1cr.Parameters.Clear() ' Saneamiento preventivo de parámetros

        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()

            ' 🌟 CERO BUCLES CRUZADOS: Insertamos los datos limpios directamente
            While drMdb1.Read()
                Dim descLimpia As String = drMdb1("DescripcionAPU").ToString().Trim()
                If Not String.IsNullOrEmpty(descLimpia) Then
                    CmbDescripcion.Items.Add(descLimpia)
                End If
            End While

            drMdb1.Close()
        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(resManager.GetString("ErrorLlenarCmbDescripcion") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try

        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        DateTimePicker1.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(0).Value.ToString
        CmbConcepto.SelectedValue = Convert.ToInt32(frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(9).Value)
        CmbCuenta.SelectedValue = Convert.ToInt32(frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(10).Value)
        CmbDescripcion.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
        vimporteAPU = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(3).Value
        TxtImporte.Text = Convert.ToDecimal(vimporteAPU).ToString("N2")
        TxtNota.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(5).Value.ToString
        vCodigoAPU = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(7).Value

        If vEditar = "SI" Then
            'LblEditando.Text = "EDITANDO APUNTE CONTABLE"
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = rmse.GetString("LblEliminando")
            BtnHoy.Enabled = False
            BtnCalculadora.Enabled = False
            DateTimePicker1.Enabled = False
            CmbConcepto.Enabled = False
            CmbDescripcion.Enabled = False
            TxtImporte.Enabled = False
            CmbCuenta.Enabled = False
            TxtNota.Enabled = False
            BtnAceptar.Enabled = False
            BtnEliminar.Select()
        End If

        ' Truco del vaivén maestro: Obligamos a que pinte la descripción larga UNA SOLA VEZ en el arranque
        If CmbConcepto.Items.Count > 0 Then
            ' Si tienes guardado qué concepto tenía la fila, lo heredas (ej: asignando su posición o ID)
            ' Si no, forzamos un refresco limpio del primer elemento para rellenar los cuadros de texto
            Dim indiceGuardado As Integer = CmbConcepto.SelectedIndex
            CmbConcepto.SelectedIndex = -1
            CmbConcepto.SelectedIndex = If(indiceGuardado >= 0, indiceGuardado, 0)
        End If

        ' =========================================================================
        ' 🌟 LA CORRECCIÓN CLAVE: FORZAMOS A QUE LEA EL TEXTO REAL DE LA REJILLA
        ' =========================================================================
        ' Después de que el combo haga sus movimientos, le imponemos a la fuerza el 
        ' texto exacto que el usuario está viendo en la fila de la pantalla principal.
        ' (Asegúrate de que la Celda 2 corresponde a la columna DescripcionAPU en tu Dgv)
        Dim filaPrincipal As Integer = frmApuntesContables.DgvApuntes.CurrentRow.Index
        Dim descripcionRealGrid As String = frmApuntesContables.DgvApuntes.Rows(filaPrincipal).Cells(2).Value.ToString()

        CmbDescripcion.Text = descripcionRealGrid
        cargandoFormulario = False
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 🌟 SANEAMIENTO INDISPENSABLE: Eliminamos las variables de texto plano antiguas (vConceptoAPU, etc.)
        ' Guardamos únicamente el objeto Date puro para el control del ejercicio si lo necesitas
        vDate3 = DateTimePicker1.Value.Date

        ' =========================================================================
        ' CONTROL DE DECIMALES E IMPORTES (Mantenido intacto por seguridad contable)
        ' =========================================================================
        Dim importeDecimal As Decimal = 0.0D
        Dim textoLimpio As String = TxtImporte.Text.Trim()

        ' Intenta leer con la cultura regional del usuario (respeta su panel de control)
        If Not Decimal.TryParse(textoLimpio,
                        System.Globalization.NumberStyles.Number,
                        System.Globalization.CultureInfo.CurrentCulture,
                        importeDecimal) Then

            ' PLAN B SEGURO: Intentamos con la cultura invariante (punto decimal universal)
            If Not Decimal.TryParse(textoLimpio,
                            System.Globalization.NumberStyles.Number,
                            System.Globalization.CultureInfo.InvariantCulture,
                            importeDecimal) Then
                importeDecimal = 0.0D
            End If
        End If

        ' Guardamos el importe limpio y aplicamos el signo contable según el tipo
        vimporteAPU = importeDecimal

        ' =========================================================================
        ' 🎯 VALIDACIÓN DEL SIGNO INMUNE A IDIOMAS (Castellano / Catalán / Inglés)
        ' =========================================================================
        Dim tipoDeFabrica As String = ""

        ' Extraemos de forma segura el TipoCON original oculto en la RAM del combo
        If CmbConcepto.SelectedItem IsNot Nothing Then
            Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)
            If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                tipoDeFabrica = filaSeleccionada("TipoCON").ToString().Trim().ToUpper()
            End If
        End If

        ' Evaluamos el signo usando el dato rígido de la Base de Datos, nunca el texto de la pantalla
        If tipoDeFabrica = "GASTO" Then
            vimporteAPU = -Math.Abs(vimporteAPU)
        Else
            vimporteAPU = Math.Abs(vimporteAPU)
        End If

        ' =========================================================================
        ' 🌟 AQUÍ ENLAZA TU CONSULTA UPDATE PARAMETRIZADA CON IDs NUMÉRICOS
        ' =========================================================================
        Dim idConceptoEditado As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
        Dim idCuentaEditada As Integer = Convert.ToInt32(CmbCuenta.SelectedValue)

        Dim sqlUpdate As String = "UPDATE apuntes SET " &
                                  "FechaAPU = ?, " &
                                  "ConceptoAPU = ?, " &
                                  "DescripcionAPU = ?, " &
                                  "ImporteAPU = ?, " &
                                  "NotasAPU = ?, " &
                                  "CuentaAPU = ?, " &
                                  "EjercicioAPU = ? " &
                                  "WHERE CodigoAPU = ?" ' ⬅️ Tu campo ID único en Access (vCodigoAPU)

        cmdMdb1cr.CommandText = sqlUpdate
        cmdMdb1cr.Parameters.Clear()

        ' Inyectamos los parámetros en estricto orden biológico de los signos '?'
        cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vDate3
        cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConceptoEditado
        cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = CmbDescripcion.Text.Trim()
        cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = Convert.ToDecimal(vimporteAPU)
        cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = TxtNota.Text.Trim()
        cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuentaEditada
        cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)

        ' Clave primaria del WHERE al final (Usando tu Convert.ToInt32 limpio sin CInt)
        cmdMdb1cr.Parameters.Add("@id", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vCodigoAPU)

        Try
            cmdMdb1cr.ExecuteNonQuery()
            Me.Close() ' Cerramos la ventana modal al terminar con éxito
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorActualizarApunte") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click
        ' Lanzamos la pregunta de confirmación contable de seguridad
        ' 🚀 REPARADO MODO PREMIUM: Sustituimos el MsgBox rígido por tu motor elástico traducido
        Dim respuesta As MsgBoxResult = ConfirmarAccionTraducida(rmse.GetString("MsgBoxEliminarApunte"), rmse.GetString("$this.Text"))

        If respuesta = vbYes Then
            ' Ejecutamos el borrado físico usando el identificador único del registro actual
            vtipoSql = "DELETE FROM apuntes WHERE apuntes.CodigoAPU = " & CInt(vCodigoAPU)
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                ' Cerramos la ventana tras borrar
                Me.Close()
            Catch ex As Exception
                MsgBox(rmse.GetString("MsgBoxErrorEliminarRegistro") & ": " & ex.Message, MsgBoxStyle.Critical, rmse.GetString("$this.Text"))
            End Try
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 Quitamos el escudo de carga solo para el Tipo de Concepto, para que pinte al editar
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            ' 3. Traducir y asignar los textos a la interfaz de forma segura
            If Not String.IsNullOrEmpty(codigoOriginal) Then
                vConcepto = codigoOriginal

                ' --- TRADUCIR EL TIPO (Gasto / Ingreso / Especial) ---
                Dim tradTipo As String = ""
                Select Case tipoOriginal.ToUpper()
                    Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                    Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                    Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                End Select
                If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal

                ' 🎯 INYECCIÓN BLINDADA: Rellenamos el cuadro de la derecha sí o sí
                TxtTipoConcepto.Text = tradTipo

                ' 🌟 ESCUDO REDUCIDO: Las descripciones solo cambian si el usuario interactúa con el ratón
                If cargandoFormulario Then Exit Sub

                ' --- TRADUCIR LAS DESCRIPCIONES ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal
                CmbDescripcion.Text = tradDesc
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorSincronizarCON") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        ' 1. 🛡️ EL ESCUDO UNIVERSAL ADMITE TODO: Números, borrar (Control), punto, coma o el Intro
        If Not Char.IsDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) AndAlso e.KeyChar <> "."c AndAlso e.KeyChar <> ","c AndAlso e.KeyChar <> ChrW(Keys.Enter) Then
            e.Handled = True
            Exit Sub
        End If

        ' 2. 🎯 AL PULSAR INTRO: Pasamos el rodillo internacional e inyectamos en la variable de apuntes
        If e.KeyChar = ChrW(Keys.Enter) Then
            e.Handled = True

            ' Invocamos tu función global centralizada: Cero grasa digital en la RAM
            Dim importeFinal As Decimal = ParsearImporteUniversal(TxtImporte.Text)

            ' Guardamos de forma segura en tu variable global de doble precisión (vImporteAPU)
            vimporteAPU = Convert.ToDouble(importeFinal)

            ' Formateamos la caja visual con el estándar de dos decimales de gala
            TxtImporte.Text = importeFinal.ToString("N2")

            ' Mandamos el cursor directo al combo de la Cuenta de forma dócil
            CmbCuenta.Select()
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
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs)
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

End Class