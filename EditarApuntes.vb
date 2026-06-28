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
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ToolTipIngresarImporte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))

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
        cargandoFormulario = False

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
        TxtImporte.Text = Math.Abs(vimporteAPU).ToString("N2")
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
        cargandoFormulario = False
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

        ' Usamos ToUpper para que la validación del signo sea 100% inmune a mayúsculas/minúsculas
        If TxtTipoConcepto.Text.ToUpper() = "GASTO" Then
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
        Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("MsgBoxEliminarApunte"), vbQuestion + vbYesNo + vbDefaultButton2, rmse.GetString("$this.Text"))
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
        ' 🌟 ESCUDO DE CARGA: Si el formulario se está iniciando o limpiando, salimos de inmediato
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim codigoOriginal As String = ""
            Dim descripcionOriginal As String = ""
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access)
            ' Como el combo está enlazado a un DataTable, convertimos el ítem actual en un DataRowView
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                codigoOriginal = filaSeleccionada("CodigoCON").ToString().Trim()
                descripcionOriginal = filaSeleccionada("DescripcionCON").ToString().Trim()

                ' Leemos el TipoCON de forma segura desde la memoria de la app
                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            ' 3. Traducir y asignar los textos a la interfaz de forma segura sin tocar la BD
            If Not String.IsNullOrEmpty(codigoOriginal) Then
                vConcepto = codigoOriginal ' Guardamos el código original en español para tus lógicas

                ' --- TRADUCIR EL TIPO (Gasto / Ingreso / Especial) ---
                Dim tradTipo As String = ""
                Select Case tipoOriginal.ToUpper()
                    Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                    Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                    Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
                End Select
                If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
                TxtTipoConcepto.Text = tradTipo

                ' --- TRADUCIR LAS DESCRIPCIONES (Desc_NOMBRE) ---
                Dim llaveDesc As String = "Desc_" & codigoOriginal.Replace(" ", "_")
                Dim tradDesc As String = resManager.GetString(llaveDesc)

                ' Si no tiene traducción en el ResX, dejamos la descripción original de la BD
                If String.IsNullOrEmpty(tradDesc) Then tradDesc = descripcionOriginal

                CmbDescripcion.Text = tradDesc
                ' Si tienes TxtDescripcion en este form, lo rellenas; si no, deja solo el combo
                ' TxtDescripcion.Text = tradDesc 
            End If

        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorSincronizarCON") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
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