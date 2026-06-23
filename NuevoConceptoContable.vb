Imports System.Data
Imports System.Windows.Forms

Public Class NuevoConceptoContable
    Private dtConceptosMemoria As New DataTable()
    Public vtipoSql, vtipoGrid, vConcepto, tipoSql, vTxtNombre, vTxtDescripcion, vTxtTipo, vTxtNotas As String
    Public TL(4) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub NuevoConceptoContable_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.CmbTipoConcepto, rmse.GetString("SeleccionarTipo"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtDescripcion, rmse.GetString("IntroducirDescripcion"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtNombre, rmse.GetString("MsgNombre"))

        ' 1. Configuramos primero el estilo del ComboBox
        CmbTipoConcepto.DropDownStyle = ComboBoxStyle.DropDownList

        ' 2. LLENAMOS EL COMBO PRIMERO (Así tendrá elementos antes de seleccionar el índice 0)
        ActualizarIdiomaComboConcepto(Me.CmbTipoConcepto, False)

        ' 3. Ahora que ya tiene filas, seleccionamos de forma segura la primera
        CmbTipoConcepto.SelectedIndex = 0
        CmbTipoConcepto.Select()


        Try
            ' 1. Configuramos las columnas de la tabla en memoria si no existen
            If dtConceptosMemoria.Columns.Count = 0 Then
                dtConceptosMemoria.Columns.Add("CodigoCON", GetType(String))
                dtConceptosMemoria.Columns.Add("ConceptoTraducido", GetType(String))
            End If

            ' 2. Limpiamos cualquier dato residual de forma segura
            dtConceptosMemoria.Clear()

            ' 3. Cargamos los códigos estables desde tu base de datos Access
            Dim sqlCarga As String = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON"
            cmdMdb1cr.CommandText = sqlCarga
            cmdMdb1cr.Parameters.Clear()

            Using reader As OleDb.OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While reader.Read()
                    Dim codigoCON As String = reader("CodigoCON").ToString().Trim()
                    Dim textoTraducido As String = codigoCON.Replace("_", " ").ToUpper()

                    ' Si tu resManager global tiene traducción para este código, la aplicamos
                    If resManager IsNot Nothing Then
                        Dim claveRecurso As String = codigoCON.Replace(" ", "_")
                        Dim trad As String = resManager.GetString(claveRecurso)
                        If Not String.IsNullOrEmpty(trad) Then
                            textoTraducido = trad.Trim().ToUpper()
                        End If
                    End If

                    ' Guardamos el registro traducido en nuestra tabla local
                    dtConceptosMemoria.Rows.Add(codigoCON, textoTraducido)
                End While
            End Using
        Catch ex As Exception
            ' Previene caídas en el arranque si la base de datos está ocupada
        End Try

    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        ' 1. Controlar las mayúsculas sin perder la posición del cursor
        Dim posicionCursor As Integer = TxtNombre.SelectionStart
        Dim textoMayusculas As String = TxtNombre.Text.ToUpper()

        If TxtNombre.Text <> textoMayusculas Then
            TxtNombre.Text = textoMayusculas
            TxtNombre.SelectionStart = posicionCursor
        End If

        Dim vBusca As String = TxtNombre.Text.Trim()

        ' Si el cuadro de texto está vacío, limpiamos el Grid y lo ocultamos
        If vBusca = "" Then
            DgvExistente.DataSource = Nothing
            DgvExistente.Visible = False
            Exit Sub
        End If

        Try
            ' 2. FILTRADO: Buscamos en la columna traducida en memoria lo que coincida con la inicial
            Dim vistaFiltro As New DataView(dtConceptosMemoria)
            vistaFiltro.RowFilter = "ConceptoTraducido LIKE '" & vBusca & "%'"

            ' 3. ENLACE AL GRID: Si hay coincidencias, las mostramos en el mini-grid
            If vistaFiltro.Count > 0 Then
                ' Pasamos solo la tabla filtrada para que el Grid se refresque de golpe
                DgvExistente.DataSource = vistaFiltro.ToTable()

                ' Cambiamos el título de la cabecera al idioma del formulario
                DgvExistente.Columns("ConceptoTraducido").HeaderText = resManager.GetString("Existentes")

                ' Ocultamos el código técnico si se muestra de forma automática
                If DgvExistente.Columns.Contains("CodigoCON") Then
                    DgvExistente.Columns("CodigoCON").Visible = False
                End If

                DgvExistente.Visible = True
            Else
                DgvExistente.DataSource = Nothing
                DgvExistente.Visible = False
            End If

        Catch ex As Exception
            DgvExistente.DataSource = Nothing
            DgvExistente.Visible = False
        End Try
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtDescripcion.Select()
        End If
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtDescripcion.Select()
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtNota.Select()
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Guardar la palabra escrita en mayúsculas y sin espacios a los lados
        Dim nombreLimpio As String = TxtNombre.Text.Trim().ToUpper()

        If nombreLimpio <> "" Then

            ' Obtener de forma segura la traducción de "SALDO" del idioma actual (para los 6 idiomas)
            Dim saldoTraducido As String = ""
            Try
                saldoTraducido = rmse.GetString("PalabraSaldo").Trim().ToUpper()
            Catch ex As Exception
                saldoTraducido = "SALDO" ' Respaldo por si no se encuentra la clave en el recurso
            End Try

            ' 2. Validación de bloqueo: No permite "SALDO" en español ni su traducción internacional
            If nombreLimpio = "SALDO" OrElse (saldoTraducido <> "" AndAlso nombreLimpio = saldoTraducido) Then
                MessageBox.Show(rmse.GetString("NoNombreSaldo"),
                resManager.GetString("$this.Text"),
                MessageBoxButtons.OK,
                MessageBoxIcon.Information)
                TxtNombre.Select()
                TxtNombre.SelectAll()
                Exit Sub ' Detiene el guardado inmediatamente
            End If

            ' =========================================================================
            ' ¡BLOQUEO DE DUPLICADOS MULTIIDIOMA REAL CON DATAVIEW!
            ' =========================================================================
            ' Buscamos si lo que el usuario ha escrito ya existe en la columna traducida en memoria
            Try
                Dim vistaValidar As New DataView(dtConceptosMemoria)
                ' Buscamos una coincidencia exacta e insensible a mayúsculas/minúsculas
                vistaValidar.RowFilter = "ConceptoTraducido = '" & nombreLimpio.Replace("'", "''") & "'"

                If vistaValidar.Count > 0 Then
                    ' Si encuentra registros, significa que el concepto YA EXISTE (sea en el idioma que sea)
                    MsgBox(resManager.GetString("Nombre") & ":  " & TxtNombre.Text.Trim() & ", " & rmse.GetString("YaExisteConcepto"), vbOKOnly + vbExclamation, rmse.GetString("$this.Text"))
                    TxtNombre.Select()
                    TxtNombre.SelectAll()
                    Exit Sub ' Detiene el guardado de inmediato
                End If
            Catch ex As Exception
                ' Si falla la validación en memoria por seguridad, dejamos que continúe
            End Try
            ' =========================================================================

            ' Si pasa todas las validaciones, preparamos el resto de variables
            ' ¡Truco de consistencia!: Guardamos el código en la base de datos normalizado (espacios por guiones)
            ' para que cuando se genere el .resx la clave sea limpia ("LUZ_Y_AGUA" en vez de "LUZ Y AGUA")
            Dim codigoEstableBD As String = nombreLimpio.Replace(" ", "_")

            vTxtNombre = TxtNombre.Text.Trim()
            vTxtDescripcion = ApostrofePorAcentoAgudo(TxtDescripcion.Text)
            vTxtNotas = TxtNota.Text

            ' 3. Guardado multiidioma del tipo (Inamovible en la BD como GASTO o INGRESO)
            If CmbTipoConcepto.SelectedIndex = 0 Then
                vTxtTipo = "GASTO"
            ElseIf CmbTipoConcepto.SelectedIndex = 1 Then
                vTxtTipo = "INGRESO"
            Else
                vTxtTipo = CmbTipoConcepto.Text ' Respaldo en caso de que cambies el orden
            End If

            ' =========================================================================
            ' NUEVO: CALCULAR EL SIGUIENTE ID DISPONIBLE (MAX + 1)
            ' =========================================================================
            Dim siguienteID As Integer = 1
            Try
                cmdMdb1cr.CommandText = "SELECT MAX(IdConceptoCON) FROM conceptos"
                cmdMdb1cr.Parameters.Clear()
                Dim resultado As Object = cmdMdb1cr.ExecuteScalar()
                If resultado IsNot DBNull.Value AndAlso resultado IsNot Nothing Then
                    siguienteID = Convert.ToInt32(resultado) + 1
                End If
            Catch ex As Exception
                siguienteID = 1 ' Respaldo por si la tabla estuviera vacía
            End Try
            ' =========================================================================

            ' 4. INSERCIÓN TOTALMENTE PARAMETRIZADA Y SEGURA (¡Añadimos el ID!)
            vtipoSql = "INSERT INTO conceptos (IdConceptoCON, CodigoCON, DescripcionCON, TipoCON, NotasCON) VALUES (?, ?, ?, ?, ?)"
            cmdMdb1cr.CommandText = vtipoSql

            ' Limpiamos y asignamos los parámetros en el orden ESTRICTO del SQL
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("@IdConceptoCON", siguienteID) ' <--- Inyectamos el nuevo número entero
            cmdMdb1cr.Parameters.AddWithValue("@CodigoCON", codigoEstableBD)
            cmdMdb1cr.Parameters.AddWithValue("@DescripcionCON", vTxtDescripcion.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@TipoCON", vTxtTipo.Trim())
            cmdMdb1cr.Parameters.AddWithValue("@NotasCON", vTxtNotas.Trim())

            Try
                cmdMdb1cr.ExecuteNonQuery()

                ' ¡TRUCO MAESTRO!: Añadimos el nuevo concepto también a tu tabla en memoria 
                ' para que si el usuario sigue escribiendo, el filtro TextChanged lo detecte al vuelo
                If dtConceptosMemoria IsNot Nothing Then
                    dtConceptosMemoria.Rows.Add(codigoEstableBD, vTxtNombre.ToUpper())
                End If

                Me.Close()
            Catch ex As Exception
                MsgBox(ex.ToString(), vbCritical, resManager.GetString("Error"))
            End Try
        Else
            ' Mensaje de error si el campo nombre está completamente vacío
            MsgBox(rmse.GetString("MsgDatosNombre"), vbCritical, rmse.GetString("$this.Text"))
            TxtNombre.Select()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtNombre_GotFocus(sender As Object, e As EventArgs) Handles TxtNombre.GotFocus
        PintaTxt()
    End Sub

    Public Sub PintaTxt()
        Dim Texto As TextBox
        Texto = Me.ActiveControl
        Texto.SelectionStart = 0
        Texto.SelectionLength = Texto.Text.Length
    End Sub

End Class