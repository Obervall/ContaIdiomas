Imports System.Diagnostics
Imports System.Drawing
Imports System.Globalization
Imports System.Linq
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntes

    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, strText, vIntro, vLetras, vCombo, vDescripcion As String
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(13) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Label7.Text = vMoneda
        vIntro = "NO"
        vFecha1Enero = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        If vAñoEjercicio <> vAñoActual Then
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If

        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptarOtro, rmse.GetString("BtnAceptarOtro.Text"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptarSalir, rmse.GetString("BtnAceptarSalir.Text"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, rmse.GetString("BtnCancelar.Text") & " " & rmse.GetString("$this.Text"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, rmse.GetString("SelecConcepto"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, rmse.GetString("SelecCuenta"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, rmse.GetString("SelecDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteAsiento"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, rmse.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, resManager.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnDescripcion, rmse.GetString("BtnDescripcion"))
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.TxtBuscarLetras, rmse.GetString("TxtABuscar"))
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.BtnAyuda, rmse.GetString("BtnAyuda"))

        CmbConcepto.DropDownStyle = ComboBoxStyle.DropDownList

        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbConcepto.Items.Add(drMdb1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            'MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try

        'LlenarConcepto()
        ' 3. Ahora que ya tiene filas, seleccionamos de forma segura la primera
        'CmbConcepto.SelectedIndex = 0

        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()

        ' Llenar el Combo Cuenta
        '***********************
        drMdb1.Close()
        cmdMdb1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuenta.Items.Add(drMdb1.GetValue(0))
                End While
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    CmbCuenta.Text = CmbCuenta.Items(frmApuntesContables.CmbCuenta.SelectedIndex)
                Else
                    CmbCuenta.Text = CmbCuenta.Items(0)
                End If
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        TxtImporte.Text = 0
    End Sub

    Private Sub CmbConcepto_MouseClick(sender As Object, e As MouseEventArgs) Handles CmbConcepto.MouseClick
        TxtBuscarLetras.Enabled = False
        vIntro = "NO"
        ' Solo forzamos el despliegue automático si el usuario NO ha pulsado la flecha nativa
        ' (Nos aseguramos comprobando si la lista ya está abierta o abriéndola suavemente)
        If CmbConcepto.Items.Count <> 0 AndAlso Not CmbConcepto.DroppedDown Then
            CmbConcepto.DroppedDown = True
            CmbConcepto.SelectedIndex = 0
        End If
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado para mostrar su descripción y tipo en los cuadros de abajo
        '*****************************************************************************************************
        If vIntro = "NO" Then
            vConcepto = CmbConcepto.Text.ToString
            drMdb1.Close()
            cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' ORDER BY conceptos.CodigoCON ASC"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            drMdb1.Read()
            If drMdb1.HasRows Then
                TxtTipoConcepto.Text = drMdb1.GetValue(2)
                CmbDescripcion.Text = drMdb1.GetValue(1)
                TxtDescripcion.Text = drMdb1.GetValue(1)
                drMdb1.Close()
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtBuscarLetras.KeyDown
        ' Si el usuario pulsa FLECHA ABAJO desde el cuadro de búsqueda, saltamos al combo de forma inteligente
        If e.KeyCode = Keys.Down Then
            If vCombo = "descripcion" AndAlso CmbDescripcion.Items.Count > 0 Then
                e.Handled = True
                CmbDescripcion.Focus()
                CmbDescripcion.SelectedIndex = 0
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtBuscarLetras.KeyPress
        ' Detectamos si se pulsa ENTER (Ascii 13)
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True ' Evita el molesto pitido de Windows

            ' 🛠️ CORRECCIÓN DE ANIDACIÓN: Cambiado a un IF limpio y directo para corregir el teclado
            If vCombo = "descripcion_vacia" Then
                ' Bloque de alta de descripción nueva
                Dim respuesta As MsgBoxResult = MsgBox(rmse.GetString("NoExistenDescripciones") & ": -" & TxtBuscarLetras.Text.ToUpper() & "-" & vbCrLf & "¿" & rmse.GetString("AñadimosDescripcion") & "?", vbQuestion + vbYesNo + vbDefaultButton1, rmse.GetString("$this.Text"))

                If respuesta = vbYes Then
                    vIntro = "SI"
                    CmbDescripcion.Text = TxtBuscarLetras.Text
                    vDescripcion = TxtBuscarLetras.Text

                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    TxtBuscarLetras.Text = ""
                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

                    TxtBuscarLetras.Enabled = True
                    vLetras = ""

                    CmbDescripcion.Focus()
                    CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
                    CmbDescripcion.SelectionLength = 0
                Else
                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
                    If TxtBuscarLetras.Text.Length > 0 Then
                        TxtBuscarLetras.Text = TxtBuscarLetras.Text.Substring(0, TxtBuscarLetras.Text.Length - 1)
                    End If
                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

                    TxtBuscarLetras.Focus()
                    TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
                    vLetras = TxtBuscarLetras.Text
                    vCombo = "descripcion"
                    BuscarLetras(vCombo)
                End If
            Else
                ' 🛠️ CORRECCIÓN: Si vCombo es "descripcion" normal (sí hay registros en el desplegable)
                ' Eliminamos el MsgBox molesto para que al pulsar Enter baje directamente al combo,
                ' despliegue las opciones y seleccione el primer elemento para moverte con las flechas.
                If CmbDescripcion.Items.Count > 0 Then
                    CmbDescripcion.Focus()
                    CmbDescripcion.DroppedDown = True
                    CmbDescripcion.SelectedIndex = 0
                End If
            End If
        Else
            ' Si es cualquier otra letra (Mayúscula, minúscula, espacio, etc.), 
            ' indicamos de forma limpia que estamos en proceso de escritura ordinaria
            vIntro = "NO"
        End If
    End Sub

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        ' Sincronizamos la variable global con lo que el usuario escribe o borra arriba
        vLetras = TxtBuscarLetras.Text

        ' Evitamos ejecutar la base de datos si el cuadro superior se limpia de forma automática
        If String.IsNullOrEmpty(vLetras) Then
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.SelectedIndex = -1
            CmbDescripcion.Text = ""
            If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            Exit Sub
        End If

        ' Forzamos a evaluar el bloque de consulta SQL de descripción siempre
        BuscarLetras("descripcion")
    End Sub

    Private Function GuardarApunteEnBaseDatos() As Boolean
        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

        Dim txtConcepto As String = Trim(CmbConcepto.Text)

        ' TRADUCCIÓN INVERSA: Buscamos la clave neutra para guardarla limpia en la BD (ej: GAS_NATURAL)
        vConceptoAPU = ObtenerClaveNeutral(txtConcepto, resManager)
        If String.IsNullOrEmpty(vConceptoAPU) Then vConceptoAPU = txtConcepto

        vDescripcionAPU = Trim(CmbDescripcion.Text)
        vNotasAPU = Trim(TxtNota.Text)
        vCuentaAPU = Trim(CmbCuenta.Text)
        vDate3 = DateTimePicker1.Value.ToString("yyyy/MM/dd")

        vImporteAPU = Val(TxtImporte.Text.Replace(",", "."))
        If TxtTipoConcepto.Text = "GASTO" Then
            vImporteAPU = -Math.Abs(vImporteAPU)
        Else
            vImporteAPU = Math.Abs(vImporteAPU)
        End If

        ' Construcción SQL estricta para Access (#yyyy/MM/dd# y Str())
        vAñadirSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) VALUES "
        vAñadirSql += "(#" & vDate3 & "#, '" & vConceptoAPU & "', '" & vDescripcionAPU & "', " & Str(vImporteAPU) & ", " & CInt(vAñoEjercicio) & ", '" & vNotasAPU & "', '" & vCuentaAPU & "')"

        cmdMdb1cr.CommandText = vAñadirSql
        Try
            cmdMdb1cr.ExecuteNonQuery()

            vtipoSql = "Select apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
            vtipoGrid = "APUNTES_CONTABLES"
            LlenarGrid(vtipoSql, vtipoGrid, "1")

            Return True
        Catch ex As Exception
            MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.Message, MsgBoxStyle.Critical)
            Return False
        End Try
    End Function

    Private Sub CmbDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbDescripcion.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ' 1. Guardamos la descripción seleccionada con las flechas
            Dim textoSeleccionado As String = CmbDescripcion.Text

            ' 2. Apagamos el buscador de arriba de forma controlada
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Enabled = False

            ' 3. Aseguramos el texto en el combo y en tu variable global
            CmbDescripcion.Text = textoSeleccionado
            vDescripcion = textoSeleccionado

            ' 4. Cerramos el desplegable y mandamos el cursor al Importe
            If CmbDescripcion.DroppedDown Then CmbDescripcion.DroppedDown = False

            ' 💡 REEMPLAZA "TxtImporte" por el nombre exacto de tu caja de texto del importe
            TxtImporte.Select()
        End If
    End Sub

    Private Sub CmbDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbDescripcion.KeyPress
        ' Si es un Intro, salimos inmediatamente
        If Asc(e.KeyChar) = 13 Then
            e.Handled = True
            Exit Sub
        End If

        ' 🛠️ BLINDAJE ABSOLUTO MANTENIDO: Si vIntro es "SI", el usuario edita libremente
        If vIntro = "SI" Then
            ' Sincronizamos la variable de forma segura reteniendo mayúsculas/minúsculas nativas
            BeginInvoke(Sub() vDescripcion = CmbDescripcion.Text)
            Exit Sub
        End If

        ' 🛠️ NUEVA MEJORA EN CALIENTE: Si el usuario borra todo el texto a mano, 
        ' liberamos el combo para poder escribir una descripción nueva desde cero.
        If String.IsNullOrEmpty(CmbDescripcion.Text) Then
            vIntro = "SI"
            vDescripcion = ""
            Exit Sub
        End If

        Dim letra As Char = e.KeyChar

        ' Flujo por defecto si sólo está buscando palabras existentes
        If Char.IsLetterOrDigit(letra) OrElse letra = " "c Then
            vCombo = "descripcion"

            TxtBuscarLetras.Enabled = True
            TxtBuscarLetras.Text += letra.ToString()

            TxtBuscarLetras.Focus()
            TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
            TxtBuscarLetras.SelectionLength = 0

            vLetras = TxtBuscarLetras.Text

            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.DroppedDown = False
            CmbDescripcion.SelectedIndex = -1

            If CmbDescripcion.Items.Count > 0 Then
                CmbDescripcion.Items.Clear()
            End If

            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.Text = ""
            e.Handled = True
        End If
    End Sub

    Private Sub CmbDescripcion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDescripcion.SelectedIndexChanged
        ' OPTIMIZACIÓN CRÍTICA: Eliminamos la consulta SQL redundante a la base de datos.
        ' Cuando el usuario hace clic en un elemento de la lista, el combo ya adquiere ese texto de forma nativa.
        ' No necesitas abrir drMdb1, ni hacer un SELECT, ni volver a setear el .Text.

        If vIntro = "NO" Then
            ' Guardamos el valor seleccionado en tu variable global por si la usas en otro sitio
            vDescripcion = CmbDescripcion.Text.Trim()
        End If
    End Sub

    Private Sub CmbDescripcion_GotFocus(sender As Object, e As EventArgs) Handles CmbDescripcion.GotFocus
        ' Si venimos del combo de conceptos, nos aseguramos de dejar guardado su valor real
        If vCombo = "concepto" Then
            TxtBuscarLetras.Enabled = False
            CmbConcepto.DroppedDown = False
            CmbConcepto.Text = vConcepto
        End If

        ' Preparamos el entorno para la descripción
        ' 🛠️ AJUSTE: Solo ponemos vIntro en "NO" si no estábamos ya editando una descripción nueva ("SI")
        If vIntro <> "SI" Then vIntro = "NO"
        vCombo = "descripcion"

        ' 🛠️ CORRECCIÓN ABSOLUTA: Cambiamos Lower por Normal para liberar las Mayúsculas/Minúsculas
        TxtBuscarLetras.CharacterCasing = CharacterCasing.Normal

        ' =====================================================================
        ' 🛠️ CONTROL INTELIGENTE DE BORRADO (MANTENIDO)
        ' =====================================================================
        ' SI EL COMBO YA TIENE TEXTO (porque el concepto le ha metido la descripción por defecto),
        ' NO lo borramos. Solo seleccionamos el texto para que el usuario pueda escribir encima si quiere.
        If String.IsNullOrEmpty(CmbDescripcion.Text) Then
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
            CmbDescripcion.SelectedIndex = -1
            CmbDescripcion.Text = ""
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        Else
            ' Si ya tiene texto, colocamos el cursor al final de la palabra de forma limpia
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
        End If

        ' 🛠️ MEJORA DE NAVEGACIÓN: Si estamos en modo de alta nueva, controlamos el cursor parpadeante
        If vIntro = "SI" Then
            CmbDescripcion.SelectionStart = CmbDescripcion.Text.Length
            CmbDescripcion.SelectionLength = 0
        End If
    End Sub

    Private Sub TxtImporte_GotFocus(sender As Object, e As EventArgs) Handles TxtImporte.GotFocus
        ' Limpiamos el texto del buscador de forma segura
        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
        TxtBuscarLetras.Text = ""
        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

        ' Lo dejamos deshabilitado momentáneamente mientras se introduce el dinero
        TxtBuscarLetras.Enabled = False
        vCombo = ""
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()

        ' Llenar el Combo Concepto al cerrar
        '***********************************
        If CmbConcepto.Items.Count <> 0 Then
            CmbConcepto.Items.Clear()
        End If
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbConcepto.Items.Add(drMdb1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            'MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try

        'LlenarConcepto()
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            CmbCuenta.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptarOtro.Select()
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox(rmse.GetString("EjercicioActual"), MsgBoxStyle.Information, rmse.GetString("Fecha31Diciembre"))
            DateTimePicker1.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker1.Value = vfechaHoy
        End If
    End Sub

    Private Sub BtnAceptarSalir_Click(sender As Object, e As EventArgs) Handles BtnAceptarSalir.Click
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        Else

        End If
        If TxtImporte.Text <> "0" Then
            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    vDate1 = frmApuntesContables.DateTimePicker1.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    vDate2 = frmApuntesContables.DateTimePicker2.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                vFilaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
                If vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1 Then
                    MsgBox(resManager.GetString("MsgFila2"))
                Else
                    vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            Else
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try
                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If frmApuntesContables.BtnFechasClick = "SI" Then
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
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                        End If
                    Else
                        vtipoSql += " Or "
                        If frmApuntesContables.BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
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
            Me.Close()
        Else
            MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
            TxtImporte.Select()
        End If
    End Sub

    Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
        If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
            'Close()
        Else

        End If

        If TxtImporte.Text <> "0" Then
            If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString

                ' Sincronizamos vConcepto con el texto actual en pantalla antes de guardar
                vConcepto = Trim(CmbConcepto.Text)

                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try

                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                    vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
                End If
                If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                    vDate1 = frmApuntesContables.DateTimePicker1.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    vDate2 = frmApuntesContables.DateTimePicker2.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
                    vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                    vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")

                ' Enfoque seguro al último registro añadido
                frmApuntesContables.DgvApuntes.Refresh()
                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                If vFila >= 0 Then
                    frmApuntesContables.DgvApuntes.ClearSelection()
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            Else
                vDate3 = DateTimePicker1.Value
                vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
                vImporteAPU = TxtImporte.Text
                If TxtTipoConcepto.Text = "GASTO" Then
                    vImporteAPU = "-" & vImporteAPU.ToString
                End If
                vNotasAPU = TxtNota.Text
                vCuentaAPU = CmbCuenta.Text.ToString

                ' Sincronizamos vConcepto con el texto actual en pantalla antes de guardar
                vConcepto = Trim(CmbConcepto.Text)

                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
                End Try

                Dim i As Integer
                vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
                If frmApuntesContables.BtnFechasClick = "SI" Then
                    vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                Else
                    vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                End If

                For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
                    vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString()
                    If i = 0 Then
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                        End If
                    Else
                        vtipoSql += " Or "
                        If frmApuntesContables.BtnFechasClick = "SI" Then
                            vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
                        Else
                            vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                        End If
                        vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                        If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
                            vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
                        End If
                        If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
                            vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
                            vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
                            vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
                            vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
                        End If
                    End If
                Next

                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")

                ' Enfoque seguro al último registro añadido en modo ListBox
                frmApuntesContables.DgvApuntes.Refresh()
                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                If vFila >= 0 Then
                    frmApuntesContables.DgvApuntes.ClearSelection()
                    frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                    frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
                End If
            End If

            ' =====================================================================
            ' 🛠️ REINICIO ASISTIDO DE PASOS PARA INTRODUCIR OTRO APUNTE
            ' =====================================================================
            ' 1. Apagamos el TextChanged superior para evitar llamadas falsas a la BD al limpiar
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

            ' 2. Reestablecemos el estado inicial del buscador de descripciones
            TxtBuscarLetras.Enabled = True
            vLetras = ""
            vIntro = "NO"

            ' 3. Vaciamos las cajas de texto numéricas y complementarias
            TxtImporte.Text = "0"
            TxtNota.Text = ""
            TxtDescripcion.Text = ""
            CmbConcepto.Text = ""
            CmbDescripcion.Text = ""

            ' 4. Regresamos el cursor a la Fecha para arrancar de nuevo el flujo de introducción
            DateTimePicker1.Focus()
        Else
            MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
            TxtImporte.Select()
            TxtImporte.SelectAll()
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

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs) Handles BtnCuenta.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub BtnAyuda_Click(sender As Object, e As EventArgs) Handles BtnAyuda.Click
        ' 🛠️ CONTROL DE VENTANA FLOTANTE INDEPENDIENTE
        ' Comprobamos si la ventana de ayuda ya está abierta en pantalla
        Dim frmExistente As AyudaApuntes = Application.OpenForms.OfType(Of AyudaApuntes)().FirstOrDefault()

        If frmExistente IsNot Nothing Then
            ' Si ya estaba abierta, la cerramos (esto disparará automáticamente el evento FormClosed)
            frmExistente.Close()
        Else
            ' --- DESPLAZAR EL FORMULARIO ACTUAL A LA IZQUIERDA ---
            Dim pixelesDesplazamiento As Integer = 150
            Me.Left -= pixelesDesplazamiento

            ' Creamos la instancia de la ventana de ayuda
            Dim frmAyuda As New AyudaApuntes()

            ' --- DETECTAR EL CIERRE (BOTÓN O CRUZ X) ---
            ' Usamos AddHandler para ejecutar código cuando frmAyuda se cierre por cualquier motivo
            AddHandler frmAyuda.FormClosed, Sub(s, ev)
                                                ' Cuando la ayuda se cierra, devolvemos el formulario principal a la derecha
                                                Me.Left += pixelesDesplazamiento
                                            End Sub

            ' CÁLCULO DE POSICIÓN: Se calcula usando la NUEVA posición del formulario actual
            Dim x As Integer = Me.Location.X + Me.Width
            Dim y As Integer = Me.Location.Y

            frmAyuda.Location = New Point(x, y)

            ' La mostramos en modo "Show"
            frmAyuda.Show(Me)
        End If
    End Sub

    Private Sub CmbCuenta_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbCuenta.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            TxtNota.Select()
        End If
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub BtnDescripcion_Click(sender As Object, e As EventArgs) Handles BtnDescripcion.Click
        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()
    End Sub

    Private Sub DateTimePicker1_LostFocus(sender As Object, e As EventArgs) Handles DateTimePicker1.LostFocus
        BtnCalculadora.TabIndex = 0
        BtnConcepto.TabIndex = 0
        BtnDescripcion.TabIndex = 0
        BtnHoy.TabIndex = 0
        BtnCuenta.TabIndex = 0
    End Sub

    Private Sub CmbConcepto_Click(sender As Object, e As EventArgs) Handles CmbConcepto.Click
        CmbConcepto.DroppedDown = True
    End Sub

    Private Sub CmbDescripcion_Click(sender As Object, e As EventArgs) Handles CmbDescripcion.Click
        CmbDescripcion.DroppedDown = True
    End Sub

    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            CmbConcepto.Select()
            ' Opcional: Ejecutar una búsqueda o guardar valor
            'BtnHoy.PerformClick()
        End If
    End Sub

    Public Function LlenarDescripcion() As String
        ' 1. BLINDAJE: Vaciamos el combo de forma segura desactivando eventos para evitar errores de índice
        RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        CmbDescripcion.SelectedIndex = -1
        CmbDescripcion.Items.Clear()
        AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

        ' 2. OPTIMIZACIÓN CRÍTICA: Usamos SELECT DISTINCT para que la base de datos filtre los duplicados al vuelo
        cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE DescripcionAPU <> 'Saldo Inicial' ORDER BY DescripcionAPU ASC"

        Try
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            drMdb1 = cmdMdb1cr.ExecuteReader()

            If drMdb1.HasRows Then
                While drMdb1.Read()
                    Dim desc As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
                    If Not String.IsNullOrEmpty(desc) Then
                        CmbDescripcion.Items.Add(desc)
                    End If
                End While
            End If
            drMdb1.Close()
        Catch ex As Exception
            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            MsgBox(rmse.GetString("ErrorLlenarDesplegable") & rmse.GetString("Label3.Text") & " " & ex.Message, vbExclamation, rmse.GetString("$this.Text"))
        End Try
        Return ""
    End Function

    Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True

            ' 1. Guardamos el texto que ha seleccionado el usuario con las flechas
            Dim textoSeleccionado As String = CmbConcepto.Text
            vConcepto = textoSeleccionado

            ' 2. Apagamos el buscador de arriba para que al borrarlo no active la base de datos
            RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
            TxtBuscarLetras.Text = ""
            AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged

            ' 🛠️ AJUSTE AQUÍ: Lo dejamos habilitado y configuramos vCombo en "descripcion"
            ' para que cuando el usuario empiece a buscar descripciones, el sistema sepa dónde está.
            TxtBuscarLetras.Enabled = True
            vCombo = "descripcion"

            ' 3. Restauramos el texto en el combo por si el borrado de arriba hizo amago de limpiarlo
            CmbConcepto.Text = textoSeleccionado

            ' =====================================================================
            ' 🛠️ MANTENIDO: CARGAR LA DESCRIPCIÓN POR DEFECTO DEL CONCEPTO SELECCIONADO
            ' =====================================================================
            If Not String.IsNullOrEmpty(vConcepto) Then
                Try
                    ' Cerramos el lector si estuviera abierto por seguridad
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

                    ' Consultamos la tabla de conceptos para traer su tipo y descripción asociada
                    cmdMdb1cr.CommandText = "SELECT * FROM conceptos WHERE CodigoCON = '" & vConcepto.Replace("'", "''") & "' ORDER BY CodigoCON ASC"
                    drMdb1 = cmdMdb1cr.ExecuteReader()

                    If drMdb1.Read() Then
                        ' Rellenamos el tipo de concepto
                        TxtTipoConcepto.Text = Convert.ToString(drMdb1.GetValue(2))

                        ' Rellenamos el combo de descripciones con la descripción por defecto de la base de datos
                        CmbDescripcion.Text = Convert.ToString(drMdb1.GetValue(1))
                        vDescripcion = CmbDescripcion.Text
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    ' Manejo silencioso o registro de error para no romper el flujo
                    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
                End Try
            End If
            ' =====================================================================

            ' 4. Cerramos la lista y saltamos limpiamente al combo de descripción o al buscador
            If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False

            ' Si el concepto ya traía una descripción automática por defecto, saltas al importe.
            ' Si venía vacío, puedes saltar directamente a CmbDescripcion para empezar a buscar letras.
            CmbDescripcion.Select()
        End If
    End Sub

    Public Function BuscarLetras(combo As String) As String
        ' Cerramos cualquier lector abierto preventivamente
        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then
            drMdb1.Close()
        End If

        ' =====================================================================
        ' MODO: DESCRIPCIÓN
        ' =====================================================================
        If combo = "descripcion" Then
            ' 1. Desconectamos el evento por seguridad
            RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

            ' 🛠️ CONTROL DE TEXTO CORTO: Si borras y queda en 2 letras o menos, limpiamos la lista
            If vLetras.Length <= 2 Then
                CmbDescripcion.DroppedDown = False
                CmbDescripcion.SelectedIndex = -1
                If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()
                vCombo = "descripcion"

                ' Volvemos a conectar el evento y salimos de la función de forma limpia
                AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
                Return ""
            End If

            Dim letrasLimpias As String = vLetras.Replace("'", "''")
            cmdMdb1cr.CommandText = "SELECT DISTINCT DescripcionAPU FROM apuntes WHERE UCase(DescripcionAPU) LIKE '%" & letrasLimpias.ToUpper() & "%' AND DescripcionAPU <> 'Saldo Inicial'"

            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    ' Si hay registros, limpiamos y llenamos las sugerencias
                    CmbDescripcion.SelectedIndex = -1
                    If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

                    While drMdb1.Read()
                        Dim desc As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
                        If Not String.IsNullOrEmpty(desc) Then CmbDescripcion.Items.Add(desc)
                    End While

                    ' Desplegamos la persiana gráfica con los resultados
                    CmbDescripcion.DroppedDown = True
                    vCombo = "descripcion"
                Else
                    ' =====================================================================
                    ' 🛠️ CONTROL DE BÚSQUEDA VACÍA SANEADO
                    ' =====================================================================
                    drMdb1.Close()

                    ' Cerramos la persiana y vaciamos la lista vieja para que no muestre datos erróneos
                    CmbDescripcion.DroppedDown = False
                    CmbDescripcion.SelectedIndex = -1
                    If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()

                    ' Guardamos el estado especial de que esta descripción es NUEVA
                    vCombo = "descripcion_vacia"
                End If

                ' Doble comprobación de seguridad para asegurar el cierre del lector
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

            Catch ex As Exception
                MsgBox(rmse.GetString("ErrorBuscarLetrasDescripcion") & ": " & ex.Message)
                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
            End Try

            ' Volvemos a conectar el evento limpiamente al finalizar
            AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
        End If

        Return ""
    End Function

End Class

'======================================================================================================

'Public Function LlenarConcepto() As String
'    ' 1. BLINDAJE: Vaciamos el combo de forma segura desactivando eventos
'    RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
'    CmbConcepto.SelectedIndex = -1
'    CmbConcepto.Items.Clear()
'    AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

'    ' Forzamos solo la columna que necesitamos
'    cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON ASC"

'    Try
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        drMdb1 = cmdMdb1cr.ExecuteReader()

'        If drMdb1.HasRows Then
'            While drMdb1.Read()
'                Dim valor As String = Convert.ToString(drMdb1.GetValue(0))
'                If valor <> "TRASPASO" Then
'                    CmbConcepto.Items.Add(valor)
'                End If
'            End While

'            ' 2. Asignación del concepto inicial por defecto
'            If CmbConcepto.Items.Count > 0 Then
'                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
'                    If frmApuntesContables.ListBox1.SelectedItems.Count <> 0 Then
'                        CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
'                    Else
'                        Dim index As Integer = frmApuntesContables.CmbConcepto.SelectedIndex - 1
'                        If index >= 0 AndAlso index < CmbConcepto.Items.Count Then
'                            CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(index))
'                        Else
'                            CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
'                        End If
'                    End If
'                Else
'                    CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
'                End If
'            End If
'        End If
'        drMdb1.Close()
'    Catch ex As Exception
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        MsgBox(rmse.GetString("ErrorLlenarDesplegable") & rmse.GetString("Label2.Text") & " " & ex.Message, vbExclamation, rmse.GetString("$this.Text"))
'    End Try
'    Return ""
'End Function


''esta función llena el combo de descripción con las claves neutras traducidas a texto visible, tomando como referencia la tabla apuntes para obtener las claves únicas de descripción, y luego traducirlas usando el ResourceManager. Se ha optimizado para evitar duplicados y se ha blindado con manejo de excepciones y cierre adecuado del DataReader.
'Public Function LlenarDescripcion() As String
'    RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
'    CmbDescripcion.SelectedIndex = -1 : CmbDescripcion.Items.Clear()
'    AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

'    cmdMdb1cr.CommandText = "SELECT DISTINCT ConceptoAPU FROM apuntes WHERE ConceptoAPU <> 'SALDO' ORDER BY ConceptoAPU ASC"
'    Try
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        drMdb1 = cmdMdb1cr.ExecuteReader()
'        If drMdb1.HasRows Then
'            While drMdb1.Read()
'                Dim claveNeutro As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
'                If Not String.IsNullOrEmpty(claveNeutro) Then
'                    Dim textoTraducido As String = resManager.GetString("Desc_" & claveNeutro)
'                    If Not String.IsNullOrEmpty(textoTraducido) Then CmbDescripcion.Items.Add(textoTraducido.Trim())
'                End If
'            End While
'        End If
'        drMdb1.Close()
'    Catch ex As Exception
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        MsgBox("Error al cargar Descripciones: " & ex.Message)
'    End Try
'    Return ""
'End Function


'' con traducción inversa y carga de descripción traducida
'Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
'    If e.KeyCode = Keys.Enter Then
'        e.SuppressKeyPress = True

'        Dim textoSeleccionado As String = CmbConcepto.Text

'        RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'        TxtBuscarLetras.Text = ""
'        AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'        TxtBuscarLetras.Enabled = False

'        CmbConcepto.Text = textoSeleccionado

'        ' TRADUCCIÓN INVERSA: Buscamos la clave neutra original
'        vConcepto = ObtenerClaveNeutral(textoSeleccionado, resManager)
'        If String.IsNullOrEmpty(vConcepto) Then vConcepto = textoSeleccionado

'        ' Cargar descripción por defecto traducida
'        If Not String.IsNullOrEmpty(vConcepto) Then
'            Try
'                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'                cmdMdb1cr.CommandText = "SELECT * FROM conceptos WHERE CodigoCON = '" & vConcepto.Replace("'", "''") & "' ORDER BY CodigoCON ASC"
'                drMdb1 = cmdMdb1cr.ExecuteReader()

'                If drMdb1.Read() Then
'                    TxtTipoConcepto.Text = Convert.ToString(drMdb1.GetValue(2))
'                    Dim keyDesc As String = "Desc_" & vConcepto
'                    Dim descTraducida As String = resManager.GetString(keyDesc)

'                    If Not String.IsNullOrEmpty(descTraducida) Then
'                        CmbDescripcion.Text = descTraducida
'                    Else
'                        CmbDescripcion.Text = Convert.ToString(drMdb1.GetValue(1))
'                    End If
'                    vDescripcion = CmbDescripcion.Text
'                    MsgBox("CmbConcepto_KeyDown" & vDescripcion)
'                End If
'                drMdb1.Close()
'            Catch ex As Exception
'                If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'            End Try
'        End If

'        If CmbConcepto.DroppedDown Then CmbConcepto.DroppedDown = False
'        CmbDescripcion.Select()
'    End If
'End Sub

'' Versión optimizada y comentada de tu función BuscarLetras, con traducción inversa y control de descripción nueva
'Public Function BuscarLetras(combo As String) As String
'    If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()

'    If combo = "concepto" Then
'        RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
'        CmbConcepto.DroppedDown = False : CmbConcepto.SelectedIndex = -1
'        If CmbConcepto.Items.Count > 0 Then CmbConcepto.Items.Clear()
'        AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

'        Dim letrasLimpias As String = vLetras.Replace("'", "''")
'        cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos WHERE CodigoCON LIKE '%" & letrasLimpias & "%' ORDER BY CodigoCON ASC"
'        Try
'            drMdb1 = cmdMdb1cr.ExecuteReader()
'            If drMdb1.HasRows Then
'                While drMdb1.Read()
'                    Dim valor As String = Convert.ToString(drMdb1.GetValue(0))
'                    If valor <> "TRASPASO" Then
'                        Dim trans As String = resManager.GetString(valor)
'                        CmbConcepto.Items.Add(If(String.IsNullOrEmpty(trans), valor, trans).ToUpper())
'                    End If
'                End While
'                CmbConcepto.DroppedDown = True
'            Else
'                CmbConcepto.DroppedDown = False
'            End If
'            drMdb1.Close()
'        Catch ex As Exception
'            MsgBox("Error en búsqueda Concepto: " & ex.Message)
'        End Try
'    End If

'    If combo = "descripcion" AndAlso vLetras.Length > 2 Then
'        RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
'        CmbDescripcion.DroppedDown = False : CmbDescripcion.SelectedIndex = -1
'        If CmbDescripcion.Items.Count > 0 Then CmbDescripcion.Items.Clear()
'        AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

'        Dim letrasLimpias As String = vLetras.Replace("'", "''")
'        cmdMdb1cr.CommandText = "SELECT DISTINCT ConceptoAPU FROM apuntes WHERE ConceptoAPU LIKE '%" & letrasLimpias & "%' AND ConceptoAPU <> 'SALDO'"
'        Try
'            drMdb1 = cmdMdb1cr.ExecuteReader()
'            Dim huboFilas As Boolean = False

'            While drMdb1.Read()
'                Dim valor As String = Convert.ToString(drMdb1.GetValue(0))
'                Dim trans As String = resManager.GetString("Desc_" & valor)
'                If Not String.IsNullOrEmpty(trans) Then
'                    CmbDescripcion.Items.Add(trans.Trim())
'                    huboFilas = True
'                End If
'            End While
'            drMdb1.Close()

'            If huboFilas Then
'                CmbDescripcion.DroppedDown = True
'                vCombo = "descripcion"
'            Else
'                CmbDescripcion.DroppedDown = False
'                Dim respuesta As MsgBoxResult = MsgBox("No existen Descripciones con: -" & vLetras.ToUpper() & "-" & vbCrLf & "¿Añadimos la descripción?", vbQuestion + vbYesNo + vbDefaultButton1, "Introducir Apunte")

'                If respuesta = vbYes Then
'                    vIntro = "SI"
'                    CmbDescripcion.Text = vLetras : vDescripcion = vLetras
'                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'                    TxtBuscarLetras.Text = ""
'                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'                    TxtBuscarLetras.Enabled = False
'                    TxtImporte.Focus() : TxtImporte.SelectAll()
'                Else
'                    RemoveHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'                    If TxtBuscarLetras.Text.Length > 0 Then TxtBuscarLetras.Text = TxtBuscarLetras.Text.Substring(0, TxtBuscarLetras.Text.Length - 1)
'                    AddHandler TxtBuscarLetras.TextChanged, AddressOf TxtBuscarLetras_TextChanged
'                    TxtBuscarLetras.Focus() : TxtBuscarLetras.SelectionStart = TxtBuscarLetras.Text.Length
'                    vLetras = TxtBuscarLetras.Text
'                End If
'            End If
'        Catch ex As Exception
'            MsgBox("Error en búsqueda Descripción: " & ex.Message)
'            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        End Try
'    End If
'    Return ""
'End Function

'' con traducción de claves neutras a texto visible
'Public Function LlenarConcepto() As String
'    RemoveHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged
'    CmbConcepto.SelectedIndex = -1 : CmbConcepto.Items.Clear()
'    AddHandler CmbConcepto.SelectedIndexChanged, AddressOf CmbConcepto_SelectedIndexChanged

'    cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON ASC"
'    Try
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        drMdb1 = cmdMdb1cr.ExecuteReader()
'        If drMdb1.HasRows Then
'            While drMdb1.Read()
'                Dim claveNeutro As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
'                If claveNeutro <> "TRASPASO" AndAlso Not String.IsNullOrEmpty(claveNeutro) Then
'                    Dim textoTraducido As String = resManager.GetString(claveNeutro)
'                    If String.IsNullOrEmpty(textoTraducido) Then textoTraducido = claveNeutro
'                    CmbConcepto.Items.Add(textoTraducido.ToUpper())
'                End If
'            End While
'            'If CmbConcepto.Items.Count > 0 Then CmbConcepto.Text = Convert.ToString(CmbConcepto.Items(0))
'            If CmbConcepto.Items.Count > 0 Then TxtDescripcion.Text = Convert.ToString(CmbConcepto.Items(0))
'        End If
'        'drMdb1.Close()
'    Catch ex As Exception
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        MsgBox("Error al cargar Conceptos: " & ex.Message)
'    End Try
'    drMdb1.Close()
'    Return ""
'End Function

''con traducción de claves neutras a texto visible y doble intento de carga (primero desde apuntes, si no hay nada, desde conceptos)
'Public Function LlenarDescripcion() As String
'    ' 1. Desvincular eventos y limpiar el ComboBox
'    RemoveHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged
'    CmbDescripcion.SelectedIndex = -1 : CmbDescripcion.Items.Clear()
'    AddHandler CmbDescripcion.SelectedIndexChanged, AddressOf CmbDescripcion_SelectedIndexChanged

'    Try
'        Dim cargoDesdeApuntes As Boolean = False

'        ' INTERVENTOR 1: Intentar cargar desde la tabla 'apuntes' (Flujo normal)
'        cmdMdb1cr.CommandText = "SELECT DISTINCT ConceptoAPU FROM apuntes WHERE ConceptoAPU <> 'SALDO' ORDER BY ConceptoAPU ASC"

'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        drMdb1 = cmdMdb1cr.ExecuteReader()

'        If drMdb1.HasRows Then
'            cargoDesdeApuntes = True
'            While drMdb1.Read()
'                Dim claveNeutro As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
'                If Not String.IsNullOrEmpty(claveNeutro) Then
'                    Dim textoTraducido As String = resManager.GetString("Desc_" & claveNeutro)
'                    If Not String.IsNullOrEmpty(textoTraducido) Then
'                        CmbDescripcion.Items.Add(textoTraducido.Trim())
'                    End If
'                End If
'            End While
'        End If
'        drMdb1.Close()

'        ' INTERVENTOR 2: Si 'apuntes' estaba vacía, cargamos basándonos en la tabla 'conceptos'
'        If Not cargoDesdeApuntes Then
'            cmdMdb1cr.CommandText = "SELECT CodigoCON FROM conceptos ORDER BY CodigoCON ASC"

'            If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'            drMdb1 = cmdMdb1cr.ExecuteReader()

'            If drMdb1.HasRows Then
'                While drMdb1.Read()
'                    Dim claveNeutro As String = Convert.ToString(drMdb1.GetValue(0)).Trim()
'                    ' Excluimos TRASPASO igual que haces en LlenarConcepto si es necesario
'                    If claveNeutro <> "TRASPASO" AndAlso Not String.IsNullOrEmpty(claveNeutro) Then
'                        ' Buscamos la descripción asociada al concepto base
'                        Dim textoTraducido As String = resManager.GetString("Desc_" & claveNeutro)

'                        ' Si existe la traducción "Desc_CONCEPTO", la añadimos al combo
'                        If Not String.IsNullOrEmpty(textoTraducido) Then
'                            CmbDescripcion.Items.Add(textoTraducido.Trim())
'                        End If
'                    End If
'                End While
'            End If
'            drMdb1.Close()
'        End If

'        ' Seleccionar el primer elemento por defecto si se llenó algo
'        If CmbDescripcion.Items.Count > 0 Then
'            CmbDescripcion.Text = Convert.ToString(CmbDescripcion.Items(0))
'        End If

'    Catch ex As Exception
'        If drMdb1 IsNot Nothing AndAlso Not drMdb1.IsClosed Then drMdb1.Close()
'        MsgBox("Error al cargar Descripciones: " & ex.Message)
'    End Try

'    Return ""
'End Function

'Private Sub BtnAceptarOtro_Click(sender As Object, e As EventArgs) Handles BtnAceptarOtro.Click
'    If frmApuntesContables.DgvApuntes.RowCount >= 25 And My.Settings.Autorizar = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
'        'MsgBox("Software No Activado, Máximo 25 Apuntes", MsgBoxStyle.Critical, "Falta Activación")
'        'Close()
'    Else

'    End If
'    If TxtImporte.Text <> "0" Then
'        If frmApuntesContables.ListBox1.SelectedItems.Count = 0 Then
'            vDate3 = DateTimePicker1.Value
'            vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
'            vImporteAPU = TxtImporte.Text
'            If TxtTipoConcepto.Text = "GASTO" Then
'                vImporteAPU = "-" & vImporteAPU.ToString
'            End If
'            vNotasAPU = TxtNota.Text
'            vCuentaAPU = CmbCuenta.Text.ToString
'            vAñadirSql = "INSERT INTO apuntes "
'            vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
'            vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
'            cmdMdb1cr.CommandText = vAñadirSql
'            Try
'                cmdMdb1cr.ExecuteNonQuery()
'                'MsgBox("Registro, Grabado Correctamente")
'            Catch ex As Exception
'                MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
'            End Try
'            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
'            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
'            If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
'                vtipoSql += " And apuntes.CuentaAPU = '" & frmApuntesContables.CmbCuenta.Text & "' "
'            End If
'            If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
'                vtipoSql += " And apuntes.ConceptoAPU = '" & frmApuntesContables.CmbConcepto.Text & "' "
'            End If
'            If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
'                vDate1 = frmApuntesContables.DateTimePicker1.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
'                vDate2 = frmApuntesContables.DateTimePicker2.Value.ToString("yyyy/MM/dd", CultureInfo.InvariantCulture)
'                vtipoSql += " And apuntes.FechaAPU >= ?"
'                vtipoSql += " And apuntes.FechaAPU <= ?"
'            End If
'            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
'            vtipoGrid = "APUNTES_CONTABLES"
'            LlenarGrid(vtipoSql, vtipoGrid, "1")
'            vFila = frmApuntesContables.DgvApuntes.RowCount - 1
'            frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
'            frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
'        Else
'            vDate3 = DateTimePicker1.Value
'            vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)
'            vImporteAPU = TxtImporte.Text
'            If TxtTipoConcepto.Text = "GASTO" Then
'                vImporteAPU = "-" & vImporteAPU.ToString
'            End If
'            vNotasAPU = TxtNota.Text
'            vCuentaAPU = CmbCuenta.Text.ToString
'            vAñadirSql = "INSERT INTO apuntes "
'            vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
'            vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
'            cmdMdb1cr.CommandText = vAñadirSql
'            Try
'                cmdMdb1cr.ExecuteNonQuery()
'                'MsgBox("Registro, Grabado Correctamente")
'            Catch ex As Exception
'                MsgBox(rmse.GetString("ErrorGrabarRegistro") & ": " & ex.ToString, vbExclamation, rmse.GetString("$this.Text"))
'            End Try
'            Dim i As Integer
'            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
'            If frmApuntesContables.BtnFechasClick = "SI" Then
'                vtipoSql += " WHERE apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
'            Else
'                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
'            End If
'            For i = 0 To frmApuntesContables.ListBox1.SelectedItems.Count - 1
'                vConcepto = frmApuntesContables.ListBox1.SelectedItems(i).ToString
'                If i = 0 Then
'                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
'                    If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
'                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
'                    End If
'                    If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
'                        vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
'                        vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
'                        vtipoSql += " And apuntes.FechaAPU >= #" & vDate1 & "#"
'                        vtipoSql += " And apuntes.FechaAPU <= #" & vDate2 & "#"
'                    End If
'                Else
'                    vtipoSql += " Or "
'                    If frmApuntesContables.BtnFechasClick = "SI" Then
'                        vtipoSql += "apuntes.ConceptoAPU <> 'SALDO' And apuntes.EjercicioAPU <> 0 "
'                    Else
'                        vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
'                    End If
'                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
'                    If frmApuntesContables.BtnFiltroCuenta.Enabled = False Then
'                        vtipoSql += " And apuntes.CuentaAPU = '" & CmbCuenta.Text & "' "
'                    End If
'                    If frmApuntesContables.BtnFiltroFecha.Enabled = False Then
'                        vDate1 = frmApuntesContables.DateTimePicker1.Value.Date
'                        vDate2 = frmApuntesContables.DateTimePicker2.Value.Date
'                        vtipoSql += " And apuntes.FechaAPU >= ?"
'                        vtipoSql += " And apuntes.FechaAPU <= ?"
'                    End If
'                End If
'            Next
'            vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
'            vtipoGrid = "APUNTES_CONTABLES"
'            LlenarGrid(vtipoSql, vtipoGrid, "1")
'            If frmApuntesContables.DgvApuntes.RowCount - 1 >= 0 Then
'                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
'                frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
'                frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
'            End If
'        End If
'        TxtImporte.Text = 0
'        TxtNota.Text = ""
'        DateTimePicker1.Select()
'    Else
'        MsgBox(rmse.GetString("NoCantidadImporte"), vbExclamation, rmse.GetString("$this.Text"))
'        TxtImporte.Select()
'        TxtImporte.SelectAll()
'    End If
'End Sub
