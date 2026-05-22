Imports System.Diagnostics
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class IntroApuntes

    Public vConcepto, vtipoSql, vtipoGrid, vAñadirSql As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU, strText, vIntro, vLetras, vCombo, vDescripcion As String
    Public vImporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

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

        Dim TL(12) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, "Ir a Hoy")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptarOtro, "Aceptar y Otro")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptarSalir, "Aceptar y Salir")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, "Seleccionar el Concepto o Escribir para Buscar (2 mínimo)")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, "Seleccionar la Cuenta a la que se refiere la transacción")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, "Seleccionar la Descripción o Escribir para Buscar (3 mínimo)")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, "Importe del Asiento")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, "Activar la Calculadora")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, "Editar Consulta Conceptos Contables y al salir Rellena Conceptos")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuenta, "Añade, Edita, Borra o Consulta Cuentas Bancarias")
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnDescripcion, "Rellena Descripciones de Apuntes guardados")
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.TxtBuscarLetras, "Texto a Buscar")

        ' Llenar el Combo Concepto
        '*************************
        LlenarConcepto()

        ' Llenar el Combo Descripción
        '****************************
        LlenarDescripcion()

        ' Llenar el Combo Cuenta
        '***********************
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
            MsgBox("Error al llenar el Combo Cuenta")
            MsgBox(ex.ToString)
        End Try
        TxtImporte.Text = 0
    End Sub

    Private Sub CmbConcepto_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbConcepto.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            CmbDescripcion.Select()
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
        TxtBuscarLetras.Enabled = True
        TxtBuscarLetras.Text = e.KeyChar
        TxtBuscarLetras.Select()
        TxtBuscarLetras.Select(TxtBuscarLetras.Text.Length, 1)
        vLetras = TxtBuscarLetras.Text
        If CmbConcepto.Items.Count >= 1 Then
            CmbConcepto.Items.Clear()
        End If
        CmbConcepto.Text = ""
        CmbDescripcion.Text = ""
        vCombo = "concepto"
    End Sub

    Private Sub CmbDescripcion_KeyDown(sender As Object, e As KeyEventArgs) Handles CmbDescripcion.KeyDown
        ' Verificamos si la tecla presionada es Enter
        If e.KeyCode = Keys.Enter Then
            ' 1. Evitar el sonido de "beep" al pulsar Enter
            e.SuppressKeyPress = True
            TxtImporte.Select()
        End If
    End Sub

    Private Sub CmbDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbDescripcion.KeyPress
        If vIntro = "NO" Then
            TxtBuscarLetras.Enabled = True
            TxtBuscarLetras.Text = e.KeyChar
            TxtBuscarLetras.Select()
            TxtBuscarLetras.Select(TxtBuscarLetras.Text.Length, 1)
            vLetras = TxtBuscarLetras.Text
            If CmbDescripcion.Items.Count >= 1 Then
                CmbDescripcion.Items.Clear()
            End If
            CmbDescripcion.Text = ""
            vCombo = "descripcion"
        End If
    End Sub

    Private Sub TxtBuscarLetras_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtBuscarLetras.KeyPress
        If Asc(e.KeyChar) = 13 Then
            vIntro = "SI"
            If vCombo = "concepto" Then
                MsgBox("Pulsar Tabulador para ir a Concepto", MsgBoxStyle.Exclamation, "Tecla a Pulsar")
            Else
                MsgBox("Pulsar Tabulador para ir a Descripción", MsgBoxStyle.Exclamation, "Tecla a Pulsar")
            End If
            If vCombo = "concepto" Then
                If CmbConcepto.Items.Count <> 0 Then
                    CmbConcepto.DroppedDown = True
                    CmbConcepto.SelectedIndex = 0
                End If
            End If
            If vCombo = "descripcion" Then
                If CmbDescripcion.Items.Count <> 0 Then
                    CmbDescripcion.DroppedDown = True
                    CmbDescripcion.SelectedIndex = 0
                End If
            End If
        Else
            vIntro = "NO"
            If vCombo = "concepto" Then
                e.KeyChar = Char.ToUpper(e.KeyChar)
                vLetras = TxtBuscarLetras.Text
                BuscarLetras(vCombo)
                If CmbConcepto.Items.Count <> 0 Then
                    CmbConcepto.DroppedDown = True
                    CmbConcepto.SelectedIndex = 0
                End If
            End If
            If vCombo = "descripcion" Then
                vLetras = TxtBuscarLetras.Text
                BuscarLetras(vCombo)
                If CmbDescripcion.Items.Count <> 0 Then
                    CmbDescripcion.DroppedDown = True
                    CmbDescripcion.SelectedIndex = 0
                End If
            End If
        End If
    End Sub

    Private Sub TxtBuscarLetras_TextChanged(sender As Object, e As EventArgs) Handles TxtBuscarLetras.TextChanged
        If vCombo = "concepto" Then
            CmbConcepto.Text = ""
            CmbConcepto.TabIndex = 4
            CmbDescripcion.TabIndex = 5
            If CmbConcepto.Items.Count <> 0 Then
                CmbConcepto.DroppedDown = True
                CmbConcepto.SelectedIndex = 0
            End If
        End If
        If vCombo = "descripcion" Then
            CmbDescripcion.Text = ""
            CmbConcepto.TabIndex = 2
            CmbDescripcion.TabIndex = 4
            If CmbDescripcion.Items.Count <> 0 Then
                CmbDescripcion.DroppedDown = True
                CmbDescripcion.SelectedIndex = 0
            End If
        End If
        vLetras = TxtBuscarLetras.Text
        BuscarLetras(vCombo)
    End Sub

    Function BuscarLetras(combo)
        drMdb1.Close()
        If combo = "concepto" Then
            If CmbConcepto.Items.Count <> 0 Then
                CmbConcepto.Items.Clear()
            End If
            cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON like '%" & vLetras & "%' ORDER BY conceptos.CodigoCON ASC"
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        If drMdb1.GetValue(0) <> "TRASPASO" Then
                            CmbConcepto.Items.Add(drMdb1.GetValue(0))
                        End If
                    End While
                    CmbConcepto.DroppedDown = True
                Else
                    MsgBox("No existen Conceptos con: " & vLetras)
                    TxtBuscarLetras.Text = Mid(TxtBuscarLetras.Text, 1, Len(TxtBuscarLetras.Text) - 1)
                    TxtBuscarLetras.Select(TxtBuscarLetras.Text.Length, 1)
                End If
                drMdb1.Close()
            Catch ex As Exception
                MsgBox("Error al Buscar Letras en Concepto")
                MsgBox(ex.ToString)
            End Try
        End If
        If combo = "descripcion" And Len(vLetras) > 2 Then
            If CmbDescripcion.Items.Count <> 0 Then
                CmbDescripcion.Items.Clear()
            End If
            cmdMdb1cr.CommandText = "SELECT * FROM apuntes Where apuntes.DescripcionAPU like '%" & vLetras & "%' "
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    primero = 1
                    While drMdb1.Read()
                        If Trim(drMdb1.GetValue(3)) <> "Saldo Inicial" Then
                            If primero = 1 Then
                                CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                                primero = 2
                            Else
                                nuevo = 0
                                For i = 0 To CmbDescripcion.Items.Count - 1
                                    If Trim(drMdb1.GetValue(3)) = Trim(CmbDescripcion.Items(i)) Then
                                        nuevo = 0
                                        Exit For
                                    Else
                                        nuevo = 1
                                    End If
                                Next
                                If nuevo = 1 Then
                                    CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                                    nuevo = 0
                                End If
                            End If
                        End If
                    End While
                    CmbDescripcion.DroppedDown = True
                Else
                    respuesta = MsgBox("No existen Descripciones con: -" & vLetras.ToUpper & "-" & vbCrLf & "¿Añadimos la descripción?", vbQuestion + vbYesNo + vbDefaultButton1, "Introducir Apunte")
                    If respuesta = vbYes Then
                        'No tiene que hacer nada, se añade la descripción escrita en el desplegable
                        vIntro = "SI"
                        CmbDescripcion.Text = vLetras
                        CmbDescripcion.Select()
                        CmbDescripcion.Select(CmbDescripcion.Text.Length, 1)
                        TxtBuscarLetras.Text = ""
                        TxtBuscarLetras.Enabled = False
                    Else
                        TxtBuscarLetras.Text = Mid(TxtBuscarLetras.Text, 1, Len(TxtBuscarLetras.Text) - 1)
                        TxtBuscarLetras.Select(TxtBuscarLetras.Text.Length, 1)
                    End If
                End If
                drMdb1.Close()
            Catch ex As Exception
                MsgBox("Error al Buscar Letras en Descripción")
                MsgBox(ex.ToString)
            End Try
        End If
        Return ""
    End Function

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        If vIntro = "NO" Then
            vConcepto = CmbConcepto.Text.ToString
            drMdb1.Close()
            cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' ORDER BY conceptos.CodigoCON ASC"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            drMdb1.Read()
            If drMdb1.HasRows Then
                TxtTipoConcepto.Text = drMdb1.GetValue(2)
                CmbDescripcion.Text = drMdb1.GetValue(1)
                drMdb1.Close()
            End If
        End If
    End Sub

    Private Sub CmbDescripcion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbDescripcion.SelectedIndexChanged
        ' Se busca Descripción según lo seleccionado
        '*******************************************
        If vIntro = "NO" Then
            vDescripcion = CmbDescripcion.Text.ToString
            drMdb1.Close()
            cmdMdb1cr.CommandText = "SELECT * FROM apuntes Where apuntes.DescripcionAPU = '" & vDescripcion & "' ORDER BY apuntes.DescripcionAPU ASC"
            drMdb1 = cmdMdb1cr.ExecuteReader()
            drMdb1.Read()
            CmbDescripcion.Text = drMdb1.GetValue(3)
            drMdb1.Close()
        End If
    End Sub

    Private Sub CmbDescripcion_GotFocus(sender As Object, e As EventArgs) Handles CmbDescripcion.GotFocus
        If vCombo = "concepto" Then
            CmbConcepto.Text = ""
            TxtBuscarLetras.Enabled = False
            CmbConcepto.DroppedDown = False
            CmbConcepto.Text = vConcepto
        End If
        CmbDescripcion.Select()
        vCombo = ""
    End Sub

    Private Sub TxtImporte_GotFocus(sender As Object, e As EventArgs) Handles TxtImporte.GotFocus
        TxtBuscarLetras.Text = ""
        TxtBuscarLetras.Enabled = False
        vCombo = ""
    End Sub

    Private Sub CmbConcepto_GotFocus(sender As Object, e As EventArgs) Handles CmbConcepto.GotFocus
        TxtBuscarLetras.Enabled = False
        vIntro = "NO"
        If CmbConcepto.Items.Count <> 0 Then
            CmbConcepto.DroppedDown = True
            CmbConcepto.SelectedIndex = 0
        End If
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.TsLabelFormulario.Text = "Conceptos Contables"

        ' Comprobamos si existe un identificador asociado.
        If ((frmConceptosContables Is Nothing) OrElse (Not frmConceptosContables.IsHandleCreated)) Then
            frmConceptosContables = New ConceptosContables
        End If

        ' Llamamos al formulario de manera modal.
        frmConceptosContables.ShowDialog()

        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmConceptosContables.Dispose()
        frmPrincipal.TsLabelFormulario.Text = resManager.GetString("MsgEspera")

        ' Llenar el Combo Concepto al cerrar
        '***********************************
        If CmbConcepto.Items.Count <> 0 Then
            CmbConcepto.Items.Clear()
        End If
        LlenarConcepto()
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
            MsgBox("El año del ejercicio no coincide con el año actual," & vbCrLf & "se establecerá la fecha del 31 de Diciembre del año del ejercicio", MsgBoxStyle.Information, "Fecha establecida al 31 de Diciembre")
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
                    MsgBox("Error al Grabar el Registro, Verificar que la Fecha es Correcta y el Importe no tiene Letras ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Grabar")
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
                    vDate1 = Format(frmApuntesContables.DateTimePicker1.Value, "yyyy/MM/dd")
                    vDate2 = Format(frmApuntesContables.DateTimePicker2.Value, "yyyy/MM/dd")
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                vFilaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
                If vFilaActual = frmApuntesContables.DgvApuntes.RowCount - 1 Then
                    MsgBox("Fila Ultima Seleccionada")
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
                    MsgBox("Error al Grabar el Registro, Verificar que la Fecha es Correcta y el Importe no tiene Letras ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Grabar")
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
            MsgBox("NO hay Cantidad en Importe ...", vbExclamation)
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
                vAñadirSql = "INSERT INTO apuntes "
                vAñadirSql += "(FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) "
                vAñadirSql += "VALUES (#" & vDate3.ToString("yyyy/MM/dd") & "#,'" & vConcepto & "','" & vDescripcionAPU & "','" & vImporteAPU & "','" & vAñoEjercicio & "','" & vNotasAPU & "','" & vCuentaAPU & "')"
                cmdMdb1cr.CommandText = vAñadirSql
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox("Error al Grabar el Registro, Verificar que la Fecha es Correcta y el Importe no tiene Letras ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Grabar")
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
                    vDate1 = Format(frmApuntesContables.DateTimePicker1.Value, "yyyy/MM/dd")
                    vDate2 = Format(frmApuntesContables.DateTimePicker2.Value, "yyyy/MM/dd")
                    vtipoSql += " And apuntes.FechaAPU >= ?"
                    vtipoSql += " And apuntes.FechaAPU <= ?"
                End If
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC, apuntes.ImporteAPU ASC"
                vtipoGrid = "APUNTES_CONTABLES"
                LlenarGrid(vtipoSql, vtipoGrid, "1")
                vFila = frmApuntesContables.DgvApuntes.RowCount - 1
                frmApuntesContables.DgvApuntes.Rows(vFila).Selected = True
                frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(vFila).Cells(0)
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
                    MsgBox("Error al Grabar el Registro, Verificar que la Fecha es Correcta y el Importe no tiene Letras ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Grabar")
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
            TxtImporte.Text = 0
            DateTimePicker1.Select()
        Else
            MsgBox("NO hay Cantidad en Importe ...", vbExclamation)
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
        frmPrincipal.TsLabelFormulario.Text = "Cuentas Bancarias"

        ' Comprobamos si existe un identificador asociado.
        If ((frmCuentasBancarias Is Nothing) OrElse (Not frmCuentasBancarias.IsHandleCreated)) Then
            frmCuentasBancarias = New CuentasBancarias
        End If

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

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
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
        If CmbDescripcion.Items.Count <> 0 Then
            MsgBox("El Combo Descripción ya esta Lleno.", vbExclamation, "Combo Descripción")
            CmbDescripcion.Items.Clear()
        End If
        LlenarDescripcion()
    End Sub

    Function LlenarConcepto()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    If drMdb1.GetValue(0) <> "TRASPASO" Then
                        CmbConcepto.Items.Add(drMdb1.GetValue(0))
                    End If
                End While
                If frmApuntesContables.BtnFiltroConcepto.Enabled = False Then
                    If frmApuntesContables.ListBox1.SelectedItems.Count <> 0 Then
                        CmbConcepto.Text = CmbConcepto.Items(0)
                    Else
                        CmbConcepto.Text = CmbConcepto.Items(frmApuntesContables.CmbConcepto.SelectedIndex - 1)
                    End If
                Else
                    CmbConcepto.Text = CmbConcepto.Items(0)
                End If
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al Llenar el Combo Concepto, Verificar que la Base de Datos no esta Dañada ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Llenar")
        End Try
        Return ""
    End Function

    Function LlenarDescripcion()
        cmdMdb1cr.CommandText = "SELECT * FROM apuntes ORDER BY apuntes.DescripcionAPU ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                primero = 1
                While drMdb1.Read()
                    If Trim(drMdb1.GetValue(3)) <> "Saldo Inicial" Then
                        If primero = 1 Then
                            CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                            primero = 2
                        Else
                            nuevo = 0
                            For i = 0 To CmbDescripcion.Items.Count - 1
                                If Trim(drMdb1.GetValue(3)) = Trim(CmbDescripcion.Items(i)) Then
                                    nuevo = 0
                                    Exit For
                                Else
                                    nuevo = 1
                                End If
                            Next
                            If nuevo = 1 Then
                                CmbDescripcion.Items.Add(Trim(drMdb1.GetValue(3)))
                                nuevo = 0
                            End If
                        End If
                    End If
                End While
            Else
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al Llenar el Combo Descripción, Verificar que la Base de Datos no esta Dañada ..." & vbCrLf & "Error: " & ex.ToString, vbExclamation, "Error al Llenar")
        End Try
        Return ""
    End Function

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
End Class