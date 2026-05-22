Imports System.Data
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class EditarApuntes

    Public vConcepto, vtipoSql, vtipoGrid, vAñadirSql As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU As String
    Public vCodigoAPU As Integer
    Public vimporteAPU As Double
    Public i, primero, nuevo As Integer
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Label7.Text = vMoneda

        Dim TL(8) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, "Ir a Hoy")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnEliminar, "Eliminar Registro")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, "Aceptar")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnCancelar, "Cancelar la introducción del Apunte")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbConcepto, "Seleccionar el Concepto a la que se refiere la transacción")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuenta, "Seleccionar la Cuenta a la que se refiere la transacción")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.CmbDescripcion, "Introducir una descripción para el Asiento")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, "Importe del Asiento")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, "Activar la Calculadora")

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
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Descripción
        '****************************
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
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Descripción")
            MsgBox(ex.ToString)
        End Try

        ' Llenar el Combo Cuenta
        '***********************
        cmdMdb1cr.CommandText = "SELECT * FROM cuentas ORDER BY cuentas.NombreCUE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbCuenta.Items.Add(drMdb1.GetValue(0))
                End While
                CmbCuenta.Text = CmbCuenta.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Cuenta")
            MsgBox(ex.ToString)
        End Try

        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        DateTimePicker1.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(0).Value.ToString
        CmbConcepto.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString
        CmbDescripcion.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
        vimporteAPU = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(3).Value
        TxtImporte.Text = Format(Math.Abs(vimporteAPU), "###,##0.00").ToString
        TxtNota.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(5).Value.ToString
        CmbCuenta.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(6).Value.ToString
        vCodigoAPU = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(7).Value

        If vEditar = "SI" Then
            LblEditando.Text = "EDITANDO APUNTE CONTABLE"
            BtnEliminar.Enabled = False
        Else
            LblEditando.Text = "¡¡ ELIMINAR APUNTE CONTABLE !!"
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
        respuesta = MsgBox("¿Estas seguro de Eliminar el Apunte Contable seleccionado?.", vbQuestion & vbYesNo & vbDefaultButton2, "Eliminar Apunte Periódico")
        If respuesta = vbYes Then
            ' Eliminar Registro Apunte
            vtipoSql = "DELETE FROM apuntes"
            vtipoSql += " WHERE apuntes.CodigoAPU = " & vCodigoAPU.ToString
            cmdMdb1cr.CommandText = vtipoSql
            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox("Registro Apunte Contable, Borrado !!!")
            Catch ex As Exception
                MsgBox("Error al Eliminar el Registro Apunte Contable")
                MsgBox(ex.ToString)
            End Try
        Else
            frmApuntesContables.DgvApuntes.Rows(filaActual).Selected = True
            frmApuntesContables.DgvApuntes.CurrentCell = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(0)
        End If
        Me.Close()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged

        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMdb1.Close()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMdb1 = cmdMdb1cr.ExecuteReader()
        drMdb1.Read()
        TxtTipoConcepto.Text = drMdb1.GetValue(2)
        CmbDescripcion.Text = drMdb1.GetValue(1)
        drMdb1.Close()
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
        If TxtImporte.Text = "" Then
            TxtImporte.Text = 0
        End If
        If TxtImporte.Text <> "0" Then

            ' Modificar Registro
            '*******************
            vDate3 = DateTimePicker1.Value
            vimporteAPU = TxtImporte.Text
            If TxtTipoConcepto.Text = "GASTO" Then
                vimporteAPU = "-" + vimporteAPU.ToString
            End If
            vDescripcionAPU = ApostrofePorAcentoAgudo(CmbDescripcion.Text)

            Dim query As String = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) " &
                     "VALUES (?, ?, ?, ?, ?, ?, ?)"
            Using cmd As New OleDb.OleDbCommand(query, conexion1)
                cmd.Parameters.Add("@FechaAPU", OleDb.OleDbType.Date).Value = vDate3 'es importante usar el tipo de dato correcto para la fecha
                cmd.Parameters.AddWithValue("?", CmbConcepto.Text)
                cmd.Parameters.AddWithValue("?", vDescripcionAPU)
                cmd.Parameters.Add("@ImporteAPU", OleDb.OleDbType.Currency).Value = vimporteAPU 'es importante usar el tipo de dato correcto para el importe
                cmd.Parameters.AddWithValue("?", vAñoEjercicio)
                cmd.Parameters.AddWithValue("?", TxtNota.Text)
                cmd.Parameters.AddWithValue("?", CmbCuenta.Text)
                Try
                    cmd.ExecuteNonQuery()
                    'MsgBox("Registro, Grabado Correctamente")
                Catch ex As Exception
                    MsgBox("Error al insertar el Apunte de Saldo Inicial de la Cuenta: " & vNombreCuenta & " del Ejercicio " & vAñoEjercicio.ToString & vbCrLf & ex.ToString)
                End Try
            End Using
            vtipoSql = "UPDATE apuntes SET  FechaAPU =?, ConceptoAPU = '" & CmbConcepto.Text & "' , DescripcionAPU = '" & vDescripcionAPU & "' , ImporteAPU = '" & vimporteAPU & "' , CuentaAPU = '" & CmbCuenta.Text & "' , NotasAPU = '" & TxtNota.Text & "' "
            vtipoSql += " WHERE apuntes.CodigoAPU = " & vCodigoAPU.ToString
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = vDate3
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                'MsgBox("Registro, Grabado Correctamente")
            Catch ex As Exception
                MsgBox("Error al Grabar el Registro Apunte Contable")
                MsgBox(ex.ToString)
            End Try
            drMdb1.Close()
        Else
            MsgBox("NO hay Cantidad en Importe ...", vbExclamation)
            TxtImporte.Select()
        End If
        Me.Close()
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
    End Sub

    Private Sub BtnCuenta_Click(sender As Object, e As EventArgs)
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

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuenta_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuenta.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class