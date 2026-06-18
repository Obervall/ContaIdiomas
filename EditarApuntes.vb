Imports System.Diagnostics
Imports System.Windows.Forms

Public Class EditarApuntes

    Public vConcepto, vtipoSql, vtipoGrid As String
    Public vDescripcionAPU, vNotasAPU, vCuentaAPU As String
    Public vCodigoAPU As Integer
    Public vimporteAPU As Double
    Public i, primero, nuevo As Integer
    Private TL(8) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub EditarApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True
        'ActualizarTextosFormulario(Me)

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
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            'MsgBox("Error al llenar el Combo Descripción")
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
            'MsgBox("Error al llenar el Combo Cuenta")
            MsgBox(ex.ToString)
        End Try

        filaActual = frmApuntesContables.DgvApuntes.CurrentRow.Index
        DateTimePicker1.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(0).Value.ToString
        CmbConcepto.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(1).Value.ToString
        CmbDescripcion.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(2).Value.ToString
        vimporteAPU = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(3).Value
        TxtImporte.Text = Math.Abs(vimporteAPU).ToString("N2")
        TxtNota.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(5).Value.ToString
        CmbCuenta.Text = frmApuntesContables.DgvApuntes.Rows(filaActual).Cells(6).Value.ToString
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
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Sincronizar y limpiar las variables de texto
        vDate3 = DateTimePicker1.Value
        vConceptoAPU = Trim(CmbConcepto.Text)
        vDescripcionAPU = Trim(CmbDescripcion.Text)
        vNotasAPU = Trim(TxtNota.Text)
        vCuentaAPU = Trim(CmbCuenta.Text)

        ' 1. Usamos el tipo Decimal (imprescindible para contabilidad)
        Dim importeDecimal As Decimal = 0.0D

        ' 2. Forzamos la limpieza básica de espacios
        Dim textoLimpio As String = TxtImporte.Text.Trim()

        ' 3. Intentamos leer con la cultura regional del usuario (respeta su panel de control)
        If Not Decimal.TryParse(textoLimpio,
                        System.Globalization.NumberStyles.Number,
                        System.Globalization.CultureInfo.CurrentCulture,
                        importeDecimal) Then

            ' 4. PLAN B SEGURO: Intentamos con la cultura invariante (punto decimal universal) 
            ' SIN hacer .Replace manual, para que .NET gestione los miles correctamente.
            If Not Decimal.TryParse(textoLimpio,
                            System.Globalization.NumberStyles.Number,
                            System.Globalization.CultureInfo.InvariantCulture,
                            importeDecimal) Then
                ' Si introduce letras o formato corrupto, se queda en 0 por seguridad
                importeDecimal = 0.0D
            End If
        End If

        ' 4. Guardamos el importe limpio en tu variable y aplicamos el signo
        vimporteAPU = importeDecimal
        If TxtTipoConcepto.Text = "GASTO" Then
            vimporteAPU = -Math.Abs(vimporteAPU)
        Else
            vimporteAPU = Math.Abs(vimporteAPU)
        End If

        ' Creamos la consulta limpia usando comodines '?' (así lo requiere Access/OleDb)
        vtipoSql = "UPDATE apuntes SET " &
           "FechaAPU = ?, " &
           "ConceptoAPU = ?, " &
           "DescripcionAPU = ?, " &
           "ImporteAPU = ?, " &
           "NotasAPU = ?, " &
           "CuentaAPU = ? " &
           "WHERE CodigoAPU = ?"

        cmdMdb1cr.CommandText = vtipoSql

        ' ¡MUY IMPORTANTE! En Access, los parámetros se deben añadir en el MISMO ORDEN en que aparecen en el SQL
        cmdMdb1cr.Parameters.Clear()
        cmdMdb1cr.Parameters.AddWithValue("@FechaAPU", vDate3) ' .NET y Access gestionan la fecha internamente
        cmdMdb1cr.Parameters.AddWithValue("@ConceptoAPU", vConceptoAPU) ' Admite comillas de forma segura
        cmdMdb1cr.Parameters.AddWithValue("@DescripcionAPU", vDescripcionAPU)
        cmdMdb1cr.Parameters.AddWithValue("@ImporteAPU", vimporteAPU) ' Envía el Decimal puro, sin importar puntos o comas de Windows
        cmdMdb1cr.Parameters.AddWithValue("@NotasAPU", vNotasAPU)
        cmdMdb1cr.Parameters.AddWithValue("@CuentaAPU", vCuentaAPU)
        cmdMdb1cr.Parameters.AddWithValue("@CodigoAPU", CInt(vCodigoAPU)) ' Filtro WHERE

        Try
            ' Ejecutamos la consulta en Access de forma segura
            cmdMdb1cr.ExecuteNonQuery()

            ' Cerramos la ventana de edición al terminar con éxito
            Me.Close()
        Catch ex As Exception
            MsgBox(rmse.GetString("MsgBoxErrorInsertar") & ": " & ex.Message, MsgBoxStyle.Critical, rmse.GetString("$this.Text"))
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
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMdb1.Close()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto.Replace("'", "''") & "' "
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