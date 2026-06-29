Imports System.Data
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class TraspasoCuentas

    Private cargandoFormulario As Boolean = True
    Public vConcepto, vAñadirOrigenSql, vAñadirDestinoSql As String
    Public vImporteAPU As Double
    Public vDescripcionAPU, vNotasAPU, vCuentaOrigenAPU, vCuentaDestinoAPU As String
    Public vfechaHoyOrigen As Date = DateTime.Today
    Private TL(11) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub TraspasoCuentas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cargandoFormulario = True
        Me.KeyPreview = True

        Label7.Text = vMoneda
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoyOrigen, resManager.GetString("IrAHoy"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.CmbConcepto, rmse.GetString("ToolTipSeleccionarConcepto"))
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.CmbCuentaOrigen, rmse.GetString("SeleccionarCuentaOrigen"))
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.CmbCuentaDestino, rmse.GetString("SeleccionarCuentaDestino"))
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtDescripcion, rmse.GetString("IntroDescripcion"))
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtImporte, rmse.GetString("ImporteApunte"))
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnCalculadora, resManager.GetString("ToolTipCalculadora"))
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnConcepto, resManager.GetString("BtnConcepto"))
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.BtnCuentaOrigen, resManager.GetString("BtnCuenta"))
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.BtnCuentaDestino, resManager.GetString("BtnCuenta"))

        ' Llenar los Combo
        '*****************
        Try
            ' Usamos la función exclusiva que no carga los 'ESPECIALES' si es para introducir/editar ordinarios
            ' (O la que uses en este formulario, pero asegurando que use DataTable)
            LlenarComboConceptoExclusivoTraspaso(Me.CmbConcepto)
            LlenarComboCuentasGenerico(Me.CmbCuentaOrigen)
            LlenarComboCuentasGenerico(Me.CmbCuentaDestino)
        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorCargarCONyCUE") & ": " & ex.Message, MsgBoxStyle.Critical)
        End Try
        TxtImporte.Text = 0
        cargandoFormulario = False
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' 🌟 ESCUDO PROTECTOR AUTOMÁTICO: Si el formulario se está iniciando o limpiando, salimos de inmediato
        If cargandoFormulario Then Exit Sub
        If CmbConcepto.SelectedIndex < 0 Then Exit Sub

        Try
            Dim tipoOriginal As String = ""

            ' 🌟 EXTRACCIÓN MAESTRA DESDE MEMORIA (Cero consultas DataReader a Access)
            ' Como el combo está enlazado a un DataTable, convertimos el ítem actual en un DataRowView
            If CmbConcepto.SelectedItem IsNot Nothing Then
                Dim filaSeleccionada As DataRowView = CType(CmbConcepto.SelectedItem, DataRowView)

                ' Leemos la columna TipoCON directamente de la memoria caché de la app
                If filaSeleccionada.Row.Table.Columns.Contains("TipoCON") Then
                    tipoOriginal = filaSeleccionada("TipoCON").ToString().Trim()
                End If
            End If

            ' Sincronizamos y traducimos el tipo a la interfaz de forma dócil y limpia
            Dim tradTipo As String = ""
            Select Case tipoOriginal.ToUpper()
                Case "GASTO" : tradTipo = resManager.GetString("Tipo_Gasto")
                Case "INGRESO" : tradTipo = resManager.GetString("Tipo_Ingreso")
                Case "ESPECIAL" : tradTipo = resManager.GetString("Tipo_Especial")
            End Select

            If String.IsNullOrEmpty(tradTipo) Then tradTipo = tipoOriginal
            TxtTipoConcepto.Text = tradTipo

        Catch ex As Exception
            MsgBox(resManager.GetString("ErrorSincronizarCON") & ": " & ex.Message, MsgBoxStyle.Critical, resManager.GetString("Error"))
        End Try
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtImporte.Select()
            TxtImporte.SelectAll()
        End If
    End Sub

    Private Sub TxtImporte_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtImporte.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            TxtNota.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtImporte_Click(sender As Object, e As EventArgs) Handles TxtImporte.Click
        TxtImporte.SelectAll()
    End Sub

    Private Sub TxtNota_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNota.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnHoyOrigen_Click(sender As Object, e As EventArgs) Handles BtnHoyOrigen.Click
        DtpOrigen.Value = vfechaHoyOrigen
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 🌟 COMPARACIÓN DE SEGURIDAD: Comparamos los IDs numéricos directamente para saber si son la misma cuenta
        If CmbCuentaOrigen.SelectedValue IsNot Nothing AndAlso CmbCuentaDestino.SelectedValue IsNot Nothing AndAlso
           CmbCuentaOrigen.SelectedValue.ToString() <> CmbCuentaDestino.SelectedValue.ToString() Then

            If TxtDescripcion.Text.Trim() <> "" Then
                If TxtImporte.Text.Trim() <> "" And TxtImporte.Text.Trim() <> "0" Then

                    ' 1. EXTRAEMOS LOS IDs NUMÉRICOS PUROS DESDE LOS COMBOS (Nueva era relacional)
                    Dim idConcepto As Integer = Convert.ToInt32(CmbConcepto.SelectedValue)
                    Dim idCuentaOrigen As Integer = Convert.ToInt32(CmbCuentaOrigen.SelectedValue)
                    Dim idCuentaDestino As Integer = Convert.ToInt32(CmbCuentaDestino.SelectedValue)

                    ' 2. Capturamos los textos limpios (¡Los parámetros ya protegen los apóstrofes solos!)
                    vDescripcionAPU = TxtDescripcion.Text.Trim()
                    vNotasAPU = TxtNota.Text.Trim()

                    ' 3. Procesamos los importes contables en formato Decimal seguro
                    Dim importeOrigen As Decimal = -Math.Abs(ConvertirDecimalSeguro(TxtImporte.Text))
                    Dim importeDestino As Decimal = Math.Abs(ConvertirDecimalSeguro(TxtImporte.Text))

                    ' =========================================================================
                    ' 🌟 TRASPASO RAMA 1: GRABACIÓN DE LA CUENTA ORIGEN (NEGATIVO)
                    ' =========================================================================
                    vAñadirOrigenSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) " &
                                       "VALUES (?, ?, ?, ?, ?, ?, ?)"
                    cmdMdb1cr.CommandText = vAñadirOrigenSql
                    cmdMdb1cr.Parameters.Clear()

                    ' Inyectamos los parámetros en estricto orden biológico de los signos '?'
                    cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = DtpOrigen.Value.Date
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConcepto       ' ID Numérico
                    cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = vDescripcionAPU
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = Math.Round(importeOrigen, 2)
                    cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
                    cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = vNotasAPU
                    cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuentaOrigen   ' ID Numérico

                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        MsgBox(rmse.GetString("RegistroOrigenGrabadoCorrectamente"))
                    Catch ex As Exception
                        MsgBox(resManager.GetString("Error") & resManager.GetString("Origen") & ": " & ex.Message, MsgBoxStyle.Critical)
                    End Try

                    ' =========================================================================
                    ' 🌟 TRASPASO RAMA 2: GRABACIÓN DE LA CUENTA DESTINO (POSITIVO)
                    ' =========================================================================
                    vAñadirDestinoSql = "INSERT INTO apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, NotasAPU, CuentaAPU) " &
                                        "VALUES (?, ?, ?, ?, ?, ?, ?)"
                    cmdMdb1cr.CommandText = vAñadirDestinoSql
                    cmdMdb1cr.Parameters.Clear()

                    ' Inyectamos los parámetros en estricto orden para el contraasiento
                    cmdMdb1cr.Parameters.Add("@fec", OleDb.OleDbType.Date).Value = DtpDestino.Value.Date
                    cmdMdb1cr.Parameters.Add("@con", OleDb.OleDbType.Integer).Value = idConcepto       ' ID Numérico
                    cmdMdb1cr.Parameters.Add("@des", OleDb.OleDbType.VarWChar).Value = vDescripcionAPU
                    cmdMdb1cr.Parameters.Add("@imp", OleDb.OleDbType.Currency).Value = Math.Round(importeDestino, 2)
                    cmdMdb1cr.Parameters.Add("@eje", OleDb.OleDbType.Integer).Value = Convert.ToInt32(vAñoEjercicio)
                    cmdMdb1cr.Parameters.Add("@not", OleDb.OleDbType.VarWChar).Value = vNotasAPU
                    cmdMdb1cr.Parameters.Add("@cue", OleDb.OleDbType.Integer).Value = idCuentaDestino  ' ID Numérico

                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        MsgBox(rmse.GetString("RegistroDestinoGrabadoCorrectamente"))
                    Catch ex As Exception
                        MsgBox(resManager.GetString("Error") & resManager.GetString("Destino") & ": " & ex.Message, MsgBoxStyle.Critical)
                    End Try

                    ' =========================================================================
                    ' 🌟 REFRESCAMOS LA REJILLA DE ATRÁS AUTOMÁTICAMENTE ANTES DE SALIR
                    ' =========================================================================
                    ' Reutilizamos la rutina mágica global que creamos ayer para que la pantalla principal
                    ' se actualice al instante con los dos nuevos apuntes sin escribir código repetido
                    If TypeOf frmApuntesContables Is Form Then
                        RefrescarGridApuntesContables()
                    End If

                    Me.Close()
                Else
                    MsgBox(rmse.GetString("NoHayImporte") & "...", vbExclamation)
                    TxtImporte.Select()
                End If
            Else
                MsgBox(rmse.GetString("DescripcionVacia"), vbExclamation)
            End If
        Else
            MsgBox(rmse.GetString("CuentasDiferentes"), vbExclamation)
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub BtnCuentaOrigen_Click(sender As Object, e As EventArgs) Handles BtnCuentaOrigen.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnCuentaDestino_Click(sender As Object, e As EventArgs) Handles BtnCuentaDestino.Click
        frmPrincipal.CuentasToolStripMenuItem.PerformClick()
    End Sub

    Private Sub BtnConcepto_Click(sender As Object, e As EventArgs) Handles BtnConcepto.Click
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()
    End Sub

    Private Sub DtpOrigen_ValueChanged(sender As Object, e As EventArgs) Handles DtpOrigen.ValueChanged
        DtpDestino.Value = DtpOrigen.Value
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

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuentaOrigen_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuentaOrigen.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub

    Private Sub CmbCuentaDestino_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbCuentaDestino.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class