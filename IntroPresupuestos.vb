Imports System.Windows.Forms

Public Class IntroPresupuestos

    Public vConcepto, vtipoSql, vAñadirSql, vFDesde, vBorrarPresu As String
    Public vMensual, vAnual, vEnero, vFebrero, vMarzo, vAbril, vMayo, vJunio, vSaldoAnualPresupuesto, vImporte As Double
    Public vJulio, vAgosto, vSeptiembre, vOctubre, vNoviembre, vDiciembre As Double
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Label16.Text = vMoneda
        Label17.Text = vMoneda
        Label18.Text = vMoneda
        Label19.Text = vMoneda
        Label20.Text = vMoneda
        Label21.Text = vMoneda
        Label22.Text = vMoneda
        Label23.Text = vMoneda
        Label24.Text = vMoneda
        Label25.Text = vMoneda
        Label26.Text = vMoneda
        Label27.Text = vMoneda
        Label28.Text = vMoneda

        Dim TL(18) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnConcepto, "Añade, Edita, Borra o Consulta Conceptos Contables")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnAceptar, "Aceptar, Guardar y Salir")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnCancelar, "Cancelar y Salir")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.CmbConcepto, "Seleccionar el Concepto Contable para definir el Presupuesto")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.TxtAnual, "Introducir el Total Anual Presupuestado, que se repartirá mensualmente")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.TxtEnero, "Introducir la cantidad Presupuestada para este mes")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.TxtFebrero, "Introducir la cantidad Presupuestada para este mes")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.TxtMarzo, "Introducir la cantidad Presupuestada para este mes")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.TxtAbril, "Introducir la cantidad Presupuestada para este mes")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.TxtMayo, "Introducir la cantidad Presupuestada para este mes")
        TL(10) = New ToolTip
        TL(10).SetToolTip(Me.TxtJunio, "Introducir la cantidad Presupuestada para este mes")
        TL(11) = New ToolTip
        TL(11).SetToolTip(Me.TxtJulio, "Introducir la cantidad Presupuestada para este mes")
        TL(12) = New ToolTip
        TL(12).SetToolTip(Me.TxtAgosto, "Introducir la cantidad Presupuestada para este mes")
        TL(13) = New ToolTip
        TL(13).SetToolTip(Me.TxtSeptiembre, "Introducir la cantidad Presupuestada para este mes")
        TL(14) = New ToolTip
        TL(14).SetToolTip(Me.TxtOctubre, "Introducir la cantidad Presupuestada para este mes")
        TL(15) = New ToolTip
        TL(15).SetToolTip(Me.TxtNoviembre, "Introducir la cantidad Presupuestada para este mes")
        TL(16) = New ToolTip
        TL(16).SetToolTip(Me.TxtDiciembre, "Introducir la cantidad Presupuestada para este mes")
        TL(17) = New ToolTip
        TL(17).SetToolTip(Me.RdbAnual, "Con esta selección introducir una cantidad en la casilla Total Anual y se repartirá a partes iguales entre todos los meses")
        TL(18) = New ToolTip
        TL(18).SetToolTip(Me.RdbMensual, "Con esta selección introducir una cantidad en la casillas de cada mes")


        ' Llenar el Combo Concepto
        '*************************
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                While drMySql1.Read()
                    CmbConcepto.Items.Add(drMySql1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar el Combo Concepto")
            MsgBox(ex.ToString)
        End Try

        If RdbAnual.Checked = True Then
            GBoxAnual.Enabled = True
            GBoxMensual.Enabled = False
            TxtAnual.Select()
            TxtAnual.SelectAll()
        End If
        If RdbMensual.Checked = True Then
            GBoxAnual.Enabled = False
            GBoxMensual.Enabled = True
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If
        LlenarTextBox()
    End Sub

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMySql1.Close()
        cmdMySql1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMySql1 = cmdMySql1cr.ExecuteReader()
        drMySql1.Read()
        TxtTipoConcepto.Text = drMySql1.GetValue(2)
        TxtDescripcion.Text = drMySql1.GetValue(1)
        drMySql1.Close()
        If RdbAnual.Checked = True Then
            TxtAnual.Select()
            TxtAnual.SelectAll()
        Else
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If
        LlenarTextBox()
    End Sub

    Function LlenarTextBox()
        TxtEnero.Text = "0,00"
        TxtFebrero.Text = "0,00"
        TxtMarzo.Text = "0,00"
        TxtAbril.Text = "0,00"
        TxtMayo.Text = "0,00"
        TxtJunio.Text = "0,00"
        TxtJulio.Text = "0,00"
        TxtAgosto.Text = "0,00"
        TxtSeptiembre.Text = "0,00"
        TxtOctubre.Text = "0,00"
        TxtNoviembre.Text = "0,00"
        TxtDiciembre.Text = "0,00"
        TxtAnual.Text = "0,00"

        ' Llenar TextBox si hay datos guardados en PRESUPUESTOS
        '******************************************************
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " And presupuesto.ConceptoPRE = '" & vConcepto & "' "
        cmdMySql1cr.CommandText = vtipoSql
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                vSaldoAnualPresupuesto = 0
                While drMySql1.Read()
                    vSaldoAnualPresupuesto += drMySql1.GetValue(1)
                    vFecha = drMySql1.GetValue(2).ToString
                    vMes = Mid(vFecha, 4, 2).ToString
                    If vMes = "01" Then
                        vEnero = drMySql1.GetValue(1).ToString
                        TxtEnero.Text = Format(vEnero, "###,##0.00").ToString
                    End If
                    If vMes = "02" Then
                        vFebrero = drMySql1.GetValue(1).ToString
                        TxtFebrero.Text = Format(vFebrero, "###,##0.00").ToString
                    End If
                    If vMes = "03" Then
                        vMarzo = drMySql1.GetValue(1).ToString
                        TxtMarzo.Text = Format(vMarzo, "###,##0.00").ToString
                    End If
                    If vMes = "04" Then
                        vAbril = drMySql1.GetValue(1).ToString
                        TxtAbril.Text = Format(vAbril, "###,##0.00").ToString
                    End If
                    If vMes = "05" Then
                        vMayo = drMySql1.GetValue(1).ToString
                        TxtMayo.Text = Format(vMayo, "###,##0.00").ToString
                    End If
                    If vMes = "06" Then
                        vJunio = drMySql1.GetValue(1).ToString
                        TxtJunio.Text = Format(vJunio, "###,##0.00").ToString
                    End If
                    If vMes = "07" Then
                        vJulio = drMySql1.GetValue(1).ToString
                        TxtJulio.Text = Format(vJulio, "###,##0.00").ToString
                    End If
                    If vMes = "08" Then
                        vAgosto = drMySql1.GetValue(1).ToString
                        TxtAgosto.Text = Format(vAgosto, "###,##0.00").ToString
                    End If
                    If vMes = "09" Then
                        vSeptiembre = drMySql1.GetValue(1).ToString
                        TxtSeptiembre.Text = Format(vSeptiembre, "###,##0.00").ToString
                    End If
                    If vMes = "10" Then
                        vOctubre = drMySql1.GetValue(1).ToString
                        TxtOctubre.Text = Format(vOctubre, "###,##0.00").ToString
                    End If
                    If vMes = "11" Then
                        vNoviembre = drMySql1.GetValue(1).ToString
                        TxtNoviembre.Text = Format(vNoviembre, "###,##0.00").ToString
                    End If
                    If vMes = "12" Then
                        vDiciembre = drMySql1.GetValue(1).ToString
                        TxtDiciembre.Text = Format(vDiciembre, "###,##0.00").ToString
                    End If
                End While
                TxtAnual.Text = Format(vSaldoAnualPresupuesto, "###,##0.00").ToString
                TxtAnual.SelectAll()
            Else
                'MsgBox("No existen registros en " & cmdMySql1cr.CommandText)
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox("Error al llenar los TextBox con los datos guardados en presupuesto")
            MsgBox(ex.ToString)
        End Try
        Return vSaldoAnualPresupuesto
    End Function

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vConcepto = CmbConcepto.Text
        If TxtAnual.Text <> 0 Then
            ' Comprobamos que el Concepto si esta en presupuesto
            cmdMySql1cr.CommandText = "SELECT * FROM presupuesto"
            cmdMySql1cr.CommandText += " WHERE "
            cmdMySql1cr.CommandText += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
            cmdMySql1cr.CommandText += " And presupuesto.ConceptoPRE = '" & vConcepto & "' "
            Try
                drMySql1 = cmdMySql1cr.ExecuteReader()
                If drMySql1.HasRows Then
                    respuesta = MsgBox("Hay datos guardados con este Concepto ¿Se actualizan valores?.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar presupuesto")
                    If respuesta = vbYes Then
                        drMySql1.Close()
                        ' Eliminar Registros mismo Concepto Apunte
                        vtipoSql = "DELETE FROM presupuesto"
                        vtipoSql += " WHERE presupuesto.ConceptoPRE = '" & vConcepto & "' "
                        cmdMySql1cr.CommandText = vtipoSql
                        Try
                            cmdMySql1cr.ExecuteNonQuery()
                            MsgBox("Registros en presupuesto, Borrados !!!")
                        Catch ex As Exception
                            MsgBox("Error al Borrar el presupuesto")
                            MsgBox(ex.ToString)
                        End Try
                    Else
                        MsgBox("No actualiza nada y sale")
                        Me.Close()
                    End If
                Else
                    MsgBox("No existen registros en " & vConcepto)
                End If
                drMySql1.Close()
            Catch ex As Exception
                MsgBox("Error al verificar que el Concepto existe en presupuesto")
                MsgBox(ex.ToString)
            End Try
            If TxtAnual.Text <> "0,00" Then
                For i = 1 To 12
                    If i = 1 Then
                        vImporte = TxtEnero.Text
                        vFDesde = "01/01/" & vAñoEjercicio.ToString
                    End If
                    If i = 2 Then
                        vImporte = TxtFebrero.Text
                        vFDesde = "01/02/" & vAñoEjercicio.ToString
                    End If
                    If i = 3 Then
                        vImporte = TxtMarzo.Text
                        vFDesde = "01/03/" & vAñoEjercicio.ToString
                    End If
                    If i = 4 Then
                        vImporte = TxtAbril.Text
                        vFDesde = "01/04/" & vAñoEjercicio.ToString
                    End If
                    If i = 5 Then
                        vImporte = TxtMayo.Text
                        vFDesde = "01/05/" & vAñoEjercicio.ToString
                    End If
                    If i = 6 Then
                        vImporte = TxtJunio.Text
                        vFDesde = "01/06/" & vAñoEjercicio.ToString
                    End If
                    If i = 7 Then
                        vImporte = TxtJulio.Text
                        vFDesde = "01/07/" & vAñoEjercicio.ToString
                    End If
                    If i = 8 Then
                        vImporte = TxtAgosto.Text
                        vFDesde = "01/08/" & vAñoEjercicio.ToString
                    End If
                    If i = 9 Then
                        vImporte = TxtSeptiembre.Text
                        vFDesde = "01/09/" & vAñoEjercicio.ToString
                    End If
                    If i = 10 Then
                        vImporte = TxtOctubre.Text
                        vFDesde = "01/10/" & vAñoEjercicio.ToString
                    End If
                    If i = 11 Then
                        vImporte = TxtNoviembre.Text
                        vFDesde = "01/11/" & vAñoEjercicio.ToString
                    End If
                    If i = 12 Then
                        vImporte = TxtDiciembre.Text
                        vFDesde = "01/12/" & vAñoEjercicio.ToString
                    End If
                    vAñadirSql = "INSERT INTO presupuesto "
                    vAñadirSql += "(ConceptoPRE, ImportePRE, FDesdePRE, EjercicioPRE) "
                    vAñadirSql += "VALUES ('" & vConcepto & "','" & vImporte & "','" & vFDesde & "','" & vAñoEjercicio & "')"
                    cmdMySql1cr.CommandText = vAñadirSql
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        If i = 12 Then
                            'MsgBox("Registro, Grabado Correctamente")
                        End If
                    Catch ex As Exception
                        MsgBox("Error al Grabar el Presupuesto")
                        MsgBox(ex.ToString)
                    End Try
                Next
            End If
            Me.Close()
        Else
            MsgBox("No hay Importes")
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
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
    End Sub

    Private Sub RdbAnual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbAnual.CheckedChanged
        If RdbAnual.Checked = True Then
            GBoxAnual.Enabled = True
            GBoxMensual.Enabled = False
            TxtAnual.Select()
            TxtAnual.SelectAll()
        End If
    End Sub

    Private Sub RdbMensual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbMensual.CheckedChanged
        If RdbMensual.Checked = True Then
            GBoxAnual.Enabled = False
            GBoxMensual.Enabled = True
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtAnual_Click(sender As Object, e As EventArgs) Handles TxtAnual.Click
        TxtAnual.Select()
        TxtAnual.SelectAll()
    End Sub

    Private Sub TxtAnual_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAnual.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vAnual = TxtAnual.Text
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString

            vMensual = TxtAnual.Text / 12
            TxtEnero.Text = Format(vMensual, "###,##0.00").ToString
            vEnero = TxtEnero.Text
            TxtFebrero.Text = Format(vMensual, "###,##0.00").ToString
            vFebrero = TxtFebrero.Text
            TxtMarzo.Text = Format(vMensual, "###,##0.00").ToString
            vMarzo = TxtMarzo.Text
            TxtAbril.Text = Format(vMensual, "###,##0.00").ToString
            vAbril = TxtAbril.Text
            TxtMayo.Text = Format(vMensual, "###,##0.00").ToString
            vMayo = TxtMayo.Text
            TxtJunio.Text = Format(vMensual, "###,##0.00").ToString
            vJunio = TxtJunio.Text
            TxtJulio.Text = Format(vMensual, "###,##0.00").ToString
            vJulio = TxtJulio.Text
            TxtAgosto.Text = Format(vMensual, "###,##0.00").ToString
            vAgosto = TxtAgosto.Text
            TxtSeptiembre.Text = Format(vMensual, "###,##0.00").ToString
            vSeptiembre = TxtSeptiembre.Text
            TxtOctubre.Text = Format(vMensual, "###,##0.00").ToString
            vOctubre = TxtOctubre.Text
            TxtNoviembre.Text = Format(vMensual, "###,##0.00").ToString
            vNoviembre = TxtNoviembre.Text
            TxtDiciembre.Text = Format(vMensual, "###,##0.00").ToString
            vDiciembre = TxtDiciembre.Text
            RdbMensual.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtEnero_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtEnero.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vEnero = TxtEnero.Text
            TxtEnero.Text = Format(vEnero, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtFebrero.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtFebrero_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtFebrero.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vFebrero = TxtFebrero.Text
            TxtFebrero.Text = Format(vFebrero, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtMarzo.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtMarzo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtMarzo.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vMarzo = TxtMarzo.Text
            TxtMarzo.Text = Format(vMarzo, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtAbril.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtAbril_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAbril.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vAbril = TxtAbril.Text
            TxtAbril.Text = Format(vAbril, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtMayo.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtMayo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtMayo.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vMayo = TxtMayo.Text
            TxtMayo.Text = Format(vMayo, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtJunio.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtJunio_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtJunio.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vJunio = TxtJunio.Text
            TxtJunio.Text = Format(vJunio, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtJulio.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtJulio_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtJulio.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vJulio = TxtJulio.Text
            TxtJulio.Text = Format(vJulio, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtAgosto.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtAgosto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtAgosto.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vAgosto = TxtAgosto.Text
            TxtAgosto.Text = Format(vAgosto, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtSeptiembre.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtSeptiembre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtSeptiembre.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vSeptiembre = TxtSeptiembre.Text
            TxtSeptiembre.Text = Format(vSeptiembre, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtOctubre.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtOctubre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtOctubre.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vOctubre = TxtOctubre.Text
            TxtOctubre.Text = Format(vOctubre, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtNoviembre.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtNoviembre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNoviembre.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vNoviembre = TxtNoviembre.Text
            TxtNoviembre.Text = Format(vNoviembre, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            TxtDiciembre.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub TxtDiciembre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDiciembre.KeyPress
        SoloNumerosConPunto(e)
        If e.KeyChar = ChrW(Keys.Enter) Then
            vDiciembre = TxtDiciembre.Text
            TxtDiciembre.Text = Format(vDiciembre, "###,##0.00").ToString
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00").ToString
            BtnAceptar.Select()
        End If
        If e.KeyChar.ToString() = "." Then
            e.KeyChar = ","
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class