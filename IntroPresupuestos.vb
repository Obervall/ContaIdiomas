Imports System.Data.OleDb
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

    'Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
    '    ' Se buscan Conceptos según lo seleccionado
    '    '******************************************
    '    vConcepto = CmbConcepto.Text.ToString
    '    drMdb1.Close()
    '    cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
    '    drMdb1 = cmdMdb1cr.ExecuteReader()
    '    drMdb1.Read()
    '    TxtTipoConcepto.Text = drMdb1.GetValue(2)
    '    TxtDescripcion.Text = drMdb1.GetValue(1)
    '    drMdb1.Close()
    '    If RdbAnual.Checked = True Then
    '        TxtAnual.Select()
    '        TxtAnual.SelectAll()
    '    Else
    '        TxtEnero.Select()
    '        TxtEnero.SelectAll()
    '    End If
    '    LlenarTextBox()
    'End Sub

    'Function LlenarTextBox()
    '    TxtEnero.Text = "0,00"
    '    TxtFebrero.Text = "0,00"
    '    TxtMarzo.Text = "0,00"
    '    TxtAbril.Text = "0,00"
    '    TxtMayo.Text = "0,00"
    '    TxtJunio.Text = "0,00"
    '    TxtJulio.Text = "0,00"
    '    TxtAgosto.Text = "0,00"
    '    TxtSeptiembre.Text = "0,00"
    '    TxtOctubre.Text = "0,00"
    '    TxtNoviembre.Text = "0,00"
    '    TxtDiciembre.Text = "0,00"
    '    TxtAnual.Text = "0,00"

    '    ' Llenar TextBox si hay datos guardados en PRESUPUESTOS
    '    '******************************************************
    '    vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
    '    vtipoSql += " WHERE "
    '    vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
    '    vtipoSql += " And presupuesto.ConceptoPRE = '" & vConcepto & "' "
    '    cmdMdb1cr.CommandText = vtipoSql
    '    Try
    '        drMdb1 = cmdMdb1cr.ExecuteReader()
    '        If drMdb1.HasRows Then
    '            vSaldoAnualPresupuesto = 0
    '            While drMdb1.Read()
    '                vSaldoAnualPresupuesto += drMdb1.GetValue(1)
    '                vFecha = drMdb1.GetValue(2).ToString
    '                vMes = Mid(vFecha, 4, 2).ToString
    '                If vMes = "01" Then
    '                    vEnero = drMdb1.GetValue(1).ToString
    '                    TxtEnero.Text = Format(vEnero, "###,##0.00").ToString
    '                End If
    '                If vMes = "02" Then
    '                    vFebrero = drMdb1.GetValue(1).ToString
    '                    TxtFebrero.Text = Format(vFebrero, "###,##0.00").ToString
    '                End If
    '                If vMes = "03" Then
    '                    vMarzo = drMdb1.GetValue(1).ToString
    '                    TxtMarzo.Text = Format(vMarzo, "###,##0.00").ToString
    '                End If
    '                If vMes = "04" Then
    '                    vAbril = drMdb1.GetValue(1).ToString
    '                    TxtAbril.Text = Format(vAbril, "###,##0.00").ToString
    '                End If
    '                If vMes = "05" Then
    '                    vMayo = drMdb1.GetValue(1).ToString
    '                    TxtMayo.Text = Format(vMayo, "###,##0.00").ToString
    '                End If
    '                If vMes = "06" Then
    '                    vJunio = drMdb1.GetValue(1).ToString
    '                    TxtJunio.Text = Format(vJunio, "###,##0.00").ToString
    '                End If
    '                If vMes = "07" Then
    '                    vJulio = drMdb1.GetValue(1).ToString
    '                    TxtJulio.Text = Format(vJulio, "###,##0.00").ToString
    '                End If
    '                If vMes = "08" Then
    '                    vAgosto = drMdb1.GetValue(1).ToString
    '                    TxtAgosto.Text = Format(vAgosto, "###,##0.00").ToString
    '                End If
    '                If vMes = "09" Then
    '                    vSeptiembre = drMdb1.GetValue(1).ToString
    '                    TxtSeptiembre.Text = Format(vSeptiembre, "###,##0.00").ToString
    '                End If
    '                If vMes = "10" Then
    '                    vOctubre = drMdb1.GetValue(1).ToString
    '                    TxtOctubre.Text = Format(vOctubre, "###,##0.00").ToString
    '                End If
    '                If vMes = "11" Then
    '                    vNoviembre = drMdb1.GetValue(1).ToString
    '                    TxtNoviembre.Text = Format(vNoviembre, "###,##0.00").ToString
    '                End If
    '                If vMes = "12" Then
    '                    vDiciembre = drMdb1.GetValue(1).ToString
    '                    TxtDiciembre.Text = Format(vDiciembre, "###,##0.00").ToString
    '                End If
    '            End While
    '            TxtAnual.Text = Format(vSaldoAnualPresupuesto, "###,##0.00").ToString
    '            TxtAnual.SelectAll()
    '        Else
    '            'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
    '        End If
    '        drMdb1.Close()
    '    Catch ex As Exception
    '        MsgBox("Error al llenar los TextBox con los datos guardados en presupuesto")
    '        MsgBox(ex.ToString)
    '    End Try
    '    Return vSaldoAnualPresupuesto
    'End Function

    Private Sub CmbConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Aseguramos que haya un concepto seleccionado
        If CmbConcepto.SelectedIndex = -1 Then Exit Sub
        vConcepto = CmbConcepto.Text.ToString().Trim()

        ' 1. Buscamos los datos estáticos del concepto (Tipo y Descripción) usando un comando limpio
        Dim sqlConcepto As String = "SELECT TipoCON, DescripcionCON FROM conceptos WHERE CodigoCON = ?"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using cmd As New OleDbCommand(sqlConcepto, conexion)
                cmd.Parameters.AddWithValue("@cod", vConcepto)
                Try
                    conexion.Open()
                    Using dr As OleDbDataReader = cmd.ExecuteReader()
                        If dr.Read() Then
                            TxtTipoConcepto.Text = dr("TipoCON").ToString()
                            TxtDescripcion.Text = dr("DescripcionCON").ToString()
                        End If
                    End Using
                Catch ex As Exception
                    MsgBox("Error al buscar detalles del concepto: " & ex.Message)
                End Try
            End Using
        End Using

        ' 2. Enfocamos la caja correspondiente según la selección
        If RdbAnual.Checked = True Then
            TxtAnual.Select()
            TxtAnual.SelectAll()
        Else
            TxtEnero.Select()
            TxtEnero.SelectAll()
        End If

        ' 3. Rellenamos las 12 cajas mensuales con lo que haya en los presupuestos
        LlenarTextBox()
    End Sub

    Public Sub LlenarTextBox()
        ' 1. Ponemos todas las cajas a cero por defecto
        TxtEnero.Text = "0,00" : TxtFebrero.Text = "0,00" : TxtMarzo.Text = "0,00"
        TxtAbril.Text = "0,00" : TxtMayo.Text = "0,00" : TxtJunio.Text = "0,00"
        TxtJulio.Text = "0,00" : TxtAgosto.Text = "0,00" : TxtSeptiembre.Text = "0,00"
        TxtOctubre.Text = "0,00" : TxtNoviembre.Text = "0,00" : TxtDiciembre.Text = "0,00"
        TxtAnual.Text = "0,00"

        ' Array local para almacenar y comparar los 12 meses en memoria
        Dim importesMensuales(11) As Double
        Dim sumaAnual As Double = 0

        ' 2. Consulta SQL sobre tu estructura MDB actual
        vtipoSql = "SELECT ImportePRE, FDesdePRE FROM presupuesto WHERE EjercicioPRE = ? AND ConceptoPRE = ?"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using cmd As New OleDbCommand(vtipoSql, conexion)
                cmd.Parameters.AddWithValue("@eje", CInt(vAñoEjercicio))
                cmd.Parameters.AddWithValue("@con", vConcepto)

                Try
                    conexion.Open()
                    Using dr As OleDbDataReader = cmd.ExecuteReader()

                        ' Recorremos los registros que existan en tu MDB para este presupuesto
                        While dr.Read()
                            Dim importe As Double = Convert.ToDouble(dr("ImportePRE"))
                            Dim fecha As Date = Convert.ToDateTime(dr("FDesdePRE"))
                            Dim mes As Integer = fecha.Month ' Extrae el número de mes (1 al 12)

                            sumaAnual += importe

                            ' Guardamos en el array (índice 0 a 11)
                            If mes >= 1 AndAlso mes <= 12 Then
                                importesMensuales(mes - 1) = importe
                            End If

                            ' Asignamos el valor formateado a la caja de texto correspondiente
                            Select Case mes
                                Case 1 : TxtEnero.Text = Format(importe, "###,##0.00")
                                Case 2 : TxtFebrero.Text = Format(importe, "###,##0.00")
                                Case 3 : TxtMarzo.Text = Format(importe, "###,##0.00")
                                Case 4 : TxtAbril.Text = Format(importe, "###,##0.00")
                                Case 5 : TxtMayo.Text = Format(importe, "###,##0.00")
                                Case 6 : TxtJunio.Text = Format(importe, "###,##0.00")
                                Case 7 : TxtJulio.Text = Format(importe, "###,##0.00")
                                Case 8 : TxtAgosto.Text = Format(importe, "###,##0.00")
                                Case 9 : TxtSeptiembre.Text = Format(importe, "###,##0.00")
                                Case 10 : TxtOctubre.Text = Format(importe, "###,##0.00")
                                Case 11 : TxtNoviembre.Text = Format(importe, "###,##0.00")
                                Case 12 : TxtDiciembre.Text = Format(importe, "###,##0.00")
                            End Select
                        End While

                        ' Mostramos la suma total acumulada en la casilla Anual
                        TxtAnual.Text = Format(sumaAnual, "###,##0.00")

                        ' 3. DETECTAR AUTOMÁTICAMENTE SI ERA REPARTO ANUAL O MENSUAL
                        Dim todosIguales As Boolean = True
                        Dim primerImporte As Double = importesMensuales(0)

                        For i As Integer = 1 To 11
                            ' Si un solo mes es diferente al primero, es un presupuesto mensual personalizado
                            If importesMensuales(i) <> primerImporte Then
                                todosIguales = False
                                Exit For
                            End If
                        Next

                        ' Desvinculamos temporalmente los eventos para que el cambio de RadioButton no limpie los TextBox
                        RemoveHandler RdbAnual.CheckedChanged, AddressOf RdbAnual_CheckedChanged
                        RemoveHandler RdbMensual.CheckedChanged, AddressOf RdbAnual_CheckedChanged

                        ' Si todos los meses son iguales y el presupuesto no está vacío, es Anual. Si no, Mensual.
                        If todosIguales AndAlso sumaAnual > 0 Then
                            RdbAnual.Checked = True
                            GBoxAnual.Enabled = True
                            GBoxMensual.Enabled = False
                        Else
                            RdbMensual.Checked = True
                            GBoxAnual.Enabled = False
                            GBoxMensual.Enabled = True
                        End If

                        ' Volvemos a activar los escuchadores de los RadioButtons
                        AddHandler RdbAnual.CheckedChanged, AddressOf RdbAnual_CheckedChanged
                        AddHandler RdbMensual.CheckedChanged, AddressOf RdbAnual_CheckedChanged

                    End Using
                Catch ex As Exception
                    MsgBox("Error al cargar importes presupuestados: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Validar que tengamos un concepto contable seleccionado
        Dim concepto As String = CmbConcepto.Text.Trim()
        If String.IsNullOrEmpty(concepto) Then
            MessageBox.Show("Por favor, seleccione un concepto contable.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        ' 2. Mapeamos los valores de las cajas de texto a un array numérico en memoria (0 = Ene, 11 = Dic)
        Dim importesMensuales(11) As Double

        If RdbAnual.Checked Then
            ' Si es anual, dividimos el total entre 12 y redondeamos de forma limpia
            Dim totalAnual As Double = 0
            Double.TryParse(TxtAnual.Text, totalAnual)
            Dim importeRepartido As Double = Math.Round(totalAnual / 12, 2)
            For i As Integer = 0 To 11
                importesMensuales(i) = importeRepartido
            Next
        Else
            ' Si es mensual, parseamos cada una de las 12 cajas de tu formulario
            Double.TryParse(TxtEnero.Text, importesMensuales(0))
            Double.TryParse(TxtFebrero.Text, importesMensuales(1))
            Double.TryParse(TxtMarzo.Text, importesMensuales(2))
            Double.TryParse(TxtAbril.Text, importesMensuales(3))
            Double.TryParse(TxtMayo.Text, importesMensuales(4))
            Double.TryParse(TxtJunio.Text, importesMensuales(5))
            Double.TryParse(TxtJulio.Text, importesMensuales(6))
            Double.TryParse(TxtAgosto.Text, importesMensuales(7))
            Double.TryParse(TxtSeptiembre.Text, importesMensuales(8))
            Double.TryParse(TxtOctubre.Text, importesMensuales(9))
            Double.TryParse(TxtNoviembre.Text, importesMensuales(10))
            Double.TryParse(TxtDiciembre.Text, importesMensuales(11))
        End If

        ' 3. GRABACIÓN SEGURA EN LA MDB ACTUAL DE LOS USUARIOS
        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Try
                conexion.Open()

                ' Abrimos una transacción para asegurar la operación en bloque
                Using transaccion As OleDbTransaction = conexion.BeginTransaction()

                    ' PLAN A: Limpiamos cualquier presupuesto anterior que tuviera este concepto en este año
                    Dim sqlDelete As String = "DELETE FROM presupuesto WHERE ConceptoPRE = ? AND EjercicioPRE = ?"
                    Using cmdDelete As New OleDbCommand(sqlDelete, conexion, transaccion)
                        cmdDelete.Parameters.AddWithValue("@con", concepto)
                        cmdDelete.Parameters.AddWithValue("@eje", CInt(vAñoEjercicio))
                        cmdDelete.ExecuteNonQuery()
                    End Using

                    ' PLAN B: Inserción masiva de las 12 mensualidades con tus nombres de columna reales
                    Dim sqlInsert As String = "INSERT INTO presupuesto (ConceptoPRE, ImportePRE, EjercicioPRE, FDesdePRE) VALUES (?, ?, ?, ?)"
                    Using cmdInsert As New OleDbCommand(sqlInsert, conexion, transaccion)

                        ' Declaramos parámetros tipados para evitar fallos de comillas y formatos de fecha de Access
                        cmdInsert.Parameters.Add("@con", OleDbType.VarWChar)
                        cmdInsert.Parameters.Add("@imp", OleDbType.Double)
                        cmdInsert.Parameters.Add("@eje", OleDbType.Integer)
                        cmdInsert.Parameters.Add("@fec", OleDbType.Date)

                        ' Ejecutamos el bucle para los 12 meses del año
                        For mes As Integer = 1 To 12
                            ' Generamos la fecha del primer día de cada mes (01/01/Año, 01/02/Año...)
                            Dim fechaMes As New Date(CInt(vAñoEjercicio), mes, 1)

                            cmdInsert.Parameters(0).Value = concepto
                            cmdInsert.Parameters(1).Value = importesMensuales(mes - 1)
                            cmdInsert.Parameters(2).Value = CInt(vAñoEjercicio)
                            cmdInsert.Parameters(3).Value = fechaMes

                            cmdInsert.ExecuteNonQuery()
                        Next
                    End Using

                    ' Si todo ha ido bien sin errores, consolidamos los cambios en el archivo físico .mdb
                    transaccion.Commit()

                    ' Opcional: Si quieres usar tu lógica de resManager para el mensaje de éxito
                    Dim msgExito As String = resManager.GetString("PresupuestoGuardado")
                    If String.IsNullOrEmpty(msgExito) Then msgExito = "Presupuesto guardado correctamente."
                    MessageBox.Show(msgExito, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Me.Close()
                End Using
            Catch ex As Exception
                MessageBox.Show("Error crítico al guardar en la base de datos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub


    'Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
    '    vConcepto = CmbConcepto.Text
    '    If TxtAnual.Text <> 0 Then
    '        ' Comprobamos que el Concepto si esta en presupuesto
    '        cmdMdb1cr.CommandText = "SELECT * FROM presupuesto"
    '        cmdMdb1cr.CommandText += " WHERE "
    '        cmdMdb1cr.CommandText += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
    '        cmdMdb1cr.CommandText += " And presupuesto.ConceptoPRE = '" & vConcepto & "' "
    '        Try
    '            drMdb1 = cmdMdb1cr.ExecuteReader()
    '            If drMdb1.HasRows Then
    '                respuesta = MsgBox("Hay datos guardados con este Concepto ¿Se actualizan valores?.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar presupuesto")
    '                If respuesta = vbYes Then
    '                    drMdb1.Close()
    '                    ' Eliminar Registros mismo Concepto Apunte
    '                    vtipoSql = "DELETE FROM presupuesto"
    '                    vtipoSql += " WHERE presupuesto.ConceptoPRE = '" & vConcepto & "' "
    '                    cmdMdb1cr.CommandText = vtipoSql
    '                    Try
    '                        cmdMdb1cr.ExecuteNonQuery()
    '                        MsgBox("Registros en presupuesto, Borrados !!!")
    '                    Catch ex As Exception
    '                        MsgBox("Error al Borrar el presupuesto")
    '                        MsgBox(ex.ToString)
    '                    End Try
    '                Else
    '                    MsgBox("No actualiza nada y sale")
    '                    Me.Close()
    '                End If
    '            Else
    '                MsgBox("No existen registros en " & vConcepto)
    '            End If
    '            drMdb1.Close()
    '        Catch ex As Exception
    '            MsgBox("Error al verificar que el Concepto existe en presupuesto")
    '            MsgBox(ex.ToString)
    '        End Try
    '        If TxtAnual.Text <> "0,00" Then
    '            For i = 1 To 12
    '                If i = 1 Then
    '                    vImporte = TxtEnero.Text
    '                    vFDesde = "01/01/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 2 Then
    '                    vImporte = TxtFebrero.Text
    '                    vFDesde = "01/02/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 3 Then
    '                    vImporte = TxtMarzo.Text
    '                    vFDesde = "01/03/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 4 Then
    '                    vImporte = TxtAbril.Text
    '                    vFDesde = "01/04/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 5 Then
    '                    vImporte = TxtMayo.Text
    '                    vFDesde = "01/05/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 6 Then
    '                    vImporte = TxtJunio.Text
    '                    vFDesde = "01/06/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 7 Then
    '                    vImporte = TxtJulio.Text
    '                    vFDesde = "01/07/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 8 Then
    '                    vImporte = TxtAgosto.Text
    '                    vFDesde = "01/08/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 9 Then
    '                    vImporte = TxtSeptiembre.Text
    '                    vFDesde = "01/09/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 10 Then
    '                    vImporte = TxtOctubre.Text
    '                    vFDesde = "01/10/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 11 Then
    '                    vImporte = TxtNoviembre.Text
    '                    vFDesde = "01/11/" & vAñoEjercicio.ToString
    '                End If
    '                If i = 12 Then
    '                    vImporte = TxtDiciembre.Text
    '                    vFDesde = "01/12/" & vAñoEjercicio.ToString
    '                End If
    '                vAñadirSql = "INSERT INTO presupuesto "
    '                vAñadirSql += "(ConceptoPRE, ImportePRE, FDesdePRE, EjercicioPRE) "
    '                vAñadirSql += "VALUES ('" & vConcepto & "','" & vImporte & "','" & vFDesde & "','" & vAñoEjercicio & "')"
    '                cmdMdb1cr.CommandText = vAñadirSql
    '                Try
    '                    cmdMdb1cr.ExecuteNonQuery()
    '                    If i = 12 Then
    '                        'MsgBox("Registro, Grabado Correctamente")
    '                    End If
    '                Catch ex As Exception
    '                    MsgBox("Error al Grabar el Presupuesto")
    '                    MsgBox(ex.ToString)
    '                End Try
    '            Next
    '        End If
    '        Me.Close()
    '    Else
    '        MsgBox("No hay Importes")
    '    End If
    'End Sub

    Private Sub TxtAnual_Validating(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles TxtAnual.Validating
        ' Solo actuamos si está seleccionada la opción de reparto Anual
        If RdbAnual.Checked Then
            Dim totalAnual As Double = 0
            ' Convertimos el texto a número de forma segura
            If Double.TryParse(TxtAnual.Text.Trim(), totalAnual) Then
                ' Dividimos entre 12 y redondeamos a 2 decimales
                Dim importeMensual As Double = Math.Round(totalAnual / 12, 2)
                Dim textoFormateado As String = Format(importeMensual, "###,##0.00")

                ' Rellenamos las 12 cajas mensuales visualmente
                TxtEnero.Text = textoFormateado : TxtFebrero.Text = textoFormateado
                TxtMarzo.Text = textoFormateado : TxtAbril.Text = textoFormateado
                TxtMayo.Text = textoFormateado : TxtJunio.Text = textoFormateado
                TxtJulio.Text = textoFormateado : TxtAgosto.Text = textoFormateado
                TxtSeptiembre.Text = textoFormateado : TxtOctubre.Text = textoFormateado
                TxtNoviembre.Text = textoFormateado : TxtDiciembre.Text = textoFormateado

                ' Reajustamos el total anual por si el redondeo de decimales varió un céntimo
                TxtAnual.Text = Format(importeMensual * 12, "###,##0.00")
            End If
        End If
    End Sub

    Private Sub CalcularSumaMensualidades()
        ' Solo actuamos si está seleccionada la opción de introducción Mensual
        If RdbMensual.Checked Then
            Dim suma As Double = 0
            Dim temp As Double = 0

            ' Sumamos el valor de cada caja de texto de forma segura
            If Double.TryParse(TxtEnero.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtFebrero.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtMarzo.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtAbril.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtMayo.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtJunio.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtJulio.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtAgosto.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtSeptiembre.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtOctubre.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtNoviembre.Text, temp) Then PointToSuma(suma, temp)
            If Double.TryParse(TxtDiciembre.Text, temp) Then PointToSuma(suma, temp)

            ' Mostramos el resultado totalizado en la caja anual
            TxtAnual.Text = Format(suma, "###,##0.00")
        End If
    End Sub

    ' Función auxiliar rápida para acumular los valores
    Private Sub PointToSuma(ByRef total As Double, valor As Double)
        total += valor
    End Sub

    ' Enlazamos las 12 cajas al mismo evento para ahorrar código
    Private Sub TxtMeses_Leave(sender As Object, e As EventArgs) Handles _
    TxtEnero.Leave, TxtFebrero.Leave, TxtMarzo.Leave, TxtAbril.Leave,
    TxtMayo.Leave, TxtJunio.Leave, TxtJulio.Leave, TxtAgosto.Leave,
    TxtSeptiembre.Leave, TxtOctubre.Leave, TxtNoviembre.Leave, TxtDiciembre.Leave

        Dim txt As TextBox = CType(sender, TextBox)
        Dim valor As Double = 0

        ' Damos formato de moneda a la caja en la que estábamos parados
        If Double.TryParse(txt.Text.Trim(), valor) Then
            txt.Text = Format(valor, "###,##0.00")
        Else
            txt.Text = "0,00"
        End If

        ' Recalculamos el total anual reflejado en la pantalla
        CalcularSumaMensualidades()
    End Sub

    Private Sub TxtMeses_Enter(sender As Object, e As EventArgs) Handles _
    TxtEnero.Enter, TxtFebrero.Enter, TxtMarzo.Enter, TxtAbril.Enter,
    TxtMayo.Enter, TxtJunio.Enter, TxtJulio.Enter, TxtAgosto.Enter,
    TxtSeptiembre.Enter, TxtOctubre.Enter, TxtNoviembre.Enter, TxtDiciembre.Enter

        Dim txt As TextBox = CType(sender, TextBox)
        Dim valor As Double = 0

        ' Al entrar, quitamos los puntos de millar para facilitar la escritura manual
        If Double.TryParse(txt.Text.Trim(), valor) Then
            If valor = 0 Then
                txt.Text = "" ' Si es cero, vaciamos la caja para que no tenga que borrar el "0,00"
            Else
                txt.Text = valor.ToString("F2") ' Formato limpio sin separador de miles (ej: 1250,00)
            End If
        End If
        txt.SelectAll()
    End Sub

    Private Sub RdbAnual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbAnual.CheckedChanged, RdbMensual.CheckedChanged
        ' 1. Habilitamos o deshabilitamos los contenedores visuales según el RadioButton activo
        GBoxAnual.Enabled = RdbAnual.Checked
        GBoxMensual.Enabled = RdbMensual.Checked

        ' 2. Lógica específica al activar cada opción
        If RdbAnual.Checked Then
            ' Si pasa a ANUAL, ponemos el foco en el Total Anual para que defina la nueva cifra macro
            TxtAnual.Select()
            TxtAnual.SelectAll()
        ElseIf RdbMensual.Checked Then
            ' Si pasa a MENSUAL, calculamos la suma de lo que ya haya en las cajas mensuales 
            ' para que el Total Anual refleje la realidad actual de los meses inmediatamente.
            CalcularSumaMensualidades()

            ' Llevamos el foco al primer mes del año para que empiece a editar cómodamente
            TxtEnero.Select()
            TxtEnero.SelectAll()
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

    'Private Sub RdbAnual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbAnual.CheckedChanged
    '    If RdbAnual.Checked = True Then
    '        GBoxAnual.Enabled = True
    '        GBoxMensual.Enabled = False
    '        TxtAnual.Select()
    '        TxtAnual.SelectAll()
    '    End If
    'End Sub

    'Private Sub RdbMensual_CheckedChanged(sender As Object, e As EventArgs) Handles RdbMensual.CheckedChanged
    '    If RdbMensual.Checked = True Then
    '        GBoxAnual.Enabled = False
    '        GBoxMensual.Enabled = True
    '        TxtEnero.Select()
    '        TxtEnero.SelectAll()
    '    End If
    'End Sub

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