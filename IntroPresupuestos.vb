Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class IntroPresupuestos

    Public vConcepto, vtipoSql, vAñadirSql, vFDesde, vBorrarPresu As String
    Public vMensual, vAnual, vEnero, vFebrero, vMarzo, vAbril, vMayo, vJunio, vSaldoAnualPresupuesto, vImporte As Double
    Public vJulio, vAgosto, vSeptiembre, vOctubre, vNoviembre, vDiciembre As Double
    Public TL(18) As ToolTip
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub IntroApuntes_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Optimización: Bucle dinámico para las etiquetas dentro del GroupBox de meses
        For i As Integer = 16 To 28
            Dim lbl() As Control = Me.Controls.Find("Label" & i, True)
            If lbl.Length > 0 Then lbl(0).Text = vMoneda
        Next

        Me.KeyPreview = True

        ' Inicialización centralizada de ToolTips
        Dim controlesToolTip As Control() = {
            BtnConcepto, BtnAceptar, BtnCancelar, CmbConcepto, TxtAnual,
            TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
            TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre,
            RdbAnual, RdbMensual
        }

        Dim clavesToolTip As String() = {
            "AñadeEditaBorraCC", "BtnAceptar.Text", "BtnCancelar.Text", "SeleccionarConcepto", "TotalAnual",
            "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual",
            "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual", "TxtMensual",
            "RadioAnual", "RadioMensual"
        }

        For i As Integer = 0 To TL.Length - 1
            TL(i) = New ToolTip()
            TL(i).SetToolTip(controlesToolTip(i), rmse.GetString(clavesToolTip(i)))
        Next

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
                    MsgBox(ex.Message)
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
        ' 1. Ponemos todas las cajas a cero por defecto usando un array para no repetir líneas
        Dim cajasMeses As TextBox() = {TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
                                       TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre}

        For Each txt In cajasMeses
            txt.Text = "0,00"
        Next
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

                            ' Guardamos en el array y asignamos a la caja correspondiente de forma automática
                            ' sin Select Case
                            If mes >= 1 AndAlso mes <= 12 Then
                                importesMensuales(mes - 1) = importe
                                cajasMeses(mes - 1).Text = Format(importe, "###,##0.00")
                            End If
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
                    MsgBox(ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Validar que tengamos un concepto contable seleccionado
        Dim concepto As String = CmbConcepto.Text.Trim()
        If String.IsNullOrEmpty(concepto) Then
            MessageBox.Show(rmse.GetString("SeleccionarCC"), resManager.GetString("Aviso"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
                    Dim msgExito As String = rmse.GetString("PresupuestoGuardado")
                    MessageBox.Show(msgExito, resManager.GetString("Éxito"), MessageBoxButtons.OK, MessageBoxIcon.Information)

                    Me.Close()
                End Using
            Catch ex As Exception
                MessageBox.Show(ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

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
        frmPrincipal.ConceptosContablesToolStripMenuItem.PerformClick()
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
            ' 1. Convertir el texto a número de forma segura (eliminando el formateo si ya existiera)
            Dim importeAnualPure As Double = 0
            Double.TryParse(TxtAnual.Text, importeAnualPure)

            ' Guardamos en tu variable global y formateamos la caja anual
            vAnual = importeAnualPure
            TxtAnual.Text = Format(vAnual, "###,##0.00")

            ' 2. Calcular el reparto mensual exacto
            vMensual = vAnual / 12
            Dim textoMensualFormateado As String = Format(vMensual, "###,##0.00")

            ' 3. Asignar el valor a tus 12 variables globales de meses
            vEnero = vMensual : vFebrero = vMensual : vMarzo = vMensual : vAbril = vMensual
            vMayo = vMensual : vJunio = vMensual : vJulio = vMensual : vAgosto = vMensual
            vSeptiembre = vMensual : vOctubre = vMensual : vNoviembre = vMensual : vDiciembre = vMensual

            ' 4. Rellenar las 12 cajas de la interfaz usando un bucle limpio
            Dim cajasMeses As TextBox() = {TxtEnero, TxtFebrero, TxtMarzo, TxtAbril, TxtMayo, TxtJunio,
                                       TxtJulio, TxtAgosto, TxtSeptiembre, TxtOctubre, TxtNoviembre, TxtDiciembre}

            For Each txt In cajasMeses
                txt.Text = textoMensualFormateado
            Next

            RdbMensual.Select()
        End If

        ' Reemplazo del punto por la coma para el teclado numérico
        If e.KeyChar = "."c Then
            e.KeyChar = ","c
        End If
    End Sub

    Private Sub TxtMeses_KeyPress(sender As Object, e As KeyPressEventArgs) Handles _
    TxtEnero.KeyPress, TxtFebrero.KeyPress, TxtMarzo.KeyPress, TxtAbril.KeyPress,
    TxtMayo.KeyPress, TxtJunio.KeyPress, TxtJulio.KeyPress, TxtAgosto.KeyPress,
    TxtSeptiembre.KeyPress, TxtOctubre.KeyPress, TxtNoviembre.KeyPress, TxtDiciembre.KeyPress

        SoloNumerosConPunto(e)

        ' Reemplazo del punto por la coma para el teclado numérico
        If e.KeyChar = "."c Then
            e.KeyChar = ","c
        End If

        If e.KeyChar = ChrW(Keys.Enter) Then
            Dim txt As TextBox = CType(sender, TextBox)
            Dim valorIngresado As Double = 0
            Double.TryParse(txt.Text, valorIngresado)

            ' Formateamos la caja actual inmediatamente
            txt.Text = Format(valorIngresado, "###,##0.00")

            ' Controlamos el foco siguiente y guardamos en la variable global correcta según la caja pulsada
            Select Case txt.Name
                Case "TxtEnero" : vEnero = valorIngresado : TxtFebrero.Select()
                Case "TxtFebrero" : vFebrero = valorIngresado : TxtMarzo.Select()
                Case "TxtMarzo" : vMarzo = valorIngresado : TxtAbril.Select()
                Case "TxtAbril" : vAbril = valorIngresado : TxtMayo.Select()
                Case "TxtMayo" : vMayo = valorIngresado : TxtJunio.Select()
                Case "TxtJunio" : vJunio = valorIngresado : TxtJulio.Select()
                Case "TxtJulio" : vJulio = valorIngresado : TxtAgosto.Select()
                Case "TxtAgosto" : vAgosto = valorIngresado : TxtSeptiembre.Select()
                Case "TxtSeptiembre" : vSeptiembre = valorIngresado : TxtOctubre.Select()
                Case "TxtOctubre" : vOctubre = valorIngresado : TxtNoviembre.Select()
                Case "TxtNoviembre" : vNoviembre = valorIngresado : TxtDiciembre.Select()
                Case "TxtDiciembre" : vDiciembre = valorIngresado : BtnAceptar.Select()
            End Select

            ' Calculamos la suma total usando tus variables globales actualizadas
            vAnual = vEnero + vFebrero + vMarzo + vAbril + vMayo + vJunio + vJulio + vAgosto + vSeptiembre + vOctubre + vNoviembre + vDiciembre
            TxtAnual.Text = Format(vAnual, "###,##0.00")

            ' Evitamos el pitido molesto de Windows al pulsar Enter
            e.Handled = True
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class