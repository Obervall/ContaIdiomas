Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class SeleccionEjercicio

    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub SeleccionEjercicio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        ' Llenar el Combo Ejercicio
        '**************************
        cmdMySql1cr.CommandText = "SELECT * FROM ejercicios ORDER BY ejercicios.EjercicioEJE ASC"
        Try
            drMySql1 = cmdMySql1cr.ExecuteReader()
            If drMySql1.HasRows Then
                vCantAños = -1
                While drMySql1.Read()
                    vCantAños += 1
                    CmbEjercicio.Items.Add(drMySql1.GetValue(0))
                End While
                CmbEjercicio.Text = Str(vAñoActual).ToString
                vAñoEjercicio = vAñoActual
            Else
                MsgBox(rmse.GetString("MsgNoExistenReg"))
            End If
            drMySql1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub CmbEjercicio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbEjercicio.SelectedIndexChanged
        vAñoEjercicio = CmbEjercicio.Text
        Text = Label1.Text & " - " & vAñoEjercicio.ToString
        ' 2. Aplicar a todos los formularios abiertos, o sea a Principal
        For Each f As Form In Application.OpenForms
            '' 3. Refrescar el formulario Principal
            If TypeOf f Is Principal Then
                ActualizarTextosFormulario(f)
            End If
        Next
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        'Quitamos el Concepto SALDO del Ejercicio marcado
        vConceptoAPU = "SALDO"
        vtipoSql = "DELETE FROM apuntes"
        vtipoSql += " WHERE apuntes.ConceptoAPU = '" & vConceptoAPU & "' "
        cmdMySql1cr.CommandText = vtipoSql
        Try
            cmdMySql1cr.ExecuteNonQuery()
            'MsgBox("Registros SALDO, Borrados !!!")
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        ' 1. Consulta SQL: Agrupamos y sumamos SOLO los años ESTRICTAMENTE MENORES al seleccionado
        Dim consulta As String =
        "SELECT A.EjercicioAPU, A.CuentaAPU, SUM(A.ImporteAPU) AS SumaAño " &
        "FROM (Ejercicios AS E INNER JOIN Apuntes AS A ON E.EjercicioEJE = A.EjercicioAPU) " &
        "WHERE E.EjercicioEJE < ? " &
        "GROUP BY A.EjercicioAPU, A.CuentaAPU " &
        "ORDER BY A.EjercicioAPU ASC"

        Dim dtMovimientos As New DataTable()

        ' 2. Carga de datos históricos en memoria
        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using comando As New OleDbCommand(consulta, conexion)
                comando.Parameters.AddWithValue("@AñoSeleccionado", vAñoEjercicio)
                Using adaptador As New OleDbDataAdapter(comando)
                    Try
                        conexion.Open()
                        adaptador.Fill(dtMovimientos)
                        MsgBox("Datos históricos cargados en memoria: " & dtMovimientos.Rows.Count.ToString & " registros encontrados.")
                    Catch ex As Exception
                        MessageBox.Show("Error al leer históricos: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End Try
                End Using
            End Using
        End Using

        ' 3. Procesamiento en memoria para obtener el saldo final acumulado por cuenta
        Dim saldosAcumulados As New Dictionary(Of String, Decimal)()

        For Each fila As DataRow In dtMovimientos.Rows
            Dim cuenta As String = fila("CuentaAPU").ToString()
            Dim importeAño As Decimal = Convert.ToDecimal(fila("SumaAño"))

            If saldosAcumulados.ContainsKey(cuenta) Then
                saldosAcumulados(cuenta) += importeAño
            Else
                saldosAcumulados.Add(cuenta, importeAño)
            End If
        Next

        ' Si no hay saldos que arrastrar del pasado, salimos del proceso de inserción
        If saldosAcumulados.Count = 0 Then
            MessageBox.Show("No hay datos históricos en años anteriores para generar saldos iniciales.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        End If

        ' 4. Inserción de los Saldos Iniciales en la tabla Apuntes
        ' Definimos la fecha de inserción fija: 01/01/vAñoEjercicio
        Dim fechaSaldoInicial As New Date(vAñoEjercicio, 1, 1)

        Dim consultaInsert As String =
        "INSERT INTO Apuntes (FechaAPU, ConceptoAPU, DescripcionAPU, ImporteAPU, EjercicioAPU, CuentaAPU) " &
        "VALUES (?, ?, ?, ?, ?, ?)"

        Using conexion As New OleDbConnection(conexion1.ConnectionString)
            Using comandoInsert As New OleDbCommand(consultaInsert, conexion)

                ' Declaramos los parámetros en el orden exacto de los signos de interrogación '?'
                comandoInsert.Parameters.Add("@Fecha", OleDbType.Date)
                comandoInsert.Parameters.Add("@Concepto", OleDbType.VarWChar)
                comandoInsert.Parameters.Add("@Descripcion", OleDbType.VarWChar)
                comandoInsert.Parameters.Add("@Importe", OleDbType.Currency)
                comandoInsert.Parameters.Add("@Ejercicio", OleDbType.Integer)
                comandoInsert.Parameters.Add("@Cuenta", OleDbType.VarWChar)

                Try
                    conexion.Open()
                    ' Usamos una transacción para asegurar que se guarden todos los saldos o ninguno
                    Using transaccion As OleDbTransaction = conexion.BeginTransaction()
                        comandoInsert.Transaction = transaccion

                        ' Recorremos cada cuenta calculada e inyectamos su saldo acumulado
                        For Each par In saldosAcumulados
                            Dim cuenta As String = par.Key
                            Dim saldoFinalPasado As Decimal = par.Value

                            ' Omitimos cuentas con saldo acumulado cero si no deseas registros vacíos
                            If saldoFinalPasado <> 0 Then
                                comandoInsert.Parameters("@Fecha").Value = fechaSaldoInicial
                                comandoInsert.Parameters("@Concepto").Value = "SALDO"
                                comandoInsert.Parameters("@Descripcion").Value = "Saldo Inicial"
                                comandoInsert.Parameters("@Importe").Value = saldoFinalPasado
                                comandoInsert.Parameters("@Ejercicio").Value = vAñoEjercicio
                                comandoInsert.Parameters("@Cuenta").Value = cuenta ' Guarda el identificador/nombre de la cuenta
                                comandoInsert.ExecuteNonQuery()
                            End If
                        Next

                        ' Confirmamos los cambios en la base de datos
                        transaccion.Commit()
                        MessageBox.Show("Saldos iniciales del año " & vAñoEjercicio & " generados e insertados correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error al insertar los saldos iniciales: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using

        ' Aplicar a todos los formularios abiertos, o sea a Principal
        For Each f As Form In Application.OpenForms
            ' Refrescar el formulario Principal
            If TypeOf f Is Principal Then
                ActualizarTextosFormulario(f)
            End If
        Next
        Me.Close()
    End Sub

    Private Sub BtnCrearNuevo_Click(sender As Object, e As EventArgs) Handles BtnCrearNuevo.Click
        Dim message, title, myValue, vAñoNuevo, tipoSql As String
        Dim yearValue As Integer
        Dim esValido As Boolean = False
        ' Set prompt.
        message = rmse.GetString("MsgIntroEjercicio")
        ' Set title.
        title = rmse.GetString("MsgTituloEjercicio")
        Do
            ' Display message, title, and default value.
            myValue = InputBox(message, title, Val(vAñoActual + 1))
            ' 1. Si el usuario cancela o no escribe nada, salimos del bucle inmediatamente
            If String.IsNullOrEmpty(myValue) Then Exit Sub
            ' 2. Intentamos convertir a número y validar el rango
            If Integer.TryParse(myValue, yearValue) Then
                If yearValue >= 1900 AndAlso yearValue < 3000 Then
                    esValido = True ' El año es correcto, se marcará para salir
                Else
                    MsgBox(rmse.GetString("MsgAvisoAño"), MsgBoxStyle.Exclamation)
                End If
            Else
                MsgBox(rmse.GetString("MsgAvisoAño"), MsgBoxStyle.Exclamation)
            End If
        Loop While Not esValido ' Se repite MIENTRAS NO sea válido
        ' Código posterior con la variable 'yearValue' ya validada
        If yearValue = vAñoActual Then
            vAñoEjercicio = vAñoActual
            MsgBox(rmse.GetString("MsgEjercicioAño"))
        Else
            ' Añadimos un año si no está en la tabla
            vAñoNuevo = yearValue
            cmdMySql1cr.CommandText = "SELECT * FROM ejercicios"
            cmdMySql1cr.CommandText += " WHERE "
            cmdMySql1cr.CommandText += "ejercicios.EjercicioEJE = " & vAñoNuevo.ToString
            Try
                drMySql1 = cmdMySql1cr.ExecuteReader()
                If drMySql1.HasRows Then
                    MsgBox(rmse.GetString("MsgExiste") & vAñoNuevo)
                Else
                    MsgBox(rmse.GetString("MsgNoExiste") & vAñoNuevo)
                    drMySql1.Close()
                    tipoSql = "INSERT INTO ejercicios "
                    tipoSql += "(EjercicioEJE) "
                    tipoSql += "VALUES ('" & vAñoNuevo & "')"
                    cmdMySql1cr.CommandText = tipoSql
                    Try
                        cmdMySql1cr.ExecuteNonQuery()
                        'MsgBox("Registro, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox(rmse.GetString("MsgErrorAñadirEjercicio") & vAñoNuevo.ToString & vbCrLf & ex.ToString)
                    End Try
                End If
                drMySql1.Close()
            Catch ex As Exception
                MsgBox(rmse.GetString("MsgErrorBuscarEjercicio") & vAñoNuevo.ToString & vbCrLf & ex.ToString)
            End Try
        End If
        Text = rmse.GetString("Label1.Text") & " - " & vAñoEjercicio.ToString
        ' 2. Aplicar a todos los formularios abiertos, o sea a Principal
        For Each f As Form In Application.OpenForms
            '' 3. Refrescar el formulario Principal
            If TypeOf f Is Principal Then
                ActualizarTextosFormulario(f)
            End If
        Next
        Me.Close()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class