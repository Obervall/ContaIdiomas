Imports System.Windows.Forms

Public Class SeleccionEjercicio

    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub SeleccionEjercicio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        ' Llenar el Combo Ejercicio
        '**************************
        cmdMdb1cr.CommandText = "SELECT * FROM ejercicios ORDER BY ejercicios.EjercicioEJE ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                vCantAños = -1
                While drMdb1.Read()
                    vCantAños += 1
                    CmbEjercicio.Items.Add(drMdb1.GetValue(0))
                End While
                CmbEjercicio.Text = Str(vAñoActual).ToString
                vAñoEjercicio = vAñoActual
            Else
                MsgBox(rmse.GetString("MsgNoExistenReg"))
            End If
            drMdb1.Close()
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
        ' Ejecutamos la función y guardamos el resultado
        Dim procesoExitoso As Boolean = IniciarSaldosIniciales(vAñoEjercicio)

        If Not procesoExitoso Then
            ' Si devolvió False, hacemos EXACTAMENTE lo que hacía tu código original:
            ' Mostramos el aviso en el idioma del usuario y CERRAMOS la ventana
            Dim msgSinDatos As String = resManager.GetString("NoHayDatosHistoricos")
            If String.IsNullOrEmpty(msgSinDatos) Then msgSinDatos = resManager.GetString("NoHayDatosHistoricos")

            MessageBox.Show(msgSinDatos, resManager.GetString("Aviso"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
            Exit Sub
        Else
            ' Si devolvió True, mostramos el mensaje de éxito en su idioma
            Dim msgExito As String = resManager.GetString("SaldosGeneradosExito")
            If String.IsNullOrEmpty(msgExito) Then msgExito = resManager.GetString("SaldosGeneradosExito")

            MessageBox.Show(msgExito, resManager.GetString("Exito"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.Close()
        End If
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
            cmdMdb1cr.CommandText = "SELECT * FROM ejercicios"
            cmdMdb1cr.CommandText += " WHERE "
            cmdMdb1cr.CommandText += "ejercicios.EjercicioEJE = " & vAñoNuevo.ToString
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                If drMdb1.HasRows Then
                    MsgBox(rmse.GetString("MsgExiste") & vAñoNuevo)
                Else
                    MsgBox(rmse.GetString("MsgNoExiste") & vAñoNuevo)
                    drMdb1.Close()
                    tipoSql = "INSERT INTO ejercicios "
                    tipoSql += "(EjercicioEJE) "
                    tipoSql += "VALUES ('" & vAñoNuevo & "')"
                    cmdMdb1cr.CommandText = tipoSql
                    Try
                        cmdMdb1cr.ExecuteNonQuery()
                        'MsgBox("Registro, Grabado Correctamente")
                    Catch ex As Exception
                        MsgBox(rmse.GetString("MsgErrorAñadirEjercicio") & vAñoNuevo.ToString & vbCrLf & ex.ToString)
                    End Try
                End If
                drMdb1.Close()
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