Imports System.Data
Imports System.Windows.Forms

Public Class NuevoTipoCuentaBancaria

    Private dtTiposMemoria As New DataTable()
    Public vtipoSql, vtipoGrid, vTxtNombre, vTxtDescripcion As String
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub NuevaCuentaBancaria_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        Dim TL(3) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnAceptar, resManager.GetString("ToolTipAceptar"))
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnCancelar, resManager.GetString("ToolTipCancelar"))
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.TxtDescripcion, resManager.GetString("Descripcion"))
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.TxtNombre, resManager.GetString("Nombre"))
        TxtNombre.Select()

        Try
            ' 1. Configuramos las columnas de la tabla en memoria si no existen
            If dtTiposMemoria.Columns.Count = 0 Then
                dtTiposMemoria.Columns.Add("CodigoTIP", GetType(String))
                dtTiposMemoria.Columns.Add("TipoTraducido", GetType(String))
            End If

            ' 2. Limpiamos cualquier dato residual de forma segura
            dtTiposMemoria.Clear()

            ' 3. Cargamos los códigos estables desde tu base de datos Access
            Dim sqlCarga As String = "SELECT CodigoTIP FROM tipocuentas ORDER BY CodigoTIP"
            cmdMdb1cr.CommandText = sqlCarga
            cmdMdb1cr.Parameters.Clear()

            Using reader As OleDb.OleDbDataReader = cmdMdb1cr.ExecuteReader()
                While reader.Read()
                    Dim codigoTIP As String = reader("CodigoTIP").ToString().Trim()
                    Dim textoTraducido As String = codigoTIP.Replace("_", " ").ToUpper()

                    ' Buscamos la traducción en tu archivo de recursos local (rmse)
                    If rmse IsNot Nothing Then
                        Dim claveRecurso As String = codigoTIP.Replace(" ", "_")
                        Dim trad As String = rmse.GetString(claveRecurso)
                        If Not String.IsNullOrEmpty(trad) Then
                            textoTraducido = trad.Trim().ToUpper()
                        End If
                    End If

                    ' Guardamos el registro traducido en nuestra tabla local
                    dtTiposMemoria.Rows.Add(codigoTIP, textoTraducido)
                End While
            End Using
        Catch ex As Exception
            ' Previene caídas en el arranque
        End Try
    End Sub

    Private Sub TxtNombre_TextChanged(sender As Object, e As EventArgs) Handles TxtNombre.TextChanged
        ' 1. Controlar las mayúsculas de forma visual sin perder la posición del cursor
        Dim posicionCursor As Integer = TxtNombre.SelectionStart
        Dim textoMayusculas As String = TxtNombre.Text.ToUpper()

        If TxtNombre.Text <> textoMayusculas Then
            TxtNombre.Text = textoMayusculas
            TxtNombre.SelectionStart = posicionCursor
        End If

        Dim vBusca As String = TxtNombre.Text.Trim()

        ' Si el cuadro está vacío, ocultamos el mini-grid para que no estorbe
        If vBusca = "" Then
            DgvExistente.Visible = False
            Exit Sub
        End If

        ' 2. BUSCA EN LA BD (Como ahora todo es Mayúsculas, el LIKE funciona impecable)
        vtipoSql = "SELECT CodigoTIP FROM tipocuentas WHERE CodigoTIP Like '" & vBusca.Replace("'", "''") & "%' ORDER BY CodigoTIP"
        vtipoGrid = "NOMBRESEXISTENTES3"

        ' Llena tu mini-grid nativo
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' =========================================================================
        ' ✨ TRUCO MULTIIDIOMA SANEADO: Traducción Visual Limpia sin Guiones
        ' =========================================================================
        If DgvExistente.Rows.Count > 0 Then
            DgvExistente.Visible = True

            For Each row As DataGridViewRow In DgvExistente.Rows
                If row.IsNewRow Then Continue For

                ' Leemos el código base en mayúsculas (Ej: "CUENTA_CORRIENTE")
                Dim codigoTIP As String = row.Cells(0).Value?.ToString().Trim()

                If Not String.IsNullOrEmpty(codigoTIP) Then
                    ' 🚀 REPARADO 1: Buscamos en el resManager local reemplazando espacios por guiones
                    Dim claveRecurso As String = codigoTIP.Replace(" ", "_")
                    Dim traduccion As String = ""

                    If resManager IsNot Nothing Then
                        traduccion = resManager.GetString(claveRecurso)
                    End If

                    ' 🚀 REPARADO 2: Si tiene traducción, limpiamos los guiones bajos visuales.
                    ' Si es un tipo nuevo del usuario y no tiene traducción, le quitamos los guiones también al original.
                    If Not String.IsNullOrEmpty(traduccion) Then
                        row.Cells(0).Value = traduccion.Replace("_", " ").Trim().ToUpper()
                    Else
                        row.Cells(0).Value = codigoTIP.Replace("_", " ").Trim().ToUpper()
                    End If
                End If
            Next
        Else
            DgvExistente.Visible = False
        End If
    End Sub

    Private Sub TxtNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtNombre.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            DgvExistente.Visible = False
            TxtDescripcion.Select()
        End If
    End Sub

    Private Sub TxtNombre_LostFocus(sender As Object, e As EventArgs) Handles TxtNombre.LostFocus
        DgvExistente.Visible = False
        TxtDescripcion.Select()
    End Sub

    Private Sub TxtDescripcion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TxtDescripcion.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Guardamos el texto forzando MAYÚSCULAS desde el teclado
        Dim nombreMayusculas As String = TxtNombre.Text.Trim().ToUpper()

        If nombreMayusculas <> "" Then
            vTxtNombre = nombreMayusculas
            vTxtDescripcion = TxtDescripcion.Text.Trim()

            ' =========================================================================
            ' 2. CONTROL DE DUPLICADOS EN MAYÚSCULAS (Inmune a fallos de OleDb)
            ' =========================================================================
            vtipoSql = "SELECT COUNT(*) FROM tipocuentas WHERE CodigoTIP = ?"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("?", vTxtNombre) ' Va en mayúsculas puras

            Try
                Dim existe As Integer = Convert.ToInt32(cmdMdb1cr.ExecuteScalar())

                If existe > 0 Then
                    ' Si encuentra coincidencia exacta (ej: CASH contra CASH), bloquea
                    Dim msgExiste As String = resManager.GetString("Nombre") & ": " & TxtNombre.Text.Trim() & ", " & resManager.GetString("Existe") & " " & frmTipoCuentaBancaria.rmse.GetString("$this.Text")
                    MessageBox.Show(msgExiste, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    TxtNombre.Select()
                    TxtNombre.SelectAll()
                    Exit Sub
                End If
            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorDuplicados") & ": " & vbNewLine & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try

            ' =========================================================================
            ' 3. CALCULAR EL SIGUIENTE ID DISPONIBLE (MAX + 1)
            ' =========================================================================
            Dim siguienteID As Integer = 1
            Try
                cmdMdb1cr.CommandText = "SELECT MAX(IdTipoCUE) FROM tipocuentas"
                cmdMdb1cr.Parameters.Clear()
                Dim resultado As Object = cmdMdb1cr.ExecuteScalar()
                If resultado IsNot DBNull.Value AndAlso resultado IsNot Nothing Then
                    siguienteID = Convert.ToInt32(resultado) + 1
                End If
            Catch ex As Exception
                siguienteID = 1
            End Try

            ' =========================================================================
            ' 4. INSERCIÓN DIRECTA PARAMETRIZADA
            ' =========================================================================
            vtipoSql = "INSERT INTO tipocuentas (IdTipoCUE, CodigoTIP, DescripcionTIP) VALUES (?, ?, ?)"
            cmdMdb1cr.CommandText = vtipoSql
            cmdMdb1cr.Parameters.Clear()
            cmdMdb1cr.Parameters.AddWithValue("@IdTipoCUE", siguienteID)
            cmdMdb1cr.Parameters.AddWithValue("@CodigoTIP", vTxtNombre) ' Se guarda en MAYÚSCULAS puras
            cmdMdb1cr.Parameters.AddWithValue("@DescripcionTIP", vTxtDescripcion)

            Try
                cmdMdb1cr.ExecuteNonQuery()
                Me.Close() ' Guardado impecable
            Catch ex As Exception
                MessageBox.Show(resManager.GetString("ErrorGrabarRegistro") & ": " & vbNewLine & ex.Message, resManager.GetString("Error"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Else
            Dim msgVacio As String = resManager.GetString("NoHayDatos") & ": " & resManager.GetString("Nombre")
            MessageBox.Show(msgVacio, rmse.GetString("$this.Text"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            TxtNombre.Select()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
    End Sub

    Private Sub TxtNombre_GotFocus(sender As Object, e As EventArgs) Handles TxtNombre.GotFocus
        PintaTxt()
    End Sub

    Public Sub PintaTxt()
        Dim Texto As TextBox
        Texto = Me.ActiveControl
        Texto.SelectionStart = 0
        Texto.SelectionLength = Texto.Text.Length
    End Sub

End Class