Imports System.Collections.Generic
Imports System.Data.OleDb

Public Class SeleccionDatosIngresos

    Public i As Integer
    Public vConcepto As String
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub SeleccionDatosIngresos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cmdMdb1cr.CommandText = "SELECT * FROM conceptos "
        cmdMdb1cr.CommandText += "Where conceptos.TipoCON = 'INGRESO' ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    ListBox1.Items.Add(drMdb1.GetValue(0))
                End While
                ListBox1.Text = ListBox1.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnTodos_Click(sender As Object, e As EventArgs) Handles BtnTodos.Click
        If ListBox1.SelectedItems.Count = ListBox1.Items.Count Then
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, False)
            Next
            ' 2. Seleccionamos el primero UNA sola vez fuera del bucle
            If ListBox1.Items.Count > 0 Then ListBox1.SetSelected(0, True)
            BtnTodos.Text = rmse.GetString("BtnTodos.Text")
        Else
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, True)
            Next
            BtnTodos.Text = rmse.GetString("MsgDeseleccionar")
        End If
    End Sub

    Private Sub BtnContinuar_Click(sender As Object, e As EventArgs) Handles BtnContinuar.Click
        ' 1. SANEAMIENTO PREVENTIVO: Limpiamos la memoria de consultas previas
        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 SENTENCIA RELACIONAL INTEGRAL NUMÉRICA PURA INDESTRUCTIBLE
        ' =========================================================================
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU " &
                   "FROM apuntes " &
                   "WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString

        If ListBox1.SelectedItems.Count <> 0 Then
            ' Saco dinámico en la RAM para juntar los IDs numéricos enteros puros
            Dim listaIDs As New List(Of Integer)()

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim textoSeleccionado As String = ListBox1.SelectedItems(i).ToString().Trim()

                ' 🚀 EL CORTAFUEGOS DEFINITIVO DE ESPACIOS Y GUIONES:
                ' Fabricamos todas las variantes posibles para que Access no tenga escapatoria
                Dim textoConGuion As String = textoSeleccionado.Replace(" ", "_").ToUpper()
                Dim textoSinGuion As String = textoSeleccionado.Replace("_", " ").ToUpper()

                ' Vamos a buscar el ID numérico entero real correspondiente a ese concepto en Access
                Dim idEncontrado As Integer = 0
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    ' Buscamos de forma elástica en todas las columnas posibles de tu maestro de conceptos
                    Dim sqlBuscarID As String = "SELECT IdConceptoCON FROM conceptos WHERE " &
                                                  "UCASE(CodigoCON) = ? OR UCASE(CodigoCON) = ? OR " &
                                                  "UCASE(DescripcionCON) = ? OR UCASE(DescripcionCON) = ?"

                    Using cmd As New OleDbCommand(sqlBuscarID, con)
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add("@c1", OleDbType.VarWChar).Value = textoConGuion
                        cmd.Parameters.Add("@c2", OleDbType.VarWChar).Value = textoSinGuion
                        cmd.Parameters.Add("@d1", OleDbType.VarWChar).Value = textoSeleccionado.ToUpper()
                        cmd.Parameters.Add("@d2", OleDbType.VarWChar).Value = textoSinGuion

                        Try
                            con.Open()
                            Dim res = cmd.ExecuteScalar()
                            If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                                idEncontrado = Convert.ToInt32(res)
                            End If
                        Catch
                        End Try
                    End Using
                End Using

                ' Si localizamos el número de ID entero en el disco duro, lo guardamos en el saco
                If idEncontrado > 0 AndAlso Not listaIDs.Contains(idEncontrado) Then
                    listaIDs.Add(idEncontrado)
                End If
            Next

            ' 🚀 LA ESTOCADA FINAL AL REVENTÓN DE TIPOS:
            ' Construimos la cláusula IN estrictamente con números enteros puros limpios
            If listaIDs.Count > 0 Then
                Dim cadenaIDs As String = String.Join(",", listaIDs)
                vtipoSql += " And apuntes.ConceptoAPU IN (" & cadenaIDs & ")"
            Else
                ' Salvavidas por si acaso la búsqueda del ID fallara: inyectamos un número imposible (-1)
                ' de esta forma la SQL se ejecuta limpia devolviendo 0 filas en lugar de reventar por tipos
                vtipoSql += " And apuntes.ConceptoAPU = -1"
            End If
        End If

        ' Ordenamos por el campo nativo numérico para tu bucle acumulador clásico de Pastebin
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql

        ' =========================================================================
        ' 🛠️ TU MSGBOX DE CONTROL: Verás cómo ahora el texto muta en números limpios
        ' =========================================================================
        'MsgBox("SQL Generada con Éxito (Nueva Era Numérica):" & vbNewLine & vtipoSql, MsgBoxStyle.Information, "DEBUG: SQL Final")

        ' Invocamos la macro modular a medida: limpia, rápida y directa
        vtipoGrid = "PRINT_GRAFICOS_SOLO"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        '' 🚀 CHIVATO 1: VALIDACIÓN DE FILAS INYECTADAS
        'If frmImprimirForm.DgvApuntes.Rows.Count = 0 Then
        '    MsgBox("CHIVATO 1 (Filtro): ¡Alerta! La rejilla frmImprimirForm.DgvApuntes se ha quedado con 0 filas. La SQL no ha devuelto nada de la base de datos.")
        'Else
        '    Dim vEjemploConcepto As String = If(frmImprimirForm.DgvApuntes.Rows(0).Cells(1).Value?.ToString(), "NULO")
        '    Dim vEjemploImporte As String = If(frmImprimirForm.DgvApuntes.Rows(0).Cells(3).Value?.ToString(), "NULO")
        '    MsgBox("CHIVATO 1 (Filtro): Registros cargados: " & frmImprimirForm.DgvApuntes.Rows.Count.ToString() & vbNewLine &
        '           "Fila 0 Celda 1 (Concepto): " & vEjemploConcepto & vbNewLine &
        '           "Fila 0 Celda 3 (Importe): " & vEjemploImporte)
        'End If

        ' =========================================================================
        ' 1. APERTURA DE TUS GRÁFICOS POR SOLO CONCEPTOS 2D (Tu lógica original intacta)
        ' =========================================================================
        If (frmGraficosSoloConceptos Is Nothing) OrElse (Not frmGraficosSoloConceptos.IsHandleCreated) Then
            frmGraficosSoloConceptos = New GraficosSoloConceptos
        End If
        frmGraficosSoloConceptos.EsGrafico3D = False
        frmGraficosSoloConceptos.ShowDialog()
        frmGraficosSoloConceptos.Dispose()
    End Sub

    Private Sub BtnContinuar3D_Click(sender As Object, e As EventArgs) Handles BtnContinuar3D.Click
        cmdMdb1cr.Parameters.Clear()

        ' =========================================================================
        ' 🌟 SENTENCIA RELACIONAL INTEGRAL NUMÉRICA PURA INDESTRUCTIBLE 3D
        ' =========================================================================
        vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU " &
                   "FROM apuntes " &
                   "WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString

        If ListBox1.SelectedItems.Count <> 0 Then
            ' Saco dinámico en la RAM para juntar los IDs numéricos enteros puros
            Dim listaIDs As New List(Of Integer)()

            For i As Integer = 0 To ListBox1.SelectedItems.Count - 1
                Dim textoSeleccionado As String = ListBox1.SelectedItems(i).ToString().Trim()

                ' Fabricamos las variantes posibles para que Access no tenga escapatoria con los guiones
                Dim textoConGuion As String = textoSeleccionado.Replace(" ", "_").ToUpper()
                Dim textoSinGuion As String = textoSeleccionado.Replace("_", " ").ToUpper()

                ' Vamos a buscar el ID numérico entero real correspondiente a ese concepto en Access
                Dim idEncontrado As Integer = 0
                Using con As New OleDbConnection(conexion1.ConnectionString)
                    Dim sqlBuscarID As String = "SELECT IdConceptoCON FROM conceptos WHERE " &
                                                  "UCASE(CodigoCON) = ? OR UCASE(CodigoCON) = ? OR " &
                                                  "UCASE(DescripcionCON) = ? OR UCASE(DescripcionCON) = ?"

                    Using cmd As New OleDbCommand(sqlBuscarID, con)
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add("@c1", OleDbType.VarWChar).Value = textoConGuion
                        cmd.Parameters.Add("@c2", OleDbType.VarWChar).Value = textoSinGuion
                        cmd.Parameters.Add("@d1", OleDbType.VarWChar).Value = textoSeleccionado.ToUpper()
                        cmd.Parameters.Add("@d2", OleDbType.VarWChar).Value = textoSinGuion

                        Try
                            con.Open()
                            Dim res = cmd.ExecuteScalar()
                            If res IsNot Nothing AndAlso Not IsDBNull(res) Then
                                idEncontrado = Convert.ToInt32(res)
                            End If
                        Catch
                        End Try
                    End Using
                End Using

                ' Si localizamos el número de ID entero, lo guardamos en el saco
                If idEncontrado > 0 AndAlso Not listaIDs.Contains(idEncontrado) Then
                    listaIDs.Add(idEncontrado)
                End If
            Next

            ' Construimos la cláusula IN estrictamente con números enteros puros limpios
            If listaIDs.Count > 0 Then
                Dim cadenaIDs As String = String.Join(",", listaIDs)
                vtipoSql += " And apuntes.ConceptoAPU IN (" & cadenaIDs & ")"
            Else
                vtipoSql += " And apuntes.ConceptoAPU = -1"
            End If
        End If

        ' Ordenamos por el campo nativo numérico para tu bucle acumulador clásico de Pastebin
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql

        ' Invocamos nuestra macro modular a medida: limpia, rápida y directa
        vtipoGrid = "PRINT_GRAFICOS_SOLO"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' =========================================================================
        ' 🚀 APERTURA EN FORMATO DE ALTA INGENIERÍA 3D
        ' =========================================================================
        If (frmGraficosSoloConceptos Is Nothing) OrElse (Not frmGraficosSoloConceptos.IsHandleCreated) Then
            frmGraficosSoloConceptos = New GraficosSoloConceptos
        End If

        ' Activamos el gatillo tridimensional de fábrica
        frmGraficosSoloConceptos.EsGrafico3D = True
        frmGraficosSoloConceptos.ShowDialog()
        frmGraficosSoloConceptos.Dispose()
    End Sub
End Class