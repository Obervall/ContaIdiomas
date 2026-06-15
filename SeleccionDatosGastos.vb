Public Class SeleccionDatosGastos

    Public i As Integer
    Public vConcepto As String
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub SeleccionDatosGastos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        cmdMdb1cr.CommandText = "SELECT * FROM conceptos "
        cmdMdb1cr.CommandText += "Where conceptos.TipoCON = 'GASTO' ORDER BY conceptos.CodigoCON ASC"
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
        If ListBox1.SelectedItems.Count <> 0 Then
            Dim i As Integer
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString
                If i = 0 Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                Else
                    vtipoSql += " Or "
                    vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                End If
            Next
        Else

        End If
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql
        vtipoGrid = "PRINT_APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' =================================
        ' 1. GRÁFICOS POR SOLO CONCEPTOS 2D
        ' =================================
        If (frmGraficosSoloConceptos Is Nothing) OrElse (Not frmGraficosSoloConceptos.IsHandleCreated) Then
            frmGraficosSoloConceptos = New GraficosSoloConceptos
        End If
        frmGraficosSoloConceptos.EsGrafico3D = False
        frmGraficosSoloConceptos.ShowDialog()
        frmGraficosSoloConceptos.Dispose()
    End Sub

    Private Sub BtnContinuar3D_Click(sender As Object, e As EventArgs) Handles BtnContinuar3D.Click
        If ListBox1.SelectedItems.Count <> 0 Then
            Dim i As Integer
            vtipoSql = "Select apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString
                If i = 0 Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                Else
                    vtipoSql += " Or "
                    vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto.Replace("'", "''") & "' "
                End If
            Next
        Else

        End If
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql
        vtipoGrid = "PRINT_APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        ' =================================
        ' 1. GRÁFICOS POR SOLO CONCEPTOS 3D
        ' =================================
        If (frmGraficosSoloConceptos Is Nothing) OrElse (Not frmGraficosSoloConceptos.IsHandleCreated) Then
            frmGraficosSoloConceptos = New GraficosSoloConceptos
        End If
        frmGraficosSoloConceptos.EsGrafico3D = True
        frmGraficosSoloConceptos.ShowDialog()
        frmGraficosSoloConceptos.Dispose()
    End Sub
End Class