Imports System.Windows.Forms

Public Class SeleccionDatosGastos

    Public i As Integer
    Public vConcepto As String

    Private Sub SeleccionDatos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

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
            MsgBox("Error al cargar el ListBox de Conceptos de Gastos" & vbCrLf & ex.ToString)
        End Try
    End Sub

    Private Sub BtnTodos_Click(sender As Object, e As EventArgs) Handles BtnTodos.Click
        If BtnTodos.Text = "Seleccionar Todos" Then
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, True)
            Next
            BtnTodos.Text = "Deseleccionar Todos"
        Else
            For i = 0 To ListBox1.Items.Count - 1
                ListBox1.SetSelected(i, False)
                ListBox1.SetSelected(0, True)
            Next
            BtnTodos.Text = "Seleccionar Todos"
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
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                Else
                    vtipoSql += " Or "
                    vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                End If
            Next
        Else

        End If
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql
        vtipoGrid = "PRINT_APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        'vGrafico = "SOLO GASTOS / INGRESOS CONCEPTOS"
        ' Comprobamos si existe un identificador asociado.
        If (frmGraficosSoloConceptos Is Nothing) OrElse (Not frmGraficosSoloConceptos.IsHandleCreated) Then
            frmGraficosSoloConceptos = New GraficosSoloConceptos
        End If
        ' Llamamos al formulario de manera modal.
        frmGraficosSoloConceptos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmGraficosSoloConceptos.Dispose()
    End Sub

    Private Sub BtnContinuar3D_Click(sender As Object, e As EventArgs) Handles BtnContinuar3D.Click
        If ListBox1.SelectedItems.Count <> 0 Then
            Dim i As Integer
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            For i = 0 To ListBox1.SelectedItems.Count - 1
                vConcepto = ListBox1.SelectedItems(i).ToString
                If i = 0 Then
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                Else
                    vtipoSql += " Or "
                    vtipoSql += "apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                    vtipoSql += " And apuntes.ConceptoAPU = '" & vConcepto & "' "
                End If
            Next
        Else

        End If
        vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC, apuntes.FechaAPU ASC"
        vtipoSqlChk = vtipoSql
        vtipoGrid = "PRINT_APUNTES_CONTABLES"
        LlenarGrid(vtipoSql, vtipoGrid, "1")

        'vGrafico = "SOLO GASTOS / INGRESOS CONCEPTOS"
        ' Comprobamos si existe un identificador asociado.
        If (frmGraficosSoloConceptos3D Is Nothing) OrElse (Not frmGraficosSoloConceptos3D.IsHandleCreated) Then
            frmGraficosSoloConceptos3D = New GraficosSoloConceptos3D
        End If
        ' Llamamos al formulario de manera modal.
        frmGraficosSoloConceptos3D.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmGraficosSoloConceptos3D.Dispose()
    End Sub
End Class