Public Class TipoGrafico

    Private Sub TipoGrafico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ActualizarTextosFormulario(Me)
    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vGrafico = "" ' Origen desde Apuntes

        ' =======================================================================
        ' 1. GRÁFICOS POR CONCEPTOS (Unificado 2D y 3D) -> RadioButton1 y 6
        ' =======================================================================
        If RadioButton1.Checked = True Or RadioButton6.Checked = True Then
            If (frmGraficosConceptos Is Nothing) OrElse (Not frmGraficosConceptos.IsHandleCreated) Then
                frmGraficosConceptos = New GraficosConceptos
            End If
            frmGraficosConceptos.EsGrafico3D = RadioButton6.Checked
            frmGraficosConceptos.ShowDialog()
            frmGraficosConceptos.Dispose()

            ' =======================================================================
            ' 2. GRÁFICOS POR CUENTAS (Unificado 2D y 3D) -> RadioButton2 y 7
            ' =======================================================================
        ElseIf RadioButton2.Checked = True Or RadioButton7.Checked = True Then
            If (frmGraficosCuentas Is Nothing) OrElse (Not frmGraficosCuentas.IsHandleCreated) Then
                frmGraficosCuentas = New GraficosCuentas
            End If
            frmGraficosCuentas.EsGrafico3D = RadioButton7.Checked
            frmGraficosCuentas.ShowDialog()
            frmGraficosCuentas.Dispose()

            ' =======================================================================
            ' 3. GRÁFICOS POR FECHAS (Unificado 2D y 3D) -> RadioButton3 y 8
            ' =======================================================================
        ElseIf RadioButton3.Checked = True Or RadioButton8.Checked = True Then
            If (frmGraficosFechas Is Nothing) OrElse (Not frmGraficosFechas.IsHandleCreated) Then
                frmGraficosFechas = New GraficosFechas
            End If
            frmGraficosFechas.EsGrafico3D = RadioButton8.Checked
            frmGraficosFechas.ShowDialog()
            frmGraficosFechas.Dispose()

            ' =======================================================================
            ' 4. GRÁFICOS POR MESES (Unificado 2D y 3D) -> RadioButton4 y 9
            ' =======================================================================
        ElseIf RadioButton4.Checked = True Or RadioButton9.Checked = True Then
            If (frmGraficosMeses Is Nothing) OrElse (Not frmGraficosMeses.IsHandleCreated) Then
                frmGraficosMeses = New GraficosMeses
            End If
            frmGraficosMeses.EsGrafico3D = RadioButton9.Checked
            frmGraficosMeses.ShowDialog()
            frmGraficosMeses.Dispose()

            ' =======================================================================
            ' OPCIÓN EXTRA (Sin programar) -> RadioButton5
            ' =======================================================================
        ElseIf RadioButton5.Checked = True Then
            MsgBox("Sin Programar ...")
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class