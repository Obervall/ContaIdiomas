Public Class TipoGraficoPeriodico

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click

        ' =======================================================================
        ' 1. GRÁFICOS POR CONCEPTOS PERIODICOS (Unificado 2D y 3D) -> RadioButton1 y 6
        ' =======================================================================
        If RadioButton1.Checked = True Or RadioButton6.Checked = True Then
            If RadioButton6.Checked = True Then
                vGrafico = "CONCEPTOS3D_PERIODICOS"
            Else
                vGrafico = "CONCEPTOS_PERIODICOS"
            End If

            If (frmGraficosConceptos Is Nothing) OrElse (Not frmGraficosConceptos.IsHandleCreated) Then
                frmGraficosConceptos = New GraficosConceptos
            End If

            frmGraficosConceptos.EsGrafico3D = RadioButton6.Checked
            frmGraficosConceptos.ShowDialog()
            frmGraficosConceptos.Dispose()

            ' =======================================================================
            ' 2. GRÁFICOS POR CUENTAS PERIODICOS (Unificado 2D y 3D) -> RadioButton2 y 7
            ' =======================================================================
        ElseIf RadioButton2.Checked = True Or RadioButton7.Checked = True Then
            If RadioButton7.Checked = True Then
                vGrafico = "CUENTAS3D_PERIODICOS"
            Else
                vGrafico = "CUENTAS_PERIODICOS"
            End If

            If (frmGraficosCuentas Is Nothing) OrElse (Not frmGraficosCuentas.IsHandleCreated) Then
                frmGraficosCuentas = New GraficosCuentas
            End If

            frmGraficosCuentas.EsGrafico3D = RadioButton7.Checked
            frmGraficosCuentas.ShowDialog()
            frmGraficosCuentas.Dispose()

            ' =======================================================================
            ' 3. GRÁFICOS POR FECHAS PERIODICOS (Unificado 2D y 3D) -> RadioButton3 y 8
            ' =======================================================================
        ElseIf RadioButton3.Checked = True Or RadioButton8.Checked = True Then
            If RadioButton8.Checked = True Then
                vGrafico = "FECHAS3D_PERIODICOS"
            Else
                vGrafico = "FECHAS_PERIODICOS"
            End If

            If (frmGraficosFechas Is Nothing) OrElse (Not frmGraficosFechas.IsHandleCreated) Then
                frmGraficosFechas = New GraficosFechas
            End If

            frmGraficosFechas.EsGrafico3D = RadioButton8.Checked
            frmGraficosFechas.ShowDialog()
            frmGraficosFechas.Dispose()

            ' =======================================================================
            ' 4. GRÁFICOS POR MESES PERIODICOS (Unificado 2D y 3D) -> RadioButton4 y 9
            ' =======================================================================
        ElseIf RadioButton4.Checked = True Or RadioButton9.Checked = True Then
            If RadioButton9.Checked = True Then
                vGrafico = "MESES3D_PERIODICOS"
            Else
                vGrafico = "MESES_PERIODICOS"
            End If

            If (frmGraficosMeses Is Nothing) OrElse (Not frmGraficosMeses.IsHandleCreated) Then
                frmGraficosMeses = New GraficosMeses
            End If

            frmGraficosMeses.EsGrafico3D = RadioButton9.Checked
            frmGraficosMeses.ShowDialog()
            frmGraficosMeses.Dispose()

            ' =======================================================================
            ' 5. OPCIÓN EXTRA (Sin programar) -> RadioButton5
            ' =======================================================================
        ElseIf RadioButton5.Checked = True Then
            MsgBox(resManager.GetString("Vacio"))
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub TipoGraficoPeriodico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'ActualizarTextosFormulario(Me)
    End Sub
End Class