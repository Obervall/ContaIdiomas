Public Class TipoGraficoPeriodico

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        If RadioButton1.Checked = True Then
            vGrafico = "CONCEPTOS_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosConceptos Is Nothing) OrElse (Not frmGraficosConceptos.IsHandleCreated) Then
                frmGraficosConceptos = New GraficosConceptos
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosConceptos.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosConceptos.Dispose()
        End If
        If RadioButton2.Checked = True Then
            vGrafico = "CUENTAS_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosCuentas Is Nothing) OrElse (Not frmGraficosCuentas.IsHandleCreated) Then
                frmGraficosCuentas = New GraficosCuentas
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosCuentas.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosCuentas.Dispose()
        End If
        If RadioButton3.Checked = True Then
            vGrafico = "FECHAS_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosFechas Is Nothing) OrElse (Not frmGraficosFechas.IsHandleCreated) Then
                frmGraficosFechas = New GraficosFechas
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosFechas.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosFechas.Dispose()
        End If
        If RadioButton4.Checked = True Then
            vGrafico = "MESES_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosMeses Is Nothing) OrElse (Not frmGraficosMeses.IsHandleCreated) Then
                frmGraficosMeses = New GraficosMeses
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosMeses.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosMeses.Dispose()
        End If
        If RadioButton5.Checked = True Then
            MsgBox("Sin Progamar ...")
            'vGrafico = "SI_PERIODICOS"
        End If
        If RadioButton6.Checked = True Then
            vGrafico = "CONCEPTOS3D_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosConceptos3D Is Nothing) OrElse (Not frmGraficosConceptos3D.IsHandleCreated) Then
                frmGraficosConceptos3D = New GraficosConceptos3D
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosConceptos3D.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosConceptos3D.Dispose()
        End If
        If RadioButton7.Checked = True Then
            vGrafico = "CUENTAS3D_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosCuentas3D Is Nothing) OrElse (Not frmGraficosCuentas3D.IsHandleCreated) Then
                frmGraficosCuentas3D = New GraficosCuentas3D
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosCuentas3D.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosCuentas3D.Dispose()
        End If
        If RadioButton8.Checked = True Then
            vGrafico = "FECHAS3D_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosFechas3D Is Nothing) OrElse (Not frmGraficosFechas3D.IsHandleCreated) Then
                frmGraficosFechas3D = New GraficosFechas3D
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosFechas3D.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosFechas3D.Dispose()
        End If
        If RadioButton9.Checked = True Then
            vGrafico = "MESES3D_PERIODICOS"
            ' Comprobamos si existe un identificador asociado.
            If (frmGraficosMeses3D Is Nothing) OrElse (Not frmGraficosMeses3D.IsHandleCreated) Then
                frmGraficosMeses3D = New GraficosMeses3D
            End If
            ' Llamamos al formulario de manera modal.
            frmGraficosMeses3D.ShowDialog()
            'MessageBox.Show("Se ha cerrado el formulario.")
            ' Destruimos el formulario.
            frmGraficosMeses3D.Dispose()
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub TipoGraficoPeriodico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
    End Sub
End Class