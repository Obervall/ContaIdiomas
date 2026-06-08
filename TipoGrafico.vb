Public Class TipoGrafico

    Private Sub TipoGrafico_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)
    End Sub


    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        vGrafico = "" 'Es para saber que son gráficos desde Apuntes, si no es vacio es para usar con los Periódicos.

        ' ------------------ GRÁFICOS 2D (PLANOS) ------------------
        If RadioButton1.Checked = True Then
            If (frmGraficosConceptos Is Nothing) OrElse (Not frmGraficosConceptos.IsHandleCreated) Then
                frmGraficosConceptos = New GraficosConceptos
            End If
            frmGraficosConceptos.EsGrafico3D = False ' <-- Le decimos que sea plano
            frmGraficosConceptos.ShowDialog()
            frmGraficosConceptos.Dispose()
        End If

        If RadioButton2.Checked = True Then
            If (frmGraficosCuentas Is Nothing) OrElse (Not frmGraficosCuentas.IsHandleCreated) Then
                frmGraficosCuentas = New GraficosCuentas
            End If
            frmGraficosCuentas.EsGrafico3D = False ' <-- Le decimos que sea plano
            frmGraficosCuentas.ShowDialog()
            frmGraficosCuentas.Dispose()
        End If

        If RadioButton3.Checked = True Then
            If (frmGraficosFechas Is Nothing) OrElse (Not frmGraficosFechas.IsHandleCreated) Then
                frmGraficosFechas = New GraficosFechas
            End If
            frmGraficosFechas.EsGrafico3D = False ' <-- Le decimos que sea plano
            frmGraficosFechas.ShowDialog()
            frmGraficosFechas.Dispose()
        End If

        If RadioButton4.Checked = True Then
            If (frmGraficosMeses Is Nothing) OrElse (Not frmGraficosMeses.IsHandleCreated) Then
                frmGraficosMeses = New GraficosMeses
            End If
            frmGraficosMeses.EsGrafico3D = False ' <-- Le decimos que sea plano
            frmGraficosMeses.ShowDialog()
            frmGraficosMeses.Dispose()
        End If

        If RadioButton5.Checked = True Then
            MsgBox("Sin Progamar ...")
        End If

        ' ------------------ GRÁFICOS 3D (RELIEVE) ------------------
        ' NOTA: ¡Apuntamos a tus formularios 3D actuales pero activando la propiedad en True!
        If RadioButton6.Checked = True Then
            If (frmGraficosConceptos3D Is Nothing) OrElse (Not frmGraficosConceptos3D.IsHandleCreated) Then
                frmGraficosConceptos3D = New GraficosConceptos3D
            End If
            frmGraficosConceptos3D.EsGrafico3D = True ' <-- ¡Le decimos que active el 3D!
            frmGraficosConceptos3D.ShowDialog()
            frmGraficosConceptos3D.Dispose()
        End If

        If RadioButton7.Checked = True Then
            If (frmGraficosCuentas3D Is Nothing) OrElse (Not frmGraficosCuentas3D.IsHandleCreated) Then
                frmGraficosCuentas3D = New GraficosCuentas3D
            End If
            frmGraficosCuentas3D.EsGrafico3D = True ' <-- ¡Le decimos que active el 3D!
            frmGraficosCuentas3D.ShowDialog()
            frmGraficosCuentas3D.Dispose()
        End If

        If RadioButton8.Checked = True Then
            If (frmGraficosFechas3D Is Nothing) OrElse (Not frmGraficosFechas3D.IsHandleCreated) Then
                frmGraficosFechas3D = New GraficosFechas3D
            End If
            frmGraficosFechas3D.EsGrafico3D = True ' <-- ¡Le decimos que active el 3D!
            frmGraficosFechas3D.ShowDialog()
            frmGraficosFechas3D.Dispose()
        End If

        If RadioButton9.Checked = True Then
            If (frmGraficosMeses3D Is Nothing) OrElse (Not frmGraficosMeses3D.IsHandleCreated) Then
                frmGraficosMeses3D = New GraficosMeses3D
            End If
            frmGraficosMeses3D.EsGrafico3D = True ' <-- ¡Le decimos que active el 3D!
            frmGraficosMeses3D.ShowDialog()
            frmGraficosMeses3D.Dispose()
        End If
    End Sub

    'Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
    '    vGrafico = "" 'Es para saber que son gráficos desde Apuntes, si no es vacio es para usar con los Periódicos.
    '    If RadioButton1.Checked = True Then
    '        'vGrafico = "CONCEPTOS"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosConceptos Is Nothing) OrElse (Not frmGraficosConceptos.IsHandleCreated) Then
    '            frmGraficosConceptos = New GraficosConceptos
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosConceptos.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosConceptos.Dispose()
    '    End If
    '    If RadioButton2.Checked = True Then
    '        'vGrafico = "CUENTAS"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosCuentas Is Nothing) OrElse (Not frmGraficosCuentas.IsHandleCreated) Then
    '            frmGraficosCuentas = New GraficosCuentas
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosCuentas.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosCuentas.Dispose()
    '    End If
    '    If RadioButton3.Checked = True Then
    '        'vGrafico = "FECHAS"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosFechas Is Nothing) OrElse (Not frmGraficosFechas.IsHandleCreated) Then
    '            frmGraficosFechas = New GraficosFechas
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosFechas.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosFechas.Dispose()
    '    End If
    '    If RadioButton4.Checked = True Then
    '        'vGrafico = "MESES"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosMeses Is Nothing) OrElse (Not frmGraficosMeses.IsHandleCreated) Then
    '            frmGraficosMeses = New GraficosMeses
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosMeses.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosMeses.Dispose()
    '    End If
    '    If RadioButton5.Checked = True Then
    '        MsgBox("Sin Progamar ...")
    '        'vGrafico = "SI"
    '    End If
    '    If RadioButton6.Checked = True Then
    '        'vGrafico = "CONCEPTOS 3D"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosConceptos3D Is Nothing) OrElse (Not frmGraficosConceptos3D.IsHandleCreated) Then
    '            frmGraficosConceptos3D = New GraficosConceptos3D
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosConceptos3D.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosConceptos3D.Dispose()
    '    End If
    '    If RadioButton7.Checked = True Then
    '        'vGrafico = "CUENTAS 3D"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosCuentas3D Is Nothing) OrElse (Not frmGraficosCuentas3D.IsHandleCreated) Then
    '            frmGraficosCuentas3D = New GraficosCuentas3D
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosCuentas3D.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosCuentas3D.Dispose()
    '    End If
    '    If RadioButton8.Checked = True Then
    '        'vGrafico = "FECHAS 3D"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosFechas3D Is Nothing) OrElse (Not frmGraficosFechas3D.IsHandleCreated) Then
    '            frmGraficosFechas3D = New GraficosFechas3D
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosFechas3D.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosFechas3D.Dispose()
    '    End If
    '    If RadioButton9.Checked = True Then
    '        'vGrafico = "MESES 3D"
    '        ' Comprobamos si existe un identificador asociado.
    '        If (frmGraficosMeses3D Is Nothing) OrElse (Not frmGraficosMeses3D.IsHandleCreated) Then
    '            frmGraficosMeses3D = New GraficosMeses3D
    '        End If
    '        ' Llamamos al formulario de manera modal.
    '        frmGraficosMeses3D.ShowDialog()
    '        'MessageBox.Show("Se ha cerrado el formulario.")
    '        ' Destruimos el formulario.
    '        frmGraficosMeses3D.Dispose()
    '    End If
    'End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class