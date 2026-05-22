Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms

Public Class SeleccionFechas

    Public vtipoSql As String
    Public PrintLine, Contador As Integer
    Public PosicionSinEncabezado As Integer = frmImprimirForm.Punto1.Top

    Private Sub SeleccionFechas_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        vFecha1Enero = Val(vAñoEjercicio)
        vFecha31Diciembre = Val(vAñoEjercicio)
        DateTimePicker1.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker2.MinDate = New Date(vFecha1Enero, 1, 1)
        DateTimePicker1.Value = New Date(vFecha1Enero, 1, 1)
        DateTimePicker1.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.MaxDate = New Date(vFecha31Diciembre, 12, 31)
        DateTimePicker2.Value = New Date(vFecha31Diciembre, 12, 31)

        Dim TL(2) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnHoy, "Ir a Hoy")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnHoy2, "Ir a Hoy")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnAceptar, "Aceptar")

    End Sub

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click

        If vOrdenadoPorFechasAPU = 1 Or vOrdenadoPorConceptosAPU = 1 Or vOrdenadoPorImportesAPU = 1 Then
            vtipoSql = "SELECT apuntes.FechaAPU, apuntes.ConceptoAPU, apuntes.DescripcionAPU, apuntes.ImporteAPU, apuntes.ImporteAPU, apuntes.NotasAPU, apuntes.CuentaAPU, apuntes.CodigoAPU FROM apuntes"
            vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
            If vSoloIngresosAPU = 1 Then
                vtipoSql += " And apuntes.ImporteAPU > 0 "
            End If
            If vSoloGastosAPU = 1 Then
                vtipoSql += " And apuntes.ImporteAPU < 0 "
            End If
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            frmImprimirForm.LblEntreFechas.Text = "Desde: " & DateTimePicker1.Value & "    Hasta: " & DateTimePicker2.Value
            vtipoSql += " And apuntes.FechaAPU >= ?"
            vtipoSql += " And apuntes.FechaAPU <= ?"
            If vOrdenadoPorFechasAPU = 1 Then
                vtipoSql += " ORDER BY apuntes.FechaAPU ASC"
            End If
            If vOrdenadoPorConceptosAPU = 1 Then
                vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC"
            End If
            If vOrdenadoPorImportesAPU = 1 Then
                vtipoSql += " ORDER BY apuntes.ImporteAPU ASC"
            End If
        End If

        If vOrdenadoPorFechasAPP = 1 Or vOrdenadoPorConceptosAPP = 1 Or vOrdenadoPorImportesAPP = 1 Then
            vtipoSql = "SELECT apuper.FechaAPP, apuper.ConceptoAPP, apuper.DescripcionAPP, apuper.ImporteAPP, apuper.ImporteAPP, apuper.NotasAPP, apuper.CuentaAPP, apuper.CodigoAPP FROM apuper"
            vtipoSql += " WHERE apuper.EjercicioAPP = " & vAñoEjercicio.ToString
            If vSoloIngresosAPP = 1 Then
                vtipoSql += " And apuper.ImporteAPP > 0 "
            End If
            If vSoloGastosAPP = 1 Then
                vtipoSql += " And apuper.ImporteAPP < 0 "
            End If
            vDate1 = DateTimePicker1.Value.Date
            vDate2 = DateTimePicker2.Value.Date
            frmImprimirForm.LblEntreFechas.Text = "Desde: " & DateTimePicker1.Value & "    Hasta: " & DateTimePicker2.Value
            vtipoSql += " And apuper.FechaAPP >= ?"
            vtipoSql += " And apuper.FechaAPP <= ?"
            If vOrdenadoPorFechasAPP = 1 Then
                vtipoSql += " ORDER BY apuper.FechaAPP ASC"
            End If
            If vOrdenadoPorConceptosAPP = 1 Then
                vtipoSql += " ORDER BY apuper.ConceptoAPP ASC"
            End If
            If vOrdenadoPorImportesAPP = 1 Then
                vtipoSql += " ORDER BY apuper.ImporteAPP ASC"
            End If
        End If
        LlenarGrid(vtipoSql, "PRINT_INFORME_APUNTES", 1)
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString

        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        If My.Settings.Previsualizar = True Then
            'Te deja ver un preview del reporte antes de imprimir
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.WindowState = FormWindowState.Maximized
            PrintPreviewDialog1.ShowDialog()
        End If

        If My.Settings.ElegirImpresora = True Then
            'Te deja elegir la impresora
            PrintDialog1.Document = PrintDocument1
            PrintDialog1.PrinterSettings = PrintDocument1.PrinterSettings
            PrintDialog1.AllowSomePages = True
            If PrintDialog1.ShowDialog = DialogResult.OK Then
                PrintDocument1.PrinterSettings = PrintDialog1.PrinterSettings
                PrintDocument1.Print()
            End If
        End If

        If My.Settings.DirectoImpresora = True Then
            'Imprime en la impresora por defecto
            PrintDocument1.Print()
        End If

        Me.Close()
    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Cualquier variable que desees que conserve su valor debes declararla fuera del Printdocument
        'Todas las variable declaradas dentro de printdocument pierden su valor al cambiar de pagina
        'Definimos los tipos de letras a utilizar en el reporte
        '******************************************************
        Dim FuenteTitulo As New Font("Microsoft Sans Serif", 15)
        Dim FuenteSubtitulo As New Font("Microsoft Sans Serif", 16)
        Dim FuenteNegrita As New Font("Microsoft Sans Serif", 9, FontStyle.Bold)
        Dim FuenteDetalles As New Font("Microsoft Sans Serif", 9)
        Dim FuenteSubrayada As New Font("Microsoft Sans Serif", 9, FontStyle.Underline Xor FontStyle.Bold)
        Dim sf As New StringFormat With {.Alignment = StringAlignment.Far}

        frmImprimirForm.LblTitulo.Text = vTituloInforme

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)
        e.Graphics.DrawString(frmImprimirForm.LblEntreFechas.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblEntreFechas.Right, frmImprimirForm.LblEntreFechas.Top)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString("Fecha:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString("Concepto Contable:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto2.Left, frmImprimirForm.Punto2.Top - 30)
        e.Graphics.DrawString("Descripción:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        e.Graphics.DrawString("Importe " & vMoneda & ":", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Left, frmImprimirForm.Punto5.Top - 30)

        'imprimimos la linea debajo de los encabezados
        '*********************************************
        e.Graphics.DrawString(frmImprimirForm.LineaTop.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaTop.Left, frmImprimirForm.LineaTop.Top)

        'Imprimimos los detalles del reporte, es decir el listado de Apuntes
        '*******************************************************************
        Dim startX As Integer = frmImprimirForm.Punto1.Left 'Tomamos la posicion horinzontal de la letra 'Punto1'
        Dim startY As Integer = frmImprimirForm.Punto1.Top 'Tomamos la posicion vertical de la letra 'Punto1'
        Do While PrintLine < frmImprimirForm.DgvApuntes.Rows.Count
            If startY + frmImprimirForm.Punto1.Height > e.MarginBounds.Bottom Then
                'Esta parte se activa solo si 'startY' que es la posicion vertical almacenada supera el borde inferior de la pagina
                'Este se reinicia con cada pagina necesitada
                e.HasMorePages = True
                Exit Do
            End If
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(0).Value, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto2.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(2).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Left, startY)
            e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(3).Value, "###,##0.00").ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto5.Right + 50, startY, sf)
            'Aqui estoy usando un tipo de letras mas grande
            'LabelCodigo' mas grande que 'Punto1' para crear mas espacio entre filas

            'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
            'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
            ''Imprimimos los valores que salen despues del datagridview al final del reporte

            startY += frmImprimirForm.LblFecha.Height
            PrintLine += 1
            Contador += 1
        Loop
        'Con el contador solamente imprimimos la parte final del reporte si ha alcanzado el total de registros
        'Si deseamos repetir la parte final del reporte en cada pagina, debemos quitar en contador
        'Imprimimos los valores que salen despues del datagridview al final del reporte
        If Contador >= frmImprimirForm.DgvApuntes.Rows.Count Then
            e.Graphics.DrawString(frmImprimirForm.LineaFondo.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LineaFondo.Left, startY)
            e.Graphics.DrawString(frmImprimirForm.LblTotal.Text, FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto5.Right + 50, startY + 15, sf)

            'Para volver a dejar a 0, cuando se imprime desde la Vista Previa
            PrintLine = 0
            Contador = 0
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)

        'Para volver a dejar a 0 las páginas, cuando se imprime desde la Vista Previa
        If Contador = 0 Then
            frmImprimirForm.LblNumeroPagina.Text = "0"
        End If
    End Sub

    Private Sub BtnHoy_Click(sender As Object, e As EventArgs) Handles BtnHoy.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox("El año del ejercicio no coincide con el año actual," & vbCrLf & "se establecerá la fecha del 1 de Enero del año del ejercicio", MsgBoxStyle.Information, "Fecha establecida al 1 de Enero")
            DateTimePicker1.Value = New Date(vAñoEjercicio, 1, 1)
        Else
            DateTimePicker1.Value = DateTime.Today
        End If
    End Sub

    Private Sub BtnHoy2_Click(sender As Object, e As EventArgs) Handles BtnHoy2.Click
        If vAñoEjercicio <> vAñoActual Then
            MsgBox("El año del ejercicio no coincide con el año actual," & vbCrLf & "se establecerá la fecha del 31 de Diciembre del año del ejercicio", MsgBoxStyle.Information, "Fecha establecida al 31 de Diciembre")
            DateTimePicker2.Value = New Date(vAñoEjercicio, 12, 31)
        Else
            DateTimePicker2.Value = DateTime.Today
        End If
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub

    Private Sub DateTimePicker2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles DateTimePicker2.KeyPress
        If e.KeyChar = ChrW(Keys.Enter) Then
            BtnAceptar.Select()
        End If
    End Sub
End Class