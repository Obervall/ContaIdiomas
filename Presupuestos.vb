Imports System.Data
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Windows.Forms
Imports ToolTip = System.Windows.Forms.ToolTip

Public Class Presupuestos

    Public vtipoSql, vtipoGrid, vConcepto, vBorrarPresu, vAñadir, vAñadir2 As String
    Public vTmpprint As String
    Public vSaldoAnualPresupuesto, vImporteConcepto, vImporteConcepto2, vExistenteImporteConcepto, vNewImporteConcepto As Double
    Public PrintLine, Contador As Integer

    Private Sub Presupuestos_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ActualizarTextosFormulario(Me)

        Dim TL(9) As ToolTip
        TL(0) = New ToolTip
        TL(0).SetToolTip(Me.BtnGraficos, "Mostrar Gráficos")
        TL(1) = New ToolTip
        TL(1).SetToolTip(Me.BtnSalir, "Salir de Presupuestos")
        TL(2) = New ToolTip
        TL(2).SetToolTip(Me.BtnFiltroConcepto, "Aplica el filtro a los Registros")
        TL(3) = New ToolTip
        TL(3).SetToolTip(Me.BtnSinFiltroConcepto, "Quitar el filtro a los Registros")
        TL(4) = New ToolTip
        TL(4).SetToolTip(Me.BtnImprimir, "Imprimir")
        TL(5) = New ToolTip
        TL(5).SetToolTip(Me.BtnPrimero, "Ir al Primer Registro")
        TL(6) = New ToolTip
        TL(6).SetToolTip(Me.BtnAnterior, "Ir al Anterior Registro")
        TL(7) = New ToolTip
        TL(7).SetToolTip(Me.BtnSiguiente, "Ir al Siguiente Registro")
        TL(8) = New ToolTip
        TL(8).SetToolTip(Me.BtnUltimo, "Ir al Ultimo Registro")
        TL(9) = New ToolTip
        TL(9).SetToolTip(Me.BtnEliminarRegistro, "Eliminar Concepto en Presupuestos")


        ' Llenar el Combo Concepto
        '*************************
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos ORDER BY conceptos.CodigoCON ASC"
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    CmbConcepto.Items.Add(drMdb1.GetValue(0))
                End While
                CmbConcepto.Text = CmbConcepto.Items(0)
            Else
                'MsgBox("No existen registros en " & tipoSql)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("No se han podido cargar los Conceptos en el Combo, revise el código !!!")
            MsgBox(ex.ToString)
        End Try

        ' Llenar Grid de PRESUPUESTOS
        '****************************
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
    End Sub

    Private Sub BtnSinFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnSinFiltroConcepto.Click
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        LblDesviacion.Enabled = False
        LblObjetivo.Visible = False
        TxtDesviacion.Text = ""
    End Sub

    Private Sub BtnFiltroConcepto_Click(sender As Object, e As EventArgs) Handles BtnFiltroConcepto.Click
        BtnFiltroConcepto.Enabled = False
        BtnSinFiltroConcepto.Enabled = True
        vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
        vtipoSql += " WHERE "
        vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
        vtipoSql += " And presupuesto.ConceptoPRE = '" & CmbConcepto.Text & "' "
        vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
        vtipoGrid = "PRESUPUESTOS"
        LlenarGrid(vtipoSql, vtipoGrid, "1")
        ' Calcular el Toatal Anual de lo Presupuestado
        '*********************************************
        cmdMdb1cr.CommandText = vtipoSql
        Try
            drMdb1 = cmdMdb1cr.ExecuteReader()
            vSaldoAnualPresupuesto = 0
            If drMdb1.HasRows Then
                While drMdb1.Read()
                    vSaldoAnualPresupuesto += drMdb1.GetValue(2)
                End While
                LblDesviacion.Enabled = True
                TxtDesviacion.Text = vSaldoAnualPresupuesto - vSaldoAnualReal
                LblObjetivo.Visible = True
                If TxtDesviacion.Text >= 0 Then
                    LblObjetivo.ForeColor = Color.DarkGreen
                    LblObjetivo.Text = "Objetivo Logrado!"
                Else
                    LblObjetivo.ForeColor = Color.DarkRed
                    LblObjetivo.Text = "Objetivo NO Logardo!"
                End If
            Else
                LblDesviacion.Enabled = False
                TxtDesviacion.Text = ""
                LblObjetivo.Visible = False
                'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
            End If
            drMdb1.Close()
        Catch ex As Exception
            MsgBox("No se ha podido calcular el total anual presupuestado, revise el código !!!")
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub BtnGraficos_Click(sender As Object, e As EventArgs) Handles BtnGraficos.Click
        ' Comprobamos si existe un identificador asociado.
        If ((frmGraficosPresupuestos Is Nothing) OrElse (Not frmGraficosPresupuestos.IsHandleCreated)) Then
            frmGraficosPresupuestos = New GraficosPresupuestos
        End If
        ' Llamamos al formulario de manera modal.
        frmGraficosPresupuestos.ShowDialog()
        'MessageBox.Show("Se ha cerrado el formulario.")
        ' Destruimos el formulario.
        frmGraficosPresupuestos.Dispose()
    End Sub

    Private Sub CmbTipoConcepto_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmbConcepto.SelectedIndexChanged
        ' Se buscan Conceptos según lo seleccionado
        '******************************************
        vConcepto = CmbConcepto.Text.ToString
        drMdb1.Close()
        cmdMdb1cr.CommandText = "SELECT * FROM conceptos Where conceptos.CodigoCON = '" & vConcepto & "' "
        drMdb1 = cmdMdb1cr.ExecuteReader()
        drMdb1.Read()
        TxtConcepto.Text = drMdb1.GetValue(1)
        drMdb1.Close()

        If BtnFiltroConcepto.Enabled = False Then
            vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
            vtipoSql += " WHERE "
            vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
            vtipoSql += " And presupuesto.ConceptoPRE = '" & CmbConcepto.Text & "' "
            vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
            vtipoGrid = "PRESUPUESTOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
            ' Calcular el Toatal Anual de lo Presupuestado
            '*********************************************
            cmdMdb1cr.CommandText = vtipoSql
            Try
                drMdb1 = cmdMdb1cr.ExecuteReader()
                vSaldoAnualPresupuesto = 0
                If drMdb1.HasRows Then
                    While drMdb1.Read()
                        vSaldoAnualPresupuesto += drMdb1.GetValue(3)
                    End While
                    LblDesviacion.Enabled = True
                    TxtDesviacion.Text = vSaldoAnualPresupuesto - vSaldoAnualReal
                    LblObjetivo.Visible = True
                    If TxtDesviacion.Text >= 0 Then
                        LblObjetivo.ForeColor = Color.DarkGreen
                        LblObjetivo.Text = "Objetivo Logrado!"
                    Else
                        LblObjetivo.ForeColor = Color.DarkRed
                        LblObjetivo.Text = "Objetivo NO Logardo!"
                    End If
                Else
                    LblDesviacion.Enabled = False
                    TxtDesviacion.Text = ""
                    LblObjetivo.Visible = False
                    'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                End If
                drMdb1.Close()
            Catch ex As Exception
                MsgBox("No se ha podido calcular el total anual presupuestado, revise el código !!!")
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
    End Sub

    Private Sub Main_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        ' Se verifica si la razón para cerrar es la 3, es decir, el botón X.
        If e.CloseReason = 3 Then
            e.Cancel = False ' NO Se cancela la solicitud de cerrar
        End If
        BtnFiltroConcepto.Enabled = True
        BtnSinFiltroConcepto.Enabled = False
    End Sub

    Private Sub BtnPrimero_Click(sender As Object, e As EventArgs) Handles BtnPrimero.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = 0
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnAnterior_Click(sender As Object, e As EventArgs) Handles BtnAnterior.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = 0 Then
            MsgBox("Fila Primera Seleccionada")
        Else
            vFila = vFilaActual - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnImprimir_Click(sender As Object, e As EventArgs) Handles BtnImprimir.Click

        'Iniciamos Tabla Tmpprint
        '************************
        vTmpprint = "DELETE FROM tmpprint"
        cmdMdb1cr.CommandText = vTmpprint
        Try
            cmdMdb1cr.ExecuteNonQuery()
            'MsgBox("Registros Tmpprint, Borrados !!!")
        Catch ex As Exception
            MsgBox("No se han podido eliminar los registros en tmpprint, revise el código !!!")
            MsgBox(ex.ToString)
        End Try

        'Ordenamos la columna Concepto, antes de calcular los totales parciales.
        '***********************************************************************
        frmPresupuestos.DgvPresupuestos.Sort(frmPresupuestos.DgvPresupuestos.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

        'Llenamos la tabla tmpprint con los Conceptos Agrupados desde DgvPresupuestos
        '****************************************************************************
        vNombreConcepto = ""
        For Each fila As DataGridViewRow In frmPresupuestos.DgvPresupuestos.Rows
            vImporteConcepto = fila.Cells(3).Value
            If vNombreConcepto <> fila.Cells(0).Value.ToString Then
                vNombreConcepto = fila.Cells(0).Value.ToString
                vImporteConcepto = 0
                vImporteConcepto = fila.Cells(3).Value

                'Buscamos la suma total del Concepto
                '***********************************
                vtipoSql = "SELECT * FROM apuntes"
                vtipoSql += " WHERE apuntes.EjercicioAPU = " & vAñoEjercicio.ToString
                vtipoSql += " And apuntes.ConceptoAPU = '" & vNombreConcepto & "' "
                vtipoSql += " ORDER BY apuntes.ConceptoAPU ASC"
                LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")  'Agrupado Conceptos

                vImporteConcepto2 = "0,00"
                For Each filas As DataGridViewRow In frmImprimirForm.DgvApuntes.Rows
                    vImporteConcepto2 += Math.Abs(filas.Cells(4).Value)
                Next
                vAñadir = "INSERT INTO tmpprint"
                vAñadir += "(FechaTMP, ConceptoTMP, DescripcionTMP, CuentaTMP, NotasTMP, ImporteTMP, SaldoTMP) "
                vAñadir += "VALUES ('2023-01-11', '" & vNombreConcepto & "', '', '', '' , '" & vImporteConcepto2 & "', '" & vImporteConcepto & "')"
                cmdMdb1cr.CommandText = vAñadir
                Try
                    cmdMdb1cr.ExecuteNonQuery()
                    'MsgBox("Registro1, Grabado Correctamente " & vNombreConcepto)
                Catch ex As Exception
                    MsgBox("No se ha podido grabar el registro en tmpprint, revise el código !!!")
                    MsgBox(ex.ToString)
                End Try
            Else
                cmdMdb1cr.CommandType = CommandType.Text
                cmdMdb1cr.CommandText = "SELECT * FROM tmpprint WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    If drMdb1.HasRows Then
                        While drMdb1.Read()
                            vExistenteImporteConcepto = drMdb1.GetValue(6)
                        End While
                    Else
                        'MsgBox("No existen registros en " & cmdMdb1cr.CommandText)
                    End If
                    drMdb1.Close()
                Catch ex As Exception
                    MsgBox("No se ha podido leer el registro en tmpprint, revise el código !!!")
                    MsgBox(ex.ToString)
                End Try
                vNewImporteConcepto = vImporteConcepto + vExistenteImporteConcepto
                vAñadir2 = "UPDATE tmpprint SET SaldoTMP = '" & vNewImporteConcepto & "' "
                vAñadir2 += " WHERE tmpprint.ConceptoTMP = '" & vNombreConcepto & "' "
                cmdMdb1cr.CommandText = vAñadir2
                Try
                    drMdb1 = cmdMdb1cr.ExecuteReader()
                    'MsgBox(vImporteConcepto & " Registro2, Grabado Correctamente " & vNombreConcepto & " " & vNewImporteConcepto)
                Catch ex As Exception
                    MsgBox("No se ha podido actualizar el registro en tmpprint, revise el código !!!")
                    MsgBox(ex.ToString)
                End Try
                drMdb1.Close()
            End If
        Next

        vtipoSql = "SELECT * FROM tmpprint"
        vtipoSql += " ORDER BY tmpprint.ConceptoTMP ASC"
        LlenarGrid(vtipoSql, "PRINT_APUNTES_CONTABLES", "2")

        'Iniciamos Código para Imprimir
        '******************************
        frmImprimirForm.LblFecha.Text = Date.Today.ToLongDateString
        PrintLine = 0
        Contador = 0
        frmImprimirForm.LblNumeroPagina.Text = "0"

        'Para ver la plantilla de impresión
        'frmImprimirForm.Show()

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

        If BtnFiltroConcepto.Enabled = False Then
            frmImprimirForm.LblTitulo.Text = "Listado de Presupuestos. Filtrado por: " & CmbConcepto.Text
        Else
            frmImprimirForm.LblTitulo.Text = "Listado de Presupuestos"
        End If

        'Imprimimos el encabezado los datos que están antes del datagridview
        '*******************************************************************
        'e.Graphics.DrawString(frmImprimirForm.LblUsuario.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblUsuario.Left, frmImprimirForm.LblUsuario.Top)
        e.Graphics.DrawString(frmImprimirForm.LblFecha.Text, FuenteNegrita, Brushes.Black, frmImprimirForm.LblFecha.Right, frmImprimirForm.LblFecha.Top)
        e.Graphics.DrawString(frmImprimirForm.LblTitulo.Text, FuenteTitulo, Brushes.Black, frmImprimirForm.LblTitulo.Left, frmImprimirForm.LblTitulo.Top)
        Dim newImage As Image = frmImprimirForm.PictureBox1.Image : e.Graphics.DrawImage(newImage, frmImprimirForm.PictureBox1.Left, frmImprimirForm.PictureBox1.Top, frmImprimirForm.PictureBox1.Width, frmImprimirForm.PictureBox1.Height)

        'Imprimimos el encabezado o titulo de la lista de materias por encima de los puntos definidos
        '********************************************************************************************
        e.Graphics.DrawString("Concepto Contable:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto1.Left, frmImprimirForm.Punto1.Top - 30)
        e.Graphics.DrawString("Real:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto3.Left, frmImprimirForm.Punto3.Top - 30)
        e.Graphics.DrawString("Presupuestado:", FuenteSubrayada, Brushes.Black, frmImprimirForm.Punto4.Left, frmImprimirForm.Punto4.Top - 30)

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
            e.Graphics.DrawString(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(1).Value.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto1.Left, startY)
            e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(5).Value, "###,##0.00 ") & vMoneda.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto3.Right + 20, startY, sf)
            e.Graphics.DrawString(Format(frmImprimirForm.DgvApuntes.Rows(PrintLine).Cells(6).Value, "###,##0.00 ") & vMoneda.ToString, FuenteDetalles, Brushes.Black, frmImprimirForm.Punto4.Right + 50, startY, sf)
            startY += frmImprimirForm.LblFecha.Height
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
        End If

        'Si deseamos poner un contador de páginas
        'Esta parte siempre va a salir en todas las paginas
        frmImprimirForm.LblNumeroPagina.Text = CInt(frmImprimirForm.LblNumeroPagina.Text) + 1
        e.Graphics.DrawString(frmImprimirForm.Label2.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.Label2.Left, e.MarginBounds.Bottom)
        e.Graphics.DrawString(frmImprimirForm.LblNumeroPagina.Text, FuenteDetalles, Brushes.Black, frmImprimirForm.LblNumeroPagina.Left, e.MarginBounds.Bottom)
    End Sub

    Private Sub BtnEliminarRegistro_Click(sender As Object, e As EventArgs) Handles BtnEliminarRegistro.Click
        respuesta = MsgBox("¿Estas seguro de Eliminar Todos los presupuestos del Concepto seleccionado?.", vbQuestion + vbYesNo + vbDefaultButton2, "Eliminar presupuesto")
        If respuesta = vbYes Then
            ' Eliminar Registro Apunte
            filaActual = DgvPresupuestos.CurrentRow.Index
            vBorrarPresu = DgvPresupuestos.Rows(filaActual).Cells(0).Value.ToString

            vtipoSql = "DELETE FROM presupuesto"
            vtipoSql += " WHERE presupuesto.ConceptoPRE = '" & vBorrarPresu & "' "
            cmdMdb1cr.CommandText = vtipoSql

            Try
                cmdMdb1cr.ExecuteNonQuery()
                MsgBox("Registros en Presupuestos, Borrados !!!")
            Catch ex As Exception
                MsgBox("No se han podido eliminar los registros en Presupuestos, revise que no existan apuntes asociados al concepto seleccionado !!!")
                MsgBox(ex.ToString)
            End Try
            ' Llenar Grid de PRESUPUESTOS
            '****************************
            vtipoSql = "SELECT presupuesto.ConceptoPRE, presupuesto.ConceptoPRE, presupuesto.ImportePRE, presupuesto.ImportePRE, presupuesto.FDesdePRE FROM presupuesto"
            vtipoSql += " WHERE "
            vtipoSql += "presupuesto.EjercicioPRE = " & vAñoEjercicio.ToString
            vtipoSql += " ORDER BY presupuesto.ConceptoPRE ASC, presupuesto.FDesdePRE ASC"
            vtipoGrid = "PRESUPUESTOS"
            LlenarGrid(vtipoSql, vtipoGrid, "1")
        End If
    End Sub

    Private Sub BtnSiguiente_Click(sender As Object, e As EventArgs) Handles BtnSiguiente.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = vFilaActual + 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub BtnUltimo_Click(sender As Object, e As EventArgs) Handles BtnUltimo.Click
        vFilaActual = DgvPresupuestos.CurrentRow.Index
        If vFilaActual = DgvPresupuestos.RowCount - 1 Then
            MsgBox("Fila Ultima Seleccionada")
        Else
            vFila = DgvPresupuestos.RowCount - 1
            DgvPresupuestos.Rows(vFila).Selected = True
            DgvPresupuestos.CurrentCell = DgvPresupuestos.Rows(vFila).Cells(0)
        End If
    End Sub

    Private Sub CmbConcepto_KeyPress(sender As Object, e As KeyPressEventArgs) Handles CmbConcepto.KeyPress
        e.KeyChar = Char.ToUpper(e.KeyChar)
    End Sub
End Class