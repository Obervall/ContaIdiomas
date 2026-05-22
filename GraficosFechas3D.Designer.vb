<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GraficosFechas3D
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GraficosFechas3D))
        Dim ChartArea1 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim Series1 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Series2 As System.Windows.Forms.DataVisualization.Charting.Series = New System.Windows.Forms.DataVisualization.Charting.Series()
        Dim Title1 As System.Windows.Forms.DataVisualization.Charting.Title = New System.Windows.Forms.DataVisualization.Charting.Title()
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.TsBtnColumnas = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripLabel3 = New System.Windows.Forms.ToolStripLabel()
        Me.TsBtnAreas = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripLabel4 = New System.Windows.Forms.ToolStripLabel()
        Me.TsBtnLineas = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripLabel5 = New System.Windows.Forms.ToolStripLabel()
        Me.TsBtnPastel = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripLabel2 = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.TSBtnImprimir = New System.Windows.Forms.ToolStripButton()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TsBtnColumnas, Me.ToolStripLabel3, Me.TsBtnAreas, Me.ToolStripLabel4, Me.TsBtnLineas, Me.ToolStripLabel5, Me.TsBtnPastel, Me.ToolStripLabel2, Me.ToolStripSeparator2, Me.TSBtnImprimir})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.Size = New System.Drawing.Size(1100, 31)
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'TsBtnColumnas
        '
        Me.TsBtnColumnas.AutoSize = False
        Me.TsBtnColumnas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnColumnas.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TsBtnColumnas.Image = CType(resources.GetObject("TsBtnColumnas.Image"), System.Drawing.Image)
        Me.TsBtnColumnas.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsBtnColumnas.Name = "TsBtnColumnas"
        Me.TsBtnColumnas.Size = New System.Drawing.Size(100, 28)
        Me.TsBtnColumnas.Text = "Columnas"
        Me.TsBtnColumnas.ToolTipText = "Gráficos Tipo Columnas"
        '
        'ToolStripLabel3
        '
        Me.ToolStripLabel3.AutoSize = False
        Me.ToolStripLabel3.Name = "ToolStripLabel3"
        Me.ToolStripLabel3.Size = New System.Drawing.Size(25, 28)
        '
        'TsBtnAreas
        '
        Me.TsBtnAreas.AutoSize = False
        Me.TsBtnAreas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnAreas.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TsBtnAreas.Image = CType(resources.GetObject("TsBtnAreas.Image"), System.Drawing.Image)
        Me.TsBtnAreas.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsBtnAreas.Name = "TsBtnAreas"
        Me.TsBtnAreas.Size = New System.Drawing.Size(100, 28)
        Me.TsBtnAreas.Text = "Areas"
        Me.TsBtnAreas.ToolTipText = "Gráficos Tipo Areas"
        '
        'ToolStripLabel4
        '
        Me.ToolStripLabel4.AutoSize = False
        Me.ToolStripLabel4.Name = "ToolStripLabel4"
        Me.ToolStripLabel4.Size = New System.Drawing.Size(25, 28)
        '
        'TsBtnLineas
        '
        Me.TsBtnLineas.AutoSize = False
        Me.TsBtnLineas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnLineas.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TsBtnLineas.Image = CType(resources.GetObject("TsBtnLineas.Image"), System.Drawing.Image)
        Me.TsBtnLineas.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsBtnLineas.Name = "TsBtnLineas"
        Me.TsBtnLineas.Size = New System.Drawing.Size(100, 28)
        Me.TsBtnLineas.Text = "Lineas"
        Me.TsBtnLineas.ToolTipText = "Gráficos Tipo Lineas"
        '
        'ToolStripLabel5
        '
        Me.ToolStripLabel5.AutoSize = False
        Me.ToolStripLabel5.Name = "ToolStripLabel5"
        Me.ToolStripLabel5.Size = New System.Drawing.Size(25, 28)
        '
        'TsBtnPastel
        '
        Me.TsBtnPastel.AutoSize = False
        Me.TsBtnPastel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnPastel.Font = New System.Drawing.Font("Arial Narrow", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TsBtnPastel.Image = CType(resources.GetObject("TsBtnPastel.Image"), System.Drawing.Image)
        Me.TsBtnPastel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TsBtnPastel.Name = "TsBtnPastel"
        Me.TsBtnPastel.Size = New System.Drawing.Size(100, 28)
        Me.TsBtnPastel.Text = "Pastel"
        Me.TsBtnPastel.ToolTipText = "Gráficos Tipo Pastel"
        '
        'ToolStripLabel2
        '
        Me.ToolStripLabel2.AutoSize = False
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        Me.ToolStripLabel2.Size = New System.Drawing.Size(25, 28)
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 31)
        '
        'Chart1
        '
        ChartArea1.Area3DStyle.Enable3D = True
        ChartArea1.Area3DStyle.IsClustered = True
        ChartArea1.AxisX.Interval = 1.0R
        ChartArea1.AxisX.IsLabelAutoFit = False
        ChartArea1.AxisX.LabelStyle.Angle = -45
        ChartArea1.AxisX.Title = "FECHAS"
        ChartArea1.AxisX.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.AxisX2.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.AxisY.Title = "EUROS"
        ChartArea1.AxisY.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        Me.Chart1.Location = New System.Drawing.Point(0, 31)
        Me.Chart1.Name = "Chart1"
        Series1.ChartArea = "ChartArea1"
        Series1.Color = System.Drawing.Color.Red
        Series1.IsValueShownAsLabel = True
        Series1.IsVisibleInLegend = False
        Series1.Legend = "Legend1"
        Series1.Name = "Gastos"
        Series1.SmartLabelStyle.AllowOutsidePlotArea = System.Windows.Forms.DataVisualization.Charting.LabelOutsidePlotAreaStyle.No
        Series2.ChartArea = "ChartArea1"
        Series2.Color = System.Drawing.Color.Blue
        Series2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!)
        Series2.IsValueShownAsLabel = True
        Series2.IsVisibleInLegend = False
        Series2.Legend = "Legend1"
        Series2.Name = "Ingresos"
        Series2.SmartLabelStyle.AllowOutsidePlotArea = System.Windows.Forms.DataVisualization.Charting.LabelOutsidePlotAreaStyle.No
        Me.Chart1.Series.Add(Series1)
        Me.Chart1.Series.Add(Series2)
        Me.Chart1.Size = New System.Drawing.Size(1100, 604)
        Me.Chart1.TabIndex = 3
        Me.Chart1.Text = "Chart1"
        Title1.Name = "Title1"
        Title1.Text = "Gráfico por Fechas 3D"
        Me.Chart1.Titles.Add(Title1)
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintPreviewDialog1
        '
        Me.PrintPreviewDialog1.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.PrintPreviewDialog1.ClientSize = New System.Drawing.Size(400, 300)
        Me.PrintPreviewDialog1.Enabled = True
        Me.PrintPreviewDialog1.Icon = CType(resources.GetObject("PrintPreviewDialog1.Icon"), System.Drawing.Icon)
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        Me.PrintPreviewDialog1.Visible = False
        '
        'TSBtnImprimir
        '
        Me.TSBtnImprimir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TSBtnImprimir.Image = CType(resources.GetObject("TSBtnImprimir.Image"), System.Drawing.Image)
        Me.TSBtnImprimir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.TSBtnImprimir.Name = "TSBtnImprimir"
        Me.TSBtnImprimir.Size = New System.Drawing.Size(29, 28)
        Me.TSBtnImprimir.Text = "ToolStripButton1"
        '
        'GraficosFechas3D
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1100, 635)
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "GraficosFechas3D"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Graficos"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As Windows.Forms.ToolStrip
    Friend WithEvents Chart1 As Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents ToolStripLabel2 As Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripSeparator2 As Windows.Forms.ToolStripSeparator
    Friend WithEvents TsBtnColumnas As Windows.Forms.ToolStripButton
    Friend WithEvents TsBtnAreas As Windows.Forms.ToolStripButton
    Friend WithEvents TsBtnLineas As Windows.Forms.ToolStripButton
    Friend WithEvents TsBtnPastel As Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripLabel3 As Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripLabel4 As Windows.Forms.ToolStripLabel
    Friend WithEvents ToolStripLabel5 As Windows.Forms.ToolStripLabel
    Friend WithEvents PrintDialog1 As Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As Windows.Forms.PrintPreviewDialog
    Friend WithEvents TSBtnImprimir As Windows.Forms.ToolStripButton
End Class
