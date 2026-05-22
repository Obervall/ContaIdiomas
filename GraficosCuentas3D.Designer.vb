<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class GraficosCuentas3D
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(GraficosCuentas3D))
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
        Me.TSBtnImprimir = New System.Windows.Forms.ToolStripButton()
        Me.Chart1 = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.ToolStrip1.SuspendLayout()
        CType(Me.Chart1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TsBtnColumnas, Me.ToolStripLabel3, Me.TsBtnAreas, Me.ToolStripLabel4, Me.TsBtnLineas, Me.ToolStripLabel5, Me.TsBtnPastel, Me.ToolStripLabel2, Me.ToolStripSeparator2, Me.TSBtnImprimir})
        resources.ApplyResources(Me.ToolStrip1, "ToolStrip1")
        Me.ToolStrip1.Name = "ToolStrip1"
        '
        'TsBtnColumnas
        '
        resources.ApplyResources(Me.TsBtnColumnas, "TsBtnColumnas")
        Me.TsBtnColumnas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnColumnas.Name = "TsBtnColumnas"
        '
        'ToolStripLabel3
        '
        resources.ApplyResources(Me.ToolStripLabel3, "ToolStripLabel3")
        Me.ToolStripLabel3.Name = "ToolStripLabel3"
        '
        'TsBtnAreas
        '
        resources.ApplyResources(Me.TsBtnAreas, "TsBtnAreas")
        Me.TsBtnAreas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnAreas.Name = "TsBtnAreas"
        '
        'ToolStripLabel4
        '
        resources.ApplyResources(Me.ToolStripLabel4, "ToolStripLabel4")
        Me.ToolStripLabel4.Name = "ToolStripLabel4"
        '
        'TsBtnLineas
        '
        resources.ApplyResources(Me.TsBtnLineas, "TsBtnLineas")
        Me.TsBtnLineas.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnLineas.Name = "TsBtnLineas"
        '
        'ToolStripLabel5
        '
        resources.ApplyResources(Me.ToolStripLabel5, "ToolStripLabel5")
        Me.ToolStripLabel5.Name = "ToolStripLabel5"
        '
        'TsBtnPastel
        '
        resources.ApplyResources(Me.TsBtnPastel, "TsBtnPastel")
        Me.TsBtnPastel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.TsBtnPastel.Name = "TsBtnPastel"
        '
        'ToolStripLabel2
        '
        resources.ApplyResources(Me.ToolStripLabel2, "ToolStripLabel2")
        Me.ToolStripLabel2.Name = "ToolStripLabel2"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        resources.ApplyResources(Me.ToolStripSeparator2, "ToolStripSeparator2")
        '
        'TSBtnImprimir
        '
        Me.TSBtnImprimir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        resources.ApplyResources(Me.TSBtnImprimir, "TSBtnImprimir")
        Me.TSBtnImprimir.Name = "TSBtnImprimir"
        '
        'Chart1
        '
        ChartArea1.Area3DStyle.Enable3D = True
        ChartArea1.Area3DStyle.IsClustered = True
        ChartArea1.AxisX.Interval = 1.0R
        ChartArea1.AxisX.IsLabelAutoFit = False
        ChartArea1.AxisX.LabelStyle.Angle = -45
        ChartArea1.AxisX.Title = "CUENTAS"
        ChartArea1.AxisX.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.AxisX2.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.AxisY.Title = "EUROS"
        ChartArea1.AxisY.TitleFont = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        ChartArea1.Name = "ChartArea1"
        Me.Chart1.ChartAreas.Add(ChartArea1)
        Legend1.Name = "Legend1"
        Me.Chart1.Legends.Add(Legend1)
        resources.ApplyResources(Me.Chart1, "Chart1")
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
        Title1.Name = "Title1"
        Title1.Text = "Gráfico por Cuentas Bancarias 3D"
        Me.Chart1.Titles.Add(Title1)
        '
        'PrintDialog1
        '
        Me.PrintDialog1.UseEXDialog = True
        '
        'PrintDocument1
        '
        '
        'PrintPreviewDialog1
        '
        resources.ApplyResources(Me.PrintPreviewDialog1, "PrintPreviewDialog1")
        Me.PrintPreviewDialog1.Name = "PrintPreviewDialog1"
        '
        'GraficosCuentas3D
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Chart1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "GraficosCuentas3D"
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
    Friend WithEvents TSBtnImprimir As Windows.Forms.ToolStripButton
    Friend WithEvents PrintDialog1 As Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As Windows.Forms.PrintPreviewDialog
End Class
