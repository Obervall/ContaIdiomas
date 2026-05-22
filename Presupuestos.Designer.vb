<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Presupuestos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Presupuestos))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtNumRegistros = New System.Windows.Forms.TextBox()
        Me.LblNumRegistros = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.BtnGraficos = New System.Windows.Forms.Button()
        Me.BtnImprimir = New System.Windows.Forms.Button()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.TxtConcepto = New System.Windows.Forms.TextBox()
        Me.BtnSinFiltroConcepto = New System.Windows.Forms.Button()
        Me.BtnFiltroConcepto = New System.Windows.Forms.Button()
        Me.CmbConcepto = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DgvPresupuestos = New System.Windows.Forms.DataGridView()
        Me.LblDesviacion = New System.Windows.Forms.Label()
        Me.TxtDesviacion = New System.Windows.Forms.TextBox()
        Me.LblObjetivo = New System.Windows.Forms.Label()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.BtnUltimo = New System.Windows.Forms.Button()
        Me.BtnSiguiente = New System.Windows.Forms.Button()
        Me.BtnAnterior = New System.Windows.Forms.Button()
        Me.BtnPrimero = New System.Windows.Forms.Button()
        Me.BtnEliminarRegistro = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DgvPresupuestos, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtNumRegistros)
        Me.GroupBox1.Controls.Add(Me.LblNumRegistros)
        Me.GroupBox1.Controls.Add(Me.Label1)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'TxtNumRegistros
        '
        resources.ApplyResources(Me.TxtNumRegistros, "TxtNumRegistros")
        Me.TxtNumRegistros.Name = "TxtNumRegistros"
        '
        'LblNumRegistros
        '
        resources.ApplyResources(Me.LblNumRegistros, "LblNumRegistros")
        Me.LblNumRegistros.Name = "LblNumRegistros"
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.BtnGraficos)
        Me.GroupBox3.Controls.Add(Me.BtnImprimir)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'BtnGraficos
        '
        resources.ApplyResources(Me.BtnGraficos, "BtnGraficos")
        Me.BtnGraficos.Name = "BtnGraficos"
        Me.BtnGraficos.UseVisualStyleBackColor = True
        '
        'BtnImprimir
        '
        resources.ApplyResources(Me.BtnImprimir, "BtnImprimir")
        Me.BtnImprimir.Name = "BtnImprimir"
        Me.BtnImprimir.UseVisualStyleBackColor = True
        '
        'BtnSalir
        '
        resources.ApplyResources(Me.BtnSalir, "BtnSalir")
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.TxtConcepto)
        Me.GroupBox4.Controls.Add(Me.BtnSinFiltroConcepto)
        Me.GroupBox4.Controls.Add(Me.BtnFiltroConcepto)
        Me.GroupBox4.Controls.Add(Me.CmbConcepto)
        Me.GroupBox4.Controls.Add(Me.Label3)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'TxtConcepto
        '
        resources.ApplyResources(Me.TxtConcepto, "TxtConcepto")
        Me.TxtConcepto.Name = "TxtConcepto"
        '
        'BtnSinFiltroConcepto
        '
        resources.ApplyResources(Me.BtnSinFiltroConcepto, "BtnSinFiltroConcepto")
        Me.BtnSinFiltroConcepto.Name = "BtnSinFiltroConcepto"
        Me.BtnSinFiltroConcepto.UseVisualStyleBackColor = True
        '
        'BtnFiltroConcepto
        '
        resources.ApplyResources(Me.BtnFiltroConcepto, "BtnFiltroConcepto")
        Me.BtnFiltroConcepto.Name = "BtnFiltroConcepto"
        Me.BtnFiltroConcepto.UseVisualStyleBackColor = True
        '
        'CmbConcepto
        '
        Me.CmbConcepto.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbConcepto.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbConcepto.FormattingEnabled = True
        resources.ApplyResources(Me.CmbConcepto, "CmbConcepto")
        Me.CmbConcepto.Name = "CmbConcepto"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'DgvPresupuestos
        '
        Me.DgvPresupuestos.AllowUserToAddRows = False
        Me.DgvPresupuestos.AllowUserToDeleteRows = False
        Me.DgvPresupuestos.AllowUserToResizeRows = False
        Me.DgvPresupuestos.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        resources.ApplyResources(Me.DgvPresupuestos, "DgvPresupuestos")
        Me.DgvPresupuestos.Name = "DgvPresupuestos"
        Me.DgvPresupuestos.ReadOnly = True
        Me.DgvPresupuestos.RowHeadersVisible = False
        Me.DgvPresupuestos.RowTemplate.Height = 24
        Me.DgvPresupuestos.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'LblDesviacion
        '
        resources.ApplyResources(Me.LblDesviacion, "LblDesviacion")
        Me.LblDesviacion.Name = "LblDesviacion"
        '
        'TxtDesviacion
        '
        resources.ApplyResources(Me.TxtDesviacion, "TxtDesviacion")
        Me.TxtDesviacion.Name = "TxtDesviacion"
        '
        'LblObjetivo
        '
        resources.ApplyResources(Me.LblObjetivo, "LblObjetivo")
        Me.LblObjetivo.Name = "LblObjetivo"
        '
        'GroupBox8
        '
        Me.GroupBox8.Controls.Add(Me.BtnUltimo)
        Me.GroupBox8.Controls.Add(Me.BtnSiguiente)
        Me.GroupBox8.Controls.Add(Me.BtnAnterior)
        Me.GroupBox8.Controls.Add(Me.BtnPrimero)
        resources.ApplyResources(Me.GroupBox8, "GroupBox8")
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.TabStop = False
        '
        'BtnUltimo
        '
        resources.ApplyResources(Me.BtnUltimo, "BtnUltimo")
        Me.BtnUltimo.Name = "BtnUltimo"
        Me.BtnUltimo.UseVisualStyleBackColor = True
        '
        'BtnSiguiente
        '
        resources.ApplyResources(Me.BtnSiguiente, "BtnSiguiente")
        Me.BtnSiguiente.Name = "BtnSiguiente"
        Me.BtnSiguiente.UseVisualStyleBackColor = True
        '
        'BtnAnterior
        '
        resources.ApplyResources(Me.BtnAnterior, "BtnAnterior")
        Me.BtnAnterior.Name = "BtnAnterior"
        Me.BtnAnterior.UseVisualStyleBackColor = True
        '
        'BtnPrimero
        '
        resources.ApplyResources(Me.BtnPrimero, "BtnPrimero")
        Me.BtnPrimero.Name = "BtnPrimero"
        Me.BtnPrimero.UseVisualStyleBackColor = True
        '
        'BtnEliminarRegistro
        '
        resources.ApplyResources(Me.BtnEliminarRegistro, "BtnEliminarRegistro")
        Me.BtnEliminarRegistro.Name = "BtnEliminarRegistro"
        Me.BtnEliminarRegistro.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnEliminarRegistro)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
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
        'Presupuestos
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.LblObjetivo)
        Me.Controls.Add(Me.TxtDesviacion)
        Me.Controls.Add(Me.LblDesviacion)
        Me.Controls.Add(Me.DgvPresupuestos)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Presupuestos"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DgvPresupuestos, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents TxtNumRegistros As Windows.Forms.TextBox
    Friend WithEvents LblNumRegistros As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents BtnImprimir As Windows.Forms.Button
    Friend WithEvents BtnSalir As Windows.Forms.Button
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents BtnSinFiltroConcepto As Windows.Forms.Button
    Friend WithEvents BtnFiltroConcepto As Windows.Forms.Button
    Friend WithEvents CmbConcepto As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents DgvPresupuestos As Windows.Forms.DataGridView
    Friend WithEvents BtnGraficos As Windows.Forms.Button
    Friend WithEvents TxtConcepto As Windows.Forms.TextBox
    Friend WithEvents LblDesviacion As Windows.Forms.Label
    Friend WithEvents TxtDesviacion As Windows.Forms.TextBox
    Friend WithEvents LblObjetivo As Windows.Forms.Label
    Friend WithEvents GroupBox8 As Windows.Forms.GroupBox
    Friend WithEvents BtnUltimo As Windows.Forms.Button
    Friend WithEvents BtnSiguiente As Windows.Forms.Button
    Friend WithEvents BtnAnterior As Windows.Forms.Button
    Friend WithEvents BtnPrimero As Windows.Forms.Button
    Friend WithEvents BtnEliminarRegistro As Windows.Forms.Button
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents PrintDialog1 As Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As Windows.Forms.PrintPreviewDialog
End Class
