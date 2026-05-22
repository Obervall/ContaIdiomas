<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CuentasBancarias
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(CuentasBancarias))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtNumRegistros = New System.Windows.Forms.TextBox()
        Me.LblNumRegistros = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnEliminaSeleccion = New System.Windows.Forms.Button()
        Me.BtnEditarRegistro = New System.Windows.Forms.Button()
        Me.BtnAñadirRegistro = New System.Windows.Forms.Button()
        Me.BtnEliminarRegistro = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.BtnImprimir = New System.Windows.Forms.Button()
        Me.BtnSeguirBuscando = New System.Windows.Forms.Button()
        Me.BtnBuscarRegistro = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtnSinFiltroTipoCuenta = New System.Windows.Forms.Button()
        Me.BtnFiltroTipoCuenta = New System.Windows.Forms.Button()
        Me.CmbTipoCuenta = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.DgvCuentas = New System.Windows.Forms.DataGridView()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.BtnUltimo = New System.Windows.Forms.Button()
        Me.BtnSiguiente = New System.Windows.Forms.Button()
        Me.BtnAnterior = New System.Windows.Forms.Button()
        Me.BtnPrimero = New System.Windows.Forms.Button()
        Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
        Me.LblCuentas = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.TxtSaldo = New System.Windows.Forms.TextBox()
        Me.TxtGastos = New System.Windows.Forms.TextBox()
        Me.TxtIngresos = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        CType(Me.DgvCuentas, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox8.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
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
        Me.TxtNumRegistros.ForeColor = System.Drawing.SystemColors.WindowText
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnEliminaSeleccion)
        Me.GroupBox2.Controls.Add(Me.BtnEditarRegistro)
        Me.GroupBox2.Controls.Add(Me.BtnAñadirRegistro)
        Me.GroupBox2.Controls.Add(Me.BtnEliminarRegistro)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'BtnEliminaSeleccion
        '
        resources.ApplyResources(Me.BtnEliminaSeleccion, "BtnEliminaSeleccion")
        Me.BtnEliminaSeleccion.Name = "BtnEliminaSeleccion"
        Me.BtnEliminaSeleccion.UseVisualStyleBackColor = True
        '
        'BtnEditarRegistro
        '
        resources.ApplyResources(Me.BtnEditarRegistro, "BtnEditarRegistro")
        Me.BtnEditarRegistro.Name = "BtnEditarRegistro"
        Me.BtnEditarRegistro.UseVisualStyleBackColor = True
        '
        'BtnAñadirRegistro
        '
        resources.ApplyResources(Me.BtnAñadirRegistro, "BtnAñadirRegistro")
        Me.BtnAñadirRegistro.Name = "BtnAñadirRegistro"
        Me.BtnAñadirRegistro.UseVisualStyleBackColor = True
        '
        'BtnEliminarRegistro
        '
        resources.ApplyResources(Me.BtnEliminarRegistro, "BtnEliminarRegistro")
        Me.BtnEliminarRegistro.Name = "BtnEliminarRegistro"
        Me.BtnEliminarRegistro.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.BtnImprimir)
        Me.GroupBox3.Controls.Add(Me.BtnSeguirBuscando)
        Me.GroupBox3.Controls.Add(Me.BtnBuscarRegistro)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'BtnImprimir
        '
        resources.ApplyResources(Me.BtnImprimir, "BtnImprimir")
        Me.BtnImprimir.Name = "BtnImprimir"
        Me.BtnImprimir.UseVisualStyleBackColor = True
        '
        'BtnSeguirBuscando
        '
        resources.ApplyResources(Me.BtnSeguirBuscando, "BtnSeguirBuscando")
        Me.BtnSeguirBuscando.Name = "BtnSeguirBuscando"
        Me.BtnSeguirBuscando.UseVisualStyleBackColor = True
        '
        'BtnBuscarRegistro
        '
        resources.ApplyResources(Me.BtnBuscarRegistro, "BtnBuscarRegistro")
        Me.BtnBuscarRegistro.Name = "BtnBuscarRegistro"
        Me.BtnBuscarRegistro.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.BtnSinFiltroTipoCuenta)
        Me.GroupBox4.Controls.Add(Me.BtnFiltroTipoCuenta)
        Me.GroupBox4.Controls.Add(Me.CmbTipoCuenta)
        Me.GroupBox4.Controls.Add(Me.Label3)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'BtnSinFiltroTipoCuenta
        '
        resources.ApplyResources(Me.BtnSinFiltroTipoCuenta, "BtnSinFiltroTipoCuenta")
        Me.BtnSinFiltroTipoCuenta.Name = "BtnSinFiltroTipoCuenta"
        Me.BtnSinFiltroTipoCuenta.UseVisualStyleBackColor = True
        '
        'BtnFiltroTipoCuenta
        '
        resources.ApplyResources(Me.BtnFiltroTipoCuenta, "BtnFiltroTipoCuenta")
        Me.BtnFiltroTipoCuenta.Name = "BtnFiltroTipoCuenta"
        Me.BtnFiltroTipoCuenta.UseVisualStyleBackColor = True
        '
        'CmbTipoCuenta
        '
        Me.CmbTipoCuenta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbTipoCuenta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbTipoCuenta.FormattingEnabled = True
        resources.ApplyResources(Me.CmbTipoCuenta, "CmbTipoCuenta")
        Me.CmbTipoCuenta.Name = "CmbTipoCuenta"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'DgvCuentas
        '
        Me.DgvCuentas.AllowUserToAddRows = False
        Me.DgvCuentas.AllowUserToDeleteRows = False
        Me.DgvCuentas.AllowUserToResizeRows = False
        Me.DgvCuentas.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        resources.ApplyResources(Me.DgvCuentas, "DgvCuentas")
        Me.DgvCuentas.Name = "DgvCuentas"
        Me.DgvCuentas.ReadOnly = True
        Me.DgvCuentas.RowHeadersVisible = False
        Me.DgvCuentas.RowTemplate.Height = 24
        Me.DgvCuentas.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        '
        'BtnSalir
        '
        resources.ApplyResources(Me.BtnSalir, "BtnSalir")
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.UseVisualStyleBackColor = True
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
        'LblCuentas
        '
        resources.ApplyResources(Me.LblCuentas, "LblCuentas")
        Me.LblCuentas.Name = "LblCuentas"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.TxtSaldo)
        Me.GroupBox7.Controls.Add(Me.TxtGastos)
        Me.GroupBox7.Controls.Add(Me.TxtIngresos)
        Me.GroupBox7.Controls.Add(Me.Label9)
        Me.GroupBox7.Controls.Add(Me.Label8)
        Me.GroupBox7.Controls.Add(Me.Label7)
        resources.ApplyResources(Me.GroupBox7, "GroupBox7")
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.TabStop = False
        '
        'TxtSaldo
        '
        resources.ApplyResources(Me.TxtSaldo, "TxtSaldo")
        Me.TxtSaldo.Name = "TxtSaldo"
        '
        'TxtGastos
        '
        resources.ApplyResources(Me.TxtGastos, "TxtGastos")
        Me.TxtGastos.Name = "TxtGastos"
        '
        'TxtIngresos
        '
        resources.ApplyResources(Me.TxtIngresos, "TxtIngresos")
        Me.TxtIngresos.Name = "TxtIngresos"
        '
        'Label9
        '
        resources.ApplyResources(Me.Label9, "Label9")
        Me.Label9.Name = "Label9"
        '
        'Label8
        '
        resources.ApplyResources(Me.Label8, "Label8")
        Me.Label8.Name = "Label8"
        '
        'Label7
        '
        resources.ApplyResources(Me.Label7, "Label7")
        Me.Label7.Name = "Label7"
        '
        'CuentasBancarias
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.LblCuentas)
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.DgvCuentas)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "CuentasBancarias"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        CType(Me.DgvCuentas, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox8.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents TxtNumRegistros As Windows.Forms.TextBox
    Friend WithEvents LblNumRegistros As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents CmbTipoCuenta As Windows.Forms.ComboBox
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents DgvCuentas As Windows.Forms.DataGridView
    Friend WithEvents BtnAñadirRegistro As Windows.Forms.Button
    Friend WithEvents BtnEditarRegistro As Windows.Forms.Button
    Friend WithEvents BtnEliminarRegistro As Windows.Forms.Button
    Friend WithEvents BtnImprimir As Windows.Forms.Button
    Friend WithEvents BtnSeguirBuscando As Windows.Forms.Button
    Friend WithEvents BtnBuscarRegistro As Windows.Forms.Button
    Friend WithEvents BtnSinFiltroTipoCuenta As Windows.Forms.Button
    Friend WithEvents BtnFiltroTipoCuenta As Windows.Forms.Button
    Friend WithEvents BtnSalir As Windows.Forms.Button
    Friend WithEvents GroupBox8 As Windows.Forms.GroupBox
    Friend WithEvents BtnUltimo As Windows.Forms.Button
    Friend WithEvents BtnSiguiente As Windows.Forms.Button
    Friend WithEvents BtnAnterior As Windows.Forms.Button
    Friend WithEvents BtnPrimero As Windows.Forms.Button
    Friend WithEvents PrintDialog1 As Windows.Forms.PrintDialog
    Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
    Friend WithEvents PrintPreviewDialog1 As Windows.Forms.PrintPreviewDialog
    Friend WithEvents LblCuentas As Windows.Forms.Label
    Friend WithEvents BtnEliminaSeleccion As Windows.Forms.Button
    Friend WithEvents GroupBox7 As Windows.Forms.GroupBox
    Friend WithEvents TxtSaldo As Windows.Forms.TextBox
    Friend WithEvents TxtGastos As Windows.Forms.TextBox
    Friend WithEvents TxtIngresos As Windows.Forms.TextBox
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
End Class
