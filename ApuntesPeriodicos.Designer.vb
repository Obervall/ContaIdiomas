<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ApuntesPeriodicos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ApuntesPeriodicos))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtNumRegistros = New System.Windows.Forms.TextBox()
        Me.LblNumRegistros = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnEliminaSeleccion = New System.Windows.Forms.Button()
        Me.BtnCalculadora = New System.Windows.Forms.Button()
        Me.BtnEditarRegistro = New System.Windows.Forms.Button()
        Me.BtnEliminarRegistro = New System.Windows.Forms.Button()
        Me.BtnAñadirRegistro = New System.Windows.Forms.Button()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.BtnGraficos = New System.Windows.Forms.Button()
        Me.BtnImprimir = New System.Windows.Forms.Button()
        Me.BtnSeguirBuscando = New System.Windows.Forms.Button()
        Me.BtnBuscarRegistro = New System.Windows.Forms.Button()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.BtnSinFiltroCuenta = New System.Windows.Forms.Button()
        Me.BtnFiltroCuenta = New System.Windows.Forms.Button()
        Me.CmbCuenta = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.TxtConcepto = New System.Windows.Forms.TextBox()
        Me.BtnSinFiltroConcepto = New System.Windows.Forms.Button()
        Me.BtnFiltroConcepto = New System.Windows.Forms.Button()
        Me.CmbConcepto = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.BtnSinFiltroFecha = New System.Windows.Forms.Button()
        Me.BtnFiltroFecha = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.TxtSaldo = New System.Windows.Forms.TextBox()
        Me.TxtGastos = New System.Windows.Forms.TextBox()
        Me.TxtIngresos = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.DgvApuper = New System.Windows.Forms.DataGridView()
        Me.GroupBox8 = New System.Windows.Forms.GroupBox()
        Me.BtnUltimo = New System.Windows.Forms.Button()
        Me.BtnSiguiente = New System.Windows.Forms.Button()
        Me.BtnAnterior = New System.Windows.Forms.Button()
        Me.BtnPrimero = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        CType(Me.DgvApuper, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox8.SuspendLayout()
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnEliminaSeleccion)
        Me.GroupBox2.Controls.Add(Me.BtnCalculadora)
        Me.GroupBox2.Controls.Add(Me.BtnEditarRegistro)
        Me.GroupBox2.Controls.Add(Me.BtnEliminarRegistro)
        Me.GroupBox2.Controls.Add(Me.BtnAñadirRegistro)
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
        'BtnCalculadora
        '
        resources.ApplyResources(Me.BtnCalculadora, "BtnCalculadora")
        Me.BtnCalculadora.Name = "BtnCalculadora"
        Me.BtnCalculadora.UseVisualStyleBackColor = True
        '
        'BtnEditarRegistro
        '
        resources.ApplyResources(Me.BtnEditarRegistro, "BtnEditarRegistro")
        Me.BtnEditarRegistro.Name = "BtnEditarRegistro"
        Me.BtnEditarRegistro.UseVisualStyleBackColor = True
        '
        'BtnEliminarRegistro
        '
        resources.ApplyResources(Me.BtnEliminarRegistro, "BtnEliminarRegistro")
        Me.BtnEliminarRegistro.Name = "BtnEliminarRegistro"
        Me.BtnEliminarRegistro.UseVisualStyleBackColor = True
        '
        'BtnAñadirRegistro
        '
        resources.ApplyResources(Me.BtnAñadirRegistro, "BtnAñadirRegistro")
        Me.BtnAñadirRegistro.Name = "BtnAñadirRegistro"
        Me.BtnAñadirRegistro.UseVisualStyleBackColor = True
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.BtnGraficos)
        Me.GroupBox3.Controls.Add(Me.BtnImprimir)
        Me.GroupBox3.Controls.Add(Me.BtnSeguirBuscando)
        Me.GroupBox3.Controls.Add(Me.BtnBuscarRegistro)
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
        'BtnSalir
        '
        resources.ApplyResources(Me.BtnSalir, "BtnSalir")
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.BtnSinFiltroCuenta)
        Me.GroupBox4.Controls.Add(Me.BtnFiltroCuenta)
        Me.GroupBox4.Controls.Add(Me.CmbCuenta)
        Me.GroupBox4.Controls.Add(Me.Label3)
        resources.ApplyResources(Me.GroupBox4, "GroupBox4")
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.TabStop = False
        '
        'BtnSinFiltroCuenta
        '
        resources.ApplyResources(Me.BtnSinFiltroCuenta, "BtnSinFiltroCuenta")
        Me.BtnSinFiltroCuenta.Name = "BtnSinFiltroCuenta"
        Me.BtnSinFiltroCuenta.UseVisualStyleBackColor = True
        '
        'BtnFiltroCuenta
        '
        resources.ApplyResources(Me.BtnFiltroCuenta, "BtnFiltroCuenta")
        Me.BtnFiltroCuenta.Name = "BtnFiltroCuenta"
        Me.BtnFiltroCuenta.UseVisualStyleBackColor = True
        '
        'CmbCuenta
        '
        Me.CmbCuenta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbCuenta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbCuenta.FormattingEnabled = True
        resources.ApplyResources(Me.CmbCuenta, "CmbCuenta")
        Me.CmbCuenta.Name = "CmbCuenta"
        '
        'Label3
        '
        resources.ApplyResources(Me.Label3, "Label3")
        Me.Label3.Name = "Label3"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.TxtConcepto)
        Me.GroupBox5.Controls.Add(Me.BtnSinFiltroConcepto)
        Me.GroupBox5.Controls.Add(Me.BtnFiltroConcepto)
        Me.GroupBox5.Controls.Add(Me.CmbConcepto)
        Me.GroupBox5.Controls.Add(Me.Label4)
        resources.ApplyResources(Me.GroupBox5, "GroupBox5")
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.TabStop = False
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
        'Label4
        '
        resources.ApplyResources(Me.Label4, "Label4")
        Me.Label4.Name = "Label4"
        '
        'GroupBox6
        '
        Me.GroupBox6.Controls.Add(Me.DateTimePicker2)
        Me.GroupBox6.Controls.Add(Me.DateTimePicker1)
        Me.GroupBox6.Controls.Add(Me.BtnSinFiltroFecha)
        Me.GroupBox6.Controls.Add(Me.BtnFiltroFecha)
        Me.GroupBox6.Controls.Add(Me.Label6)
        Me.GroupBox6.Controls.Add(Me.Label5)
        resources.ApplyResources(Me.GroupBox6, "GroupBox6")
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.TabStop = False
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        resources.ApplyResources(Me.DateTimePicker2, "DateTimePicker2")
        Me.DateTimePicker2.Name = "DateTimePicker2"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        resources.ApplyResources(Me.DateTimePicker1, "DateTimePicker1")
        Me.DateTimePicker1.Name = "DateTimePicker1"
        '
        'BtnSinFiltroFecha
        '
        resources.ApplyResources(Me.BtnSinFiltroFecha, "BtnSinFiltroFecha")
        Me.BtnSinFiltroFecha.Name = "BtnSinFiltroFecha"
        Me.BtnSinFiltroFecha.UseVisualStyleBackColor = True
        '
        'BtnFiltroFecha
        '
        resources.ApplyResources(Me.BtnFiltroFecha, "BtnFiltroFecha")
        Me.BtnFiltroFecha.Name = "BtnFiltroFecha"
        Me.BtnFiltroFecha.UseVisualStyleBackColor = True
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
        '
        'Label5
        '
        resources.ApplyResources(Me.Label5, "Label5")
        Me.Label5.Name = "Label5"
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
        'DgvApuper
        '
        Me.DgvApuper.AllowUserToAddRows = False
        Me.DgvApuper.AllowUserToDeleteRows = False
        Me.DgvApuper.AllowUserToResizeRows = False
        Me.DgvApuper.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        resources.ApplyResources(Me.DgvApuper, "DgvApuper")
        Me.DgvApuper.Name = "DgvApuper"
        Me.DgvApuper.ReadOnly = True
        Me.DgvApuper.RowHeadersVisible = False
        Me.DgvApuper.RowTemplate.Height = 24
        Me.DgvApuper.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvApuper.StandardTab = True
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
        'ApuntesPeriodicos
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox8)
        Me.Controls.Add(Me.DgvApuper)
        Me.Controls.Add(Me.GroupBox7)
        Me.Controls.Add(Me.GroupBox6)
        Me.Controls.Add(Me.GroupBox5)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ApuntesPeriodicos"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        CType(Me.DgvApuper, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox8.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents BtnSalir As Windows.Forms.Button
    Friend WithEvents TxtNumRegistros As Windows.Forms.TextBox
    Friend WithEvents LblNumRegistros As Windows.Forms.Label
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents BtnEditarRegistro As Windows.Forms.Button
    Friend WithEvents BtnEliminarRegistro As Windows.Forms.Button
    Friend WithEvents BtnAñadirRegistro As Windows.Forms.Button
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox5 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox6 As Windows.Forms.GroupBox
    Friend WithEvents GroupBox7 As Windows.Forms.GroupBox
    Friend WithEvents DgvApuper As Windows.Forms.DataGridView
    Friend WithEvents BtnCalculadora As Windows.Forms.Button
    Friend WithEvents BtnGraficos As Windows.Forms.Button
    Friend WithEvents BtnImprimir As Windows.Forms.Button
    Friend WithEvents BtnSeguirBuscando As Windows.Forms.Button
    Friend WithEvents BtnBuscarRegistro As Windows.Forms.Button
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label5 As Windows.Forms.Label
    Friend WithEvents Label9 As Windows.Forms.Label
    Friend WithEvents Label8 As Windows.Forms.Label
    Friend WithEvents Label7 As Windows.Forms.Label
    Friend WithEvents TxtSaldo As Windows.Forms.TextBox
    Friend WithEvents TxtGastos As Windows.Forms.TextBox
    Friend WithEvents TxtIngresos As Windows.Forms.TextBox
    Friend WithEvents BtnSinFiltroCuenta As Windows.Forms.Button
    Friend WithEvents BtnFiltroCuenta As Windows.Forms.Button
    Friend WithEvents CmbCuenta As Windows.Forms.ComboBox
    Friend WithEvents BtnSinFiltroConcepto As Windows.Forms.Button
    Friend WithEvents BtnFiltroConcepto As Windows.Forms.Button
    Friend WithEvents CmbConcepto As Windows.Forms.ComboBox
    Friend WithEvents DateTimePicker2 As Windows.Forms.DateTimePicker
    Friend WithEvents DateTimePicker1 As Windows.Forms.DateTimePicker
    Friend WithEvents BtnSinFiltroFecha As Windows.Forms.Button
    Friend WithEvents BtnFiltroFecha As Windows.Forms.Button
    Friend WithEvents TxtConcepto As Windows.Forms.TextBox
    Friend WithEvents GroupBox8 As Windows.Forms.GroupBox
    Friend WithEvents BtnUltimo As Windows.Forms.Button
    Friend WithEvents BtnSiguiente As Windows.Forms.Button
    Friend WithEvents BtnAnterior As Windows.Forms.Button
    Friend WithEvents BtnPrimero As Windows.Forms.Button
    Friend WithEvents BtnEliminaSeleccion As Windows.Forms.Button
End Class
