<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AprendizajeBancario
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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AprendizajeBancario))
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.BtnAyuda = New System.Windows.Forms.Button()
		Me.TxtBuscarLetras = New System.Windows.Forms.TextBox()
		Me.Label8 = New System.Windows.Forms.Label()
		Me.TxtDescripcion = New System.Windows.Forms.TextBox()
		Me.LblBuscarLetras = New System.Windows.Forms.Label()
		Me.BtnDescripcion = New System.Windows.Forms.Button()
		Me.CmbDescripcion = New System.Windows.Forms.ComboBox()
		Me.BtnConcepto = New System.Windows.Forms.Button()
		Me.TxtTipoConcepto = New System.Windows.Forms.TextBox()
		Me.TxtImporte = New System.Windows.Forms.TextBox()
		Me.TxtNota = New System.Windows.Forms.TextBox()
		Me.CmbCuenta = New System.Windows.Forms.ComboBox()
		Me.CmbConcepto = New System.Windows.Forms.ComboBox()
		Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
		Me.Label7 = New System.Windows.Forms.Label()
		Me.Label6 = New System.Windows.Forms.Label()
		Me.Label5 = New System.Windows.Forms.Label()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.BtnCancelar = New System.Windows.Forms.Button()
		Me.BtnSaltarApuntes = New System.Windows.Forms.Button()
		Me.BtnAceptarOtro = New System.Windows.Forms.Button()
		Me.GroupBox1.SuspendLayout()
		Me.SuspendLayout()
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.BtnAyuda)
		Me.GroupBox1.Controls.Add(Me.TxtBuscarLetras)
		Me.GroupBox1.Controls.Add(Me.Label8)
		Me.GroupBox1.Controls.Add(Me.TxtDescripcion)
		Me.GroupBox1.Controls.Add(Me.LblBuscarLetras)
		Me.GroupBox1.Controls.Add(Me.BtnDescripcion)
		Me.GroupBox1.Controls.Add(Me.CmbDescripcion)
		Me.GroupBox1.Controls.Add(Me.BtnConcepto)
		Me.GroupBox1.Controls.Add(Me.TxtTipoConcepto)
		Me.GroupBox1.Controls.Add(Me.TxtImporte)
		Me.GroupBox1.Controls.Add(Me.TxtNota)
		Me.GroupBox1.Controls.Add(Me.CmbCuenta)
		Me.GroupBox1.Controls.Add(Me.CmbConcepto)
		Me.GroupBox1.Controls.Add(Me.DateTimePicker1)
		Me.GroupBox1.Controls.Add(Me.Label7)
		Me.GroupBox1.Controls.Add(Me.Label6)
		Me.GroupBox1.Controls.Add(Me.Label5)
		Me.GroupBox1.Controls.Add(Me.Label4)
		Me.GroupBox1.Controls.Add(Me.Label3)
		Me.GroupBox1.Controls.Add(Me.Label2)
		Me.GroupBox1.Controls.Add(Me.Label1)
		resources.ApplyResources(Me.GroupBox1, "GroupBox1")
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.TabStop = False
		'
		'BtnAyuda
		'
		resources.ApplyResources(Me.BtnAyuda, "BtnAyuda")
		Me.BtnAyuda.Name = "BtnAyuda"
		Me.BtnAyuda.TabStop = False
		Me.BtnAyuda.UseVisualStyleBackColor = True
		'
		'TxtBuscarLetras
		'
		resources.ApplyResources(Me.TxtBuscarLetras, "TxtBuscarLetras")
		Me.TxtBuscarLetras.Name = "TxtBuscarLetras"
		'
		'Label8
		'
		resources.ApplyResources(Me.Label8, "Label8")
		Me.Label8.Name = "Label8"
		'
		'TxtDescripcion
		'
		resources.ApplyResources(Me.TxtDescripcion, "TxtDescripcion")
		Me.TxtDescripcion.Name = "TxtDescripcion"
		'
		'LblBuscarLetras
		'
		resources.ApplyResources(Me.LblBuscarLetras, "LblBuscarLetras")
		Me.LblBuscarLetras.Name = "LblBuscarLetras"
		'
		'BtnDescripcion
		'
		resources.ApplyResources(Me.BtnDescripcion, "BtnDescripcion")
		Me.BtnDescripcion.Name = "BtnDescripcion"
		Me.BtnDescripcion.UseVisualStyleBackColor = True
		'
		'CmbDescripcion
		'
		Me.CmbDescripcion.FormattingEnabled = True
		resources.ApplyResources(Me.CmbDescripcion, "CmbDescripcion")
		Me.CmbDescripcion.Name = "CmbDescripcion"
		'
		'BtnConcepto
		'
		resources.ApplyResources(Me.BtnConcepto, "BtnConcepto")
		Me.BtnConcepto.Name = "BtnConcepto"
		Me.BtnConcepto.UseVisualStyleBackColor = True
		'
		'TxtTipoConcepto
		'
		resources.ApplyResources(Me.TxtTipoConcepto, "TxtTipoConcepto")
		Me.TxtTipoConcepto.Name = "TxtTipoConcepto"
		'
		'TxtImporte
		'
		resources.ApplyResources(Me.TxtImporte, "TxtImporte")
		Me.TxtImporte.Name = "TxtImporte"
		'
		'TxtNota
		'
		resources.ApplyResources(Me.TxtNota, "TxtNota")
		Me.TxtNota.Name = "TxtNota"
		'
		'CmbCuenta
		'
		Me.CmbCuenta.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
		Me.CmbCuenta.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
		resources.ApplyResources(Me.CmbCuenta, "CmbCuenta")
		Me.CmbCuenta.FormattingEnabled = True
		Me.CmbCuenta.Name = "CmbCuenta"
		'
		'CmbConcepto
		'
		Me.CmbConcepto.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
		Me.CmbConcepto.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
		Me.CmbConcepto.FormattingEnabled = True
		resources.ApplyResources(Me.CmbConcepto, "CmbConcepto")
		Me.CmbConcepto.Name = "CmbConcepto"
		'
		'DateTimePicker1
		'
		resources.ApplyResources(Me.DateTimePicker1, "DateTimePicker1")
		Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
		Me.DateTimePicker1.Name = "DateTimePicker1"
		'
		'Label7
		'
		resources.ApplyResources(Me.Label7, "Label7")
		Me.Label7.Name = "Label7"
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
		'Label4
		'
		resources.ApplyResources(Me.Label4, "Label4")
		Me.Label4.Name = "Label4"
		'
		'Label3
		'
		resources.ApplyResources(Me.Label3, "Label3")
		Me.Label3.Name = "Label3"
		'
		'Label2
		'
		resources.ApplyResources(Me.Label2, "Label2")
		Me.Label2.Name = "Label2"
		'
		'Label1
		'
		resources.ApplyResources(Me.Label1, "Label1")
		Me.Label1.Name = "Label1"
		'
		'BtnCancelar
		'
		resources.ApplyResources(Me.BtnCancelar, "BtnCancelar")
		Me.BtnCancelar.Name = "BtnCancelar"
		Me.BtnCancelar.UseVisualStyleBackColor = True
		'
		'BtnSaltarApuntes
		'
		resources.ApplyResources(Me.BtnSaltarApuntes, "BtnSaltarApuntes")
		Me.BtnSaltarApuntes.Name = "BtnSaltarApuntes"
		Me.BtnSaltarApuntes.UseVisualStyleBackColor = True
		'
		'BtnAceptarOtro
		'
		resources.ApplyResources(Me.BtnAceptarOtro, "BtnAceptarOtro")
		Me.BtnAceptarOtro.Name = "BtnAceptarOtro"
		Me.BtnAceptarOtro.UseVisualStyleBackColor = True
		'
		'AprendizajeBancario
		'
		resources.ApplyResources(Me, "$this")
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Controls.Add(Me.BtnAceptarOtro)
		Me.Controls.Add(Me.BtnSaltarApuntes)
		Me.Controls.Add(Me.BtnCancelar)
		Me.Controls.Add(Me.GroupBox1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "AprendizajeBancario"
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
	Friend WithEvents Label7 As Windows.Forms.Label
	Friend WithEvents Label6 As Windows.Forms.Label
	Friend WithEvents Label5 As Windows.Forms.Label
	Friend WithEvents Label4 As Windows.Forms.Label
	Friend WithEvents Label3 As Windows.Forms.Label
	Friend WithEvents Label2 As Windows.Forms.Label
	Friend WithEvents Label1 As Windows.Forms.Label
	Friend WithEvents BtnCancelar As Windows.Forms.Button
	Friend WithEvents BtnSaltarApuntes As Windows.Forms.Button
	Friend WithEvents BtnAceptarOtro As Windows.Forms.Button
	Friend WithEvents DateTimePicker1 As Windows.Forms.DateTimePicker
	Friend WithEvents BtnConcepto As Windows.Forms.Button
	Friend WithEvents TxtTipoConcepto As Windows.Forms.TextBox
	Friend WithEvents TxtImporte As Windows.Forms.TextBox
	Friend WithEvents TxtNota As Windows.Forms.TextBox
	Friend WithEvents CmbCuenta As Windows.Forms.ComboBox
	Friend WithEvents CmbConcepto As Windows.Forms.ComboBox
	Friend WithEvents CmbDescripcion As Windows.Forms.ComboBox
	Friend WithEvents BtnDescripcion As Windows.Forms.Button
	Friend WithEvents LblBuscarLetras As Windows.Forms.Label
	Friend WithEvents Label8 As Windows.Forms.Label
	Friend WithEvents TxtDescripcion As Windows.Forms.TextBox
	Friend WithEvents TxtBuscarLetras As Windows.Forms.TextBox
	Friend WithEvents BtnAyuda As Windows.Forms.Button
End Class
