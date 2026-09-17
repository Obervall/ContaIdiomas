<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FiltroEvolutivo
	Inherits System.Windows.Forms.Form

	'Form reemplaza a Dispose para limpiar la lista de componentes.
	<System.Diagnostics.DebuggerNonUserCode()> _
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
	<System.Diagnostics.DebuggerStepThrough()> _
	Private Sub InitializeComponent()
		Me.NudDia = New System.Windows.Forms.NumericUpDown()
		Me.NudMes = New System.Windows.Forms.NumericUpDown()
		Me.CmbElemento = New System.Windows.Forms.ComboBox()
		Me.BtnAceptar = New System.Windows.Forms.Button()
		CType(Me.NudDia, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.NudMes, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'NudDia
		'
		Me.NudDia.Location = New System.Drawing.Point(65, 49)
		Me.NudDia.Name = "NudDia"
		Me.NudDia.Size = New System.Drawing.Size(185, 22)
		Me.NudDia.TabIndex = 0
		'
		'NudMes
		'
		Me.NudMes.Location = New System.Drawing.Point(60, 110)
		Me.NudMes.Name = "NudMes"
		Me.NudMes.Size = New System.Drawing.Size(189, 22)
		Me.NudMes.TabIndex = 1
		'
		'CmbElemento
		'
		Me.CmbElemento.FormattingEnabled = True
		Me.CmbElemento.Location = New System.Drawing.Point(58, 169)
		Me.CmbElemento.Name = "CmbElemento"
		Me.CmbElemento.Size = New System.Drawing.Size(234, 24)
		Me.CmbElemento.TabIndex = 2
		'
		'BtnAceptar
		'
		Me.BtnAceptar.Location = New System.Drawing.Point(122, 259)
		Me.BtnAceptar.Name = "BtnAceptar"
		Me.BtnAceptar.Size = New System.Drawing.Size(216, 44)
		Me.BtnAceptar.TabIndex = 3
		Me.BtnAceptar.Text = "Button1"
		Me.BtnAceptar.UseVisualStyleBackColor = True
		'
		'FiltroEvolutivo
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.BtnAceptar)
		Me.Controls.Add(Me.CmbElemento)
		Me.Controls.Add(Me.NudMes)
		Me.Controls.Add(Me.NudDia)
		Me.Name = "FiltroEvolutivo"
		Me.Text = "Filtro Evolutivo"
		CType(Me.NudDia, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.NudMes, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents NudDia As Windows.Forms.NumericUpDown
	Friend WithEvents NudMes As Windows.Forms.NumericUpDown
	Friend WithEvents CmbElemento As Windows.Forms.ComboBox
	Friend WithEvents BtnAceptar As Windows.Forms.Button
End Class
