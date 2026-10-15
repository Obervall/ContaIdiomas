<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Cabecera
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
		Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Cabecera))
		Me.DgvUsuarios = New System.Windows.Forms.DataGridView()
		Me.LblTitulo = New System.Windows.Forms.Label()
		Me.BtnSalir = New System.Windows.Forms.Button()
		CType(Me.DgvUsuarios, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'DgvUsuarios
		'
		Me.DgvUsuarios.AllowUserToAddRows = False
		Me.DgvUsuarios.AllowUserToDeleteRows = False
		Me.DgvUsuarios.AllowUserToResizeRows = False
		DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
		DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
		DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
		DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Info
		DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
		DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
		Me.DgvUsuarios.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
		Me.DgvUsuarios.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DgvUsuarios.Location = New System.Drawing.Point(12, 153)
		Me.DgvUsuarios.Name = "DgvUsuarios"
		Me.DgvUsuarios.ReadOnly = True
		Me.DgvUsuarios.RowHeadersWidth = 51
		Me.DgvUsuarios.RowTemplate.Height = 24
		Me.DgvUsuarios.Size = New System.Drawing.Size(383, 186)
		Me.DgvUsuarios.TabIndex = 0
		'
		'LblTitulo
		'
		Me.LblTitulo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
		Me.LblTitulo.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblTitulo.Location = New System.Drawing.Point(13, 120)
		Me.LblTitulo.Name = "LblTitulo"
		Me.LblTitulo.Size = New System.Drawing.Size(382, 30)
		Me.LblTitulo.TabIndex = 1
		Me.LblTitulo.Text = "Num. EMPRESAS / USUARIOS"
		Me.LblTitulo.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'BtnSalir
		'
		Me.BtnSalir.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.BtnSalir.Image = CType(resources.GetObject("BtnSalir.Image"), System.Drawing.Image)
		Me.BtnSalir.ImageAlign = System.Drawing.ContentAlignment.TopCenter
		Me.BtnSalir.Location = New System.Drawing.Point(268, 12)
		Me.BtnSalir.Name = "BtnSalir"
		Me.BtnSalir.Size = New System.Drawing.Size(127, 90)
		Me.BtnSalir.TabIndex = 2
		Me.BtnSalir.Text = "Salir"
		Me.BtnSalir.TextAlign = System.Drawing.ContentAlignment.BottomCenter
		Me.BtnSalir.UseVisualStyleBackColor = True
		'
		'Cabecera
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(410, 352)
		Me.ControlBox = False
		Me.Controls.Add(Me.BtnSalir)
		Me.Controls.Add(Me.LblTitulo)
		Me.Controls.Add(Me.DgvUsuarios)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "Cabecera"
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
		Me.Text = "Empresas / Usuarios"
		CType(Me.DgvUsuarios, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents DgvUsuarios As Windows.Forms.DataGridView
	Friend WithEvents LblTitulo As Windows.Forms.Label
	Friend WithEvents BtnSalir As Windows.Forms.Button
End Class
