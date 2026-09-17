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
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FiltroEvolutivo))
		Me.NudDia = New System.Windows.Forms.NumericUpDown()
		Me.NudMes = New System.Windows.Forms.NumericUpDown()
		Me.CmbCuenta = New System.Windows.Forms.ComboBox()
		Me.BtnAceptar = New System.Windows.Forms.Button()
		Me.CmbConcepto = New System.Windows.Forms.ComboBox()
		Me.PrintDialog1 = New System.Windows.Forms.PrintDialog()
		Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
		Me.PrintPreviewDialog1 = New System.Windows.Forms.PrintPreviewDialog()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label4 = New System.Windows.Forms.Label()
		CType(Me.NudDia, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.NudMes, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'NudDia
		'
		resources.ApplyResources(Me.NudDia, "NudDia")
		Me.NudDia.Name = "NudDia"
		'
		'NudMes
		'
		resources.ApplyResources(Me.NudMes, "NudMes")
		Me.NudMes.Name = "NudMes"
		'
		'CmbCuenta
		'
		Me.CmbCuenta.FormattingEnabled = True
		resources.ApplyResources(Me.CmbCuenta, "CmbCuenta")
		Me.CmbCuenta.Name = "CmbCuenta"
		'
		'BtnAceptar
		'
		resources.ApplyResources(Me.BtnAceptar, "BtnAceptar")
		Me.BtnAceptar.Name = "BtnAceptar"
		Me.BtnAceptar.UseVisualStyleBackColor = True
		'
		'CmbConcepto
		'
		Me.CmbConcepto.FormattingEnabled = True
		resources.ApplyResources(Me.CmbConcepto, "CmbConcepto")
		Me.CmbConcepto.Name = "CmbConcepto"
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
		'Label1
		'
		resources.ApplyResources(Me.Label1, "Label1")
		Me.Label1.Name = "Label1"
		'
		'Label2
		'
		resources.ApplyResources(Me.Label2, "Label2")
		Me.Label2.Name = "Label2"
		'
		'Label3
		'
		resources.ApplyResources(Me.Label3, "Label3")
		Me.Label3.Name = "Label3"
		'
		'Label4
		'
		resources.ApplyResources(Me.Label4, "Label4")
		Me.Label4.Name = "Label4"
		'
		'FiltroEvolutivo
		'
		resources.ApplyResources(Me, "$this")
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.CmbConcepto)
		Me.Controls.Add(Me.BtnAceptar)
		Me.Controls.Add(Me.CmbCuenta)
		Me.Controls.Add(Me.NudMes)
		Me.Controls.Add(Me.NudDia)
		Me.Name = "FiltroEvolutivo"
		CType(Me.NudDia, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.NudMes, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents NudDia As Windows.Forms.NumericUpDown
	Friend WithEvents NudMes As Windows.Forms.NumericUpDown
	Friend WithEvents CmbCuenta As Windows.Forms.ComboBox
	Friend WithEvents BtnAceptar As Windows.Forms.Button
	Friend WithEvents CmbConcepto As Windows.Forms.ComboBox
	Friend WithEvents PrintDialog1 As Windows.Forms.PrintDialog
	Friend WithEvents PrintDocument1 As Drawing.Printing.PrintDocument
	Friend WithEvents PrintPreviewDialog1 As Windows.Forms.PrintPreviewDialog
	Friend WithEvents Label1 As Windows.Forms.Label
	Friend WithEvents Label2 As Windows.Forms.Label
	Friend WithEvents Label3 As Windows.Forms.Label
	Friend WithEvents Label4 As Windows.Forms.Label
End Class
