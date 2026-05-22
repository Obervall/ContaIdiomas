<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class TraspasoCuentas
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TraspasoCuentas))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.BtnCuentaOrigen = New System.Windows.Forms.Button()
        Me.BtnCalculadora = New System.Windows.Forms.Button()
        Me.BtnConcepto = New System.Windows.Forms.Button()
        Me.BtnHoyOrigen = New System.Windows.Forms.Button()
        Me.TxtTipoConcepto = New System.Windows.Forms.TextBox()
        Me.TxtDescripcion = New System.Windows.Forms.TextBox()
        Me.TxtImporte = New System.Windows.Forms.TextBox()
        Me.TxtNota = New System.Windows.Forms.TextBox()
        Me.CmbCuentaOrigen = New System.Windows.Forms.ComboBox()
        Me.CmbConcepto = New System.Windows.Forms.ComboBox()
        Me.DtpOrigen = New System.Windows.Forms.DateTimePicker()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.BtnAceptar = New System.Windows.Forms.Button()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.BtnCuentaDestino = New System.Windows.Forms.Button()
        Me.CmbCuentaDestino = New System.Windows.Forms.ComboBox()
        Me.DtpDestino = New System.Windows.Forms.DateTimePicker()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.BtnCuentaOrigen)
        Me.GroupBox1.Controls.Add(Me.BtnCalculadora)
        Me.GroupBox1.Controls.Add(Me.BtnConcepto)
        Me.GroupBox1.Controls.Add(Me.BtnHoyOrigen)
        Me.GroupBox1.Controls.Add(Me.TxtTipoConcepto)
        Me.GroupBox1.Controls.Add(Me.TxtDescripcion)
        Me.GroupBox1.Controls.Add(Me.TxtImporte)
        Me.GroupBox1.Controls.Add(Me.TxtNota)
        Me.GroupBox1.Controls.Add(Me.CmbCuentaOrigen)
        Me.GroupBox1.Controls.Add(Me.CmbConcepto)
        Me.GroupBox1.Controls.Add(Me.DtpOrigen)
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
        'BtnCuentaOrigen
        '
        resources.ApplyResources(Me.BtnCuentaOrigen, "BtnCuentaOrigen")
        Me.BtnCuentaOrigen.Name = "BtnCuentaOrigen"
        Me.BtnCuentaOrigen.UseVisualStyleBackColor = True
        '
        'BtnCalculadora
        '
        resources.ApplyResources(Me.BtnCalculadora, "BtnCalculadora")
        Me.BtnCalculadora.Name = "BtnCalculadora"
        Me.BtnCalculadora.UseVisualStyleBackColor = True
        '
        'BtnConcepto
        '
        resources.ApplyResources(Me.BtnConcepto, "BtnConcepto")
        Me.BtnConcepto.Name = "BtnConcepto"
        Me.BtnConcepto.UseVisualStyleBackColor = True
        '
        'BtnHoyOrigen
        '
        resources.ApplyResources(Me.BtnHoyOrigen, "BtnHoyOrigen")
        Me.BtnHoyOrigen.Name = "BtnHoyOrigen"
        Me.BtnHoyOrigen.UseVisualStyleBackColor = True
        '
        'TxtTipoConcepto
        '
        resources.ApplyResources(Me.TxtTipoConcepto, "TxtTipoConcepto")
        Me.TxtTipoConcepto.Name = "TxtTipoConcepto"
        '
        'TxtDescripcion
        '
        resources.ApplyResources(Me.TxtDescripcion, "TxtDescripcion")
        Me.TxtDescripcion.Name = "TxtDescripcion"
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
        'CmbCuentaOrigen
        '
        Me.CmbCuentaOrigen.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbCuentaOrigen.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbCuentaOrigen.FormattingEnabled = True
        resources.ApplyResources(Me.CmbCuentaOrigen, "CmbCuentaOrigen")
        Me.CmbCuentaOrigen.Name = "CmbCuentaOrigen"
        '
        'CmbConcepto
        '
        Me.CmbConcepto.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbConcepto.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbConcepto.FormattingEnabled = True
        resources.ApplyResources(Me.CmbConcepto, "CmbConcepto")
        Me.CmbConcepto.Name = "CmbConcepto"
        '
        'DtpOrigen
        '
        Me.DtpOrigen.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        resources.ApplyResources(Me.DtpOrigen, "DtpOrigen")
        Me.DtpOrigen.Name = "DtpOrigen"
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
        'BtnAceptar
        '
        resources.ApplyResources(Me.BtnAceptar, "BtnAceptar")
        Me.BtnAceptar.Name = "BtnAceptar"
        Me.BtnAceptar.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.BtnCuentaDestino)
        Me.GroupBox2.Controls.Add(Me.CmbCuentaDestino)
        Me.GroupBox2.Controls.Add(Me.DtpDestino)
        Me.GroupBox2.Controls.Add(Me.Label10)
        Me.GroupBox2.Controls.Add(Me.Label14)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'BtnCuentaDestino
        '
        resources.ApplyResources(Me.BtnCuentaDestino, "BtnCuentaDestino")
        Me.BtnCuentaDestino.Name = "BtnCuentaDestino"
        Me.BtnCuentaDestino.UseVisualStyleBackColor = True
        '
        'CmbCuentaDestino
        '
        Me.CmbCuentaDestino.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbCuentaDestino.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbCuentaDestino.FormattingEnabled = True
        resources.ApplyResources(Me.CmbCuentaDestino, "CmbCuentaDestino")
        Me.CmbCuentaDestino.Name = "CmbCuentaDestino"
        '
        'DtpDestino
        '
        resources.ApplyResources(Me.DtpDestino, "DtpDestino")
        Me.DtpDestino.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DtpDestino.Name = "DtpDestino"
        '
        'Label10
        '
        resources.ApplyResources(Me.Label10, "Label10")
        Me.Label10.Name = "Label10"
        '
        'Label14
        '
        resources.ApplyResources(Me.Label14, "Label14")
        Me.Label14.Name = "Label14"
        '
        'TraspasoCuentas
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.BtnAceptar)
        Me.Controls.Add(Me.BtnCancelar)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TraspasoCuentas"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
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
    Friend WithEvents BtnAceptar As Windows.Forms.Button
    Friend WithEvents DtpOrigen As Windows.Forms.DateTimePicker
    Friend WithEvents BtnCuentaOrigen As Windows.Forms.Button
    Friend WithEvents BtnCalculadora As Windows.Forms.Button
    Friend WithEvents BtnConcepto As Windows.Forms.Button
    Friend WithEvents BtnHoyOrigen As Windows.Forms.Button
    Friend WithEvents TxtTipoConcepto As Windows.Forms.TextBox
    Friend WithEvents TxtDescripcion As Windows.Forms.TextBox
    Friend WithEvents TxtImporte As Windows.Forms.TextBox
    Friend WithEvents TxtNota As Windows.Forms.TextBox
    Friend WithEvents CmbCuentaOrigen As Windows.Forms.ComboBox
    Friend WithEvents CmbConcepto As Windows.Forms.ComboBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents BtnCuentaDestino As Windows.Forms.Button
    Friend WithEvents CmbCuentaDestino As Windows.Forms.ComboBox
    Friend WithEvents DtpDestino As Windows.Forms.DateTimePicker
    Friend WithEvents Label10 As Windows.Forms.Label
    Friend WithEvents Label14 As Windows.Forms.Label
End Class
