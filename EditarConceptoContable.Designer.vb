<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class EditarConceptoContable
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(EditarConceptoContable))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.TxtTipoConcepto = New System.Windows.Forms.TextBox()
        Me.LblEditando = New System.Windows.Forms.Label()
        Me.TxtDescripcion = New System.Windows.Forms.TextBox()
        Me.TxtNombre = New System.Windows.Forms.TextBox()
        Me.TxtNota = New System.Windows.Forms.TextBox()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.BtnAceptar = New System.Windows.Forms.Button()
        Me.BtnEliminar = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.TxtTipoConcepto)
        Me.GroupBox1.Controls.Add(Me.LblEditando)
        Me.GroupBox1.Controls.Add(Me.TxtDescripcion)
        Me.GroupBox1.Controls.Add(Me.TxtNombre)
        Me.GroupBox1.Controls.Add(Me.TxtNota)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'TxtTipoConcepto
        '
        resources.ApplyResources(Me.TxtTipoConcepto, "TxtTipoConcepto")
        Me.TxtTipoConcepto.Name = "TxtTipoConcepto"
        '
        'LblEditando
        '
        Me.LblEditando.ForeColor = System.Drawing.Color.IndianRed
        resources.ApplyResources(Me.LblEditando, "LblEditando")
        Me.LblEditando.Name = "LblEditando"
        '
        'TxtDescripcion
        '
        resources.ApplyResources(Me.TxtDescripcion, "TxtDescripcion")
        Me.TxtDescripcion.Name = "TxtDescripcion"
        '
        'TxtNombre
        '
        resources.ApplyResources(Me.TxtNombre, "TxtNombre")
        Me.TxtNombre.Name = "TxtNombre"
        '
        'TxtNota
        '
        resources.ApplyResources(Me.TxtNota, "TxtNota")
        Me.TxtNota.Name = "TxtNota"
        '
        'Label6
        '
        resources.ApplyResources(Me.Label6, "Label6")
        Me.Label6.Name = "Label6"
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
        'BtnEliminar
        '
        resources.ApplyResources(Me.BtnEliminar, "BtnEliminar")
        Me.BtnEliminar.Name = "BtnEliminar"
        Me.BtnEliminar.UseVisualStyleBackColor = True
        '
        'EditarConceptoContable
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.BtnEliminar)
        Me.Controls.Add(Me.BtnAceptar)
        Me.Controls.Add(Me.BtnCancelar)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "EditarConceptoContable"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents Label6 As Windows.Forms.Label
    Friend WithEvents Label4 As Windows.Forms.Label
    Friend WithEvents Label3 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents BtnCancelar As Windows.Forms.Button
    Friend WithEvents BtnAceptar As Windows.Forms.Button
    Friend WithEvents TxtDescripcion As Windows.Forms.TextBox
    Friend WithEvents TxtNombre As Windows.Forms.TextBox
    Friend WithEvents TxtNota As Windows.Forms.TextBox
    Friend WithEvents LblEditando As Windows.Forms.Label
    Friend WithEvents BtnEliminar As Windows.Forms.Button
    Friend WithEvents TxtTipoConcepto As Windows.Forms.TextBox
End Class
