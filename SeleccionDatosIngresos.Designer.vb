<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SeleccionDatosIngresos
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SeleccionDatosIngresos))
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.BtnTodos = New System.Windows.Forms.Button()
        Me.BtnContinuar = New System.Windows.Forms.Button()
        Me.BtnContinuar3D = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        resources.ApplyResources(Me.ListBox1, "ListBox1")
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
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
        'BtnTodos
        '
        resources.ApplyResources(Me.BtnTodos, "BtnTodos")
        Me.BtnTodos.Name = "BtnTodos"
        Me.BtnTodos.UseVisualStyleBackColor = True
        '
        'BtnContinuar
        '
        resources.ApplyResources(Me.BtnContinuar, "BtnContinuar")
        Me.BtnContinuar.Name = "BtnContinuar"
        Me.BtnContinuar.UseVisualStyleBackColor = True
        '
        'BtnContinuar3D
        '
        resources.ApplyResources(Me.BtnContinuar3D, "BtnContinuar3D")
        Me.BtnContinuar3D.Name = "BtnContinuar3D"
        Me.BtnContinuar3D.UseVisualStyleBackColor = True
        '
        'SeleccionDatosIngresos
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.BtnContinuar3D)
        Me.Controls.Add(Me.BtnContinuar)
        Me.Controls.Add(Me.BtnTodos)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ListBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SeleccionDatosIngresos"
        Me.ShowIcon = False
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ListBox1 As Windows.Forms.ListBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents BtnTodos As Windows.Forms.Button
    Friend WithEvents BtnContinuar As Windows.Forms.Button
    Friend WithEvents BtnContinuar3D As Windows.Forms.Button
End Class
