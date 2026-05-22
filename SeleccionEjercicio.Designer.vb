<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SeleccionEjercicio
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SeleccionEjercicio))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.CmbEjercicio = New System.Windows.Forms.ComboBox()
        Me.BtnCrearNuevo = New System.Windows.Forms.Button()
        Me.BtnAceptar = New System.Windows.Forms.Button()
        Me.BtnCancelar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'CmbEjercicio
        '
        Me.CmbEjercicio.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.CmbEjercicio.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.CmbEjercicio.FormattingEnabled = True
        resources.ApplyResources(Me.CmbEjercicio, "CmbEjercicio")
        Me.CmbEjercicio.Name = "CmbEjercicio"
        '
        'BtnCrearNuevo
        '
        resources.ApplyResources(Me.BtnCrearNuevo, "BtnCrearNuevo")
        Me.BtnCrearNuevo.Name = "BtnCrearNuevo"
        Me.BtnCrearNuevo.UseVisualStyleBackColor = True
        '
        'BtnAceptar
        '
        resources.ApplyResources(Me.BtnAceptar, "BtnAceptar")
        Me.BtnAceptar.Name = "BtnAceptar"
        Me.BtnAceptar.UseVisualStyleBackColor = True
        '
        'BtnCancelar
        '
        resources.ApplyResources(Me.BtnCancelar, "BtnCancelar")
        Me.BtnCancelar.Name = "BtnCancelar"
        Me.BtnCancelar.UseVisualStyleBackColor = True
        '
        'SeleccionEjercicio
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.BtnCancelar)
        Me.Controls.Add(Me.BtnAceptar)
        Me.Controls.Add(Me.BtnCrearNuevo)
        Me.Controls.Add(Me.CmbEjercicio)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "SeleccionEjercicio"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents CmbEjercicio As Windows.Forms.ComboBox
    Friend WithEvents BtnCrearNuevo As Windows.Forms.Button
    Friend WithEvents BtnAceptar As Windows.Forms.Button
    Friend WithEvents BtnCancelar As Windows.Forms.Button
End Class
