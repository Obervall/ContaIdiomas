<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AyudaApuntes
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(AyudaApuntes))
        Me.MiCuadroTextoAyuda = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'MiCuadroTextoAyuda
        '
        resources.ApplyResources(Me.MiCuadroTextoAyuda, "MiCuadroTextoAyuda")
        Me.MiCuadroTextoAyuda.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.MiCuadroTextoAyuda.Name = "MiCuadroTextoAyuda"
        Me.MiCuadroTextoAyuda.ReadOnly = True
        '
        'AyudaApuntes
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.MiCuadroTextoAyuda)
        Me.Name = "AyudaApuntes"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents MiCuadroTextoAyuda As Windows.Forms.RichTextBox
End Class
