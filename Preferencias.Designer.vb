<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Preferencias
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Preferencias))
        Me.GBoxPantalla = New System.Windows.Forms.GroupBox()
        Me.ChbPantallaCompleta = New System.Windows.Forms.CheckBox()
        Me.ChbPantallaCierre = New System.Windows.Forms.CheckBox()
        Me.BtnSalir = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.TxtBaseDatos = New System.Windows.Forms.TextBox()
        Me.BtnUbicacion = New System.Windows.Forms.Button()
        Me.ChkBuscarActualizacion = New System.Windows.Forms.CheckBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.BtnBuscarBackup = New System.Windows.Forms.Button()
        Me.TxtPathExportar = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.BtnBuscarActualizacion = New System.Windows.Forms.Button()
        Me.CmbElegirIdioma = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.CmbMonedas = New System.Windows.Forms.ComboBox()
        Me.GBoxPantalla.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'GBoxPantalla
        '
        Me.GBoxPantalla.Controls.Add(Me.ChbPantallaCompleta)
        Me.GBoxPantalla.Controls.Add(Me.ChbPantallaCierre)
        resources.ApplyResources(Me.GBoxPantalla, "GBoxPantalla")
        Me.GBoxPantalla.Name = "GBoxPantalla"
        Me.GBoxPantalla.TabStop = False
        '
        'ChbPantallaCompleta
        '
        resources.ApplyResources(Me.ChbPantallaCompleta, "ChbPantallaCompleta")
        Me.ChbPantallaCompleta.Name = "ChbPantallaCompleta"
        Me.ChbPantallaCompleta.UseVisualStyleBackColor = True
        '
        'ChbPantallaCierre
        '
        resources.ApplyResources(Me.ChbPantallaCierre, "ChbPantallaCierre")
        Me.ChbPantallaCierre.Name = "ChbPantallaCierre"
        Me.ChbPantallaCierre.UseVisualStyleBackColor = True
        '
        'BtnSalir
        '
        resources.ApplyResources(Me.BtnSalir, "BtnSalir")
        Me.BtnSalir.Name = "BtnSalir"
        Me.BtnSalir.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.CheckBox3)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        resources.ApplyResources(Me.GroupBox1, "GroupBox1")
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.TabStop = False
        '
        'CheckBox3
        '
        resources.ApplyResources(Me.CheckBox3, "CheckBox3")
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        resources.ApplyResources(Me.CheckBox2, "CheckBox2")
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox1
        '
        resources.ApplyResources(Me.CheckBox1, "CheckBox1")
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.TxtBaseDatos)
        Me.GroupBox2.Controls.Add(Me.BtnUbicacion)
        resources.ApplyResources(Me.GroupBox2, "GroupBox2")
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.TabStop = False
        '
        'TxtBaseDatos
        '
        resources.ApplyResources(Me.TxtBaseDatos, "TxtBaseDatos")
        Me.TxtBaseDatos.Name = "TxtBaseDatos"
        '
        'BtnUbicacion
        '
        resources.ApplyResources(Me.BtnUbicacion, "BtnUbicacion")
        Me.BtnUbicacion.Name = "BtnUbicacion"
        Me.BtnUbicacion.UseVisualStyleBackColor = True
        '
        'ChkBuscarActualizacion
        '
        resources.ApplyResources(Me.ChkBuscarActualizacion, "ChkBuscarActualizacion")
        Me.ChkBuscarActualizacion.Name = "ChkBuscarActualizacion"
        Me.ChkBuscarActualizacion.UseVisualStyleBackColor = True
        '
        'Label1
        '
        resources.ApplyResources(Me.Label1, "Label1")
        Me.Label1.Name = "Label1"
        '
        'BtnBuscarBackup
        '
        resources.ApplyResources(Me.BtnBuscarBackup, "BtnBuscarBackup")
        Me.BtnBuscarBackup.Name = "BtnBuscarBackup"
        Me.BtnBuscarBackup.UseVisualStyleBackColor = True
        '
        'TxtPathExportar
        '
        resources.ApplyResources(Me.TxtPathExportar, "TxtPathExportar")
        Me.TxtPathExportar.Name = "TxtPathExportar"
        '
        'BtnBuscarActualizacion
        '
        resources.ApplyResources(Me.BtnBuscarActualizacion, "BtnBuscarActualizacion")
        Me.BtnBuscarActualizacion.Name = "BtnBuscarActualizacion"
        Me.BtnBuscarActualizacion.UseVisualStyleBackColor = True
        '
        'CmbElegirIdioma
        '
        Me.CmbElegirIdioma.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.CmbElegirIdioma, "CmbElegirIdioma")
        Me.CmbElegirIdioma.FormattingEnabled = True
        Me.CmbElegirIdioma.Items.AddRange(New Object() {resources.GetString("CmbElegirIdioma.Items"), resources.GetString("CmbElegirIdioma.Items1"), resources.GetString("CmbElegirIdioma.Items2"), resources.GetString("CmbElegirIdioma.Items3"), resources.GetString("CmbElegirIdioma.Items4"), resources.GetString("CmbElegirIdioma.Items5")})
        Me.CmbElegirIdioma.Name = "CmbElegirIdioma"
        '
        'Label2
        '
        resources.ApplyResources(Me.Label2, "Label2")
        Me.Label2.Name = "Label2"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.TxtPathExportar)
        Me.GroupBox3.Controls.Add(Me.BtnBuscarBackup)
        resources.ApplyResources(Me.GroupBox3, "GroupBox3")
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.TabStop = False
        '
        'CmbMonedas
        '
        Me.CmbMonedas.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        resources.ApplyResources(Me.CmbMonedas, "CmbMonedas")
        Me.CmbMonedas.FormattingEnabled = True
        Me.CmbMonedas.Items.AddRange(New Object() {resources.GetString("CmbMonedas.Items"), resources.GetString("CmbMonedas.Items1"), resources.GetString("CmbMonedas.Items2"), resources.GetString("CmbMonedas.Items3"), resources.GetString("CmbMonedas.Items4"), resources.GetString("CmbMonedas.Items5")})
        Me.CmbMonedas.Name = "CmbMonedas"
        '
        'Preferencias
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        resources.ApplyResources(Me, "$this")
        Me.Controls.Add(Me.CmbMonedas)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmbElegirIdioma)
        Me.Controls.Add(Me.BtnBuscarActualizacion)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.ChkBuscarActualizacion)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnSalir)
        Me.Controls.Add(Me.GBoxPantalla)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Preferencias"
        Me.GBoxPantalla.ResumeLayout(False)
        Me.GBoxPantalla.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents GBoxPantalla As Windows.Forms.GroupBox
    Friend WithEvents BtnSalir As Windows.Forms.Button
    Friend WithEvents ChbPantallaCompleta As Windows.Forms.CheckBox
    Friend WithEvents ChbPantallaCierre As Windows.Forms.CheckBox
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents CheckBox3 As Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As Windows.Forms.CheckBox
    Friend WithEvents CheckBox1 As Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents TxtBaseDatos As Windows.Forms.TextBox
    Friend WithEvents BtnUbicacion As Windows.Forms.Button
    Friend WithEvents ChkBuscarActualizacion As Windows.Forms.CheckBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents BtnBuscarBackup As Windows.Forms.Button
    Friend WithEvents TxtPathExportar As Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As Windows.Forms.FolderBrowserDialog
    Friend WithEvents BtnBuscarActualizacion As Windows.Forms.Button
    Friend WithEvents CmbElegirIdioma As Windows.Forms.ComboBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents CmbMonedas As Windows.Forms.ComboBox
End Class
