<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ImprimirForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ImprimirForm))
        Me.LblUsuario = New System.Windows.Forms.Label()
        Me.TxtUsuario = New System.Windows.Forms.TextBox()
        Me.LblTitulo = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.LineaTop = New System.Windows.Forms.Label()
        Me.Punto5 = New System.Windows.Forms.Label()
        Me.Punto4 = New System.Windows.Forms.Label()
        Me.Punto3 = New System.Windows.Forms.Label()
        Me.Punto2 = New System.Windows.Forms.Label()
        Me.Punto1 = New System.Windows.Forms.Label()
        Me.LineaFondo = New System.Windows.Forms.Label()
        Me.DgvApuntes = New System.Windows.Forms.DataGridView()
        Me.LblFecha = New System.Windows.Forms.Label()
        Me.LblTotal = New System.Windows.Forms.Label()
        Me.LblNumeroPagina = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.LblEntreFechas = New System.Windows.Forms.Label()
        CType(Me.DgvApuntes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'LblUsuario
        '
        Me.LblUsuario.AutoSize = True
        Me.LblUsuario.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblUsuario.Location = New System.Drawing.Point(137, 49)
        Me.LblUsuario.Name = "LblUsuario"
        Me.LblUsuario.Size = New System.Drawing.Size(80, 20)
        Me.LblUsuario.TabIndex = 0
        Me.LblUsuario.Text = "Usuario:"
        '
        'TxtUsuario
        '
        Me.TxtUsuario.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtUsuario.Location = New System.Drawing.Point(223, 46)
        Me.TxtUsuario.Name = "TxtUsuario"
        Me.TxtUsuario.Size = New System.Drawing.Size(140, 27)
        Me.TxtUsuario.TabIndex = 1
        '
        'LblTitulo
        '
        Me.LblTitulo.AutoSize = True
        Me.LblTitulo.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTitulo.Location = New System.Drawing.Point(45, 107)
        Me.LblTitulo.Name = "LblTitulo"
        Me.LblTitulo.Size = New System.Drawing.Size(409, 25)
        Me.LblTitulo.TabIndex = 3
        Me.LblTitulo.Text = "Listado de Apuntes ordenado por Fechas"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(832, 683)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 20)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Pag # "
        '
        'LineaTop
        '
        Me.LineaTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LineaTop.ForeColor = System.Drawing.Color.Blue
        Me.LineaTop.Location = New System.Drawing.Point(47, 170)
        Me.LineaTop.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LineaTop.Name = "LineaTop"
        Me.LineaTop.Size = New System.Drawing.Size(894, 27)
        Me.LineaTop.TabIndex = 45
        Me.LineaTop.Text = "---------------------------------------------------------------------------------" &
    "--------------------------------------------------------------------------------" &
    "-----------------"
        '
        'Punto5
        '
        Me.Punto5.AutoSize = True
        Me.Punto5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Punto5.ForeColor = System.Drawing.Color.Blue
        Me.Punto5.Location = New System.Drawing.Point(912, 197)
        Me.Punto5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Punto5.Name = "Punto5"
        Me.Punto5.Size = New System.Drawing.Size(29, 20)
        Me.Punto5.TabIndex = 44
        Me.Punto5.Text = "P5"
        '
        'Punto4
        '
        Me.Punto4.AutoSize = True
        Me.Punto4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Punto4.ForeColor = System.Drawing.Color.Blue
        Me.Punto4.Location = New System.Drawing.Point(766, 197)
        Me.Punto4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Punto4.Name = "Punto4"
        Me.Punto4.Size = New System.Drawing.Size(29, 20)
        Me.Punto4.TabIndex = 43
        Me.Punto4.Text = "P4"
        '
        'Punto3
        '
        Me.Punto3.AutoSize = True
        Me.Punto3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Punto3.ForeColor = System.Drawing.Color.Blue
        Me.Punto3.Location = New System.Drawing.Point(490, 197)
        Me.Punto3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Punto3.Name = "Punto3"
        Me.Punto3.Size = New System.Drawing.Size(29, 20)
        Me.Punto3.TabIndex = 42
        Me.Punto3.Text = "P3"
        '
        'Punto2
        '
        Me.Punto2.AutoSize = True
        Me.Punto2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Punto2.ForeColor = System.Drawing.Color.Blue
        Me.Punto2.Location = New System.Drawing.Point(249, 197)
        Me.Punto2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Punto2.Name = "Punto2"
        Me.Punto2.Size = New System.Drawing.Size(29, 20)
        Me.Punto2.TabIndex = 41
        Me.Punto2.Text = "P2"
        '
        'Punto1
        '
        Me.Punto1.AutoSize = True
        Me.Punto1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Punto1.ForeColor = System.Drawing.Color.Blue
        Me.Punto1.Location = New System.Drawing.Point(61, 197)
        Me.Punto1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Punto1.Name = "Punto1"
        Me.Punto1.Size = New System.Drawing.Size(29, 20)
        Me.Punto1.TabIndex = 40
        Me.Punto1.Text = "P1"
        '
        'LineaFondo
        '
        Me.LineaFondo.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LineaFondo.ForeColor = System.Drawing.Color.Blue
        Me.LineaFondo.Location = New System.Drawing.Point(47, 371)
        Me.LineaFondo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LineaFondo.Name = "LineaFondo"
        Me.LineaFondo.Size = New System.Drawing.Size(894, 22)
        Me.LineaFondo.TabIndex = 46
        Me.LineaFondo.Text = "---------------------------------------------------------------------------------" &
    "--------------------------------------------------------------------------------" &
    "-----------------"
        '
        'DgvApuntes
        '
        Me.DgvApuntes.AllowUserToAddRows = False
        Me.DgvApuntes.AllowUserToDeleteRows = False
        Me.DgvApuntes.AllowUserToResizeRows = False
        Me.DgvApuntes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DgvApuntes.Location = New System.Drawing.Point(50, 230)
        Me.DgvApuntes.Name = "DgvApuntes"
        Me.DgvApuntes.ReadOnly = True
        Me.DgvApuntes.RowHeadersVisible = False
        Me.DgvApuntes.RowHeadersWidth = 51
        Me.DgvApuntes.RowTemplate.Height = 24
        Me.DgvApuntes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DgvApuntes.Size = New System.Drawing.Size(891, 138)
        Me.DgvApuntes.StandardTab = True
        Me.DgvApuntes.TabIndex = 47
        '
        'LblFecha
        '
        Me.LblFecha.AutoSize = True
        Me.LblFecha.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblFecha.Location = New System.Drawing.Point(459, 49)
        Me.LblFecha.Name = "LblFecha"
        Me.LblFecha.Size = New System.Drawing.Size(60, 20)
        Me.LblFecha.TabIndex = 48
        Me.LblFecha.Text = "Fecha"
        '
        'LblTotal
        '
        Me.LblTotal.AutoSize = True
        Me.LblTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblTotal.Location = New System.Drawing.Point(884, 415)
        Me.LblTotal.Name = "LblTotal"
        Me.LblTotal.Size = New System.Drawing.Size(57, 20)
        Me.LblTotal.TabIndex = 49
        Me.LblTotal.Text = "Total:"
        '
        'LblNumeroPagina
        '
        Me.LblNumeroPagina.AutoSize = True
        Me.LblNumeroPagina.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblNumeroPagina.Location = New System.Drawing.Point(901, 683)
        Me.LblNumeroPagina.Name = "LblNumeroPagina"
        Me.LblNumeroPagina.Size = New System.Drawing.Size(19, 20)
        Me.LblNumeroPagina.TabIndex = 50
        Me.LblNumeroPagina.Text = "0"
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(36, 25)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(95, 60)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 51
        Me.PictureBox1.TabStop = False
        '
        'LblEntreFechas
        '
        Me.LblEntreFechas.AutoSize = True
        Me.LblEntreFechas.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LblEntreFechas.Location = New System.Drawing.Point(304, 145)
        Me.LblEntreFechas.Name = "LblEntreFechas"
        Me.LblEntreFechas.Size = New System.Drawing.Size(150, 25)
        Me.LblEntreFechas.TabIndex = 52
        Me.LblEntreFechas.Text = "Desde: Hasta:"
        '
        'ImprimirForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1046, 794)
        Me.Controls.Add(Me.LblEntreFechas)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.LblNumeroPagina)
        Me.Controls.Add(Me.LblTotal)
        Me.Controls.Add(Me.LblFecha)
        Me.Controls.Add(Me.DgvApuntes)
        Me.Controls.Add(Me.LineaFondo)
        Me.Controls.Add(Me.LineaTop)
        Me.Controls.Add(Me.Punto5)
        Me.Controls.Add(Me.Punto4)
        Me.Controls.Add(Me.Punto3)
        Me.Controls.Add(Me.Punto2)
        Me.Controls.Add(Me.Punto1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.LblTitulo)
        Me.Controls.Add(Me.TxtUsuario)
        Me.Controls.Add(Me.LblUsuario)
        Me.Name = "ImprimirForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ImprimirForm"
        CType(Me.DgvApuntes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents LblUsuario As Windows.Forms.Label
    Friend WithEvents TxtUsuario As Windows.Forms.TextBox
    Friend WithEvents LblTitulo As Windows.Forms.Label
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents LineaTop As Windows.Forms.Label
    Friend WithEvents Punto5 As Windows.Forms.Label
    Friend WithEvents Punto4 As Windows.Forms.Label
    Friend WithEvents Punto3 As Windows.Forms.Label
    Friend WithEvents Punto2 As Windows.Forms.Label
    Friend WithEvents Punto1 As Windows.Forms.Label
    Friend WithEvents LineaFondo As Windows.Forms.Label
    Friend WithEvents DgvApuntes As Windows.Forms.DataGridView
    Friend WithEvents LblFecha As Windows.Forms.Label
    Friend WithEvents LblTotal As Windows.Forms.Label
    Friend WithEvents LblNumeroPagina As Windows.Forms.Label
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents LblEntreFechas As Windows.Forms.Label
End Class
