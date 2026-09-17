Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml.Drawing.Charts
Imports Microsoft.Office.Interop.Excel

Public Class FiltroEvolutivo

    Public Property TipoInforme As String ' "CUENTAS" o "CONCEPTOS"
    Private dtDatosInforme As System.Data.DataTable

    Private Sub BtnAceptar_Click(sender As Object, e As EventArgs) Handles BtnAceptar.Click
        ' 1. Recuperar los datos de la BD (.mdb) usando el SQL parametrizado
        'dtDatosInforme = ObtenerDatosEvolutivos(CInt(CmbElemento.SelectedValue), CInt(NudMes.Value), CInt(NudDia.Value))

        ' 2. Lanzar el PrintDocument o el PrintPreviewDialog
        Dim pd As New Printing.PrintDocument()
        AddHandler pd.PrintPage, AddressOf ImprimirInformeEvolutivo

        Dim preview As New PrintPreviewDialog()
        preview.Document = pd
        preview.ShowDialog()

        Me.Close()
    End Sub

    Private ReadOnly Property ObtenerDatosEvolutivos(v1 As Integer, v2 As Integer, v3 As Integer) As DocumentFormat.OpenXml.Drawing.Charts.DataTable
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private Sub ImprimirInformeEvolutivo(sender As Object, e As Printing.PrintPageEventArgs)

        Dim fuenteEjes As New System.Drawing.Font("Arial", 12, FontStyle.Bold)
        Dim fuenteTitulo As New System.Drawing.Font("Arial", 16, FontStyle.Bold)
        Dim fuenteTexto As New System.Drawing.Font("Arial", 11, FontStyle.Regular)
        Dim y As Single = 50

        ' Dibujar Encabezado
        e.Graphics.DrawString("EVOLUCIÓN DE SALDOS A FECHA: " & NudDia.Value & "/" & NudMes.Value, fuenteTitulo, Brushes.Black, 50, y)
        y += 50

		' Dibujar las columnas de la tabla (Año | Saldo)
		For Each row As DataRow In dtDatosInforme.Rows
			Dim anio As String = row("Anio").ToString()
			' Formateamos el importe como moneda limpia
			Dim importe As String = String.Format("{0:C2}", row(1))

			e.Graphics.DrawString(anio, fuenteTexto, Brushes.Black, 100, y)
			e.Graphics.DrawString(importe, fuenteTexto, Brushes.Black, 250, y)
			y += 25
		Next
	End Sub
End Class