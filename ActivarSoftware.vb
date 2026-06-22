Imports System.Management
Imports System.Security.Cryptography
Imports System.Text

Public Class ActivarSoftware
    Dim CpuInfo As String = String.Empty
    Dim StrMotherBoardId As String = String.Empty
    Public vIdMaquina, vTxtPropietario As String
    Public rmse As New System.ComponentModel.ComponentResourceManager(Me.GetType())

    Private Sub ActivarSoftware_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True

        TxtPropietario.Text = ""
        StrMotherBoardId = ""

        Dim query As New SelectQuery("Win32_BaseBoard")
        Dim search As New ManagementObjectSearcher(query)

        Try
            For Each info As ManagementObject In search.Get()
                If info("SerialNumber") IsNot Nothing Then
                    StrMotherBoardId = info("SerialNumber").ToString().Trim()
                End If
            Next

            Dim idUpper As String = StrMotherBoardId.ToUpper()

            ' Filtramos "To be filled by O.E.M." y otros valores basura habituales
            If String.IsNullOrEmpty(idUpper) OrElse
           idUpper.Length <= 5 OrElse
           idUpper.Contains("TO BE FILLED") OrElse
           idUpper.Contains("NONE") OrElse
           idUpper.Contains("DEFAULT") OrElse
           idUpper.Contains("00000000") Then

                Dim mc As New ManagementClass("win32_processor")
                Dim moc As ManagementObjectCollection = mc.GetInstances()

                For Each mo As ManagementObject In moc
                    If mo.Properties("processorID").Value IsNot Nothing Then
                        CpuInfo = mo.Properties("processorID").Value.ToString().Trim()
                        Exit For
                    End If
                Next
                vIdMaquina = CpuInfo
            Else
                vIdMaquina = StrMotherBoardId
            End If

        Catch ex As Exception
            ' Mantenemos tu objeto rmse original para el mensaje de error
            MsgBox(rmse.GetString("ErrorObtenerID") & ": " & Environment.NewLine & ex.Message, MsgBoxStyle.Exclamation, resManager.GetString("Error"))
        End Try

        'Private Sub ActivarSoftware_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '    Me.KeyPreview = True

        '    TxtPropietario.Text = ""
        '    Dim query As New SelectQuery("Win32_BaseBoard")
        '    Dim search As New ManagementObjectSearcher(query)
        '    Dim info As ManagementObject
        '    Try
        '        For Each info In search.Get()
        '            StrMotherBoardId = info("SerialNumber").ToString()
        '        Next
        '        vIdMaquina = StrMotherBoardId
        '        If Len(vIdMaquina) <= 5 Then
        '            Dim mc As New ManagementClass("win32_processor")
        '            Dim moc As ManagementObjectCollection = mc.GetInstances
        '            For Each mo As ManagementObject In moc
        '                If CpuInfo = "" Then
        '                    CpuInfo = mo.Properties("processorID").Value.ToString
        '                    Exit For
        '                End If
        '            Next
        '            vIdMaquina = CpuInfo
        '        End If
        '    Catch ex As Exception
        '        MsgBox(rmse.GetString("ErrorObtenerID") & ": " & NL & ex.Message, MsgBoxStyle.Exclamation, resManager.GetString("Error"))
        '    End Try
        vTxtPropietario = My.Settings.Autorizar
        TxtPropietario.Text = Mid(My.Settings.Autorizar, 41)
        If vTxtPropietario = "Se autoriza el uso de ContaHogar 3.0 a: Modo Demo" Then
            MsgBox(frmPrincipal.rmse.GetString("LblAvisoActivacion"), MsgBoxStyle.Exclamation, rmse.GetString("ModoDemo"))
            IdMaquina.Text = vIdMaquina
        Else
            TxtPropietario.Text = My.Settings.Autorizar
            TxtPropietario.Enabled = False
            IdMaquina.Text = vIdMaquina
            TxtCodigoActivacion.Text = My.Settings.Codigo
            BtnActivar.Text = "ACTIVADO"
            BtnActivar.Enabled = False
        End If
    End Sub

    Public Shared Function GenerarSerial(ByVal InputString) As String
        Dim sha512 As SHA512 = SHA512Managed.Create()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes(InputString)
        Dim hash As Byte() = sha512.ComputeHash(bytes)
        Dim stringBuilder As New StringBuilder()
        For i As Integer = 0 To (hash.Length - 1) / 3
            stringBuilder.Append(hash(i).ToString("X2"))
        Next
        Return stringBuilder.ToString()
    End Function

    Private Sub BtnActivar_Click(sender As Object, e As EventArgs) Handles BtnActivar.Click
        If TxtCodigoActivacion.Text = GenerarSerial(TxtPropietario.Text & IdMaquina.Text & "LLkjhMN1") Then
            My.Settings.Autorizar = "Se autoriza el uso a: " & TxtPropietario.Text
            My.Settings.Codigo = "Codigo Activación: " & TxtCodigoActivacion.Text
            My.Settings.Save()
            My.Settings.Reload()
            vActivado = True
            MsgBox(rmse.GetString("SoftwareActivado"), MsgBoxStyle.Information, rmse.GetString("ActivacionCorrecta"))
        Else
            MsgBox(rmse.GetString("SoftwareNoActivado"), MsgBoxStyle.Exclamation, rmse.GetString("ErrorActivacion"))
        End If
    End Sub

    Private Sub BtnSalir_Click(sender As Object, e As EventArgs) Handles BtnSalir.Click
        Me.Close()
    End Sub
End Class