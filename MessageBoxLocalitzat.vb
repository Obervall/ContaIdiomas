Imports System.Runtime.InteropServices
Imports System.Windows.Forms

Public Module MsgBoxTraductorGlobal

    ' --- DECLARACIÓN CORREGIDA DE APIS NATIVAS (COMPATIBLE 32/64 BITS) ---
    Private Declare Auto Function SetWindowsHookEx Lib "user32.dll" (ByVal idHook As Integer, ByVal lpfn As HookProc, ByVal hmod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    Private Declare Auto Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As IntPtr) As Boolean
    Private Declare Auto Function SetDlgItemText Lib "user32.dll" (ByVal hDlg As IntPtr, ByVal nIDDlgItem As Integer, ByVal lpString As String) As Boolean
    Private Declare Auto Function GetCurrentThreadId Lib "kernel32.dll" () As UInteger

    Private Delegate Function HookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Private hHook As IntPtr = IntPtr.Zero

    ' Textos globales de los botones (Se rellenan en CambiarIdiomaGlobal)
    Public TextBotoOk As String = "OK"
    Public TextBotoCancel As String = "Cancel"
    Public TextBotoYes As String = "Yes"
    Public TextBotoNo As String = "No"

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        ' HCBT_ACTIVATE = 5
        If nCode = 5 Then
            ' Forzamos los textos exactos usando los IDs nativos de Windows
            SetDlgItemText(wParam, 1, TextBotoOk)      ' ID 1 = OK / Aceptar
            SetDlgItemText(wParam, 2, TextBotoCancel)  ' ID 2 = Cancel / Cancelar
            SetDlgItemText(wParam, 6, TextBotoYes)     ' ID 6 = Yes / Sí
            SetDlgItemText(wParam, 7, TextBotoNo)      ' ID 7 = No

            ' Una vez aplicados, liberamos el gancho inmediatamente
            UnhookWindowsHookEx(hHook)
        End If

        ' Pasamos el control al siguiente gancho de la cadena de Windows
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    Private Declare Auto Function CallNextHookEx Lib "user32.dll" (ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    ' ========================================================
    ' INTERCEPTOR DE LOS "MsgBox" CLÁSICOS
    ' ========================================================
    Public Function MsgBox(ByVal Prompt As Object, Optional ByVal Buttons As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional ByVal Title As Object = Nothing) As MsgBoxResult
        Dim btnNet As MessageBoxButtons = CType(Buttons And 7, MessageBoxButtons)
        Dim icoNet As MessageBoxIcon = CType(Buttons And &H70, MessageBoxIcon)
        Dim defNet As MessageBoxDefaultButton = CType(Buttons And &H300, MessageBoxDefaultButton)
        Dim txtTitol As String = If(Title IsNot Nothing, Title.ToString(), "")

        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, CInt(GetCurrentThreadId()))
        Dim res As DialogResult = System.Windows.Forms.MessageBox.Show(Prompt.ToString(), txtTitol, btnNet, icoNet, defNet)

        Return CType(res, MsgBoxResult)
    End Function

    ' ========================================================
    ' INTERCEPTOR DE LOS "MessageBox.Show"
    ' ========================================================
    Public Function Show(ByVal text As String) As DialogResult
        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, GetCurrentThreadId())
        Return System.Windows.Forms.MessageBox.Show(text)
    End Function

    Public Function Show(ByVal text As String, ByVal caption As String) As DialogResult
        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, GetCurrentThreadId())
        Return System.Windows.Forms.MessageBox.Show(text, caption)
    End Function

    Public Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons) As DialogResult
        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, GetCurrentThreadId())
        Return System.Windows.Forms.MessageBox.Show(text, caption, buttons)
    End Function

    Public Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon) As DialogResult
        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, GetCurrentThreadId())
        Return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon)
    End Function

    Public Function Show(ByVal text As String, ByVal caption As String, ByVal buttons As MessageBoxButtons, ByVal icon As MessageBoxIcon, ByVal defaultButton As MessageBoxDefaultButton) As DialogResult
        hHook = SetWindowsHookEx(5, AddressOf HookCallback, IntPtr.Zero, GetCurrentThreadId())
        Return System.Windows.Forms.MessageBox.Show(text, caption, buttons, icon, defaultButton)
    End Function

End Module
