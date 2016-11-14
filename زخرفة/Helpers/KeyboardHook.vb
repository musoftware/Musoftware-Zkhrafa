Public Class KeyboardHook
    Private Const HC_ACTION As Integer = 0
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN = &H100
    Private Const WM_KEYUP = &H101
    Private Const WM_SYSKEYDOWN = &H104
    Private Const WM_SYSKEYUP = &H105

    ''Keypress Structure 
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As Integer
        Public scancode As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure
    ''API Functions 
    Private Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" _
    (ByVal idHook As Integer, _
    ByVal lpfn As KeyboardProcDelegate, _
    ByVal hmod As Integer, _
    ByVal dwThreadId As Integer) As Integer

    Private Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Integer, _
    ByVal nCode As Integer, _
    ByVal wParam As Integer, _
    ByVal lParam As KBDLLHOOKSTRUCT) As Integer

    Private Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Integer) As Integer

    ''Our Keyboard Delegate 
    Private Delegate Function KeyboardProcDelegate _
    (ByVal nCode As Integer, _
    ByVal wParam As Integer, _
    ByRef lParam As KBDLLHOOKSTRUCT) As Integer

    ''The KeyPress events 
    Public Shared Event KeyDown(ByVal Key As System.Windows.Forms.Keys)
    Public Shared Event KeyUp(ByVal Key As System.Windows.Forms.Keys)
    ''The identifyer for our KeyHook 
    Private Shared KeyHook As Integer
    ''KeyHookDelegate 
    Private Shared KeyHookDelegate As KeyboardProcDelegate

    Public Sub New()
        ''Installs a Low Level Keyboard Hook 
        KeyHookDelegate = New KeyboardProcDelegate(AddressOf KeyboardProc)
        KeyHook = SetWindowsHookEx(WH_KEYBOARD_LL, KeyHookDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
    End Sub

    Private Shared Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As KBDLLHOOKSTRUCT) As Integer
        ''If it is a keypress 
        If (nCode = HC_ACTION) Then
            Select Case wParam
                ''If it is a Keydown Event 
                Case WM_KEYDOWN, WM_SYSKEYDOWN
                    ''Activates the KeyDown event in Form 1 
                    RaiseEvent KeyDown(CType(lParam.vkCode, System.Windows.Forms.Keys))
                Case WM_KEYUP, WM_SYSKEYUP
                    ''Activates the KeyUp event in Form 1 
                    RaiseEvent KeyUp(CType(lParam.vkCode, System.Windows.Forms.Keys))
            End Select
        End If
        ''Next 
        Return CallNextHookEx(KeyHook, nCode, wParam, lParam)
    End Function

    Protected Overrides Sub Finalize()
        ''On close it UnHooks the Hook 
        UnhookWindowsHookEx(KeyHook)
        MyBase.Finalize()
    End Sub
End Class