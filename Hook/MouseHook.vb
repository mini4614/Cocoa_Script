Imports System.Runtime.InteropServices

Public Class MouseHook
    Implements IDisposable
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Integer, ByVal lpfn As MouseProcDelegate, ByVal hmod As Integer, ByVal dwThreadId As Integer) As Integer
    'Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Integer, ByVal nCode As Integer, ByVal wParam As Integer, ByVal lParam As MSLLHOOKSTRUCT) As Integer
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As IntPtr) As Integer

    <DllImport("user32.dll")> _
    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function


    <Flags()> _
    Public Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum
    <StructLayout(LayoutKind.Sequential)> _
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    Private Delegate Function MouseProcDelegate(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer

    Private Structure MSLLHOOKSTRUCT
        Public pt As Point
        Public mouseData As Integer
        Public flags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure

    Public Enum Wheel_Direction
        WheelUp
        WheelDown
    End Enum

    Private Const HC_ACTION As Integer = 0
    Private Const WH_MOUSE_LL As Integer = 14

    Private Const WM_MOUSEMOVE As Integer = &H200

    Private Const WM_LBUTTONDOWN As Integer = &H201
    Private Const WM_LBUTTONUP As Integer = &H202

    Private Const WM_RBUTTONDOWN As Integer = &H204
    Private Const WM_RBUTTONUP As Integer = &H205



    Private MouseHook As Integer
    Private MouseHookDelegate As MouseProcDelegate

    Public Event Mouse_Move(ByVal ptLocat As Point)

    Public Event Mouse_Left_Down(ByVal ptLocat As Point)
    Public Event Mouse_Left_Up(ByVal ptLocat As Point)

    Public Event Mouse_Right_Down(ByVal ptLocat As Point)
    Public Event Mouse_Right_Up(ByVal ptLocat As Point)


    Public Sub New()
        MouseHookDelegate = New MouseProcDelegate(AddressOf MouseProc)
        MouseHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
    End Sub

    Private Function MouseProc(ByVal nCode As Integer, ByVal wParam As Integer, ByRef lParam As MSLLHOOKSTRUCT) As Integer

        If (nCode = HC_ACTION) Then
            Select Case wParam
                Case WM_MOUSEMOVE
                    RaiseEvent Mouse_Move(lParam.pt)

                Case WM_LBUTTONDOWN
                    RaiseEvent Mouse_Left_Down(lParam.pt)
                Case WM_LBUTTONUP
                    RaiseEvent Mouse_Left_Up(lParam.pt)

                Case WM_RBUTTONDOWN
                    RaiseEvent Mouse_Right_Down(lParam.pt)
                Case WM_RBUTTONUP
                    RaiseEvent Mouse_Right_Up(lParam.pt)

            End Select
        End If

        Return CallNextHookEx(MouseHook, nCode, wParam, lParam.dwExtraInfo)
    End Function

    Protected Overrides Sub Finalize()
        UnhookWindowsHookEx(MouseHook)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            UnhookWindowsHookEx(MouseHook)
        Catch ex As Exception
        End Try
    End Sub
End Class