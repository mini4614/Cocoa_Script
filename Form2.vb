
Public Class Form2
    Public graph As Graphics
    Public p As New Pen(Color.FromArgb(0, 0, 0))
    Public Declare Sub SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer)
    Private Declare Sub ReleaseCapture Lib "User32" ()
    Const WM_CHAR As Integer = &H102
    Const WM_KEYDOWN As Integer = &H100
    Const VK_RETURN As Integer = &HD
    Const WM_NCLBUTTONDOWN = &HA1
    Const WM_IME_CHAR As Long = &H286

    Const HTCAPTION = 2



    Public Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Height = Form3.XY1.Y * 2
        Width = Form3.XY1.X * 2


        If Form3.RadioButton3.Checked = True Then ' 마우스 일경우
            Top = MousePosition.Y - Height / 2 - Form3.yjo
            Left = MousePosition.X - Width / 2
        Else
            Top = Form3.aim.Y - Height / 2 - Form3.yjo
            Left = Form3.aim.X - Width / 2

        End If


    End Sub

End Class