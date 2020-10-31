Imports System.Threading
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.Math

Public Class Form3
    Public Declare Function ReleaseCapture Lib "user32" () As Integer
    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (
    ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
    Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal x As Integer, ByVal y As Integer) As Integer
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Dim myPort As Array  'COM Ports detected on the system will be stored here
    Delegate Sub SetTextCallback(ByVal [text] As String) 'Added to prevent threading errors during receiveing of data


    <DllImport("user32.dll")> Private Shared Function GetSystemMetrics(ByVal smIndex As Integer) As Integer
    End Function
    <DllImport("user32.dll")> Private Shared Function SendInput(ByVal nInputs As Integer, ByRef pInputs As INPUT, ByVal cbSize As Integer) As Integer
    End Function
    Private Structure INPUT
        Dim dwType As Integer
        Dim mkhi As MOUSEKEYBDHARDWAREINPUT
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Private Structure MOUSEKEYBDHARDWAREINPUT
        <FieldOffset(0)> Public mi As MOUSEINPUT

    End Structure
    Private Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As Integer
    End Structure
    Dim it As New INPUT
    Public slp, slp2, slp3, slp4 As String
    Public errorlevel, red, green, blue, fixx, fixy, movex, movey, bg, bg2, speedx, speedy, er, eg, eb, cspeedx, cspeedy, check, cl, cr, loading, start
    Public aimx, aimy, dirx, diry, AimOffsetX, AimOffsetY, RootX, RootY, flick, m, az, x, y, xx, yy, yjo, yjo1, yjo2
    Public XY1, XY2, result, aim, sa, ea As Point
    Public key, sendx, sendy As String
    Public kbHook, kbhook2 As FullHook
    Public t1 As New Thread(AddressOf thread1)
    Public t2 As New Thread(AddressOf thread2)
    Public t3 As New Thread(AddressOf thread3)
    Public t4 As New Thread(AddressOf thread4)
    Public Running As Boolean = False
    Dim find As Integer = 0
    Dim rfd, zbux, zbuy




    Private Sub thread1()
        Do
            result = PixelSearch(aim.X - XY1.X, aim.Y - yjo1, aim.X + XY1.X, aim.Y + yjo2, red, green, blue, eb)

            ' If Not find = 0 Then
            '     rfd = find
            '  End If

            Thread.Sleep(slp)
            '    Delay(slp)
        Loop
    End Sub

    Private Sub thread3()
        Do
            result = PixelSearch(sa.X, sa.Y, ea.X, ea.Y, red, green, blue, eb)

            ' If Not find = 0 Then
            '     rfd = find
            '  End If

            Thread.Sleep(slp)
            '    Delay(slp)
        Loop
    End Sub

    Private Sub thread2()
        Do

            If speedx = 0 Then
                it.mkhi.mi.dx = 0
            Else
                xx = (result.X + fixx)
                '  xx = (result.X + fixx + (rfd * zbux))

                it.mkhi.mi.dx = (xx - aim.X) * errorlevel / speedx

                If it.mkhi.mi.dx > 127 Then
                    it.mkhi.mi.dx = 127
                End If
                If it.mkhi.mi.dx < -127 Then
                    it.mkhi.mi.dx = -127
                End If

            End If

            If speedy = 0 Then
                it.mkhi.mi.dy = 0
            Else
                yy = (result.Y + fixy)
                ' yy = (result.Y + fixy + (rfd * zbuy))

                it.mkhi.mi.dy = (yy - aim.Y) * errorlevel / speedy

                If it.mkhi.mi.dy > 127 Then
                    it.mkhi.mi.dy = 127
                End If
                If it.mkhi.mi.dy < -127 Then
                    it.mkhi.mi.dy = -127
                End If
            End If



            If it.mkhi.mi.dx < 0 Then
                sendx = "-" & Microsoft.VisualBasic.Right("000" & it.mkhi.mi.dx, 4).Replace("-", "")
            Else
                sendx = "+" & Microsoft.VisualBasic.Right("000" & it.mkhi.mi.dx, 3)
            End If
            If it.mkhi.mi.dy < 0 Then
                sendy = "-" & Microsoft.VisualBasic.Right("000" & it.mkhi.mi.dy, 4).Replace("-", "")
            Else
                sendy = "+" & Microsoft.VisualBasic.Right("000" & it.mkhi.mi.dy, 3)
            End If


            If check = True Then
                If ((it.mkhi.mi.dx <= 0) And (it.mkhi.mi.dy >= 0) And (xx - it.mkhi.mi.dx <= aim.X) And (yy - it.mkhi.mi.dy >= aim.Y)) Or ((it.mkhi.mi.dx >= 0) And (it.mkhi.mi.dy >= 0) And (xx - it.mkhi.mi.dx >= aim.X) And (yy - it.mkhi.mi.dy >= aim.Y)) Or ((it.mkhi.mi.dx <= 0) And (it.mkhi.mi.dy <= 0) And (xx - it.mkhi.mi.dx <= aim.X) And (yy - it.mkhi.mi.dy <= aim.Y)) Or ((it.mkhi.mi.dx >= 0) And (it.mkhi.mi.dy <= 0) And (xx - it.mkhi.mi.dx >= aim.X) And (yy - it.mkhi.mi.dy <= aim.Y)) Then
                    SerialPort1.Write(sendx & sendy & vbCr)

                End If

            End If



            Thread.Sleep(slp2)

        Loop

    End Sub

    Private Sub thread4()
        Do

            If flick = False And check = True Then
                m = 1
            End If

            aimx = ((result.X + fixx) - aim.X)



            If speedx = 0 Then
                it.mkhi.mi.dx = 0
            Else
                it.mkhi.mi.dx = ((result.X + fixx) - aim.X) * errorlevel * speedx
            End If

            If it.mkhi.mi.dx > 124 Then
                it.mkhi.mi.dx = 124
            End If
            If it.mkhi.mi.dx < -124 Then
                it.mkhi.mi.dx = -124
            End If

            If m = 1 Then
                SerialPort1.Write(it.mkhi.mi.dx & vbCr)
                m = 0
            End If

            flick = check

            Thread.Sleep(slp2)

        Loop

    End Sub



    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TID.TextChanged
        If Not TPW.Text = "" And Not TID.Text = "" Then
            PictureBox4.Visible = True
            PictureBox3.Visible = False

            '   PictureBox5.Visible = False
        Else
            PictureBox3.Visible = True
            PictureBox4.Visible = False
            '   PictureBox5.Visible = False

        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TPW.TextChanged
        If Not TPW.Text = "" And Not TID.Text = "" Then
            PictureBox4.Visible = True
            PictureBox3.Visible = False

            ' PictureBox5.Visible = False
        Else
            PictureBox3.Visible = True
            PictureBox4.Visible = False
            '  PictureBox5.Visible = False

        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not ComboBox2.Text = "" Then
            SerialPort1.PortName = ComboBox2.Text         'Set SerialPort1 to the selected COM port at startup
            SerialPort1.BaudRate = "57600"     'Set Baud rate to the selected value on
            SerialPort1.Parity = IO.Ports.Parity.None
            SerialPort1.StopBits = IO.Ports.StopBits.One
            SerialPort1.DataBits = 8            'Open our serial port
            SerialPort1.Open()
            Button1.Enabled = False
            Button2.Enabled = True
        End If


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        SerialPort1.Close()             'Close our Serial Port

        Button1.Enabled = True

        Button2.Enabled = False
    End Sub


    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load






        ' 연결된 port 추가



        Width = 364
        Height = 496


        Panel1.Location = New Point(55, 140)

        Panel3.Location = New Point(55, 140)

        Panel21.Location = New Point(55, 140)

        Panel23.Location = New Point(55, 140)

        Panel17.Location = New Point(55, 140)

        Panel27.Location = New Point(55, 140)

        Panel19.Location = New Point(55, 140)

        Panel25.Location = New Point(55, 140)


        PictureBox3.Location = New Point(55, 327)
        PictureBox4.Location = New Point(55, 327)
        PictureBox5.Location = New Point(55, 327)

        PictureBox8.Location = New Point(55, 374)
        PictureBox9.Location = New Point(55, 374)

        PictureBox16.Location = New Point(25, 44)
        PictureBox17.Location = New Point(64, 44)

        PictureBox20.Location = New Point(305, 25)
        PictureBox21.Location = New Point(305, 25)

        PictureBox24.Location = New Point(310, 44)


        PictureBox20.Visible = True
        PictureBox21.Visible = False


        PictureBox14.Visible = False
        PictureBox15.Visible = True

        PictureBox16.Visible = True
        PictureBox17.Visible = False

        PictureBox8.Visible = True
        PictureBox9.Visible = False

        PictureBox3.Visible = True
        PictureBox4.Visible = False
        PictureBox5.Visible = False

        Panel5.Width = 350
        Panel5.Location = New Point(0, -94)
        Panel6.Location = New Point(0, 93)
        Panel7.Location = New Point(0, 93)
        Panel5.Visible = False
        Panel6.Visible = True
        Panel7.Visible = False


        it.dwType = 0
        it.mkhi.mi.dwFlags = &H1

        Dim a As DataGridViewCellEventArgs
        Call Grid_CellContentClick(sender, a)
        PictureBox14.Visible = True
        PictureBox17.Visible = True
        PictureBox15.Visible = False


        PictureBox16.Visible = False

        Panel7.Visible = True

        kbhook2 = New FullHook()
        kbhook2.kstart(Me)

        Call PictureBox15_Click(sender, e)
        '     

        ' 프로그램설정

        'Me.ShowInTaskbar = False

        '   프로그램숨기기ToolStripMenuItem.Checked = ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "1", "")


        'If 프로그램숨기기ToolStripMenuItem.Checked = True Then
        'Me.ShowInTaskbar = False
        'Else
        'Me.ShowInTaskbar = True
        'End If


        ' If ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "2", "") = "True" Then

        ' PictureBox9.Visible = True
        'PictureBox8.Visible = False

        'TID.Text = ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "3", "")
        'TPW.Text = ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "4", "")

        'Call TextBox2_TextChanged(sender, e)
        'Call PictureBox5_Click(sender, e)

        'End If






    End Sub

    Private Sub Form3_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        t1.Abort()
        t1 = New Threading.Thread(AddressOf thread1)
        t2.Abort()
        t2 = New Threading.Thread(AddressOf thread2)
        t3.Abort()
        t3 = New Threading.Thread(AddressOf thread3)
        t4.Abort()
        t4 = New Threading.Thread(AddressOf thread4)

        If kbHook IsNot Nothing Then
            kbHook.DropAll()
        End If
        If kbhook2 IsNot Nothing Then
            kbhook2.kDropAll()
        End If


        Call Label4_Click(sender, e)



    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If CheckBox2.Checked = True Then
            RadioButton1.Enabled = False
            RadioButton2.Enabled = False
        Else
            RadioButton1.Enabled = True
            RadioButton2.Enabled = True

        End If
    End Sub

    Private Sub PictureBox4_MouseMove(sender As Object, e As MouseEventArgs) Handles PictureBox4.MouseMove
        If PictureBox4.Visible = True Then
            ' PictureBox3.Visible = False
            PictureBox5.Visible = True
            PictureBox4.Visible = False



        End If
    End Sub

    Private Sub Label24_MouseMove(sender As Object, e As EventArgs) Handles Label24.MouseMove
        Label24.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label24_MouseLeave(sender As Object, e As EventArgs) Handles Label24.MouseLeave
        Label24.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label30_MouseMove(sender As Object, e As EventArgs) Handles Label30.MouseMove
        Label30.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label30_MouseLeave(sender As Object, e As EventArgs) Handles Label30.MouseLeave
        Label30.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label31_MouseMove(sender As Object, e As EventArgs) Handles Label31.MouseMove
        Label31.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label31_MouseLeave(sender As Object, e As EventArgs) Handles Label31.MouseLeave
        Label31.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label32_MouseMove(sender As Object, e As EventArgs) Handles Label32.MouseMove
        Label32.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label32_MouseLeave(sender As Object, e As EventArgs) Handles Label32.MouseLeave
        Label32.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label33_MouseMove(sender As Object, e As EventArgs) Handles Label33.MouseMove
        Label33.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label33_MouseLeave(sender As Object, e As EventArgs) Handles Label33.MouseLeave
        Label33.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label34_MouseMove(sender As Object, e As EventArgs) Handles Label34.MouseMove
        Label34.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label34_MouseLeave(sender As Object, e As EventArgs) Handles Label34.MouseLeave
        Label34.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label35_MouseMove(sender As Object, e As EventArgs) Handles Label35.MouseMove
        Label35.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label35_MouseLeave(sender As Object, e As EventArgs) Handles Label35.MouseLeave
        Label35.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub Label36_MouseMove(sender As Object, e As EventArgs) Handles Label36.MouseMove
        Label36.BackColor = Color.FromArgb(245, 245, 245)
    End Sub
    Private Sub Label36_MouseLeave(sender As Object, e As EventArgs) Handles Label36.MouseLeave
        Label36.BackColor = Color.FromArgb(255, 255, 255)
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub PictureBox5_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox5.MouseLeave
        If PictureBox5.Visible = True Then
            'PictureBox3.Visible = False
            PictureBox4.Visible = True
            PictureBox5.Visible = False
        End If
    End Sub

    Private Sub Label37_Click(sender As Object, e As EventArgs) Handles Label37.MouseDown
        Panel17.BringToFront()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "1", TextBox5.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "2", TextBox6.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "3", TextBox7.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "4", TextBox8.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "5", TextBox4.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "COLOR", "6", TextBox9.Text)

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "SPEED", "1", HScrollBar1.Value)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "SPEED", "2", HScrollBar2.Value)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "SPEED", "3", RadioButton5.Checked)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "SPEED", "4", RadioButton6.Checked)



        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FOV", "1", TextBox1.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FOV", "2", TextBox2.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FOV", "3", TextBox99.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FOV", "4", RadioButton3.Checked)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FOV", "5", RadioButton4.Checked)

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FIX", "1", HScrollBar3.Value)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "FIX", "2", HScrollBar4.Value)

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "HOTKEY", "1", RadioButton1.Checked)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "HOTKEY", "2", RadioButton2.Checked)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "HOTKEY", "3", TextBox17.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "HOTKEY", "4", CheckBox2.Checked)

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "DELAY", "1", TextBox11.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "DELAY", "2", TextBox12.Text)
        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "DELAY", "3", TextBox13.Text)

        writeIni(Application.StartupPath & "\Config" & "\" & Label1.Text & ".ini", "OVERLAY", "1", CheckBox1.Checked)


    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click
        PictureBox8.Visible = False
        PictureBox9.Visible = True
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click

        If Grid.RowCount > 0 Then
            Do
                Grid.Rows.RemoveAt(Grid.SelectedRows(0).Index)
            Loop While Grid.RowCount > 0
        End If

        Dim files = My.Computer.FileSystem.GetFiles(Application.StartupPath + "\Config\", FileIO.SearchOption.SearchAllSubDirectories)
        For i = 0 To files.Count - 1
            Dim sp As String = files.Item(i).ToString()
            Dim sp2 As String = Application.StartupPath & "\Config\"
            Dim sp3 = Split(Split(sp, sp2)(1), ".")
            If sp3(1) = "ini" Then
                Grid.Rows.Add("          " & sp3(0) & "                                                            ")
            End If
        Next

        For i = 0 To Grid.Rows.Count - 1
            Dim r As DataGridViewRow = Grid.Rows(i)
            r.Height = 57
        Next

        Dim a As DataGridViewCellEventArgs
        Call Grid_CellContentClick(sender, a)

        '  Panel5.Visible = True



    End Sub


    Private Sub Tsc_TextChanged(sender As Object, e As EventArgs) Handles Tsc.TextChanged
        '   MsgBox(Grid.RowCount)
        If Tsc.Text = "" Then
            Call Label21_Click(sender, e)
        Else

            If Grid.RowCount > 0 Then
                Do
                    Grid.Rows.RemoveAt(Grid.SelectedRows(0).Index)
                Loop While Grid.RowCount > 0
            End If


            Dim files = My.Computer.FileSystem.GetFiles(Application.StartupPath + "\Config\", FileIO.SearchOption.SearchAllSubDirectories)
            For i = 0 To files.Count - 1
                Dim sp As String = files.Item(i).ToString()
                Dim sp2 As String = Application.StartupPath & "\Config\"
                Dim sp3 = Split(Split(sp, sp2)(1), ".")
                If sp3(1) = "ini" Then
                    If sp3(0).Contains(Tsc.Text) Then
                        Grid.Rows.Add("          " & sp3(0) & "                                                            ")
                    End If

                End If
            Next

            For i = 0 To Grid.Rows.Count - 1
                Dim r As DataGridViewRow = Grid.Rows(i)
                r.Height = 57
            Next

            Dim a As DataGridViewCellEventArgs
            ' Call Grid_CellContentClick(sender, a)

        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click
        PictureBox8.Visible = True
        PictureBox9.Visible = False
    End Sub


    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        Panel17.Visible = True
        Panel17.BringToFront()
    End Sub

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        Panel21.Visible = True
        Panel21.BringToFront()
    End Sub

    Private Sub Label31_Click(sender As Object, e As EventArgs) Handles Label31.Click
        Panel19.Visible = True
        Panel19.BringToFront()
    End Sub

    Private Sub Label32_Click(sender As Object, e As EventArgs) Handles Label32.Click
        Panel23.Visible = True
        Panel23.BringToFront()
    End Sub

    Private Sub Label33_Click(sender As Object, e As EventArgs) Handles Label33.Click
        Panel25.Visible = True
        Panel25.BringToFront()
    End Sub

    Private Sub Label34_Click(sender As Object, e As EventArgs) Handles Label34.Click
        Panel27.Visible = True
        Panel27.BringToFront()
    End Sub
    Private Sub Label35_Click(sender As Object, e As EventArgs) Handles Label35.Click
        Panel1.Visible = True
        Panel1.BringToFront()
    End Sub



    Private Sub Label57_Click(sender As Object, e As EventArgs) Handles Label57.Click
        Panel1.Visible = False
        Panel1.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub
    Private Sub Label40_Click(sender As Object, e As EventArgs) Handles Label40.Click
        Panel17.Visible = False
        Panel17.Location = New Point(55, 140)

        Call Label4_Click(sender, e)

    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        Panel21.Visible = False
        Panel21.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub

    Private Sub Label41_Click(sender As Object, e As EventArgs) Handles Label41.Click
        Panel19.Visible = False
        Panel19.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        Panel23.Visible = False
        Panel23.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub


    Private Sub Label53_Click(sender As Object, e As EventArgs) Handles Label53.Click
        Panel25.Visible = False
        Panel25.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        Panel27.Visible = False
        Panel27.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub

    Private Sub Label44_Click(sender As Object, e As EventArgs) Handles Label44.MouseDown
        Panel19.BringToFront()
    End Sub

    Private Sub Label47_Click(sender As Object, e As EventArgs) Handles Label47.MouseDown
        Panel21.BringToFront()
    End Sub

    Private Sub Label50_Click(sender As Object, e As EventArgs) Handles Label50.MouseDown
        Panel23.BringToFront()
    End Sub

    Private Sub Label56_Click(sender As Object, e As EventArgs) Handles Label56.MouseDown
        Panel25.BringToFront()
    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.MouseDown
        Panel27.BringToFront()
    End Sub

    Private Sub PictureBox23_MouseMove(sender As Object, e As EventArgs) Handles PictureBox23.MouseMove
        PictureBox23.Visible = False
        PictureBox24.Visible = True
    End Sub

    Private Sub PictureBox24_MouseLeave(sender As Object, e As EventArgs) Handles PictureBox24.MouseLeave
        PictureBox24.Visible = False
        PictureBox23.Visible = True
    End Sub



    Private Sub Label47_Click_1(sender As Object, e As EventArgs) Handles Label47.Click

    End Sub

    Private Sub Panel21_Click(sender As Object, e As EventArgs) Handles Panel21.MouseDown
        Panel21.BringToFront()
    End Sub



    Private Sub 종료ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 종료ToolStripMenuItem.Click
        Call Label4_Click(sender, e)

        Me.Close()
    End Sub


    Private Sub Panel17_Click(sender As Object, e As EventArgs) Handles Panel17.MouseDown
        Panel17.BringToFront()
    End Sub

    Private Sub TextBox2_TextChanged_1(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Panel5_Paint(sender As Object, e As PaintEventArgs) Handles Panel5.Paint

    End Sub

    Private Sub 프로그램숨기기ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 프로그램숨기기ToolStripMenuItem.Click

        If 프로그램숨기기ToolStripMenuItem.Checked = False Then
            프로그램숨기기ToolStripMenuItem.Checked = True
            Me.ShowInTaskbar = False

        Else
            프로그램숨기기ToolStripMenuItem.Checked = False
            Me.ShowInTaskbar = True
        End If
        writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "1", 프로그램숨기기ToolStripMenuItem.Checked)


    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub 자동로그인ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 자동로그인ToolStripMenuItem.Click

        If 자동로그인ToolStripMenuItem.Checked = False Then
            자동로그인ToolStripMenuItem.Checked = True
            writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "2", True)
            '   writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "3", TID.Text)
            '     writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "4", TPW.Text)
        Else
            자동로그인ToolStripMenuItem.Checked = False
            writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "2", False)
        End If

    End Sub

    Private Sub 로그아웃ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles 로그아웃ToolStripMenuItem1.Click

        If ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "2", "") = "True" Then

            PictureBox9.Visible = True
            PictureBox8.Visible = False

            TID.Text = ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "3", "")
            TPW.Text = ReadIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "4", "")

            Call TextBox2_TextChanged(sender, e)

        End If



        PictureBox6.Visible = True
        PictureBox7.Visible = True


        Panel5.Visible = False
    End Sub

    Private Sub Panel23_Click(sender As Object, e As EventArgs) Handles Panel23.MouseDown
        Panel23.BringToFront()
    End Sub



    Private Sub Label60_Click(sender As Object, e As EventArgs) Handles Label60.Click
        Panel1.BringToFront()
    End Sub

    Private Sub Label63_Click(sender As Object, e As EventArgs) Handles Label63.Click
        Panel3.Visible = False
        Panel3.Location = New Point(55, 140)

        Call Label4_Click(sender, e)
    End Sub

    Private Sub Label65_Click(sender As Object, e As EventArgs) Handles Label65.Click

    End Sub

    Private Sub Panel27_Click(sender As Object, e As EventArgs) Handles Panel27.MouseDown
        Panel27.BringToFront()
    End Sub


    Private Sub Panel19_Click(sender As Object, e As EventArgs) Handles Panel19.MouseDown
        Panel19.BringToFront()
    End Sub

    Private Sub Panel23_Paint(sender As Object, e As PaintEventArgs) Handles Panel23.Paint

    End Sub

    Private Sub Panel19_Paint(sender As Object, e As PaintEventArgs) Handles Panel19.Paint

    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

    Private Sub Label36_Click(sender As Object, e As EventArgs) Handles Label36.Click
        Panel3.Visible = True
        Panel3.BringToFront()

        myPort = IO.Ports.SerialPort.GetPortNames() 'Get all com ports available
        For i = 0 To UBound(myPort)
            ComboBox2.Items.Add(myPort(i))
        Next

    End Sub

    Private Sub Panel25_Click(sender As Object, e As EventArgs) Handles Panel25.MouseDown
        Panel25.BringToFront() '''''''''''''''''
    End Sub
    Private Sub Panel1_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel1.MouseDown
        Panel1.BringToFront()
    End Sub




    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        Me.Close()
    End Sub



    Private Sub PictureBox6_Click(sender As Object, e As EventArgs) Handles PictureBox6.Click
        Me.WindowState = FormWindowState.Minimized

    End Sub


    Private Sub PictureBox5_Click(sender As Object, e As EventArgs) Handles PictureBox5.Click
        PictureBox22.Visible = True
        PictureBox3.Visible = True
        PictureBox4.Visible = False
        PictureBox5.Visible = False
        TID.Enabled = False

        TPW.Enabled = False

        Delay(1000)

        If 네이버로그인(TID.Text, TPW.Text) = True Then

            loading = True







            ' 폴더존재여부

            If System.IO.Directory.Exists(Application.StartupPath & "\Config") = True Then
            Else
                System.IO.Directory.CreateDirectory(Application.StartupPath & "\Config")
            End If


            ' 스크립트 읽어오기


            Dim files = My.Computer.FileSystem.GetFiles(Application.StartupPath + "\Config\", FileIO.SearchOption.SearchAllSubDirectories)
            For i = 0 To files.Count - 1
                Dim sp As String = files.Item(i).ToString()
                Dim sp2 As String = Application.StartupPath & "\Config\"
                Dim sp3 = Split(Split(sp, sp2)(1), ".")
                If sp3(1) = "ini" Then
                    Grid.Rows.Add("          " & sp3(0) & "                                                            ")
                End If
            Next

            For i = 0 To Grid.Rows.Count - 1
                Dim r As DataGridViewRow = Grid.Rows(i)
                r.Height = 57
            Next

            Dim a As DataGridViewCellEventArgs
            Call Grid_CellContentClick(sender, a)
            Call Label4_Click(sender, e)

            '아디비번 저장

            ' writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "2", PictureBox9.Visible)

            'If PictureBox9.Visible = True Then
            '     자동로그인ToolStripMenuItem.Checked = True
            '  Else
            '        자동로그인ToolStripMenuItem.Checked = False
            '    End If

            writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "3", TID.Text)
            writeIni(Application.StartupPath & "\Config\Default.ini", "SETTING", "4", TPW.Text)






            '   Delay(1000)

            '로드전 새로고침
            Call Label21_Click(sender, e)

            PictureBox6.Visible = False
            PictureBox7.Visible = False
            PictureBox8.Visible = True
            PictureBox9.Visible = False


            '   Delay(50)

            Panel5.Visible = True

            TID.Text = ""
            TPW.Text = ""


            loading = False
        Else


            Label23.Text = "로그인에 실패했습니다."
        End If

        TID.Enabled = True

        TPW.Enabled = True
        PictureBox22.Visible = False
        PictureBox3.Visible = False
        PictureBox4.Visible = True
        PictureBox5.Visible = False

    End Sub

    Private Sub PictureBox1_MouseMove(sender As Object, e As MouseEventArgs)
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Me.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Panel5_MouseMove(sender As Object, e As MouseEventArgs) Handles Panel5.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Me.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click
        Call Label4_Click(sender, e)

        Me.Close()
    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Grid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles Grid.CellContentClick

        If Not Grid.RowCount = 0 Then
            Dim s As EventArgs
            Call Label4_Click(sender, s)

            Label1.Text = Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1)
            TextBox5.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "1", "")
            TextBox6.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "2", "")
            TextBox7.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "3", "")
            TextBox8.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "4", "")
            TextBox4.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "5", "")
            TextBox9.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "COLOR", "6", "")




            HScrollBar1.Value = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "SPEED", "1", "")
            HScrollBar2.Value = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "SPEED", "2", "")
            RadioButton5.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "SPEED", "3", "")
            RadioButton6.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "SPEED", "4", "")



            TextBox1.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FOV", "1", "")
            TextBox2.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FOV", "2", "")
            TextBox99.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FOV", "3", "")
            RadioButton3.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FOV", "4", "")
            RadioButton4.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FOV", "5", "")

            HScrollBar3.Value = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FIX", "1", "")
            HScrollBar4.Value = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "FIX", "2", "")

            Dim a As ScrollEventArgs
            Call HScrollBar1_Scroll(sender, a)
            Call HScrollBar2_Scroll(sender, a)
            Call HScrollBar3_Scroll(sender, a)
            Call HScrollBar4_Scroll(sender, a)

            RadioButton1.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "HOTKEY", "1", "")
            RadioButton2.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "HOTKEY", "2", "")
            CheckBox2.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "HOTKEY", "4", "")
            TextBox17.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "HOTKEY", "3", "")

            Call CheckBox2_CheckedChanged(sender, e)


            TextBox11.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "DELAY", "1", "")
            TextBox12.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "DELAY", "2", "")
            TextBox13.Text = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "DELAY", "3", "")


            CheckBox1.Checked = ReadIni(Application.StartupPath & "\Config" & "\" & Split(Grid.SelectedRows(0).Cells(0).Value, "          ")(1) & ".ini", "OVERLAY", "1", "")
            '설정값 읽어오기

            Call TextBox5_TextChanged(sender, e)
            Call TextBox6_TextChanged(sender, e)
            Call TextBox7_TextChanged(sender, e)


            '끗
        End If


    End Sub

    Private Sub PictureBox14_Click(sender As Object, e As EventArgs) Handles PictureBox14.Click


        If Panel17.Visible = False And Panel21.Visible = False And Panel23.Visible = False And Panel27.Visible = False And Panel19.Visible = False And Panel25.Visible = False Then
            PictureBox16.Visible = True
            PictureBox15.Visible = True
            PictureBox14.Visible = False

            PictureBox17.Visible = False

            Panel6.Visible = True
            Panel7.Visible = False

            kbhook2.kDropAll()
            kbhook2 = Nothing


        End If


    End Sub

    Private Sub PictureBox15_Click(sender As Object, e As EventArgs) Handles PictureBox15.Click
        If Not Grid.RowCount = 0 Then

            Dim a As DataGridViewCellEventArgs
            Call Grid_CellContentClick(sender, a)
            PictureBox14.Visible = True
            PictureBox17.Visible = True
            PictureBox15.Visible = False


            PictureBox16.Visible = False

            Panel7.Visible = True

            kbhook2 = New FullHook()
            kbhook2.kstart(Me)

        End If

    End Sub

    Public Sub PictureBox20_Click(sender As Object, e As EventArgs) Handles PictureBox20.Click
        ' If PictureBox20.Visible = False Then
        '    Call PictureBox21_Click(sender, e)
        '       End If
        PictureBox20.Visible = False
        PictureBox21.Visible = True


        Grid.Enabled = False
        Tsc.Enabled = False



        XY1.X = TextBox1.Text
        XY1.Y = TextBox2.Text
        yjo = TextBox99.Text
        yjo1 = XY1.Y + yjo
        yjo2 = XY1.Y - yjo



        red = TextBox5.Text
        green = TextBox6.Text
        blue = TextBox7.Text

        er = TextBox8.Text
        eg = TextBox4.Text
        eb = TextBox9.Text



        fixx = Label28.Text
        fixy = Label9.Text


        slp = TextBox11.Text
        slp2 = TextBox12.Text
        slp3 = TextBox13.Text

        'zbux = TextBox14.Text
        ' zbuy = TextBox15.Text

        If RadioButton5.Checked = True Then



            speedx = (101 - Label18.Text)
            speedy = (101 - Label19.Text)

        Else
            speedx = (Label18.Text - 95)
            speedy = (Label19.Text - 95)

        End If

        If Label18.Text = 0 Then
            speedx = 0
        End If
        If Label19.Text = 0 Then
            speedy = 0
        End If


        kbHook = New FullHook()
        kbHook.Start(Me)




        If slp < 0 Or slp2 < 0 Or slp3 <= 0 Or XY1.X <= 0 Or XY1.Y <= 0 Or Button1.Enabled = True Then
            Call PictureBox21_Click(sender, e)
        Else





            start = True

            If RadioButton3.Checked = True Then
                t1.Start()
            Else
                aim.X = Screen.PrimaryScreen.Bounds.Width / 2
                aim.Y = Screen.PrimaryScreen.Bounds.Height / 2
                sa.X = aim.X - XY1.X
                sa.Y = aim.Y - (XY1.Y + yjo)
                ea.X = aim.X + XY1.X
                ea.Y = aim.Y + XY1.Y - yjo


                t3.Start()


            End If

            If CheckBox1.Checked = True Then
                Form2.Show()
            End If


            If RadioButton5.Checked = True Then
                t2.Start()
            Else
                          t4.Start()
            End If




        End If






    End Sub

    Public Sub PictureBox21_Click(sender As Object, e As EventArgs) Handles PictureBox21.Click
        start = False

        PictureBox20.Visible = True
        PictureBox21.Visible = False

        Grid.Enabled = True
        Tsc.Enabled = True

        kbHook.DropAll()
        kbHook = Nothing



        Form2.Close()

        t1.Abort()
        t1 = New Threading.Thread(AddressOf thread1)
        t2.Abort()
        t2 = New Threading.Thread(AddressOf thread2)
        t3.Abort()
        t3 = New Threading.Thread(AddressOf thread3)
        t4.Abort()
        t4 = New Threading.Thread(AddressOf thread4)

    End Sub

    Private Sub HScrollBar3_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar3.Scroll
        Label28.Text = HScrollBar3.Value
    End Sub

    Private Sub HScrollBar1_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar1.Scroll
        Label18.Text = HScrollBar1.Value
    End Sub

    Private Sub HScrollBar2_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar2.Scroll
        Label19.Text = HScrollBar2.Value
    End Sub

    Private Sub HScrollBar4_Scroll(sender As Object, e As ScrollEventArgs) Handles HScrollBar4.Scroll
        Label9.Text = HScrollBar4.Value
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If IsNumeric(TextBox5.Text) = True And IsNumeric(TextBox6.Text) = True And IsNumeric(TextBox7.Text) = True Then
            If TextBox5.Text <= 255 And TextBox5.Text >= 0 And TextBox6.Text <= 255 And TextBox6.Text >= 0 And TextBox7.Text <= 255 And TextBox7.Text >= 0 Then
                PictureBox19.BackColor = Color.FromArgb(TextBox5.Text, TextBox6.Text, TextBox7.Text)
            End If
        Else
            PictureBox19.BackColor = Color.FromArgb("255", "255", "255")
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If IsNumeric(TextBox5.Text) = True And IsNumeric(TextBox6.Text) = True And IsNumeric(TextBox7.Text) = True Then
            If TextBox5.Text <= 255 And TextBox5.Text >= 0 And TextBox6.Text <= 255 And TextBox6.Text >= 0 And TextBox7.Text <= 255 And TextBox7.Text >= 0 Then
                PictureBox19.BackColor = Color.FromArgb(TextBox5.Text, TextBox6.Text, TextBox7.Text)
            End If
        Else
            PictureBox19.BackColor = Color.FromArgb("255", "255", "255")
        End If
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        If IsNumeric(TextBox5.Text) = True And IsNumeric(TextBox6.Text) = True And IsNumeric(TextBox7.Text) = True Then
            If TextBox5.Text <= 255 And TextBox5.Text >= 0 And TextBox6.Text <= 255 And TextBox6.Text >= 0 And TextBox7.Text <= 255 And TextBox7.Text >= 0 Then
                PictureBox19.BackColor = Color.FromArgb(TextBox5.Text, TextBox6.Text, TextBox7.Text)
            End If
        Else
            PictureBox19.BackColor = Color.FromArgb("255", "255", "255")
        End If
    End Sub

    Private Sub Form3_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub


    Private Function PixelSearch(ByVal a As Integer, ByVal b As Integer, ByVal c As Integer, ByVal d As Integer, ByVal R As Integer, ByVal G As Integer, ByVal Blue As Integer, ByVal errorb As Integer) As Point

        Dim screenshot As Bitmap = New Bitmap(c - a, d - b, PixelFormat.Format24bppRgb) ' 비트맵 선언 (공간)
        Dim gra As Graphics = Graphics.FromImage(screenshot)
        gra.CopyFromScreen(a, b, 0, 0, screenshot.Size, CopyPixelOperation.SourceCopy) ' 선언한 공간 크기만큼 캡쳐


        Dim tBd As BitmapData = screenshot.LockBits(New Rectangle(Point.Empty, screenshot.Size), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb)
        Dim rgbs(Math.Abs(tBd.Stride) * screenshot.Height - 1) As Byte
        Marshal.Copy(tBd.Scan0, rgbs, 0, rgbs.Length)
        screenshot.UnlockBits(tBd)

        Dim csi(screenshot.Width - 1)() As Color
        'Dim fd As Integer

        ' Dim xyp As Point
        ' find = 0
        For x As Integer = 0 To screenshot.Width - 1 Step slp3
            ReDim csi(x)(screenshot.Height - 1)
            For y As Integer = 0 To screenshot.Height - 1 Step slp3

                Dim idx As Integer = (tBd.Stride * y) + (x * 3)
                csi(x)(y) = Color.FromArgb(rgbs(idx + 2), rgbs(idx + 1), rgbs(idx))

                If R = csi(x)(y).R And G = csi(x)(y).G And (Blue - errorb) <= csi(x)(y).B And (Blue + errorb) >= csi(x)(y).B Then
                    '     If (R - errorr) <= csi(x)(y).R And (G - errorg) <= csi(x)(y).G And (Blue - errorb) <= csi(x)(y).B And
                    '    (R + errorr) >= csi(x)(y).R And (G + errorg) >= csi(x)(y).G And (Blue + errorb) >= csi(x)(y).B Then
                    errorlevel = 1

                    Return New Point(x + a, y + b)
                    ' find = find + 1

                    ' If find = 1 Then
                    ' xyp = New Point(x + a, y + b)
                    ' End If
                    ' Return xyp
                    ' ElseIf find > 0 Then

                    ' Return xyp

                End If


            Next
        Next
        ' find = 0

        errorlevel = 0
        Return Nothing
    End Function

    Private Sub TPW_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TPW.KeyPress

        If e.KeyChar = Chr(13) Then
            If PictureBox4.Visible = True Or PictureBox5.Visible = True Then

                Call PictureBox5_Click(sender, e)
            End If
        End If

    End Sub

    Private Sub Label37_MouseMove(sender As Object, e As MouseEventArgs) Handles Label37.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel17.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Label47_MouseMove(sender As Object, e As MouseEventArgs) Handles Label47.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel21.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Label22_MouseMove(sender As Object, e As MouseEventArgs) Handles Label22.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel27.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Label44_MouseMove(sender As Object, e As MouseEventArgs) Handles Label44.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel19.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Label56_MouseMove(sender As Object, e As MouseEventArgs) Handles Label56.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel25.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Label50_MouseMove(sender As Object, e As MouseEventArgs) Handles Label50.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel23.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub
    Public Sub Delay(ByVal mSecond As Integer)
        Dim _startsecond As Integer
        _startsecond = Environment.TickCount
        While (True)
            If (Environment.TickCount - _startsecond >= mSecond) Then 'ms이기 때문에 * 1024로 Delay(1) = 1s가 되게끔
                Exit While
            End If
            Application.DoEvents()
        End While
    End Sub


    Private Sub PictureBox24_MouseClick(sender As Object, e As MouseEventArgs) Handles PictureBox24.MouseClick
        ContextMenuStrip1.Show(PictureBox24, New Point(e.Location.X - 70, e.Location.Y + 20))

    End Sub

    Private Sub TextBox17_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox17.KeyDown
        TextBox17.Text = key
    End Sub



    Private Sub Label60_MouseMove(sender As Object, e As MouseEventArgs) Handles Label60.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel1.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub



    Private Sub Label65_MouseDown(sender As Object, e As MouseEventArgs) Handles Label65.MouseDown
        Panel3.BringToFront()
    End Sub

    Private Sub Label65_MouseMove(sender As Object, e As MouseEventArgs) Handles Label65.MouseMove
        If e.Button = System.Windows.Forms.MouseButtons.Left Then
            ReleaseCapture()

            SendMessage(Panel3.Handle.ToInt32, &HA1S, 2, 0)

        End If
    End Sub

    Private Sub Panel3_MouseDown(sender As Object, e As MouseEventArgs) Handles Panel3.MouseDown
        Panel3.BringToFront()

    End Sub
End Class