Public Class FullHook

    Private MouseListener As MouseHook
    Private KeyListener As KeyboardHook

    Public Sub Start(ByRef InvokeBy As Control)
        DropAll()

        MouseListener = New MouseHook()

        AddHandler MouseListener.Mouse_Move, AddressOf Mouse_Move


        If Form3.CheckBox1.Checked = True And Form3.RadioButton3.Checked = True Then

            AddHandler MouseListener.Mouse_Move, AddressOf Overlay

        End If



        If Form3.CheckBox2.Checked = True Then
            Form3.check = True
        Else
            If Form3.RadioButton1.Checked = True Then
                AddHandler MouseListener.Mouse_Left_Down, AddressOf Mouse_Left_Down
                AddHandler MouseListener.Mouse_Left_Up, AddressOf Mouse_Left_Up

            End If
            If Form3.RadioButton2.Checked = True Then
                AddHandler MouseListener.Mouse_Right_Down, AddressOf Mouse_Right_Down
                AddHandler MouseListener.Mouse_Right_Up, AddressOf Mouse_Right_Up
            End If
        End If




    End Sub
    Public Sub kstart(ByRef InvokeBy As Control)
        kDropAll()

        KeyListener = New KeyboardHook()

        AddHandler KeyListener.KeyDown, AddressOf KeyDown





    End Sub


    Public Sub DropAll()
        If MouseListener IsNot Nothing Then
            RemoveHandler MouseListener.Mouse_Move, AddressOf Mouse_Move
            RemoveHandler MouseListener.Mouse_Move, AddressOf Overlay

            RemoveHandler MouseListener.Mouse_Left_Down, AddressOf Mouse_Left_Down
            RemoveHandler MouseListener.Mouse_Left_Up, AddressOf Mouse_Left_Up

            RemoveHandler MouseListener.Mouse_Right_Down, AddressOf Mouse_Right_Down
            RemoveHandler MouseListener.Mouse_Right_Up, AddressOf Mouse_Right_Up

            MouseListener.Dispose()
            MouseListener = Nothing
        End If





    End Sub
    Public Sub kDropAll()


        If KeyListener IsNot Nothing Then
            RemoveHandler KeyListener.KeyDown, AddressOf KeyDown
            KeyListener.Dispose()
            KeyListener = Nothing
        End If





    End Sub

    '======================================================handles====================================================

    Public Sub Mouse_Move(ByVal ptLocat As Point)
        Form3.aim = ptLocat

    End Sub

    Public Sub Overlay(ByVal ptLocat As Point)

        Form2.Top = ptLocat.Y - Form2.Height / 2 - Form3.yjo
        Form2.Left = ptLocat.X - Form2.Width / 2

    End Sub

    Private Sub Mouse_Left_Down(ByVal ptLocat As Point)
        Form3.check = True

    End Sub

    Private Sub Mouse_Left_Up(ByVal ptLocat As Point)
        Form3.check = False

    End Sub

    Private Sub Mouse_Right_Down(ByVal ptLocat As Point)
        Form3.check = True

    End Sub

    Private Sub Mouse_Right_Up(ByVal ptLocat As Point)
        Form3.check = False

    End Sub

    Private Sub KeyDown(ByVal Key As Keys)

        If Form3.TextBox17.Text = Key.ToString Then

            Dim sender As Object, e As EventArgs

            If Form3.PictureBox20.Visible = True Then
                Call Form3.PictureBox20_Click(sender, e)
            Else
                Call Form3.PictureBox21_Click(sender, e)
            End If
        End If

        Form3.key = Key.ToString
        ' MsgBox(Form3.key)
    End Sub

End Class
