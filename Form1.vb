Public Class Form1
    Dim kbHook As FullHook
    Dim AllData As DataSession
    Dim a
    Public Running As Boolean = False
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not Running Then
            kbHook = New FullHook()
            AllData = New DataSession(True, Me)
            AddHandler AllData.ActionAdded, AddressOf NewEvent
            kbHook.Start(Me, AllData)
            Button1.Text = "Unhook Computer"
        Else
            kbHook.StopAll()
            kbHook.Dispose()
            kbHook = Nothing
            AllData = Nothing
            Button1.Text = "Hook Computer"
        End If
        Running = Not Running
    End Sub


    Private Sub NewEvent(ByRef _action As OneAction)
        a = _action.ToString
        ListBox1.Items.Add(a)
        ListBox1.TopIndex = ListBox1.Items.Count - 1
        '  MsgBox(_action.ToString)
        If a = "F1" Then

            If CheckBox1.Checked = True Then
                CheckBox1.Checked = False
            Else
                CheckBox1.Checked = True
            End If
        End If


    End Sub



    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If kbHook IsNot Nothing Then
            kbHook.Dispose()
        End If
    End Sub
End Class
