Imports System.Runtime.InteropServices
Imports System.Diagnostics
Public Class WindowsHook
   
    Private HistoryId As New List(Of Integer)
    Private HistoryWindowName As New List(Of String)
    Private HistoryProName As New List(Of String)

    Private DegBy As Control
    Private Delay As Integer
    Private MainThred As System.Threading.Thread

    Public Event WindowsClosed(ByVal WindowsName As String, ByVal ProssName As String)
    Public Event WindowsOpend(ByVal WindowsName As String, ByVal ProssName As String)
    Public Event WindowsModifed(ByVal WindowsName As String, ByVal ProssName As String)
    Public Delegate Sub TwoString(ByVal a As String, ByVal b As String)

    Public Sub New(ByRef con As Control, Optional _delay As Integer = 10000)
        LoadHistory()
        DegBy = con
        Delay = _delay
        MainThred = New System.Threading.Thread(AddressOf Runner)
        MainThred.Start()
    End Sub

    Private Sub WinOpenDel(ByVal _a As String, ByVal _b As String)
        If DegBy.InvokeRequired Then
            DegBy.Invoke(New TwoString(AddressOf WinOpenDel), _a, _b)
        Else
            RaiseEvent WindowsOpend(_a, _b)
        End If
    End Sub
    Private Sub WinCloseDel(ByVal _a As String, ByVal _b As String)
        If DegBy.InvokeRequired Then
            DegBy.Invoke(New TwoString(AddressOf WinCloseDel), _a, _b)
        Else
            RaiseEvent WindowsClosed(_a, _b)
        End If
    End Sub
    Private Sub WinModDel(ByVal _a As String, ByVal _b As String)
        If DegBy.InvokeRequired Then
            DegBy.Invoke(New TwoString(AddressOf WinModDel), _a, _b)
        Else
            RaiseEvent WindowsModifed(_a, _b)
        End If
    End Sub
    Private Sub Runner()
        While True
            RunCheck()
            System.Threading.Thread.Sleep(Delay)
        End While
    End Sub
    Public Sub StopAll()
        If MainThred IsNot Nothing Then
            MainThred.Abort()
        End If
    End Sub

    Private Sub LoadHistory()
        HistoryId.Clear()
        HistoryWindowName.Clear()
        HistoryProName.Clear()
        For Each p As Process In Process.GetProcesses()
            ' If p.MainWindowTitle.Length > 0 Then
            HistoryId.Add(p.Id)
            HistoryWindowName.Add(p.MainWindowTitle)
            HistoryProName.Add(p.ProcessName)
            ' End If
        Next
    End Sub

    Public Sub RunCheck()
        Dim TempId As New List(Of Integer)
        Dim TempWindowName As New List(Of String)
        Dim TempProName As New List(Of String)
        For Each p As Process In Process.GetProcesses()
            ' If p.MainWindowTitle.Length > 0 Then
            TempId.Add(p.Id)
            TempWindowName.Add(p.MainWindowTitle)
            TempProName.Add(p.ProcessName)
            '  End If
        Next
        Dim temp As Integer
        For i As Integer = 0 To TempId.Count - 1 Step 1
            temp = HistoryId.IndexOf(TempId.Item(i))
            If temp <> -1 Then
                If HistoryWindowName.Item(temp) <> TempWindowName.Item(i) Then
                    'WinModDel(TempWindowName.Item(i), TempProName.Item(i))
                    RaiseEvent WindowsModifed(TempWindowName.Item(i), TempProName.Item(i))
                End If
                HistoryWindowName.RemoveAt(temp)
                HistoryProName.RemoveAt(temp)
                HistoryId.RemoveAt(temp)
            Else
                ' WinOpenDel(TempWindowName.Item(i), TempProName.Item(i))
                RaiseEvent WindowsOpend(TempWindowName.Item(i), TempProName.Item(i))
            End If
        Next
        For i As Integer = 0 To HistoryId.Count - 1 Step 1
            'WinCloseDel(HistoryWindowName.Item(i), HistoryProName.Item(i))
            RaiseEvent WindowsClosed(TempWindowName.Item(i), TempProName.Item(i))
        Next
        HistoryWindowName = TempWindowName
        HistoryId = TempId
        HistoryProName = TempProName
    End Sub
End Class
