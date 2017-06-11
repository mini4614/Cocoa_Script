Public Class FullHook
    Implements IDisposable

    Private KeyListener As KeyboardHook
    Private MouseListener As MouseHook
    Private WinListener As WindowsHook

    Private CurrentDataSession As DataSession

    Public Enum MouseState
        L = 0
        DL = 1
        R = 2
        DR = 3
        M = 4
        DM = 5
    End Enum
    Public Enum WinState
        O = 0
        C = 1
        M = 2
    End Enum
    Public Sub Start(ByRef InvokeBy As Control, ByVal _DataSiosn As DataSession)
        DropAll()
        KeyListener = New KeyboardHook()
        MouseListener = New MouseHook()
        WinListener = New WindowsHook(InvokeBy, 1000)
        SetAllConnections(_DataSiosn)
    End Sub
    Public Sub NewDataSession(ByRef _NewData As DataSession)
        CurrentDataSession = _NewData
    End Sub
    Public Sub StopAll()
        DropAll()
    End Sub
    Private Sub SetAllConnections(ByRef _DataSession As DataSession)
        CurrentDataSession = _DataSession

        AddHandler KeyListener.KeyDown, AddressOf KeyDown

        AddHandler MouseListener.Mouse_Left_Down, AddressOf Mouse_Left_Down
        'AddHandler MouseListener.Mouse_Left_DoubleClick, AddressOf Mouse_Left_DoubleClick
        '   AddHandler MouseListener.Mouse_Right_Down, AddressOf Mouse_Right_Down
        'AddHandler MouseListener.Mouse_Right_DoubleClick, AddressOf Mouse_Right_DoubleClick
        'AddHandler MouseListener.Mouse_Middle_Down, AddressOf Mouse_Middle_Down
        ' AddHandler MouseListener.Mouse_Middle_DoubleClick, AddressOf Mouse_Middle_DoubleClick
        '  AddHandler MouseListener.Mouse_Wheel, AddressOf Mouse_Wheel

        ' AddHandler WinListener.WindowsOpend, AddressOf WindowsOpend
        ' AddHandler WinListener.WindowsClosed, AddressOf WindowsClosed
        '  AddHandler WinListener.WindowsModifed, AddressOf WindowsModifed
    End Sub
    Private Sub UnSetAllConnection()
        If KeyListener IsNot Nothing Then
            RemoveHandler KeyListener.KeyDown, AddressOf KeyDown
        End If

        If MouseListener IsNot Nothing Then
            RemoveHandler MouseListener.Mouse_Left_Down, AddressOf Mouse_Left_Down
            '    RemoveHandler MouseListener.Mouse_Left_DoubleClick, AddressOf Mouse_Left_DoubleClick
            '    RemoveHandler MouseListener.Mouse_Right_Down, AddressOf Mouse_Right_Down
            '    RemoveHandler MouseListener.Mouse_Right_DoubleClick, AddressOf Mouse_Right_DoubleClick
            '  RemoveHandler MouseListener.Mouse_Middle_Down, AddressOf Mouse_Middle_Down
            '     RemoveHandler MouseListener.Mouse_Middle_DoubleClick, AddressOf Mouse_Middle_DoubleClick
            '   RemoveHandler MouseListener.Mouse_Wheel, AddressOf Mouse_Wheel
        End If
        If WinListener IsNot Nothing Then
            '   RemoveHandler WinListener.WindowsOpend, AddressOf WindowsOpend
            '  RemoveHandler WinListener.WindowsClosed, AddressOf WindowsClosed
            '  RemoveHandler WinListener.WindowsModifed, AddressOf WindowsModifed
        End If
    End Sub
    Private Sub DropAll()
        UnSetAllConnection()
        If KeyListener IsNot Nothing Then
            KeyListener.Dispose()
            KeyListener = Nothing
        End If
        If MouseListener IsNot Nothing Then
            MouseListener.Dispose()
            MouseListener = Nothing
        End If
        If WinListener IsNot Nothing Then
            WinListener.StopAll()
            WinListener = Nothing
        End If
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        DropAll()
    End Sub

    '======================================================handles====================================================
    ' Private Sub WindowsClosed(ByVal WindowsName As String, ByVal ProssName As String)
    '     If CurrentDataSession IsNot Nothing Then
    '         CurrentDataSession.AddWinEvent(WinState.C, WindowsName, ProssName)
    '     End If
    ' End Sub
    '  Private Sub WindowsOpend(ByVal WindowsName As String, ByVal ProssName As String)
    '     If CurrentDataSession IsNot Nothing Then
    '          CurrentDataSession.AddWinEvent(WinState.O, WindowsName, ProssName)
    '       End If
    '  End Sub
    ' Private Sub WindowsModifed(ByVal WindowsName As String, ByVal ProssName As String)
    '     If CurrentDataSession IsNot Nothing Then
    '        CurrentDataSession.AddWinEvent(WinState.M, WindowsName, ProssName)
    '    End If
    ' End Sub
    Private Sub KeyDown(ByVal Key As Keys)
        If CurrentDataSession IsNot Nothing Then
            CurrentDataSession.AddKeyboradEvent(Key)
        End If
    End Sub
    Private Sub Mouse_Left_Down(ByVal ptLocat As Point)
        If CurrentDataSession IsNot Nothing Then
            CurrentDataSession.AddMouseEvent(MouseState.L, ptLocat)
        End If
    End Sub
    ' Private Sub Mouse_Left_DoubleClick(ByVal ptLocat As Point)
    '      If CurrentDataSession IsNot Nothing Then
    '          CurrentDataSession.AddMouseEvent(MouseState.DL, ptLocat)
    '     End If
    ' End Sub
    ' Private Sub Mouse_Right_Down(ByVal ptLocat As Point)
    '     If CurrentDataSession IsNot Nothing Then
    '          CurrentDataSession.AddMouseEvent(MouseState.R, ptLocat)
    '      End If
    '   End Sub
    '  Private Sub Mouse_Right_DoubleClick(ByVal ptLocat As Point)
    '      If CurrentDataSession IsNot Nothing Then
    '          CurrentDataSession.AddMouseEvent(MouseState.DR, ptLocat)
    '      End If
    '  End Sub
    '  Private Sub Mouse_Middle_Down(ByVal ptLocat As Point)
    '  If CurrentDataSession IsNot Nothing Then
    '         CurrentDataSession.AddMouseEvent(MouseState.M, ptLocat)
    '      End If
    '  End Sub
    '  Private Sub Mouse_Middle_DoubleClick(ByVal ptLocat As Point)
    '     If CurrentDataSession IsNot Nothing Then
    '         CurrentDataSession.AddMouseEvent(MouseState.DM, ptLocat)
    '    End If
    '   End Sub
    '  Private Sub Mouse_Wheel(ByVal ptLocat As Point, ByVal Direction As MouseHook.Wheel_Direction)
    '   If CurrentDataSession IsNot Nothing Then
    '         CurrentDataSession.AddScrolEvent(Direction, ptLocat)
    '    End If
    'End Sub
End Class
