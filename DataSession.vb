Public Class DataSession

    Public StartingTime As Date
    Public EventList As List(Of OneAction)
    Public LastEvent As Date
    Private teapTime As Date
    Private tempAction As OneAction
    Private EnableEvent As Boolean
    Private InvokeBy As Control


    Public Event ActionAdded(ByRef _Action As OneAction)


    Public Sub New(Optional _EnableEvent As Boolean = False, Optional ByRef _InvokeBy As Control = Nothing)
        EnableEvent = _EnableEvent
        InvokeBy = _InvokeBy
        StartingTime = Now
        LastEvent = Now
        EventList = New List(Of OneAction)
    End Sub

    Private Delegate Sub NoneInput(ByRef _Action As OneAction)
    Private Sub DoEvent(ByRef _Action As OneAction)
        If InvokeBy.InvokeRequired Then
            InvokeBy.Invoke(New NoneInput(AddressOf DoEvent), _Action)
        Else
            RaiseEvent ActionAdded(_Action)
        End If
    End Sub

    Public Sub AddKeyboradEvent(ByVal Downkey As Keys)
        teapTime = Now
        tempAction = New KeyAction(Downkey, (teapTime - LastEvent).TotalMilliseconds)
        EventList.Add(tempAction)
        LastEvent = teapTime
        If EnableEvent Then
            DoEvent(tempAction)
        End If
    End Sub
    Public Sub AddWinEvent(ByVal State As FullHook.WinState, ByVal WinName As String, ByVal ProName As String)
        teapTime = Now
        tempAction = New WinAction(State, WinName, ProName, (teapTime - LastEvent).TotalMilliseconds)
        EventList.Add(tempAction)
        LastEvent = teapTime
        If EnableEvent Then
            DoEvent(tempAction)
        End If
    End Sub
    Public Sub AddMouseEvent(ByVal ActionState As FullHook.MouseState, ByVal p As Point)
        teapTime = Now
        tempAction = New MouseAction(ActionState, p, (teapTime - LastEvent).TotalMilliseconds)
        EventList.Add(tempAction)
        LastEvent = teapTime
        If EnableEvent Then
            DoEvent(tempAction)
        End If
    End Sub
    Public Sub AddScrolEvent(ByVal dir As MouseHook.Wheel_Direction, ByVal p As Point)
        teapTime = Now
        tempAction = New MouseAction(MouseAction.MouseAction.S, p, (teapTime - LastEvent).TotalMilliseconds, dir)
        EventList.Add(tempAction)
        LastEvent = teapTime
        If EnableEvent Then
            DoEvent(tempAction)
        End If
    End Sub
    Public Enum ActionType
        Keyboard = 0
        Win = 1
        Mouse = 2
        TimeSet = 3
    End Enum
End Class

Public MustInherit Class OneAction
    Public ActionState As DataSession.ActionType
    Public Delay As Long
    Public Sub New(ByVal _ActionState As DataSession.ActionType, ByVal _Delay As Integer)
        ActionState = _ActionState
        Delay = _Delay
    End Sub

    Public MustOverride Overrides Function ToString() As String

End Class
Public Class TimeSet
    Inherits OneAction
    Public MainTime As Date

    Public Sub New(ByVal _MainTime As Date)
        MyBase.New(DataSession.ActionType.TimeSet, 0)
        MainTime = _MainTime
    End Sub

    Public Overrides Function toString() As String
        Return "New Time >> " & MainTime.ToString
    End Function
End Class
Public Class WinAction
    Inherits OneAction

    Public WinActionType As FullHook.WinState
    Public WindowsName As String
    Public ProName As String
    Public Sub New(ByVal _WinActionType As FullHook.WinState, ByVal _WindowsName As String, ByVal _ProName As String, ByVal _Delay As Integer)
        MyBase.New(DataSession.ActionType.Win, _Delay)
        WindowsName = _WindowsName
        ProName = _ProName
        WinActionType = _WinActionType
    End Sub

    Public Overrides Function ToString() As String
        Select Case WinActionType
            Case FullHook.WinState.C
                Return "Win Close >> " & ProName & " -> " & WindowsName
            Case FullHook.WinState.M
                Return "Win Modify >> " & ProName & " -> " & WindowsName
            Case FullHook.WinState.O
                Return "Win Open >> " & ProName & " -> " & WindowsName
        End Select
        Throw New Exception("Bad Window Action Type")
    End Function
End Class
Public Class MouseAction
    Inherits OneAction

    Public MainType As MouseAction
    Public p As Point

    Public ScrolType As ScrolDir

    Public Enum MouseAction
        L = 0
        DL = 1
        R = 2
        DR = 3
        M = 4
        DM = 5
        S = 6
    End Enum
    Public Enum ScrolDir
        up = 0
        down = 1
    End Enum
    Public Sub New(ByVal _MainType As MouseAction, ByVal _p As Point, ByVal _Delay As Integer, Optional ByVal dir As MouseHook.Wheel_Direction = MouseHook.Wheel_Direction.WheelUp)
        MyBase.New(DataSession.ActionType.Mouse, _Delay)
        MainType = _MainType
        p = _p
        ScrolType = dir
    End Sub
    Public Overrides Function toString() As String
        Select Case MainType
            Case MouseAction.L
                Return p.ToString
            Case MouseAction.DL
                Return p.ToString
            Case MouseAction.R
                Return p.ToString
            Case MouseAction.DR
                Return p.ToString
            Case MouseAction.M
                Return p.ToString
            Case MouseAction.DM
                Return p.ToString
            Case MouseAction.S
                Select Case ScrolType
                    Case ScrolDir.down
                        Return p.ToString
                    Case ScrolDir.up
                        Return p.ToString
                End Select
        End Select
        Throw New Exception("Bad Mouse Type")
    End Function
End Class
Public Class KeyAction
    Inherits OneAction

    Public MainKey As Keys
    Public StartingValue As String

    Public Sub New(ByVal _MainKey As Keys, ByVal _Delay As Integer)
        MyBase.New(DataSession.ActionType.Keyboard, _Delay)
        MainKey = _MainKey

        If My.Computer.Keyboard.CapsLock Then
            StartingValue = MainKey.ToString.ToUpper
        Else
            StartingValue = MainKey.ToString.ToLower
        End If

    End Sub

    Public Overrides Function ToString() As String
        Return MainKey.ToString
    End Function
End Class

