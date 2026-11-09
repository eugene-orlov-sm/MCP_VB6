Attribute VB_Name = "NativeDebugger"
Option Explicit

Private Declare Function SendInput Lib "user32" (ByVal count As Long, ByRef inputs As Any, ByVal size As Long) As Long
Private Declare Function VkKeyScanW Lib "user32" (ByVal character As Integer) As Integer
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, ByRef pid As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal key As Long) As Integer
Private Declare Function EnumWindows Lib "user32" (ByVal callback As Long, ByVal param As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hwnd As Long, ByVal callback As Long, ByVal param As Long) As Long
Private Declare Function GetClassNameW Lib "user32" (ByVal hwnd As Long, ByVal text As Long, ByVal size As Long) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hwnd As Long, ByVal text As Long, ByVal size As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal command As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsChild Lib "user32" (ByVal parent As Long, ByVal hwnd As Long) As Long

Private m_inputs() As Byte
Private m_count As Long
Private m_main As Long
Private m_pane As Long
Private m_want As String


' Enumerate only windows owned by this IDE without entering VBIDE COM
Public Function ScanWindow(ByVal hwnd As Long, ByVal param As Long) As Long
    ScanWindow = 1
    If Not OwnWindow(hwnd) Then Exit Function
    Dim buffer As String, size As Long, kind As String, caption As String
    buffer = String$(260, vbNullChar)
    size = GetClassNameW(hwnd, StrPtr(buffer), 260)
    kind = LCase$(Left$(buffer, size))
    If kind = "wndclass_desked_gsk" Then m_main = hwnd
    If kind <> "vbawindow" Then Exit Function
    buffer = String$(260, vbNullChar)
    size = GetWindowTextW(hwnd, StrPtr(buffer), 260)
    caption = LCase$(Left$(buffer, size))
    If caption = m_want Then m_pane = hwnd
End Function


Private Function FindPane(ByVal name As String) As Long
    m_main = 0
    m_pane = 0
    m_want = LCase$(name)
    EnumWindows AddressOf ScanWindow, 0
    If m_main <> 0 Then EnumChildWindows m_main, AddressOf ScanWindow, 0
    If m_want = "mainwindow" Then FindPane = m_main Else FindPane = m_pane
End Function


Private Sub ShowPane(ByVal hwnd As Long)
    Dim current As Long
    current = hwnd
    Do While OwnWindow(current)
        ShowWindow current, 4
        If current = m_main Then Exit Do
        current = GetParent(current)
    Loop
End Sub


' Show existing debugger panes without calling the suspended VBIDE window model
Public Function ControlWindow(ByVal name As String, ByVal action As String) As Object
    Dim result As Object, hwnd As Long
    Set result = CreateObject("Scripting.Dictionary")
    result("tool") = "control_window"
    result("ok") = False
    result("commandBarsTouched") = False
    result("windowModelTouched") = False
    result("state") = "break"
    result("foregroundBefore") = GetForegroundWindow()
    hwnd = FindPane(name)
    If hwnd = 0 Then
        result("reason") = "debugger_window_unavailable"
        result("howToProceed") = "Open this pane in DESIGN before running; unsupported paused window operations are refused without entering VBIDE COM"
    ElseIf LCase$(action) = "show" Or LCase$(action) = "hide" Then
        If LCase$(action) = "show" Then ShowPane hwnd Else ShowWindow hwnd, 0
        result("ok") = (CBool(IsWindowVisible(hwnd)) = (LCase$(action) = "show"))
        result("hwnd") = hwnd
        result("visible") = CBool(IsWindowVisible(hwnd))
        result("delivery") = "native_window"
    Else
        result("reason") = "paused_window_action_unsupported"
    End If
    result("foregroundAfter") = GetForegroundWindow()
    result("foregroundTookFocus") = (result("foregroundBefore") <> result("foregroundAfter"))
    result("paneAssertion") = "notApplicable"
    result("paneAssertionNote") = "Native debugger-pane visibility does not enter or edit code panes"
    Set ControlWindow = result
End Function


Private Sub AddKey(ByVal key As Long, Optional ByVal released As Boolean = False)
    Dim offset As Long, i As Long
    offset = m_count * 28
    ReDim Preserve m_inputs(0 To offset + 27)
    For i = offset To offset + 27
        m_inputs(i) = 0
    Next
    m_inputs(offset) = 1
    m_inputs(offset + 4) = key And &HFF&
    If released Then m_inputs(offset + 8) = 2
    m_count = m_count + 1
End Sub


Private Sub AddStroke(ByVal key As Long, Optional ByVal modifiers As Long = 0)
    If (modifiers And 1) <> 0 Then AddKey &H10
    If (modifiers And 2) <> 0 Then AddKey &H11
    If (modifiers And 4) <> 0 Then AddKey &H12
    AddKey key
    AddKey key, True
    If (modifiers And 4) <> 0 Then AddKey &H12, True
    If (modifiers And 2) <> 0 Then AddKey &H11, True
    If (modifiers And 1) <> 0 Then AddKey &H10, True
End Sub


Private Function OwnWindow(ByVal hwnd As Long) As Boolean
    Dim pid As Long
    If hwnd = 0 Then Exit Function
    If IsWindow(hwnd) = 0 Then Exit Function
    GetWindowThreadProcessId hwnd, pid
    OwnWindow = (pid = GetCurrentProcessId())
End Function


Private Function PrepareInput(Owner As CodeTools, ByVal target As Long, ByRef oldForeground As Long, ByRef oldFocus As Long, ByRef reason As String) As Boolean
    Dim key As Variant
    reason = ""
    For Each key In Array(&H10, &H11, &H12, &H5B, &H5C)
        If GetAsyncKeyState(CLng(key)) < 0 Then
            reason = "modifier_key_held"
            Exit Function
        End If
    Next
    If Not OwnWindow(target) Then
        reason = "native_input_target_unavailable"
        Exit Function
    End If
    oldForeground = GetForegroundWindow()
    oldFocus = GetFocus()
    SetForegroundWindow m_main
    If Not OwnWindow(GetForegroundWindow()) Then
        reason = "native_input_foreground_unavailable"
        Exit Function
    End If
    SetFocus target
    If GetFocus() <> target And (target <> m_main Or IsChild(m_main, GetFocus()) = 0) Then
        reason = "native_input_focus_unavailable"
        RestoreInput oldForeground, oldFocus
        Exit Function
    End If
    PrepareInput = True
End Function


Private Sub RestoreInput(ByVal foreground As Long, ByVal focus As Long)
    If OwnWindow(focus) Then SetFocus focus
    If foreground <> 0 Then
        If IsWindow(foreground) <> 0 Then SetForegroundWindow foreground
    End If
End Sub


Private Function SendKeysNow(ByRef reason As String) As Boolean
    Dim sent As Long
    sent = SendInput(m_count, m_inputs(0), 28)
    SendKeysNow = (sent = m_count)
    If Not SendKeysNow Then
        reason = "native_input_incomplete"
        m_count = 0
        AddKey &H10, True
        AddKey &H11, True
        AddKey &H12, True
        SendInput m_count, m_inputs(0), 28
    End If
End Function


' Execute debugger shortcuts through the native input dispatcher
Public Function ExecuteCommand(Owner As CodeTools, ByVal mode As String, ByRef reason As String) As Boolean
    Dim key As Long, modifiers As Long
    Select Case LCase$(mode)
        Case "continue", "resume", "run": key = &H74
        Case "step_into", "stepinto": key = &H77
        Case "step_over", "stepover": key = &H77: modifiers = 1
        Case "step_out", "stepout": key = &H77: modifiers = 3
        Case "show_immediate": key = &H47: modifiers = 2
        Case Else: reason = "native_command_unsupported": Exit Function
    End Select
    Dim oldForeground As Long, oldFocus As Long
    Dim target As Long
    target = FindPane("MainWindow")
    If Not PrepareInput(Owner, target, oldForeground, oldFocus, reason) Then Exit Function
    On Error GoTo Failed
    m_count = 0
    AddStroke key, modifiers
    If Not SendKeysNow(reason) Then GoTo Done
    ' Dispatch while the keyboard focus still belongs to this IDE
    DoEvents
    ExecuteCommand = True
Done:
    RestoreInput oldForeground, oldFocus
    Exit Function
Failed:
    reason = "native_command_error: " & Err.Description
    Resume Done
End Function


Private Function ClosedSink(ByVal path As String) As Boolean
    Dim file As Integer
    On Error GoTo NotReady
    If Len(Dir$(path)) = 0 Then Exit Function
    file = FreeFile
    Open path For Binary Access Read Lock Read Write As #file
    ClosedSink = (LOF(file) > 0)
    Close #file
    Exit Function
NotReady:
    On Error Resume Next
    If file <> 0 Then Close #file
End Function


' Evaluate in the real Immediate window so locals use the paused frame
Public Function EvaluateLine(Owner As CodeTools, ByVal line As String, ByVal sink As String, ByVal timeoutMs As Long, ByRef reason As String) As Boolean
    Dim pane As Long
    pane = FindPane("Immediate")
    If pane = 0 Then
        If Not ExecuteCommand(Owner, "show_immediate", reason) Then Exit Function
        pane = FindPane("Immediate")
        If pane = 0 Then
            reason = "immediate_window_unavailable"
            Exit Function
        End If
    End If
    ShowPane pane
    Dim i As Long, code As Integer, stroke As Integer
    m_count = 0
    ' Append on a new line without executing any pending Immediate input
    AddStroke &H23, 2
    AddStroke &HD, 2
    For i = 1 To Len(line)
        code = AscW(Mid$(line, i, 1))
        stroke = VkKeyScanW(code)
        If stroke = -1 Then
            reason = "native_input_character_unavailable"
            Exit Function
        End If
        AddStroke CLng(stroke) And &HFF&, (CLng(stroke) And &HFF00&) \ &H100&
    Next
    AddStroke &HD
    Dim oldForeground As Long, oldFocus As Long
    If Not PrepareInput(Owner, pane, oldForeground, oldFocus, reason) Then Exit Function
    On Error GoTo Failed
    If Not SendKeysNow(reason) Then GoTo Done
    Dim started As Long
    started = GetTickCount()
    Do
        DoEvents
        If ClosedSink(sink) Then
            EvaluateLine = True
            Exit Do
        End If
        If GetTickCount() - started >= timeoutMs Then
            reason = "native_immediate_timeout"
            Exit Do
        End If
    Loop
Done:
    RestoreInput oldForeground, oldFocus
    Exit Function
Failed:
    reason = "native_immediate_error: " & Err.Description
    Resume Done
End Function
