Attribute VB_Name = "ModGestShortCut"
Option Explicit

Const MOUSEEVENTF_LEFTDOWN = &H2                ' left button down
Const MOUSEEVENTF_LEFTUP = &H4                  ' left button up
Const MOUSEEVENTF_RIGHTDOWN = &H8               ' right button down
Const MOUSEEVENTF_RIGHTUP = &H10                ' right button up
Const MOUSEEVENTF_MOVE = &H1
Const MOUSEEVENTF_ABSOLUTE = &H8000


Private Const KEYEVENTF_KEYUP = &H2
'Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
'Private Const INPUT_HARDWARE = 2

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
    End Type

'Private Type HARDWAREINPUT
'    uMsg As Long
'    wParamL As Integer
'    wParamH As Integer
'    End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
    End Type

Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)


'Public Type Rect
'  left As Long
'  top As Long
'  right As Long
'  bottom As Long
'End Type
'Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As Rect) As Long
'
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)


Public Sub MouseClick(Button As Integer, X As Single, Y As Single)
'Dim inputevents(0 To 2) As GENERALINPUT  ' holds information about each event
'Dim keyevent As KEYBDINPUT  ' temporarily hold keyboard input info
'Dim mouseevent As MOUSEINPUT  ' temporarily hold mouse input info
'Dim r As Rect  ' receives window rectangle
    
    SetCursorPos X, Y
    
    If Button = vbKeyRButton Then
        mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
    Else
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    End If

'**+ Con SendKeys o con Send_input il click nn funziona **/
'    SendKeys vbKeyLButton
'
'    'GetWindowRect Frm.hwnd, r  ' set r equal to Form1's rectangle
'
'    x = x * Screen.TwipsPerPixelX
'    y = y * Screen.TwipsPerPixelY
'
'    ' Load the information needed to synthesize pressing the right mouse button.
'    mouseevent.dx = x 'r.left + 120 ' no horizontal movement
'    mouseevent.dy = y 'r.top + 120 'r.bottom * Screen.TwipsPerPixelY  ' no vertical movement
'    mouseevent.mouseData = 0  ' not needed
'    mouseevent.dwFlags = MOUSEEVENTF_MOVE ' right button down
'    mouseevent.dwtime = 0  ' use the default
'    mouseevent.dwExtraInfo = 0  ' not needed
'    ' Copy the structure into the input array's buffer.
'    inputevents(0).dwType = INPUT_MOUSE  ' mouse input
'    CopyMemory inputevents(0).xi(0), mouseevent, Len(mouseevent)
'
'    ' Load the information needed to synthesize pressing the right mouse button.
'    mouseevent.dx = x 'r.left + 120 ' no horizontal movement
'    mouseevent.dy = y 'r.top + 120 'r.bottom * Screen.TwipsPerPixelY  ' no vertical movement
'    mouseevent.mouseData = 0  ' not needed
'    mouseevent.dwFlags = MOUSEEVENTF_LEFTDOWN ' right button down
'    mouseevent.dwtime = 0  ' use the default
'    mouseevent.dwExtraInfo = 0  ' not needed
'    ' Copy the structure into the input array's buffer.
'    inputevents(1).dwType = INPUT_MOUSE  ' mouse input
'    CopyMemory inputevents(1).xi(0), mouseevent, Len(mouseevent)
'
'    ' Do the same as above, but for releasing the right mouse button.
'    mouseevent.dx = x 'r.left + 120 ' no horizontal movement
'    mouseevent.dy = y 'r.top + 120 ' no vertical movement
'    mouseevent.mouseData = 0  ' not needed
'    mouseevent.dwFlags = MOUSEEVENTF_LEFTUP ' right button up
'    mouseevent.dwtime = 0  ' use the default
'    mouseevent.dwExtraInfo = 0  ' not needed
'    ' Copy the structure into the input array's buffer.
'    inputevents(2).dwType = INPUT_MOUSE  ' mouse input
'    CopyMemory inputevents(2).xi(0), mouseevent, Len(mouseevent)
'
'    ' Now that all the information for the four input events has been placed
'    ' into the array, finally send it into the input stream.
'    SendInput 3, inputevents(0), Len(inputevents(0))  ' place the events into the stream
'**********************************************************************************/
    
End Sub

Public Function GetShift(ByVal Tool As ActiveBar3LibraryCtl.Tool) As Integer
    Dim iShift As Integer
    Dim sValue As String
    
    On Error GoTo errHandler
    
    sValue = Tool.ShortCuts(0)
    
    If InStr(sValue, "Ctrl") > 0 Then
        iShift = vbCtrlMask
    End If
    If InStr(sValue, "Shift") > 0 Then
        iShift = iShift + vbShiftMask
    End If
    If InStr(sValue, "Alt") > 0 Then
        iShift = iShift + vbAltMask
    End If
    
    GetShift = iShift
    
    Exit Function
    
errHandler:
    GetShift = 0
End Function

Public Function GetSendKeys(ByVal Tool As ActiveBar3LibraryCtl.Tool) As String
    Dim iShift As Integer
    Dim sStr  As String

    iShift = GetShift(Tool)
    
    If (iShift And vbShiftMask = iShift) Then
        sStr = "+"
    End If
    If (iShift And vbCtrlMask = iShift) Then
        sStr = sStr & "^"
    End If
    If (iShift And vbAltMask = iShift) Then
        sStr = sStr & "%"
    End If

    GetSendKeys = sStr
End Function

Public Function GetKeyCode(ByVal Tool As ActiveBar3LibraryCtl.Tool) As KeyCodeConstants
    Dim iKeyCode As KeyCodeConstants
    Dim sValue As String
    Dim iPos As Integer
    
    On Error GoTo errHandler
    
    sValue = Tool.ShortCuts(0)
    
    iPos = InStrRev(sValue, "+")
    If iPos > 0 Then
        sValue = Right$(sValue, Len(sValue) - iPos)
    End If
    
    iKeyCode = MapKeyCode(sValue)
    
    GetKeyCode = iKeyCode
    
    Exit Function
    
errHandler:
    GetKeyCode = 0
End Function

Public Function GetKey(ByVal Tool As ActiveBar3LibraryCtl.Tool) As String
    Dim sValue As String
    Dim iPos As Integer
    
    On Error GoTo errHandler
    
    sValue = Tool.ShortCuts(0)
    
    iPos = InStrRev(sValue, "+")
    If iPos > 0 Then
        sValue = Right$(sValue, Len(sValue) - iPos)
    End If
    
    Select Case sValue
        Case "CANCELLA", "DELETE", "CANC"
            sValue = ""
    End Select
    
    GetKey = sValue
    
    Exit Function
    
errHandler:
    GetKey = 0
End Function

Private Function MapKeyCode(ByVal Value As String) As KeyCodeConstants
    Select Case UCase(Value)
        Case "A"
            MapKeyCode = vbKeyA
        Case "B"
            MapKeyCode = vbKeyB
        Case "C"
            MapKeyCode = vbKeyC
        Case "D"
            MapKeyCode = vbKeyD
        Case "E"
            MapKeyCode = vbKeyE
        Case "F"
            MapKeyCode = vbKeyF
        Case "G"
            MapKeyCode = vbKeyG
        Case "H"
            MapKeyCode = vbKeyH
        Case "I"
            MapKeyCode = vbKeyI
        Case "J"
            MapKeyCode = vbKeyJ
        Case "K"
            MapKeyCode = vbKeyK
        Case "L"
            MapKeyCode = vbKeyL
        Case "M"
            MapKeyCode = vbKeyM
        Case "N"
            MapKeyCode = vbKeyN
        Case "O"
            MapKeyCode = vbKeyO
        Case "P"
            MapKeyCode = vbKeyP
        Case "Q"
            MapKeyCode = vbKeyQ
        Case "R"
            MapKeyCode = vbKeyR
        Case "S"
            MapKeyCode = vbKeyS
        Case "T"
            MapKeyCode = vbKeyT
        Case "U"
            MapKeyCode = vbKeyU
        Case "V"
            MapKeyCode = vbKeyV
        Case "W"
            MapKeyCode = vbKeyW
        Case "X"
            MapKeyCode = vbKeyX
        Case "Y"
            MapKeyCode = vbKeyY
        Case "Z"
            MapKeyCode = vbKeyZ
        Case "0"
            MapKeyCode = vbKey0
        Case "1"
            MapKeyCode = vbKey1
        Case "2"
            MapKeyCode = vbKey2
        Case "3"
            MapKeyCode = vbKey3
        Case "4"
            MapKeyCode = vbKey4
        Case "5"
            MapKeyCode = vbKey5
        Case "6"
            MapKeyCode = vbKey6
        Case "7"
            MapKeyCode = vbKey7
        Case "8"
            MapKeyCode = vbKey8
        Case "9"
            MapKeyCode = vbKey9
        Case "0 (TN)"
            MapKeyCode = vbKeyNumpad0
        Case "1 (TN)"
            MapKeyCode = vbKeyNumpad1
        Case "2 (TN)"
            MapKeyCode = vbKeyNumpad2
        Case "3 (TN)"
            MapKeyCode = vbKeyNumpad3
        Case "4 (TN)"
            MapKeyCode = vbKeyNumpad4
        Case "5 (TN)"
            MapKeyCode = vbKeyNumpad5
        Case "6 (TN)"
            MapKeyCode = vbKeyNumpad6
        Case "7 (TN)"
            MapKeyCode = vbKeyNumpad7
        Case "8 (TN)"
            MapKeyCode = vbKeyNumpad8
        Case "9 (TN)"
            MapKeyCode = vbKeyNumpad9
        Case "F1"
            MapKeyCode = vbKeyF1
        Case "F2"
            MapKeyCode = vbKeyF2
        Case "F3"
            MapKeyCode = vbKeyF3
        Case "F4"
            MapKeyCode = vbKeyF4
        Case "F5"
            MapKeyCode = vbKeyF5
        Case "F6"
            MapKeyCode = vbKeyF6
        Case "F7"
            MapKeyCode = vbKeyF7
        Case "F8"
            MapKeyCode = vbKeyF8
        Case "F9"
            MapKeyCode = vbKeyF9
        Case "F10"
            MapKeyCode = vbKeyF10
        Case "F11"
            MapKeyCode = vbKeyF11
        Case "F12"
            MapKeyCode = vbKeyF12
        Case "INS", "INSERT"
            MapKeyCode = vbKeyInsert
        Case "CANCELLA", "DELETE", "DEL", "CANC"
            MapKeyCode = vbKeyDelete
        Case "BACKSPACE"
            MapKeyCode = vbKeyCancel
        Case "INVIO"
            MapKeyCode = vbKeyReturn
    End Select
End Function

Public Sub SendKey(ByVal vKey As Integer, Optional booDown As Boolean = False)
    Dim GInput(0) As GENERALINPUT
    Dim KInput As KEYBDINPUT
    KInput.wVk = vKey
    If Not booDown Then
        KInput.dwFlags = KEYEVENTF_KEYUP
    End If
    GInput(0).dwType = INPUT_KEYBOARD
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

'Public Sub SendKeys(ByVal vKey As Integer, Optional booDown As Boolean = False)
'
'    If (iShift And vbCtrlMask) > 0 Then
'        SendKey vbKeyControl, True
'    End If
'    SendKey iKeyCode, True
'    SendKey iKeyCode
'    If (iShift And vbCtrlMask) > 0 Then
'        SendKey vbKeyControl
'    End If
'
'End Sub

Public Sub MySendKeys(ByVal sKey As String)
Dim iShift As Integer
Dim iKeyCode As Integer
Dim iPos As Integer

    If InStr(sKey, "^") > 0 Then
        iShift = vbCtrlMask
    End If
    If InStr(sKey, "+") > 0 Then
        iShift = iShift + vbShiftMask
    End If
    If InStr(sKey, "%") > 0 Then
        iShift = iShift + vbAltMask
    End If
    
    iPos = InStrRev(sKey, "(")
    If iPos > 0 Then
        sKey = Right$(sKey, Len(sKey) - iPos)
    End If
    
    iPos = InStrRev(sKey, ")")
    If iPos > 0 Then
        sKey = Left$(sKey, iPos - 1)
    End If
    
    iPos = InStrRev(sKey, "{")
    If iPos > 0 Then
        sKey = Right$(sKey, Len(sKey) - iPos)
    End If
    
    iPos = InStrRev(sKey, "}")
    If iPos > 0 Then
        sKey = Left$(sKey, iPos - 1)
    End If
    
    iKeyCode = MapKeyCode(sKey)

' Non è necessario anzi comporta lo scatenarsi
' di una ripetizione dello shortcut attivato
' Si è scelto di fare quindi solo il keyup
' xke fare solo il keydown comportava degli
' effetti collaterali...
'
'    '/**+ KEY_DOWN **********************/
'    If (iShift And vbCtrlMask) > 0 Then
'        SendKey vbKeyControl, True
'    End If
'    If (iShift And vbShiftMask) > 0 Then
'        SendKey vbKeyShift, True
'    End If
'    If (iShift And vbAltMask) > 0 Then
'        SendKey vbKeyMenu, True
'    End If
'
'    SendKey iKeyCode, True
'    '/***********************************/
    
    '/**+ KEY_UP **********************/
    SendKey iKeyCode

    If (iShift And vbCtrlMask) > 0 Then
        SendKey vbKeyControl
    End If
    If (iShift And vbShiftMask) > 0 Then
        SendKey vbKeyShift
    End If
    If (iShift And vbAltMask) > 0 Then
        SendKey vbKeyMenu
    End If
    '/***********************************/
End Sub

