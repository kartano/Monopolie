Attribute VB_Name = "mMsgbox"
'---------------------------------------------------------------------------------------
' Module    : mMsgBox
' Date      : Unknown
' Author    : Unknown
' Purpose   : General messagebox utilities
' 08/21/2004    SM:  Fixed bug where Screen.ACtiveForm can be Nothing
' 03/03/2006    SM:  Removed some cruft, tidied up the indentation
'---------------------------------------------------------------------------------------

Option Explicit
'api declares
'api constants
Private Const NV_CLOSEMSGBOX           As Long = &H5000
Private Const NV_MOVEMSGBOX            As Long = &H5001
Private Const NV_ALTERNATECHECK        As Long = &H5003
Private Const NV_ALTERNATEMSGBOX       As Long = &H5002
Private Const SWP_NOSIZE               As Long = &H1
Private Const WS_CHILD                 As Long = &H40000000
Private Const WS_VISIBLE               As Long = &H10000000
Private Const WS_TABSTOP               As Long = &H10000
Private Const WM_SETFONT               As Long = &H30

'types
Public Type RECT
    Left                                 As Long
    Top                                  As Long
    Right                                As Long
    Bottom                               As Long
End Type

'This function allows me to perform a form of subclasing of the Windows MessageBox.
'This is performed by firing a timer that waits for the created Windows Messagbox, then
'modifies it as required
Private Const BS_AUTOCHECKBOX          As Long = &H3
Private Const HWND_TOPMOST             As Integer = -1
Private Const BM_GETSTATE              As Long = &HF2
Private Const WM_GETFONT               As Long = &H31
Private m_sTitle                       As String
Private m_X                            As Long
Private m_Y                            As Long
Private m_lPause                       As Long
Private m_lHandle                      As Long
Private m_Buttons                      As VbMsgBoxStyle
Private m_sPrompt                      As String
Private m_sCheckboxText                As String
Private m_sButtonText(1 To 3)          As String
Private m_lCheckHwnd                   As Long
Private m_bCheckState                  As Boolean
Private m_lHwnd                        As Long
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, _
                                                ByVal nIDEvent As Long, _
                                                ByVal uElapse As Long, _
                                                ByVal lpTimerFunc As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, _
                                                                      ByVal lpText As String, _
                                                                      ByVal lpCaption As String, _
                                                                      ByVal wType As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal nIDEvent As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                      ByVal lpWindowName As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          ByVal lpsz2 As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, _
                                                                            ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
                                                                              ByVal lpClassName As String, _
                                                                              ByVal lpWindowName As String, _
                                                                              ByVal dwStyle As Long, _
                                                                              ByVal X As Long, _
                                                                              ByVal Y As Long, _
                                                                              ByVal nWidth As Long, _
                                                                              ByVal nHeight As Long, _
                                                                              ByVal hWndParent As Long, _
                                                                              ByVal hMenu As Long, _
                                                                              ByVal hInstance As Long, _
                                                                              lpParam As Any) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Public Function AlternateMsgbox(ByVal Prompt As String, _
                                Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                                Optional ByVal Title As String = vbNullString, _
                                Optional ByVal Button1Text As String = vbNullString, _
                                Optional ByVal Button2Text As String = vbNullString, _
                                Optional ByVal Button3Text As String = vbNullString, _
                                Optional ByVal CheckboxText As String = vbNullString, _
                                Optional ByRef CheckboxReturn As Boolean) As VbMsgBoxResult


    SetHwnd
    'This function allows the devloper to change the captions in the buttons and add
    'a checkbox control, the checkbox control is used for "Don't show again" type messages
    'Set Locals
    m_sTitle = Title
    m_sPrompt = Prompt
    m_sButtonText(1) = Button1Text
    m_sButtonText(2) = Button2Text
    m_sButtonText(3) = Button3Text
    m_sCheckboxText = CheckboxText
    m_Buttons = Buttons
    'Fire the checkbox state checking timer
    SetTimer m_lHwnd, NV_ALTERNATECHECK, 0&, AddressOf NewTimerProc
    'Fire the standard mdofy event
    SetTimer m_lHwnd, NV_ALTERNATEMSGBOX, 0&, AddressOf NewTimerProc
    AlternateMsgbox = MessageBox(m_lHwnd, Prompt, Title, Buttons)
    'cancel the checkbox fire timer
    KillTimer m_lHwnd, NV_ALTERNATECHECK
    'return its state
    CheckboxReturn = m_bCheckState
End Function

Public Function NewTimerProc(ByVal hwnd As Long, _
                             ByVal msg As Long, _
                             ByVal wParam As Long, _
                             ByVal lParam As Long) As Long

  
    Dim oForm              As Form
    Dim W                  As Single
    Dim H                  As Single
    Dim rBox               As RECT
    Dim hButton            As Long
    Dim msgButtons(1 To 3) As String
    Dim i                  As Integer
    Dim hFont              As Long
    Dim lCaptionHwnd       As Long
    Dim R                  As RECT
    Dim nHeight            As Integer
    If wParam = NV_ALTERNATECHECK Then
        If m_lCheckHwnd > 0 Then
        'Returns the value of the checkbox on extended MsgBox
            m_bCheckState = (SendMessage(m_lCheckHwnd, BM_GETSTATE, 0, 0&) <> 0)
        End If
    Else 'NOT WPARAM...
        'Cancel timer
        KillTimer hwnd, wParam
        Select Case wParam
        Case NV_CLOSEMSGBOX
            ' A system class is a window class registered by the system which cannot
            ' be destroyed by a processed, e.g. #32768 (a menu), #32769 (desktop
            ' window), #32770 (dialog box), #32771 (task switch window) and
            ' #32770 is a MessageBox
            m_lHandle = FindWindow("#32770", m_sTitle)
            If m_lHandle <> 0 Then
                SetForegroundWindow m_lHandle
                SendKeys "{enter}"
            End If
        Case NV_MOVEMSGBOX
            m_lHandle = FindWindow("#32770", m_sTitle)
            If m_lHandle <> 0 Then
                W = Screen.Width / Screen.TwipsPerPixelX
                H = Screen.Height / Screen.TwipsPerPixelY
                GetWindowRect m_lHandle, rBox
                With rBox
                    If m_X > (W - (.Right - .Left) - 1) Then
                        m_X = (W - (.Right - .Left) - 1)
                    End If
                    If m_Y > (H - (.Bottom - .Top) - 1) Then
                        m_Y = (H - (.Bottom - .Top) - 1)
                    End If
                End With 'RBOX
                If m_X < 1 Then
                    m_X = 1
                    If m_Y < 1 Then
                        m_Y = 1
                    End If
                End If
                ' SWP_NOSIZE is to use current size, ignoring 3rd & 4th parameters.
                SetWindowPos m_lHandle, HWND_TOPMOST, m_X, m_Y, 0, 0, SWP_NOSIZE
            End If
        Case NV_ALTERNATEMSGBOX
            m_lHandle = FindWindow("#32770", m_sTitle)
            If m_lHandle <> 0 Then
                'setup standard msgbox caption array
                If (m_Buttons Or vbRetryCancel) = m_Buttons Then
                    msgButtons(1) = "&Retry"
                    msgButtons(2) = "Cancel"
                ElseIf (m_Buttons Or vbYesNo) = m_Buttons Then 'NOT (M_BUTTONS...
                    msgButtons(1) = "&Yes"
                    msgButtons(2) = "&No"
                ElseIf (m_Buttons Or vbYesNoCancel) = m_Buttons Then 'NOT (M_BUTTONS...
                    msgButtons(1) = "&Yes"
                    msgButtons(2) = "&No"
                    msgButtons(3) = "Cancel"
                ElseIf (m_Buttons Or vbAbortRetryIgnore) = m_Buttons Then 'NOT (M_BUTTONS...
                    msgButtons(1) = "&Abort"
                    msgButtons(2) = "&Retry"
                    msgButtons(3) = "&Ignore"
                ElseIf (m_Buttons Or vbOKCancel) = m_Buttons Then 'NOT (M_BUTTONS...
                    msgButtons(1) = "OK"
                    msgButtons(2) = "Cancel"
                ElseIf (m_Buttons Or vbOKOnly) = m_Buttons Then 'NOT (M_BUTTONS...
                    msgButtons(1) = "OK"
                End If
                'replace the captions where required
                For i = LBound(msgButtons) To UBound(msgButtons)
                    If Len(m_sButtonText(i)) > 0 Then
                        hButton = FindWindowEx(m_lHandle, 0&, "Button", msgButtons(i))
                        If hButton <> 0 Then
                            SetWindowText hButton, m_sButtonText(i)
                        End If
                    End If
                Next i
                'should I add a checkbox
                If Len(m_sCheckboxText) > 0 Then
                    'Find the window
                    lCaptionHwnd = FindWindowEx(m_lHandle, 0, "Static", m_sPrompt)
                    GetWindowRect m_lHandle, R
                    Set oForm = Screen.ActiveForm
                    nHeight = oForm.TextHeight(m_sCheckboxText) / Screen.TwipsPerPixelY
                    'Create the checkbox control
                    m_lCheckHwnd = CreateWindowEx(0, "Button", m_sCheckboxText, WS_CHILD Or WS_VISIBLE Or WS_TABSTOP Or BS_AUTOCHECKBOX, 5, (R.Bottom - R.Top) - nHeight - 20, (oForm.TextWidth(m_sCheckboxText) / Screen.TwipsPerPixelX) + 22, nHeight, m_lHandle, 0, App.hInstance, ByVal 0&)
                    ' set the font of the checkbox to the same as the messagebox
                    hFont = SendMessage(lCaptionHwnd, WM_GETFONT, 0, 0&)
                    SendMessage m_lCheckHwnd, WM_SETFONT, hFont, 0&
                    'move the new checkbox to the correct position
                    MoveWindow m_lHandle, R.Left, R.Top, R.Right - R.Left, R.Bottom - R.Top + nHeight, 1&
                End If
            End If
        End Select
        Set oForm = Nothing
    End If
End Function

Public Function RepositionedMsgbox(ByVal Prompt As String, _
                                   Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                                   Optional ByVal Title As String = vbNullString, _
                                   Optional ByVal inX As Long, _
                                   Optional ByVal inY As Long) As VbMsgBoxResult

    SetHwnd
    'This function allows the devloper to modify the default start-up position of
    'the Windows Msgbox - it defaults to center screen
    m_sTitle = Title
    m_X = inX
    m_Y = inY
    'set the timer to fire
    SetTimer m_lHwnd, NV_MOVEMSGBOX, 0&, AddressOf NewTimerProc
    'invoke Msgbox
    RepositionedMsgbox = MessageBox(m_lHwnd, Prompt, Title, Buttons)
End Function

Public Function SelfClosingMsgbox(ByVal Prompt As String, _
                                  Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, _
                                  Optional ByVal Title As String = vbNullString, _
                                  Optional ByVal nSecs As Integer = 4) As VbMsgBoxResult

    m_sTitle = Title
    'This function will close the MsgBox after x seconds, defaulting to 4
    m_lPause = nSecs * 1000
    SetHwnd
    'set the timer to fire
    SetTimer m_lHwnd, NV_CLOSEMSGBOX, m_lPause, AddressOf NewTimerProc
    'invoke Msgbox
    SelfClosingMsgbox = MessageBox(m_lHwnd, Prompt, Title, Buttons)
    'kill timer so if you show another msgbox it will not autoclose it unless you want it to.
    KillTimer m_lHwnd, NV_CLOSEMSGBOX
End Function

Private Sub SetHwnd()
    m_lHwnd = Screen.ActiveForm.hwnd
End Sub
