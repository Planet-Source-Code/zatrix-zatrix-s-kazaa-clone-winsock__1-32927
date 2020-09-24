Attribute VB_Name = "modMSGBox"
Option Explicit
'global
Global Settings_Reg(1 To 3) As Boolean
Global item_chk As Integer, peeps_f As Integer, peeps_v As Integer

';;
Global MsgBox_Title As String

Declare Function MessageBoxEx Lib "user32" Alias "MessageBoxExA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal uType As Long, ByVal wLanguageId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetDlgItemText Lib "user32" Alias "SetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String) As Long
Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Const IDOK = 1
Const IDCANCEL = 2
Const IDABORT = 3
Const IDRETRY = 4
Const IDIGNORE = 5
Const IDYES = 6
Const IDNO = 7
Const IDHELP = 7

Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Dim hHook As Long
Dim parenthWnd As Long
Dim centerScreen As Boolean
Dim buttonText1 As String
Dim buttonText2 As String
Dim buttonText3 As String
Dim buttonType1 As Integer
Dim buttonType2 As Integer
Dim buttonType3 As Integer
Dim buttonTextHelp As String
Dim buttonTypeHelp As Integer

Public Function MsgBoxEx(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String = "", Optional ByVal HelpFile As String, Optional ByVal Context, Optional ByVal centerForm As Boolean = True, Optional ByVal hwnd As Long, Optional button1Text As String, Optional button2Text As String, Optional button3Text As String, Optional helpButtonText As String) As VbMsgBoxResult
  Dim ret As Long
  Dim hInst As Long
  Dim Thread As Long
  'Set up the CBT hook
  parenthWnd = hwnd
  buttonText1 = button1Text
  buttonText2 = button2Text
  buttonText3 = button3Text
  buttonTextHelp = helpButtonText

  MsgBox_Title = Title

  If (Buttons And vbRetryCancel) = vbRetryCancel Then
    buttonType1 = IDRETRY
    buttonType2 = IDCANCEL
  ElseIf (Buttons And vbYesNo) = vbYesNo Then
    buttonType1 = IDYES
    buttonType2 = IDNO
  ElseIf (Buttons And vbYesNoCancel) = vbYesNoCancel Then
    buttonType1 = IDYES
    buttonType2 = IDNO
    buttonType3 = IDCANCEL
  ElseIf (Buttons And vbAbortRetryIgnore) = vbAbortRetryIgnore Then
    buttonType1 = IDABORT
    buttonType2 = IDRETRY
    buttonType3 = IDIGNORE
  ElseIf (Buttons And vbOKCancel) = vbOKCancel Then
    buttonType1 = IDOK
    buttonType2 = IDCANCEL
  ElseIf (Buttons And vbOKOnly) = vbOKOnly Then
    buttonType1 = IDOK
  End If
  If (Buttons And vbMsgBoxHelpButton) = vbMsgBoxHelpButton Then
    buttonTypeHelp = 1
  Else
    buttonTypeHelp = 0
  End If

  hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  centerScreen = Not centerForm
  hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenter, hInst, Thread)

  ret = MessageBoxEx(hwnd, Prompt, Title, Buttons, 0)
  MsgBoxEx = ret
End Function
Private Function WinProcCenter(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rectForm As RECT, rectMsg As RECT
  Dim X As Long, Y As Long
  If lMsg = HCBT_ACTIVATE Then
    If centerScreen = True Then
      'Show the MsgBox at a fixed location (0,0)
      GetWindowRect wParam, rectMsg
      X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
      Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
    Else
      'Get the coordinates of the form and the message box so that
      'you can determine where the center of the form is located
      GetWindowRect parenthWnd, rectForm
      GetWindowRect wParam, rectMsg
      X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
      Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
    End If
    SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    SetButtonText wParam
    'Release the CBT hook
    UnhookWindowsHookEx hHook
  End If
  WinProcCenter = False
End Function

Private Sub SetButtonText(ByVal wParam As Long)
  If buttonText1 <> "" Then
    SetDlgItemText wParam, buttonType1, buttonText1
  End If
  If buttonText2 <> "" Then
    SetDlgItemText wParam, buttonType2, buttonText2
  End If
  If buttonText3 <> "" Then
    SetDlgItemText wParam, buttonType3, buttonText3
  End If
  If buttonTextHelp <> "" And buttonTypeHelp = 1 Then
    SetHelpButtonText wParam
  End If
End Sub
Private Sub SetHelpButtonText(ByVal wParam As Long)
  Dim Btn(0 To 3) As Long
  Dim T As Integer
  Dim cName As String
  Dim Length As Long

  Btn(0) = FindWindowEx(wParam, 0, vbNullString, vbNullString)

  For T = 1 To 3
    Btn(T) = FindWindowEx(wParam, Btn(T - 1), vbNullString, vbNullString)
    If Btn(T) = 0 Then
      Exit For
    End If
  Next T

  For T = 3 To 0 Step -1
    If Btn(T) <> 0 And Btn(T) <> wParam Then
      cName = Space(255)
      cName = Left(cName, GetClassName(Btn(T), cName, 255))
      If UCase(cName) = "BUTTON" Then
        SetWindowText Btn(T), buttonTextHelp
        Exit For
      End If
    End If
  Next T

End Sub

