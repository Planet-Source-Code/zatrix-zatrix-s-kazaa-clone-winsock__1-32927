VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ZATRiX's KaZaA Clone"
   ClientHeight    =   4905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   Picture         =   "frmMain.frx":014A
   ScaleHeight     =   4905
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox txtStatus 
      Height          =   825
      Left            =   525
      TabIndex        =   18
      Top             =   1755
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1455
      _Version        =   393217
      BackColor       =   9275527
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":24717
   End
   Begin VB.Timer tmMain 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6315
      Top             =   2010
   End
   Begin VB.CommandButton cmbScan 
      Appearance      =   0  'Flat
      Caption         =   "&Scan"
      Height          =   1170
      Left            =   2325
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Scan For Users With KaZaA"
      Top             =   570
      Width           =   825
   End
   Begin VB.ListBox lstMain 
      Appearance      =   0  'Flat
      BackColor       =   &H008D8887&
      Height          =   1155
      Left            =   3270
      Style           =   1  'Checkbox
      TabIndex        =   9
      ToolTipText     =   "Found IP addresses"
      Top             =   570
      Width           =   3180
   End
   Begin VB.TextBox IpGroupE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1050
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   5
      ToolTipText     =   "Ending IP address"
      Top             =   1320
      Width           =   405
   End
   Begin VB.TextBox IpGroupE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1455
      MaxLength       =   3
      TabIndex        =   6
      ToolTipText     =   "Ending IP address"
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox IpGroupE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1830
      MaxLength       =   3
      TabIndex        =   7
      Text            =   "255"
      ToolTipText     =   "Ending IP address"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Timer tmCount 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2625
      Top             =   630
   End
   Begin VB.PictureBox Maskk 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   -15
      Picture         =   "frmMain.frx":24799
      ScaleHeight     =   570
      ScaleWidth      =   375
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox IpGroup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   645
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "24"
      ToolTipText     =   "Starting IP address"
      Top             =   765
      Width           =   405
   End
   Begin VB.TextBox IpGroup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1815
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Starting IP address"
      Top             =   765
      Width           =   390
   End
   Begin VB.TextBox IpGroup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "100"
      ToolTipText     =   "Starting IP address"
      Top             =   765
      Width           =   375
   End
   Begin VB.TextBox IpGroup 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1050
      MaxLength       =   3
      TabIndex        =   1
      Text            =   "66"
      ToolTipText     =   "Starting IP address"
      Top             =   765
      Width           =   390
   End
   Begin VB.TextBox IpGroupE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   645
      Locked          =   -1  'True
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "555"
      ToolTipText     =   "Ending IP address"
      Top             =   1320
      Width           =   405
   End
   Begin InetCtlsObjects.Inet Ineter 
      Left            =   3225
      Top             =   2025
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RemotePort      =   1214
      URL             =   "http://"
   End
   Begin SHDocVwCtl.WebBrowser www 
      Height          =   1815
      Left            =   510
      TabIndex        =   10
      Top             =   2580
      Width           =   5940
      ExtentX         =   10477
      ExtentY         =   3201
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSWinsockLib.Winsock Connector 
      Index           =   0
      Left            =   6360
      Top             =   1065
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   1214
   End
   Begin VB.Label lblSF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End At:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   1
      Left            =   1050
      TabIndex        =   17
      Top             =   1110
      Width           =   840
   End
   Begin VB.Label lblSF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start From:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   795
      TabIndex        =   16
      Top             =   555
      Width           =   1320
   End
   Begin VB.Label lblMini 
      BackStyle       =   0  'Transparent
      Height          =   165
      Left            =   6105
      TabIndex        =   15
      ToolTipText     =   "Minimize"
      Top             =   270
      Width           =   240
   End
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Height          =   285
      Left            =   6405
      TabIndex        =   14
      ToolTipText     =   "Exit"
      Top             =   90
      Width           =   255
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Thank You For Downloading My Prog!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1470
      TabIndex        =   13
      ToolTipText     =   "Status Label"
      Top             =   4425
      Width           =   4080
   End
   Begin VB.Label MoveForm 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   780
      TabIndex        =   12
      ToolTipText     =   "Click and Drag to move!"
      Top             =   90
      Width           =   5970
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long
Dim names As String, counter As Integer
Sub Scan()
Dim temp As Integer
tmCount = True
frmMain.Enabled = False
frmAbort.Show , frmMain
frmAbort.Left = frmMain.Width + frmMain.Left
frmAbort.Top = frmMain.Top
If Val(IpGroupE(2).Text) < Val(IpGroup(2).Text) Then
    temp = Val(IpGroup(2).Text)
    IpGroup(2).Text = Val(IpGroupE(2).Text)
    IpGroupE(2).Text = temp
ElseIf Val(IpGroupE(3).Text) < Val(IpGroup(3).Text) Then
    temp = Val(IpGroup(3).Text)
    IpGroup(3).Text = Val(IpGroupE(3).Text)
    IpGroupE(3).Text = temp
End If
lblStatus.Caption = "Scanning................"
End Sub

Private Sub cmbScan_Click()
Call Scan
End Sub

Private Sub cmbScan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu frmAbort.mnuScanM
End Sub

Private Sub Connector_Connect(Index As Integer)
Connector(Index).SendData "PASS Admin" & vbCrLf & "ZATRiX" & vbCrLf & "USER KaZaAClone " & Connector(Index).LocalIP & ":KaZaA"
End Sub

Private Sub Connector_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String, KaZaAUser, added As Boolean, rIP As String, add As String, startt As Integer, endd As Integer
'On Error GoTo X
On Error Resume Next
rIP = Connector(Index).RemoteHostIP
Connector(Index).GetData data, vbString
txtStatus.Text = data & txtStatus.Text
If InStr(data, "???") <> 0 Then
    add = "?"
    peeps_v = peeps_v + 1
Else
    add = "!"
End If
startt = InStr(1, data, "X-Kazaa-Username: ")
startt = startt + Len("X-Kazaa-Username: ")
endd = InStr(startt, data, Chr(10))
data = Mid(data, startt, endd - startt - 1)
If U_Found(rIP) = False Then
    lstMain.AddItem add & data & add & "  " & rIP & " 1214"
    peeps_f = peeps_f + 1
    lblStatus = "Found So Far: " & peeps_f & "..." & "?Searchable Users?: " & peeps_v
End If
add = ""
End Sub
Public Function U_Found(ip As String) As Boolean
Static i As Integer
For i = 0 To lstMain.ListCount - 1
    If InStr(lstMain.List(i), ip) <> 0 Then
        i = lstMain.ListCount - 1
        U_Found = True
    Else
        U_Found = False
    End If
Next
End Function

Function RegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
  Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
  Dim lngRgnFinal As Long, lngRgnTmp As Long
  Dim lngStart As Long, lngRow As Long
  Dim lngCol As Long
  If lngTransColor& < 1 Then
    lngTransColor& = GetPixel(picSource.hDC, 0, 0)
  End If
  lngHeight& = picSource.Height / Screen.TwipsPerPixelY
  lngWidth& = picSource.Width / Screen.TwipsPerPixelX
  lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
  For lngRow& = 0 To lngHeight& - 1
    lngCol& = 0
    Do While lngCol& < lngWidth&
      Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
        lngCol& = lngCol& + 1
      Loop
      If lngCol& < lngWidth& Then
        lngStart& = lngCol&
        Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
          lngCol& = lngCol& + 1
        Loop
        If lngCol& > lngWidth& Then lngCol& = lngWidth&
        lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
        lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
        DeleteObject (lngRgnTmp&)
      End If
    Loop
  Next
  RegionFromBitmap& = lngRgnFinal&
  
End Function
Sub ChangeMask()
'On Error Resume Next ' In case of error
' This is also part of Dos's Dos-Shape example. To update if the skin is changed
  Dim lngRetr As Long
  lngRegion& = RegionFromBitmap(Maskk, vbWhite)
  lngRetr& = SetWindowRgn(Me.hwnd, lngRegion&, True)
End Sub
Private Sub Form_Load()
Dim rep As VbMsgBoxResult, sett As Boolean
If GetSetting("ZATRiX KaZaA Clone", "Settings", "!WARNING!", "TRUE") = False Then GoTo cont
rep = MsgBox("!WARNING!:" & vbCrLf & "My program requires some Registry values to be inserted into your registry!" & vbCrLf & "If you do NOT wish for these values to be used press [NO]! Otherwise press [YES]!", vbCritical + vbYesNo, "WARNING")
If rep = vbNo Then
    End
Else
    SaveSetting "ZATRiX KaZaA Clone", "Settings", "!WARNING!", "FALSE"
End If
cont:
IpGroupE(0).Text = IpGroup(0).Text
IpGroupE(1).Text = IpGroup(1).Text
IpGroupE(2).Text = IpGroup(2).Text
www.Navigate "http://www.kazaa.com"
Settings_Reg(3) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Abort", "TRUE")
Settings_Reg(2) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "TRUE")
Settings_Reg(1) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Search", "TRUE")
Maskk.AutoSize = True
MoveForm.ZOrder 1
ChangeMask
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
VBA.SetAttr App.path & "\tempPage.html", vbNormal
VBA.SetAttr App.path & "\errorTemp.html", vbNormal
Kill App.path & "\tempPage.html"
Kill App.path & "\errorTemp.html"
End Sub

Private Sub lblExit_Click()
Unload Me
End
End Sub

Private Sub lblMini_Click()
Me.WindowState = 1
End Sub

Private Sub MoveForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then ' Left button
    ReleaseCapture
    Call SendMessage(Me.hwnd, &HA1, 2, 0)
End If
End Sub
Private Sub IpGroup_Change(Index As Integer)
If Index <> 2 And Index <> 3 Then IpGroupE(Index).Text = IpGroup(Index).Text
If Val(IpGroup(Index).Text) > 255 Then
    MsgBox "The max is 255!", vbCritical, Me.Caption
    IpGroup(Index).Text = 255
End If
End Sub

Private Sub IpGroupE_Change(Index As Integer)
If Val(IpGroupE(Index).Text) > 255 Then
    MsgBox "The max is 255!", vbCritical, Me.Caption
    IpGroupE(Index).Text = 255
End If
End Sub

Private Sub lstMain_ItemCheck(Item As Integer)
Dim i As Integer
For i = 0 To lstMain.ListCount - 1
    lstMain.Selected(i) = False
    lstMain.Selected(Item) = True
Next
item_chk = Item
frmAbort.mnuView.Caption = "&View Files for [" & lstMain.List(Item) & "]"
End Sub

Private Sub lstMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu frmAbort.mnuMain
End Sub

Private Sub tmCount_Timer()
Static rep As VbMsgBoxResult
If Val(IpGroup(3)) < Val(IpGroupE(3)) Then
            IpGroup(3).Text = Val(IpGroup(3)) + 1
        Else
            IpGroup(3) = 0
            If Val(IpGroup(2)) < Val(IpGroupE(2)) Then
                IpGroup(2) = Val(IpGroup(2)) + 1
            Else
                If Settings_Reg(2) Then
                    rep = MsgBoxEx("Scan Complete", vbOKCancel + vbInformation, Me.Caption, , , , Me.hwnd, "&OK", "Don't Show")
                    If rep = vbCancel Then SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "FALSE"
                    Settings_Reg(2) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "TRUE")
                End If
                tmCount = False
                frmMain.Enabled = True
                Unload frmAbort
                lblStatus.Caption = "!Found Users!: " & lstMain.ListCount & "..." & "?Searchable Users?: " & peeps_v
                peeps_f = 0
                peeps_v = 0
            End If
        End If
tmMain_Timer
End Sub

Private Sub tmMain_Timer()
Dim hostt As String
Load Connector(Connector.UBound + 1)
        hostt = IpGroup(0) & "." & IpGroup(1) & "." & IpGroup(2) & "." & IpGroup(3)
        Connector(Connector.UBound).Close
        Connector(Connector.UBound).Connect hostt, 1214
        If Connector.UBound > 100 Then Unload Connector(Connector.UBound - 100)
End Sub

Private Sub txtStatus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu frmAbort.mnuStat
End Sub

Private Sub www_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
End Sub

