VERSION 5.00
Begin VB.Form frmAbort 
   BorderStyle     =   0  'None
   ClientHeight    =   945
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   1785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   1785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmbAbort 
      Caption         =   "&Abort Search"
      Height          =   855
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   1695
   End
   Begin VB.Label lblMisc 
      Caption         =   "<a href="""
      Height          =   90
      Left            =   285
      TabIndex        =   1
      Top             =   225
      Width           =   75
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Main"
      Visible         =   0   'False
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear List"
      End
      Begin VB.Menu mnuRef 
         Caption         =   "&Refresh List"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Files of Selected"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "&Search For ..."
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuStat 
      Caption         =   "Status"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "&Save To Disk"
      End
      Begin VB.Menu mnuTCl 
         Caption         =   "&Clear"
      End
   End
   Begin VB.Menu mnuScanM 
      Caption         =   "ScanTypes"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuScanIP 
         Caption         =   "&Scan For A Specific IP"
      End
      Begin VB.Menu mnuBuff 
         Caption         =   "&Reset Settings"
      End
   End
End
Attribute VB_Name = "frmAbort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbAbort_Click()
Dim rep As VbMsgBoxResult
frmMain.Enabled = True
frmMain.tmCount = False
If Settings_Reg(3) Then
    rep = MsgBoxEx("Scan Aborted!", vbCritical + vbOKCancel, Me.Caption, , , , frmMain.hwnd, "&OK", "Don't Show")
    If rep = vbCancel Then SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Abort", "FALSE"
    Settings_Reg(3) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Abort", "TRUE")
End If
Unload Me
End Sub

Private Sub Form_Load()
Me.Width = cmbAbort.Width + 110
Me.Height = cmbAbort.Height + 110
End Sub

Private Sub mnuBuff_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Abort", "TRUE"
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "TRUE"
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Search", "TRUE"
SaveSetting "ZATRiX KaZaA Clone", "Settings", "!WARNING!", "TRUE"
Settings_Reg(3) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Abort", "TRUE")
Settings_Reg(2) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "TRUE")
Settings_Reg(1) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Search", "TRUE")
End Sub

Private Sub mnuScanIP_Click()
frmMain.Connector(0).Close
frmMain.Connector(0).RemoteHost = InputBox("What specific IP do you wish to check?")
frmMain.Connector(0).RemotePort = 1214
frmMain.Connector(0).Connect
End Sub
Private Sub mnuAbout_Click()
MsgBox "Thank you for downloading my code!" & vbCrLf & " by: ZATRiX", vbInformation, ".::Proud To Be Canadian::."
End Sub

Private Sub mnuClear_Click()
frmMain.txtStatus.Text = ""
End Sub

Private Sub mnuRef_Click()
frmMain.lstMain.Clear
Call frmMain.Scan
End Sub

Private Sub mnuSAF_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Abort", "FALSE"
Unload Me
Me.Show
End Sub

Private Sub mnuSAT_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Abort", "TRUE"
Unload Me
Me.Show
End Sub

Private Sub mnuSave_Click()
Dim inf As Integer
Dim path As String
inf = FreeFile
path = InputBox("Enter A location for the file to be stored", Me.Caption, App.path & "\logstatus.txt")
Open path For Output As #inf
    Print #inf, frmMain.txtStatus.Text
Close
MsgBox "The file was saved!", vbExclamation, path
End Sub

Private Sub mnuSearch_Click()
Dim temp, i As Integer, data As String, topic As String, iplist As String, rep As VbMsgBoxResult
topic = InputBox("What would you like to search for?", Me.Caption, "case SENSITIVE")
For i = 0 To frmMain.lstMain.ListCount - 1
    temp = Split(frmMain.lstMain.List(i), " ")
    data = frmMain.Ineter.OpenURL("http://" & temp(2) & ":1214")
    If InStr(data, topic) <> 0 Then iplist = temp(0) & vbCrLf & iplist
    frmMain.lblStatus.Caption = "Searching in......." & frmMain.lstMain.List(i)
Next
If Settings_Reg(1) Then
    If iplist <> "" Then
        MsgBox "The following users have a possible match for what you requested:" & vbCrLf & iplist, vbExclamation, Me.Caption
        If rep = vbCancel Then SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Search", "FALSE"
    Else
        rep = MsgBox("Sorry nothing was found.", vbCritical + vbRetryCancel, Me.Caption)
        If rep = vbRetry Then mnuSearch_Click
    End If
    Settings_Reg(1) = GetSetting("ZATRiX KaZaA Clone", "Settings", "Show Search", "TRUE")
End If
frmMain.lblStatus.Caption = "Done Search"
End Sub

Private Sub mnuSFF_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "FALSE"
Unload Me
Me.Show
End Sub

Private Sub mnuSFT_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Scan Finished", "TRUE"
Unload Me
Me.Show
End Sub

Private Sub mnuSSF_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Search", "FALSE"
Unload Me
Me.Show
End Sub

Private Sub mnuSST_Click()
SaveSetting "ZATRiX KaZaA Clone", "Settings", "Show Search", "TRUE"
Unload Me
Me.Show
End Sub

Private Sub mnuView_Click()
On Error Resume Next
If frmMain.lstMain.List(item_chk) = "" Then Exit Sub
Dim temp, data As String, inf As Integer
temp = Split(frmMain.lstMain.List(item_chk), " ")
data = frmMain.Ineter.OpenURL("Http://" & temp(2) & ":1214")
If data = "" Then
    Call Errorr(frmMain.lstMain.List(item_chk))
    Exit Sub
End If
data = Replace(data, "<body>", "<body bgcolor='87888D' text='white' link='yellow' vlink='black'>")
data = Replace(data, lblMisc.Caption & "/", lblMisc.Caption & "http://" & temp(2) & ":1214/")
data = Replace(data, "><table>", "><img src='" & App.path & "\logo.gif" & "'><br><table>")
inf = FreeFile
VBA.SetAttr App.path & "\tempPage.html", vbNormal
Open App.path & "\tempPage.html" For Output As #inf
    Print #inf, data
Close #inf
VBA.SetAttr App.path & "\tempPage.html", vbHidden
frmMain.www.Navigate App.path & "\tempPage.html"
End Sub
Sub Errorr(namer As String)
On Error Resume Next
Dim data As String, inf As Integer
inf = FreeFile
VBA.SetAttr App.path & "\errorTemp.html", vbNormal
Open App.path & "\error.html" For Binary As #inf
    data = Space(LOF(inf))
    Get #inf, , data
Close #inf
namer = Replace(namer, "!", "", 1, Len(namer))
data = Replace(data, "name", namer)
data = Replace(data, "mm/dd/yy", VBA.Date)
Open App.path & "\errorTemp.html" For Output As #inf
    Print #inf, data
Close #inf
VBA.SetAttr App.path & "\errorTemp.html", vbHidden
frmMain.www.Navigate App.path & "\errorTemp.html"
End Sub
