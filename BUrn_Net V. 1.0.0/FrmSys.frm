VERSION 5.00
Begin VB.Form FrmSys 
   BorderStyle     =   0  'None
   Caption         =   "BUrn_Net V. 1.0.0"
   ClientHeight    =   90
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   90
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu Mnulin 
      Caption         =   "&-"
      Visible         =   0   'False
      Begin VB.Menu MnuBUrn 
         Caption         =   "&ShutDown BUrn_Net"
      End
      Begin VB.Menu Mnulin2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuNet 
         Caption         =   "&ShutDown NetWork"
      End
      Begin VB.Menu Mnulin3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSys 
         Caption         =   "&ShutDown System"
      End
      Begin VB.Menu Mnulin4 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbort 
         Caption         =   "&Abort ShutDown"
      End
      Begin VB.Menu Mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAb 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmSys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MnuAbort.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
Sys = X / Screen.TwipsPerPixelX
Select Case Sys
Case WM_LBUTTONDOWN:
'Shell_NotifyIcon NIM_DELETE, nid
FrmNet.Show
Case WM_RBUTTONUP:
PopupMenu Mnulin
End Select
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = FrmNet.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub

Private Sub MnuAb_Click()
FrmNet.Enabled = False
FrmAbout.Show
End Sub

Private Sub MnuAbort_Click()
Dim RetVal As String
On Error GoTo err53:

RetVal = Shell("shutdown -a", 1)
MnuSys.Enabled = True
MnuAbort.Enabled = False
err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If
End Sub

Private Sub MnuBUrn_Click()

Call quitting


End Sub

Private Sub MnuNet_Click()
Dim RetVal As String
On Error GoTo err53:
RetVal = Shell("shutdown -i", 1)

err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If


End Sub

Private Sub MnuSys_Click()
On Error GoTo err53:
RetVal = Shell("shutdown -s", 1)
MnuSys.Enabled = False
MnuAbort.Enabled = True

err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If
End Sub
