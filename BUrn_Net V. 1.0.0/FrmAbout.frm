VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bit Source Software"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Sys-Info"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Copyright @ 2002-2003 - Mark Joseph Aspiras"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FrmNet.Enabled = True
Unload Me

End Sub

Private Sub Command2_Click()
Sysinfo
End Sub

Private Sub Form_Load()
Label2.Caption = "BUrn_Net is a trademark of BitSource Software. This program is a FreeWare. All Rights Reserved."
Label3.Caption = "Parts of Code and Design by -> BUrnSysTem, SlickzShady, Eyescube, HeadShit, and Kel_muppy"


Label4.Caption = "Microsoft Windows is a trademark of Microsoft Corporation."


End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmNet.Enabled = True
End Sub

