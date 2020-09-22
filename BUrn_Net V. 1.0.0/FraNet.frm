VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmNet 
   BackColor       =   &H80000008&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUrn_NeT V. 1.0.0"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   6435
   Icon            =   "FraNet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   6015
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Height          =   615
         Left            =   4080
         TabIndex        =   10
         Top             =   120
         Width           =   1575
         Begin VB.CommandButton Command2 
            Caption         =   "&OK"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "SHutDown System"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000008&
         Caption         =   "SHutDown NetWork"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         MaskColor       =   &H0000FF00&
         TabIndex        =   8
         ToolTipText     =   "THis is for Administrator Control Only"
         Top             =   360
         Width           =   1935
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   4470
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1376
            MinWidth        =   776
            Text            =   "Message:"
            TextSave        =   "Message:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9895
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List1 
      BackColor       =   &H80000007&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   2085
      ItemData        =   "FraNet.frx":030A
      Left            =   240
      List            =   "FraNet.frx":030C
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1080
      Width           =   1965
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1965
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Send Message"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4560
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000008&
      Enabled         =   0   'False
      ForeColor       =   &H0000FF00&
      Height          =   2535
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   240
      Width           =   3645
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   5520
      TabIndex        =   0
      Top             =   5880
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   6015
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent To"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   480
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Sent From"
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblCaption 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Message"
         ForeColor       =   &H0000FF00&
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   13
         Top             =   2880
         Width           =   1335
      End
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      Height          =   195
      Index           =   0
      Left            =   3480
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuShut 
         Caption         =   "&ShutDown"
      End
      Begin VB.Menu Mnulin 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbor 
         Caption         =   "&Abort ShutDown"
         Enabled         =   0   'False
      End
      Begin VB.Menu Mnulin1 
         Caption         =   "-"
      End
      Begin VB.Menu Mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MnuTool 
      Caption         =   "&Tool"
      Begin VB.Menu MnuNet 
         Caption         =   "SHutDown NetWork"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "FrmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAJOR_VERSION_MASK        As Long = &HF

Private Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_name As Long
End Type

Private Declare Function NetServerEnum Lib "netapi32" _
        (ByVal servername As Long, _
        ByVal level As Long, _
        buf As Any, _
        ByVal prefmaxlen As Long, _
        entriesread As Long, _
        totalentries As Long, _
        ByVal servertype As Long, _
        ByVal domain As Long, _
        resume_handle As Long) As Long

Private Declare Function NetApiBufferFree Lib "netapi32" _
        (ByVal Buffer As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" _
        (pTo As Any, uFrom As Any, _
        ByVal lSize As Long)

Private Declare Function lstrlenW Lib "kernel32" _
        (ByVal lpString As Long) As Long

Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_BAD_NETPATH As Long = 53
Private Const ERROR_INVALID_PARAMETER As Long = 87
Private Const ERROR_NOT_SUPPORTED As Long = 50
Private Const ERROR_INVALID_NAME As Long = 123
Private Const NERR_BASE As Long = 2100

Private Const NERR_NetworkError As Long = (NERR_BASE + 36)
Private Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Private Const NERR_UseNotFound As Long = (NERR_BASE + 150)

Private Const MAX_COMPUTERNAME As Long = 15
Private Const VER_PLATFORM_WIN32s As Long = 0
Private Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Private Const VER_PLATFORM_WIN32_NT As Long = 2

Private Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type

Private Type NetMessageData
    sServerName As String
    sSendTo As String
    sSendFrom As String
    sMessage As String
End Type


Private Declare Function NetMessageBufferSend Lib "netapi32" _
        (ByVal servername As String, _
        ByVal msgname As String, _
        ByVal fromname As String, _
        ByVal msgbuf As String, _
        ByRef msgbuflen As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" _
        (ByVal lpBuffer As String, _
        nSize As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long


Private Sub Command2_Click()
Dim RetVal As String
If Option1.Value = True Then
    If MsgBox("See the Remote Shutdown Dialog on the TaskBar", vbInformation + vbYesNo, "Remote Shutdown Control") = vbYes Then
    FrmNet.WindowState = 0
    RetVal = Shell("shutdown -i", 1)
    Else
    FrmNet.WindowState = 0
    End If
ElseIf Option2.Value = True Then
 
    If MsgBox("Warning: This Control will SHutDown the System", vbInformation + vbYesNo, "ShutDown Computer") = vbYes Then
    On Error GoTo err53:
      
         RetVal = Shell("shutdown -s", 1)
            MnuShut.Enabled = False
            MnuAbor.Enabled = True
      
      
    
    End If
    
End If

err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If
'AdjustToken
'ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_REBOOT), 65535

End Sub

Private Sub Form_Load()

    Dim tmp As String

    'pre-load the text boxes with
    'the local computer name for testing
    tmp = Space$(MAX_COMPUTERNAME + 1)
    Call GetComputerName(tmp, Len(tmp))

    Text1.Text = TrimNull(tmp)
    'Text3.Text = TrimNull(tmp)
    Call GetServers(vbNullString)
End Sub



Private Sub Command1_Click()

    Dim msgData As NetMessageData
    Dim sSuccess As String

    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            With msgData
                .sServerName = Text1.Text
                .sSendTo = List1.List(i)
                .sSendFrom = Text3.Text
                .sMessage = Text4.Text
            End With
            sSuccess = NetSendMessage(msgData)
        End If
    Next

    StatusBar1.Panels(2).Text = sSuccess
    
    If Text4.Text = "" Then
    StatusBar1.Panels(2) = "Message was not Successfully Sent"
    End If
End Sub


Private Function IsWinNT() As Boolean

    'returns True if running WinNT/Win2000/WinXP
    #If Win32 Then

        Dim OSV As OSVERSIONINFO

        OSV.OSVSize = Len(OSV)

        If GetVersionEx(OSV) = 1 Then

            'PlatformId contains a value representing the OS.
            IsWinNT = (OSV.PlatformID = VER_PLATFORM_WIN32_NT)

        End If

    #End If

End Function


Private Function NetSendMessage(msgData As NetMessageData) As String

    Dim success As Long

    'assure that the OS is NT ..
    'NetMessageBufferSend  can not
    'be called on Win9x
    If IsWinNT() Then

        With msgData

            'if To name omitted return error and exit
            If .sSendTo = "" Then

                NetSendMessage = GetNetSendMessageStatus(ERROR_INVALID_PARAMETER)
                Exit Function

            Else

                'if there is a message
                If Len(.sMessage) Then

                    'convert the strings to unicode
                    .sSendTo = StrConv(.sSendTo, vbUnicode)
                    .sMessage = StrConv(.sMessage, vbUnicode)

                    'Note that the API could be called passing
                    'vbNullString as the SendFrom and sServerName
                    'strings. This would generate the message on
                    'the sending machine.
                    If Len(.sServerName) > 0 Then
                        .sServerName = StrConv(.sServerName, vbUnicode)
                        Else: .sServerName = vbNullString
                    End If

                    If Len(.sSendFrom) > 0 Then
                        .sSendFrom = StrConv(.sSendFrom, vbUnicode)
                        Else: .sSendFrom = vbNullString
                    End If

                    'change the cursor and show. Control won't return
                    'until the call has completed.
                    Screen.MousePointer = vbHourglass

                    success = NetMessageBufferSend(.sServerName, _
                            .sSendTo, _
                            .sSendFrom, _
                            .sMessage, _
                            ByVal Len(.sMessage))

                    Screen.MousePointer = vbNormal

                    NetSendMessage = GetNetSendMessageStatus(success)

                End If 'If Len(.sMessage)
            End If  'If .sSendTo
        End With  'With msgData
    End If  'If IsWinNT

End Function


Private Function GetNetSendMessageStatus(nError As Long) As String

    Dim msg As String

    Select Case nError

        Case NERR_SUCCESS:            msg = "The message was successfully sent"
        Case NERR_NameNotFound:       msg = "Send To not found"
        Case NERR_NetworkError:       msg = "General network error occurred"
        Case NERR_UseNotFound:        msg = "Network connection not found"
        Case ERROR_ACCESS_DENIED:     msg = "Access to computer denied"
        Case ERROR_BAD_NETPATH:       msg = "Sent From server name not found."
        Case ERROR_INVALID_PARAMETER: msg = "Invalid parameter(s) specified."
        Case ERROR_NOT_SUPPORTED:     msg = "Network request not supported."
        Case ERROR_INVALID_NAME:      msg = "Illegal character or malformed name."
        Case Else:                    msg = "Unknown error executing command."

    End Select

    GetNetSendMessageStatus = msg

End Function


Private Function TrimNull(item As String)

    'return string before the terminating null
    Dim pos As Integer

    pos = InStr(item, Chr$(0))

    If pos Then
        TrimNull = Left$(item, pos - 1)
        Else: TrimNull = item
    End If

End Function



Private Function GetServers(sDomain As String) As Long

    'lists all servers of the specified type
    'that are visible in a domain.

    Dim bufptr          As Long
    Dim dwEntriesread   As Long
    Dim dwTotalentries  As Long
    Dim dwResumehandle  As Long
    Dim se100           As SERVER_INFO_100
    Dim success         As Long
    Dim nStructSize     As Long
    Dim cnt             As Long

    nStructSize = LenB(se100)

    'Call passing MAX_PREFERRED_LENGTH to have the
    'API allocate required memory for the return values.
    '
    'The call is enumerating all machines on the
    'network (SV_TYPE_ALL); however, by Or'ing
    'specific bit masks for defined types you can
    'customize the returned data. For example, a
    'value of 0x00000003 combines the bit masks for
    'SV_TYPE_WORKSTATION (0x00000001) and
    'SV_TYPE_SERVER (0x00000002).
    '
    'dwServerName must be Null. The level parameter
    '(100 here) specifies the data structure being
    'used (in this case a SERVER_INFO_100 structure).
    '
    'The domain member is passed as Null, indicating
    'machines on the primary domain are to be retrieved.
    'If you decide to use this member, pass
    'StrPtr("domain name"), not the string itself.
    success = NetServerEnum(0&, _
            100, _
            bufptr, _
            MAX_PREFERRED_LENGTH, _
            dwEntriesread, _
            dwTotalentries, _
            SV_TYPE_ALL, _
            0&, _
            dwResumehandle)

    'if all goes well
    If success = NERR_SUCCESS And _
            success <> ERROR_MORE_DATA Then

        'loop through the returned data, adding each
        'machine to the list
        For cnt = 0 To dwEntriesread - 1

            'get one chunk of data and cast
            'into an SERVER_INFO_100 struct
            'in order to add the name to a list
            CopyMemory se100, ByVal bufptr + (nStructSize * cnt), nStructSize

            List1.AddItem GetPointerToByteStringW(se100.sv100_name)

        Next

    End If

    'clean up regardless of success
    Call NetApiBufferFree(bufptr)

    'return entries as sign of success
    GetServers = dwEntriesread

End Function


Public Function GetPointerToByteStringW(ByVal dwData As Long) As String

    Dim tmp() As Byte
    Dim tmplen As Long

    If dwData <> 0 Then

        tmplen = lstrlenW(dwData) * 2

        If tmplen <> 0 Then

            ReDim tmp(0 To (tmplen - 1)) As Byte
            CopyMemory tmp(0), ByVal dwData, tmplen
            GetPointerToByteStringW = tmp

        End If

    End If

End Function



Private Sub Form_Unload(Cancel As Integer)
FrmSys.Show
'Call quit
End Sub

Private Sub MnuAbor_Click()
Dim RetVal As String
On Error GoTo err53:
RetVal = Shell("shutdown -a", 1)

MnuAbor.Enabled = False
MnuShut.Enabled = True

err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If
End Sub

Private Sub MnuAbout_Click()
FrmNet.Enabled = False
FrmAbout.Show
End Sub

Private Sub Mnuexit_Click()
Unload Me
FrmSys.Show
End Sub

Private Sub MnuNet_Click()
If MsgBox("See the Remote Shutdown Dialog on the TaskBar", vbInformation + vbYesNo, "Remote Shutdown Control") = vbYes Then
    FrmNet.WindowState = 0
    RetVal = Shell("shutdown -i", 1)
    Else
    FrmNet.WindowState = 0
    End If
End Sub

Private Sub MnuShut_Click()
Dim RetVal As String
If MsgBox("Warning: This Control will SHutDown the System", vbInformation + vbYesNo, "ShutDown Computer") = vbYes Then
    On Error GoTo err53:
    RetVal = Shell("shutdown -s", 1)
    MnuAbor.Enabled = True
    MnuShut.Enabled = False

End If
    
err53:
If Err = 53 Then
MsgBox "This is for XP machine only", vbInformation, "Error"
End If
'AdjustToken
'ExitWindowsEx (EWX_SHUTDOWN Or EWX_FORCE Or EWX_REBOOT), 65535
    

End Sub

Private Sub Text3_Change()
If Len(Text3) >= 1 Then
List1.Enabled = True
Text4.Enabled = True
Command1.Enabled = True
ElseIf Len(Text3) < 1 Then
List1.Enabled = False
Text4.Enabled = False
Command1.Enabled = False

End If
End Sub
