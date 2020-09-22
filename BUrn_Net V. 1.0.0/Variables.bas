Attribute VB_Name = "Variables"



Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const HWND_TOPMOST = -1

Public nid As NOTIFYICONDATA

Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type

Public Const MAX_PREFERRED_LENGTH As Long = -1
Public Const NERR_SUCCESS As Long = 0&
Public Const ERROR_MORE_DATA As Long = 234&

Public Const SV_TYPE_WORKSTATION         As Long = &H1
Public Const SV_TYPE_SERVER              As Long = &H2
Public Const SV_TYPE_SQLSERVER           As Long = &H4
Public Const SV_TYPE_DOMAIN_CTRL         As Long = &H8
Public Const SV_TYPE_DOMAIN_BAKCTRL      As Long = &H10
Public Const SV_TYPE_TIME_SOURCE         As Long = &H20
Public Const SV_TYPE_AFP                 As Long = &H40
Public Const SV_TYPE_NOVELL              As Long = &H80
Public Const SV_TYPE_DOMAIN_MEMBER       As Long = &H100
Public Const SV_TYPE_PRINTQ_SERVER       As Long = &H200
Public Const SV_TYPE_DIALIN_SERVER       As Long = &H400
Public Const SV_TYPE_XENIX_SERVER        As Long = &H800
Public Const SV_TYPE_SERVER_UNIX         As Long = SV_TYPE_XENIX_SERVER
Public Const SV_TYPE_NT                  As Long = &H1000
Public Const SV_TYPE_WFW                 As Long = &H2000
Public Const SV_TYPE_SERVER_MFPN         As Long = &H4000
Public Const SV_TYPE_SERVER_NT           As Long = &H8000
Public Const SV_TYPE_POTENTIAL_BROWSER   As Long = &H10000
Public Const SV_TYPE_BACKUP_BROWSER      As Long = &H20000
Public Const SV_TYPE_MASTER_BROWSER      As Long = &H40000
Public Const SV_TYPE_DOMAIN_MASTER       As Long = &H80000
Public Const SV_TYPE_SERVER_OSF          As Long = &H100000
Public Const SV_TYPE_SERVER_VMS          As Long = &H200000
Public Const SV_TYPE_WINDOWS             As Long = &H400000  'Windows95 and above
Public Const SV_TYPE_DFS                 As Long = &H800000  'Root of a DFS tree
Public Const SV_TYPE_CLUSTER_NT          As Long = &H1000000 'NT Cluster
Public Const SV_TYPE_TERMINALSERVER      As Long = &H2000000 'Terminal Server
Public Const SV_TYPE_DCE                 As Long = &H10000000 'IBM DSS
Public Const SV_TYPE_ALTERNATE_XPORT     As Long = &H20000000 'rtn alternate transport
Public Const SV_TYPE_LOCAL_LIST_ONLY     As Long = &H40000000 'rtn local only
Public Const SV_TYPE_DOMAIN_ENUM         As Long = &H80000000
Public Const SV_TYPE_ALL                 As Long = &HFFFFFFFF

Public Const SV_PLATFORM_ID_OS2       As Long = 400
Public Const SV_PLATFORM_ID_NT        As Long = 500

Public Const MAJOR_VERSION_MASK        As Long = &HF

Public Type SERVER_INFO_100
    sv100_platform_id As Long
    sv100_name As Long
End Type

Public Declare Function NetServerEnum Lib "netapi32" _
        (ByVal servername As Long, _
        ByVal level As Long, _
        buf As Any, _
        ByVal prefmaxlen As Long, _
        entriesread As Long, _
        totalentries As Long, _
        ByVal servertype As Long, _
        ByVal domain As Long, _
        resume_handle As Long) As Long

Public Declare Function NetApiBufferFree Lib "netapi32" _
        (ByVal Buffer As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" _
        (pTo As Any, uFrom As Any, _
        ByVal lSize As Long)

Public Declare Function lstrlenW Lib "kernel32" _
        (ByVal lpString As Long) As Long

Public Const ERROR_ACCESS_DENIED As Long = 5
Public Const ERROR_BAD_NETPATH As Long = 53
Public Const ERROR_INVALID_PARAMETER As Long = 87
Public Const ERROR_NOT_SUPPORTED As Long = 50
Public Const ERROR_INVALID_NAME As Long = 123
Public Const NERR_BASE As Long = 2100

Public Const NERR_NetworkError As Long = (NERR_BASE + 36)
Public Const NERR_NameNotFound As Long = (NERR_BASE + 173)
Public Const NERR_UseNotFound As Long = (NERR_BASE + 150)

Public Const MAX_COMPUTERNAME As Long = 15
Public Const VER_PLATFORM_WIN32s As Long = 0
Public Const VER_PLATFORM_WIN32_WINDOWS As Long = 1
Public Const VER_PLATFORM_WIN32_NT As Long = 2

Public Type OSVERSIONINFO
    OSVSize         As Long
    dwVerMajor      As Long
    dwVerMinor      As Long
    dwBuildNumber   As Long
    PlatformID      As Long
    szCSDVersion    As String * 128
End Type


Public Type NetMessageData
    sServerName As String
    sSendTo As String
    sSendFrom As String
    sMessage As String
End Type


Public Declare Function NetMessageBufferSend Lib "netapi32" _
        (ByVal servername As String, _
        ByVal msgname As String, _
        ByVal fromname As String, _
        ByVal msgbuf As String, _
        ByRef msgbuflen As Long) As Long

Public Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" _
        (ByVal lpBuffer As String, _
        nSize As Long) As Long

Public Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long


Public Function NetSendMessage(msgData As NetMessageData) As String

    Dim success As Long

    
    If IsWinNT = True Then

        With msgData

    
            If .sSendTo = "" Then

                NetSendMessage = GetNetSendMessageStatus(ERROR_INVALID_PARAMETER)
                Exit Function

            Else

    
                If Len(.sMessage) Then

    
                    .sSendTo = StrConv(.sSendTo, vbUnicode)
                    .sMessage = StrConv(.sMessage, vbUnicode)

                    If Len(.sServerName) > 0 Then
                        .sServerName = StrConv(.sServerName, vbUnicode)
                        Else: .sServerName = vbNullString
                    End If

                    If Len(.sSendFrom) > 0 Then
                        .sSendFrom = StrConv(.sSendFrom, vbUnicode)
                        Else: .sSendFrom = vbNullString
                    End If

    
    
                    Screen.MousePointer = vbHourglass

                    success = NetMessageBufferSend(.sServerName, _
                            .sSendTo, _
                            .sSendFrom, _
                            .sMessage, _
                            ByVal Len(.sMessage))

                    Screen.MousePointer = vbNormal

                    NetSendMessage = GetNetSendMessageStatus(success)

                End If
            End If
        End With
    End If

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


Public Function TrimNull(item As String)

    
    Dim pos As Integer

    pos = InStr(item, Chr$(0))

    If pos Then
        TrimNull = Left$(item, pos - 1)
        Else: TrimNull = item
    End If

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
