Attribute VB_Name = "ShuttingDown"
Sub quitting()
If MsgBox("Are you sure you want to shutdown", vbQuestion + vbYesNo, "ShutDown BUrnTiMeR") = vbYes Then
    Shell_NotifyIcon NIM_DELETE, nid
    msg = msg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Email:      darksystem@blackcode.com"
    msg = msg & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "Website:  www.burnsystem.com"
    sstyle = vbInformation
    ttitle = "BUrn_Net V. 2.0.0"
    MsgBox msg, sstyle, ttitle
    Shell_NotifyIcon NIM_DELETE, nid
    End
Else
Exit Sub
End If
End Sub


