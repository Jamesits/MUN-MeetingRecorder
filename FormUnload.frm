VERSION 5.00
Begin VB.Form FormUnload 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "FormUnload.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "FormUnload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
FormUnload.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSplash
Unload Clocksetting
Unload frmAbout
If UnloadactiveX Then
    If SystemInf = 64 Then
    Runprog ("regsvr32 /s /u %windir%\SysWOW64\COMDLG32.OCX")
    Runprog ("regsvr32 /s /u %windir%\SysWOW64\mscomctl.ocx")
    Runprog ("del %windir%\SysWOW64\COMDLG32.OCX")
    Runprog ("del %windir%\SysWOW64\mscomctl.ocx")

    Else
    Runprog ("regsvr32 /s /u %windir%\System32\COMDLG32.OCX")
    Runprog ("regsvr32 /s /u %windir%\System32\mscomctl.ocx")
    Runprog ("del %windir%\System32\COMDLG32.OCX")
    Runprog ("del %windir%\System32\mscomctl.ocx")

    End If
    End If
'----------------卸载所有窗体------------------------------
    Dim frm As Form
    For Each frm In Forms
    Unload frm
    Next

End Sub
