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
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
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
If Common.UnloadactiveX Then
    If SystemInf = 64 Then
    Common.Runprog ("regsvr32 /s /u %windir%\SysWOW64\COMDLG32.OCX")
    Common.Runprog ("regsvr32 /s /u %windir%\SysWOW64\mscomctl.ocx")
    Common.Runprog ("del %windir%\SysWOW64\COMDLG32.OCX")
    Common.Runprog ("del %windir%\SysWOW64\mscomctl.ocx")

    Else
    Common.Runprog ("regsvr32 /s /u %windir%\System32\COMDLG32.OCX")
    Common.Runprog ("regsvr32 /s /u %windir%\System32\mscomctl.ocx")
    Common.Runprog ("del %windir%\System32\COMDLG32.OCX")
    Common.Runprog ("del %windir%\System32\mscomctl.ocx")

    End If
    End If
    
    Dim frm As Form
    For Each frm In Forms
    Unload frm
    Next

End Sub
