VERSION 5.00
Begin VB.Form Common 
   Caption         =   "���ó��������洢���ô���"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Common.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '����ȱʡ
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public iniFileName As String
Public glasseffect As Byte
Public Debugmode, Loadactivex, UnloadactiveX, Beta, InitMeet As Boolean
Public registername, registercompany, isreged As String
 
Private Sub Form_Load()
Common.Hide
'---------------------------------���ñ���������---------------------------
Debugmode = False
Loadactivex = False
UnloadactiveX = False
Beta = False
InitMeet = False
registername = ""
registercompany = ""
readreg
readcfg

End Sub

'-----------------------Open Config File-------------------------
Public Sub readcfg()
iniFileName = "MUNREC.cfg"
If GetIniS("Program", "isuninstalledversion", "True") = "True" Then Loadactivex = True
If GetIniS("Program", "isdebugversion", "False") = "True" Then Debugmode = True
If GetIniS("Program", "cleanmeetrecord", "False") = "True" Then InitMeet = True
End Sub

Public Sub readreg()
iniFileName = "MUNreg.cfg"
isreged = GetIniS("Registry", "isreged", "False")
registername = GetIniS("Registry", "RegName", "û��ע����Ϣ")
registercompany = GetIniS("Registry", "RegCorp", "û��ע����Ϣ")
End Sub

'----------------------------------ini�ļ���д�����˵�ģ�飩---------------------------------
    
    '****************************************��ȡIni�ַ���ֵ(Function)******************************************
    Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
    '�����ؼ��ʵ�ֵ
    If Temp% > 0 Then '�ؼ��ʵ�ֵ��Ϊ��
    s = ""
    For i = 1 To 144
    If Asc(Mid$(ResultString, i, 1)) = 0 Then
    Exit For
    Else
    s = s & Mid$(ResultString, i, 1)
    End If
    Next
    Else
    Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, AppProFileName(iniFileName))
    '��ȱʡֵд��INI�ļ�
    s = DefString
    End If
    GetIniS = s
    End Function

    '**************************************��ȡIni��ֵ(Function)***************************************************
    Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Integer
    Dim d As Long, s As String
    d = DefValue
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
    If d <> DefValue Then
    s = "" & d
    d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
    End If
    End Function

    '***************************************д���ַ���ֵ(Sub)**************************************************
    Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
    End Sub
    '****************************************д����ֵ(Sub)******************************************************
    Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
    Dim res%, s$
    s$ = Str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
    End Sub
    
    '���������Ѳ�֪���������һ����(keyword) ʱ
    'д��һ������ַ���ֵ�Ĺ��̣�����write����д��һ���յ�ֵʵ�ֵģ�'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
    'Dim retval As Integer
    'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
