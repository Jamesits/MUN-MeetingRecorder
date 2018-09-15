Attribute VB_Name = "Functions"
Public Sub Main()

Debugmode = True
Loadactivex = False
UnloadactiveX = False
Beta = True
InitMeet = False

Set frm = Form1440900
Load Common
Load frmSplash
End Sub

'---------------------------------���ú���----------------------------------
Public Sub Runprog(ByVal path As String) '����DOS����
Dim pid As Long
pid = Shell("cmd /c " & path, vbHide)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
Do
Call GetExitCodeProcess(hProcess, ExitCode)
Loop While ExitCode = STILL_ALIVE
Call CloseHandle(hProcess)
End Sub

Public Function formattime(ByVal t As Integer) As String  '������ʽ��Ϊʱ���֣�����ַ���
Dim g, h, i As String
g = Int(t / 3600)
h = Int((t - 3600 * Val(g)) / 60)
i = Int(t - 3600 * Val(g) - 60 * Val(h))
If Val(g) < 10 Then g = "0" & g
If Val(h) < 10 Then h = "0" & h
If Val(i) < 10 Then i = "0" & i
formattime = g & " : " & h & " : " & i
End Function

Public Sub reg(ByVal pathfrom, pathto, name As String)
 If Dir(pathto & "\" & name) = "" Then
      Runprog ("Copy " & App.path & "\" & pathfrom & "\" & name & " " & pathto)
      Runprog ("regsvr32 /s " & pathto & "\" & name)
      Else
      Runprog ("regsvr32 /s /u " & pathto & "\" & name)
      Runprog ("regsvr32 /s " & pathto & "\" & name)
     End If
End Sub

Public Sub Delay(ByVal ms As Long)
Call Sleep(ms)
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
    'End Sub
    '��ʵ0&��ʾǰ���һ����������Ҷ�д��һ����������������section����дһ��Key��һ��������

    '***************************************���KeyWord"��"(Sub)*************************************************
    Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
    End Sub

    '��������section����дһ��Key��һ��������
    '**************************************��� Section"��"(Sub)***********************************************
    Sub DelIniSec(ByVal SectionName As String) '���section
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
    End Sub

    '*************************************����Ini�ļ���(Function)***************************************************
    '����ini�ļ���
    Function AppProFileName(iniFileName)
    AppProFileName = Trim(App.path & "\" & iniFileName)
    End Function


    '�÷�: ���� ����iniFileName="�ļ���" ����Ҫ
    '�����˵������Ը�ֵ��iniFileName�Ϳ���д���¼�������������ʱд�벻ͬ��ini�ļ�(��������ļ��Ƿ��Ѵ��ڣ���ͨ���޸�������ñ�����

    'Ȼ�� DelInikey��ByVal SectionName As String, ByVal KeyWord As String�� �����
              'DelIniSec(ByVal SectionName As String)) �����
              'SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long) д����
              'GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long)��ȡ��
              'SetIniS (ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) д���ַ�
              'GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) ��ȡ�ַ�


'------------------------------------------��ȡ�����ļ�--------------------------------------------------------
Private Sub readlang(ByVal langpath As String)
iniFileMame = lang.ini
End Sub

'-------------------------------------------����ocx
Sub LoadOCX(ByVal path As String, ByVal filename As String, ByVal resnumber As Byte)
Dim byt() As Byte
Dim File As String

'File = Environ("windir") & "\system32\PICCLP32.OCX"
File = path + "\" + filename
If Len(Dir(File)) = 0 Then
byt = LoadResData(resnumber, "CUSTOM")
Open File For Binary As #1
Put #1, 1, byt()
Close #1
Shell ("regsvr32 /s filename")
End If
End Sub

