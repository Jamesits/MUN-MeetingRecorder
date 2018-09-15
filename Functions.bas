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

'---------------------------------公用函数----------------------------------
Public Sub Runprog(ByVal path As String) '运行DOS命令
Dim pid As Long
pid = Shell("cmd /c " & path, vbHide)
hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
Do
Call GetExitCodeProcess(hProcess, ExitCode)
Loop While ExitCode = STILL_ALIVE
Call CloseHandle(hProcess)
End Sub

Public Function formattime(ByVal t As Integer) As String  '秒数格式化为时：分：秒的字符串
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


'----------------------------------ini文件读写（别人的模块）---------------------------------
    
    '****************************************获取Ini字符串值(Function)******************************************
    Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String) As String
    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, AppProFileName(iniFileName))
    '检索关键词的值
    If Temp% > 0 Then '关键词的值不为空
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
    '将缺省值写入INI文件
    s = DefString
    End If
    GetIniS = s
    End Function

    '**************************************获取Ini数值(Function)***************************************************
    Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long) As Integer
    Dim d As Long, s As String
    d = DefValue
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, AppProFileName(iniFileName))
    If d <> DefValue Then
    s = "" & d
    d = WritePrivateProfileString(SectionName, KeyWord, s, AppProFileName(iniFileName))
    End If
    End Function

    '***************************************写入字符串值(Sub)**************************************************
    Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String)
    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, AppProFileName(iniFileName))
    End Sub
    '****************************************写入数值(Sub)******************************************************
    Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long)
    Dim res%, s$
    s$ = Str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, AppProFileName(iniFileName))
    End Sub
    
    '这是我自已不知道怎样清除一个键(keyword) 时
    '写的一个清除字符串值的过程，是有write函数写入一个空的值实现的，'Sub DelIniS(ByVal SectionName As String, ByVal KeyWord As String)
    'Dim retval As Integer
    'retval = WritePrivateProfileString(SectionName, KeyWord, "", AppProFileName(iniFileName))
    'End Sub
    '其实0&表示前面的一个被清除，我多写了一个“”，如果是清除section就少写一个Key多一个“”。

    '***************************************清除KeyWord"键"(Sub)*************************************************
    Sub DelIniKey(ByVal SectionName As String, ByVal KeyWord As String)
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, KeyWord, 0&, AppProFileName(iniFileName))
    End Sub

    '如果是清除section就少写一个Key多一个“”。
    '**************************************清除 Section"段"(Sub)***********************************************
    Sub DelIniSec(ByVal SectionName As String) '清除section
    Dim RetVal As Integer
    RetVal = WritePrivateProfileString(SectionName, 0&, "", AppProFileName(iniFileName))
    End Sub

    '*************************************定义Ini文件名(Function)***************************************************
    '定义ini文件名
    Function AppProFileName(iniFileName)
    AppProFileName = Trim(App.path & "\" & iniFileName)
    End Function


    '用法: 首先 定义iniFileName="文件名" 不需要
    '这就是说，你可以赋值给iniFileName就可以写入记录，而且你可以随时写入不同的ini文件(不管这个文件是否已存在），通过修改这个公用变量。

    '然后　 DelInikey（ByVal SectionName As String, ByVal KeyWord As String） 清除键
              'DelIniSec(ByVal SectionName As String)) 清除部
              'SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Long) 写入数
              'GetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefValue As Long)读取数
              'SetIniS (ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) 写入字符
              'GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String) 读取字符


'------------------------------------------读取语言文件--------------------------------------------------------
Private Sub readlang(ByVal langpath As String)
iniFileMame = lang.ini
End Sub

'-------------------------------------------加载ocx
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

