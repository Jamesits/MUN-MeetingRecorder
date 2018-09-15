VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{E7A0891C-94FE-4D84-B694-12E7EF672CAF}#2.0#0"; "ControlContainer.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MUN Meeting Recorder"
   ClientHeight    =   7425
   ClientLeft      =   4140
   ClientTop       =   3330
   ClientWidth     =   11760
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin ComctlLib.Toolbar TB1 
      Height          =   420
      Left            =   4200
      Negotiate       =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "IL1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   20
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Description     =   "新建"
            Object.ToolTipText     =   "新建"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "打开"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "保存"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "打印"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "撤销"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.ToolTipText     =   "重复"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "查找"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "剪切"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "复制"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "粘贴"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "粗体"
            Object.Tag             =   ""
            ImageIndex      =   11
            Style           =   1
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "下划线"
            Object.Tag             =   ""
            ImageIndex      =   12
            Style           =   1
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "斜体"
            Object.Tag             =   ""
            ImageIndex      =   13
            Style           =   1
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "字体"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "插入图像"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.ToolTipText     =   "菜单"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin CContainer.ControlContainer CCConButtons 
      Height          =   675
      Left            =   -3780
      Top             =   1080
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   1191
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "微软雅黑"
      FontSize        =   9
      Begin VB.CommandButton CmdStart 
         Caption         =   "开始/暂停"
         Height          =   435
         Left            =   60
         TabIndex        =   6
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton CmdStop 
         Caption         =   "停止/重置"
         Height          =   435
         Left            =   1320
         TabIndex        =   5
         Top             =   120
         Width           =   1155
      End
      Begin VB.CommandButton CmdTimeSet 
         Caption         =   "时间设置"
         Height          =   435
         Left            =   2580
         TabIndex        =   4
         Top             =   120
         Width           =   1155
      End
      Begin VB.Label LblArray 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3770
         TabIndex        =   12
         Top             =   0
         Width           =   165
      End
   End
   Begin CContainer.ControlContainer CContainer1 
      Height          =   1755
      Left            =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3096
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "微软雅黑"
      FontSize        =   9
      Begin VB.Label LblLeftTime 
         BackColor       =   &H00FFFFFF&
         Caption         =   "剩余 00:00:00"
         Height          =   255
         Left            =   2820
         TabIndex        =   11
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label LblPassTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "已经过 00:00:00"
         Height          =   255
         Left            =   2715
         TabIndex        =   10
         Top             =   1260
         Width           =   1320
      End
      Begin VB.Label LblMeetTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "会议时间 00:00:00"
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label LblSetTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "定时 00:00:00"
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   1500
         Width           =   1140
      End
      Begin VB.Label LblSysTime 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "系统时间 0000-00-00 00:00:00"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   1260
         Width           =   2550
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   48
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   180
         TabIndex        =   3
         Top             =   60
         Width           =   3870
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      IMEMode         =   2  'OFF
      IntegralHeight  =   0   'False
      ItemData        =   "FrmMain.frx":0000
      Left            =   0
      List            =   "FrmMain.frx":0002
      TabIndex        =   2
      Top             =   1800
      Width           =   4155
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   10320
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer NormalEvents 
      Interval        =   35
      Left            =   11340
      Top             =   6960
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   7035
      Left            =   4140
      TabIndex        =   0
      Top             =   420
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12409
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      BulletIndent    =   4
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmMain.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList IL1 
      Left            =   10920
      Top             =   6900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483636
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   16
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":00B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":01C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":02D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":03E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":04F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":060B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":071D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":082F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0941
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0A53
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0B65
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0C77
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0D89
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":0E9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":11ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":153F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu PopMNU 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu MTitle 
         Caption         =   "模拟联合国会议记录器"
         Enabled         =   0   'False
      End
      Begin VB.Menu MVer 
         Caption         =   "版本"
         Enabled         =   0   'False
      End
      Begin VB.Menu MSeperator_jahflhfdjkahklsjhfafdas 
         Caption         =   "-"
      End
      Begin VB.Menu MnuDyn 
         Caption         =   "动态菜单测试"
         Index           =   0
      End
      Begin VB.Menu MnuSeperator_fadjgkajsgj 
         Caption         =   "-"
      End
      Begin VB.Menu MFile 
         Caption         =   "文件(&F)"
      End
      Begin VB.Menu MEdit 
         Caption         =   "编辑(&E)"
      End
      Begin VB.Menu MClock 
         Caption         =   "定时器(&C)"
         Begin VB.Menu MClkStart 
            Caption         =   "启动(&S)"
            Shortcut        =   ^S
         End
         Begin VB.Menu MClkPause 
            Caption         =   "暂停(&P)"
            Shortcut        =   ^P
         End
         Begin VB.Menu MClkReset 
            Caption         =   "复位(&R)"
            Shortcut        =   ^R
         End
         Begin VB.Menu MNUSeperator_fajhasdfdashlfkashfjkh 
            Caption         =   "-"
         End
         Begin VB.Menu MClkSetting 
            Caption         =   "定时器设定(&T)"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu MSetting 
         Caption         =   "设置(&S)"
         Begin VB.Menu MLang 
            Caption         =   "语言"
            Begin VB.Menu MDelpLang 
               Caption         =   "开发语言（简体中文）"
               Checked         =   -1  'True
            End
            Begin VB.Menu MNewLang 
               Caption         =   "语言"
               Checked         =   -1  'True
               Index           =   0
               Visible         =   0   'False
            End
         End
         Begin VB.Menu MNUSeperator_sjdahfljkfhaksf 
            Caption         =   "-"
         End
         Begin VB.Menu MSysSetting 
            Caption         =   "系统设置"
         End
      End
      Begin VB.Menu MHelp 
         Caption         =   "帮助(&H)"
         Begin VB.Menu MHlpTopic 
            Caption         =   "帮助主题"
         End
         Begin VB.Menu MNUSeperator_afhdlkfhajkfhsdl 
            Caption         =   "-"
         End
         Begin VB.Menu MAbout 
            Caption         =   "关于(&A)"
         End
      End
      Begin VB.Menu MNUSeperator_fahjfhlkajhfkjdhljkfhaks 
         Caption         =   "-"
      End
      Begin VB.Menu MNUExit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function SetCapture Lib "user32 " (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32 " () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type POINTAPI
        x As Long
        Y As Long
End Type
Const ratio = 2 / 3 '窗口切分比例
Const stratio = 3 / 4 '纵向切分比例
'改变窗口大小
Private InitWidth As Long
Private InitHeight As Long
'/改变窗口大小
Private RTFundoflag As Boolean
Private butPopedFlag As Boolean



Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
PopMNU.Visible = False
Me.WindowState = vbMaximized
'改变窗口大小
InitWidth = ScaleWidth
InitHeight = ScaleHeight
Dim Ctl As Control
On Error Resume Next
For Each Ctl In Me
    Ctl.Tag = Ctl.Left & " " & Ctl.Top & " " & Ctl.Width & " " & Ctl.Height & " "
    Ctl.Tag = Ctl.Tag & Ctl.FontSize & " "
Next Ctl
On Error GoTo 0
'/改变窗口大小
MVer.Caption = "版本 " & getVersionString
RTFundoflag = False
With CCConButtons
.BackColor = vbWhite
.Left = 0
.Width = 3795
.Top = 1140
.Height = 675
End With
CCConButtons.Left = 0
CCConButtons.Top = 1080
butPopedFlag = True
LblArray = "<"
Me.Show
Me.Refresh
DebugEvent
Sleep 500
MovePort CCConButtons, 20, -CCConButtons.Width + LblArray.Width
butPopedFlag = False
LblArray = ">"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'注意！在这行命令执行之前保存！！！
RichTextBox1.TextRTF = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Height <= 945 Or Me.Width <= 2250 Then
'Me.Height = IIf(Me.Height <= 945, 945, Me.Height)
'Me.Width = IIf(Me.Width <= 2250, 2250, Me.Width)
ReleaseCapture
Else

'改变窗口大小
Dim D(4) As Double
Dim i As Long
Dim TempPos As Long
Dim StartPos As Long
Dim Ctl As Control
Dim TempVisible As Boolean
Dim ScaleX As Double
Dim ScaleY As Double

ScaleX = ScaleWidth / InitWidth
ScaleY = ScaleHeight / InitHeight
On Error Resume Next
For Each Ctl In Me
    TempVisible = Ctl.Visible
    Ctl.Visible = False
    StartPos = 1
    ' 读取 Control 的原始位置、大小、字型大小
    For i = 0 To 4
        TempPos = InStr(StartPos, Ctl.Tag, " ", vbTextCompare)
        If TempPos > 0 Then
            D(i) = Mid(Ctl.Tag, StartPos, TempPos - StartPos)
            StartPos = TempPos + 1
        Else
            D(i) = 0
        End If
        ' 根据比例设定 Control 的位置、大小、字型大小
        Ctl.Move D(0) * ScaleX, D(1) * ScaleY, D(2) * ScaleX, D(3) * ScaleY
        'Ctl.Width = D(2) * ScaleX
        'Ctl.Height = D(3) * ScaleY
        If ScaleX < ScaleY Then
            Ctl.FontSize = D(4) * ScaleX
        Else
            Ctl.FontSize = D(4) * ScaleY
        End If
    Next i
    Ctl.Visible = TempVisible
Next Ctl
On Error GoTo 0
'/改变窗口大小



arrangeControls

End If
End Sub

Private Sub LblArray_Click()
If butPopedFlag = False Then
MovePort CCConButtons, 20, 0
butPopedFlag = True
LblArray = "<"
Else
MovePort CCConButtons, 20, -CCConButtons.Width + LblArray.Width
butPopedFlag = False
LblArray = ">"
End If
End Sub

Private Sub MAbout_Click()
frmAbout.Show
End Sub

Private Sub MnuDyn_Click(Index As Integer)
Load MnuDyn(MnuDyn.UBound + 1)
End Sub

Private Sub MNUExit_Click()
End
End Sub

Public Sub InsertPicture(ByVal path As String)
Dim a As Object, tempd As Variant, tempt As String
On Error Resume Next
Set a = CreateObject("WScript.shell")
tempd = Clipboard.GetData
tempt = Clipboard.GetText
Clipboard.Clear
Clipboard.SetData LoadPicture(path)
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^V"
Sleep 1000
'Clipboard.Clear
Clipboard.SetData tempd
Clipboard.SetText tempt
Set a = Nothing
End Sub

Public Sub copytext()
Dim a As Object
Set a = CreateObject("WScript.shell")
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^C"
Set a = Nothing
End Sub

Public Sub cuttext()
Dim a As Object
Set a = CreateObject("WScript.shell")
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^X"
Set a = Nothing
End Sub

Public Sub pastetext()
Dim a As Object
Set a = CreateObject("WScript.shell")
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^V"
Set a = Nothing
End Sub



Private Sub NormalEvents_Timer()
Me.Refresh
LblSysTime = "系统时间 " & Format$(Date, "yyyy-mm-dd") & " " & Format$(Time, "hh:mm:ss")
'Dim p As POINTAPI
'GetCursorPos p
'If p.X < 500 And p.Y > CCConButtons.Top And p.Y < CCConButtons.Top + CCConButtons.Height And butPopedFlag = False Then
'MovePort CCConButtons, 20, 0
'End If


End Sub

Private Sub arrangeControls()
With TB1
.Top = 0
.Width = Int(Me.Width * ratio)
.Left = Me.Width - .Width
End With

With RichTextBox1
.Top = TB1.Top + TB1.Height
.Width = TB1.Width - 220
.Left = TB1.Left
.Height = Me.Height - .Top - TB1.Height - 130
End With

With List1
.Left = 0
.Top = Int(Me.Height * (1 - stratio))
.Height = Me.Height - .Top - 550
.Width = Int(Me.Width * (1 - ratio)) + 10
End With

With CContainer1
.Top = 0
.Left = 0
.Width = Int(Me.Width * (1 - ratio))
.Height = Int(Me.Height * (1 - stratio))
End With
With Label1
.Top = 0
.Left = Int((CContainer1.Width - Label1.Width) / 2)

End With

With LblArray
.Left = CCConButtons.Width - .Width
.Height = CCConButtons.Height
End With
End Sub

Public Sub DebugEvent()
Dim i As Integer
For i = 1 To 50: List1.AddItem ("Speaker" & i): Next i
End Sub

Private Sub RichTextBox1_Change()
If RTFundoflag = False Then
TB1.Buttons(6).Enabled = True
TB1.Buttons(7).Enabled = False
End If
End Sub

Private Sub RichTextBox1_SelChange()
On Error Resume Next
TB1.Buttons(14).Value = -CInt(RichTextBox1.SelBold)
TB1.Buttons(15).Value = -CInt(RichTextBox1.SelUnderline)
TB1.Buttons(16).Value = -CInt(RichTextBox1.SelItalic)
End Sub

Private Sub TB1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1 '新建
NewFile
Case 2 '打开
OpenFile
Case 3 '保存
SaveFile
Case 4 '打印

Case 6 '撤销
RTFundoflag = True
TB1.Buttons(6).Enabled = False
TB1.Buttons(7).Enabled = True
undo
RTFundoflag = False
Case 7 '重复
RTFundoflag = True
TB1.Buttons(6).Enabled = True
TB1.Buttons(7).Enabled = False
undo
RTFundoflag = False
Case 9 '查找
Case 10 '剪切
cuttext
Case 11 '复制
copytext
Case 12 '粘贴
pastetext
Case 14 '粗体
RichTextBox1.SelBold = TB1.Buttons(14).Value
Case 15 '下划线
RichTextBox1.SelUnderline = TB1.Buttons(15).Value
Case 16 '斜体
RichTextBox1.SelItalic = TB1.Buttons(16).Value
Case 17 '字体
On Error Resume Next
CD1.FontName = RichTextBox1.SelFontName
CD1.FontSize = RichTextBox1.SelFontSize
CD1.FontBold = RichTextBox1.SelBold
CD1.FontStrikethru = RichTextBox1.SelStrikeThru
CD1.FontItalic = RichTextBox1.SelItalic
CD1.FontUnderline = RichTextBox1.SelUnderline
On Error GoTo ext
CD1.CancelError = True
CD1.ShowFont
RichTextBox1.SelFontName = CD1.FontName
RichTextBox1.SelFontSize = CD1.FontSize
RichTextBox1.SelBold = CD1.FontBold
RichTextBox1.SelStrikeThru = CD1.FontStrikethru
RichTextBox1.SelItalic = CD1.FontItalic
RichTextBox1.SelUnderline = CD1.FontUnderline
Case 18 '插入图像
On Error GoTo ext
CD1.CancelError = True
CD1.Filter = "Images(*.jpg;*.bmp)|*.jpg;*.bmp"
CD1.ShowOpen
InsertPicture (CD1.FileName)
Case 20 '菜单
'PopupMenu PopMNU, , TB1.Buttons(20).Left + TB1.Left + 15, TB1.Top + TB1.Buttons(20).Top + TB1.Buttons(20).Height + 15
PopupMenu PopMNU, , TB1.Buttons(20).Left + TB1.Left + 15, TB1.Top + TB1.Height
End Select
ext:
End Sub

Public Sub undo()
Dim a As Object
Set a = CreateObject("WScript.shell")
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^Z"
Set a = Nothing
End Sub


Private Sub OpenFile()
On Error GoTo ext
CD1.CancelError = True
CD1.DialogTitle = "打开"
CD1.Filter = "所有支持的格式(*.rtf;*.txt)|*.rtf;*.txt|Rich Text Format(*.rtf)|*.rtf|文本文档(*.txt)|*.txt|所有文件(*.*)|*.*"
CD1.ShowOpen
RichTextBox1.LoadFile CD1.FileName
ext:
End Sub


Private Sub SaveFile()
On Error GoTo ext
CD1.CancelError = True
CD1.DialogTitle = "保存"
CD1.Filter = "Rich Text Format(*.rtf)|*.rtf|文本文档(*.txt)|*.txt|所有文件(*.*)|*.*"
CD1.ShowSave
If CD1.FilterIndex = 2 Then
RichTextBox1.SaveFile CD1.FileName, 1
Else
RichTextBox1.SaveFile CD1.FileName, 0
End If

ext:
End Sub

Private Sub NewFile()

End Sub
