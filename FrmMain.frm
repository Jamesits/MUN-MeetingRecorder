VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{E7A0891C-94FE-4D84-B694-12E7EF672CAF}#1.0#0"; "ControlContainer.ocx"
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MUN Meeting Recorder"
   ClientHeight    =   7425
   ClientLeft      =   4140
   ClientTop       =   3330
   ClientWidth     =   11760
   FillColor       =   &H0000FF00&
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
   WindowState     =   2  'Maximized
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
   Begin 工程1.ControlContainer CContainer1 
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3201
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
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
         Height          =   1245
         Left            =   45
         TabIndex        =   3
         Top             =   180
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
      HideSelection   =   0   'False
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      BulletIndent    =   4
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmMain.frx":0004
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
        X As Long
        Y As Long
End Type
Const ratio = 2 / 3 '窗口切分比例
Const stratio = 3 / 4 '纵向切分比例
'改变窗口大小
Private InitWidth As Long
Private InitHeight As Long
'/改变窗口大小
Private RTFundoflag As Boolean

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
Me.Show

'MNUPop.Refresh

DebugEvent
End Sub

Private Sub Form_Resize()
If Me.Height <= 945 Or Me.Width <= 2250 Then
Me.Height = IIf(Me.Height <= 945, 945, Me.Height)
Me.Width = IIf(Me.Width <= 2250, 2250, Me.Width)
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

Private Sub MnuDyn_Click(Index As Integer)
Load MnuDyn(MnuDyn.UBound + 1)
End Sub

Private Sub MNUExit_Click()
End
End Sub

Public Sub InsertPicture(ByVal path As String)
'Dim tempclip, clipformat As Integer
Dim a As Object
Set a = CreateObject("WScript.shell")
'tempclip = Clipboard.GetText
'clipformat = Clipboard.GetFormat
Clipboard.Clear
Clipboard.SetData LoadPicture(path)
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^V"
'Clipboard.SetText tempclip
Set a = Nothing
End Sub

Private Sub NormalEvents_Timer()
Me.Refresh
End Sub

Private Sub arrangeControls()

'With Shape1
'.Left = MNUPop.Left - 150
'.Top = MNUPop.Top - 50
'.Height = MNUPop.Height + 100
'.Width = MNUPop.Width + 300
'MNUPop.AutoSize = False
'MNUPop.Left = .Left
''MNUPop.Top = .Top
'MNUPop.Height = .Height
'MNUPop.Width = .Width
'.BackColor = RGB(20, 143, 203)
'End With
'TODO:Solve the problem below
With TB1
.Top = 0
.Width = Int(Me.Width * ratio)
.Left = Me.Width - .Width
End With

With RichTextBox1
.Top = TB1.Top + TB1.Height
.Width = TB1.Width - 220
.Left = TB1.Left
.Height = Me.Height - .Top - TB1.Height - 172
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

Private Sub TB1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Index
Case 1 '新建
Case 2 '打开
Case 3 '保存
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
Case 11 '复制
Case 12 '粘贴
Case 14 '粗体
Case 15 '下划线
Case 16 '斜体
Case 17 '字体
Case 18 '插入图像
CD1.Filter = "Images(*.jpg;*.bmp)|*.jpg;*.bmp"
CD1.ShowOpen
InsertPicture (CD1.FileName)
Case 20 '菜单
PopupMenu PopMNU
End Select
End Sub

Public Sub undo()
Dim a As Object
Set a = CreateObject("WScript.shell")
FrmMain.RichTextBox1.SetFocus
a.SendKeys "^Z"
Set a = Nothing
End Sub
