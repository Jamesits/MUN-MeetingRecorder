VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   9720
      Begin VB.Label Regcomp 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label Regname 
         AutoSize        =   -1  'True
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   60
      End
      Begin VB.Label Register 
         AutoSize        =   -1  'True
         Caption         =   "注册到"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   540
      End
      Begin VB.Label LoadStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "加载中……"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2400
         TabIndex        =   7
         Top             =   2280
         Width           =   1395
      End
      Begin VB.Label Author 
         Caption         =   "朱焕杰"
         Height          =   255
         Left            =   6960
         TabIndex        =   6
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Image imgLogo 
         Height          =   585
         Left            =   840
         Picture         =   "frmSplash.frx":058A
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblCopyright 
         Caption         =   "版权所有"
         Height          =   255
         Left            =   6960
         TabIndex        =   2
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Createnhance Team"
         Height          =   255
         Left            =   6960
         TabIndex        =   1
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8910
         TabIndex        =   3
         Top             =   2160
         Width           =   480
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "产品"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   32.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Createnhance Team"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2355
         TabIndex        =   4
         Top             =   705
         Width           =   3600
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   1085
      ButtonWidth     =   609
      ButtonHeight    =   926
      Appearance      =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SystemInf As Integer

Private Sub Form_Load()
Dim i As Byte
    DoEvents
    lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision & IIf(Common.Beta, " Beta", " 正式版") & IIf(Common.Debugmode, "(Debug mode)", "")
    lblProductName.Caption = App.Title
    LoadStatus.Caption = "检测注册信息……"
    Regname.Caption = Common.registername
    Regcomp.Caption = Common.registercompany
    LoadStatus.Caption = "加载中……"
    SetTransparentWindow Me.hwnd, 100
    frmSplash.Show
    frmSplash.Refresh
    For i = 1 To 100
    SetTransparentWindow Me.hwnd, 100 - i
    Delay 1
    Next i
    LoadStatus.Caption = "加载组件……"
    Load FormUnload
    LoadStatus.Caption = "检查更新……"
    DoEvents
    If Dir("Update.exe") <> "" Then
    Shell "Update.exe /silent"
    End If
    LoadStatus.Caption = "检测系统状况……"
    Getsysteminf
    LoadStatus.Caption = "注册控件……"
    frmSplash.Refresh
    Delay 100
    DoEvents
    If Loadactivex Then
    If SystemInf = 64 Then
      reg "ActiveX", "%windir%\Syswow64", "comdlg32.ocx"
      frmSplash.Refresh
      reg "ActiveX", "%windir%\Syswow64", "mscomctl.ocx"
      frmSplash.Refresh
      reg "ActiveX", "%windir%\Syswow64", "RICHTX32.OCX"
    Else
      reg "ActiveX", "%windir%\System32", "comdlg32.ocx"
      frmSplash.Refresh
      reg "ActiveX", "%windir%\System32", "mscomctl.ocx"
            frmSplash.Refresh
      reg "ActiveX", "%windir%\System32", "RICHTX32.OCX"
    End If
    End If
    frmSplash.Refresh
    LoadStatus.Caption = "加载中……"
    If InitMeet Then
    LoadStatus.Caption = "重置会议"
    initini
    End If
    LoadStatus.Caption = "检测分辨率……"
    pixels.Show
End Sub

Public Sub Getsysteminf()
If Dir("%windir%\SysWOW64", vbDirectory) <> "" Then SystemInf = 64 Else SystemInf = 32
End Sub

Public Sub initini()
iniFileName = "setting"
SetIniS "Meeting", "Name", "新会议1"
SetIniS "Meeting", "Start_Y", 2011
SetIniS "Meeting", "Start_M", 7
SetIniS "Meeting", "Start_D", 23
SetIniS "Meeting", "Start_H", 7
SetIniS "Meeting", "Start_M", 0
SetIniS "Meeting", "Start_S", 0
SetIniN "Meeting", "Lasttime", 3600
End Sub
