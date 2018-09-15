VERSION 5.00
Begin VB.Form frmreg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "注册"
   ClientHeight    =   2550
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5880
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label4 
      Caption         =   "Createnhance Team保证不会将此信息发送到网络的任何地方，且不会将此信息用于标注用户外的任何用途。"
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "公司："
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "姓名："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请输入您的注册信息："
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1800
   End
End
Attribute VB_Name = "frmreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then OKButton_Click
End Sub

Private Sub Form_Load()
iniFileName = "MUNreg.cfg"
Text1.Text = GetIniS("Registry", "RegName", "")
Text2.Text = GetIniS("Registry", "RegCorp", "")
frmreg.Show
End Sub

Private Sub OKButton_Click()
iniFileName = "MUNreg.cfg"
Common.isreged = "True"
SetIniS "Registry", "isreged", Common.isreged
registername = Text1.Text
SetIniS "Registry", "RegName", Text1.Text
registercompany = Text2.Text
SetIniS "Registry", "RegCorp", Text2.Text
MessageBox Me.hwnd, "注册信息将在下次启动时生效！", "注册软件", vbOKOnly
Unload Me
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then OKButton_Click
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then OKButton_Click
End Sub

