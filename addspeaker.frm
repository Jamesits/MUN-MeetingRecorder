VERSION 5.00
Begin VB.Form addspeaker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加到主发言名单"
   ClientHeight    =   1215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
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
   ScaleHeight     =   1215
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
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
      TabIndex        =   1
      Top             =   600
      Width           =   1215
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
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "请输入要添加到主发言名单中的成员姓名："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "addspeaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub OKButton_Click()
Unload Me
End Sub

Private Sub Text1_Change()
inputdata = Text1.Text
End Sub
