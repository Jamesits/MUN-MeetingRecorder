VERSION 5.00
Begin VB.Form meetsetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "会议设置"
   ClientHeight    =   3195
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
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "会议名称"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "meetsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

