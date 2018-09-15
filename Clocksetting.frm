VERSION 5.00
Begin VB.Form Clocksetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定时器设置"
   ClientHeight    =   2895
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5055
   Icon            =   "Clocksetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.Frame viewsetting 
      Caption         =   "显示设置"
      Height          =   2175
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton Option2 
         Caption         =   "显示剩余时间"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "显示已经过时间"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "定时器设置"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "秒"
         Height          =   255
         Left            =   2160
         TabIndex        =   9
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "分钟"
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "小时"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   480
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "时间："
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Clocksetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public newsetting As Integer


Private Sub CancelButton_Click()
Clocksetting.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Clocksetting.Hide
End Sub

Private Sub OKButton_Click()
newsetting = Val(Text1.Text) * 3600 + Val(Text2.Text) * 60 + Val(Text3.Text)
Clocksetting.Hide
Form1.Timesetting1 = newsetting
Form1.timesetting.Caption = Form1.formattime(newsetting)
Form1.Refresh
MsgBox "时间将被重置！", vbOKCancel, "计时器"
Form1.resetclock
End Sub

Private Sub Option1_Click()
Form1.showgonetime
End Sub

Private Sub Option2_Click()
Form1.showtimeleft
End Sub
