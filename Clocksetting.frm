VERSION 5.00
Begin VB.Form Clocksetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "定时器设置"
   ClientHeight    =   2670
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5235
   Icon            =   "Clocksetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Addset 
      Caption         =   "修改设置"
      Height          =   975
      Left            =   2880
      TabIndex        =   13
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton Option4 
         Caption         =   "在原时间上增加"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "覆盖设置"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame viewsetting 
      Caption         =   "显示设置（立即生效）"
      Height          =   855
      Left            =   2880
      TabIndex        =   10
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton Option2 
         Caption         =   "显示剩余时间"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "显示已经过时间"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "定时器设置"
      Height          =   1935
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
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton clocksettingOKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2160
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


Private Sub Form_Load()
Option1.Value = True
Option3.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Clocksetting.Hide
End Sub

Public Sub clocksettingOKButton_Click()
If Option3.Value = True Then
newsetting = Val(Text1.Text) * 3600 + Val(Text2.Text) * 60 + Val(Text3.Text)
Clocksetting.Hide
frm.Timesetting1 = newsetting
frm.timesetting.Caption = frm.formattime(newsetting)
frm.Refresh
MsgBox "时间将被重置！", vbOKCancel, "计时器"
frm.resetclock
Else
newsetting = Val(Text1.Text) * 3600 + Val(Text2.Text) * 60 + Val(Text3.Text)
Clocksetting.Hide
frm.Timesetting1 = newsetting + frm.Timesetting1
frm.timesetting.Caption = frm.formattime(frm.Timesetting1)
frm.Refresh
'MsgBox "时间将被重置！", vbOKCancel, "计时器"
End If
Option1.Value = True
Option3.Value = True
End Sub

Public Sub clocksettingOK()
If Option3.Value = True Then
Clocksetting.Hide
frm.Timesetting1 = newsetting
frm.timesetting.Caption = frm.formattime(newsetting)
frm.Refresh
MsgBox "时间将被重置！", vbOKCancel, "计时器"
frm.resetclock
Else
Clocksetting.Hide
frm.Timesetting1 = newsetting + frm.Timesetting1
frm.timesetting.Caption = frm.formattime(frm.Timesetting1)
frm.Refresh
'MsgBox "时间将被重置！", vbOKCancel, "计时器"
End If
Option1.Value = True
Option3.Value = True
End Sub

Private Sub Option1_Click()
frm.showgonetime
End Sub

Private Sub Option2_Click()
frm.showtimeleft
End Sub

