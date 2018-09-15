VERSION 5.00
Begin VB.Form pixels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选择分辨率"
   ClientHeight    =   1725
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "pixels.frx":0000
      Left            =   240
      List            =   "pixels.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "建议分辨率："
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请选择适合您的屏幕的分辨率："
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2520
   End
End
Attribute VB_Name = "pixels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Combo1.Text = "1024*768"
Label3.Caption = GetSystemMetrics(SM_CXSCREEN) & "*" & GetSystemMetrics(SM_CYSCREEN)
End Sub

Private Sub OKButton_Click()
pixels.Hide
Select Case Combo1.Text
Case "640*480"
Set frm = Form640480
Case "800*600"
Set frm = Form800600
Case "1024*768"
Set frm = Form1024768
Case "1280*800"
Set frm = Form1280800
Case "1440*900"
Set frm = Form1440900
End Select
pixels.Hide
    frmSplash.LoadStatus.Caption = "加载主窗体……"
    Load frm
    DoEvents
    For i = 1 To 50
    SetTransparentWindow frmSplash.hwnd, i * 2
    Delay 1
    Next i
    frm.Show
    frmSplash.Hide
End Sub
