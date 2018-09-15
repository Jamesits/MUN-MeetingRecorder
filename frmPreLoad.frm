VERSION 5.00
Begin VB.Form frmPreLoad 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MUN Recorder"
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8445
   FillColor       =   &H00FFFFFF&
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
   Picture         =   "frmPreLoad.frx":0000
   ScaleHeight     =   3450
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
End
Attribute VB_Name = "frmPreLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_Load()
Me.Hide
Me.Width = 8415
Me.Height = 3450
Me.Show
Me.Refresh
Print "MUN Meeting Recorder 4.0 Alpha"
Print "Version:" & getVersionString
Print "This version is written by James Swineson."
Print "Model United Nation Meeting Recorder is licensed under a Creative Commons 署名-非商业性使用-相同方式共享 2.5 中国大陆 License基于sourceforge.net上的作品创作。"
Print "Copyright (c) 2009-2011 Createnhance Solutions.All rights reserved."
Me.Refresh
Print "Status:Check Updates"
'Shell "updater\Update.exe /checknow"
Sleep 1000
FrmMain.Show
Unload Me
End Sub
