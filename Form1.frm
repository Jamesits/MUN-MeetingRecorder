VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1024768 
   Caption         =   "ģ�����Ϲ������¼��(1024*768)"
   ClientHeight    =   10230
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   26.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   682
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   Begin VB.Frame Frame3 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3840
      TabIndex        =   40
      Top             =   2160
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton Command17 
         Caption         =   "+1����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   48
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Command16 
         Caption         =   "+2����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   47
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+5����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   46
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command14 
         Caption         =   "+30��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   45
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "30��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command13 
         Caption         =   "5����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   43
         Top             =   1440
         Width           =   975
      End
      Begin VB.CommandButton Command12 
         Caption         =   "2����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "1����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   41
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Timer Refreshform 
      Interval        =   10
      Left            =   0
      Top             =   600
   End
   Begin VB.Frame frame2 
      Caption         =   "�����¼"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   5160
      TabIndex        =   21
      Top             =   0
      Width           =   9975
      Begin RichTextLib.RichTextBox Text1 
         Height          =   9255
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   16325
         _Version        =   393217
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Form1.frx":058A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "΢���ź�"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   390
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "�½�"
               Object.ToolTipText     =   "�½�"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "��"
               Object.ToolTipText     =   "��"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Undo"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "�ظ�"
               Object.ToolTipText     =   "�ظ�"
               ImageKey        =   "Redo"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ճ��"
               Object.ToolTipText     =   "ճ��"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "����"
               Object.ToolTipText     =   "����"
               ImageKey        =   "Bold"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "�»���"
               Object.ToolTipText     =   "�»���"
               ImageKey        =   "Underline"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "б��"
               Object.ToolTipText     =   "б��"
               ImageKey        =   "Italic"
            EndProperty
         EndProperty
         Begin VB.CommandButton Fontset 
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   38
            Top             =   0
            Width           =   735
         End
         Begin VB.ComboBox Fontsizebox 
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7680
            TabIndex        =   37
            Text            =   "ѡ���ֺ�"
            Top             =   0
            Width           =   1335
         End
         Begin VB.ComboBox Fontlist 
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   0
            Width           =   2775
         End
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9975
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   3493
            MinWidth        =   1129
            Text            =   "ģ�����Ϲ������¼��"
            TextSave        =   "ģ�����Ϲ������¼��"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2011-8-29"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "19:01"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            TextSave        =   "Ins"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Enabled         =   0   'False
            TextSave        =   "SCRL"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Speaklist 
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   4935
      Begin VB.Frame Framedel 
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton Command8 
            Caption         =   "ȷ��"
            Default         =   -1  'True
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   27
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ȡ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   26
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "ȷʵҪɾ����"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   28
            Top             =   360
            Width           =   1260
         End
      End
      Begin VB.Frame Frameadd 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton Command5 
            Caption         =   "ȡ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   20
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "ȷ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   19
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox Textadd 
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label7 
            Caption         =   "�������ˣ�"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frameedit 
         Caption         =   "�޸�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   4575
         Begin VB.TextBox Textedit 
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   480
            Width           =   4215
         End
         Begin VB.CommandButton Command6 
            Cancel          =   -1  'True
            Caption         =   "ȡ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3480
            TabIndex        =   31
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            Caption         =   "ȷ��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   30
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "�޸���������"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   1080
         End
      End
      Begin VB.CommandButton command1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "�޸�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ɾ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   24
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6825
         ItemData        =   "Form1.frx":0637
         Left            =   120
         List            =   "Form1.frx":0639
         TabIndex        =   12
         Top             =   600
         Width           =   4575
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   7005
      Top             =   4875
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":063B
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":074D
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":085F
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0971
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0A83
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B95
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0CA7
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DB9
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0ECB
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FDD
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10EF
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1201
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1313
            Key             =   "Italic"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "��ʱ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin ģ�����Ϲ������¼��.cSysTray cSysTray1 
         Left            =   360
         Top             =   0
         _ExtentX        =   900
         _ExtentY        =   900
         InTray          =   0   'False
         TrayIcon        =   "Form1.frx":1425
         TrayTip         =   "ģ�����Ϲ������¼��"
      End
      Begin VB.CommandButton Timersetting 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3720
         TabIndex        =   39
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton stopclock 
         Caption         =   "��λ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton pauseclock 
         Caption         =   "��ͣ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Startclock 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
      Begin VB.Timer timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   4560
         Top             =   120
      End
      Begin VB.Timer clockcontrol 
         Interval        =   50
         Left            =   4200
         Top             =   120
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " 00 : 00 : 00"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   42
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1125
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4650
      End
      Begin VB.Label timeleft 
         AutoSize        =   -1  'True
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ʣ��ʱ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   1320
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�Ѿ���"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label timesetting 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ʱʱ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label systemtime 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ϵͳʱ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   720
      End
   End
   Begin VB.Menu MFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu MNew 
         Caption         =   "�½�"
         Shortcut        =   ^N
      End
      Begin VB.Menu MOpen 
         Caption         =   "��..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MSave 
         Caption         =   "����..."
         Shortcut        =   ^S
      End
      Begin VB.Menu M_5 
         Caption         =   "-"
      End
      Begin VB.Menu MSet 
         Caption         =   "����"
         Enabled         =   0   'False
      End
      Begin VB.Menu M_1 
         Caption         =   "-"
      End
      Begin VB.Menu MExit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu MClock 
      Caption         =   "��ʱ��(&C)"
      Begin VB.Menu MClockstart 
         Caption         =   "����"
      End
      Begin VB.Menu MClockpause 
         Caption         =   "��ͣ"
      End
      Begin VB.Menu MClockclear 
         Caption         =   "��λ"
      End
      Begin VB.Menu M_3 
         Caption         =   "-"
      End
      Begin VB.Menu MClockSetting 
         Caption         =   "��ʱ������"
      End
   End
   Begin VB.Menu MHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu MHelpindex 
         Caption         =   "��������"
         Enabled         =   0   'False
      End
      Begin VB.Menu MInternet 
         Caption         =   "�ٷ���վ"
      End
      Begin VB.Menu Mreg 
         Caption         =   "ע��"
      End
      Begin VB.Menu M_4 
         Caption         =   "-"
      End
      Begin VB.Menu MCheckUpdate 
         Caption         =   "������"
      End
      Begin VB.Menu MUpdateSetting 
         Caption         =   "��������"
      End
      Begin VB.Menu M_2 
         Caption         =   "-"
      End
      Begin VB.Menu MAbout 
         Caption         =   "����"
      End
   End
   Begin VB.Menu MTray 
      Caption         =   "���������̲˵�"
      Visible         =   0   'False
      Begin VB.Menu Mshowmainwindows 
         Caption         =   "��ʾ������"
      End
      Begin VB.Menu M12 
         Caption         =   "-"
      End
      Begin VB.Menu MQuit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1024768"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Timesetting1 As Double
Public Timegone As Double
Public timeleft1 As Double
Public tempstr As String
Public Timernum As Double
Public filename As String
Public onloading As Boolean
Public editing As Integer
Public Bigclockshow As Integer
Public isclockstarted As Boolean
Public starth, startm, starts As Integer
Public Meettiming As Integer
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Command10_Click()
Clocksetting.Option3.Value = True
Clocksetting.newsetting = 60
Clocksetting.clocksettingOK
End Sub

Private Sub Command11_Click()
Clocksetting.Option3.Value = True
Clocksetting.newsetting = 30
Clocksetting.clocksettingOK
End Sub

Private Sub Command12_Click()
Clocksetting.Option3.Value = True
Clocksetting.newsetting = 120
Clocksetting.clocksettingOK
End Sub

Private Sub Command13_Click()
Clocksetting.Option3.Value = True
Clocksetting.newsetting = 300
Clocksetting.clocksettingOK
End Sub

Private Sub Command14_Click()
Clocksetting.Option3.Value = False
Clocksetting.newsetting = 30
Clocksetting.clocksettingOK
End Sub

Private Sub Command15_Click()
Clocksetting.Option3.Value = False
Clocksetting.newsetting = 300
Clocksetting.clocksettingOK
End Sub

Private Sub Command16_Click()
Clocksetting.Option3.Value = False
Clocksetting.newsetting = 120
Clocksetting.clocksettingOK
End Sub

Private Sub Command17_Click()
Clocksetting.Option3.Value = False
Clocksetting.newsetting = 60
Clocksetting.clocksettingOK
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

'------------------------------------------�������ж��--------------------------------
Private Sub Form_Load()
onloading = True
List1.Clear
Timernum = 0
Timesetting1 = 300
Meettiming = 0
a = 0
b = 0
c = 0
timesetting.Caption = " " & formattime(Timesetting1)
List1.Clear
Bigclockshow = 0
StatusBar1.Panels(1).AutoSize = sbrContents
StatusBar1.Panels(1).Text = App.Title & " �汾 " & App.Major & "." & App.Minor & "." & App.Revision & IIf(Common.Beta, " Beta", " ��ʽ��") & IIf(Common.Debugmode, "(Debug mode)", "")
If Dir("MUNAutoSave.rtf") <> "" Then
If MessageBox(Me.hwnd, "��⵽�Զ�����Ļ����¼���Ƿ����룿", "ģ�����Ϲ������¼��", vbOKCancel) = 1 Then Text1.LoadFile "MUNAutoSave.rtf"
DoEvents
Runprog "del MUNAutoSave.rtf"
End If
If Common.Debugmode Then
Command13.Enabled = True
Command14.Enabled = True
Command15.Enabled = True
Command16.Enabled = True
Command11.Enabled = True
Command12.Enabled = True
Command10.Enabled = True
Startmeeting.Enabled = True
MNew.Enabled = True
MOpen.Enabled = True
MSave.Enabled = True
MSaveas.Enabled = True
Medit.Enabled = True
MOption.Enabled = True
MHelpindex.Enabled = True
For i = 0 To 19
List1.List(i) = "Speaker" & (i + 1)
Next i
For i = 1 To 16
Toolbar1.Buttons(i).Enabled = True
Next i
End If
readini
 For i = 0 To Screen.FontCount - 1
  Fontlist.AddItem Screen.Fonts(i)
 Next i
Fontlist.Text = Text1.Font
Fontsizebox.AddItem 8
Fontsizebox.AddItem 9
Fontsizebox.AddItem 10
Fontsizebox.AddItem 11
Fontsizebox.AddItem 12
Fontsizebox.AddItem 14
Fontsizebox.AddItem 16
Fontsizebox.AddItem 18
Fontsizebox.AddItem 20
Fontsizebox.AddItem 22
Fontsizebox.AddItem 24
Fontsizebox.AddItem 26
Fontsizebox.AddItem 28
Fontsizebox.AddItem 36
Fontsizebox.AddItem 48
Fontsizebox.AddItem 72
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Fontsizebox.Text = Text1.Font.Size
h = False
onloading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Text1.Text <> "" Then Text1.SaveFile "MUNAutoSave.rtf"
Unload FormUnload
End Sub

'--------------------------------�����������------------------------------
Private Sub Command1_Click()
Textadd.Text = ""
Frameadd.Visible = True
Textadd.SetFocus
End Sub

Private Sub Command4_Click()
applyadd
End Sub

Private Sub Command5_Click()
Frameadd.Visible = False
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub






Private Sub Frameadd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub



Private Sub Framedel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub


Private Sub Frameedit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

'----------------------------------------������--------------------------------
Private Sub MCheckUpdate_Click()
Shell "update.exe /checknow", vbNormalFocus
End Sub


Private Sub MInternet_Click()
Wait.Show
DoEvents
Runprog "start iexplore http://createnhanceteam.1.vg"
Unload Wait
End Sub

Private Sub MQuit_Click()
MExit_Click
End Sub

Private Sub Mreg_Click()
frmreg.Show
End Sub

Private Sub MSet_Click()
frmOptions.Show
End Sub

Private Sub Mshowmainwindows_Click()
Me.WindowState = 0 '����ظ���Normal״̬
Delay 2
Me.Visible = True '�������������ͼ��
cSysTray1.InTray = False '��������ɼ�
End Sub

Private Sub MUpdateSetting_Click()
Shell "update.exe /configure", vbNormalFocus
End Sub

Private Sub Startmeeting_Click()
Meettime.Enabled = True
End Sub


'---------------------------------�༭��������----------------------------
Private Sub Command2_Click()
If List1.ListCount > 0 Then
Frameedit.Visible = True
Textedit.SetFocus
For i = 0 To List1.ListCount - 1
If List1.Selected(i) Then
Textedit.Text = List1.List(i)
editing = i
Exit For
End If
Next i
End If
End Sub

Private Sub Command6_Click()
Frameedit.Visible = False
End Sub

Private Sub Command7_Click()
applyedit
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

Private Sub Speaklist_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = False
End Sub

Private Sub Textadd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then applyadd
End Sub

Private Sub Textedit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then applyedit
End Sub
'-----------------------------------ɾ����������---------------------------
Private Sub Command3_Click()
If List1.ListCount > 0 Then
i = 0
Framedel.Visible = True
Command8.SetFocus
While i < List1.ListCount - 1
If List1.Selected(i) Then
Exit Sub
End If
i = i + 1
Wend
End If
End Sub

Private Sub Command8_Click()
del
End Sub

Private Sub Command9_Click()
Framedel.Visible = False
End Sub

'-----------------------------ʱ�ӿؼ����������Ҽ��˵�---------------
Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MClock
End Sub

Private Sub label6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then PopupMenu MClock
End Sub
'---------------------------------���ô�ʱ�ӵ���ʾ-----------------
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then showgonetime
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then showgonetime
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then showtimeleft
End Sub


Private Sub timeleft_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then showtimeleft
End Sub

'------------------------------------��ʱ�����ư�ť--------------------------------
Private Sub Startclock_Click()
Start
End Sub

Private Sub stopclock_Click()
resetclock
End Sub

Private Sub pauseclock_Click()
pause
End Sub

Private Sub Timersetting_Click()
Clocksetting.Show
End Sub

'-----------------------------------�˵������------------------------
Private Sub MAbout_Click()
frmAbout.Show
End Sub

Private Sub MClockclear_Click()
resetclock
End Sub

Private Sub MClockpause_Click()
pause
End Sub

Private Sub MClockSetting_Click()
Clocksetting.Show
End Sub

Private Sub MClockstart_Click()
Start
End Sub
Private Sub MExit_Click()
Selected = MessageBox(Me.hwnd, "ȷʵҪ�˳���", "ģ�����Ϲ������¼��", vbOKCancel)
If Selected = 1 Then Unload Me
End Sub

Private Sub MOpen_Click()
openf
End Sub

Private Sub MNew_Click()
newf
End Sub

Private Sub MSave_Click()
savef
End Sub

'-----------------------------------------����Timer�ؼ�����----------------------------------


Private Sub Refreshform_Timer() 'ˢ��frm
If isclockstarted = True Then
Startclock.Enabled = False
pauseclock.Enabled = True
Else
Startclock.Enabled = True
pauseclock.Enabled = False
End If
If Bigclockshow = 0 Then
Label6.ForeColor = vbBlack
Else
Label6.ForeColor = vbRed
End If
frm.Refresh
End Sub

Private Sub clockcontrol_Timer() 'ˢ��ϵͳʱ��
systemtime.Caption = " " & Date & " " & Time
End Sub

Private Sub timer3_Timer() 'ˢ�¶�ʱ��
Timernum = Timernum + 1
updatetimer
End Sub



'--------------------------��ʱ����ع���---------------------------------

Public Sub Start() '������ʱ��
isclockstarted = True
If Bigclockshow = 0 Then
Label6.ForeColor = vbBlack
Label6.Caption = " " & formattime(Timernum)
Else
Label6.ForeColor = vbRed
Label6.Caption = "-" & formattime(Timesetting1)
End If
timeleft.Caption = "-" & formattime(t)
Label4.Caption = " " & "00 : 00 : 00"
timer3.Enabled = True
End Sub

Public Sub pause() '��ͣ��ʱ��
timer3.Enabled = False
isclockstarted = False
End Sub

Public Sub resetclock()  '���ü�ʱ��
isclockstarted = False
timer3.Enabled = False
Timernum = 0
If Bigclockshow = 0 Then
Label6.ForeColor = vbBlack
Label6.Caption = " " & "00 : 00 : 00"
Else
Label6.ForeColor = vbGreen
Label6.Caption = "-" & "00 : 00 : 00"
End If
timeleft.Caption = "N/A"
Label4.Caption = "N/A"
d = "00"
e = "00"
F = "00"
End Sub

Public Sub updatetimer()  '���¼�ʱ��
Label4.Caption = " " & formattime(Timernum)
t = Timesetting1 - Timernum
timeleft.Caption = "-" & formattime(t)
timer3.Enabled = True
If Bigclockshow = 0 Then
Label6.ForeColor = vbBlack
Label6.Caption = " " & formattime(Timernum)
Else
Label6.ForeColor = vbRed
Label6.Caption = "-" & formattime(t)
End If
frm.Refresh
If t <= 0 Then    'ʱ�䵽
MsgBox "ʱ�䵽", vbOKOnly, "��ʾ"
timer3.Enabled = False
resetclock
End If
End Sub

Public Sub showtimeleft() '�Ŵ���ʾʣ�µ�ʱ��
Bigclockshow = 1
Label6.Caption = "-" & formattime(t)
End Sub

Public Sub showgonetime() '�Ŵ���ʾ�Ѿ���ʱ��
Bigclockshow = 0
Label6.Caption = " " & formattime(Timernum)
End Sub


'--------------------------------------�б����ع���------------------------
Public Sub applyadd() 'ȷ�����ѡ����
If Textadd.Text <> "" Then
Frameadd.Visible = False
List1.AddItem (Textadd.Text)
command1.SetFocus
End If
End Sub

Public Sub applyedit() 'ȷ���޸�ѡ����
If Textedit.Text <> "" Then
List1.List(editing) = Textedit.Text
Frameedit.Visible = False
End If
End Sub

Public Sub del() 'ɾ���б���
List1.RemoveItem (i)
Framedel.Visible = False
End Sub


'------------------------------------ͨ�ú�������---------------------------
Public Function formattime(ByVal t As Integer) As String  '������ʽ��Ϊʱ���֣�����ַ���
Dim g, h, i As String
g = Int(t / 3600)
h = Int((t - 3600 * Val(g)) / 60)
i = Int(t - 3600 * Val(g) - 60 * Val(h))
If Val(g) < 10 Then g = "0" & g
If Val(h) < 10 Then h = "0" & h
If Val(i) < 10 Then i = "0" & i
formattime = g & " : " & h & " : " & i
End Function


'--------------------------Open ini File------------------------
Public Sub readini()
Common.iniFileName = "setting.ini"
StatusBar1.Panels(2).AutoSize = sbrSpring
StatusBar1.Panels(2).Text = "             "
End Sub


Private Sub Timersetting_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Frame3.Visible = True

End Sub


'---------------------------------------�����¼��------------------------------
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)         '������

If Button.Index = 1 Then   '�½�
newf
End If

If Button.Index = 2 Then   '��
openf
End If

If Button.Index = 3 Then   '����
savef
End If

If Button.Index = 4 Then   '��ӡ
MessageBox Me.hwnd, "������", "�����¼", vbOKOnly
End If

If Button.Index = 6 Then   '����
SendMessage Me.Text1.hwnd, &HC7, 0, 0
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = True
End If

If Button.Index = 7 Then   '�ظ�
SendMessage Me.Text1.hwnd, &HC7, 0, 0
Toolbar1.Buttons(6).Enabled = True
Toolbar1.Buttons(7).Enabled = False
End If

If Button.Index = 8 Then   '����
MessageBox Me.hwnd, "������", "�����¼", vbOKOnly
End If

If Button.Index = 10 Then   '����
Clipboard.SetText Text1.SelRTF
Text1.SelRTF = ""
End If

If Button.Index = 11 Then   '����
Clipboard.SetText Text1.SelRTF
End If

If Button.Index = 12 Then   'ճ��
Text1.SelText = ""
Text1.SelRTF = Clipboard.GetText
End If

If Button.Index = 14 Then   '�Ӵ�
Text1.SelBold = Not Text1.SelBold
End If

If Button.Index = 15 Then   '�»���
Text1.SelUnderline = Not Text1.SelUnderline
End If

If Button.Index = 16 Then   'б��
Text1.SelItalic = Not Text1.SelItalic
End If

End Sub

Private Sub Fontlist_Change()        '��������
Text1.SelFontName = Fontlist.Text
If Not onloading Then Text1.SetFocus
End Sub

Private Sub Fontlist_Click()
Fontlist_Change
End Sub

Private Sub Fontsizebox_Change()         '�����ֺ�
If Fontsizebox.Text <> "" And h = True Then
Text1.SelFontSize = Val(Fontsizebox.Text)
h = False
If Not onloading Then Text1.SetFocus
End If

End Sub

Private Sub Fontsizebox_Click()
If Fontsizebox.Text <> "" Then
Text1.SelFontSize = Val(Fontsizebox.Text)
h = False
End If
If Not onloading Then Text1.SetFocus
End Sub

Private Sub Fontsizebox_KeyPress(KeyAscii As Integer)      '�س���Text1�ػ���
If KeyAscii = 13 Then
h = True
Text1.SelFontSize = Val(Fontsizebox.Text)
If Not onloading Then Text1.SetFocus
End If


End Sub

Private Sub text1_SelChange()  '�����ֺŻ��Ե��б��
If Text1.SelFontName <> Null Then Fontlist.Text = Text1.SelFontName
If Text1.SelFontSize <> Null Then Fontsizebox.Text = Text1.SelFontSize
End Sub


Private Sub Text1_Change()
Toolbar1.Buttons(6).Enabled = True
Toolbar1.Buttons(7).Enabled = False
End Sub

Private Sub Fontset_Click()
Dialog1.CancelError = True '������
On Error GoTo Cancel
Dialog1.Flags = cdlCFEffects Or cdlCFBoth '�趨ѡ���Ҫ����
'��ʼ��������
Dialog1.FontName = IIf(Text1.SelFontName <> Null, Text1.SelFontName, "΢���ź�")
Dialog1.FontSize = IIf(Text1.SelFontSize <> Null, Text1.SelFontSize, 24)
Dialog1.Color = IIf(Text1.SelColor <> Null, Text1.SelColor, vbBlack)
Dialog1.FontBold = IIf(Text1.SelBold <> Null, Text1.SelBold, False)
Dialog1.FontItalic = IIf(Text1.SelItalic <> Null, Text1.SelItalic, False)
Dialog1.FontUnderline = IIf(Text1.SelUnderline <> Null, Text1.SelUnderline, False)
Dialog1.FontStrikethru = IIf(Text1.SelStrikeThru <> Null, Text1.SelStrikeThru, False)
If Dialog1.FontName = Null Then Dialog1.FontName = False
If Dialog1.FontSize = Null Then Dialog1.FontSize = 24
'If Dialog1.Color = Null Then Dialog1.Color = vbBlack
'If Dialog1.FontBold = Null Then Dialog1.FontBold = False
'If Dialog1.FontItalic = Null Then Dialog1.FontItalic = False
'If Dialog1.FontUnderline = Null Then Dialog1.FontUnderline = False
'If Dialog1.FontStrikethru = Null Then Dialog1.FontStrikethru = False
Dialog1.ShowFont
Text1.SelFontName = Dialog1.FontName '������һһ��Ӧ�Ĺ�ϵ
Text1.SelFontSize = Dialog1.FontSize
Text1.SelColor = Dialog1.Color
Text1.SelBold = Dialog1.FontBold
Text1.SelItalic = Dialog1.FontItalic
Text1.SelUnderline = Dialog1.FontUnderline
Text1.SelStrikeThru = Dialog1.FontStrikethru
Cancel:
End Sub

Public Sub savef()
save:
Dialog1.CancelError = True '������
On Error GoTo cancel1
Dialog1.Filter = "����֧�ֵ��ĵ�(*.rtf;*.txt;*.doc;*.docx)|*.rtf;*.txt;*.doc;*.docx|Microsoft Word�ĵ�|*.doc;*.docx|Rich Text Format(*.rtf)|*.rtf|�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*"
Dialog1.FilterIndex = 3
Dialog1.filename = ""
Dialog1.ShowSave
If Dialog1.FilterIndex = 1 Then
If LCase(Right(Dialog1.filename, 3)) = "rtf" Then
Dialog1.FilterIndex = 2
ElseIf LCase(Right(Dialog1.filename, 3)) = "txt" Then
Dialog1.FilterIndex = 3
Else: Dialog1.FilterIndex = 4
End If
End If
If Dialog1.FilterIndex = 3 Then
h = MessageBox(Me.hwnd, "��ѡ���˱���Ϊ�ı��ļ���ʽ�����е�ͼƬ��������ļ�����������ý�ȫ���������ȷ�ϣ�", "����", vbOKCancel)
If h = 2 Then GoTo save
End If
If Dialog1.FilterIndex = 4 Then
h = MessageBox(hwnd, "�������޷�ʶ����Ҫ����ĸ�ʽ�����Դ��ı���ʽ������ļ����ļ��е�ͼƬ�͸�ʽ�趨�Ȼ�ȫ����ʧ��������ȷ����������������ȡ�����Ա���RTF��ʽ���ĵ���������ʽ���ú�ͼƬ����Ϣ��", "����", vbOKCancel)
If h = 1 Then Dialog1.FilterIndex = 3
If h = 2 Then Dialog1.FilterIndex = 2
End If

Text1.SaveFile Dialog1.filename, Dialog1.FilterIndex - 2
cancel1:
End Sub



Public Sub openf()
Dim isdoc As Boolean
isdoc = False
Dialog1.CancelError = True '������
On Error GoTo Cancel
Dialog1.Filter = "����֧�ֵ��ĵ�(*.rtf;*.txt;*.doc;*.docx)|*.rtf;*.txt;*.doc;*.docx|Microsoft Word�ĵ�|*.doc;*.docx|Rich Text Format(*.rtf)|*.rtf|�ı��ļ�(*.txt)|*.txt|�����ļ�(*.*)|*.*"
Dialog1.FilterIndex = 3
Dialog1.filename = ""
Dialog1.ShowOpen
If LCase(Right(Dialog1.filename, 3)) = "doc" Or LCase(Right(Dialog1.filename, 4)) = "docx" Then
isdoc = True
Wait.Label1.Caption = "����ת���ļ���ʽ�����Ժ򡭡�"
Wait.Show
DoEvents
Runprog "Copy """ + Dialog1.filename + """ c:\MUNTemp.doc"
DoEvents
Runprog "doctotext\doctotext.exe c:\MUNTemp.doc>c:\MUNTemp.txt"
Unload Wait
Wait.Label1.Caption = "����ת���ļ����룬���Ժ򡭡�"
Wait.Show
DoEvents
Runprog "iconv\iconv.exe -f utf-8//IGNORE -t gb2312 c:\MUNTemp.txt > c:\MUNTGB.txt"
Dialog1.filename = "c:\MUNTGB.txt"
End If
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Wait.Label1.Caption = "�����ļ������Ժ򡭡�"
Wait.Show
DoEvents
Text1.LoadFile Dialog1.filename
If isdoc Then
Runprog "del c:\MUNTemp.doc"
Runprog "del c:\MUNTemp.txt"
Runprog "del c:\MUNTGB.txt"
End If
Unload Wait
Cancel:
End Sub

Public Sub newf()
h = MessageBox(hwnd, "����������������¼��������", "�����¼", vbOKCancel)
If h = 1 Then
Text1.Text = ""
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(7).Enabled = False
End If
End Sub


'---------------------------------------����ͼ��----------------------
Private Sub Form_Resize()
If Me.WindowState = 1 Then '�����Ϊ��С���򡪡�
cSysTray1.InTray = True '���ص�������
Me.Visible = False '�ó�����治�ɼ�
End If
End Sub
'�֏ͳ�����Ļ
Private Sub CsysTray1_MouseDown(Button As Integer, Id As Long)
If Button = 1 Then
Me.WindowState = 0 '����ظ���Normal״̬
Delay 2
Me.Visible = True '�������������ͼ��
cSysTray1.InTray = False '��������ɼ�
Else
PopupMenu MTray
End If
End Sub

