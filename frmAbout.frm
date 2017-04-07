VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 PiZYDS-极杀-BestKiller V3"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Text            =   "frmAbout.frx":0ECA
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "访问官网"
      Height          =   345
      Left            =   4140
      TabIndex        =   1
      Top             =   3075
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   795
   End
   Begin VB.Image Image1 
      Height          =   465
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "开发者：章鱼DS"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "官网地址：http://www.pizyds.com/"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblTitle 
      Caption         =   "PiZYDS-极杀-BestKiller"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本 V3"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "作者信息："
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   255
      TabIndex        =   2
      Top             =   2625
      Width           =   990
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
 Unload frmAbout
End Sub

Private Sub cmdSysInfo_Click()
Call ShellExecute(hWnd, "open", "http://www.pizyds.com/", vbNullString, vbNullString, &H0)
End Sub

Private Sub Form_Load()
 On Error Resume Next
 frmAbout.Caption = "关于 " & MyCHName
 lblVersion.Caption = "版本 " & MyVision
 Image2.Picture = LoadResPicture(101, vbResBitmap)
 Image1.Picture = LoadResPicture(102, vbResBitmap)
End Sub

Private Sub Label8_Click()
On Error GoTo Err
 Call ShellExecute(hWnd, "open", "http://www.pizyds.com/", vbNullString, vbNullString, &H0)
Err:
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label8.ForeColor = &HFF0000
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label8.ForeColor = &H0&
End Sub

