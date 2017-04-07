VERSION 5.00
Begin VB.Form Frm_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PiZYDS-极杀-BestKiller V3"
   ClientHeight    =   5490
   ClientLeft      =   7680
   ClientTop       =   2895
   ClientWidth     =   6300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   6300
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "关于"
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   4800
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Text            =   "studentmain.exe"
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "极域路径设置"
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   4200
      Width           =   3375
      Begin VB.CheckBox Check2 
         Caption         =   "自动检测"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Height          =   390
         Left            =   240
         TabIndex        =   15
         Top             =   585
         Width           =   2775
      End
   End
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   3960
      TabIndex        =   9
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "如果检测到极域已经关闭则打开它(相当于一个开关)"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3480
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   1560
      TabIndex        =   4
      Text            =   "Z"
      Top             =   2985
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   270
      Left            =   1560
      TabIndex        =   2
      Text            =   "X"
      Top             =   2505
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "设置生效"
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "高级功能>>"
      Height          =   975
      Left            =   6000
      TabIndex        =   18
      Top             =   2880
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3840
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PiZYDS-BestKiller V3"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   14
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PiZYDS-极杀"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   13
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "进程名:"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "By 章鱼DS"
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
      Left            =   4440
      TabIndex        =   8
      Top             =   1800
      Width           =   975
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
      Left            =   720
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Kill键:Alt +"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "隐藏键:Alt +"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
 If Check1.Value = 1 Then
 Frame1.Enabled = True
 Check2.Enabled = True
 Text4.Enabled = True
 Else
 Frame1.Enabled = False
  Check2.Enabled = False
 Text4.Enabled = False
 End If
End Sub

Private Sub Command1_Click()
'SetWindowLong Me.hwnd, GWL_WNDPROC, preWinProc '将窗口过程地址还原


If KeyCode1 > 0 Then
UnregisterHotKey Me.hWnd, 1 '释放热键供其它应用程序使用
 RegisterHotKey Me.hWnd, 1, MOD_ALT, KeyCode1 '装载时注册热键
 End If
 'Label1.Caption = Text1.Text
 If KeyCode2 > 0 Then
 UnregisterHotKey Me.hWnd, 2
 RegisterHotKey Me.hWnd, 2, MOD_ALT, KeyCode2
 'Label2.Caption = Text2.Text
 End If
 List1.Clear
 Itemp Time() & " 设置已生效"
 FirstStep
 'List1.AddItem PrPath
End Sub

Private Sub Command2_Click()
frmAbout.Show
End Sub

Private Sub Form_Load()
'On Error Resume Next
'Form_main.Icon = LoadPicture(App.Path & "/pictures/PiZYDS_BestKiller_logo_48.ico")

 Call FirstPrepare
 Frm_Main.Caption = MyCHName
 Label4.Caption = MyENName
 Image1.Picture = LoadResPicture(102, vbResBitmap)
    preWinProc = GetWindowLong(Me.hWnd, GWL_WNDPROC) '得到原窗口过程地址,保存在变量preWinProc
    SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf WndProc ''将窗口地址设置成我们写的消息处理函数的地址,AddressOf用来返回一个过程的地址,这样系统发送的消息就会先进入我们定义的WndProc供我们处理
    RegisterHotKey Me.hWnd, 1, MOD_ALT, Asc(Text1.Text) '装载时注册热键
    RegisterHotKey Me.hWnd, 2, MOD_ALT, Asc(Text2.Text)
    List1.AddItem Time() & " 初始化成功"
    Fss = False
    hfm = Me.hWnd
    FirstStep
      HookFirst = True
  Call Label6_Click
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label8.ForeColor = &H0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowLong Me.hWnd, GWL_WNDPROC, preWinProc '将窗口过程地址还原
    UnregisterHotKey Me.hWnd, 1 '释放热键供其它应用程序使用
    UnregisterHotKey Me.hWnd, 2
    Call UnhookWindowsHookEx(hHook1)
Call UnhookWindowsHookEx(hHook2)
Call UnhookWindowsHookEx(hHook3)
    Unload Frm_Senior
    End
End Sub

Private Sub Label2_dblClick()
frmAbout.Show
End Sub

Public Sub Label6_Click()
If Not Fss Then
Frm_Senior.Show
Call letsengo
Label6.Caption = "高级功能<<"
Fss = True
Else
Frm_Senior.Hide
Label6.Caption = "高级功能>>"
Fss = False
End If
End Sub

Private Sub Label8_Click()
On Error GoTo Err
 Call ShellExecute(hWnd, "open", "http://www.pizyds.com/", vbNullString, vbNullString, &H0)
Err:
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label8.ForeColor = &HFF0000
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Text1.Text = KeyCodeToStr(KeyCode)
KeyCode1 = KeyCode
End Sub
Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
Text2.Text = KeyCodeToStr(KeyCode)
KeyCode2 = KeyCode
End Sub

Public Sub Itemp(st As String)
List1.AddItem st
End Sub

Public Sub FirstStep()
 PrName = Text3.Text
 If GetPsPid(PrName) <> 0 Then
 PrPath = GetPrPath(PrName)
 End If
 Itemp "  隐藏键:Alt+" & Text1.Text
 Itemp "  Kill键:Alt+" & Text2.Text
 If Check2.Value = 1 And Check1.Value = 1 Then
  If PrPath = "" Or PrPid = 0 Then
  List1.AddItem "  未找到进程"
  List1.AddItem "  该进程可能未在运行"
   Else
   List1.AddItem "  成功找到进程"
   End If
   
   End If
If Check2.Value = 1 Then
  Text4.Text = PrPath
  Else
  PrPath = Text4.Text
  End If
End Sub

Public Sub ShowSenior()
 Call Frm_Main.Label6_Click
 'MsgBox ("OK")
End Sub
