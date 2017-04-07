VERSION 5.00
Begin VB.Form Frm_Senior 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¸ß¼¶ÉèÖÃ"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2250
   ControlBox      =   0   'False
   Icon            =   "Frm_Senior.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   2250
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "Ñ­»·ÇÀ¹³"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.ListBox List1 
         Height          =   960
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ÆÁ±Î°´¼ü²âÊÔ"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Text            =   "500"
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "¿ªÊ¼ÇÀ¹³"
         Height          =   495
         Left            =   600
         TabIndex        =   1
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   0
         Top             =   2040
      End
      Begin VB.Label Label4 
         Caption         =   "×´Ì¬£º"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "ÑÓ³Ù(ºÁÃë)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Caption         =   "More..."
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
End
Attribute VB_Name = "Frm_Senior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Testhook As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
Testhook = True
 Else
 Testhook = False
End If
End Sub

Private Sub Command1_Click()
If Timer1.Enabled = False Then
Timer1.Interval = Val(Text1.Text)
Timer1.Enabled = True
Command1.Caption = "Í£Ö¹ÇÀ¹³"
List1.Clear
List1.AddItem "Stop.."
Else
Timer1.Enabled = False
Call UnhookWindowsHookEx(hHook1)
Call UnhookWindowsHookEx(hHook2)
Call UnhookWindowsHookEx(hHook3)
Command1.Caption = "¿ªÊ¼ÇÀ¹³"
'Label3.Caption = "Stop.."
End If
End Sub

Private Sub Form_Load()
Testhook = False
hfs = Me.hWnd
Call letsengo
Hooknum = 13
List1.Clear
List1.AddItem "Stop.."
If HookFirst = True Then
 Call Command1_Click
 HookFirst = False
 End If
End Sub


Private Sub Timer1_Timer()
Call GetHook
List1.Clear
List1.AddItem Hookst(0)
List1.AddItem Hookst(1)
List1.AddItem Hookst(2)
List1.AddItem Hookst(3)
End Sub
