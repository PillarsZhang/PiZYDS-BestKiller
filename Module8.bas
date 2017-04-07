Attribute VB_Name = "Others"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' SetWindowPos Flags
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering

Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

' SetWindowPos() hwndInsertAfter values
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Public rfm As RECT
Public hfm, hfs As Long
Public Sub letsengo()
Frm_Senior.Height = Frm_Main.Height
GetWindowRect hfm, rfm
'MsgBox "左上角坐标(" & rfm.Left & "," & rfm.Top & ")" & vbCrLf & "右下角坐标(" & rfm.Right & "," & rfm.Bottom & ")" & vbCrLf & "窗口高" & rfm.Bottom - rfm.Top & "窗口宽" & rfm.Right - rfm.Left
SetWindowPos hfs, -1, rfm.Right, rfm.Top, 0, 0, 1
End Sub

'https://zhidao.baidu.com/question/75611930.html


'Dim h As Long, r As RECT
'h = FindWindow(vbNullString, "酷狗") '这里写上你的窗口标题，必须一字不差
'GetWindowRect h, r
