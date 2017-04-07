Attribute VB_Name = "HotKey"
'变量
Public preWinProc As Long '存储原本的窗口过程的地址
  
'常量
Public Const GWL_WNDPROC = (-4) '这个常数供GetWindowLong和SetWindowLong使用以得到和设置窗口过程地址
Public Const WM_HOTKEY = &H312 '热键消息常数,用来判断消息是否为热键消息的常数
Public Const MOD_ALT = &H1 'RegisterHotKey和UnregisterHotKey用到的表示按下Alt键的常数
  
  
'API声明
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long, ByVal _
                        dwNewLong As Long) As Long
  
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal _
                        hWnd As Long, ByVal _
                        nIndex As Long) As Long
  
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal _
                        lpPrevWndFunc As Long, ByVal _
                        hWnd As Long, ByVal _
                        Msg As Long, ByVal _
                        wParam As Long, ByVal _
                        lParam As Long) As Long
  
Public Declare Function RegisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long, ByVal _
                        fsModifiers As Long, ByVal _
                        vk As Long) As Long '向系统注册热键
  
Public Declare Function UnregisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long) As Long
  
'过程
Sub Main()
    If App.PrevInstance = True Then    '如果如果已经运行就自己退出
        MsgBox "程序已经运行!", vbOKOnly, "提示"
        End
    End If
    Frm_Main.Show
End Sub
  
  
Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then '如果是热键消息
        If wParam = 1 Then '如果是本程序定义的(系统消息中的wParam参数在热键消息中代表热键标示符,是在RegisterHotKey注册热键的时候定义的一个整数,如果热键是系统定义的,则标示符取值为-1或-2,详见开头
                Call WindowShowHide ' 热键对应上了之后就调用指定的过程
                ElseIf wParam = 2 Then
                Call bestkill
                Exit Function '消息已处理,不需要发回窗口
        End If
    End If
    WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam) '不是热键消息,就把消息发给原来窗口过程交给它处理
End Function
