Attribute VB_Name = "GetHooking"
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" _
             (ByVal idHook As Long, _
               ByVal lpfn As Long, _
               ByVal hmod As Long, _
               ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" _
               (ByVal hHook As Long, _
               ByVal ncode As Long, _
               ByVal wParam As Long, _
               lParam As Any) As Long
               
Declare Function UnhookWindowsHookEx Lib "user32" _
               (ByVal hHook As Long) As Long
               
Public hHook1, hHook2, hHook3 As Long
Public Hooknum As Integer
Public Hookst(3) As String
Public Sub GetHook()
Call UnhookWindowsHookEx(hHook1)
hHook1 = SetWindowsHookEx(Hooknum, AddressOf MyKBHook, App.hInstance, 0)
Call UnhookWindowsHookEx(hHook2)
hHook2 = SetWindowsHookEx(Hooknum, AddressOf MyKBHook, App.hInstance, 0)
Call UnhookWindowsHookEx(hHook3)
hHook3 = SetWindowsHookEx(Hooknum, AddressOf MyKBHook, App.hInstance, 0)

Hookst(0) = "Succeed.."
Hookst(1) = Str(hHook1)
Hookst(2) = Str(hHook2)
Hookst(3) = Str(hHook3)
'MsgBox Hookst
End Sub
'具体的钩子程序,本例中该过程被包含在Module1中
Public Function MyKBHook(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
MyHBHook = 0
If Testhook Then MyKBHook = 1
'Call CallNextHookEx(hHook, ncode, wParam, lParam) '将消息传给下一个钩子
End Function

'http://www.cnblogs.com/ywb-lv/articles/2443868.html
