Attribute VB_Name = "HotKey"
'����
Public preWinProc As Long '�洢ԭ���Ĵ��ڹ��̵ĵ�ַ
  
'����
Public Const GWL_WNDPROC = (-4) '���������GetWindowLong��SetWindowLongʹ���Եõ������ô��ڹ��̵�ַ
Public Const WM_HOTKEY = &H312 '�ȼ���Ϣ����,�����ж���Ϣ�Ƿ�Ϊ�ȼ���Ϣ�ĳ���
Public Const MOD_ALT = &H1 'RegisterHotKey��UnregisterHotKey�õ��ı�ʾ����Alt���ĳ���
  
  
'API����
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
                        vk As Long) As Long '��ϵͳע���ȼ�
  
Public Declare Function UnregisterHotKey Lib "user32" (ByVal _
                        hWnd As Long, ByVal _
                        ID As Long) As Long
  
'����
Sub Main()
    If App.PrevInstance = True Then    '�������Ѿ����о��Լ��˳�
        MsgBox "�����Ѿ�����!", vbOKOnly, "��ʾ"
        End
    End If
    Frm_Main.Show
End Sub
  
  
Public Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then '������ȼ���Ϣ
        If wParam = 1 Then '����Ǳ��������(ϵͳ��Ϣ�е�wParam�������ȼ���Ϣ�д����ȼ���ʾ��,����RegisterHotKeyע���ȼ���ʱ�����һ������,����ȼ���ϵͳ�����,���ʾ��ȡֵΪ-1��-2,�����ͷ
                Call WindowShowHide ' �ȼ���Ӧ����֮��͵���ָ���Ĺ���
                ElseIf wParam = 2 Then
                Call bestkill
                Exit Function '��Ϣ�Ѵ���,����Ҫ���ش���
        End If
    End If
    WndProc = CallWindowProc(preWinProc, hWnd, Msg, wParam, lParam) '�����ȼ���Ϣ,�Ͱ���Ϣ����ԭ�����ڹ��̽���������
End Function
