Attribute VB_Name = "Module1"
Option Explicit

'�ڴ��ڽṹ��Ϊָ���Ĵ���������Ϣ
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'��ָ�����ڵĽṹ��ȡ����Ϣ
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
'����ָ���Ľ���
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'��ϵͳע��һ��ָ�����ȼ�
Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'ȡ���ȼ����ͷ�ռ�õ���Դ
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long) As Long
'�������API������ע��ϵͳ���ȼ�������ģ�����ʵ�ֹ����������ʾ

  '�ȼ���־����,�����жϵ����̰���������ʱ�Ƿ������������趨���ȼ�
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

'����ϵͳ���ȼ�,ԭ�жϱ�ʾ,�����ص���Ŀ���
Public preWinProc As Long, MyhWnd As Long, uVirtKey As Long

'�ȼ����ع���
Public Function WndProc(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then     '������ص��ȼ���־����
        If wParam = 1 Then      '��������ǵĶ�����ȼ�...
            HideDone            'ִ�����������ָ��Ŀ
        End If
      End If
    '��������ȼ�,���߲����������õ��ȼ�,��������Ȩ��ϵͳ,��������ȼ�
    WndProc = CallWindowProc(preWinProc, Hwnd, Msg, wParam, lParam)
End Function

'��ؼ�����Ŀ���ع���
Public Sub HideDone()
Open "time1" For Output As #16
Print #16, "0"
Close #16
End Sub
