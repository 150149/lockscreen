Attribute VB_Name = "Module1"
Option Explicit

'在窗口结构中为指定的窗口设置信息
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'从指定窗口的结构中取得信息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
'运行指定的进程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'向系统注册一个指定的热键
Public Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'取消热键并释放占用的资源
Public Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal ID As Long) As Long
'上述五个API函数是注册系统级热键所必需的，具体实现过程如后文所示

  '热键标志常数,用来判断当键盘按键被按下时是否命中了我们设定的热键
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

'定义系统的热键,原中断标示,被隐藏的项目句柄
Public preWinProc As Long, MyhWnd As Long, uVirtKey As Long

'热键拦截过程
Public Function WndProc(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then     '如果拦截到热键标志常数
        If wParam = 1 Then      '如果是我们的定义的热键...
            HideDone            '执行隐藏鼠标所指项目
        End If
      End If
    '如果不是热键,或者不是我们设置的热键,交还控制权给系统,继续监测热键
    WndProc = CallWindowProc(preWinProc, Hwnd, Msg, wParam, lParam)
End Function

'最关键的项目隐藏过程
Public Sub HideDone()
Open "time1" For Output As #16
Print #16, "0"
Close #16
End Sub
