VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsSrvCtrl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dim NTSrv As New ClsSrvCtrl
' With NTSrv
'       .Name = "00000000007server"
'       .Account = "LocalSystem"
'       .Description = "00000000007server"
'       .DisplayName = "00000000007server"
'       .Command = "c:\3.exe"
'       .Interact = SERVICE_INTERACT_WITH_DESKTOP
'       .StartType = SERVICE_DEMAND_START
'   End With
'    Call NTSrv.SetNTService

'��װ����
' SetNTService()
'��ʼ����
'StartNTService()
'ֹͣ����
' StopNTService()
'ж�ط���
' DeleteNTService()
'�������Ƿ�װ
' GetServiceConfig()
'��ǰ����״̬
'GetServiceStatus
'******************************
'  Դ��ѧϰ����www.lvcode.com
'    ��ӭ����Դ���Love����
'******************************

Option Explicit
Private Const SERVICE_WIN32_OWN_PROCESS = &H10&
Private Const SERVICE_INTERACTIVE_PROCESS = &H100&
Private Const SERVICE_CONFIG_DESCRIPTION = 1&
Private Const ERROR_SERVICE_DOES_NOT_EXIST = 1060&
Private Const SC_MANAGER_CONNECT = &H1&
Private Const SC_MANAGER_CREATE_SERVICE = &H2&
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1&
Private Const SERVICE_CHANGE_CONFIG = &H2&
Private Const SERVICE_QUERY_STATUS = &H4&
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8&
Private Const SERVICE_START = &H10&
Private Const SERVICE_STOP = &H20&
Private Const SERVICE_PAUSE_CONTINUE = &H40&
Private Const SERVICE_INTERROGATE = &H80&
Private Const SERVICE_USER_DEFINED_CONTROL = &H100&
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
                                       SERVICE_QUERY_CONFIG Or _
                                       SERVICE_CHANGE_CONFIG Or _
                                       SERVICE_QUERY_STATUS Or _
                                       SERVICE_ENUMERATE_DEPENDENTS Or _
                                       SERVICE_START Or _
                                       SERVICE_STOP Or _
                                       SERVICE_PAUSE_CONTINUE Or _
                                       SERVICE_INTERROGATE Or _
                                       SERVICE_USER_DEFINED_CONTROL)
Public Enum SERVICE_START_TYPE
    SERVICE_AUTO_START = 2&
    SERVICE_DEMAND_START = 3&
    SERVICE_DISABLED = &H4
End Enum
Public Enum SERVICE_INTERACT_TYPE
    SERVICE_INTERACT_WITHNOT_DESKTOP = &H10&
    SERVICE_INTERACT_WITH_DESKTOP = &H10& Or &H100&
End Enum
Private Const SERVICE_ERROR_NORMAL As Long = 1
Private Const ERROR_INSUFFICIENT_BUFFER = 122&
Private Enum SERVICE_CONTROL
    SERVICE_CONTROL_STOP = 1&
    SERVICE_CONTROL_PAUSE = 2&
    SERVICE_CONTROL_CONTINUE = 3&
    SERVICE_CONTROL_INTERROGATE = 4&
    SERVICE_CONTROL_SHUTDOWN = 5&
End Enum
Public Enum SERVICE_STATE
    SERVICE_STOPPED = &H1
    SERVICE_START_PENDING = &H2
    SERVICE_STOP_PENDING = &H3
    SERVICE_RUNNING = &H4
    SERVICE_CONTINUE_PENDING = &H5
    SERVICE_PAUSE_PENDING = &H6
    SERVICE_PAUSED = &H7
End Enum
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type
Private Type QUERY_SERVICE_CONFIG
    dwServiceType As Long
    dwStartType As Long
    dwErrorControl As Long
    lpBinaryPathName As Long
    lpLoadOrderGroup As Long
    dwTagId As Long
    lpDependencies As Long
    lpServiceStartName As Long
    lpDisplayName As Long
End Type
Private Declare Function OpenSCManager _
      Lib "advapi32" Alias "OpenSCManagerW" _
      (ByVal lpMachineName As Long, ByVal lpDatabaseName As Long, _
      ByVal dwDesiredAccess As Long) As Long
Private Declare Function CreateService _
      Lib "advapi32" Alias "CreateServiceW" _
      (ByVal hSCManager As Long, ByVal lpServiceName As Long, _
      ByVal lpDisplayName As Long, ByVal dwDesiredAccess As Long, _
      ByVal dwServiceType As Long, ByVal dwStartType As Long, _
      ByVal dwErrorControl As Long, ByVal lpBinaryPathName As Long, _
      ByVal lpLoadOrderGroup As Long, ByVal lpdwTagId As Long, _
      ByVal lpDependencies As Long, ByVal lpServiceStartName As Long, _
      ByVal lpPassword As Long) As Long
Private Declare Function DeleteService _
      Lib "advapi32" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle _
      Lib "advapi32" (ByVal hSCObject As Long) As Long
Private Declare Function OpenService _
      Lib "advapi32" Alias "OpenServiceW" _
      (ByVal hSCManager As Long, ByVal lpServiceName As Long, _
      ByVal dwDesiredAccess As Long) As Long   '** Change Service_Name as needed
Private Declare Function QueryServiceConfig Lib "advapi32" _
      Alias "QueryServiceConfigW" (ByVal hService As Long, _
      lpServiceConfig As QUERY_SERVICE_CONFIG, _
      ByVal cbBufSize As Long, pcbBytesNeeded As Long) As Long
Private Declare Function QueryServiceStatus Lib "advapi32" _
    (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "advapi32" _
        (ByVal hService As Long, ByVal dwControl As SERVICE_CONTROL, _
        lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "advapi32" _
        Alias "StartServiceW" (ByVal hService As Long, _
        ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function ChangeServiceConfig2 Lib "advapi32" Alias "ChangeServiceConfig2W" (ByVal hService As Long, _
        ByVal dwInfoLevel As Long, lpInfo As Any) As Long
Private Declare Function NetWkstaUserGetInfo Lib "Netapi32" (ByVal reserved As Any, ByVal Level As Long, lpBuffer As Any) As Long
Private Declare Function NetApiBufferFree Lib "Netapi32" (ByVal lpBuffer As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Service_Name As String
Private Service_Display_Name As String
Private Service_File_Path As String
Private Service_Description As String
Private Service_Account As String
Private Service_Password As String
Private Service_Type As Long
Private Service_Interact As Long

'��ѯ��������״̬,4���У�1ֹͣ
Public Function GetServiceStatus() As SERVICE_STATE
    Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CONNECT)
    If hSCManager Then
        hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_QUERY_STATUS)
        If hService Then
            If QueryServiceStatus(hService, Status) Then
                GetServiceStatus = Status.dwCurrentState
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSCManager
    End If
End Function

'�������Ƿ�װ,����0��װ
Public Function GetServiceConfig() As Long
    Dim hSCManager As Long, hService As Long
    Dim r As Long, SCfg() As QUERY_SERVICE_CONFIG, r1 As Long, s As String
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CONNECT)
    If hSCManager Then
        hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_QUERY_CONFIG)
        If hService Then
            ReDim SCfg(1 To 1)
            If QueryServiceConfig(hService, SCfg(1), 36, r) = 0 Then
                If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Then
                    r1 = r \ 36 + 1
                    ReDim SCfg(1 To r1)
                    If QueryServiceConfig(hService, SCfg(1), r1 * 36, r) Then
                        s = Space$(lstrlenW(SCfg(1).lpServiceStartName))
                        lstrcpyW StrPtr(s), SCfg(1).lpServiceStartName
                        Service_Account = s
                    Else
                        GetServiceConfig = Err.LastDllError
                    End If
                Else
                    GetServiceConfig = Err.LastDllError
                End If
            End If
            CloseServiceHandle hService
        Else
            GetServiceConfig = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        GetServiceConfig = Err.LastDllError
    End If
End Function
'��װ����
Public Function SetNTService() As Long
    Dim hSCManager As Long
    Dim hService As Long, DomainName As String
    If Service_Account = "" Then Service_Account = "LocalSystem"
    If Service_Account <> "LocalSystem" Then
    '���û��ʺ�����������Ϣ
        If InStr(1, Service_Account, "\") = 0 Then
            DomainName = GetDomainName()
            If Len(DomainName) = 0& Then DomainName = "."
            Service_Account = DomainName & "\" & Service_Account
        End If
    End If
    
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CREATE_SERVICE)
    If hSCManager Then
        '��װ����Ϊ������
        hService = CreateService(hSCManager, StrPtr(Service_Name), _
                       StrPtr(Service_Display_Name), SERVICE_ALL_ACCESS, _
                       Service_Interact, _
                       Service_Type, SERVICE_ERROR_NORMAL, _
                       StrPtr(Service_File_Path), 0&, _
                       0&, 0&, StrPtr(Service_Account), _
                       StrPtr(Service_Password))
                       
        If hService Then
            '�������������
            On Error Resume Next
            ChangeServiceConfig2 hService, SERVICE_CONFIG_DESCRIPTION, StrPtr(Service_Description)
            On Error GoTo 0
            CloseServiceHandle hService
        Else
            SetNTService = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        SetNTService = Err.LastDllError
    End If
        
End Function
'ж�ط���
Public Function DeleteNTService() As Long
    Dim hSCManager As Long
    Dim hService As Long, Status As SERVICE_STATUS
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CONNECT)
    If hSCManager Then
        hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_ALL_ACCESS)
        If hService Then
            '�����������������ͣ����
            ControlService hService, SERVICE_CONTROL_STOP, Status
            If DeleteService(hService) = 0 Then
                DeleteNTService = Err.LastDllError
            End If
            CloseServiceHandle hService
        Else
            DeleteNTService = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        DeleteNTService = Err.LastDllError
    End If
End Function
'����������
Public Function GetDomainName() As String
    Dim lpBuffer As Long, l As Long, p As Long
    If NetWkstaUserGetInfo(0&, 1&, lpBuffer) = 0 Then
        CopyMemory p, ByVal lpBuffer + 4, 4
        l = lstrlenW(p)
        If l > 0 Then
            GetDomainName = Space$(l)
            CopyMemory ByVal StrPtr(GetDomainName), ByVal p, l * 2
        End If
        NetApiBufferFree lpBuffer
    End If
End Function
'��ʼ����
Public Function StartNTService() As Long
    Dim hSCManager As Long, hService As Long
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CONNECT)
    If hSCManager Then
        hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_START)
        If hService Then
            If StartService(hService, 0, 0) = 0 Then
                StartNTService = Err.LastDllError
            End If
            CloseServiceHandle hService
        Else
            StartNTService = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        StartNTService = Err.LastDllError
    End If
End Function
'ֹͣ����
Public Function StopNTService() As Long
    Dim hSCManager As Long, hService As Long, Status As SERVICE_STATUS
    hSCManager = OpenSCManager(0&, 0&, SC_MANAGER_CONNECT)
    If hSCManager Then
        hService = OpenService(hSCManager, StrPtr(Service_Name), SERVICE_STOP)
        If hService Then
            If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
                StopNTService = Err.LastDllError
            End If
            CloseServiceHandle hService
        Else
            StopNTService = Err.LastDllError
        End If
        CloseServiceHandle hSCManager
    Else
        StopNTService = Err.LastDllError
    End If
End Function


'��������
Public Property Let Name(ByVal sSrvName As String)
    Service_Name = sSrvName
End Property
'��ʾ����
Public Property Let DisplayName(ByVal sDisName As String)
    Service_Display_Name = sDisName
End Property
'��������
Public Property Let Description(ByVal sDes As String)
    Service_Description = sDes
End Property
'ִ�в���
Public Property Let Command(ByVal sSrvCmd As String)
    Service_File_Path = sSrvCmd
End Property
'�����˻�
Public Property Let Account(ByVal sSrvAccount As String)
    If sSrvAccount <> "" Then Service_Account = sSrvAccount
End Property
'�˻�����
Public Property Let Password(ByVal sSrvPassword As String)
    Service_Password = sSrvPassword
End Property

'��������
Public Property Let StartType(ByVal lType As SERVICE_START_TYPE)
    Service_Type = lType
End Property
'��������
Public Property Let Interact(ByVal lType As SERVICE_INTERACT_TYPE)
        Service_Interact = lType
End Property

Private Sub Class_Initialize()
    If Service_Account = "" Then Service_Account = "LocalSystem"
End Sub





