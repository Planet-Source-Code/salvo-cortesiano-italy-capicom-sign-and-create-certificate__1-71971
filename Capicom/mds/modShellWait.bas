Attribute VB_Name = "modShellWait"
'Simulate multithreading with WaitForMultipleObjects
'(eg. How ICQ monitors connection state)
'By: John Galanopoulos
Option Explicit

Public Const WAIT_FAILED = &HFFFFFFFF       'Our WaitForSingleObject failed to wait and returned -1
Public Const WAIT_OBJECT_0 = &H0&           'The waitable object got signaled '
Public Const WAIT_ABANDONED = &H80&         'We got out of the waitable object
Public Const WAIT_TIMEOUT = &H102&          'the interval we used, timed out.
Public Const STANDARD_RIGHTS_ALL = &H1F0000 'No special user rights needed to open t ' his process

Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function RasConnectionNotification Lib "rasapi32.dll" Alias "RasConnectionNotificationA" (hRasConn As Long, ByVal hEvent As Long, ByVal dwFlags As Long) As Long

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Const RASCN_Connection = &H1
Public Const RASCN_Disconnection = &H2

Public Enum sShell
    vbHide = 0
    vbMaximizedFocus = 1
    vbMinimizedFocus = 2
    vbMinimizedNoFocus = 3
    vbNormalFocus = 4
    vbNormalNoFocus = 5
End Enum
Public Function ShelledAPP(ByVal FileName As String, Optional runMode As sShell = vbNormalFocus) As String
    Dim shProcID As Long
    Dim hProcess As Long
    Dim WaitRet As Long
    On Local Error GoTo ErrorShelled
    shProcID = Shell(FileName, runMode)
    hProcess = OpenProcess(STANDARD_RIGHTS_ALL, False, shProcID)
    Do
    WaitRet = WaitForSingleObject(hProcess, 10)      ' wait for 10ms to see if the hProcess was signaled
        Select Case WaitRet
                Case WAIT_TIMEOUT               'The first case must always be WAIT_TIMEOUT 'cause it is the most used option
                    ShelledAPP = "Wait"
                    DoEvents              'until the shelled process terminates
                Case WAIT_FAILED Or WAIT_ABANDONED
                    ShelledAPP = "Error"
            Exit Do
                Case WAIT_OBJECT_0              'The object got signaled so inform user and get out of the loop
                    ShelledAPP = "End"
            Exit Do
        End Select
    Loop
    Call CloseHandle(hProcess)                      'Close the process handle
    Call CloseHandle(shProcID)                      'Close the process id handle
    DoEvents
ErrorShelled:
    Err.Clear
End Function

Public Sub MonitorRASStatusAsync()
    Dim hEvents(1) As Long
    Dim RasNotif As Long
    Dim WaitRet As Long
    Dim sd As SECURITY_ATTRIBUTES
    Dim hRasConn As Long

    hRasConn = 0

    With sd
       .nLength = Len(sd)
       .lpSecurityDescriptor = 0
       .bInheritHandle = 0
    End With
    
    hEvents(0) = CreateEvent(sd, True, False, "RASStatusNotificationObject1")
    If hEvents(0) = 0 Then MsgBox "Couldn't assign an event handle": Exit Sub
    RasNotif = RasConnectionNotification(ByVal hRasConn, hEvents(0), RASCN_Connection)
    If RasNotif <> 0 Then MsgBox "Ras Notification failure": GoTo ras_TerminateEvent
    hEvents(1) = CreateEvent(sd, True, False, "RASStatusNotificationObject2")
    If hEvents(1) = 0 Then MsgBox "Couldn't assign an event handle": Exit Sub
    RasNotif = RasConnectionNotification(ByVal hRasConn, hEvents(1), RASCN_Disconnection)
    If RasNotif <> 0 Then MsgBox "Ras Notification failure": GoTo ras_TerminateEvent
    
    Do
       WaitRet = WaitForMultipleObjects(2, hEvents(0), False, 20)
                       Select Case WaitRet
                Case WAIT_TIMEOUT
                    DoEvents
                Case WAIT_FAILED Or WAIT_ABANDONED Or WAIT_ABANDONED + 1
                    GoTo ras_TerminateEvent
                Case WAIT_OBJECT_0
                    MsgBox "Connected"
                ResetEvent hEvents(0) 'Reset the event to avoid a second message box
            DoEvents    'Free any pending messages
                Case WAIT_OBJECT_0 + 1
                    MsgBox "Disconnected"
                 ResetEvent hEvents(1) 'Reset the event to place it in no signal state (Manual reset, remember?)
            DoEvents
        End Select
    Loop
ras_TerminateEvent:
    Call CloseHandle(hEvents(1))
    Call CloseHandle(hEvents(0))
DoEvents
Exit Sub
End Sub
