Attribute VB_Name = "ProcessHelper"
'--------------------------------------------------------------------------------
'    Component  : ProcessHelper
'    Project    : prjSuperBar
'
'    Description: Process utility module. Provides native process helper
'                 functions
'
'--------------------------------------------------------------------------------
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetModuleFileNameExW Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcessImageFileName Lib "PSAPI" Alias "GetProcessImageFileNameW" (ByVal hProcess As Long, ByVal lptrImageFileName As Long, ByVal nSize As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsWow64Process Lib "kernel32" _
    (ByVal hProc As Long, _
    bWow64Process As Boolean) As Long


Public Function GetProcessID(hWnd As Long) As Long
Dim Id As Long

    GetWindowThreadProcessId hWnd, Id
    GetProcessID = Id
End Function

Public Function GetProcessPath(PID As Long)

Dim lReturn As Long
Dim szFileName As String
Dim Buffer(2048)        As Byte
Dim dwLength            As Long

    lReturn = OpenProcess(PROCESS_QUERY_INFORMATION Or _
                                   PROCESS_VM_READ, 0, PID)
    If lReturn = 0 Then
        Debug.Print "Failed to open process: " & PID
        Exit Function
    End If
        
    dwLength = GetModuleFileNameExW(lReturn, 0, VarPtr(Buffer(0)), UBound(Buffer))
    If dwLength = 0 Then
        dwLength = GetProcessImageFileName(lReturn, VarPtr(Buffer(0)), UBound(Buffer))
        
        szFileName = Left$(Buffer, dwLength)
        szFileName = g_DeviceCollection.ConvertToLetterPath(szFileName)

        If dwLength = 0 Then
            CloseHandle lReturn
            Exit Function
        End If
    Else
        szFileName = Left$(Buffer, dwLength)
    End If
    
    GetProcessPath = szFileName
    
    CloseHandle lReturn
End Function

Public Function Is64bit() As Boolean
    Dim Handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False

    ' Now check to see if IsWow64Process function exists
    Handle = GetProcAddress(GetModuleHandle("kernel32"), _
                   "IsWow64Process")

    If Handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc
    End If

    Is64bit = bolFunc

End Function

Public Function IsPIDValid(ByVal PID As Long) As Long

    Debug.Print "IsPIDValid():PID:: " & IsPIDValid

    Dim hProcess As Long
    Dim dwRetval As Long
    
    If (PID = 0) Then
        IsPIDValid = 1: Exit Function
    End If
    
    If (PID < 0) Then
        IsPIDValid = 0: Exit Function
    End If
    
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, PID)
    If (hProcess = pNull) Then
        'invalid parameter means PID isn't in the system
        If (WinBase.GetLastError() = ERROR_INVALID_PARAMETER) Then
            IsPIDValid = 0: Exit Function
        End If
        'some other error
        IsPIDValid = -1
    End If

    dwRetval = WaitForSingleObject(hProcess, 0)
    Call CloseHandle(hProcess)  'otherwise you'll be losing handles
    
    Select Case dwRetval
    
    Case WAIT_OBJECT_0:
        IsPIDValid = 0: Exit Function
    Case WAIT_TIMEOUT:
        IsPIDValid = 1: Exit Function
        
    Case Else
        IsPIDValid = -1: Exit Function

    End Select


End Function
