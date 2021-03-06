VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : TaskList
'    Project    : prjSuperBar
'
'    Description: Window database manager. It manages all windows and their
'                 respective processes.
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_Processes As Collection
Private m_ActiveWindow As Window
Private m_ActiveProcess As Process

Public Function UpdateFlashStatusOfEachProcess()
    On Error GoTo Handler

Dim thisProcess As Process
    
    For Each thisProcess In m_Processes
        thisProcess.UpdateFlashStatus
    Next

    Exit Function
Handler:
    LogError Err.number, "UpdateFlashStatusOfEachProcess(); " & Err.Description, "ObjTaskList"
End Function

Public Function GetProcessByHWND(ByRef hWnd As Long) As Window
    
Dim thisProcess As Process
Dim thisWindow As Window
    
    For Each thisProcess In m_Processes
        Set thisWindow = thisProcess.GetWindowByHWND(hWnd)
        
        If Not thisWindow Is Nothing Then
            Set GetProcessByHWND = thisProcess
            Exit For
        End If
    Next
    
End Function

Public Function GetWindowByHWND(ByRef hWnd As Long) As Window
    
Dim thisProcess As Process
Dim thisWindow As Window
    
    For Each thisProcess In m_Processes
        Set thisWindow = thisProcess.GetWindowByHWND(hWnd)
        
        If Not thisWindow Is Nothing Then
            Set GetWindowByHWND = thisWindow
            Exit For
        End If
    Next
    
End Function

Public Property Get ActiveProcess()
    Set ActiveProcess = m_ActiveProcess
End Property

Public Function DeleteDeadHandles()
    On Error GoTo Handler
    
Dim thisProcess As Process
    
    Set m_ActiveWindow = New Window
    Set m_ActiveProcess = New Process
    
    For Each thisProcess In m_Processes
        thisProcess.CheckWindowHandles
        
        If Not thisProcess.HasWindows Then
        
            If thisProcess.Pinned Then
                thisProcess.Running = False
            Else
                'Object falls out of scope
                If Exists(m_Processes, thisProcess.GetKey) Then
                    m_Processes.Remove thisProcess.GetKey
                End If
            End If
        End If
    Next
    
    Exit Function
Handler:
    LogError Err.number, "DeleteDeadHandles(); " & Err.Description, "ObjTaskList"
End Function

Public Property Get Processes() As Collection
    Set Processes = m_Processes
End Property

Public Function PrintProcesses()
    
Dim thisProcess As Process
    
    For Each thisProcess In m_Processes
        Debug.Print thisProcess.Id & " - " & thisProcess.Path
        thisProcess.PrintWindows
    Next
    
End Function

Public Function RemoveWindow(ByRef theWindow As Window)

    If Not theWindow Is Nothing Then
        theWindow.Parent.RemoveWindow theWindow
        
        If Not theWindow.Parent.HasWindows Then
        
            If theWindow.Parent.Pinned Then
                theWindow.Parent.Running = False
            Else
                'Object falls out of scope
                If Exists(m_Processes, theWindow.Parent.GetKey) Then
                    m_Processes.Remove theWindow.Parent.GetKey
                End If
            End If
        End If
    End If
End Function

Public Function AddWindowByHwnd(hWnd As Long)
Dim processID As Long
Dim processPath As String


Dim thisProcess As Process
Dim ExistsInTaskList As Boolean
Dim newProcess As Process
Dim newWindow As Window

    'Check hWnd is valid
    If hWnd = 0 Then
        'Debug.Print "Invalid hWnd; 0"
        Exit Function
    End If

    processID = GetProcessID(hWnd)
    processPath = GetProcessPath(processID)
    
    ExistsInTaskList = False
    
    If processPath <> vbNullString Then
        'Check if ProcessPath exists in collection
        ExistsInTaskList = Exists(m_Processes, processPath & "_")
        
        If ExistsInTaskList Then
            Set thisProcess = m_Processes(processPath & "_")
        End If
        
    Else
        'Check if ProcessID exists in collection
        ExistsInTaskList = Exists(m_Processes, CStr(processID) & "_")
        
        If ExistsInTaskList Then
            Set thisProcess = m_Processes(CStr(processID) & "_")
        End If
    End If
    
    If Not ExistsInTaskList Then
    
        Set newProcess = New Process: newProcess.Constructor processID, processPath
        newProcess.X = -1
        
        If Not newProcess.IconIsValid Then
            If newProcess.CreateIconFromhWnd(hWnd) = False Then
                Debug.Print "Troublesome; " & processPath
            End If
        End If
        
        m_Processes.Add newProcess, newProcess.GetKey
        Set thisProcess = newProcess
    Else
        If thisProcess.WindowCount = 1 Then
            If processPath <> "" And _
                thisProcess.IsIconCreatedFromPath = False Then
                
                thisProcess.CreateIconFromPath
            End If
        End If
        
        If processID <> 0 Then
            thisProcess.Id = processID
        End If
    End If
    
    thisProcess.Running = True
    Set newWindow = thisProcess.CreateWindow(hWnd)
    If g_hwndForeGroundWindow = hWnd Then
        Set m_ActiveWindow = newWindow
        Set m_ActiveProcess = thisProcess
    End If
    
    thisProcess.AddWindow newWindow
End Function

Private Sub Class_Initialize()
    Set m_Processes = New Collection
End Sub

