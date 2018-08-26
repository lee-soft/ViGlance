Attribute VB_Name = "TaskbarHelper"
'--------------------------------------------------------------------------------
'    Component  : TaskbarHelper
'    Project    : prjSuperBar
'
'    Description: Provides information about the native OS taskbar
'
'--------------------------------------------------------------------------------
Option Explicit

Public g_ReBarWindow32Hwnd As Long
Public g_RunningProgramsHwnd As Long
Public g_StartButtonHwnd As Long
Public g_TaskBarHwnd As Long
Public g_StartMenuHwnd As Long
Public g_StartMenuOpen As Boolean
Public g_viStartRunning As Boolean
Public g_viStartOrbHwnd As Long

Private m_TargetTaskList As TaskList
Private m_taskbarRect As win.RECT

Public Function UpdateViStartHwnds() As Boolean

Dim newViStartOrbHwnd As Long
    newViStartOrbHwnd = FindWindow(0&, "##VIGLANCE_MODE##")

    If newViStartOrbHwnd <> g_viStartOrbHwnd Then
        g_viStartOrbHwnd = newViStartOrbHwnd
        
        UpdateViStartHwnds = True
    End If

End Function

Public Function UpdatehWnds() As Boolean
Dim newTaskBarHwnd As Long
Dim updatedHwnd As Boolean

    updatedHwnd = False
    newTaskBarHwnd = FindWindow("Shell_TrayWnd", "")
    
    If newTaskBarHwnd = 0 Then
        Exit Function
    End If

    If newTaskBarHwnd <> g_TaskBarHwnd Then
        g_TaskBarHwnd = newTaskBarHwnd
        g_ReBarWindow32Hwnd = FindWindowEx(ByVal g_TaskBarHwnd, ByVal 0&, "ReBarWindow32", vbNullString)
        'g_ReBarWindow32Hwnd = FindWindowEx(ByVal FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MSTaskSwWClass", "Running Applications"), _
                                    ByVal 0&, "ToolbarWindow32", "Running Applications")
    
        g_RunningProgramsHwnd = FindWindowEx(FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MsTaskSwWClass", vbNullString), ByVal 0&, "ToolbarWindow32", vbNullString)
        If g_RunningProgramsHwnd = 0 Then
            g_RunningProgramsHwnd = FindWindowEx(FindWindowEx(ByVal g_ReBarWindow32Hwnd, ByVal 0&, "MSTaskSwWClass", vbNullString), ByVal 0&, "MSTaskListWClass", vbNullString)
        
            If g_RunningProgramsHwnd = 0 Then
                'Reset update trigger (forcing routine to later update again)
                g_TaskBarHwnd = -1
            End If
        End If
        
        g_StartButtonHwnd = FindWindowEx(g_TaskBarHwnd, 0, "Button", vbNullString)
        If g_StartButtonHwnd = 0 Then
            'Windows Vista/Seven
            
            g_StartButtonHwnd = FindWindow("Button", "Start")
            If g_StartButtonHwnd = 0 Then
                'Reset update trigger (forcing routine to later update again)
                g_TaskBarHwnd = -1
                Debug.Print "Start button not found!"
                
            Else
                'g_WindowsVista = True
            End If
            
        End If
    
        updatedHwnd = True
    End If
    
    UpdatehWnds = updatedHwnd
End Function

Public Sub MakeTaskbarTransparent(ByVal bLevel As Byte)
    If ApplicationOptions.TaskBarFade = False Then
        bLevel = 255
    End If

Dim lOldStyle As Long
Static lOldLevel As Byte

    If bLevel = lOldLevel Then
        Exit Sub
    End If
    
    lOldLevel = bLevel

    If (g_TaskBarHwnd <> 0) Then
        lOldStyle = GetWindowLong(g_TaskBarHwnd, GWL_EXSTYLE)
        SetWindowLong g_TaskBarHwnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
        SetLayeredWindowAttributes g_TaskBarHwnd, 0, bLevel, &H2&
    End If
End Sub

Public Function IsViStartOpen()

    IsViStartOpen = False
    
    If FindWindow("ThunderRT6FormDC", "ViStart_PngNew") <> 0 Then
        IsViStartOpen = True
    End If
    
    'Debug.Print "VGMODE::" & IsViStartOpen
    
End Function

Private Function GetEdge(rc As RECT) As Long

Dim uEdge As Long: uEdge = -1

    If (rc.Top = rc.Left) And (rc.Bottom > rc.Right) Then
        uEdge = ABE_LEFT
    ElseIf (rc.Top = rc.Left) And (rc.Bottom < rc.Right) Then
        uEdge = ABE_TOP
    ElseIf (rc.Top > rc.Left) Then
        uEdge = abe_bottom
    Else
        uEdge = ABE_RIGHT
    End If
    
    GetEdge = uEdge

End Function

Function GetTaskBarEdge() As AbeBarEnum
        
Dim abd As APPBARDATA

    abd.cbSize = LenB(abd)
    abd.hWnd = g_TaskBarHwnd
    SHAppBarMessage ABM_GETTASKBARPOS, abd
    
    GetTaskBarEdge = GetEdge(abd.rc)

End Function

Function IsTaskBarBehindWindow(hWnd As Long)
    
    If GetZOrder(g_TaskBarHwnd) > GetZOrder(hWnd) Then
        IsTaskBarBehindWindow = True
    Else
        IsTaskBarBehindWindow = False
    End If
    
End Function

Function IsWindowTopMost(hWnd As Long)

Dim windowStyle As Long

    IsWindowTopMost = False
    windowStyle = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    If IsStyle(windowStyle, WS_EX_TOPMOST) Then
        IsWindowTopMost = True
    End If

End Function

Public Function IsStartMenuOpen() As Boolean
    If IsWindow(g_StartMenuHwnd) = False Then
        g_StartMenuHwnd = FindWindow("DV2ControlHost", "Start Menu")
        If g_StartMenuHwnd = 0 Then
            g_StartMenuHwnd = FindWindow("DV2ControlHost", vbNullString)
        End If
    End If
    
    IsStartMenuOpen = IsWindowVisible(g_StartMenuHwnd)
End Function

Public Function ShowStartMenu()
    SendMessage g_TaskBarHwnd, ByVal WM_SYSCOMMAND, ByVal SC_TASKLIST, ByVal 0
End Function

Public Function EnumWindowsAsTaskList(ByRef srcCollection As TaskList)
    On Error GoTo Handler

' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'

    If Not srcCollection Is Nothing Then
    
        Set m_TargetTaskList = srcCollection
        Call EnumWindows(AddressOf fEnumWindowsCallBack, ByVal 0)
    End If
    
    Exit Function
Handler:
    LogError Err.number, "EnumerateWindowsAsTaskObject(); " & Err.Description, "TaskbarHelper"
    
End Function

Public Function fEnumWindowsCallBack(ByVal hWnd As Long, ByVal lParam As Long) As Long

    If IsVisibleToTaskBar(hWnd) Then
        m_TargetTaskList.AddWindowByHwnd hWnd
    End If

fEnumWindowsCallBack = True
End Function

Public Function IsVisibleToTaskBar(hWnd As Long) As Boolean

Dim lExStyle    As Long
Dim bNoOwner    As Boolean

    IsVisibleToTaskBar = False
    
    ' This callback function is called by Windows (from
    ' the EnumWindows API call) for EVERY window that exists.
    ' It populates the listbox with a list of windows that we
    ' are interested in.
    '
    ' Windows to display are those that:
    '   -   are not this app's
    '   -   are visible
    '   -   do not have a parent
    '   -   have no owner and are not Tool windows OR
    '       have an owner and are App windows
    '       can be activated

    If IsWindowVisible(hWnd) Then
        If (GetParent(hWnd) = 0) Then
            bNoOwner = (GetWindow(hWnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
            
                IsVisibleToTaskBar = True
            End If
        End If
    End If
End Function

Public Function IsVisibleOnTaskbar(lExStyle As Long) As Boolean
    IsVisibleOnTaskbar = False

    If (lExStyle And WS_EX_APPWINDOW) Then
        IsVisibleOnTaskbar = True
    End If
End Function

Public Function IsMouseInsideWindowsTaskBar() As Boolean

Dim CurrentCursorPos As win.POINTL

    GetWindowRect g_TaskBarHwnd, m_taskbarRect
    GetCursorPos CurrentCursorPos

    IsMouseInsideWindowsTaskBar = PointInsideOfRect(CurrentCursorPos, m_taskbarRect)

End Function
