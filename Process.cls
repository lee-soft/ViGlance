VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Window"
'--------------------------------------------------------------------------------
'    Component  : Process
'    Project    : prjSuperBar
'
'    Description: A Process object. Holds information about native OS
'                 processes.
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_Windows As Collection

'local variable(s) to hold property value(s)
Private mvarCaption As String 'local copy
Private mvarID As Long 'local copy
Private mvarPath As String 'local copy
Private mvarFlashing As Boolean 'local copy
Private mvarJumpList As JumpList
Private mvarArguments As String

Private mvarhIcon As Long
Private m_Icon32 As AlphaIcon

Private m_createdFromPath As Boolean
Private m_X As Long

Private m_Pinned As Boolean
Private m_Running As Boolean

Public Function GetKey()
    
Dim Id As String
    If mvarPath <> "" Then
        Id = mvarPath
    Else
        Id = mvarID
    End If

    GetKey = Id & "_" & mvarArguments
End Function

Public Property Let Arguments(newArguments As String)
    mvarArguments = newArguments
End Property

Public Property Get Arguments() As String
    Arguments = mvarArguments
End Property

Public Function GetJumpLists()
    GetJumpLists = mvarJumpList.GetMRUList
End Function

Public Function UpdateWindowImages()

Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        If thisWindow.Image.width = 0 Or thisWindow.hWnd = g_hwndForeGroundWindow Then
            If thisWindow.WindowState <> vbMinimized Then
                thisWindow.UpdateImage
            End If
        End If
    Next

End Function

Public Function IsRunning() As Boolean
    IsRunning = IsPIDValid(mvarID)
End Function

Public Property Let Running(ByVal newValue As Boolean)
    m_Running = newValue
End Property

Public Property Get Running() As Boolean
    Running = m_Running
End Property

Public Property Get Pinned() As Boolean
    Pinned = m_Pinned
End Property

Public Property Let Pinned(ByVal newValue As Boolean)
    m_Pinned = newValue
End Property

Public Property Get PinnedAndClosed() As Boolean
    If m_Pinned = True And m_Running = False Then
        PinnedAndClosed = True
    Else
        PinnedAndClosed = False
    End If
End Property

Public Property Get X() As Long
    X = m_X
End Property

Public Property Let X(ByVal newValue As Long)
    m_X = newValue
End Property

Public Function RestoreAllWindows()
    
Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        ShowWindowTimeout thisWindow.hWnd, SW_RESTORE
    Next
    
End Function

Public Function MinimizeAllWindows()
    
Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        ShowWindowTimeout thisWindow.hWnd, SW_MINIMIZE
    Next
    
End Function

Public Function RequestCloseAllWindows()
    
Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        PostMessage thisWindow.hWnd, WM_CLOSE, 0&, 0&
    Next
    
End Function

Public Property Let Flashing(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Flashing = 5
    mvarFlashing = vData
End Property

Public Property Get Flashing() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Flashing
    Flashing = mvarFlashing
End Property

Public Function UpdateFlashStatus()
    
Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        If thisWindow.Flashing Then
            Me.Flashing = True
            Exit Function
        End If
    Next
    Me.Flashing = False
    
End Function

Public Function GetWindowByHWND(ByRef hWnd As Long) As Window
    
On Error GoTo ObjectDoesNotExist
    Dim theWindow As Window

    Set theWindow = m_Windows(CStr("a" & hWnd))
    Set GetWindowByHWND = theWindow
    Exit Function

ObjectDoesNotExist:
    Set GetWindowByHWND = Nothing
    
End Function

Public Property Get Window() As Collection
    Set Window = m_Windows
End Property

Public Function IsIconCreatedFromPath() As Boolean
    IsIconCreatedFromPath = m_createdFromPath
End Function

Public Function HasWindows() As Boolean
    
    If m_Windows.Count > 0 Then
        HasWindows = True
    End If
    
End Function

Public Function CheckWindowHandles()
    'Decide weather windows should be visible on taskbar

Dim thisWindow As Window
Dim lStyle As Long
    
    For Each thisWindow In m_Windows
        lStyle = GetWindowLong(thisWindow.hWnd, GWL_STYLE)
        
        If (IsWindow(thisWindow.hWnd) = False) Or _
            ((lStyle And WS_VISIBLE) = False) Then
            
            m_Windows.Remove "a" & thisWindow.hWnd
        End If
    Next

End Function

Public Function CreateIconFromhWnd(hWnd As Long) As Boolean
    On Error GoTo CreateBlankIcon
    CreateIconFromhWnd = True
    
    Set m_Icon32 = New AlphaIcon
    mvarhIcon = IconHelper.GetIconFromHwnd(hWnd)
    
    If mvarhIcon <> 0 Then
        m_Icon32.CreateFromHICON mvarhIcon
        DestroyIcon mvarhIcon
    Else
        Debug.Print "Failed to GetIconFromhWnd"
    End If
    
    Exit Function
CreateBlankIcon:

    CreateIconFromhWnd = False
    mvarhIcon = 0
    
End Function

Public Function CreateIconFromPath() As Boolean
    On Error GoTo CreateBlankIcon
    CreateIconFromPath = True

    Set m_Icon32 = New AlphaIcon
    mvarhIcon = IconHelper.GetApplicationIcon(mvarPath)
    
    If mvarhIcon <> 0 Then
        m_Icon32.CreateFromHICON mvarhIcon
        DestroyIcon mvarhIcon
        
        m_createdFromPath = True
    Else
        Debug.Print "Failed to CreateIconFromPath", mvarPath
    End If
    
    Exit Function
CreateBlankIcon:

    CreateIconFromPath = False
    mvarhIcon = 0
End Function

'Constructor
Public Function Constructor(newPID As Long, newPath As String)
    
Dim newJumpList As JumpList
    
    mvarID = newPID
    mvarPath = newPath
    
    mvarhIcon = IconHelper.GetApplicationIcon(mvarPath)
    Set m_Icon32 = New AlphaIcon

    If Len(mvarPath) > 3 Then
        If Right$(LCase$(mvarPath), 3) = "exe" Then
            Set newJumpList = GetImageJumpList(mvarPath)
            If Not newJumpList Is Nothing Then
                Set mvarJumpList = newJumpList
            End If
        
            If mvarhIcon <> 0 Then
                m_Icon32.CreateFromHICON mvarhIcon
                DestroyIcon mvarhIcon
            End If
            'Set m_Icon32 = IconToAlphaBitmap(IconHelper.GetApplicationIcon(mvarPath))
        Else
            mvarhIcon = 0
        End If
    Else
        mvarhIcon = 0
    End If
End Function

Public Function IconIsValid() As Boolean
    If mvarhIcon <> 0 Then
        IconIsValid = True
    End If
End Function

Public Property Get Image() As GDIPImage
    Set Image = m_Icon32.Image
End Property

Public Property Get Path() As String
    Path = mvarPath
End Property

Public Function PrintWindows()
    
Dim thisWindow As Window
    
    For Each thisWindow In m_Windows
        Debug.Print thisWindow.hWnd & " : " & thisWindow.Caption
    Next
    
End Function

Public Property Get WindowCount()
    WindowCount = m_Windows.Count
End Property

Private Sub Class_Initialize()
    m_Pinned = False
    m_Running = False

    Set mvarJumpList = New JumpList
    Set m_Windows = New Collection
End Sub

Public Function AddWindow(newWindow As Window)
    If Exists(m_Windows, "a" & newWindow.hWnd) = False Then
    
        Set newWindow.Parent = Me
        m_Windows.Add newWindow, "a" & newWindow.hWnd
    End If
End Function

Public Function RemoveWindow(theWindow As Window)
    If Exists(m_Windows, "a" & theWindow.hWnd) = True Then
        m_Windows.Remove "a" & theWindow.hWnd
    End If
End Function

Public Function CreateWindow(hWnd As Long) As Window

Dim newWindow As New Window
    newWindow.hWnd = hWnd
    
    Set CreateWindow = newWindow

End Function

Public Property Get Id() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    Id = mvarID
End Property

Public Property Let Id(ByVal newID As Long)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID
    mvarID = newID
End Property

Public Property Let Caption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property

Public Property Get Caption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property

Private Sub Class_Terminate()
On Error Resume Next
    Set m_Icon32 = Nothing
End Sub
