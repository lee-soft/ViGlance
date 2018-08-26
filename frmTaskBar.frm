VERSION 5.00
Begin VB.Form frmTaskbar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Running Applications"
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   210
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   14
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   14
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timAutoCloseMenu 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1200
      Top             =   780
   End
   Begin VB.Timer timAnimate 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1260
      Top             =   1200
   End
   Begin VB.Timer timEnumerateTasks 
      Interval        =   500
      Left            =   540
      Top             =   1200
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmTaskbar
'    Project    : prjSuperBar
'
'    Description: The physical taskbar form. Everything is rendered to this form
'                 It will physically replace the native Windows taskbar
'
'--------------------------------------------------------------------------------
Option Explicit

Public Event onVerticleMode(VerticleMode As Boolean)

Private Const BAND_NOTHING As Long = -1
Private Const BAND_NORMAL As Long = 0
Private Const BAND_HOVER_UNCLICKED As Long = 40
Private Const BAND_LEFT_BUTTON_DOWN As Long = 80
Private Const BAND_ACTIVE As Long = 160
Private Const BAND_NOTIFY As Long = 200

Private Const BAND_PINNED As Long = 240
Private Const BAND_PINNED_HOVER As Long = 280
Private Const BAND_PINNED_SELECTED As Long = 320

Private Const GROUP_2_NORMAL As Long = 40
Private Const GROUP_2_ACTIVE As Long = 160

Private Const GROUP_3_NORMAL As Long = 80
Private Const GROUP_3_ACTIVE As Long = 200

'Private Const APP_BAND_HEIGHT As Long = 40

Private winSize As Size
Private srcPoint As POINTAPI

' Create a Graphics object:
Private m_gfx As New GDIPGraphics
' Create the Image and load the file:
Private m_TaskBandImage As New GDIPImage
Private m_TaskGroupImage As New GDIPImage

Private m_TaskBandEmpty As New GDIPImage
Private m_TaskBandLight As New GDIPImage
Private m_TaskBandGraphics As New GDIPGraphics
'Private m_imageAttr As New GDIPImageAttributes

Private m_TaskbarBackgroundBitmap As GDIPBitmap
Private m_TaskbarBackgroundGraphics As GDIPGraphics
Private m_ToolListItem As RECT

Private m_TaskList As TaskList
Private m_VisibleList() As Process
Private m_SolidBrush As New GDIPBrush
   
Private curWinLong As Long
Private blendFunc32bpp As BLENDFUNCTION

Private m_RolloverIndex As Long
Private m_LastMouseButton As Integer
Private m_isTracking As Boolean

Private WithEvents m_SubMenu As frmSubMenu
Attribute m_SubMenu.VB_VarHelpID = -1

Private m_SubMenu_finalRight As Long
Private m_SubMenu_finalLeft As Long
Private m_SubMenu_finalBottom As Long
Private m_SubMenu_finalTop As Long


Private m_moveLeftVelocity As Single
Private m_moveTopVelocity As Single
Private m_moveBottomVelocity As Single
Private m_moveRightVelocity As Single

Private m_Shell_Hook_Msg_ID As Long

Private m_Rect As RECT
Private m_currentPosition As windowPos
Private m_userInvoked As Boolean

Private m_SubMenuLayout As RECTF
Private m_ignoreClick As Boolean
Private m_GroupMenu As clsMenu
Private m_lastHoveredGroup As Process

Private m_aspectRatio As Single
Private m_currentTaskBarEdge As AbeBarEnum

Private m_notifiedLeft As Boolean
Private m_notifiedInside As Boolean
Private m_transparentValue As Byte

Private m_MoveOffset As Long
Private m_GapPosition As Long

Private m_MouseHeldCounter As Long
Private m_targetCaptureWindow As Window

Private WithEvents timPopupDelay As Timer
Attribute timPopupDelay.VB_VarHelpID = -1
Private WithEvents timHideSubMenuDelay As Timer
Attribute timHideSubMenuDelay.VB_VarHelpID = -1
Private WithEvents timTestMouse As Timer
Attribute timTestMouse.VB_VarHelpID = -1
Private WithEvents timKeepOnTaskBar As Timer
Attribute timKeepOnTaskBar.VB_VarHelpID = -1
Private WithEvents timCheckLoadSize As Timer
Attribute timCheckLoadSize.VB_VarHelpID = -1
Private WithEvents timPopupNextDelay As Timer
Attribute timPopupNextDelay.VB_VarHelpID = -1
Private WithEvents timReCaptureImage As Timer
Attribute timReCaptureImage.VB_VarHelpID = -1

Private m_MouseDown As Boolean
Private m_currentJumpLists
Private m_hWndToTestIfValid As Long
Private m_ignoreMouseMove As Boolean

Implements IHookSink

Private Function PrepareSubMenu(ShowUpdates As Boolean) As Boolean
    Debug.Print "PrepareSubMenu:: " & ShowUpdates

Dim proposedHeight As Long
Dim proposedWidth As Long
Dim heightDifference As Long
Dim mSubMenu_FutureWidth As Long

Dim taskBarPosition As AbeBarEnum
Dim overFlow As Long
Dim temp As Long

    If m_lastHoveredGroup Is Nothing Then
        Debug.Print "Impossible m_lastHoveredGroup Object"
    
        PrepareSubMenu = False
        Exit Function
    End If
    
    If Not m_lastHoveredGroup.HasWindows Then
        Debug.Print "Doesn't have windows"
    
        PrepareSubMenu = False
        Exit Function
    End If
    
    mSubMenu_FutureWidth = GetSystemMetrics(SM_CXSCREEN) - m_Rect.Left
    'Debug.Print "mSubMenu_FutureWidth; " & mSubMenu_FutureWidth
    
    m_lastHoveredGroup.UpdateWindowImages
    Set m_SubMenu.WindowList = m_lastHoveredGroup.Window

    m_SubMenu.AspectRatio = 1
    
    proposedHeight = m_SubMenu.PredictHeight
    proposedWidth = m_SubMenu.PredictWidth
    
    'Debug.Print "proposedHeight; " & proposedHeight
    'Debug.Print "m_currentPosition.Y; " & m_currentPosition.Y
    
    'Debug.Print "m_RolloverIndex; " & m_RolloverIndex
    m_SubMenu_finalLeft = ((m_RolloverIndex * 62) / m_aspectRatio) _
                            - (proposedWidth / 2) _
                            + (m_TaskBandImage.width / 2)
                            
    If (m_SubMenu_finalLeft < 0) Then
         m_SubMenu_finalLeft = 0
    End If
    
    m_SubMenu_finalRight = m_SubMenu_finalLeft + proposedWidth
    'Debug.Print "PROPSEDHEIGHT>>: " & proposedHeight
    
    If m_SubMenu_finalRight > mSubMenu_FutureWidth Then
        If m_SubMenu_finalLeft > 0 Then
            overFlow = m_SubMenu_finalRight - mSubMenu_FutureWidth
            temp = m_SubMenu_finalLeft - overFlow
            
            If temp > -1 Then
                m_SubMenu_finalLeft = m_SubMenu_finalLeft - overFlow
                m_SubMenu_finalRight = 0
            Else
                Debug.Print "WE GOT A PROBLEM[1] ------------------------"
                m_SubMenu.AspectRatio = mSubMenu_FutureWidth / ((temp * -1) + m_SubMenu_finalRight)
                
                m_SubMenu_finalLeft = 0
                m_SubMenu_finalRight = mSubMenu_FutureWidth
            End If
        Else
            m_SubMenu.AspectRatio = (mSubMenu_FutureWidth - (m_SubMenu.ReturnPadding)) / (proposedWidth)
            m_SubMenu_finalRight = mSubMenu_FutureWidth
            
            'Debug.Print "PROPSEDHEIGHT>>: " & (mSubMenu_FutureWidth) / (proposedWidth)
            proposedHeight = ((proposedHeight - SUBMENU_Y_PADDING) * m_SubMenu.AspectRatio) + SUBMENU_Y_PADDING
        End If
    End If
                            
    taskBarPosition = GetTaskBarEdge
    
    If (m_SubMenu.ScaleHeight < proposedHeight) Or (m_SubMenu.height = 1) Then
    
        heightDifference = m_SubMenu.ScaleHeight - proposedHeight
        
        'Realign border
        m_SubMenuLayout.Top = m_SubMenuLayout.Top - heightDifference
        m_SubMenuLayout.height = m_SubMenuLayout.height - heightDifference
        'm_SubMenuLayout.Left = m_SubMenuLayout.Left - widthDifference
        'm_SubMenuLayout.Width = m_SubMenuLayout.Width - widthDifference
        
        If taskBarPosition = abe_bottom Then
        
            MoveWindow m_SubMenu.hWnd, _
                        (m_currentPosition.X), _
                        (m_currentPosition.Y) - proposedHeight - 1, _
                        GetSystemMetrics(SM_CXSCREEN) - m_Rect.Left, _
                        proposedHeight, _
                        False
        ElseIf taskBarPosition = ABE_TOP Then
        
            MoveWindow m_SubMenu.hWnd, _
                        (m_currentPosition.X), _
                        (m_Rect.Bottom + 1), _
                        GetSystemMetrics(SM_CXSCREEN) - m_Rect.Left, _
                        proposedHeight, _
                        False
        End If
        
        'Doesn't really matter
        If ShowUpdates Then
            m_SubMenu.ReInitSurface
            m_SubMenu.DrawBorder RECTFtoL(m_SubMenuLayout)
        End If
        
        'Debug.Print "New height is; " & m_SubMenu.ScaleHeight
    End If
    
    If taskBarPosition = ABE_TOP Then
        m_SubMenu_finalTop = 0
        m_SubMenu_finalBottom = m_SubMenu_finalTop + proposedHeight
        
    ElseIf taskBarPosition = abe_bottom Then
    
        m_SubMenu_finalBottom = m_SubMenu.ScaleHeight
        m_SubMenu_finalTop = m_SubMenu_finalBottom - proposedHeight
        
        Debug.Print "m_SubMenu_finalBottom:: " & m_SubMenu.ScaleHeight
        Debug.Print "m_SubMenu_finalTop:: " & m_SubMenu_finalBottom - proposedHeight
    End If
    
    'Debug.Print m_SubMenu_finalLeft & ":" & m_SubMenu_finalRight & ":" & m_SubMenu_finalTop & ":" & m_SubMenu_finalBottom
        
    PrepareSubMenu = True
End Function

Private Function TriggerSubMenu() As Boolean

    If m_lastHoveredGroup Is Nothing Then
        Exit Function
    End If
    
    TriggerSubMenu = False

    'Prevent submenu from showing if submenu contains the foreground window
    If Not m_lastHoveredGroup.GetWindowByHWND(g_hwndForeGroundWindow) Is Nothing Then
        If m_lastHoveredGroup.WindowCount < 1 Then
            Exit Function
        End If
    End If

    If PrepareSubMenu(False) = False Then
        Exit Function
    End If
    
    With m_SubMenuLayout
        .height = m_SubMenu_finalBottom
        .width = m_SubMenu_finalRight
        .Left = m_SubMenu_finalLeft
        .Top = m_SubMenu_finalTop
    End With

    m_SubMenu.ReInitSurface
    
    m_SubMenu.DrawBorder RECTFtoL(m_SubMenuLayout)
    
    m_SubMenu.AllowListRedraw = True
    
    m_SubMenu.DrawList RECTFtoL(m_SubMenuLayout)
    
    'If m_userInvoked = False Then
        'ShowWindow m_SubMenu.hWnd, SW_SHOW
    'Else
        ShowWindow m_SubMenu.hWnd, SW_SHOW
        StayOnTop m_SubMenu, True
    'End If
    
    TriggerSubMenu = True
End Function

Private Sub TriggerSubMenuWithAnimation()
    Debug.Print "TriggerSubMenuWithAnimation"
    
    If ApplicationOptions.GlideAnimation = False Then
        TriggerSubMenu
        Exit Sub
    End If
    
    If PrepareSubMenu(True) = False Then
        Exit Sub
    End If

    timPopupDelay.Enabled = False
    
    m_moveLeftVelocity = (m_SubMenu_finalLeft - (m_SubMenuLayout.Left)) / 10
    m_moveRightVelocity = (m_SubMenu_finalRight - (m_SubMenuLayout.width)) / 10
    m_moveBottomVelocity = (m_SubMenu_finalBottom - (m_SubMenuLayout.height)) / 10
    m_moveTopVelocity = (m_SubMenu_finalTop - (m_SubMenuLayout.Top)) / 10
    
    If m_moveLeftVelocity <> 0 Or m_moveTopVelocity <> 0 Or m_moveBottomVelocity <> 0 Or m_moveRightVelocity <> 0 Then
    
        m_SubMenu.AllowListRedraw = False
        timAnimate.Enabled = True
    Else
        Debug.Print "Animation isn't worth activating, no left/top velocity"
        
        m_SubMenu.DrawList RECTFtoL(m_SubMenuLayout)
    End If
End Sub

Private Sub Form_Click()
    Debug.Print "click recieved!"
End Sub

Sub Form_Initialize()

    Set m_TaskList = New TaskList
    Set m_GroupMenu = New clsMenu

    Set m_TaskbarBackgroundGraphics = New GDIPGraphics
    Set m_TaskbarBackgroundBitmap = New GDIPBitmap
    Set m_SubMenu = frmSubMenu

    Set timPopupDelay = Controls.Add("VB.Timer", "timPopupDelay", Me)
    Set timHideSubMenuDelay = Controls.Add("VB.Timer", "timHideSubMenuDelay", Me)
    Set timTestMouse = Controls.Add("VB.Timer", "timTestMouse", Me)
    Set timKeepOnTaskBar = Controls.Add("VB.Timer", "timKeepOnTaskBar", Me)
    Set timCheckLoadSize = Controls.Add("VB.Timer", "timCheckLoadSize", Me)
    Set timPopupNextDelay = Controls.Add("VB.Timer", "timPopupNextDelay", Me)
    Set timReCaptureImage = Controls.Add("VB.Timer", "timReCaptureImage", Me)

    timPopupNextDelay.Interval = ApplicationOptions.PopupNextDelay
    timPopupNextDelay.Enabled = True
    
    timReCaptureImage.Interval = ApplicationOptions.ReCaptureImageDelay
    timReCaptureImage.Enabled = True
    
    timHideSubMenuDelay.Interval = ApplicationOptions.HideChildrenDelay
    timCheckLoadSize.Interval = 3000

    If OptionsHelper.ApplicationOptions.AutoClick Then
        timPopupDelay.Interval = ApplicationOptions.PopupDelay
    ElseIf OptionsHelper.ApplicationOptions.InstantSpawn Then
        timPopupDelay.Interval = 1
    End If
    
    timTestMouse.Interval = 100
    timTestMouse.Enabled = True
    
    m_RolloverIndex = -1
    m_aspectRatio = 1
    
    m_GroupMenu.AddItem 1, "Minimize"
    m_GroupMenu.AddItem 2, "Restore"
    m_GroupMenu.AddSeperater
    m_GroupMenu.AddItem 5, "Pin"
    m_GroupMenu.AddItem 4, "Close All Windows"

    timKeepOnTaskBar.Interval = 5000
    timKeepOnTaskBar.Enabled = True
    
    ParsePinnedList
    
    DragAcceptFiles Me.hWnd, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_MouseDown = True
       
    timPopupDelay.Enabled = False
    
    If m_LastMouseButton <> Button Then
        m_LastMouseButton = Button
        DrawTaskList
    End If
End Sub

Private Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim newUserSelection As Boolean
Dim newRolloverIndex As Long
Static lastX As Single

    timHideSubMenuDelay.Enabled = False

    If Not m_isTracking Then
        m_isTracking = True
        ReTrackMouse
    End If

    If m_aspectRatio <> 1 Then
        X = X * m_aspectRatio
    End If
    
    If Button = vbLeftButton And Not m_lastHoveredGroup Is Nothing Then
        If m_MoveOffset = 0 Then
        
            If X <> lastX Then
                m_MouseHeldCounter = m_MouseHeldCounter + 1
                lastX = X
                
                If m_MouseHeldCounter >= 3 Then
                    m_MoveOffset = X - m_lastHoveredGroup.X
                    m_GapPosition = m_lastHoveredGroup.X
                End If
            Else
                m_MouseHeldCounter = 0
            End If
        Else
        
            If X - m_MoveOffset > -1 And _
                Me.ScaleWidth - m_TaskBandImage.width > (X / m_aspectRatio) - m_MoveOffset Then

                m_lastHoveredGroup.X = X - m_MoveOffset
            End If
            
            ShuffleProcessesIfNeeded
            
            
            DrawTaskList
        End If
        
        Exit Sub
    End If
    
    If IsWindowVisible(m_SubMenu.hWnd) = APIFALSE Then
        newUserSelection = True
    End If
    
    newRolloverIndex = RTS2(X, 62) - 1
    'Check new rollover is within subscript bounderies
    If newRolloverIndex < 0 Or (newRolloverIndex + 1) > m_TaskList.Processes.Count Then
        Debug.Print "Impossible newRolloverIndex Value (Dropping Event; Form_MouseMove) "
        Exit Sub
    End If
    
    m_LastMouseButton = Button
    
    If (newRolloverIndex <> m_RolloverIndex) Or newUserSelection = True Then

        If m_ignoreMouseMove Then
        
            If newRolloverIndex <> m_RolloverIndex Then
                m_ignoreMouseMove = False
            Else
                Exit Sub
            End If
        End If

        m_RolloverIndex = newRolloverIndex
        Set m_lastHoveredGroup = m_VisibleList(m_RolloverIndex)

        If newUserSelection Then
            timHideSubMenuDelay.Enabled = False
            
            m_ignoreClick = False
            
            If IsWindowVisible(m_SubMenu.hWnd) = APIFALSE Then
                timPopupDelay.Enabled = ApplicationOptions.AutoClick
                
            End If
        Else
            If timPopupDelay.Enabled = False Then
                timPopupNextDelay.Enabled = True
            End If
        End If
    End If
    
    DrawTaskList
    'Debug.Print vbCrLf & vbCrLf
End Sub

Private Sub MouseEvent(Button As Integer, X As Integer, Y As Integer)

Dim mnuCmd As Long

    'Debug.Print "MouseEvent Triggered!!"
    
    If Not m_MouseDown Then
        'Exit Sub
    End If
    
    m_MouseDown = False
    m_MouseHeldCounter = 0
    
    If m_MoveOffset <> 0 Then
        Debug.Print "Blocking event for drag-stuff; " & m_MoveOffset
        
        m_MoveOffset = 0
    
        If Not m_lastHoveredGroup Is Nothing Then
            m_lastHoveredGroup.X = RoundToSignificance(m_lastHoveredGroup.X + (m_TaskBandImage.width), m_TaskBandImage.width + 2)
            Debug.Print "Form_MouseUp; " & m_lastHoveredGroup.X
            
            If m_lastHoveredGroup.X < 0 Then
                m_lastHoveredGroup.X = 0
            End If
            
            ReAlignVisibleProcesses
        End If
        
        DrawTaskList
        Exit Sub
    End If
    
    If Button = vbLeftButton And m_ignoreClick Then
        m_ignoreClick = False
        m_LastMouseButton = -1
        
        Debug.Print "Form_MouseUp; Exited; Fake Call"
        Exit Sub
    End If
    
    'Function guard
    If m_lastHoveredGroup Is Nothing Then
        Debug.Print "Form_MouseUp; Exited; Invalid ParameterGuard"
    
        Exit Sub
    End If
    
    If Button = vbRightButton Then
        timPopupDelay.Enabled = False
        m_currentJumpLists = m_lastHoveredGroup.GetJumpLists
        
        Set m_GroupMenu = BuildMenuWithJumpList(m_currentJumpLists)
        m_GroupMenu.EditItem 5, IIf(m_lastHoveredGroup.Pinned, "Unpin", "Pin")

        If m_SubMenu.Visible Then
            ShowWindow m_SubMenu.hWnd, SW_HIDE
        End If
        
        mnuCmd = m_GroupMenu.ShowMenu(Me.hWnd)
        
        Select Case mnuCmd
        
        Case 1
            m_lastHoveredGroup.MinimizeAllWindows
        
        Case 2
            m_lastHoveredGroup.RestoreAllWindows
            
        Case 4
            m_lastHoveredGroup.RequestCloseAllWindows
            
        Case 5
            If m_lastHoveredGroup.Pinned = False Then
                If m_lastHoveredGroup.Path <> vbNullString Then
                    m_lastHoveredGroup.Pinned = True
                    
                    AddToPinnedList m_lastHoveredGroup
                Else
                    LogError -2, "Invalid path; Null", "TaskBar"
                End If
            Else
                m_lastHoveredGroup.Pinned = False
                
                SortPinnedList
                DumpOptions
                
                'RemoveFromPinnedList m_lastHoveredGroup.Path
            End If
        
        Case 6
            Unload Me
            Exit Sub
            
        Case Else
            If mnuCmd >= 7 Then
                If IsArrayInitialized(m_currentJumpLists) Then
                    ShellEx CStr(m_currentJumpLists(mnuCmd - 7))
                End If
            End If

        End Select
            
            'RaiseEvent requestResumeChecker
        'End If
        
    ElseIf Button = vbLeftButton Then
        Debug.Print "Button = LEFT BUTTON!"
        
        If ShiftKey = False And m_lastHoveredGroup.Running Then
            If m_lastHoveredGroup.WindowCount = 1 Then
                
                ShowWindow frmSubMenu.hWnd, SW_HIDE
                HandleWindow m_lastHoveredGroup.Window(1).hWnd
            Else
                
            
                If m_SubMenu.WindowList Is m_lastHoveredGroup.Window And m_SubMenu.Visible Then
                    Debug.Print "Showing the same, so closing!"
                    
                    ShowWindow frmSubMenu.hWnd, SW_HIDE
                    Set m_SubMenu.WindowList = Nothing
                    
                Else
                    timPopupDelay.Enabled = False
                
                    m_userInvoked = True
                    
                    timPopupNextDelay.Enabled = False
                    
                    'Debug.Print "TRIGGA!"
                    
                    'Debug.Print "MouseEvent--TriggerSubMenu"
                    TriggerSubMenu
                End If
            End If
        Else
            If Is64bit Then
                Dim win64Token As Win64FSToken: Set win64Token = New Win64FSToken
                ShellExecute Me.hWnd, "Open", m_lastHoveredGroup.Path, m_lastHoveredGroup.Arguments, "", SW_SHOWDEFAULT
                win64Token.EnableFS
            Else
                ShellExecute Me.hWnd, "Open", m_lastHoveredGroup.Path, m_lastHoveredGroup.Arguments, "", SW_SHOWDEFAULT
            End If
            
            HideSubMenu
        End If
        
        m_LastMouseButton = 0
        DrawTaskList
        
    End If
    'Me.Width = Me.Width + 50
    'm_TaskbarBackgroundBitmap.CreateFromSize Me.ScaleWidth, 40
    'm_TaskbarBackgroundGraphics.FromImage m_TaskbarBackgroundBitmap.Image
End Sub

Sub HideSubMenu()

    m_SubMenu.height = 0
    m_SubMenu.Hide

End Sub

Private Sub ParsePinnedList()
    On Error GoTo Handler
    
Dim newProcess As Process
Dim processIndex As Long
Dim strPath As String
Dim strArguments As String

Dim lastPinnedApp As Long: lastPinnedApp = GetLastPinnedApp()

    For processIndex = 0 To lastPinnedApp
        strPath = ApplicationOptions.PinnedApplications(processIndex).szPath
        strArguments = ApplicationOptions.PinnedApplications(processIndex).szArguments

        If FileExists(strPath, False) Then
            Set newProcess = New Process

            newProcess.Constructor 0, strPath
            newProcess.CreateIconFromPath
            newProcess.Arguments = strArguments
            
            newProcess.Pinned = True
                    
            m_TaskList.Processes.Add newProcess, newProcess.GetKey
        End If
    Next
    
    Exit Sub
Handler:
End Sub

'Private Sub RemoveFromPinnedList(ByVal strPath As String)
'    On Error GoTo Handler
'
'Dim newProcess As Process
'Dim processIndex As Long
'Dim newProcessIndex As Long
'
'Dim thisProcessPath As String
'Dim emptyByteStringArray() As ByteString
'
'    For processIndex = LBound(ApplicationOptions.PinnedApplications) To UBound(ApplicationOptions.PinnedApplications)
'        thisProcessPath = ApplicationOptions.PinnedApplications(processIndex).Value
'
'        If strPath = thisProcessPath Then
'            If processIndex < UBound(ApplicationOptions.PinnedApplications) Then
'                For newProcessIndex = processIndex To UBound(ApplicationOptions.PinnedApplications) - 1
'                    ApplicationOptions.PinnedApplications(newProcessIndex).Value = ApplicationOptions.PinnedApplications(newProcessIndex + 1).Value
'                Next
'            End If
'
'            If processIndex > 0 Then
'                ReDim Preserve ApplicationOptions.PinnedApplications(UBound(ApplicationOptions.PinnedApplications) - 1)
'            Else
'                ApplicationOptions.PinnedApplications = emptyByteStringArray
'            End If
'
'            Exit For
'        End If
'    Next
'
'    Exit Sub
'Handler:
'    Debug.Print "RemoveFromPinnedList; " & Err.Description
'
'End Sub

Private Sub AddToPinnedList(ByRef thisProcess As Process)
    On Error Resume Next
    
    ReDim Preserve ApplicationOptions.PinnedApplications(GetLastPinnedApp() + 1)
    
    With ApplicationOptions.PinnedApplications(GetLastPinnedApp())
        .szPath = thisProcess.Path
        .szArguments = thisProcess.Arguments
    End With
    
    DumpOptions
    Exit Sub
Handler:
    'AddToPinnedList strPath, True

End Sub

Private Sub ReInitSurface()
    Debug.Print "ReInitSurface"
    
    On Error GoTo Handler
    
    m_TaskbarBackgroundBitmap.Dispose
    m_TaskbarBackgroundGraphics.Dispose

    If winSize.cy < 40 Then
        m_aspectRatio = 1.4
    Else
        m_aspectRatio = 1
    End If
        
    m_TaskbarBackgroundBitmap.CreateFromSizeFormat winSize.cx * m_aspectRatio, winSize.cy * m_aspectRatio, PixelFormat.Format32bppArgb
    m_TaskbarBackgroundGraphics.FromImage m_TaskbarBackgroundBitmap.Image
    
    m_gfx.FromHDC Me.hdc

    Debug.Print "EndInitSurface"

    Exit Sub
Handler:
    LogError Err.number, "ReInitSurface(); " & Err.Description, "winTaskBar"
End Sub

Public Sub MakeTrans()
    curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW Or WS_CHILD
    
    'update layer window stuff (that will blend and show the GDI+ text)
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = Me.ScaleWidth
    winSize.cy = Me.ScaleHeight

    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    'now the text will be shown
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, Me.hdc, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)

End Sub

Private Sub Form_Load()
    RegisterShellHookWindow Me.hWnd
    'Recieve shell events
    m_Shell_Hook_Msg_ID = RegisterWindowMessage("SHELLHOOK")
    
    Load m_SubMenu
    
    ReTrackMouse
    
    frmZOrderKeeper.Show
    MakeTrans
       
    'im.FromFile App.Path & "\bottom.png"
    m_TaskBandImage.FromFile App.Path & "\resources\taskbandbutton.png"
    m_TaskGroupImage.FromFile App.Path & "\resources\button_group.png"
    
    m_TaskBandEmpty.FromFile App.Path & "\resources\light.png"
    m_TaskBandLight.FromFile App.Path & "\resources\light.png"
    
    m_TaskBandGraphics.FromImage m_TaskBandEmpty
    m_TaskBandGraphics.Clear
    m_TaskBandGraphics.DrawImage m_TaskBandLight, 0, 0, m_TaskBandLight.width, m_TaskBandLight.height
    
    'MakeAttributes

    ReInitSurface
    
    Call HookWindow(Me.hWnd, Me)
End Sub

'Private Sub MakeAttributes()
'
'    Dim theMap(255) As ColorMap
'    Dim mapIndex As Long
'
'    For mapIndex = 0 To 255
'        theMap(mapIndex).oldColor = ColorARGB(mapIndex, 0, 0, 0)
'        theMap(mapIndex).newColor = ColorARGB(255, 255, 0, 0)
'    Next
'
'    'm_imageAttr.SetColorMatrix clrMatrix
'    'm_imageAttr.SetImageAttributesRemapTable theMap
'
'End Sub

Private Sub Form_DragDropFile(sFile As String)
    
Dim thisLink As ShellLink
    thisLink = GetShortcut(CStr(sFile))

    If LCase$(Right$(GetEXEPathFromQuote(thisLink.szPath), 4)) <> ".exe" Then
        thisLink.szArguments = thisLink.szPath
        thisLink.szPath = Environ$("windir") & "\explorer.exe"
    End If
    
    If FileExists(thisLink.szPath) Then
        Dim newProcess As New Process

        newProcess.Constructor 0, thisLink.szPath
        newProcess.CreateIconFromPath
        newProcess.Arguments = thisLink.szArguments
        
        If Not Exists(m_TaskList.Processes, newProcess.GetKey) Then
            newProcess.Pinned = True
        
            m_TaskList.Processes.Add newProcess, newProcess.GetKey
            AddToPinnedList newProcess
        End If
        'MsgBox "Debug #302:: Added (" & sExePath & ")", vbCritical, "Debug Error"
    End If

End Sub

Private Sub Form_Resize()
    'Debug.Print "frmTaskBar_Resize(); " & Me.ScaleWidth & " - "; Me.ScaleHeight
    On Error Resume Next
    
    HideSubMenu
    
    GetWindowRect Me.hWnd, m_Rect
    winSize.cx = m_Rect.Right - m_Rect.Left
    winSize.cy = m_Rect.Bottom - m_Rect.Top

    ReInitSurface
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnhookWindow(Me.hWnd)

    Unload frmWindowChecker
    Unload m_SubMenu
    
    Set m_TaskList = Nothing
    Set m_GroupMenu = Nothing
    
    Set m_SolidBrush = Nothing
   ' Clear up the Graphics object (also can
   ' use 'Set m_gfx = Nothing'):
   
    m_TaskbarBackgroundGraphics.Dispose
    ' Clear up the Image object (also can use
    ' 'Set m_gfx = Nothing'):
    m_TaskbarBackgroundBitmap.Dispose
   
    m_gfx.Dispose
    
    m_TaskBandImage.Dispose
    m_TaskGroupImage.Dispose

    DisposeGDIIfLast
    TaskbarHelper.MakeTaskbarTransparent 255
    
    Unload frmZOrderKeeper
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, iMsg As Long, wParam As Long, lParam As Long) As Long
    On Error GoTo Handler

Dim theWindow As Window

    If hWnd <> Me.hWnd Then
         ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
        Exit Function
    End If

    If iMsg = WM_WINDOWPOSCHANGED Then
        'Debug.Print "WM_WINDOWPOSCHANGED"
        
        
    
        CopyMemory m_currentPosition, ByVal lParam, LenB(m_currentPosition)
        GetWindowRect Me.hWnd, m_Rect
        
        If Not m_currentPosition.Flags And SWP_NOMOVE Then
            'Position changed
            HideSubMenu
        End If
        
        'Allow VB to process it
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
            
    ElseIf iMsg = WM_DROPFILES Then
    
        Dim hFilesInfo As Long
        Dim szFileName As String
        Dim wTotalFiles As Long
        Dim wIndex As Long
        
        hFilesInfo = wParam
        wTotalFiles = DragQueryFileW(hFilesInfo, &HFFFF, ByVal 0&, 0)
    
        For wIndex = 0 To wTotalFiles
            szFileName = Space$(1024)
            
            If Not DragQueryFileW(hFilesInfo, wIndex, StrPtr(szFileName), Len(szFileName)) = 0 Then
                Form_DragDropFile szFileName
            End If
        Next wIndex
        
        DragFinish hFilesInfo
            
    ElseIf iMsg = WM_LBUTTONUP Then
        MouseEvent vbLeftButton, LoWord(lParam), HiWord(lParam)
         'Allow DefWndProc to Handle also
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
        
    ElseIf iMsg = WM_RBUTTONUP Then
        MouseEvent vbRightButton, LoWord(lParam), HiWord(lParam)
         'Allow DefWndProc to Handle also
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
    
    ElseIf iMsg = WM_MOUSEMOVE Then

        MouseMove MouseButtonState(wParam), ShiftState(wParam), LoWord(lParam), HiWord(lParam)
        
    ElseIf iMsg = WM_ACTIVATE Then
    
        If wParam = WA_INACTIVE Then

            If Not hWndBelongToUs(lParam) Then
                Debug.Print ":D @1"
                
                timPopupDelay.Enabled = False
                timAutoCloseMenu.Enabled = False
                
                m_ignoreMouseMove = True
                HideSubMenu
            End If
        End If
        
         'Allow DefWndProc to Handle also
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)

    ElseIf iMsg = WM_MOUSELEAVE Then
        m_ignoreMouseMove = False
        m_isTracking = False
    
        If IsMouseInsideMe = False Then
            
            If m_userInvoked = False Then
                timHideSubMenuDelay.Enabled = True
            End If
                
            timPopupDelay.Enabled = False
            timEnumerateTasks_Timer
            frmZOrderKeeper.timZOrderChecker.Enabled = True
        Else
        
            timTestMouse.Interval = 100
            m_notifiedLeft = False
        End If
        
    ElseIf iMsg = m_Shell_Hook_Msg_ID Then
        'Debug.Print "Shell Hook Message!"
    
        If wParam = HSHELL_REDRAW Then
            'Debug.Print lParam
            Set theWindow = m_TaskList.GetWindowByHWND(lParam)
            Set m_targetCaptureWindow = theWindow
            
            timReCaptureImage.Enabled = True
            
        ElseIf wParam = HSHELL_WINDOWDESTROYED Then
            If IsWindowVisible(m_SubMenu.hWnd) = APITRUE Then
                Debug.Print "Yes, visible"
                
                If Not m_lastHoveredGroup Is Nothing Then
                    Set theWindow = m_TaskList.GetWindowByHWND(lParam)
                
                    If Not theWindow Is Nothing Then
                        Debug.Print "HSHELL_WINDOWDESTROYED;;RedrawSubMenu"

                        m_TaskList.RemoveWindow theWindow
                        
                        If m_lastHoveredGroup.WindowCount > 0 Then
                            RedrawSubMenu
                        Else
                            HideSubMenu
                        End If
                        
                        ReAlignVisibleProcesses
                    End If
                End If
            Else
                m_TaskList.RemoveWindow m_TaskList.GetWindowByHWND(lParam)
            End If
            
            DrawTaskList
            
        ElseIf wParam = HSHELL_FLASH Then
            
            Set theWindow = m_TaskList.GetWindowByHWND(lParam)
            If Not theWindow Is Nothing Then
                If UpdateThumbnail(theWindow) Then
                
                    RedrawSubMenuIfTargetIsVisible lParam
                End If
                
                theWindow.Flashing = True
            End If
        ElseIf wParam = HSHELL_WINDOWCREATED Then
            
            If IsVisibleToTaskBar(lParam) Then
                m_TaskList.AddWindowByHwnd lParam
                        
                If IsWindowVisible(m_SubMenu.hWnd) = APITRUE Then
                    If Not m_lastHoveredGroup Is Nothing Then
                        Set theWindow = m_TaskList.GetWindowByHWND(lParam)
                    
                        If Not theWindow Is Nothing Then
                            Debug.Print "HSHELL_WINDOWCREATED;;RedrawSubMenu"
        
                            'm_TaskList.RemoveWindow theWindow
                            RedrawSubMenu
                            ReAlignVisibleProcesses
                        End If
                    End If
                End If
            End If
            
            DrawTaskList
            
        ElseIf wParam = HSHELL_WINDOWACTIVATED Then
            
            'Snapshot the window being unactivated
            
            Set theWindow = m_TaskList.GetWindowByHWND(g_hwndForeGroundWindow)
            UpdateThumbnail theWindow
            
            Set theWindow = m_TaskList.GetWindowByHWND(lParam)
            UpdateThumbnail theWindow
            
            g_hwndForeGroundWindow = lParam
            m_TaskList.AddWindowByHwnd lParam
            
            DrawTaskList
        End If
    Else
         ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
    End If

    Exit Function
Handler:
    LogError Err.number, "WindowProc(" & iMsg & "," & wParam & "," & lParam & "); " & Err.Description, "winTaskBar"

    ' Just allow VB to process it
    IHookSink_WindowProc = _
       InvokeWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Function UpdateThumbnail(ByRef cWnd As Window) As Boolean
    UpdateThumbnail = True

    If Not cWnd Is Nothing Then
        cWnd.UpdateWindowText
    
        If Not cWnd.isMinimized Then
            cWnd.UpdateImage
            'RepaintWindow cWnd.hWnd
            
            UpdateThumbnail = True
        End If
    End If

End Function

Private Sub m_SubMenu_onClicked(targetWindow As Window)
    'Debug.Print "m_SubMenu_onClicked"
    
    targetWindow.Flashing = False
    
    HideSubMenu
    HandleWindow targetWindow.hWnd
End Sub

Private Sub m_SubMenu_onClosed(targetWindow As Window)
    m_hWndToTestIfValid = targetWindow.hWnd
    
    'PostMessage targetWindow.hWnd, WM_CLOSE, 0&, 0&
    PostMessage targetWindow.hWnd, ByVal WM_SYSCOMMAND, ByVal SC_CLOSE, 0
    'timAutoCloseMenu.Enabled = True
End Sub

Private Sub m_SubMenu_onDeactivated(hWnd As Long)
    'Debug.Print "m_SubMenu_onDeactivated"

Dim ExceptionHwnd As Long
    If ApplicationOptions.ViOrb Then
        ExceptionHwnd = frmFader.hWnd
    End If
    
    If hWndBelongToUs(hWnd, ExceptionHwnd) = False Then
    
        HideSubMenu
        Set m_SubMenu.WindowList = Nothing
        
        timHideSubMenuDelay.Enabled = False
        timPopupDelay.Enabled = False
        
        If hWnd = Me.hWnd Then
            m_ignoreClick = True
        End If
    End If
End Sub

Private Sub m_SubMenu_onMouseOut()
    If m_userInvoked = False Then
        timHideSubMenuDelay.Enabled = True
    End If
End Sub

Private Sub m_SubMenu_onMouseOver()
    'Debug.Print "Disabling HideMenu"
    timHideSubMenuDelay.Enabled = False
    timPopupDelay.Enabled = False
    timPopupNextDelay.Enabled = False
End Sub

Private Sub m_SubMenu_onRightClicked(targetWindow As Window)
Dim theMenuHandle As Long
Dim thisMenu As clsMenu
Dim sysCmdID As Long
Dim currentCursorPosition As win.POINTL

    If targetWindow.IsHung Then Exit Sub

    GetCursorPos currentCursorPosition

    theMenuHandle = GetSystemMenu(targetWindow.hWnd, 0)
    Set thisMenu = CreateSystemMenu(theMenuHandle, targetWindow.WindowState)

    SetForegroundWindow m_SubMenu.hWnd
    sysCmdID = thisMenu.ShowMenu(m_SubMenu.hWnd)

    Select Case sysCmdID
    
    Case SC_RESTORE
        ShowWindow targetWindow.hWnd, SW_SHOWNORMAL
    
    Case SC_MINIMIZE
        ShowWindow targetWindow.hWnd, SW_SHOWMINIMIZED
    
    Case SC_MAXIMIZE
        ShowWindow targetWindow.hWnd, SW_SHOWMAXIMIZED
        
    Case SC_CLOSE
        PostMessage targetWindow.hWnd, WM_CLOSE, 0&, 0&
        
    Case Else
        SendMessage targetWindow.hWnd, WM_SYSCOMMAND, ByVal sysCmdID, ByVal MAKELPARAM(currentCursorPosition.X, currentCursorPosition.Y)

    
    End Select
    
    m_SubMenu.ResetRollover
End Sub

Private Sub timAnimate_Timer()
    On Error GoTo Handler

Dim actualSubMenuLeft As Long
Dim actualSubMenuWidth As Long
Dim actualSubMenuTop As Long
Dim actualSubMenuBottom As Long

    'Debug.Print m_SubMenu_finalLeft & ":" & m_SubMenuLayout.Left
    actualSubMenuLeft = m_SubMenuLayout.Left
    actualSubMenuWidth = m_SubMenuLayout.width
    actualSubMenuTop = m_SubMenuLayout.Top
    actualSubMenuBottom = m_SubMenuLayout.height

    'Debug.Print "Comparison; " & actualSubMenuLeft & ";" & m_SubMenu_finalLeft

    If actualSubMenuLeft <> m_SubMenu_finalLeft Or _
        actualSubMenuTop <> m_SubMenu_finalTop Or _
        actualSubMenuBottom <> m_SubMenu_finalBottom Or _
        actualSubMenuWidth <> m_SubMenu_finalRight Then
        
        'Debug.Print "timAnimate - Phase01"
        
        'MoveWindow m_SubMenu.hWnd, _
                        (m_SubMenu.Left / Screen.TwipsPerPixelX) + m_moveLeftVelocity, _
                        (m_SubMenu.Top / Screen.TwipsPerPixelY) + m_moveTopVelocity, _
                        (m_SubMenu.Width / Screen.TwipsPerPixelX) + m_shrinkVelocity, _
                        (m_SubMenu.Height / Screen.TwipsPerPixelY) + m_shrinkHeightVelocity, False
                        
        'm_SubMenu.Move m_SubMenu.Left + m_moveLeftVelocity * Screen.TwipsPerPixelX, _
                       m_SubMenu.Top + m_moveTopVelocity * Screen.TwipsPerPixelY, _
                       m_SubMenu.Width + m_shrinkVelocity * Screen.TwipsPerPixelX, _
                       m_SubMenu.Height + m_shrinkHeightVelocity * Screen.TwipsPerPixelY
        
        With m_SubMenuLayout
            'Debug.Print .Top & ":" & .Left & ":" & .Height & ":" & .Width
        
            .Left = .Left + m_moveLeftVelocity
            .Top = .Top + m_moveTopVelocity
            .width = .width + m_moveRightVelocity
            .height = .height + m_moveBottomVelocity
        End With
         
        'm_SubMenu.ReInitSurface2
        m_SubMenu.DrawBorder RECTFtoL(m_SubMenuLayout)
        'm_SubMenu.DrawBorder CreateRect(0, 0, 100, 100)
        
        'm_SubMenu.Refresh
    Else
    
        m_SubMenu.AllowListRedraw = True
        If m_SubMenu.WindowList.Count > 0 Then
        
            m_SubMenu.DrawList RECTFtoL(m_SubMenuLayout)
        Else
            HideSubMenu
        End If
        
        'Debug.Print "Finished Animation!"
        timAnimate.Enabled = False
        Exit Sub
    End If

Exit Sub

Handler:
    'Debug.Print "timAnimate():" & Err.Description
    LogError Err.number, "AnimateGlide(); " & Err.Description, "winTaskBar"

    timAnimate.Enabled = False
    
    'Debug.Print "timAnimate::TriggerSubMenuWithAnimation"
    TriggerSubMenuWithAnimation
End Sub

Private Sub timAutoCloseMenu_Timer()
    If IsWindow(m_hWndToTestIfValid) Then
        ForceFocus m_hWndToTestIfValid
        HideSubMenu
    End If
    
    timAutoCloseMenu.Enabled = False
End Sub

Sub timEnumerateTasks_Timer()
    On Error GoTo Handler

Dim theWindow As Window
Dim newForeGroundWindow As Long

    newForeGroundWindow = GetForegroundWindow

    If newForeGroundWindow <> Me.hWnd And _
       newForeGroundWindow <> m_SubMenu.hWnd And _
       newForeGroundWindow <> TaskbarHelper.g_ReBarWindow32Hwnd And _
       newForeGroundWindow <> TaskbarHelper.g_TaskBarHwnd Then
       
        g_hwndForeGroundWindow = newForeGroundWindow
        
        'Debug.Print "Fetching Window!"
        Set theWindow = m_TaskList.GetWindowByHWND(g_hwndForeGroundWindow)
        If Not theWindow Is Nothing Then
            theWindow.Flashing = False
            
        End If
    End If

    m_TaskList.DeleteDeadHandles
    m_TaskList.UpdateFlashStatusOfEachProcess
    
    TaskbarHelper.EnumWindowsAsTaskList m_TaskList
    
    SetPositionNewProcesses
    CopyAndSortVisibleList
    
    If m_MoveOffset = 0 Then ReAlignVisibleProcesses
    
    DrawTaskList
    
    Exit Sub
Handler:
    LogError Err.number, "EnumerateTasks(); " & Err.Description, "winTaskBar"

End Sub

Sub ReAlignVisibleProcesses()

Dim processIndex As Long
Dim processIndex2 As Long

Dim thisProcess As Process
Dim beforeThisProcess As Process

Dim visibleListCount As Long

Dim secondProcessIndex As Long
Dim lastProcessIndex As Long

    lastProcessIndex = UBound(m_VisibleList)

    If IsArrayInitialized(m_VisibleList) Then
        If Not m_VisibleList(0) Is Nothing Then
            If m_VisibleList(0).X <> 0 Then
                For processIndex2 = processIndex To lastProcessIndex
                    If Not m_VisibleList(processIndex2) Is Nothing Then
                        
                        m_VisibleList(processIndex2).X = m_VisibleList(processIndex2).X - (m_TaskBandImage.width + 2)
                    End If
                Next
            End If
        End If
        
        secondProcessIndex = LBound(m_VisibleList) + 1
    
        For processIndex = secondProcessIndex To lastProcessIndex
            Set thisProcess = m_VisibleList(processIndex)
            Set beforeThisProcess = m_VisibleList(processIndex - 1)
            
            'Debug.Print "ReAlignVisibleProcesses; " & (m_TaskBandImage.width + 2)
            
            If Not thisProcess Is Nothing And Not beforeThisProcess Is Nothing Then
                If thisProcess.X - beforeThisProcess.X <> (m_TaskBandImage.width + 2) Then
                    
                    visibleListCount = UBound(m_VisibleList)
                    
                    For processIndex2 = processIndex To visibleListCount
                        If Not m_VisibleList(processIndex2) Is Nothing Then
                            
                            'Debug.Print (m_VisibleList(processIndex2).x - m_VisibleList(processIndex2 - 1).x) / (m_TaskBandImage.width + 2)
                            m_VisibleList(processIndex2).X = m_VisibleList(processIndex2).X - (m_TaskBandImage.width + 2) * _
                                    ((m_VisibleList(processIndex2).X - m_VisibleList(processIndex2 - 1).X) / (m_TaskBandImage.width + 2) - 1)
                        End If
                    Next
                End If
            End If
        Next
    End If

End Sub

Sub CopyAndSortVisibleList()

Dim processTemp As Process
Dim processIndex As Long

Dim lngX As Long
Dim lngY As Long

Dim firstProcessIndex As Long
Dim secondToLastProcessIndex As Long

    If m_TaskList.Processes.Count > 0 Then
        ReDim m_VisibleList(m_TaskList.Processes.Count - 1)
        processIndex = 1
        
        For Each processTemp In m_TaskList.Processes
            Set m_VisibleList(processIndex - 1) = m_TaskList.Processes(processIndex)
            processIndex = processIndex + 1
        Next
        
        firstProcessIndex = LBound(m_VisibleList)
        secondToLastProcessIndex = UBound(m_VisibleList) - 1
    
        For lngX = firstProcessIndex To secondToLastProcessIndex
            For lngY = firstProcessIndex To secondToLastProcessIndex
                If m_VisibleList(lngY).X > m_VisibleList(lngY + 1).X Then
                    ' exchange the items
                    Set processTemp = m_VisibleList(lngY)
                    
                    Set m_VisibleList(lngY) = m_VisibleList(lngY + 1)
                    Set m_VisibleList(lngY + 1) = processTemp
                    
                End If
            Next
        Next
    End If

End Sub

Sub ShuffleProcessesIfNeeded()

Dim problemMaticProcess As Process

    If Not m_lastHoveredGroup Is Nothing And m_MoveOffset <> 0 Then
        'Debug.Print m_lastHoveredGroup
    
        If m_lastHoveredGroup.X < m_GapPosition - 31 Then
            Set problemMaticProcess = FindProcessOnXAxisBeforeGap(m_lastHoveredGroup.X)
            If Not problemMaticProcess Is Nothing Then
                
                Debug.Print "Left; " & problemMaticProcess.Path
                'Sleep 2000
                
                If problemMaticProcess.X + (m_TaskBandImage.width + 2) = m_GapPosition Then
                    m_GapPosition = problemMaticProcess.X
                    problemMaticProcess.X = problemMaticProcess.X + (m_TaskBandImage.width + 2)
                End If
                
                CopyAndSortVisibleList
            End If
        ElseIf m_lastHoveredGroup.X > m_GapPosition + 31 Then
            Set problemMaticProcess = FindProcessOnXAxisAfterGap(m_lastHoveredGroup.X)
            If Not problemMaticProcess Is Nothing Then
            
                Debug.Print "Right; " & m_lastHoveredGroup.X - m_GapPosition
                'Sleep 2000
                
                'm_GapPosition = problemMaticProcess.x
                'problemMaticProcess.x = problemMaticProcess.x - (m_TaskBandImage.width + 2)
                If problemMaticProcess.X - (m_TaskBandImage.width + 2) = m_GapPosition Then
                    m_GapPosition = problemMaticProcess.X
                    problemMaticProcess.X = problemMaticProcess.X - (m_TaskBandImage.width + 2)
                End If
                
                CopyAndSortVisibleList
            End If
        End If
    End If
    
End Sub

Function IsBeingShown(hWnd As Long) As Boolean
    IsBeingShown = False

    If m_lastHoveredGroup Is Nothing Then
        Exit Function
    End If
    
Dim testWindow As Window

    Set testWindow = m_lastHoveredGroup.GetWindowByHWND(hWnd)
    
    If m_SubMenu.Visible Then
        If Not testWindow Is Nothing Then
            IsBeingShown = True
        End If
    End If

End Function

Sub RedrawSubMenuIfTargetIsVisible(hWnd As Long)
    If IsBeingShown(hWnd) Then
        Debug.Print "HSHELL_FLASH;;RedrawSubMenu"
        RedrawSubMenu
    End If
End Sub

Sub RedrawSubMenu()
    Debug.Print "RedrawSubMenu"
    If m_SubMenu.Visible Then TriggerSubMenuWithAnimation
End Sub

Sub DrawTaskList()
    On Error GoTo Handler

Dim thisProcess As Process
Dim rolledOverVisible As Boolean

    m_TaskbarBackgroundGraphics.Clear

    For Each thisProcess In m_TaskList.Processes
        Debug.Print thisProcess.X, thisProcess.Path
    
        If Not thisProcess Is m_lastHoveredGroup Then
            DrawProcess thisProcess
        Else
            rolledOverVisible = True
        End If
    Next
    
    If Not m_lastHoveredGroup Is Nothing And _
        rolledOverVisible Then
        
        DrawProcess m_lastHoveredGroup
    End If
           
    m_gfx.Clear vbBlack
    m_gfx.DrawImage _
       m_TaskbarBackgroundBitmap.Image, 0, 0, CSng(winSize.cx), CSng(winSize.cy)
       
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, Me.hdc, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
    Exit Sub
Handler:
    LogError Err.number, "DrawTaskList(); " & Err.Description, "winTaskBar"

End Sub

Private Function DrawProcess(ByRef thisProcess As Process)

Dim thisIndex As Long
Dim BandType As Long
Dim IsRolledOver As Boolean
Dim IsFlashing As Boolean
Dim ProcessPadding As Long

    If thisProcess.IconIsValid Then
        IsRolledOver = False
        IsFlashing = False
        
        'm_isTracking cause of rolled out bug dragging
        If thisProcess Is m_lastHoveredGroup And m_isTracking = True Then
            IsRolledOver = True
        End If
        
        IsFlashing = thisProcess.Flashing
        If IsFlashing And Not thisProcess.GetWindowByHWND(m_hWndToTestIfValid) Is Nothing Then
            timAutoCloseMenu_Timer
        End If
        'thisProcess.Flashing = False
    
        BandType = BAND_NORMAL
        
        If thisProcess.PinnedAndClosed Then
            BandType = BAND_PINNED
            
            If m_LastMouseButton = vbLeftButton Then
                If thisProcess Is m_lastHoveredGroup And IsMouseLeftButtonDown Then
                    BandType = BAND_PINNED_SELECTED
                End If
            Else
                If IsRolledOver Then
                    BandType = BAND_PINNED_HOVER
                End If
            End If
        Else
            If m_LastMouseButton = vbLeftButton Then
                ProcessPadding = IIf(IsRolledOver, 1, 0)
            
                If thisProcess Is m_TaskList.ActiveProcess Then
                    If IsRolledOver Then
                        BandType = BAND_LEFT_BUTTON_DOWN
                    Else
                        BandType = BAND_ACTIVE
                    End If
                    
                    'BandType = IIf(IsRolledOver, BAND_ACTIVE_HOVER, BAND_ACTIVE)
                Else
                    If IsRolledOver Then
                        BandType = BAND_LEFT_BUTTON_DOWN
                    Else
                        BandType = BAND_NORMAL
                    End If
                
                    'BandType = IIf(IsRolledOver, BAND_LEFT_BUTTON_DOWN, BAND_NORMAL)
                End If
            Else
                If thisProcess Is m_TaskList.ActiveProcess Or IsFlashing Then
                    If IsFlashing And Not IsRolledOver Then
                        BandType = BAND_NOTIFY
                    ElseIf IsRolledOver Then
                        BandType = BAND_HOVER_UNCLICKED
                    Else
                        BandType = BAND_ACTIVE
                    End If
                
                    'BandType = IIf(IsRolledOver, BAND_ACTIVE_HOVER, BAND_ACTIVE)
                Else
                    If IsRolledOver Then
                        BandType = BAND_HOVER_UNCLICKED
                    Else
                        BandType = BAND_NORMAL
                    End If
                
                    'BandType = IIf(IsRolledOver, BAND_HOVER_UNCLICKED, BAND_NORMAL)
                End If
            End If
        End If
        
        If BandType > BAND_NOTHING Then
        
            'If MouseX > -1 Then
            '
            '    m_TaskBandGraphics.Clear RGB(255, 180, 0)
            '    m_TaskBandGraphics.DrawImage m_TaskBandLight, MouseX - (m_TaskBandLight.width / 2), 0, m_TaskBandLight.width, m_TaskBandLight.height
            '
            '    m_TaskbarBackgroundGraphics.DrawImage m_TaskBandEmpty, thisProcess.X, 0, m_TaskBandLight.width, m_TaskBandLight.height
            'End If
        
            If thisProcess.WindowCount < 2 Then
                m_TaskbarBackgroundGraphics.DrawImageRect _
                    m_TaskBandImage, thisProcess.X, 0, m_TaskBandImage.width, 40, 0, BandType
                    
            ElseIf thisProcess.WindowCount = 2 Then
                m_TaskbarBackgroundGraphics.DrawImageStretchAttrF _
                    m_TaskBandImage, CreateRectF(thisProcess.X, 0, 40, m_TaskBandImage.width - 4), 0, BandType, m_TaskBandImage.width, 40, UnitPixel, 0, 0, 0
                
                m_TaskbarBackgroundGraphics.DrawImageRect _
                    m_TaskGroupImage, thisProcess.X + (m_TaskBandImage.width) - 4, 0, m_TaskGroupImage.width, 40, 0, IIf(BandType = BAND_ACTIVE Or BandType = BAND_NOTIFY, GROUP_2_ACTIVE, GROUP_2_NORMAL)
            Else
                m_TaskbarBackgroundGraphics.DrawImageStretchAttrF _
                    m_TaskBandImage, CreateRectF(thisProcess.X, 0, 40, m_TaskBandImage.width - 8), 0, BandType, m_TaskBandImage.width, 40, UnitPixel, 0, 0, 0
                
                m_TaskbarBackgroundGraphics.DrawImageRect _
                    m_TaskGroupImage, thisProcess.X + (m_TaskBandImage.width) - 8, 0, m_TaskGroupImage.width, 40, 0, IIf(BandType = BAND_ACTIVE Or BandType = BAND_NOTIFY, GROUP_3_ACTIVE, GROUP_3_NORMAL)
            End If
        End If
        
        Debug.Print "Drawing Process", "Path", thisProcess.Path, "Icon_Width", thisProcess.Image.width, "Icon_Height", thisProcess.Image.height
        m_TaskbarBackgroundGraphics.DrawImage thisProcess.Image, thisProcess.X + 14 + ProcessPadding, 4 + ProcessPadding, 32, 32
    
        thisIndex = thisIndex + 1
        
    Else
        Debug.Print "Icon image is invalid!", thisProcess.Path
    End If


End Function

Private Function FindProcessOnXAxisBeforeGap(XStart As Long)

Dim thisProcess As Process
Dim suspectProcess As Process

Dim processIndex As Long

Dim firstProcessIndex As Long: firstProcessIndex = LBound(m_VisibleList)
Dim lastProcessIndex As Long: lastProcessIndex = UBound(m_VisibleList)

    For processIndex = firstProcessIndex To lastProcessIndex
        Set thisProcess = m_VisibleList(processIndex)
        
        If Not thisProcess Is Nothing Then
            If thisProcess.X < XStart Then
                Set suspectProcess = thisProcess
            Else
                Exit For
            End If
        End If
    Next
    
    Set FindProcessOnXAxisBeforeGap = suspectProcess

End Function

Private Function FindProcessOnXAxisAfterGap(XStart As Long)

Dim thisProcess As Process
Dim suspectProcess As Process

Dim processIndex As Long
Dim firstItemIndex As Long: firstItemIndex = LBound(m_VisibleList)
Dim lastItemIndex As Long: lastItemIndex = UBound(m_VisibleList)

    For processIndex = firstItemIndex To lastItemIndex
        Set thisProcess = m_VisibleList(processIndex)
        
        If Not thisProcess Is Nothing Then
    
            If thisProcess.X < XStart Then
            Else
                If Not thisProcess Is m_lastHoveredGroup Then
                    Set suspectProcess = thisProcess
                    Exit For
                End If
            End If
        End If
    Next
    
    Set FindProcessOnXAxisAfterGap = suspectProcess

End Function

Private Function ForceFocus(hWnd As Long)

    g_hwndForeGroundWindow = hWnd
    SetForegroundWindow hWnd
    
End Function

Private Function HandleWindow(hWnd As Long)

Dim currWinP As WINDOWPLACEMENT
Dim MousePosition As win.POINTL
Dim windowPos As RECT

    If IsWindowHung(hWnd) Then Exit Function

    Call GetCursorPos(MousePosition)
    
    If GetWindowPlacement(hWnd, currWinP) > 0 Then
    
        If (currWinP.ShowCmd = SW_SHOWMINIMIZED) Then
            'minimized, so restore
            ShowWindow hWnd, SW_RESTORE
            
            g_hwndForeGroundWindow = hWnd
            SetForegroundWindow hWnd
            
            GetWindowRect hWnd, windowPos
            
            MoveWindow hWnd, windowPos.Left, windowPos.Top, windowPos.Right - windowPos.Left + 1, windowPos.Bottom - windowPos.Top + 1, ByVal APITRUE
            MoveWindow hWnd, windowPos.Left, windowPos.Top, windowPos.Right - windowPos.Left, windowPos.Bottom - windowPos.Top, ByVal APITRUE
            
            
        ElseIf g_hwndForeGroundWindow = hWnd Then
            'normal, so minimize
            ShowWindow hWnd, SW_MINIMIZE
        Else
        
            SetForegroundWindow hWnd
        End If
    End If
End Function

Private Sub ReTrackMouse()
    m_isTracking = True

    Dim ET As TrackMouseEvent
    'initialize structure
    ET.cbSize = Len(ET)
    ET.hwndTrack = Me.hWnd
    ET.dwFlags = TME_LEAVE
    'start the tracking
    TrackMouseEvent ET
End Sub

Private Sub timHideSubMenuDelay_Timer()
    On Error GoTo Handler
    
    If hWndBelongToUs(GetForegroundWindow()) Then
        Exit Sub
    End If

    'Debug.Print "timHideSubMenuDelay_Timer"
    
    m_SubMenu.height = 0

    'AnimateWindow m_SubMenu.hwnd, 200, AW_HIDE Or AW_VER_POSITIVE
    ShowWindowTimeout m_SubMenu.hWnd, SW_HIDE
    timHideSubMenuDelay.Enabled = False
    
    Exit Sub
Handler:
    LogError Err.number, "HideSubMenuDelay(); " & Err.Description, "winTaskBar"
End Sub

Private Sub ViStartCheck()
    
    If UpdateViStartHwnds Then
        
    End If
End Sub

Private Sub timKeepOnTaskBar_Timer()
    On Error GoTo Handler

Dim newTaskBarEdge As AbeBarEnum
Dim taskBarRect As RECT

    ViStartCheck

    If TaskbarHelper.UpdatehWnds Or _
        IsWindowVisible(g_RunningProgramsHwnd) = APITRUE Then
        
        If Not ApplicationOptions.Floating Then
            If m_currentTaskBarEdge = ABE_TOP Or m_currentTaskBarEdge = abe_bottom Then
                GetWindowRect TaskbarHelper.g_RunningProgramsHwnd, m_ToolListItem
                If RECTWIDTH(m_ToolListItem) <> 0 Then
                    frmZOrderKeeper.DoCommand
                    If Not OptionsHelper.ApplicationOptions.DontShowSplash Then frmClock.ReInstallIcon
                    
                    ShowWindow g_RunningProgramsHwnd, SW_HIDE
                Else
                    Exit Sub
                End If
            End If
            
            If ApplicationOptions.ViOrb Then ShowWindow g_StartButtonHwnd, SW_HIDE
        End If
    End If

    newTaskBarEdge = GetTaskBarEdge
    
    If newTaskBarEdge = ABE_TOP Or newTaskBarEdge = abe_bottom Then
        
        If newTaskBarEdge <> m_currentTaskBarEdge Then
            m_currentTaskBarEdge = newTaskBarEdge
            RaiseEvent onVerticleMode(False)
        End If
        
        GetWindowRect TaskbarHelper.g_RunningProgramsHwnd, m_ToolListItem
        GetWindowRect TaskbarHelper.g_ReBarWindow32Hwnd, taskBarRect
        GetWindowRect Me.hWnd, m_Rect
        
        If RECTWIDTH(m_ToolListItem) > 0 Then
            'Debug.Print "New Height; " & m_ToolListItem.Bottom - m_ToolListItem.Top
            If m_Rect.Left <> m_ToolListItem.Left Or _
                m_Rect.Top <> m_ToolListItem.Top Or _
                RECTWIDTH(m_Rect) <> RECTWIDTH(m_ToolListItem) Or _
                RECTHEIGHT(m_Rect) <> RECTHEIGHT(taskBarRect) Then
            
                'Debug.Print "Repositioning+Resizing Window!"
                MoveWindow Me.hWnd, m_ToolListItem.Left, m_ToolListItem.Top, _
                                        m_ToolListItem.Right - m_ToolListItem.Left, _
                                        taskBarRect.Bottom - taskBarRect.Top, True
            End If
        Else
            ShowWindow g_RunningProgramsHwnd, SW_SHOW
            LogError 238, "impossible sized taskbar window - skipping - " & g_RunningProgramsHwnd, "winTaskBar"
        End If
    Else
        
        If newTaskBarEdge <> m_currentTaskBarEdge Then
            m_currentTaskBarEdge = newTaskBarEdge
            
            RaiseEvent onVerticleMode(True)
        End If
    End If
    
    Exit Sub
Handler:
    LogError Err.number, "KeepOnTask(); " & Err.Description, "winTaskBar"
End Sub

Private Sub timPopupNextDelay_Timer()
    timPopupNextDelay.Enabled = False
    
    Debug.Print "timPopupNextDelay--TrigerSubMenuWithAnimation"
    TriggerSubMenuWithAnimation
End Sub

Private Sub timPopupDelay_Timer()

    If Not IsMouseInsideMe Then
        timPopupDelay.Enabled = False
        Exit Sub
    End If
    
    'If m_lastHoveredGroup.WindowCount > 1 Then
        'Do Simulated Click instead of TriggerMenu
        'To force app into click activation
        
        m_userInvoked = False
        'm_fakeClick = True
        
        'mouse_event _
            &H2 + _
            &H4, _
            0, 0, 0, 0
            
        'SetActiveWindow m_SubMenu.hWnd
        
        Debug.Print "timPopupDelay"
        TriggerSubMenu
        
        'SetForegroundWindow m_SubMenu.hwnd
        'SetActiveWindow m_SubMenu.hwnd
        'BringWindowToTop m_SubMenu.hwnd
        'SwitchToThisWindow m_SubMenu.hwnd, False
        'SetForegroundWindow m_SubMenu.hwnd
        'SetActiveWindow m_SubMenu.hwnd
        'BringWindowToTop m_SubMenu.hwnd
        'SwitchToThisWindow m_SubMenu.hwnd, False
        'ShowWindow m_SubMenu.hwnd, SW_SHOW
    'Else
        'If we don't then the click will activate the window
        'instead of showing the submenu
        'TriggerSubMenu
    'End If
    
    timPopupDelay.Enabled = False
End Sub

Public Function IsMouseInsideMe() As Boolean

Dim currentMousePosition As win.POINTL

    GetCursorPos currentMousePosition
    IsMouseInsideMe = False

    If currentMousePosition.Y > m_Rect.Top And _
       currentMousePosition.Y < m_Rect.Bottom And _
       currentMousePosition.X > m_Rect.Left And _
       currentMousePosition.X < m_Rect.Right Then

       IsMouseInsideMe = True
    End If

End Function

'Private Function IsMouseInsideTaskBar() As Boolean
'
'Dim currentMousePosition As win.POINTL
'
'    GetCursorPos currentMousePosition
'    IsMouseInsideTaskBar = False
'
'    If currentMousePosition.y > m_Rect.Top And _
'       currentMousePosition.y < m_Rect.Bottom And _
'       currentMousePosition.x > m_Rect.Left And _
'       currentMousePosition.x < m_Rect.Right Then
'
'       IsMouseInsideTaskBar = True
'    End If
'
'End Function

Private Sub timReCaptureImage_Timer()
    'Ideally this routine should be in another thread
    'Since we are single threaded, and REDRAW messages could be sent continously
    'Allow app to "breathe"
    
    Debug.Print "timReCaptureImage!"
    
    timReCaptureImage.Enabled = False
    
    If m_targetCaptureWindow Is Nothing Then
        Exit Sub
    End If

    If IsBeingShown(m_targetCaptureWindow.hWnd) Then
        If UpdateThumbnail(m_targetCaptureWindow) Then
            RedrawSubMenu
        End If
    End If
End Sub

Private Sub timTestMouse_Timer()
    On Error GoTo Handler

    If IsMouseInsideWindowsTaskBar = False Then
        'timTestMouse.Enabled = False
        m_notifiedInside = False
        
        If m_notifiedLeft = False Then
            m_notifiedLeft = True
            'm_transparentValue = 255
            timTestMouse.Interval = 20
            
            'MakeTaskbarTransparent 255
            PostMessage Me.hWnd, WM_MOUSELEAVE, 0, 0
        ElseIf m_transparentValue > 180 Then
            
            m_transparentValue = m_transparentValue - 10
            MakeTaskbarTransparent m_transparentValue
        Else
            timTestMouse.Interval = 100
        End If
        
    Else
        m_notifiedLeft = False
    
        If m_notifiedInside = False Then
            m_notifiedInside = True
            'm_transparentValue = 180
            timTestMouse.Interval = 20
            
            'Debug.Print "Making Transparent!"
            'MakeTaskbarTransparent 118
            
        ElseIf (m_transparentValue + 25) < 255 Then
            
            m_transparentValue = m_transparentValue + 10
            MakeTaskbarTransparent m_transparentValue
        Else
            timTestMouse.Interval = 100
        End If
    End If

    Exit Sub
Handler:
    LogError Err.number, "TestMouse(); " & Err.Description, "winTaskBar"
End Sub

Private Sub SetPositionNewProcesses()

Dim thisProcess As Process
Dim thisProcessIndex As Long

    thisProcessIndex = m_TaskList.Processes.Count

    For Each thisProcess In m_TaskList.Processes
        If thisProcess.X = -1 Then
            thisProcess.X = (m_TaskList.Processes.Count - thisProcessIndex) * (m_TaskBandImage.width + 2)
        End If
        
        thisProcessIndex = thisProcessIndex - 1
    Next

End Sub

Public Function SortPinnedList()
    On Error GoTo Handler

Dim lngX As Long
Dim PinnedIndexArray As Long

Dim thisProcess As Process
Dim lengthVisibleList As Long
Dim emptyByteStringArray() As ShellLink
Dim firstItemIndex As Long

    ApplicationOptions.PinnedApplications = emptyByteStringArray
    
    lengthVisibleList = UBound(m_VisibleList)
    firstItemIndex = LBound(m_VisibleList)

    For lngX = firstItemIndex To lengthVisibleList
        Set thisProcess = m_VisibleList(lngX)
        
        If thisProcess.Pinned Then
            ReDim Preserve ApplicationOptions.PinnedApplications(PinnedIndexArray)
        
            ApplicationOptions.PinnedApplications(PinnedIndexArray).szPath = thisProcess.Path
            ApplicationOptions.PinnedApplications(PinnedIndexArray).szArguments = thisProcess.Arguments
        
            PinnedIndexArray = PinnedIndexArray + 1
        End If
    Next

Handler:
End Function
