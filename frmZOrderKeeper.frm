VERSION 5.00
Begin VB.Form frmZOrderKeeper 
   Caption         =   "Container"
   ClientHeight    =   3600
   ClientLeft      =   -76680
   ClientTop       =   17415
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timStartMenuCheck 
      Interval        =   100
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer timZOrderChecker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   480
      Top             =   2400
   End
End
Attribute VB_Name = "frmZOrderKeeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmZOrderKeeper
'    Project    : prjSuperBar
'
'    Description: Keeps our windows above the taskbar in zorder stack
'
'--------------------------------------------------------------------------------


Option Explicit

Private WithEvents Taskbar As frmTaskbar
Attribute Taskbar.VB_VarHelpID = -1
Private m_NotTopMost As Boolean
Private m_verticleTaskBar As Boolean


'--------------------------------------------------------------------------------
' Procedure  :       DoCommand
' Description:       Makes the native windows taskbar "own" this window so
'                    when the windows taskbar is activated, it's children
'                    activated as a result. Meaning we can control what stays
'                    on top
' Parameters :
'--------------------------------------------------------------------------------
Sub DoCommand()

    Debug.Print "DoCommand"
    
    TaskbarHelper.UpdatehWnds
    Set Taskbar = frmTaskbar

    frmTaskbar.Show
    
    
    
    'If ApplicationOptions.ViOrb Then
    '    frmFader.Show
    '    frmStartButton.Show
    'End If
    
    frmSubMenu.Show
    
    SetOwner frmTaskbar.hWnd, Me.hWnd
    
    'If ApplicationOptions.ViOrb Then
    '    SetOwner frmFader.hWnd, Me.hWnd
    '    SetOwner frmStartButton.hWnd, Me.hWnd
    'End If
    
    SetOwner frmSubMenu.hWnd, Me.hWnd
    
    'Debug.Print g_ReBarWindow32Hwnd
    SetParent Me.hWnd, TaskbarHelper.g_StartButtonHwnd
    'SetParent frmTaskbar.hWnd, g_ReBarWindow32Hwnd
    
    If Not ApplicationOptions.Floating And ApplicationOptions.ViOrb Then
        frmStartButton.MoveOrbIfNotOverStartButton
    End If

    StayOnTop Me, True
    StayOnTop frmTaskbar, True
    
    'If ApplicationOptions.ViOrb Then
    '    StayOnTop frmFader, True
    '    StayOnTop frmStartButton, True
    'End If
    
    StayOnTop frmSubMenu, True
    Me.Visible = False
End Sub

Private Sub Form_Load()
    DoCommand
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ShowWindow g_RunningProgramsHwnd, SW_SHOW
    ShowWindow g_StartButtonHwnd, SW_SHOW
    frmTaskbar.SortPinnedList

    Unload frmOptions
    Unload frmSplash
    Unload frmClock
    Unload frmTaskbar
    Unload frmStartButton
        
    DumpOptions
    Unload Me
End Sub

Private Sub Taskbar_onVerticleMode(VerticleMode As Boolean)
    On Error GoTo Handler
    Debug.Print "VERTICLE MODE!"

    m_verticleTaskBar = VerticleMode

    If m_verticleTaskBar Then
    
        If IsWindowVisible(g_RunningProgramsHwnd) = APIFALSE Then
            Debug.Print "VERT MODE"
            
            frmTaskbar.Hide
            ShowWindow g_RunningProgramsHwnd, SW_SHOW
        End If
    Else

        frmTaskbar.Show
        ShowWindow g_RunningProgramsHwnd, SW_HIDE
    End If
    
    Exit Sub
Handler:
    LogError Err.number, "verticleMode(" & VerticleMode & "); " & Err.Description, "winZOrder"
End Sub

Private Sub timStartMenuCheck_Timer()
    g_StartMenuOpen = IsStartMenuOpen
    g_viStartRunning = IsViStartOpen
    
    
End Sub

Private Sub timZOrderChecker_Timer()
    On Error GoTo Handler

    'Debug.Print "timZOrderChecker!"
    
    'Enforces Z-Order's order :P
Dim hWndForeGroundWindow As Long
Dim zOrderTaskBar As Long
    
    hWndForeGroundWindow = GetForegroundWindow
    zOrderTaskBar = GetZOrder(frmTaskbar.hWnd)

    'Debug.Print "AppTaskBar; " & zOrderTaskBar & " : " & GetZOrder(g_TaskBarHwnd)
    'Debug.Print "AppTaskBarVisible; " & frmTaskbar.Visible

    If (GetZOrder(g_TaskBarHwnd) < zOrderTaskBar) And m_NotTopMost = False Then
        'SetWindowPos frmTaskbar.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        
        SetWindowPos frmTaskbar.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        If Not hWndBelongToUs(hWndForeGroundWindow) And _
            hWndForeGroundWindow <> TaskbarHelper.g_TaskBarHwnd And _
            hWndForeGroundWindow <> TaskbarHelper.g_StartMenuHwnd Then
            
            'Debug.Print "Test Result 1: " & Now
            'Debug.Print "IsWindowTopMost: " & IsWindowTopMost(hWndForeGroundWindow)
            'Debug.Print "TaskBarVisible: " & IsWindowVisible(TaskbarHelper.g_TaskBarHwnd)
            'Debug.Print "IsTaskBarBehindWindow: " & IsTaskBarBehindWindow(hWndForeGroundWindow)
           
            If IsTaskBarBehindWindow(hWndForeGroundWindow) Then
                If IsWindowTopMost(hWndForeGroundWindow) = False Then
                    Debug.Print "Hiding forms!!"
                    
                    m_NotTopMost = True
                    
                    If Not m_verticleTaskBar Then frmTaskbar.Hide
                    If ApplicationOptions.ViOrb Then
                        Debug.Print "Hiding ViOrb!!"
                    
                        frmStartButton.Hide
                        frmFader.Hide
                    End If
                        
                    frmSubMenu.Hide
                Else
                    If zOrderTaskBar < GetZOrder(hWndForeGroundWindow) Then
                    'Debug.Print OptionsHelper.GetWindowClassString(hWndForeGroundWindow)
                        
                        m_NotTopMost = True 'viglance/vistart stealing focus
                        SetWindowPos hWndForeGroundWindow, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
                    End If
                End If
                
                'Me.Hide
            Else
                If m_NotTopMost = True Then
                    m_NotTopMost = False
                    
                    If Not m_verticleTaskBar Then frmTaskbar.Show
                    If ApplicationOptions.ViOrb Then
                        frmStartButton.Show
                        frmFader.Show
                    End If
                    
                    'frmSubMenu.Show
                    'Me.Show
                End If
            End If
        End If
    End If

    Exit Sub
Handler:
    LogError Err.number, "zOrderCheck(" & Err.Description & ")", "winZOrder"
End Sub
