VERSION 5.00
Begin VB.Form frmStartButton 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "#Start~ViGlance#"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form3"
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmStartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmStartButton
'    Project    : prjSuperBar
'
'    Description: A physical startorb button that optionally replaces the
'                 official windows start button. (Basically ViOrb)
'
'--------------------------------------------------------------------------------
Option Explicit

Private ORB_HEIGHT As Long
Private ORB_WIDTH As Long

Private m_theStartButton As GDIPImage
Private WithEvents m_FaderWindow As frmFader
Attribute m_FaderWindow.VB_VarHelpID = -1

Private WithEvents timKeepOnStartButton As Timer
Attribute timKeepOnStartButton.VB_VarHelpID = -1

' Create a Graphics object:
Private m_gfx As GDIPGraphics
Private m_Bitmap As GDIPBitmap
Private m_BitmapGraphics As GDIPGraphics
Private m_SourcePositionY As Long
Private m_reBar32_Rect As RECT
Private m_Rect As RECT

Private winSize As Size
Private curWinLong As Long
Private srcPoint As POINTAPI
Private blendFunc32bpp As BLENDFUNCTION

Private Function SetFaderWindow()
    'On Error Resume Next
    
    Set m_FaderWindow = frmFader
    m_FaderWindow.SetDimensions ORB_WIDTH, ORB_HEIGHT
    
    Load m_FaderWindow
    m_FaderWindow.InitializeVariables Me.hdc
End Function

Private Property Let frameIndex(ByVal newIndex As Long)
    m_SourcePositionY = newIndex * ORB_HEIGHT
End Property

Private Sub Form_Initialize()

    Set m_theStartButton = New GDIPImage
    Set m_gfx = New GDIPGraphics
    Set m_BitmapGraphics = New GDIPGraphics
    Set m_Bitmap = New GDIPBitmap
    
    Set timKeepOnStartButton = Controls.Add("VB.Timer", "timKeepOnStartButton", Me)
    timKeepOnStartButton.Interval = 1000
    timKeepOnStartButton.Enabled = True
    
End Sub

Private Sub Form_Load()
    m_theStartButton.FromFile App.Path & "\resources\start_button.png"
    
    ORB_HEIGHT = m_theStartButton.height / 3
    ORB_WIDTH = m_theStartButton.width
    
    Me.width = ORB_WIDTH * Screen.TwipsPerPixelX
    Me.height = ORB_HEIGHT * Screen.TwipsPerPixelY
    
    winSize.cx = Me.ScaleWidth
    winSize.cy = Me.ScaleHeight
    
    ' Initialise it to work on the PictureBox HDC:
    m_gfx.FromHDC Me.hdc
    
    m_Bitmap.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, PixelFormat.Format32bppArgb
    m_BitmapGraphics.FromImage m_Bitmap.Image
    
    MakeTrans
    UpdateAndReDraw
    'UpdateAndReDraw
    
    SetFaderWindow
    'm_FaderWindow.InitializeVariables Me.hDC
End Sub

Sub UpdateAndReDraw(Optional ByVal UpdateHDC As Boolean = True)
    m_BitmapGraphics.Clear
    m_BitmapGraphics.DrawImageRect m_theStartButton, 0, 0, ORB_WIDTH, ORB_HEIGHT, 0, m_SourcePositionY

    m_gfx.Clear vbBlack
    m_gfx.DrawImage _
       m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
       
    If UpdateHDC Then Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, Me.hdc, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Sub

Private Sub MakeTrans()

    curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW
    
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_FaderWindow_onClicked
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StayOnTop m_FaderWindow, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload m_FaderWindow
    Set m_FaderWindow = Nothing
    
    m_gfx.Dispose
    m_Bitmap.Dispose
    m_theStartButton.Dispose
    m_BitmapGraphics.Dispose
    
    DisposeGDIIfLast
End Sub

Private Sub m_FaderWindow_onClicked()
    If g_StartMenuOpen = False Then
        If g_viStartRunning = False Then
            TaskbarHelper.ShowStartMenu
        Else
            SetKeyDown 91
            SetKeyUp 91
        End If
    End If
End Sub

Private Sub m_FaderWindow_onRolledOut()
    UpdateAndReDraw False
    m_FaderWindow.FadeOut
End Sub

Private Sub m_FaderWindow_onRolledOver()
    frameIndex = 1
    
    UpdateAndReDraw False
    m_FaderWindow.FadeIn
End Sub

Public Function MoveOrbIfNotOverStartButton()

Dim recStartButton As RECT
Dim lngTop As Long

    GetWindowRect g_ReBarWindow32Hwnd, m_reBar32_Rect
    GetWindowRect g_StartButtonHwnd, recStartButton
    GetWindowRect Me.hWnd, m_Rect

    If GetTaskBarEdge = ABE_LEFT Or GetTaskBarEdge = ABE_RIGHT Then
        Debug.Print "Awww c'mon man!"
    
        lngTop = -1
    Else
        If m_reBar32_Rect.Bottom - m_reBar32_Rect.Top < 40 _
            Then
            
            lngTop = 12
        Else
            lngTop = (ORB_HEIGHT / 2) - (m_reBar32_Rect.Bottom - m_reBar32_Rect.Top) / 2
            'Debug.Print m_reBar32_Rect.Top - lngTop
        End If
    End If

    If lngTop <> -1 Then
        If ((recStartButton.Left) <> (m_Rect.Left) Or _
            (m_reBar32_Rect.Top - lngTop) <> m_Rect.Top) Then
    
            Debug.Print "MOVING; " & (recStartButton.Top) & "<>" & (m_Rect.Top)
            
            MoveWindow Me.hWnd, recStartButton.Left, m_reBar32_Rect.Top - lngTop, Me.ScaleWidth, Me.ScaleHeight, False
            SnapFaderOverMe
        End If
    Else
        If ((recStartButton.Left) <> (m_Rect.Left) Or _
            (recStartButton.Top - 5) <> m_Rect.Top) Then
            
            MoveWindow Me.hWnd, recStartButton.Left, recStartButton.Top - 5, Me.ScaleWidth, Me.ScaleHeight, False
            SnapFaderOverMe
        End If
    End If
End Function

Private Function SnapFaderOverMe()
    GetWindowRect Me.hWnd, m_Rect
    SetWindowPos m_FaderWindow.hWnd, 0, m_Rect.Left, m_Rect.Top, 0, 0, SWP_NOSIZE Or SWP_NOREPOSITION
End Function

Private Sub timKeepOnStartButton_Timer()
    MoveOrbIfNotOverStartButton
End Sub
