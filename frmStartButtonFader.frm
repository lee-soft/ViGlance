VERSION 5.00
Begin VB.Form frmFader 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Start_Fader"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   LinkTopic       =   "Form3"
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timWindowChecker 
      Interval        =   1000
      Left            =   360
      Top             =   960
   End
   Begin VB.Timer timFadeOut 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   600
      Top             =   240
   End
   Begin VB.Timer timFadeIn 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "frmFader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmFader
'    Project    : prjSuperBar
'
'    Description: The fading overlay for the start button
'
'--------------------------------------------------------------------------------
Option Explicit

Public Event onRolledOver()
Public Event onRolledOut()
Public Event onClicked()
Public Event onMouseUp()

Private ORB_HEIGHT As Long
Private ORB_WIDTH As Long

Private winSize As Size
Private curWinLong As Long
Private srcPoint As POINTAPI
Private blendFunc32bpp As BLENDFUNCTION
Private m_theHDC As Long

Private m_MouseInClientArea As Boolean
Private m_MouseEvents As TrackMouseEvent
Private lastAlpha As Long

Implements IHookSink

Public Function SetDimensions(newWidth As Long, newHeight As Long)
    ORB_HEIGHT = newHeight
    ORB_WIDTH = newWidth
    
    Me.height = ORB_HEIGHT * Screen.TwipsPerPixelY
    Me.width = ORB_WIDTH * Screen.TwipsPerPixelX
    
    winSize.cx = ORB_WIDTH
    winSize.cy = ORB_HEIGHT
End Function

Public Property Let Alpha(newAlpha As Byte)
    blendFunc32bpp.SourceConstantAlpha = newAlpha
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, m_theHDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Property

'Private m_FaderWindow As frmStartButton
Sub InitializeVariables(ByVal newHDC As Long)
    m_theHDC = newHDC
    UpdateAndReDraw
End Sub

Private Sub Form_Initialize()

    Call HookWindow(Me.hWnd, Me)
End Sub

Private Sub Form_Load()
    MakeTrans
End Sub

Sub UpdateAndReDraw()
    'Debug.Assert srcPoint.x
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, winSize, m_theHDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Sub

Private Sub MakeTrans()
    Debug.Print "MakeTrans"

    curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOOLWINDOW
    
    'update layer window stuff (that will blend and show the GDI+ text)
    srcPoint.X = 0
    srcPoint.Y = 0
    winSize.cx = ORB_WIDTH
    winSize.cy = ORB_HEIGHT

    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent onMouseUp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnhookWindow(Me.hWnd)
    DisposeGDIIfLast
End Sub

Private Function IHookSink_WindowProc(hWnd As Long, iMsg As Long, wParam As Long, lParam As Long) As Long
    On Error GoTo Handler

    
    If iMsg = WM_MOUSEMOVE Then
        If Not m_MouseInClientArea Then
            m_MouseInClientArea = True

            m_MouseEvents.cbSize = Len(m_MouseEvents)
            m_MouseEvents.dwFlags = TME_LEAVE
            m_MouseEvents.hwndTrack = Me.hWnd
            
            TrackMouseEvent m_MouseEvents
            
            RaiseEvent onRolledOver
        End If
        
    ElseIf iMsg = WM_MOUSELEAVE Then
        m_MouseInClientArea = False
        
        RaiseEvent onRolledOut
    ElseIf iMsg = WM_LBUTTONDOWN Then
        RaiseEvent onClicked
    Else
         ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            InvokeWindowProc(hWnd, iMsg, wParam, lParam)
    End If
    
    Exit Function
Handler:
    LogError Err.number, "WindowProc(" & iMsg & "," & wParam & "," & lParam & "); " & Err.Description, "winOrbFader"
    
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
        InvokeWindowProc(hWnd, iMsg, wParam, lParam)
End Function

Sub FadeIn()
    timFadeOut.Enabled = False
    timFadeIn.Enabled = True
End Sub

Sub FadeOut()
    timFadeOut.Enabled = True
    timFadeIn.Enabled = False
End Sub

Private Sub timFadeIn_Timer()

    If lastAlpha < 255 Then
        lastAlpha = lastAlpha + 15
        Alpha = CByte(lastAlpha)
    Else
        timFadeIn.Enabled = False
    End If
End Sub

Private Sub timFadeOut_Timer()

    If lastAlpha > 1 Then
        lastAlpha = lastAlpha - 15
        Alpha = CByte(lastAlpha)
    Else
        timFadeOut.Enabled = False
        Alpha = CByte(1)
    End If
End Sub
