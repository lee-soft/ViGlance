VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ViOrb_Fader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event onMove()

Private Const ORB_HEIGHT As Long = 44
Private Const ORB_WIDTH As Long = 44

Private m_theStartButton As GDIPImage
' Create a Graphics object:
Private m_gfx As GDIPGraphics
Private m_Bitmap As GDIPBitmap
Private m_BitmapGraphics As GDIPGraphics

Private tempBI As BITMAPINFO
Private winSize As Size
Private curWinLong As Long
Private mainBitmap As Long
Private oldBitmap As Long
Private srcPoint As POINTAPI
Private blendFunc32bpp As BLENDFUNCTION

Private m_hDC As Long
Private m_hWnd As Long

Private m_memoryDC As Long
Private m_hBmp As Long
Private m_hBmpOld As Long

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Let Alpha(newAlpha As Byte)
    blendFunc32bpp.SourceConstantAlpha = newAlpha
    'Call UpdateLayeredWindow(Me.hwnd, Me.hDC, ByVal 0&, winSize, Me.hDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Property

'Private m_FaderWindow As frmStartButton
Sub InitializeVariables(ByRef theStartButton As GDIPImage, Bitmap As GDIPBitmap, BitmapGraphics As GDIPGraphics)
    Set m_theStartButton = theStartButton
    Set m_Bitmap = Bitmap
    Set m_BitmapGraphics = BitmapGraphics
    
    MakeTrans
End Sub

Private Sub Form_Click()
    UpdateAndReDraw
End Sub

Private Sub Form_DblClick()
    blendFunc32bpp.SourceConstantAlpha = 100
    'now the text will be shown
    'Call UpdateLayeredWindow(Me.hWnd, Me.hDC, ByVal 0&, winSize, Me.hDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
    'UpdateAndReDraw
    'Unload Me
End Sub

Private Sub Class_Initialize()
    InitializeAndRegisterClass Me

    Set m_gfx = New GDIPGraphics
    
    m_hWnd = CreateWindowEx(WS_EX_TOPMOST Or WS_EX_LAYERED, ViOrb_Fader_Support.g_WndClassViOrbFader.lpszClassName, _
                                "Start_Fader", WS_OVERLAPPED Or WS_POPUP Or WS_VISIBLE, CW_USEDEFAULT, CW_USEDEFAULT, 640&, 480&, 0, 0&, _
                                    0, 0&)
    m_hDC = GetWindowDC(m_hWnd)
    
    m_memoryDC = CreateCompatibleDC(m_hDC)
    m_hBmp = CreateCompatibleBitmap(m_memoryDC, ORB_WIDTH, ORB_HEIGHT)
    m_hBmpOld = SelectObject(m_memoryDC, m_hBmp)
    
    m_gfx.FromHDC m_memoryDC
End Sub

Sub UpdateAndReDraw()
    m_BitmapGraphics.Clear
    m_BitmapGraphics.DrawImageRect m_theStartButton, 0, 0, ORB_WIDTH, ORB_HEIGHT, 0, 0

    m_gfx.Clear vbBlack
    m_gfx.DrawImage _
       m_Bitmap.Image, 0, 0, ORB_WIDTH, ORB_HEIGHT
       
    Call UpdateLayeredWindow(m_hWnd, m_hDC, ByVal 0&, winSize, m_memoryDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Sub

Private Sub MakeTrans()
    'update layer window stuff (that will blend and show the GDI+ text)
    srcPoint.x = 0
    srcPoint.y = 0
    winSize.cx = ORB_WIDTH
    winSize.cy = ORB_HEIGHT

    With blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    'now the text will be shown
    Call UpdateLayeredWindow(m_hWnd, m_hDC, ByVal 0&, winSize, m_memoryDC, srcPoint, 0, blendFunc32bpp, ULW_ALPHA)
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal message As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long

    On Error GoTo Abort
    
    Select Case message
    
    Case WM_PAINT
        
    
    Case WM_CREATE

        'MakeTrans
        
    Case WM_MOVE
        RaiseEvent onMove
        
    Case Else
        WndProc = DefWindowProc(hWnd, message, wParam, lParam)
    End Select
    
    Exit Function
Abort:
    WndProc = DefWindowProc(hWnd, message, wParam, lParam)
End Function

Private Sub Class_Terminate()
    m_gfx.Dispose
    DisposeGDIIfLast
End Sub
