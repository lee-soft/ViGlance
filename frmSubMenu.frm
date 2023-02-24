VERSION 5.00
Begin VB.Form frmSubMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   675
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   32
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmSubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event onClosed(targetWindow As Window)
Public Event onClicked(targetWindow As Window)
Public Event onRightClicked(targetWindow As Window)

Public Event onMouseOver()
Public Event onMouseOut()
Public Event onDeactivated(hWnd As Long)

Private Const FRAME_LEFT_STRETCH_SEGMENT As Long = 17

Private Const BUTTON_Y_GAP As Long = 44
Private Const ITEM_WIDTH As Long = 217

Private Const LEFT_PADDING As Long = 20
Private Const RIGHT_PADDING As Long = 26

Private Const ITEM_PADDING As Long = 40

Private Const CLOSE_BUTTON_MARGIN As Long = 9

Private m_Bitmap As GDIPBitmap
Private m_BitmapGraphics As GDIPGraphics
Private m_Graphics As GDIPGraphics
Private m_GroupMenu As ProgramFrames
Private m_GroupMenuButton As GroupMenuButton

Private m_closeButton As GDIPImage
Private m_closeButtonSize As SIZEL

Private m_Font As GDIPFont
Private m_Brush As GDIPBrush
Private m_Pen As GDIPPen
Private m_Path As GDIPGraphicPath
Private m_FontFamily As GDIPFontFamily

Private m_srcPoint As POINTF
Private m_winSize As Size
Private m_blendFunc32bpp As BLENDFUNCTION

Private m_ListPosition As POINTL
Private m_currentLayout As RECT
Private m_WindowList As Collection

Private m_hoveredOverWindow As Window
Private m_selectedWindow As Window

Private m_isTracking As Boolean
Private m_haltDeActivation As Boolean
Private m_allowListRedraw As Boolean
Private m_aspectRatio As Single

Private m_mouseOnCloseButton As Boolean

Private Type ProgramFrames
    Image As GDIPImage

    TopLeft As RECTL
    TopRight As RECTL
    BottomLeft As RECTL
    BottomRight As RECTL
    
    MiddleLeft As RECTL
    MiddleRight As RECTL
    MiddleBottom As RECTL
    MiddleTop As RECTL
    Middle As RECTL
End Type

Private Type GroupSection
    TopLeft As RECTL
    TopRight As RECTL
    BottomLeft As RECTL
    BottomRight As RECTL
    
    MiddleLeft As RECTL
    MiddleRight As RECTL
    MiddleBottom As RECTL
    MiddleTop As RECTL
    Middle As RECTL
End Type

Private Type GroupMenuButton
    Image As GDIPImage
    
    NormalSection As GroupSection
    SelectedSection As GroupSection
    FlashingSection As GroupSection
End Type

Enum GroupMenuButtonStates
    Normal = 0
    Hovered = 1
    Selected = 2
    Flashing = 3
End Enum

Private m_currentMode As WindowMode

Implements IHookSink

Private Function PotentialWindowPreviewSize() As Long
    If m_WindowList Is Nothing Then Exit Function
    PotentialWindowPreviewSize = (GetSystemMetrics(SM_CXSCREEN) / m_WindowList.Count)
    
    'Debug.Print "PotentialWindowPreviewSize:: " & PotentialWindowPreviewSize
End Function

Public Property Let AspectRatio(ByVal newAspectRatio As Single)
    m_aspectRatio = newAspectRatio
End Property

Public Property Get AspectRatio() As Single
    AspectRatio = m_aspectRatio
End Property

Public Property Let AllowListRedraw(ByVal newValue As Boolean)
    m_allowListRedraw = newValue
End Property

Public Function ResetRollover()
    Set m_hoveredOverWindow = Nothing
    Set m_selectedWindow = Nothing
End Function

Public Property Set WindowList(ByRef newList As Collection)
    Set m_WindowList = newList
    
    Debug.Print "New List!"
    
    If OptionsHelper.ApplicationOptions.TextOnlyMode Or PotentialWindowPreviewSize() < 120 Then
        If Not m_currentMode = TextMode Then
            m_currentMode = TextMode
            'RaiseEvent onResize
        End If
    Else
        If Not m_currentMode = ThumbnailMode Then
            m_currentMode = ThumbnailMode
            'RaiseEvent onResize
        End If
    End If
End Property

Public Property Get WindowList() As Collection
    Set WindowList = m_WindowList
    'Debug.Print m_WindowList.Count
    
    'Me.Height = (newList.Count * TEXTMODE_ITEM_Y_GAP) * Screen.TwipsPerPixelY
End Property

Private Sub Form_Initialize()
 
    Call HookWindow(Me.hWnd, Me)

    Set m_Graphics = New GDIPGraphics
    Set m_BitmapGraphics = New GDIPGraphics
    Set m_Bitmap = New GDIPBitmap
    Set m_Font = New GDIPFont
    Set m_Brush = New GDIPBrush
    
    Set m_Path = New GDIPGraphicPath
    Set m_Pen = New GDIPPen
    Set m_FontFamily = New GDIPFontFamily
    
    m_aspectRatio = 1
    
    InitializeProgramFrames
    InitializeGroupButton
    InitializeCloseButton
    
    'm_Brush.Colour.Value = RGB(0, 2, 92)
    'm_Pen.Colour.Value = RGB(0, 2, 92)
    m_Brush.Colour.Value = vbWhite
    
    m_FontFamily.Constructor GeneralHelper.GetClosestVistaFont
    m_Font.Constructor m_FontFamily, 12
    
    MakeTrans
    ReInitSurface
End Sub

Private Function InitializeCloseButton()

    Set m_closeButton = New GDIPImage
    
    m_closeButton.FromFile App.Path & "\resources\close_button.png"
    
    m_closeButtonSize.height = m_closeButton.height
    m_closeButtonSize.width = m_closeButton.width / 3

End Function

Private Function InitializeGroupButton() As Boolean

    Set m_GroupMenuButton.Image = New GDIPImage
    m_GroupMenuButton.Image.FromFile App.Path & "\resources\taskbar_groupmenu_button.png"
    
    InitializeGroupFrame m_GroupMenuButton.NormalSection, 1
    InitializeGroupFrame m_GroupMenuButton.SelectedSection, 2
    InitializeGroupFrame m_GroupMenuButton.FlashingSection, 3

End Function

Private Function InitializeGroupFrame(ByRef thisType As GroupSection, ByVal frameNumber As Long)

Dim frameHeight As Long
Dim frameWidth As Long
Dim actualFrameHeight As Long

Dim YMiddle As Long: YMiddle = 34
Dim XMiddle As Long: XMiddle = 1

Dim imageWidth As Long
    
    frameHeight = ((m_GroupMenuButton.Image.height / 4) / 2) - YMiddle
    '3
    
    frameWidth = (m_GroupMenuButton.Image.width / 2) - XMiddle
    '51
    
    actualFrameHeight = m_GroupMenuButton.Image.height / 4
    '74
    
    imageWidth = m_GroupMenuButton.Image.width
    '104
    
    'imageHeight = m_GroupMenuButton.Image.height
    '?

    With thisType
        .TopLeft = CreateRectL( _
                            0, _
                            actualFrameHeight * frameNumber, _
                            frameHeight, _
                            frameWidth)
        
        .MiddleTop = CreateRectL( _
                            frameWidth, _
                            (actualFrameHeight * frameNumber), _
                             frameHeight, _
                            XMiddle * 2)
        
        .TopRight = CreateRectL( _
                            imageWidth - frameWidth, _
                            actualFrameHeight * frameNumber, _
                             frameHeight, _
                            frameWidth)
                            
        .MiddleLeft = CreateRectL( _
                            0, _
                            (actualFrameHeight * frameNumber) + frameHeight, _
                             YMiddle * 2, _
                            frameWidth)
                            
        .BottomLeft = CreateRectL( _
                            0, _
                            (actualFrameHeight * frameNumber) + frameHeight + YMiddle * 2, _
                             frameHeight, _
                            frameWidth)
                            
        .MiddleBottom = CreateRectL( _
                            frameWidth, _
                            (actualFrameHeight * frameNumber) + (frameHeight + YMiddle * 2), _
                             frameHeight, _
                            XMiddle * 2)
                            
        .BottomRight = CreateRectL( _
                            imageWidth - frameWidth, _
                            (actualFrameHeight * frameNumber) + (frameHeight + YMiddle * 2), _
                             frameHeight, _
                            frameWidth)
        
        .MiddleRight = CreateRectL( _
                            imageWidth - frameWidth, _
                            (actualFrameHeight * frameNumber) + (frameHeight), _
                             YMiddle * 2, _
                            frameWidth)
                            
        .Middle = CreateRectL( _
                            frameWidth, (actualFrameHeight * frameNumber) + (frameHeight), YMiddle * 2, XMiddle * 2)
    End With

End Function

Private Function InitializeProgramFrames() As Boolean

    Set m_GroupMenu.Image = New GDIPImage
    m_GroupMenu.Image.FromFile App.Path & "\resources\taskbar_groupmenu.png"

    With m_GroupMenu.TopLeft
        .height = 18
        .width = 19
        .Left = 0
        .Top = 0
    End With
    
    With m_GroupMenu.MiddleLeft
        .Top = m_GroupMenu.TopLeft.height
        .Left = 0
        .height = FRAME_LEFT_STRETCH_SEGMENT
        .width = 19
    End With
    
    With m_GroupMenu.BottomLeft
        .height = 24
        .width = 19
        
        .Left = 0
        .Top = m_GroupMenu.Image.height - .height
    End With
    
    With m_GroupMenu.MiddleBottom
        .height = 24
        .width = 13
        .Left = m_GroupMenu.BottomLeft.width
        .Top = m_GroupMenu.BottomLeft.Top
    End With
    
    With m_GroupMenu.BottomRight
        .height = 24
        .width = 24
        
        .Left = m_GroupMenu.Image.width - .width
        .Top = m_GroupMenu.Image.height - .height
    End With
    
    With m_GroupMenu.TopRight
        .height = 18
        .width = 24
        
        .Left = m_GroupMenu.Image.width - .width
        .Top = 0
    End With
    
    With m_GroupMenu.MiddleRight
        .height = 23
        .width = 24
        
        .Left = 13 + m_GroupMenu.MiddleLeft.width
        .Top = m_GroupMenu.TopRight.height
    End With
    
    With m_GroupMenu.MiddleTop
        .height = 18
        .width = 13
        .Left = m_GroupMenu.TopLeft.width
        .Top = 0
    End With
    
    With m_GroupMenu.Middle
        .Top = m_GroupMenu.TopLeft.height
        .Left = m_GroupMenu.TopLeft.width
        .width = 13
        .height = 23
    End With

End Function

Public Function PredictHeight()
    If (m_currentMode = TextMode) Then
        PredictHeight = PredictHeightText
    Else
        PredictHeight = PredictHeightThumbnail
    End If
End Function

Public Function PredictWidth()
    If (m_currentMode = TextMode) Then
        PredictWidth = PredictWidthText
    Else
        PredictWidth = PredictWidthThumbnail
    End If
End Function

Private Function PredictHeightText()
    PredictHeightText = ((m_WindowList.Count * TEXTMODE_ITEM_Y_GAP) + THUMBNAIL_TOP_PADDING) + (m_GroupMenu.TopLeft.height + m_GroupMenu.BottomLeft.height)
End Function

Private Function PredictHeightThumbnail() As Long
'On Error GoTo Handler

Dim thisWindow As Window
Dim maxHeight As Long
Dim targetHeight As Long

    maxHeight = 0

    For Each thisWindow In m_WindowList
        If thisWindow.Image.height = 0 Then
            targetHeight = thisWindow.Parent.Image.height
        Else
            targetHeight = thisWindow.Image.height
        End If
        
        If targetHeight > maxHeight Then
            maxHeight = targetHeight
        End If
    Next
    
    PredictHeightThumbnail = maxHeight + SUBMENU_Y_PADDING
End Function

Private Function PredictWidthThumbnail() As Long
    PredictWidthThumbnail = (m_WindowList.Count * (ITEM_WIDTH + THUMBNAIL_ITEM_Y_GAP)) + (LEFT_PADDING + RIGHT_PADDING)
End Function

Private Function PredictWidthText()
On Error GoTo Handler

Dim thisWindow As Window
Dim lpRect As RECTF
Dim maxWidth As Long

    For Each thisWindow In m_WindowList
        lpRect = m_BitmapGraphics.MeasureString(thisWindow.Caption, m_Font)
        
        If lpRect.width > maxWidth Then
            maxWidth = lpRect.width
        End If
    Next
    
    PredictWidthText = LEFT_PADDING + maxWidth + ITEM_PADDING + m_GroupMenu.TopLeft.width + m_GroupMenu.TopRight.width
    
    Exit Function
Handler:
    Debug.Print "PridictWidth()" & Err.Description
End Function

Sub DrawCloseButton(X As Long, Y As Long, frameIndex As Long)

    'm_BitmapGraphics.DrawImage m_closeButton, X, Y, m_closeButton.width, m_closeButton.height
    
    m_BitmapGraphics.DrawImageRect m_closeButton, X, Y, m_closeButtonSize.width, m_closeButton.height, m_closeButtonSize.width * (frameIndex - 1), 0

End Sub

Sub DrawBorder(Layout As RECT, Optional update As Boolean = True)
On Error GoTo Handler

Dim theHeight As Long
Dim theWidth As Long

Dim layoutBottom As Long
Dim layoutRight As Long
    
    layoutBottom = Layout.Bottom - m_GroupMenu.MiddleBottom.height
    layoutRight = Layout.Right - m_GroupMenu.MiddleRight.width

    theHeight = (layoutBottom - Layout.Top) - m_GroupMenu.TopLeft.height
    theWidth = (layoutRight - Layout.Left) - m_GroupMenu.TopLeft.width

    m_BitmapGraphics.Clear
    
    DrawImageRect2 m_GroupMenu.Image, m_GroupMenu.TopLeft, _
        CreatePoint(Layout.Left, Layout.Top)
        
    DrawImageRect2 m_GroupMenu.Image, m_GroupMenu.TopRight, _
        CreatePoint(layoutRight, Layout.Top)
        
    DrawImageRect2 m_GroupMenu.Image, m_GroupMenu.BottomLeft, _
        CreatePoint(Layout.Left, layoutBottom)
        
    DrawImageRect2 m_GroupMenu.Image, m_GroupMenu.BottomRight, _
        CreatePoint(layoutRight, layoutBottom)

    DrawImageStretchRect m_GroupMenu.Image, _
        CreateRectL(Layout.Left + m_GroupMenu.TopLeft.width, _
                    Layout.Top, _
                    m_GroupMenu.MiddleTop.height, _
                   theWidth), _
        m_GroupMenu.MiddleTop
    
    DrawImageStretchRect m_GroupMenu.Image, _
        CreateRectL(Layout.Left, _
                    Layout.Top + m_GroupMenu.TopLeft.height, _
                    theHeight, _
                    m_GroupMenu.MiddleLeft.width), _
        m_GroupMenu.MiddleLeft
        
    DrawImageStretchRect m_GroupMenu.Image, _
        CreateRectL(layoutRight, _
                    Layout.Top + m_GroupMenu.TopRight.height, _
                    theHeight, _
                    m_GroupMenu.MiddleRight.width), _
        m_GroupMenu.MiddleRight
        
    DrawImageStretchRect m_GroupMenu.Image, _
        CreateRectL(Layout.Left + m_GroupMenu.BottomLeft.width, _
                    layoutBottom, _
                    m_GroupMenu.MiddleBottom.height, _
                   theWidth), _
        m_GroupMenu.MiddleBottom
        
    DrawImageStretchRect m_GroupMenu.Image, _
        CreateRectL(Layout.Left + m_GroupMenu.TopLeft.width, _
                    Layout.Top + m_GroupMenu.Middle.Top, _
                    theHeight, _
                    theWidth), m_GroupMenu.Middle
    
    If update Then UpdateBuffer
    Exit Sub
Handler:
    Debug.Print "DrawList()" & Err.Description
End Sub

Function ReturnPadding() As Long
    ReturnPadding = (LEFT_PADDING + RIGHT_PADDING)
End Function

Sub DrawList(Layout As RECT)
    If (m_currentMode = TextMode) Then
        DrawListText Layout
    Else
        DrawListThumbnail Layout
    End If
End Sub

Sub DrawListText(Layout As RECT)
On Error GoTo Handler

Dim thisWindow As Window

Dim Y As Long
Dim X As Long
Dim BorderWidth As Long

    m_ListPosition = CreatePoint(Layout.Left, Layout.Top)
    m_currentLayout = Layout

    Y = Layout.Top + TEXTMODE_TOP_PADDING + m_GroupMenu.TopLeft.height
    X = Layout.Left + TEXTMODE_LEFT_PADDING
    
    'BorderWidth = (Layout.Right - Layout.Left) - _
                    m_GroupMenu.MiddleLeft.Width - m_GroupMenu.MiddleRight.Width - LEFT_PADDING
    BorderWidth = (Layout.Right - Layout.Left) - _
                    33
    
    DrawBorder Layout, False
    m_Path.Constructor FillModeWinding
    
    For Each thisWindow In m_WindowList
        thisWindow.UpdateWindowText
    
        'm_BitmapGraphics.DrawString thisWindow.Caption, m_Font, m_Brush, CreatePointF(LEFT_PADDING, Y)
        
        If thisWindow Is m_selectedWindow Then
        
            DrawButton CreatePoint(X - 14, Y), m_GroupMenuButton.SelectedSection, BorderWidth, TEXTMODE_BUTTONHEIGHT
            
            If m_mouseOnCloseButton Then
                DrawCloseButton (X + BorderWidth) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN + 10), Y + (m_closeButtonSize.height - 9), 3
            Else
                DrawCloseButton (X + BorderWidth) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN + 10), Y + (m_closeButtonSize.height - 9), 1
            End If
        
        ElseIf thisWindow Is m_hoveredOverWindow Then
        
            DrawButton CreatePoint(X - 14, Y), m_GroupMenuButton.NormalSection, BorderWidth, TEXTMODE_BUTTONHEIGHT
            
            If m_mouseOnCloseButton Then
                DrawCloseButton (X + BorderWidth) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN + 10), Y + (m_closeButtonSize.height - 9), 2
            Else
                DrawCloseButton (X + BorderWidth) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN + 10), Y + (m_closeButtonSize.height - 9), 1
            End If
            
        ElseIf thisWindow.Flashing Then
            Debug.Print "Flashy window draw!"
            DrawButton CreatePoint(X - 14, Y), m_GroupMenuButton.FlashingSection, BorderWidth, TEXTMODE_BUTTONHEIGHT
        End If
        
        'Debug.Print "Window[]; " & thisWindow.hwnd & " - " & GetWindowClassString(thisWindow.hwnd) & " - " & thisWindow.Caption
        m_Path.AddString thisWindow.Caption, m_FontFamily, 0, 12, CreateRectF(X, Y + 3, 12, BorderWidth), 0
        
        Y = Y + TEXTMODE_ITEM_Y_GAP
    Next
    
    m_BitmapGraphics.FillPath m_Brush, m_Path
    
    UpdateBuffer
    Exit Sub
Handler:
    Debug.Print "DrawList()" & Err.Description
End Sub

Private Function ExtraX(length As Long, IncX As Double) As Long

    If length < IncX Then
        ExtraX = (IncX / 2) - (length / 2)
    End If

End Function

Sub DrawListThumbnail(Layout As RECT)
On Error GoTo Handler

Dim thisWindow As Window

Dim ImageY As Long
Dim TextY As Long
Dim TextX As Long
Dim IncrimentX As Double
Dim TextLength As Long

Dim xPrecise As Double
Dim X As Long

Dim BorderWidth As Long
Dim textMaxWidth As Long
Dim thisText As String

    m_ListPosition = CreatePoint(Layout.Left, Layout.Top)
    m_currentLayout = Layout

    TextY = Layout.Top + THUMBNAIL_TOP_PADDING + m_GroupMenu.TopLeft.height
    ImageY = TextY + 21
    
    'BorderWidth = (Layout.Right - Layout.Left) - _
                    m_GroupMenu.MiddleLeft.Width - m_GroupMenu.MiddleRight.Width - LEFT_PADDING
    BorderWidth = (Layout.Right - Layout.Left)
    
    IncrimentX = (BorderWidth - (LEFT_PADDING + RIGHT_PADDING)) / m_WindowList.Count
    xPrecise = Layout.Left + LEFT_PADDING
    
    Debug.Print "IncX:: " & IncrimentX
    
    DrawBorder Layout, False
    m_Path.Constructor FillModeWinding
    
    For Each thisWindow In m_WindowList
        thisWindow.UpdateWindowText
    
        X = CLng(xPrecise)
        'm_BitmapGraphics.DrawString thisWindow.Caption, m_Font, m_Brush, CreatePointF(LEFT_PADDING, Y)
        TextLength = m_BitmapGraphics.MeasureString(thisWindow.Caption, m_Font).width
        
 
        If TextLength < IncrimentX Then
            TextX = X + (IncrimentX / 2) - (TextLength / 2)
        Else
            TextX = X
        End If
        
        thisText = thisWindow.Caption
        textMaxWidth = CLng(IncrimentX)
        
        If thisWindow Is m_hoveredOverWindow Then

            textMaxWidth = (CLng(IncrimentX)) - (m_closeButtonSize.width + 9)
        
            If m_Graphics.MeasureString(thisText, m_Font).width + ExtraX(TextLength, IncrimentX) > textMaxWidth Then

                While Len(thisText) > 0 And m_Graphics.MeasureString(thisText & "...", m_Font).width + ExtraX(TextLength, IncrimentX) > textMaxWidth
                    thisText = Mid$(thisText, 1, Len(thisText) - 1)
                    TextLength = m_BitmapGraphics.MeasureString(thisWindow.Caption, m_Font).width
                Wend
                
                textMaxWidth = (CLng(IncrimentX))
                thisText = thisText & "..."
            End If
        End If
        
        m_Path.AddString thisText, m_FontFamily, 0, 12, CreateRectF(TextX, TextY, 12, textMaxWidth), 0
        
        If thisWindow Is m_selectedWindow Then
        
            If Not m_mouseOnCloseButton Then
                DrawButton CreatePoint(X, TextY - 4), m_GroupMenuButton.SelectedSection, CLng(IncrimentX), (RECTHEIGHT(Layout) - BUTTON_Y_GAP)
            
                DrawCloseButton (X + IncrimentX) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN), Layout.Top + 5 + (m_closeButtonSize.height + CLOSE_BUTTON_MARGIN), 1
            Else
                DrawButton CreatePoint(X, TextY - 4), m_GroupMenuButton.NormalSection, CLng(IncrimentX), (RECTHEIGHT(Layout) - BUTTON_Y_GAP)
            
                DrawCloseButton (X + IncrimentX) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN), Layout.Top + 5 + (m_closeButtonSize.height + CLOSE_BUTTON_MARGIN), 3
            End If
        
        ElseIf thisWindow Is m_hoveredOverWindow Then
        
            
            DrawButton CreatePoint(X, TextY - 4), m_GroupMenuButton.NormalSection, CLng(IncrimentX), (RECTHEIGHT(Layout) - BUTTON_Y_GAP)
            
            'if PointInsideOfRect(
            
            If Not m_mouseOnCloseButton Then
                DrawCloseButton (X + IncrimentX) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN), Layout.Top + 5 + (m_closeButtonSize.height + CLOSE_BUTTON_MARGIN), 1
            Else
                DrawCloseButton (X + IncrimentX) - (m_closeButtonSize.width + CLOSE_BUTTON_MARGIN), Layout.Top + 5 + (m_closeButtonSize.height + CLOSE_BUTTON_MARGIN), 2
            End If
            
        ElseIf thisWindow.Flashing Then
            Debug.Print "Flashy window draw!"
            DrawButton CreatePoint(X, TextY - 4), m_GroupMenuButton.FlashingSection, CLng(IncrimentX), (RECTHEIGHT(Layout) - BUTTON_Y_GAP)
        End If
        
        If thisWindow.Image.width <> 0 Then
            'm_BitmapGraphics.DrawImage thisWindow.Image, x - 14, ImageY, ITEM_WIDTH, ITEM_HEIGHT
            DrawAspectImage CreatePoint(X, ImageY), thisWindow.Image, CLng(IncrimentX)
        Else
            DrawAspectImage CreatePoint(X, ImageY), thisWindow.Parent.Image, CLng(IncrimentX)
            
        End If
        'Debug.Print "Window[]; " & thisWindow.hwnd & " - " & GetWindowClassString(thisWindow.hwnd) & " - " & thisWindow.Caption
        'm_Path.AddString thisWindow.Caption, m_FontFamily, 0, 12, CreateRectF(x, ImageY + 3, 12, BorderWidth), 0
        
        xPrecise = xPrecise + IncrimentX
        'ImageY = ImageY + TEXTMODE_ITEM_Y_GAP
    Next
    
    m_BitmapGraphics.FillPath m_Brush, m_Path
    UpdateBuffer

    Exit Sub
Handler:
    Debug.Print "DrawList()" & Err.Description
End Sub

Private Sub DrawAspectImage(Position As POINTL, srcImage As GDIPImage, ByVal targetWidth As Long)

    If srcImage.width * m_aspectRatio < (targetWidth) Then
        Position.X = Position.X + _
            ((targetWidth / 2)) - (srcImage.width * m_aspectRatio / 2)
    End If

    m_BitmapGraphics.DrawImage srcImage, CSng(Position.X), CSng(Position.Y), _
        srcImage.width * m_aspectRatio, srcImage.height * m_aspectRatio
End Sub

Private Sub DrawButton(Position As POINTL, SourceButtonSection As GroupSection, Optional width As Long = 100, Optional height As Long = 100)

    DrawImageRect2 m_GroupMenuButton.Image, SourceButtonSection.TopLeft, Position
    
    DrawImageStretchRect m_GroupMenuButton.Image, _
        CreateRectL(Position.X + SourceButtonSection.TopLeft.width, _
                    Position.Y, _
                    SourceButtonSection.MiddleTop.height, _
                    (width - SourceButtonSection.TopRight.width - SourceButtonSection.TopLeft.width)), _
        SourceButtonSection.MiddleTop
        
    DrawImageRect2 m_GroupMenuButton.Image, SourceButtonSection.TopRight, _
        CreatePoint(Position.X + (width - SourceButtonSection.TopRight.width), _
                    Position.Y)

    DrawImageStretchRect m_GroupMenuButton.Image, _
        CreateRectL(Position.X, _
                    Position.Y + SourceButtonSection.TopLeft.height, _
                    (height - SourceButtonSection.BottomLeft.height - SourceButtonSection.TopLeft.height), _
                    SourceButtonSection.MiddleLeft.width), _
        SourceButtonSection.MiddleLeft
        
    DrawImageRect2 m_GroupMenuButton.Image, SourceButtonSection.BottomLeft, _
        CreatePoint(Position.X, _
                    Position.Y + (height - SourceButtonSection.BottomLeft.height))
                    
    DrawImageStretchRect m_GroupMenuButton.Image, _
        CreateRectL(Position.X + SourceButtonSection.TopLeft.width, _
                    Position.Y + (height - SourceButtonSection.BottomLeft.height), _
                    SourceButtonSection.MiddleBottom.height, _
                    (width - SourceButtonSection.BottomRight.width - SourceButtonSection.MiddleRight.width)), _
        SourceButtonSection.MiddleBottom
        
    DrawImageRect2 m_GroupMenuButton.Image, SourceButtonSection.BottomRight, _
        CreatePoint(Position.X + (width - SourceButtonSection.BottomRight.width), _
                    Position.Y + (height - SourceButtonSection.BottomRight.height))
                    
    DrawImageStretchRect m_GroupMenuButton.Image, _
        CreateRectL(Position.X + (width - SourceButtonSection.BottomRight.width), _
                    Position.Y + SourceButtonSection.TopRight.height, _
                    (height - SourceButtonSection.BottomRight.height - SourceButtonSection.TopRight.height), _
                    SourceButtonSection.MiddleLeft.width), _
        SourceButtonSection.MiddleRight
        
    DrawImageStretchRect m_GroupMenuButton.Image, _
        CreateRectL(Position.X + SourceButtonSection.TopLeft.width, _
                    Position.Y + SourceButtonSection.TopLeft.height, _
                    (height - SourceButtonSection.BottomLeft.height - SourceButtonSection.TopLeft.height), _
                    (width - SourceButtonSection.TopRight.width - SourceButtonSection.TopLeft.width)), _
        SourceButtonSection.Middle
        
End Sub



Function DrawImageRect2(ByRef Image As GDIPImage, ByRef destRect As RECTL, ByRef Position As POINTL)
    m_BitmapGraphics.DrawImageRect Image, (Position.X), (Position.Y), (destRect.width), (destRect.height), (destRect.Left), (destRect.Top)
End Function

Function DrawImageStretchRect(ByRef Image As GDIPImage, ByRef destRect As RECTL, ByRef SourceRect As RECTL)
    m_BitmapGraphics.DrawImageStretchAttrF Image, _
        RECTLtoF(destRect), _
        SourceRect.Left, SourceRect.Top, SourceRect.width, SourceRect.height, UnitPixel, 0, 0, 0
End Function

Sub UpdateBuffer()
On Error GoTo Handler
    m_Graphics.Clear vbBlack
    m_Graphics.DrawImage _
        m_Bitmap.Image, 0, 0, Me.ScaleWidth, Me.ScaleHeight
        
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_winSize, Me.hdc, m_srcPoint, 0, m_blendFunc32bpp, ULW_ALPHA)
        
    Exit Sub
Handler:
    Debug.Print "UpdateBuffer()" & Err.Description
End Sub

Function ReInitSurface() As Boolean
    On Error GoTo Handler
    
    m_winSize.cx = Me.ScaleWidth
    m_winSize.cy = Me.ScaleHeight
    
    m_Bitmap.Dispose
    m_BitmapGraphics.Dispose

    m_Bitmap.CreateFromSizeFormat Me.ScaleWidth, Me.ScaleHeight, PixelFormat.Format32bppArgb
    m_BitmapGraphics.FromImage m_Bitmap.Image
    
    m_BitmapGraphics.TextRenderingHint = TextRenderingHintAntiAlias
    m_BitmapGraphics.SmoothingMode = SmoothingModeHighQuality
    m_BitmapGraphics.InterpolationMode = InterpolationModeHighQualityBicubic
    m_BitmapGraphics.PixelOffsetMode = PixelOffsetModeHighQuality
    
    m_Graphics.FromHDC Me.hdc
    ReInitSurface = True
    
    Exit Function
Handler:
    ReInitSurface = False
    Debug.Print "ReInitSurface():" & Me.ScaleWidth & " / " & Me.ScaleHeight & vbCrLf & Err.Description
End Function

Private Function tryGetWindow(index As Long) As Window
    On Error GoTo Handler
    
Dim targetWindow As Window
    Set targetWindow = m_WindowList(index)
    Set tryGetWindow = targetWindow
    
    Exit Function
Handler:
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not m_hoveredOverWindow Is Nothing Then

        Set m_selectedWindow = m_hoveredOverWindow
        
        If m_allowListRedraw Then
            DrawList m_currentLayout
        End If
    End If
End Sub

Private Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Handler
    
    If (m_currentMode = TextMode) Then
        MouseMoveText Button, Shift, X, Y
    Else
        MouseMoveThumbnail Button, Shift, X, Y
    End If

    Exit Sub
Handler:
End Sub

Private Sub MouseMoveText(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim newIndex As Long
Dim newWindow As Window
Dim update As Boolean

    'WM_LBUTTONDOWN (Form_MouseDown) isn't sent correctly, but it is here
    If Button = vbLeftButton Or Button = vbRightButton Then
        Form_MouseDown Button, Shift, X, Y
        Exit Sub
    End If
    
    If Not m_isTracking Then
        m_isTracking = True
        ReTrackMouse
    End If

    RaiseEvent onMouseOver
    
    If m_allowListRedraw Then

        X = X - m_ListPosition.X
        Y = Y - m_ListPosition.Y - THUMBNAIL_TOP_PADDING
    
        newIndex = (RoundToSignificance(Y, TEXTMODE_ITEM_Y_GAP) / TEXTMODE_ITEM_Y_GAP)
        Set newWindow = tryGetWindow(newIndex)
    
        If Not newWindow Is Nothing Then
            If Not newWindow Is m_hoveredOverWindow Then
                Set m_hoveredOverWindow = newWindow
                update = True
            End If
            
            If X < RECTWIDTH(m_currentLayout) And X > (RECTWIDTH(m_currentLayout) - (m_closeButtonSize.width + 24)) Then
                If m_mouseOnCloseButton <> True Then
                    update = True
                    m_mouseOnCloseButton = True
                End If
            Else
                If m_mouseOnCloseButton <> False Then
                    update = True
                    m_mouseOnCloseButton = False
                End If
            End If
        End If
        
        If update Then
            DrawListText m_currentLayout
        End If
    End If

End Sub

Private Sub MouseMoveThumbnail(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim newIndex As Long
Dim newWindow As Window
Dim windowWidth As Long

Dim windowX As Long
Dim windowY As Long
Dim update As Boolean

    'WM_LBUTTONDOWN (Form_MouseDown) isn't sent correctly, but it is here
    If Button = vbLeftButton Or Button = vbRightButton Then
        Form_MouseDown Button, Shift, X, Y
        'Exit Sub
    End If
        
    If Not m_isTracking Then
        m_isTracking = True
        ReTrackMouse
    End If

    RaiseEvent onMouseOver
    
    If m_allowListRedraw Then

        X = X - m_ListPosition.X - THUMBNAIL_LEFT_PADDING
        Y = Y - m_ListPosition.Y - THUMBNAIL_TOP_PADDING

        windowWidth = (RECTWIDTH(m_currentLayout) - (LEFT_PADDING + RIGHT_PADDING)) / m_WindowList.Count
        'realWidth = windowWidth + (LEFT_PADDING + RIGHT_PADDING)

        newIndex = RTS2(X, windowWidth)
        Set newWindow = tryGetWindow(newIndex)
        'Debug.Print "newWindow Is Nothing; " & Not newWindow Is Nothing
    
        If Not newWindow Is Nothing Then

            windowX = X - ((newIndex * windowWidth) - windowWidth)
            windowY = Y - 15
            
            If ((windowX - windowWidth) * -1) < 22 And windowY < 22 Then
                If m_mouseOnCloseButton <> True Then
                    update = True
                    m_mouseOnCloseButton = True
                End If
            Else
                If m_mouseOnCloseButton <> False Then
                    update = True
                    m_mouseOnCloseButton = False
                End If
            End If

            'if new window is different
            If Not newWindow Is m_hoveredOverWindow Then
                Set m_hoveredOverWindow = newWindow
                update = True
            End If
            
            If update Then
                DrawListThumbnail m_currentLayout
            End If
        End If
    Else
        Debug.Print "reDraw Blocked!"
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Not m_hoveredOverWindow Is Nothing Then
        If Button = vbLeftButton Then
        
            If m_mouseOnCloseButton Then
                RaiseEvent onClosed(m_hoveredOverWindow)
            Else
                RaiseEvent onClicked(m_hoveredOverWindow)
            End If
                
        ElseIf Button = vbRightButton Then
            m_haltDeActivation = True
            RaiseEvent onRightClicked(m_hoveredOverWindow)
            m_haltDeActivation = False
        End If
    End If
    
    m_mouseOnCloseButton = False
    
    Set m_selectedWindow = Nothing
    DrawList m_currentLayout
End Sub

Private Sub Form_Paint()
    UpdateBuffer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnhookWindow(Me.hWnd)
    
    m_Graphics.Dispose
    m_BitmapGraphics.Dispose

    m_Font.Dispose
    m_Brush.Dispose
    m_GroupMenu.Image.Dispose
    m_GroupMenuButton.Image.Dispose
    m_Bitmap.Dispose

    m_Path.Dispose
    m_FontFamily.Dispose
    m_Pen.Dispose

    DisposeGDIIfLast
End Sub

Private Sub MakeTrans()
'    Exit Sub

Dim curWinLong As Long

    curWinLong = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    SetWindowLong Me.hWnd, GWL_EXSTYLE, curWinLong Or WS_EX_LAYERED Or WS_EX_TOPMOST Or WS_EX_TOOLWINDOW
    
    'update layer window stuff (that will blend and show the GDI+ text)
    m_srcPoint.X = 0
    m_srcPoint.Y = 0
    m_winSize.cx = Me.ScaleWidth
    m_winSize.cy = Me.ScaleHeight

    With m_blendFunc32bpp
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    'now the text will be shown
    Call UpdateLayeredWindow(Me.hWnd, Me.hdc, ByVal 0&, m_winSize, Me.hdc, m_srcPoint, 0, m_blendFunc32bpp, ULW_ALPHA)

End Sub

Private Function IHookSink_WindowProc(hWnd As Long, iMsg As Long, wParam As Long, lParam As Long) As Long
    On Error GoTo Handler

    If iMsg = WM_MOUSELEAVE Then
        m_isTracking = False
        If Not m_haltDeActivation Then ResetRollover
        
        If m_allowListRedraw Then
            DrawList m_currentLayout
        End If
        
        RaiseEvent onMouseOut
    ElseIf iMsg = WM_MOUSEMOVE Then
    
        MouseMove MouseButtonState(wParam), ShiftState(wParam), LOWORD(lParam), HiWord(lParam)
        
    ElseIf iMsg = WM_ACTIVATE Then
        
        If wParam = WA_INACTIVE Then
            ResetRollover
        
            If Not m_haltDeActivation = True Then
                If Not ApplicationOptions.Floating Then RaiseEvent onDeactivated(lParam)
            End If
        End If
        
         ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            CallOldWindowProcessor(hWnd, iMsg, wParam, lParam)
    Else
         ' Just allow default processing for everything else.
         IHookSink_WindowProc = _
            CallOldWindowProcessor(hWnd, iMsg, wParam, lParam)
    End If

    Exit Function
Handler:
    LogError Err.number, "WindowProc(" & iMsg & "," & wParam & "," & lParam & "); " & Err.Description, "winSubMenu"
    
    ' Just allow default processing for everything else.
    IHookSink_WindowProc = _
        CallOldWindowProcessor(hWnd, iMsg, wParam, lParam)
End Function

Sub ReTrackMouse()
    m_isTracking = True

    Dim ET As TrackMouseEvent
    'initialize structure
    ET.cbSize = Len(ET)
    ET.hwndTrack = Me.hWnd
    ET.dwFlags = &H2&
    'start the tracking
    TrackMouseEvent ET
End Sub


