Attribute VB_Name = "RectHelper"
'--------------------------------------------------------------------------------
'    Component  : RectHelper
'    Project    : prjSuperBar
'
'    Description: Contains RECT helper functions
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function CreateRect(Left As Long, Top As Long, Bottom As Long, Right As Long) As RECT

Dim newRect As RECT
    With newRect
        .Left = Left
        .Top = Top
        .Bottom = Bottom
        .Right = Right
    End With
    
    CreateRect = newRect
End Function

Public Function CreateRectL(Left As Long, Top As Long, height As Long, width As Long) As RECTL

Dim newRect As RECTL
    With newRect
        .Left = Left
        .Top = Top
        .height = height
        .width = width
    End With
    
    CreateRectL = newRect
End Function

Public Function CreateRectF(Left As Long, Top As Long, height As Long, width As Long) As RECTF

Dim newRectF As RECTF

    With newRectF
        .Left = Left
        .Top = Top
        .height = height
        .width = width
    End With
    
    CreateRectF = newRectF
End Function

Public Function RECTWIDTH(ByRef srcRect As RECT)
    RECTWIDTH = srcRect.Right - srcRect.Left
End Function

Public Function RECTHEIGHT(ByRef srcRect As RECT)
    RECTHEIGHT = srcRect.Bottom - srcRect.Top
End Function

Public Function PrintRectF(ByRef srcRect As RECTF)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & _
                "Left; " & srcRect.Left & vbCrLf & _
                "Height; " & srcRect.height & vbCrLf & _
                "Width; " & srcRect.width
End Function

Public Function PrintRect(ByRef srcRect As RECT)
    Debug.Print "Top; " & srcRect.Top & vbCrLf & _
                "Left; " & srcRect.Left & vbCrLf & _
                "Bottom; " & srcRect.Bottom & vbCrLf & _
                "Right; " & srcRect.Right
End Function

Public Function RECTtoF(ByRef srcRECTL As RECT) As RECTF
    RECTtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.Bottom), CLng(srcRECTL.Right))
End Function

Public Function RECTFtoL(ByRef srcRect As RECTF) As RECT
    RECTFtoL = CreateRect(CLng(srcRect.Left), CLng(srcRect.Top), CLng(srcRect.height), CLng(srcRect.width))
End Function

Public Function RECTLtoF(ByRef srcRECTL As RECTL) As RECTF
    RECTLtoF = CreateRectF(CLng(srcRECTL.Left), CLng(srcRECTL.Top), CLng(srcRECTL.height), CLng(srcRECTL.width))
End Function

Public Function PointInsideOfRect(srcPoint As win.POINTL, srcRect As win.RECT) As Boolean

    PointInsideOfRect = False

    If srcPoint.Y > srcRect.Top And _
       srcPoint.Y < srcRect.Bottom And _
       srcPoint.X > srcRect.Left And _
       srcPoint.X < srcRect.Right Then

       PointInsideOfRect = True
    End If


End Function

