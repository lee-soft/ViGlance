VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlphaIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--------------------------------------------------------------------------------
'    Component  : AlphaIcon
'    Project    : prjSuperBar
'
'    Description: A Transparent Icon
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_Icon As GDIPBitmap
Private m_IconXP As GDIPBitmap

Private m_IsAlphaBitmap As Boolean

Public Property Get Image()
    Set Image = m_IconXP.Image
End Property

Public Function CreateFromHICON(ByVal icoHandle As Long)

Dim ii As ICONINFO

Dim bmData As BitmapData
Dim bmBounds As RECTL

Dim X As Long
Dim Y As Long

    Set m_Icon = New GDIPBitmap
    Set m_IconXP = New GDIPBitmap

    GetIconInfo icoHandle, ii
    m_Icon.CreateFromHBITMAP ii.hbmColor, 0
    
    GeneralHelper.DeleteObject ii.hbmColor
    GeneralHelper.DeleteObject ii.hbmMask

    bmBounds.width = m_Icon.Image.width
    bmBounds.height = m_Icon.Image.height

    bmData = m_Icon.LockBits(bmBounds, _
                ImageLockModeRead, _
                m_Icon.Image.ImgPixelFormat)

    m_IconXP.CreateFromSizeFormatData bmData.height, _
                bmData.width, _
                bmData.Stride, _
                Format32bppArgb, _
                bmData.Scan0Ptr
        
        
        
    m_IsAlphaBitmap = False

    For Y = 0 To bmData.height - 1
        For X = 0 To bmData.width - 1
            
            Dim PixelColor As argb
            Long2ARGB m_IconXP.GetPixel(X, Y), PixelColor
                'PixelColor = Color.FromArgb(Marshal.ReadInt32(bmData.Scan0, (bmData.Stride * y) + (4 * x)))
            If PixelColor.a > 0 And PixelColor.a < 255 Then
                m_IsAlphaBitmap = True
                Exit For
            End If
        Next
        If m_IsAlphaBitmap Then Exit For
    Next

    m_Icon.UnlockBits bmData
        
    If Not m_IsAlphaBitmap Then
        m_Icon.Dispose
        m_IconXP.Dispose
        
        m_IconXP.CreateFromHICON icoHandle
    End If
    
End Function
