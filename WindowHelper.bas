Attribute VB_Name = "WindowHelper"
'--------------------------------------------------------------------------------
'    Component  : WindowHelper
'    Project    : prjSuperBar
'
'    Description: Utility module for native Windows. Containing helper functions
'                 for Windows
'
'--------------------------------------------------------------------------------
Option Explicit

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, _
    ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" _
    (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
    ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long


'--------------------------------------------------------------------------------
' Procedure  :       hWndBelongToUs
' Description:       Checks if a given Window handle is a handle from this app
' Parameters :       hWnd (Long)
'                    ExceptionHwnd (Long = 0)
'--------------------------------------------------------------------------------
Public Function hWndBelongToUs(hWnd As Long, Optional ExceptionHwnd As Long = 0) As Boolean

Dim thisForm As Form
    hWndBelongToUs = False

    For Each thisForm In Forms
        If thisForm.hWnd = hWnd Then
            If hWnd = ExceptionHwnd Then
                hWndBelongToUs = False
            Else
                hWndBelongToUs = True
            End If
            
            Exit For
        End If
    Next
    
End Function

Public Sub RepaintWindow(ByRef hWnd As Long)
    'verified it works
    If IsWindowHung(hWnd) Then Exit Sub
    
    If hWnd <> 0 Then
        Call RedrawWindow(hWnd, ByVal 0&, ByVal 0&, _
             RDW_ERASE Or RDW_INVALIDATE Or RDW_ALLCHILDREN Or RDW_UPDATENOW)
    End If
    
End Sub


'--------------------------------------------------------------------------------
' Procedure  :       IsWindowHung
' Description:       Checks to see if a Window has hung
' Parameters :       hWnd (Long)
'--------------------------------------------------------------------------------
Public Function IsWindowHung(hWnd As Long) As Boolean

Dim lResult As Long
Dim lReturn As Long
    
    lReturn = SendMessageTimeout(hWnd, _
                        WM_NULL, _
                        0&, _
                        0&, _
                        SMTO_ABORTIFHUNG Or SMTO_BLOCK, _
                        1000, _
                        lResult)
                     
    If lReturn Then
        IsWindowHung = False
        Exit Function
    End If
    
    IsWindowHung = True

End Function

Public Function ShowWindowTimeout(ByRef hWnd As Long, ByRef nCmdShow As ESW)
    If Not IsWindowHung(hWnd) Then
        ShowWindow hWnd, nCmdShow
    End If
End Function

Function SetOwner(ByVal HwndtoUse As Long, ByVal HwndofOwner As Long) As Long
    SetOwner = SetWindowLong(HwndtoUse, GWL_HWNDPARENT, HwndofOwner)
End Function

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .width / Screen.TwipsPerPixelX
        iHeight = .height / Screen.TwipsPerPixelY
    End With
    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hWnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub


