Attribute VB_Name = "IconHelper"
'--------------------------------------------------------------------------------
'    Component  : IconHelper
'    Project    : prjSuperBar
'
'    Description: Utility class for retrieving icons program icons
'                 in various ways
'
'--------------------------------------------------------------------------------
Option Explicit

'Private Const ILD_TRANSPARENT = &H1 'display transparent

Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'system icon index
Private Const SHGFI_LARGEICON = &H0 'large icon
Private Const SHGFI_SMALLICON = &H1 'small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
                                 SHGFI_SHELLICONSIZE Or _
                                 SHGFI_SYSICONINDEX Or _
                                 SHGFI_DISPLAYNAME Or _
                                 SHGFI_EXETYPE

Public Function GetIconFromHwnd(hWnd As Long) As Long

    Call SendMessageTimeout(hWnd, WM_GETICON, ICON_BIG, 0, 0, 100, GetIconFromHwnd)
    If Not CBool(GetIconFromHwnd) Then GetIconFromHwnd = GetClassLong(hWnd, GCL_HICON)
    If Not CBool(GetIconFromHwnd) Then Call SendMessageTimeout(hWnd, WM_GETICON, 1, 0, 0, 100, GetIconFromHwnd)
    If Not CBool(GetIconFromHwnd) Then GetIconFromHwnd = GetClassLong(hWnd, GCL_HICON)
    If Not CBool(GetIconFromHwnd) Then Call SendMessageTimeout(hWnd, WM_QUERYDRAGICON, 0, 0, 0, 100, GetIconFromHwnd)
End Function

Public Function GetApplicationIcon(strExePath As String) As Long
Dim shinfo As SHFILEINFO
Dim win64Token As Win64FSToken

    If Is64bit Then
        If (InStr(LCase$(strExePath), LCase$(Environ$("windir"))) > 0) Then
            Set win64Token = New Win64FSToken
        End If
    End If
    
    'get the system icon associated with that file
    SHGetFileInfo strExePath, 0&, _
                    shinfo, Len(shinfo), _
                    BASIC_SHGFI_FLAGS Or SHGFI_ICON
    
    GetApplicationIcon = shinfo.HICON
    If Not win64Token Is Nothing Then
        win64Token.EnableFS
    End If
End Function
