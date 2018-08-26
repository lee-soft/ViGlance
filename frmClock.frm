VERSION 5.00
Begin VB.Form frmClock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "00:00 AM"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmClock
'    Project    : prjSuperBar
'
'    Description: System tray icon
'
'--------------------------------------------------------------------------------
Option Explicit

Private m_PopupSystemMenu As clsMenu
Private nid As NOTIFYICONDATA

Private Sub Form_Initialize()
    Set m_PopupSystemMenu = New clsMenu
    'Set timWindowChecker = Controls.Add("VB.Timer", "timWindowChecker", Me)
    
    m_PopupSystemMenu.AddItem 1, "&Exit"
    m_PopupSystemMenu.AddSeperater
    m_PopupSystemMenu.AddItem 3, "&Options"
    m_PopupSystemMenu.AddItem 2, "&About"

End Sub

Public Sub ReInstallIcon()
    If Not OptionsHelper.ApplicationOptions.HideTrayIcon Then Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_Load()
'the form must be fully visible before calling Shell_NotifyIcon
    'Me.Show
    'Me.Refresh
    
Dim theTip() As Byte: theTip = App.Title & vbNullChar
Dim theTipIndex As Long
    
    With nid
     .cbSize = Len(nid)
     .hWnd = Me.hWnd
     .uID = App.hInstance
     .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
     .uCallbackMessage = WM_MOUSEMOVE
     .HICON = Me.Icon
     
    For theTipIndex = 0 To UBound(theTip)
        .szTip(theTipIndex) = theTip(theTipIndex)
    Next
    
    End With
    
    ReInstallIcon
    ShowWindow Me.hWnd, SW_HIDE
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this procedure receives the callbacks from the System Tray icon.
Dim msg As Long

    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
     msg = X
    Else
     msg = X / Screen.TwipsPerPixelX
    End If
    
    Select Case msg
      
     Case WM_RBUTTONUP        '517 display popup menu
        SetForegroundWindow Me.hWnd 'dont care about the result
        
      Select Case m_PopupSystemMenu.ShowMenu(Me.hWnd)
      
      Case 1
        Unload frmTaskbar
        Exit Sub
    
      Case 2
        frmSplash.Show
        Exit Sub
        
      Case 3
        frmOptions.Show
        Exit Sub
      
      End Select
      
    End Select
End Sub

Private Sub Form_Resize()
 'this is necessary to assure that the minimized window is hidden
 On Error Resume Next
 
 If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'this removes the icon from the system tray
 Shell_NotifyIcon NIM_DELETE, nid
End Sub
