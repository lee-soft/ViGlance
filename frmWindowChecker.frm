VERSION 5.00
Begin VB.Form frmWindowChecker 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timWindowChecker 
      Interval        =   1000
      Left            =   840
      Top             =   720
   End
End
Attribute VB_Name = "frmWindowChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmWindowChecker
'    Project    : prjSuperBar
'
'    Description: Checks that our Windows/Forms have not been unloaded by
'                 explorer. Reloads them if they've been unloaded
'
'--------------------------------------------------------------------------------
Option Explicit



Private Sub timWindowChecker_Timer()
    If IsWindow(frmZOrderKeeper.hWnd) = APIFALSE Then
    
        Set frmOptions = Nothing
        Set frmZOrderKeeper = Nothing
        Set frmTaskbar = Nothing
        Set frmSubMenu = Nothing
        Set frmStartButton = Nothing
        Set frmFader = Nothing
        Set frmClock = Nothing
        Set frmSplash = Nothing
        
        RebootApplication
    End If
    
    
End Sub
