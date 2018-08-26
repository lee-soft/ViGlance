VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   248
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkStartWindows 
      Caption         =   "Start &ViGlance with Windows"
      Height          =   495
      Left            =   180
      TabIndex        =   5
      Top             =   2520
      Width           =   4215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Pinned Items"
      Height          =   420
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   420
      Left            =   2160
      TabIndex        =   8
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   3300
      TabIndex        =   7
      Top             =   3180
      Width           =   1095
   End
   Begin VB.CheckBox chkGlide 
      Caption         =   "Enable Glide &Animation"
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CheckBox chkGroupWindows 
      Caption         =   "Automatically Show &Group Submenu"
      Height          =   495
      Left            =   180
      TabIndex        =   4
      Top             =   2040
      Width           =   4215
   End
   Begin VB.CheckBox chkTransparency 
      Caption         =   "Enable &Window's Taskbar Transparency"
      Height          =   495
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   4215
   End
   Begin VB.CheckBox chkAeroPeek 
      Caption         =   "Enable Window &Thumbnails (Up-To-Date)"
      Height          =   495
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.CheckBox chkViOrb 
      Caption         =   "Enable &Start Orb (ViOrb)"
      Height          =   495
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : frmOptions
'    Project    : prjSuperBar
'
'    Description: The options window. GUI for configuring app preferences
'
'--------------------------------------------------------------------------------
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()

    Dim emptyByteStringArray() As ShellLink
    ApplicationOptions.PinnedApplications = emptyByteStringArray
    
    OptionsHelper.RebootApplication

End Sub

Private Sub cmdOK_Click()
    OptionsHelper.ApplicationOptions.ViOrb = IIf(chkViOrb.Value = vbChecked, True, False)
    OptionsHelper.ApplicationOptions.TextOnlyMode = IIf(chkAeroPeek.Value = vbChecked, False, True)
    OptionsHelper.ApplicationOptions.TaskBarFade = IIf(chkTransparency.Value = vbChecked, True, False)
    OptionsHelper.ApplicationOptions.AutoClick = IIf(chkGroupWindows.Value = vbChecked, True, False)
    OptionsHelper.ApplicationOptions.GlideAnimation = IIf(chkGlide.Value = vbChecked, True, False)
    OptionsHelper.StartWithWindows = IIf(chkStartWindows.Value = vbChecked, True, False)
    
    OptionsHelper.RebootApplication
End Sub

Private Sub Form_Load()
    chkViOrb.Value = IIf(OptionsHelper.ApplicationOptions.ViOrb, vbChecked, vbUnchecked)
    chkAeroPeek.Value = IIf(OptionsHelper.ApplicationOptions.TextOnlyMode, vbUnchecked, vbChecked)
    chkTransparency.Value = IIf(OptionsHelper.ApplicationOptions.TaskBarFade, vbChecked, vbUnchecked)
    chkGroupWindows.Value = IIf(OptionsHelper.ApplicationOptions.AutoClick, vbChecked, vbUnchecked)
    chkGlide.Value = IIf(OptionsHelper.ApplicationOptions.GlideAnimation, vbChecked, vbUnchecked)
    chkStartWindows.Value = IIf(OptionsHelper.StartWithWindows, vbChecked, vbUnchecked)
End Sub
