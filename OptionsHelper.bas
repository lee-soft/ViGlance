Attribute VB_Name = "OptionsHelper"
Option Explicit

Private Const WINDOWS_REGRUN As String = "Software\Microsoft\Windows\CurrentVersion\Run\"

' ############# TEXTMODE CONSTANTS ###############
Public Const TEXTMODE_BUTTONHEIGHT As Long = 25
Public Const TEXTMODE_TOP_PADDING As Long = 7
Public Const TEXTMODE_LEFT_PADDING As Long = 28
Public Const TEXTMODE_ITEM_Y_GAP As Long = 30

' ############# THUMBNAILMODE CONSTANTS ##########
Public Const THUMBNAIL_TOP_PADDING As Long = 5
Public Const THUMBNAIL_LEFT_PADDING As Long = 20
'Public Const THUMBNAIL_RIGHT_PADDING As Long = 26
Public Const THUMBNAIL_ITEM_Y_GAP As Long = 30

Public Const WINDOW_IMAGE_HEIGHT As Long = 120
Public Const WINDOW_IMAGE_WIDTH As Long = 187

Public Const SUBMENU_Y_PADDING As Long = 85
'Public Const SUBMENU_X_PADDING As Long = 50

Public Const JUMPLIST_CAP As Long = 9

Public g_hwndForeGroundWindow As Long

Public Enum WindowMode
    TextMode = 1
    ThumbnailMode = 2
End Enum

Public Type APP_OPTIONS
    Floating As Boolean
    DontShowSplash As Boolean
    AutoClick As Boolean
    InstantSpawn As Boolean
    ViOrb As Boolean
    GlideAnimation As Boolean
    TaskBarFade As Boolean
    TextOnlyMode As Boolean
    HideTrayIcon As Boolean
    
    PopupDelay As Long
    PopupNextDelay As Long
    ReCaptureImageDelay As Long
    HideChildrenDelay As Long
    
    PinnedApplications() As ShellLink
End Type

Public ApplicationOptions As APP_OPTIONS
Public g_DeviceCollection As DeviceCollection
Public AppProfile As WinProfile

Sub RebootApplication()
Dim F As Form
    For Each F In Forms
        Unload F
    Next

    Set frmZOrderKeeper = Nothing
    Set frmTaskbar = Nothing
    Set frmSubMenu = Nothing
    Set frmStartButton = Nothing
    Set frmFader = Nothing
    Set frmClock = Nothing
    Set frmSplash = Nothing
    
    DumpOptions
    GeneralHelper.Main
End Sub

Public Function CreatePoint(X As Long, Y As Long) As POINTL

Dim newPoint As POINTL
    With newPoint
        .X = X
        .Y = Y
    End With

    CreatePoint = newPoint
End Function

Public Function EnsureFolderExists(ByVal pathToCreate As String) _
  As Boolean
    
    Dim sSomePath As String
    Dim bAns As Boolean: bAns = False
    
   sSomePath = pathToCreate
    
    If CreatePath(sSomePath) = True Then
        bAns = True
    Else
        bAns = False
    End If
EnsureFolderExists = bAns
End Function

Private Function CreatePath(newPath) As Boolean
    Dim sPath As String
    'Add a trailing slash if none
    sPath = newPath & IIf(Right$(newPath, 1) = "\", "", "\")

    'Call API
    If MakeSureDirectoryPathExists(sPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

Public Function ReadOptions()

Dim lastPinnedApp As Long
Dim shellLinkIndex As Long

    ApplicationOptions.AutoClick = INIBool(AppProfile.ReadINIValue("General", "AutoClick", INIVal(True)))
    ApplicationOptions.DontShowSplash = INIBool(AppProfile.ReadINIValue("General", "DontShowSplash", INIVal(False)))
    ApplicationOptions.Floating = INIBool(AppProfile.ReadINIValue("General", "Floating", INIVal(False)))
    ApplicationOptions.GlideAnimation = INIBool(AppProfile.ReadINIValue("General", "GlideAnimation", INIVal(True)))
    ApplicationOptions.InstantSpawn = INIBool(AppProfile.ReadINIValue("General", "InstantSpawn", INIVal(False)))
    ApplicationOptions.TaskBarFade = INIBool(AppProfile.ReadINIValue("General", "TaskBarFade", INIVal(True)))
    ApplicationOptions.TextOnlyMode = INIBool(AppProfile.ReadINIValue("General", "TextOnlyMode", INIVal(False)))
    ApplicationOptions.ViOrb = INIBool(AppProfile.ReadINIValue("General", "ViOrb", INIVal(False)))
    ApplicationOptions.HideTrayIcon = INIBool(AppProfile.ReadINIValue("General", "HideTrayIcon", INIVal(False)))
    
    ApplicationOptions.ReCaptureImageDelay = INILong(AppProfile.ReadINIValue("Timers", "ReCaptureImageDelay", INIVal(1500)))
    ApplicationOptions.PopupNextDelay = INILong(AppProfile.ReadINIValue("Timers", "PopupNextDelay", INIVal(500)))
    ApplicationOptions.HideChildrenDelay = INILong(AppProfile.ReadINIValue("Timers", "HideChildrenDelay", INIVal(500)))
    ApplicationOptions.PopupDelay = INILong(AppProfile.ReadINIValue("Timers", "PopupDelay", INIVal(1000)))
    
    lastPinnedApp = INILong(AppProfile.ReadINIValue("General", "LastPinned", -1))

    If lastPinnedApp > -1 Then
        ReDim ApplicationOptions.PinnedApplications(lastPinnedApp)
    
        For shellLinkIndex = 0 To lastPinnedApp
            ApplicationOptions.PinnedApplications(shellLinkIndex).szPath = _
                    AppProfile.ReadINIValue("ShellLink_" & CStr(shellLinkIndex), "Path", "")
                    
            ApplicationOptions.PinnedApplications(shellLinkIndex).szArguments = _
                    AppProfile.ReadINIValue("ShellLink_" & CStr(shellLinkIndex), "Arguments", "")
        Next
    End If

End Function

Public Function GetLastPinnedApp() As Long
    On Error GoTo UnInitialized
    
    GetLastPinnedApp = CLng(UBound(ApplicationOptions.PinnedApplications))
    Exit Function
UnInitialized:
    GetLastPinnedApp = -1
End Function

Public Function INILong(Value As String) As Long
On Error GoTo Handler

    If IsNumeric(Value) Then
        INILong = CLng(Value)
    End If

    Exit Function
Handler:
    INILong = -1

End Function

Public Function INIBool(Value As String) As Boolean

    If Value = 1 Then
        INIBool = True
    Else
        INIBool = False
    End If

End Function

Public Function INIVal(Value)

    If VarType(Value) = vbBoolean Then
        If Value = True Then
            INIVal = 1
        ElseIf Value = False Then
            INIVal = 0
        End If
        
        Exit Function
    Else
        INIVal = CStr(Value)
    End If

End Function

Public Function DumpOptions()
    On Error GoTo CatchException
    
    Dim shellLnkId As Long
    
    AppProfile.WriteINIValue "General", "AutoClick", INIVal(ApplicationOptions.AutoClick)
    AppProfile.WriteINIValue "General", "DontShowSplash", INIVal(ApplicationOptions.DontShowSplash)
    AppProfile.WriteINIValue "General", "Floating", INIVal(ApplicationOptions.Floating)
    AppProfile.WriteINIValue "General", "GlideAnimation", INIVal(ApplicationOptions.GlideAnimation)
    AppProfile.WriteINIValue "General", "InstantSpawn", INIVal(ApplicationOptions.InstantSpawn)
    AppProfile.WriteINIValue "General", "TaskBarFade", INIVal(ApplicationOptions.TaskBarFade)
    AppProfile.WriteINIValue "General", "TextOnlyMode", INIVal(ApplicationOptions.TextOnlyMode)
    AppProfile.WriteINIValue "General", "ViOrb", INIVal(ApplicationOptions.ViOrb)
    AppProfile.WriteINIValue "General", "HideTrayIcon", INIVal(ApplicationOptions.HideTrayIcon)
    
    AppProfile.WriteINIValue "General", "LastPinned", GetLastPinnedApp()
    
    AppProfile.WriteINIValue "Timers", "ReCaptureImageDelay", INIVal(ApplicationOptions.ReCaptureImageDelay)
    AppProfile.WriteINIValue "Timers", "PopupDelay", INIVal(ApplicationOptions.PopupDelay)
    AppProfile.WriteINIValue "Timers", "HideChildrenDelay", INIVal(ApplicationOptions.HideChildrenDelay)
    AppProfile.WriteINIValue "Timers", "PopupNextDelay", INIVal(ApplicationOptions.PopupNextDelay)
    
    Dim lastPinnedApp As Long: lastPinnedApp = GetLastPinnedApp()
    
    For shellLnkId = 0 To lastPinnedApp
        AppProfile.WriteINIValue "ShellLink_" & shellLnkId, "Path", ApplicationOptions.PinnedApplications(shellLnkId).szPath
        AppProfile.WriteINIValue "ShellLink_" & shellLnkId, "Arguments", ApplicationOptions.PinnedApplications(shellLnkId).szArguments
    Next shellLnkId

    Exit Function
CatchException:
    MsgBox "Error dumping options file. Settings will not be preserved", vbCritical, "Saving Options Failed!"
    
End Function

Public Property Get StartWithWindows() As Boolean

    If RegistryHelper.ReadReg(HKEY_CURRENT_USER, WINDOWS_REGRUN, "ViGlance") = App.Path & "\" & App.EXEName & ".exe" Then
        StartWithWindows = True
    Else
        StartWithWindows = False
    End If

End Property

Public Property Let StartWithWindows(ByVal Value As Boolean)
    If Value = True Then
        RegistryHelper.WriteReg HKEY_CURRENT_USER, WINDOWS_REGRUN, "ViGlance", App.Path & "\" & App.EXEName & ".exe"
    Else
        RegistryHelper.DeleteReg HKEY_CURRENT_USER, WINDOWS_REGRUN, "ViGlance"
    End If
End Property

