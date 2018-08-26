Attribute VB_Name = "GeneralHelper"
'--------------------------------------------------------------------------------
'    Component  : GeneralHelper
'    Project    : prjSuperBar
'
'    Description: A place miscellaneous API declerations and functions to live
'                 TODO: Seperate this stuff into their own helper modules
'
'--------------------------------------------------------------------------------
Option Explicit

Public Declare Function CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb As Long) As Long

Public Declare Sub DragAcceptFiles Lib "shell32" (ByVal hWnd As Long, ByVal _
    bool As Integer)
Public Declare Function DragQueryFileW Lib "shell32" (ByVal wParam As Long, _
    ByVal index As Long, ByVal lpszFile As Long, ByVal BufferSize As Long) _
    As Long
Public Declare Sub DragFinish Lib "shell32" (ByVal hDrop As Integer)

Public Declare Function Wow64DisableWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long
Public Declare Function Wow64RevertWow64FsRedirection Lib "kernel32.dll" (ByRef oldValue As Long) As Long

Public Declare Function PrintWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Public Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal hdcDst As Long, pptDst As Any, PSIZE As Any, ByVal hdcSrc As Long, pptSrc As Any, ByVal crKey As Long, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Long) As Long
Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TrackMouseEvent) As Long

Public Declare Function RegisterShellHookWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32.dll" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Public Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, ByRef pData As APPBARDATA) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function GetLastError Lib "kernel32.dll" () As Long

Public Declare Function MakeSureDirectoryPathExists Lib _
        "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Function ShellExecuteExW Lib "shell32.dll" (lpExecInfo As SHELLEXECUTEINFOW) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal iFileHandle As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Private Type SHELLEXECUTEINFOW
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    ' fields
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    HICON As Long
    hProcess As Long
End Type

Private Type KEYBDINPUT
    wVk As Integer
    wScan As Integer
    dwFlags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Type GENERALINPUT
    dwType As Long
    xi(0 To 23) As Byte
End Type

Public Const RDW_ALLCHILDREN = &H80
Public Const RDW_ERASE = &H4
Public Const RDW_INVALIDATE = &H1
Public Const RDW_UPDATENOW = &H100
'Public Const RDW_INTERNALPAINT = &H2
'Public Const RDW_VALIDATE = &H8

'Public Const WM_SYSMENU As Long = &H313
Public Const ABM_GETTASKBARPOS As Long = &H5

Public Enum AbeBarEnum
   abe_bottom = 3
   ABE_LEFT = 0
   ABE_RIGHT = 2
   ABE_TOP = 1
End Enum

'Public Const AW_CENTER As Long = &H10
'Public Const AW_SLIDE As Long = &H40000
'Public Const AW_HIDE As Long = &H10000
'Public Const AW_BLEND As Long = &H80000
'Public Const AW_VER_NEGATIVE As Long = &H8
'Public Const AW_VER_POSITIVE As Long = &H4

Public Const HSHELL_REDRAW As Long = 6
'Public Const HSHELL_HIGHBIT = &H8000
Public Const HSHELL_FLASH = 32774
Public Const HSHELL_WINDOWDESTROYED As Long = 2
Public Const HSHELL_WINDOWCREATED As Long = 1
Public Const HSHELL_WINDOWACTIVATED As Long = 4


'Public Const IDANI_OPEN = &H1
'Public Const IDANI_CLOSE = &H2
'Public Const IDANI_CAPTION = &H3

'Public Const ULW_OPAQUE = &H4
'Public Const ULW_COLORKEY = &H1
Public Const ULW_ALPHA = &H2
Public Const AC_SRC_ALPHA As Long = &H1
Public Const AC_SRC_OVER = &H0

Public Const WS_EX_LAYERED = &H80000

Public Const GCL_HICON = (-14)
'Public Const GCL_HICONSM = (-34)

Public Const TME_LEAVE As Long = &H2
'Public Const WS_EX_NOACTIVATE As Long = &H8000000

Public Const SMTO_BLOCK = &H1
Public Const SMTO_ABORTIFHUNG = &H2

Private Const KEYEVENTF_KEYUP = &H2
Private Const INPUT_KEYBOARD = 1
'Private Const INPUT_HARDWARE = 2

' API Defined Types
Public Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type

Public Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Public Type Size
    cx As Long
    cy As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type argb
    r As Byte
    g As Byte
    b As Byte
    a As Byte
End Type

Public Sub Main()
    Set g_DeviceCollection = New DeviceCollection

    ' Must call this before using any GDI+ call:
    If Not (GDIPlusCreate(True)) Then
        Exit Sub
    End If

    InitializeAndCheckOptionsPath
    SetOpenSaveDocs

    App.TaskVisible = False
    OptionsHelper.ReadOptions
    
    If App.PrevInstance Then
        Exit Sub
    End If
    
    If Command$ <> vbNullString Then
        If DecideOptionsByCommandLine = True Then
            Exit Sub
        End If
    End If

    If Not ApplicationOptions.DontShowSplash Then
        ShowSplashForm
    End If
    
    Load frmClock
    Load frmWindowChecker
    frmZOrderKeeper.Show
    'frmZOrderKeeper.DoCommand
End Sub

Private Function InitializeAndCheckOptionsPath()

Dim optionsPath As String
    
    Set AppProfile = New WinProfile
    
    optionsPath = Environ$("appdata") & "\" & App.ProductName
    
    OptionsHelper.EnsureFolderExists optionsPath
    AppProfile.INIPath = optionsPath & "\options.ini"
End Function

Private Function ShowSplashForm()

    With frmSplash
        .Caption = ""
    
        .Label3.FontSize = 9
        .Label2.FontSize = 9
        .Label1.FontSize = 9
    
        .lblLink.Top = 1280
        .lblLink.Left = 1620
        
        .Label5.Visible = False
        .Label6.Visible = False
        '.Label4.Visible = False
        
        .lblBottom.Visible = False
        
        .lblEmail.Caption = ""
        .lblEmail.Top = 2000
        .lblEmail.FontBold = True

        .height = 1880
        .width = 4700
        
        .Show
    End With

End Function

Private Function DecideOptionsByCommandLine() As Boolean
    
Dim strParameters() As String
Dim paramIndex As Long
Dim strParameter As String

    OptionsHelper.ApplicationOptions.Floating = False
    OptionsHelper.ApplicationOptions.AutoClick = True
    OptionsHelper.ApplicationOptions.InstantSpawn = False
    OptionsHelper.ApplicationOptions.DontShowSplash = True
    OptionsHelper.ApplicationOptions.ViOrb = False
    OptionsHelper.ApplicationOptions.GlideAnimation = True
    OptionsHelper.ApplicationOptions.TaskBarFade = False
    OptionsHelper.ApplicationOptions.TextOnlyMode = False
    
    strParameters = Split(CStr(Command), " ")
    
    For paramIndex = LBound(strParameters) To UBound(strParameters)
        strParameter = strParameters(paramIndex)
        
        Select Case strParameter
        
        Case "/nosplash"
            OptionsHelper.ApplicationOptions.DontShowSplash = False
            
        Case "/noautoclick"
            OptionsHelper.ApplicationOptions.AutoClick = False
        
        Case "/instantspawn"
            OptionsHelper.ApplicationOptions.InstantSpawn = True
            OptionsHelper.ApplicationOptions.AutoClick = False
            
        Case "/noglide"
            OptionsHelper.ApplicationOptions.GlideAnimation = False
            
        Case "/taskbarfade"
            OptionsHelper.ApplicationOptions.TaskBarFade = True
        
        Case "/floating"
            OptionsHelper.ApplicationOptions.Floating = True
            
        Case "/?"
            MessageBox 0, "/nospash - Disables the splash screen" & vbCrLf & _
                          "/noautoclick - Disables autoclicking of groups (use with /instantspawn)" & vbCrLf & _
                          "/instantspawn - Programmer's intent; have sub menu's instantly spawn" & vbCrLf & _
                          "/noglide - Disables gliding animations on Superbar" & vbCrLf & _
                          "/taskbarfade - Enables fading animation on Windows Taskbar", _
                          "Command Line Arguments", MB_OK Or MB_ICONINFORMATION
            
            DecideOptionsByCommandLine = True
        
        Case Else
            MessageBox 0, strParameter & " isn't recognised", "Error", MB_OK
        
        End Select
    Next
    
    
End Function

Public Sub Long2ARGB(ByVal LongARGB As Long, ByRef argb As argb)
    CopyMemory argb, LongARGB, 4
End Sub

Public Function IsStyle( _
      ByVal lAll As Long, _
      ByVal lBit As Long) As Boolean
      
   IsStyle = False
   If (lAll And lBit) = lBit Then
      IsStyle = True
   End If
End Function

Public Function Exists(col, index) As Boolean
On Error GoTo ExistsTryNonObject
    Dim o As Object

    Set o = col(index)
    Exists = True
    Exit Function

ExistsTryNonObject:
    Exists = ExistsNonObject(col, index)
End Function

Private Function ExistsNonObject(col, index) As Boolean
On Error GoTo ExistsNonObjectErrorHandler
    Dim v As Variant

    v = col(index)
    ExistsNonObject = True
    Exit Function

ExistsNonObjectErrorHandler:
    ExistsNonObject = False
End Function

Public Function RTS2(ByVal number As Long, _
    ByVal significance As Long)
    
'Round number up or down to the nearest multiple of significance
Dim d As Double
    
    number = number + (significance / 2)
    d = number / significance
    d = Round(d, 0)
    RTS2 = d
End Function

Public Function RoundToSignificance(ByVal number As Integer, _
ByVal significance As Integer) As Integer
    'Round number up or down to the nearest multiple of significance
    Dim d As Double
    d = number / significance
    d = Round(d, 0)
    RoundToSignificance = d * significance
End Function

Public Function DisposeGDIIfLast()
    If Forms.Count = 1 Then
        GDIPlusDispose
    End If
End Function

Private Function TryFont(FontName As String) As Boolean

Dim testFont As New GDIPFont
    
    On Error GoTo FontNotExist
    
    testFont.Depreciated_Constructor FontName, 14
    'If Not GDIPlusWrapper.GetLastErrorStatus() = Ok Then GoTo FontNotExist
    
    testFont.Dispose
    TryFont = True
    
    Exit Function
FontNotExist:
    TryFont = False
End Function

Public Function GetClosestVistaFont()
    If TryFont("Segoe UI") = True Then
        GetClosestVistaFont = "Segoe UI"
        Exit Function
    End If
    
    If TryFont("Tahoma") = True Then
        GetClosestVistaFont = "Tahoma"
        Exit Function
    End If
    
    GetClosestVistaFont = "Arial"
End Function

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long

 'Combines two integers into a long integer
  MAKELPARAM = MakeLong(wLow, wHigh)
  
End Function

Public Function MakeLong(wLow As Long, wHigh As Long) As Long

  MakeLong = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
  
End Function

Public Function GetZOrder(ByVal hWndTarget As Long) As Long
    
Dim hWnd As Long
Dim lngZOrder As Long

    ' Loop through window list and
    ' compare to hWnd to hwndTarget, to find global ZOrder
    hWnd = GetTopWindow(0)
    lngZOrder = 0
    
    Do While hWnd And hWnd <> hWndTarget
       ' Get next window and move on.
        hWnd = GetNextWindow(hWnd, _
          GW_HWNDNEXT)
        lngZOrder = lngZOrder + 1
        
        'Debug.Print lngZOrder & ";" & GetWindowClassString(hwnd) & ";" & GetWindowNameString(hwnd)
    Loop
    
    GetZOrder = lngZOrder

End Function

Public Function IsMouseLeftButtonDown() As Boolean
    IsMouseLeftButtonDown = (GetAsyncKeyState(vbKeyLButton) And &H8000)
End Function

Public Function CreateSystemMenu(ByVal hMenu As Long, ByVal WindowState As FormWindowStateConstants) As clsMenu

Dim itemCount As Long
Dim objNewMenu As New clsMenu
Dim itemIndex As Long
Dim itemID As Long
Dim bufferString As String
Dim itemLength As Long
Dim itemState As Long
Dim menuDefault As Long

    itemCount = GetMenuItemCount(hMenu)

    For itemIndex = 0 To itemCount - 1
        itemLength = GetMenuString(hMenu, itemIndex, ByVal 0, 0, MF_BYPOSITION) + 1
        bufferString = String$(itemLength, 0)
        itemState = GetMenuState(hMenu, itemIndex, MF_BYPOSITION)
        itemID = GetMenuItemID(hMenu, itemIndex)
        GetMenuString hMenu, itemIndex, bufferString, itemLength, MF_BYPOSITION
        
        If WindowState = vbNormal Then
            If itemID = SC_RESTORE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_CLOSE Then
                menuDefault = itemID
            End If
            
        ElseIf WindowState = vbMinimized Then
            If itemID = SC_MOVE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_SIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_MINIMIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_RESTORE Then
                itemState = MF_STRING
                menuDefault = itemID
            ElseIf itemID = SC_MAXIMIZE Then
                itemState = MF_STRING
            End If
            
        ElseIf WindowState = vbMaximized Then
            If itemID = SC_MAXIMIZE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_MOVE Then
                itemState = MF_GRAYED Or MF_STRING
            ElseIf itemID = SC_SIZE Then
                itemState = MF_GRAYED Or MF_STRING
            End If
        End If
        
        
        AppendMenu objNewMenu.Handle, itemState, itemID, bufferString
    
        If itemID = menuDefault Then
            SetMenuDefaultItem objNewMenu.Handle, itemIndex, True
        End If
    Next
    
    Set CreateSystemMenu = objNewMenu

End Function

Public Function FileExists(sSource As String, Optional ByVal allowFsDirection As Boolean = True) As Boolean
    If sSource = vbNullString Then
        Exit Function
    End If

    Dim WFD As WIN32_FIND_DATA
    Dim hFile As Long
    
    hFile = FindFirstFile(sSource, WFD)
    FileExists = hFile <> INVALID_HANDLE_VALUE
    
    Call FindClose(hFile)
   
    If FileExists = False And Is64bit And allowFsDirection = False Then
        Dim win64Token As Win64FSToken: Set win64Token = New Win64FSToken
        FileExists = FileExists(sSource, True)
        win64Token.EnableFS
    End If
End Function

Public Function MouseButtonState(wParam As Long) As MouseButtonConstants
    
    MouseButtonState = 0
    
    If wParam = MK_LBUTTON Then
        MouseButtonState = vbLeftButton
    ElseIf wParam = MK_RBUTTON Then
        MouseButtonState = vbRightButton
    End If
    
End Function

Public Function ShiftState(wParam As Long) As Boolean
    
    ShiftState = False
    If wParam And MK_SHIFT Then
        ShiftState = True
    End If
    
End Function

Public Function SetKeyDown(KeyCode As Long)

Dim GInput(0 To 1) As GENERALINPUT
Dim KInput As KEYBDINPUT

    KInput.wVk = KeyCode 'the key we're going to press
    KInput.dwFlags = 0 'press the key
    'copy the structure into the input array's buffer.
    GInput(0).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(0).xi(0), KInput, Len(KInput)
    
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))


End Function

Public Function SetKeyUp(KeyCode As Long)

Dim GInput(0 To 1) As GENERALINPUT
Dim KInput As KEYBDINPUT



    'do the same as above, but for releasing the key
    KInput.wVk = KeyCode ' the key we're going to realease
    KInput.dwFlags = KEYEVENTF_KEYUP ' release the key
    GInput(1).dwType = INPUT_KEYBOARD ' keyboard input
    CopyMemory GInput(1).xi(0), KInput, Len(KInput)
    'send the input now
    Call SendInput(2, GInput(0), Len(GInput(0)))

End Function

Public Function isset(srcAny) As Boolean

    On Error GoTo Handler

Dim thisVarType As VbVarType: thisVarType = VarType(srcAny)

    If thisVarType = vbObject Then
        If Not srcAny Is Nothing Then
            isset = True
            Exit Function
        End If
    ElseIf thisVarType = vbArray Or _
           thisVarType = 8200 Then
           
            If UBound(srcAny) > 0 Then
                isset = True
                Exit Function
            End If
    Else
        isset = IsEmpty(srcAny)
        Exit Function
    End If

Handler:
    isset = False

End Function

Public Function GetWord(ByVal strVal As String) As Long
    Dim Lo As Long
    Dim Hi As Long
    
    Lo = AscB(MidB$(strVal, 1, 1))
    Hi = AscB(MidB$(strVal, 2, 1))
    
    GetWord = (Hi * 256) + Lo
End Function

Public Function GetDWord(ByVal strVal As String) As Double
    Dim LoWord As Single
    Dim HiWord As Single
    If LenB(strVal) <> 4 Then
        GetDWord = 0
        Exit Function
    End If
    
    LoWord = GetWord(MidB$(strVal, 1, 2))
    HiWord = GetWord(MidB$(strVal, 3, 2))
    GetDWord = (HiWord * 65536) + LoWord
End Function

' The state of either Shift keys
Function ShiftKey() As Boolean
    ShiftKey = (GetAsyncKeyState(vbKeyShift) And &H8000)
End Function

Public Function ShellEx(ByVal strPath As String) As bool
    
Dim ShellExInfo As SHELLEXECUTEINFOW

Dim strEXE As String
Dim strParam As String

Dim lngFirstQ As Long
Dim lngSecondQ As Long
    
    On Error GoTo Handler
    
    If strPath = vbNullString Then
        ShellEx = APIFALSE
        Exit Function
    End If
    
    If InStr(strPath, """") Then
        lngFirstQ = InStr(strPath, """") + 1
        lngSecondQ = InStr(lngFirstQ, strPath, """")
        
        If lngFirstQ = 2 And (lngSecondQ < Len(strPath)) Then
            '"The.EXE" - Paramaters
            
            strParam = Mid$(strPath, lngSecondQ + 1)
            strEXE = Mid$(strPath, lngFirstQ, lngSecondQ - lngFirstQ)
        ElseIf lngFirstQ > 2 Then
            'The.EXE "The Parameters"
            
            strEXE = Mid$(strPath, 1, lngFirstQ - 2)
            strParam = Mid$(strPath, lngFirstQ - 1)
        End If
    Else
        'The.EXE
        strEXE = strPath
    End If
    
    ShellExInfo.lpFile = StrPtr(strEXE)
    ShellExInfo.lpParameters = StrPtr(strParam)
    
    ShellExInfo.cbSize = Len(ShellExInfo)
    ShellExInfo.nShow = SW_SHOWNORMAL
    
    ShellEx = ShellExecuteExW(ShellExInfo)

    Exit Function
Handler:
    LogError 0, Err.Description, "GeneralHelper::ShellEx"

End Function

Public Function StrEnd(ByVal sData As String, ByVal sDelim As String, Optional iOffset As Integer = 1)

    If InStr(sData, sDelim) = 0 Then
        'Delim not present
    
        StrEnd = sData
        Exit Function
    End If

Dim iLen As Integer, iDLen As Integer

    iLen = Len(sData) + 1
    iDLen = Len(sDelim)

    If iLen = 1 Or iDLen = 0 Then
        StrEnd = False
        Exit Function
    End If

    While Mid$(sData, iLen, iDLen) <> sDelim And iLen > 1
        iLen = iLen - 1
    Wend

    If iLen = 0 Then
        StrEnd = False
        Exit Function
    End If
    
    StrEnd = Mid$(sData, iLen + iOffset)

End Function
