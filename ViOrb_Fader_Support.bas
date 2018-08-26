Attribute VB_Name = "ViOrb_Fader_Support"
Option Explicit

Private Const CLASS_NAME As String = "ViOrb_Fader_Class"

Public g_WndClassViOrbFader As WNDCLASSEX

Private m_RegisterdClass As Boolean
Private m_Handler As ViOrb_Fader

Public Function InitializeAndRegisterClass(ByRef newViOrbFader As ViOrb_Fader)
    If m_RegisterdClass = True Then Exit Function
    
    m_RegisterdClass = True
    Set m_Handler = newViOrbFader

    ' Fill in our WNDCLASSEX structure
    With g_WndClassViOrbFader
      .cbSize = Len(g_WndClassViOrbFader)
      .style = 0
      .lpfnWndProc = GetFunctionPtr(AddressOf ViOrb_Fader_Support.WndProc)
      .cbClsExtra = 0
      .cbWndExtra = 0
      .hInstance = App.hInstance
      .hIcon = LoadIcon(App.hInstance, IDI_APPLICATION)
      .hCursor = LoadCursor(App.hInstance, IDC_ARROW)
      .hbrBackground = GetStockObject(LTGRAY_BRUSH)
      .lpszMenuName = vbNullString
      .lpszClassName = CLASS_NAME
      .hIconSm = LoadIcon(App.hInstance, IDI_APPLICATION)
    End With
    
    RegisterClassEx g_WndClassViOrbFader
End Function

' **********************************************************************
' FUNCTION: WndProc
' PURPOSE: This is the Window Procedure for our class, it's purpose
' is to provide handling for all the messages we wish to respond
' to. All other messages are sent on to the Windows Default
' message handler.
' **********************************************************************
Public Function WndProc(ByVal hwnd As Long, ByVal message As Long, _
            ByVal wParam As Long, ByVal lParam As Long) As Long
            
On Error Resume Next
    WndProc = m_Handler.WndProc(hwnd, message, wParam, lParam)
End Function

