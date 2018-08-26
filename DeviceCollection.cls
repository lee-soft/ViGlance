VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeviceCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'    Component  : DeviceCollection
'    Project    : prjSuperBar
'
'    Description: Converts paths from device paths to drive paths
'
'--------------------------------------------------------------------------------
Option Explicit

Private Declare Function QueryDosDeviceW Lib "kernel32.dll" ( _
    ByVal lpDeviceName As Long, _
    ByVal lpTargetPath As Long, _
    ByVal ucchMax As Long _
    ) As Long
    
Private Const MAX_PATH As Long = 260
Private m_drives As Collection

Public Function ConvertToLetterPath(strPath As String)

Dim letterIndex As Long: letterIndex = 1
Dim delimNumber As Long: delimNumber = 3
Dim devicePath As String

    If Left$(strPath, 8) = "\Device\" Then
        While delimNumber > 0
            letterIndex = InStr(letterIndex, strPath, "\") + 1
            delimNumber = delimNumber - 1
        Wend
        
        devicePath = Mid$(strPath, 1, letterIndex - 2)
        strPath = Chr$(m_drives(devicePath)) & ":\" & Mid$(strPath, letterIndex)
        
    End If
    
    ConvertToLetterPath = strPath

End Function

Private Sub Class_Initialize()

Dim szLinkName(MAX_PATH) As Byte
Dim szDevName(MAX_PATH) As Byte
Dim bDrive As Byte: bDrive = 0
Dim sDevName As String
    
    Set m_drives = New Collection
    szLinkName(2) = AscW(":")
    
    'this isn't ideal, consider itterating properly, not lazily
    'AA,AB etc could also be drives so this really doesn't work
    
    For bDrive = AscB("a") To AscB("z")
        szLinkName(0) = bDrive
        
        If QueryDosDeviceW(VarPtr(szLinkName(0)), VarPtr(szDevName(0)), MAX_PATH) Then
            sDevName = szDevName
            m_drives.Add szLinkName(0), sDevName
        End If
    Next

End Sub


