VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JumpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_mruKey As Collection

Public Function AddMRURegKey(thisMRUKey As WinRegistryKey)
    m_mruKey.Add thisMRUKey
End Function

Public Function GetMRUList()

Dim thisMRUKey As WinRegistryKey
Dim strLnks()

    For Each thisMRUKey In m_mruKey
        strLnks = ConcatArray(strLnks, GetMRUListForKey(thisMRUKey))
    Next
    
    GetMRUList = UniqueValues(strLnks)

End Function

Private Sub Class_Initialize()
    Set m_mruKey = New Collection
End Sub
