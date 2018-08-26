Attribute VB_Name = "RegistryHelper"
'--------------------------------------------------------------------------------
'    Component  : RegistryHelper
'    Project    : prjSuperBar
'
'    Description: A cheap and nasty way to write/read registry strings.
'                 TODO: Remove this module and use WinRegistryKey instead
'
'--------------------------------------------------------------------------------
Option Explicit

Private Declare Function SHDeleteValue Lib "shlwapi.dll" Alias "SHDeleteValueA" (ByVal hKey As Long, ByVal pszSubKey As String, ByVal pszValue As String) As Long

Sub DeleteReg(H_KEY&, RSubKey$, ValueName$)
    SHDeleteValue H_KEY, RSubKey, ValueName$
End Sub

Sub WriteReg(H_KEY As Long, RSubKey As String, ValueName$, RegValueStr$)
    'H_KEY must be one of the Key Constants
    Dim lRtn&         'returned by registry functions, should be 0&
    Dim hKey&         'return handle to opened key
    Dim lpDisp&
    Dim Sec_Att As SECURITY_ATTRIBUTES
    Dim RegValue() As Byte
    
    If RegValueStr = "" Then RegValueStr = " "
    RegValue = CStr(RegValueStr$)
    
    Sec_Att.nLength = 12&
    Sec_Att.lpSecurityDescriptor = 0&
    Sec_Att.bInheritHandle = False
    
    
    lRtn = RegCreateKeyEx(H_KEY, RSubKey, 0&, "", 0&, KEY_WRITE, Sec_Att, hKey, lpDisp)
    If lRtn <> 0 Then
        Exit Sub       'No key open, so leave
    End If
    lRtn = RegSetValueEx(hKey, ValueName, 0&, REG_SZ, RegValue(0), CLng(UBound(RegValue) + 1))
    lRtn = RegCloseKey(hKey)
End Sub

Function ReadReg(MainKey&, SubKey$, Value$) As Variant
   ' MainKey must be one of the Publicly declared HKEY constants.
   Dim sKeyType As EREGTYPE      'to return the key type.  This function expects REG_SZ or REG_DWORD
   Dim ret&            'returned by registry functions, should be 0&
   Dim lpHKey&         'return handle to opened key
   Dim lpcbData&       'length of data in returned string
   Dim ReturnedString$ 'returned string value
   Dim ReturnedLong&   'returned long value
   
   If MainKey >= &H80000000 And MainKey <= &H80000006 Then
        ' Open key
        ret = RegOpenKeyEx(MainKey, SubKey, 0&, KEY_READ, lpHKey)
        
        If ret <> ERROR_SUCCESS Then
            ReadReg = ""
            Exit Function     'No key open, so leave
        End If
          
        ' Set up buffer for data to be returned in.
        ' Adjust next value for larger buffers.
        lpcbData = 255
        ReturnedString = Space$(lpcbData)
    
        ' Read key
        ret& = RegQueryValueEx(lpHKey, Value, ByVal 0&, sKeyType, StrPtr(ReturnedString), lpcbData)
    
        If ret <> ERROR_SUCCESS Then
            ReadReg = ""   'Value probably doesn't exist
        Else
            If sKeyType = REG_DWORD Then
                ret = RegQueryValueEx(lpHKey, Value, ByVal 0&, sKeyType, ReturnedLong, 4)
                If ret = ERROR_SUCCESS Then ReadReg = ReturnedLong
            Else
                ReadReg = MidB$(ReturnedString, 1, lpcbData - 2)
            End If
        End If
        ' Always close opened keys.
        ret = RegCloseKey(lpHKey)
    End If
End Function

