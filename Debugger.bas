Attribute VB_Name = "Debugger"
Option Explicit

Public Sub LogError(ByVal lNum As Long, ByVal sDesc As String, ByVal sFrom As String)
    
    Debug.Print "APP ERROR; " & sDesc & " ; " & sFrom
    Dim FileNum As Integer

    FileNum = FreeFile
    Open App.Path & "\errors.log" For Append As FileNum
        Write #FileNum, lNum, sDesc, sFrom, Now()
    Close FileNum
End Sub
