Attribute VB_Name = "ArrayHelper"
'--------------------------------------------------------------------------------
'    Component  : ArrayHelper
'    Project    : prjSuperBar
'
'    Description: Contains Array Helper functions
'
'--------------------------------------------------------------------------------
Option Explicit

Public Function ConcatArray(ByRef a, ByRef b)

Dim n As Long, m As Long
Dim c()

Dim firstItemIndex As Long
Dim lastItemIndex As Long

    n = 0
    
    If IsArrayInitialized(a) Then
        firstItemIndex = LBound(a)
        lastItemIndex = UBound(a)
    
        For m = firstItemIndex To lastItemIndex
            ReDim Preserve c(n)
            
            c(n) = a(m)
            n = n + 1
        Next
    End If
    
    If IsArrayInitialized(b) Then
        firstItemIndex = LBound(b)
        lastItemIndex = UBound(b)
    
        For m = firstItemIndex To lastItemIndex
            ReDim Preserve c(n)
            
            c(n) = b(m)
            n = n + 1
        Next
    End If
    
    ConcatArray = c

End Function

Public Function In_Array(ByRef a, ByRef sValue) As Boolean

Dim m As Long

Dim firstItemIndex As Long
Dim lastItemIndex As Long

    If Not IsArrayInitialized(a) Then
        In_Array = False
        Exit Function
    End If
    
    firstItemIndex = LBound(a)
    lastItemIndex = UBound(a)
    
    For m = firstItemIndex To lastItemIndex
        If (a(m) = sValue) Then
            In_Array = True
            Exit Function
        End If
    Next

End Function

Public Function IsArrayInitialized(myArray) As Boolean

Dim mySize As Long

    On Error Resume Next
    mySize = UBound(myArray) ' In this instance the error number is set as myArray has a size of -1!

    If (Err.number <> 0) Then
        IsArrayInitialized = False
    Else
        If mySize > -1 Then
            IsArrayInitialized = True
        End If
    End If

End Function

Public Function SizeOf(srcArray) As Long
On Error GoTo Handler
    
    SizeOf = UBound(srcArray)
    Exit Function
Handler:
    SizeOf = 0
    
End Function

Public Function UniqueValues(ByRef heyStack)

Dim m As Long
Dim n As Long
Dim newArray()

Dim firstItemIndex As Long
Dim lastItemIndex As Long

    If Not IsArrayInitialized(heyStack) Then
        Exit Function
    End If
    
    firstItemIndex = LBound(heyStack)
    lastItemIndex = UBound(heyStack)
    
    For m = firstItemIndex To lastItemIndex
        If Not In_Array(newArray, heyStack(m)) Then
            ReDim Preserve newArray(n)
            
            newArray(n) = heyStack(m)
            n = n + 1
        End If
    Next
    
    UniqueValues = newArray

End Function
