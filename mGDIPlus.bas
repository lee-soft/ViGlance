Attribute VB_Name = "mGDIPlus"
Option Explicit

Private m_LastError As Long
Public g_IgnoreGDIErrors As Boolean


Public Function GetErrorStatus() As GpStatus
    GetErrorStatus = m_LastError
End Function

' Use this in lieu of the Color class constructor
' Thanks to Richard Mason for help with this
'Public Function ColorARGB(ByVal Alpha As Byte, ByVal Red As Byte, ByVal Green As Byte, ByVal Blue As Byte) As Long
'   Dim bytestruct As COLORBYTES
'   Dim result As COLORLONG
'
'   With bytestruct
'      .AlphaByte = Alpha
'      .RedByte = Red
'      .GreenByte = Green
'      .BlueByte = Blue
'   End With
'
'   LSet result = bytestruct
'   ColorARGB = result.longval
'End Function

'Public Function CreateMatrix(V1 As Single, V2 As Single, V3 As Single, V4 As Single, V5 As Single, _
'                             W1 As Single, W2 As Single, W3 As Single, W4 As Single, W5 As Single, _
'                             X1 As Single, X2 As Single, X3 As Single, X4 As Single, X5 As Single, _
'                             Y1 As Single, Y2 As Single, Y3 As Single, Y4 As Single, Y5 As Single, _
'                             Z1 As Single, Z2 As Single, Z3 As Single, Z4 As Single, Z5 As Single) As ColorMatrix
'
'Dim clrMatrix As ColorMatrix
'
'    clrMatrix.m(0, 0) = V1: clrMatrix.m(1, 0) = V2: clrMatrix.m(2, 0) = V3: clrMatrix.m(3, 0) = V4: clrMatrix.m(4, 0) = V5
'    clrMatrix.m(0, 1) = W1: clrMatrix.m(1, 1) = W2: clrMatrix.m(2, 1) = W3: clrMatrix.m(3, 1) = W4: clrMatrix.m(4, 1) = W5
'    clrMatrix.m(0, 2) = X1: clrMatrix.m(1, 2) = X2: clrMatrix.m(2, 2) = X3: clrMatrix.m(3, 2) = X4: clrMatrix.m(4, 2) = X5
'    clrMatrix.m(0, 3) = Y1: clrMatrix.m(1, 3) = Y2: clrMatrix.m(2, 3) = Y3: clrMatrix.m(3, 3) = Y4: clrMatrix.m(4, 2) = Y5
'    clrMatrix.m(0, 4) = Z1: clrMatrix.m(1, 4) = Z2: clrMatrix.m(2, 4) = Z3: clrMatrix.m(3, 4) = Z4: clrMatrix.m(4, 4) = Z5
'
'    CreateMatrix = clrMatrix
'End Function

