VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myG As New GDIPGraphics

Private Sub Command1_Click()
    Dim W As New Window
    'Dim I As New GDIPImage
    
    'I.FromFile "D:\Users\lee\Documents\SuperBar\Resources\start_button.png"
    
    W.hWnd = Me.hWnd
    
    W.UpdateImage
    
    myG.FromHDC Me.hDC
    myG.Clear vbBlack
    myG.DrawImage W.Image, 0, 0, 100, 100
End Sub

Private Sub Form_Load()
    ' Must call this before using any GDI+ call:
    If Not (GDIPlusCreate()) Then
        Exit Sub
    End If
    

    


    
End Sub
