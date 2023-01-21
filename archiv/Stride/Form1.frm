VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6885
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim i As Long, j As Long, n As Long, bpp As Long
    Dim u As Long: u = 100
    Dim om As Long, fs As Long, b As Boolean
    bpp = 1
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    bpp = 4
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    bpp = 8
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    bpp = 16
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    bpp = 24
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    bpp = 32
    n = n + u
    j = 0
    For i = i To n
        om = CalcStrideOM(j, bpp)
        fs = CalcStrideFS(j, bpp)
        b = (om = fs)
        List1.AddItem "i: " & i & "; j: " & j & "; bpp: " & bpp & "; om: " & om & "; fs: " & fs & "; b: " & b
        j = j + 1
    Next
    
End Sub

Private Function CalcStrideOM(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
    'calculates the number of bytes in one horizontal line of pixels,
    'including the number of pad-bytes for the 4-aligned result
    CalcStrideOM = W * BitsPerPixel \ 8
    Dim m As Long: m = CalcStrideOM Mod 4: If m > 0 Then m = 4 - m
    CalcStrideOM = CalcStrideOM + m
End Function

Private Function CalcStrideFS(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
    Dim BytesPerPixel As Long: BytesPerPixel = BitsPerPixel \ 8
    Select Case BitsPerPixel
    Case 1:
    Case 4:
    Case 8:        CalcStrideFS = (W + 3) And Not 3
    Case 16, 24:   CalcStrideFS = ((W * BytesPerPixel) + BytesPerPixel) And Not BytesPerPixel
    Case 32:       CalcStrideFS = W * BytesPerPixel
    End Select
End Function

