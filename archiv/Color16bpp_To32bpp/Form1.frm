VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "16bpp RGB565"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "16bpp ARGB1555"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   31
      Left            =   7320
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   30
      Left            =   6840
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   29
      Left            =   6360
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   28
      Left            =   5880
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   27
      Left            =   5400
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   26
      Left            =   4920
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   25
      Left            =   4440
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   24
      Left            =   3960
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   23
      Left            =   3480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   22
      Left            =   3000
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   21
      Left            =   2520
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   20
      Left            =   2040
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   19
      Left            =   1560
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   18
      Left            =   1080
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   17
      Left            =   600
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   16
      Left            =   120
      Top             =   1080
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   15
      Left            =   7320
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   14
      Left            =   6840
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   13
      Left            =   6360
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   12
      Left            =   5880
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   11
      Left            =   5400
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   10
      Left            =   4920
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   9
      Left            =   4440
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   8
      Left            =   3960
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   7
      Left            =   3480
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   6
      Left            =   3000
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   5
      Left            =   2520
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   4
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   3
      Left            =   1560
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   2
      Left            =   1080
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   1
      Left            =   600
      Top             =   600
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      BorderStyle     =   0  'Transparent
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TByt2
    Value0 As Byte
    Value1 As Byte
End Type
Private Type TByt4
    Value0 As Byte
    Value1 As Byte
    Value2 As Byte
    Value3 As Byte
End Type
Private Type TInt
    Value As Integer
End Type
Private Type TLng
    Value As Long
End Type

Private m_data() As Byte

Private RedMask   As Long
Private GreenMask As Long
Private BlueMask  As Long
Private AlphaMask As Long

Private Function Color16bppBGRA1555_ToColor32(ByVal Value As Long) As Long
'     BlueMask = &H1F    ' 5 bit
'    GreenMask = &H3E0   ' 5 bit
'      RedMask = &H7C00  ' 5 bit
'    AlphaMask = &H8000  ' 1 bit
    Dim r As Long, g As Long, b As Long, a As Long
    'Const Mask5Bit As Long = &H1F&
    'Const Mask6Bit As Long = &H2F&
    b = ((Value And BlueMask) * 256) \ &H1F '1F ' Mask5Bit
    g = (((Value And GreenMask) / BlueMask) * 256) \ &H20 'Mask5Bit
    r = (((Value And RedMask) / GreenMask) * 256) \ &H20 'Mask5Bit
    a = ((Value And AlphaMask) / RedMask) * 256 ' alpha is only 1 bit, so it is 0 or 1, resp 0 or 255
    Color16bppBGRA1555_ToColor32 = RGB(r, g, b)
End Function

Private Function Color16bppBGR565_ToColor32(ByVal Value As Long) As Long
'     BlueMask = &H1F    ' 5 bit
'    GreenMask = &H7E0   ' 6 bit
'      RedMask = &HF800  ' 5 bit
'    AlphaMask = &H0     ' ----
    Dim r As Long, g As Long, b As Long, a As Long
    'Const Mask5Bit As Long = &H1F&
    'Const Mask6Bit As Long = &H3F&
    'Const ShL11    As Long = &H7FF&
    b = ((Value And BlueMask) * 256) \ &H1F
    g = (((Value And GreenMask) / BlueMask) * 256) \ &H40
    r = (((Value And RedMask) / GreenMask) * 256) \ &H20
    'a = (Value And AlphaMask) / RedMask * 255 ' alpha is not included
    Color16bppBGR565_ToColor32 = RGB(r, g, b)
End Function

Sub SetColors_BGRA1555()
    ReDim m_data(0 To 63) As Byte
    Dim i As Long
    m_data(i) = &HFF: i = i + 1:    m_data(i) = &H7F: i = i + 1 'FF 7F
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H0:  i = i + 1 '00 00
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H7C: i = i + 1 '00 7C
    m_data(i) = &HE0: i = i + 1:    m_data(i) = &H3:  i = i + 1 'E0 03
    m_data(i) = &H1F: i = i + 1:    m_data(i) = &H0:  i = i + 1 '1F 00
    m_data(i) = &HE0: i = i + 1:    m_data(i) = &H7F: i = i + 1 'E0 7F
    m_data(i) = &H1F: i = i + 1:    m_data(i) = &H7C: i = i + 1 '1F 7C
    m_data(i) = &HFF: i = i + 1:    m_data(i) = &H3:  i = i + 1 'FF 03
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H40: i = i + 1 '00 40
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H2:  i = i + 1 '00 02
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H0:  i = i + 1 '10 00
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H40: i = i + 1 '10 40
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H42: i = i + 1 '00 42
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H2:  i = i + 1 '10 02
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H42: i = i + 1 '10 42
    m_data(i) = &HF7: i = i + 1:    m_data(i) = &H5E: i = i + 1 'F7 5E

    
    m_data(i) = &H3D: i = i + 1:    m_data(i) = &H53: i = i + 1 '3D 53
    m_data(i) = &HB9: i = i + 1:    m_data(i) = &H53: i = i + 1 'B9 53
    m_data(i) = &H9D: i = i + 1:    m_data(i) = &H66: i = i + 1 '9D 66
    m_data(i) = &H99: i = i + 1:    m_data(i) = &H76: i = i + 1 '99 76
    m_data(i) = &H34: i = i + 1:    m_data(i) = &H77: i = i + 1 '34 77
    m_data(i) = &HB4: i = i + 1:    m_data(i) = &H67: i = i + 1 'B4 67
    m_data(i) = &H22: i = i + 1:    m_data(i) = &H66: i = i + 1 '22 66
    m_data(i) = &H51: i = i + 1:    m_data(i) = &H64: i = i + 1 '51 64
    m_data(i) = &H22: i = i + 1:    m_data(i) = &H47: i = i + 1 '22 47
    m_data(i) = &H59: i = i + 1:    m_data(i) = &H44: i = i + 1 '59 44
    m_data(i) = &H39: i = i + 1:    m_data(i) = &HA:  i = i + 1 '39 0A
    m_data(i) = &H31: i = i + 1:    m_data(i) = &HB:  i = i + 1 '31 0B
    m_data(i) = &H7E: i = i + 1:    m_data(i) = &H6:  i = i + 1 '7E 06
    m_data(i) = &HA7: i = i + 1:    m_data(i) = &H7A: i = i + 1 'A7 7A
    m_data(i) = &H41: i = i + 1:    m_data(i) = &H7F: i = i + 1 '41 7F
    m_data(i) = &HAC: i = i + 1:    m_data(i) = &H7F: i = i + 1 'AC 7F
End Sub

Sub SetColors_BGR565()
    ReDim m_data(0 To 63) As Byte
    Dim i As Long
    m_data(i) = &HFF: i = i + 1:    m_data(i) = &HFF: i = i + 1 'FF FF
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H0:  i = i + 1 '00 00
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &HF8: i = i + 1 '00 F8
    m_data(i) = &HE0: i = i + 1:    m_data(i) = &H7:  i = i + 1 'E0 07
    m_data(i) = &H1F: i = i + 1:    m_data(i) = &H0:  i = i + 1 '1F 00
    m_data(i) = &HE0: i = i + 1:    m_data(i) = &HFF: i = i + 1 'E0 FF
    m_data(i) = &H1F: i = i + 1:    m_data(i) = &HF8: i = i + 1 '1F F8
    m_data(i) = &HFF: i = i + 1:    m_data(i) = &H7:  i = i + 1 'FF 07
    m_data(i) = &H0:  i = i + 1:    m_data(i) = &H80: i = i + 1 '00 80
    m_data(i) = &H20: i = i + 1:    m_data(i) = &H4:  i = i + 1 '20 04
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H0:  i = i + 1 '10 00
    m_data(i) = &H10: i = i + 1:    m_data(i) = &H80: i = i + 1 '10 80
    m_data(i) = &H20: i = i + 1:    m_data(i) = &H84: i = i + 1 '20 84
    m_data(i) = &H30: i = i + 1:    m_data(i) = &H4:  i = i + 1 '30 04
    m_data(i) = &H30: i = i + 1:    m_data(i) = &H84: i = i + 1 '30 84
    m_data(i) = &HF7: i = i + 1:    m_data(i) = &HBD: i = i + 1 'F7 BD
    
    m_data(i) = &H7D: i = i + 1:    m_data(i) = &HA6: i = i + 1 '7D A6
    m_data(i) = &H79: i = i + 1:    m_data(i) = &HA7: i = i + 1 '79 A7
    m_data(i) = &H3D: i = i + 1:    m_data(i) = &HCD: i = i + 1 '3D CD
    m_data(i) = &H39: i = i + 1:    m_data(i) = &HED: i = i + 1 '39 ED
    m_data(i) = &H74: i = i + 1:    m_data(i) = &HEE: i = i + 1 '74 EE
    m_data(i) = &H74: i = i + 1:    m_data(i) = &HCF: i = i + 1 '74 CF
    m_data(i) = &H62: i = i + 1:    m_data(i) = &HCC: i = i + 1 '62 CC
    m_data(i) = &H91: i = i + 1:    m_data(i) = &HC8: i = i + 1 '91 C8
    m_data(i) = &H62: i = i + 1:    m_data(i) = &H8E: i = i + 1 '62 8E
    m_data(i) = &H99: i = i + 1:    m_data(i) = &H88: i = i + 1 '99 88
    m_data(i) = &H79: i = i + 1:    m_data(i) = &H14: i = i + 1 '79 14
    m_data(i) = &H71: i = i + 1:    m_data(i) = &H16: i = i + 1 '71 16
    m_data(i) = &HFE: i = i + 1:    m_data(i) = &HC:  i = i + 1 'FE 0C
    m_data(i) = &H67: i = i + 1:    m_data(i) = &HF5: i = i + 1 '67 F5
    m_data(i) = &HA1: i = i + 1:    m_data(i) = &HFE: i = i + 1 'A1 FE
    m_data(i) = &H6C: i = i + 1:    m_data(i) = &HFF: i = i + 1 '6C FF
    
End Sub

Sub SetMask_BGRA1555()
     BlueMask = &H1F&    ' 5 bit
    GreenMask = &H3E0&   ' 5 bit
      RedMask = &H7C00&  ' 5 bit
    AlphaMask = &H8000&  ' 1 bit
'    Debug.Print "BlueMask  : &H" & Hex(BlueMask) & " = " & BlueMask    'BlueMask  : &H1F = 31
'    Debug.Print "GreenMask : &H" & Hex(GreenMask) & " = " & GreenMask  'GreenMask : &H3E0 = 992
'    Debug.Print "RedMask   : &H" & Hex(RedMask) & " = " & RedMask      'RedMask   : &H7C00 = 31744
'    Debug.Print "AlphaMask : &H" & Hex(AlphaMask) & " = " & AlphaMask  'AlphaMask : &H8000 = 32768
End Sub

Sub SetMask_BGR565()
     BlueMask = &H1F&    ' 5 bit
    GreenMask = &H7E0&   ' 6 bit
      RedMask = &HF800&  ' 5 bit
    AlphaMask = &H0&     ' ----
'    Debug.Print "BlueMask  : &H" & Hex(BlueMask) & " = " & BlueMask    'BlueMask  : &H1F = 31
'    Debug.Print "GreenMask : &H" & Hex(GreenMask) & " = " & GreenMask  'GreenMask : &H7E0 = 2016
'    Debug.Print "RedMask   : &H" & Hex(RedMask) & " = " & RedMask      'RedMask   : &HF800 = 63488
'    Debug.Print "AlphaMask : &H" & Hex(AlphaMask) & " = " & AlphaMask  'AlphaMask : &H0 = 0
End Sub

Private Sub Command1_Click()
    SetColors_BGRA1555
    SetMask_BGRA1555
    Dim tb As TByt4, ti As TLng
    Dim i As Long, j As Long
    For i = 0 To 31
        tb.Value0 = m_data(j): j = j + 1
        tb.Value1 = m_data(j): j = j + 1
        LSet ti = tb
        Shape1(i).BackColor = Color16bppBGRA1555_ToColor32(ti.Value)
    Next
End Sub

Private Sub Command2_Click()
    SetColors_BGR565
    SetMask_BGR565
    Dim tb As TByt4, ti As TLng
    Dim i As Long, j As Long
    For i = 0 To 31
        tb.Value0 = m_data(j): j = j + 1
        tb.Value1 = m_data(j): j = j + 1
        LSet ti = tb
        Shape1(i).BackColor = Color16bppBGR565_ToColor32(ti.Value)
    Next
End Sub

Private Sub Command3_Click()
    Dim i As Long
    For i = 0 To 31
        Shape1(i).BackColor = &HFFFFFF
    Next
End Sub

Private Sub Command4_Click()
    Dim i As Long, h As Long
    i = 1 + 2 + 4 + 8 + 16 + 32
    h = 2 ^ 0 + 2 ^ 1 + 2 ^ 2 + 2 ^ 3 + 2 ^ 4 + 2 ^ 5
    Debug.Print i & " = &H" & Hex(i)
    Debug.Print h & " = &H" & Hex(h)
    
    
    i = 1 + 2 + 4 + 8 + 16 + 32 + 64 + 128 + 256 + 512 + 1024
    h = 2 ^ 0 + 2 ^ 1 + 2 ^ 2 + 2 ^ 3 + 2 ^ 4 + 2 ^ 5 + 2 ^ 6 + 2 ^ 7 + 2 ^ 8 + 2 ^ 9 + 2 ^ 10
    Debug.Print i & " = &H" & Hex(i)
    Debug.Print h & " = &H" & Hex(h)
    
    Debug.Print 2 ^ 11
End Sub

