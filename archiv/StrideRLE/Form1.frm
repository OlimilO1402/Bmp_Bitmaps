VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command6 
      Caption         =   "Test6"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Fill Test5"
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Fill Test4"
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fill Test3"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fill Test2"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill Test1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   840
      Width           =   8415
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   8415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   8415
   End
   Begin VB.CommandButton BtnDecode 
      Caption         =   "Decode"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Res"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "RGB"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "RLE"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private W As Long
Private H As Long
Private m_PixelDataRLE() As Byte
Private m_PixelDataRGB() As Byte
Private m_DataResult() As Byte

Private Declare Function RtlCompareMemory Lib "ntdll" (pSrc1 As Any, pSrc2 As Any, ByVal bytLength As Long) As Long

Sub UpdateView()
    Text1.Text = Bytes_ToHex(m_PixelDataRLE)
    Text2.Text = Bytes_ToHex(m_PixelDataRGB)
    Text3.Text = vbNullString
End Sub

Sub FillExample1()
    W = 4: H = 3    'yes the rle-encoded data is actually longer
    FillBytes m_PixelDataRLE, "00 04 D4 67 C8 67 00 00 04 C8 00 00 04 67 00 01" '16 Bytes
    FillBytes m_PixelDataRGB, "D4 67 C8 67 C8 C8 C8 C8 67 67 67 67"             '12Bytes
    UpdateView
End Sub
Sub FillExample2()
    W = 5: H = 3
    FillBytes m_PixelDataRLE, "00 05 D4 67 C8 67 D4 00 00 00 05 C8 00 00 05 67 00 01"                   '18 Bytes
    FillBytes m_PixelDataRGB, "D4 67 C8 67 D4 00 00 00 C8 C8 C8 C8 C8 00 00 00 67 67 67 67 67 00 00 00" '24 Bytes
    UpdateView
End Sub
Sub FillExample3()
    W = 5: H = 3
    FillBytes m_PixelDataRLE, "00 05 D4 67 C8 67 D4 00 00 00 01 67 04 C8 00 00 05 67 00 01"             '20 Bytes
    FillBytes m_PixelDataRGB, "D4 67 C8 67 D4 00 00 00 67 C8 C8 C8 C8 00 00 00 67 67 67 67 67 00 00 00" '24 Bytes
    UpdateView
End Sub
Sub FillExample4()
    W = 11: H = 7
    FillBytes m_PixelDataRLE, "00 03 5A 5A 2B 00 04 4B 00 04 55 60 84 84 00 00 00 0B 5A 30 50 4B 4B 4B 70 50 5A 84 84 00 00 00 00 0B 5A 5A 4B 4B 4C 4B 4B 4B 5A 7E 84 00 00 00 00 0B 5A 50 4B 4B 6F 4B 70 4B 50 5A 5A 00 00 00 00 0B 5A 4B 4B 4B 4C 4B 4C 4B 70 51 55 00 00 00 00 0B 5A 4F 51 4B 6F 4C 4B 70 4B 4C 6F 00 00 00 00 0B 54 54 4F 50 4B 4C 6F 4C 4B 70 4B 00 00 01"
    FillBytes m_PixelDataRGB, "5A 5A 2B 4B 4B 4B 4B 55 60 84 84 00 5A 30 50 4B 4B 4B 70 50 5A 84 84 00 5A 5A 4B 4B 4C 4B 4B 4B 5A 7E 84 00 5A 50 4B 4B 6F 4B 70 4B 50 5A 5A 00 5A 4B 4B 4B 4C 4B 4C 4B 70 51 55 00 5A 4F 51 4B 6F 4C 4B 70 4B 4C 6F 00 54 54 4F 50 4B 4C 6F 4C 4B 70 4B 00"
    UpdateView
End Sub
Sub FillExample5()
    W = 16: H = 9
    FillBytes m_PixelDataRLE, "09 51 00 07 4D 28 2C 2D 40 51 51 00 00 00 00 10 51 0A 0C 0D 0F 11 13 26 4C 51 46 1D 31 1C 51 51 00 00 00 10 51 0B 38 38 38 29 25 4B 51 51 4A 1A 35 1C 51 51 00 00 00 10 51 0C 39 39 2B 12 4B 51 51 51 42 23 30 3E 51 51 00 00 00 10 51 0C 3A 2E 35 14 3F 49 49 44 1A 34 22 43 51 51 00 00 00 10 51 0D 2A 10 16 36 21 17 18 1F 32 2F 3C 4F 51 51 00 00 00 10 51 0E 20 47 3D 15 30 37 37 33 24 3B 48 51 51 51 00 00 00 0C 51 1E 4B 51 50 45 27 18 19 1B 41 4E 04 51 00 00 01 51 01 4B 0E 51 00 01"
    FillBytes m_PixelDataRGB, "51 51 51 51 51 51 51 51 51 4D 28 2C 2D 40 51 51 51 0A 0C 0D 0F 11 13 26 4C 51 46 1D 31 1C 51 51 51 0B 38 38 38 29 25 4B 51 51 4A 1A 35 1C 51 51 51 0C 39 39 2B 12 4B 51 51 51 42 23 30 3E 51 51 51 0C 3A 2E 35 14 3F 49 49 44 1A 34 22 43 51 51 51 0D 2A 10 16 36 21 17 18 1F 32 2F 3C 4F 51 51 51 0E 20 47 3D 15 30 37 37 33 24 3B 48 51 51 51 51 1E 4B 51 50 45 27 18 19 1B 41 4E 51 51 51 51 51 4B 51 51 51 51 51 51 51 51 51 51 51 51 51 51"
    UpdateView
End Sub
Sub FillExample6()
    'Read BitmapRGB.data
    'Read BitampRLE.data
End Sub

Private Sub Command1_Click()
    FillExample1
End Sub
Private Sub Command2_Click()
    FillExample2
End Sub
Private Sub Command3_Click()
    FillExample3
End Sub
Private Sub Command4_Click()
    FillExample4
End Sub
Private Sub Command5_Click()
    FillExample5
End Sub
Private Sub Command6_Click()
    FillExample6
End Sub

Public Sub FillBytes(Dst_buffer() As Byte, ByVal s As String)
    Dim sa() As String: sa = Split(s, " ")
    ReDim Dst_buffer(0 To UBound(sa))
    Dim i As Long
    For i = 0 To UBound(sa)
        Dst_buffer(i) = CByte("&H" & sa(i))
    Next
End Sub

Function Bytes_ToHex(buffer() As Byte) As String
    ReDim sa(0 To UBound(buffer)) As String
    Dim i As Long
    For i = 0 To UBound(sa)
        sa(i) = Hex2(buffer(i))
    Next
    Bytes_ToHex = Join(sa, " ")
End Function
Function Hex2(b As Byte) As String
    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Private Sub BtnDecode_Click()
    MRLE.RLE8_Decode W, H, m_PixelDataRLE, m_DataResult
    Dim l As Long: l = UBound(m_DataResult) + 1
    Dim equals As Boolean: equals = (RtlCompareMemory(m_PixelDataRGB(0), m_DataResult(0), l) = l)
    Text3.Text = Bytes_ToHex(m_DataResult)
    If equals Then MsgBox "OK gleich"
End Sub
