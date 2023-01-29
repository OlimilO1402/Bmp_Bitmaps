VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Blur"
      Height          =   435
      Left            =   6660
      TabIndex        =   6
      Top             =   4800
      Width           =   1515
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   4215
      Left            =   3300
      ScaleHeight     =   4155
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   420
      Width           =   3075
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   420
      Width           =   3075
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":1C91
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   4800
      Width           =   6195
   End
   Begin VB.Image Image2 
      Height          =   1935
      Left            =   6660
      Stretch         =   -1  'True
      Top             =   2700
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   6660
      Stretch         =   -1  'True
      Top             =   420
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Von der Kopie"
      Height          =   195
      Index           =   3
      Left            =   6660
      TabIndex        =   5
      Top             =   2460
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Vom Orginal"
      Height          =   195
      Index           =   2
      Left            =   6660
      TabIndex        =   4
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Kopie"
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   3
      Top             =   180
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "Orginal"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   2595
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

Private Declare Function VarPtrArray Lib "msvbvm50.dll" _
        Alias "VarPtr" (ptr() As Any) As Long
        
Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal _
        ByteLen As Long)
        
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC _
        As Long, ByVal x As Long, ByVal y As Long, ByVal _
        nWidth As Long, ByVal nHeight As Long, ByVal _
        hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc _
        As Long, ByVal dwRop As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias _
        "GetObjectA" (ByVal hObject As Long, ByVal nCount _
        As Long, lpObject As Any) As Long

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
End Type

Private Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 0) As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As Long
  Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Private Const SRCCOPY = &HCC0020

Dim aa As Long, bb As Long

Private Sub DoBlur(bPic As PictureBox)
   Dim Pict() As Byte
   Dim av As Long
   Dim ptr As Long
   Dim safe As SAFEARRAY1D, bmp As BITMAP
   
    Call GetObject(bPic.Picture, Len(bmp), bmp)
    With safe
      .cbElements = 1
      .cDims = 1
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = bmp.bmHeight * bmp.bmWidthBytes
      .pvData = bmp.bmBits
    End With
    Call CopyMemory(ByVal VarPtrArray(Pict), VarPtr(safe), 4)
    
    On Error Resume Next

    'Blur algo
    ptr = bmp.bmWidthBytes + 3
    For aa = 1 To bmp.bmHeight - 3
      For bb = 0 To bmp.bmWidthBytes
        ptr = ptr + 1
        av = Pict(ptr - bmp.bmWidthBytes)
        av = av + Pict(ptr - 3)
        av = av + Pict(ptr + 3)
        av = av + Pict(ptr + bmp.bmWidthBytes)
        Pict(ptr) = av \ 4
      Next bb
    Next aa

    Call CopyMemory(ByVal VarPtrArray(Pict), 0&, 4)
End Sub

Private Sub Command2_Click()
  Call SavePicture(Picture2.Picture, App.Path & "\Temp.BMP")
  Picture2.Picture = LoadPicture(App.Path & "\Temp.BMP")
  
  Call DoBlur(Picture2)
  Image1.Picture = Picture1.Image
  Image2.Picture = Picture2.Image
End Sub

Private Sub Form_Load()
  Dim oFS As Integer
  
    oFS = 11
    Picture1.FontSize = oFS
    Picture1.CurrentY = 400
    Picture1.CurrentX = 100
    Picture1.Print "Dieses Beispiel stammt von"
    Picture1.CurrentY = 2400
    Picture1.CurrentX = 100
    Picture1.FontSize = 18
    Picture1.Print "Blur und Preview"
    Picture1.CurrentY = 2800
    Picture1.CurrentX = 100
    Picture1.Print "Demo in VB."
    
    Picture1.FontSize = oFS
    Picture1.CurrentY = 3400
    Picture1.CurrentX = 100
    Picture1.Print "Viel Spaß beim testen,"
    Picture1.CurrentY = 3700
    Picture1.CurrentX = 100
    Picture1.Print "Dirk Lietzow"
    
    Picture2.Picture = Picture1.Image
    Image1.Picture = Picture1.Image
    Image2.Picture = Picture2.Image
End Sub

Private Sub Picture1_Click()
  Picture2.Picture = Picture1.Image
  Image1.Picture = Picture1.Image
  Image2.Picture = Picture2.Image
End Sub

Private Sub Picture2_Click()
  Picture2.Picture = Picture1.Image
  Image1.Picture = Picture1.Image
  Image2.Picture = Picture2.Image
End Sub
