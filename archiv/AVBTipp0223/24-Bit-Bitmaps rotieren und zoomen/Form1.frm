VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows-Standard
   Begin VB.VScrollBar VScroll1 
      Height          =   4545
      Left            =   6120
      Max             =   200
      TabIndex        =   4
      Top             =   600
      Value           =   100
      Width           =   270
   End
   Begin VB.PictureBox ps 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3960
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   260
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.PictureBox pd 
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   60
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   720
      Width           =   6060
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   500
      Left            =   60
      SmallChange     =   50
      TabIndex        =   0
      Top             =   5280
      Width           =   6045
   End
   Begin VB.PictureBox po 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   4560
      Left            =   240
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   6060
   End
   Begin VB.Label Label1 
      Caption         =   "Achtung: Um die Effekte in akzeptabler Geschwindigkeit genießen zu können, diesen Source vorab als exe kompilieren"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6015
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

'Achtung: Um die Effekte in akzeptabler Geschwindigkeit ge-
'         nießen zu können, diesen Source vorab als exe kom-
'         pilieren!

Option Explicit

Private Declare Function GetObject Lib "GDI32" Alias _
         "GetObjectA" (ByVal hObject As Long, ByVal _
         nCount As Long, lpObject As Any) As Long

Private Declare Function VarPtrArray Lib "msvbvm50.dll" _
        Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias _
        "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal _
        ByteLen As Long)

Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC _
        As Long, ByVal x As Long, ByVal y As Long, ByVal _
        nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC _
        As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
        ByVal dwRop As Long) As Long

Private Type SAFEARRAYBOUND
  cElements As Long
  lLbound As Long
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
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const SRCAND = &H8800C6

Private b_hgt%, b_wid%, inhere%, cx%, cy%
Private pic_w%, pic_h%

Sub RotateZoom(DestP As PictureBox, x%, y%, Angle As Double, _
               Zoom As Double, SourceP As PictureBox, sx1%, _
               sy1%, sx2%, sy2%)
          
  Dim PictD() As Byte, PictS() As Byte, ox%, oy%, r%, c%
  Dim cd%, rd%, cs%, rs%, csc%, rsc%, mx%, my%, Wid%, Hgt%
  Dim sD As SAFEARRAY2D, BmpD As BITMAP
  Dim sS As SAFEARRAY2D, BmpS As BITMAP
  Dim ASin As Double, ACos As Double
  Dim TransR As Byte, transG As Byte, transB As Byte
  Dim srcR As Byte, srcG As Byte, srcB As Byte
 
 
    Call GetObject(DestP.Picture, Len(BmpD), BmpD)
    Call GetObject(SourceP.Picture, Len(BmpS), BmpS)
    
    'Es werden nur 24-Bit Bitmaps unterstützt
    If BmpS.bmBitsPixel <> 24 Then
      MsgBox "Es werden nur 24-Bit Bitmaps unterstützt"
      Exit Sub
    End If
    
    With sD
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = BmpD.bmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BmpD.bmWidthBytes
      .pvData = BmpD.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(PictD), VarPtr(sD), 4)
    
    With sS
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = BmpS.bmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BmpS.bmWidthBytes
      .pvData = BmpS.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(PictS), VarPtr(sS), 4)
    

    'max. breite und höhe setzen
    If sx2 - sx1 > sy2 - sy1 Then
      Wid = sx2 - sx1
      Hgt = sx2 - sx1
    Else
      Wid = sy2 - sy1
      Hgt = sy2 - sy1
    End If
    
    'max offsets ermitteln
    ox = sx1 + Wid / 2
    oy = sy1 + Hgt / 2
    
    'Mittelpunkt ermitteln
    mx = (sx1 + sx2) / 2
    my = (sy1 + sy2) / 2
    
    'Rotate & Zoom
    ASin = Sin(Angle) * Zoom
    ACos = Cos(Angle) * Zoom
    
    'Transparente Farbe ermitteln (0,0), Werte für RGB
    TransR = PictS(sx1, sy1)
    transG = PictS(sx1 + 1, sy1)
    transB = PictS(sx1 + 2, sy1)
    
    'Hauptschleife
    cs = sx1
    For cd = x To x + Wid
      cs = cs + 1
      rs = sy1
      For rd = y To y + Hgt
        
        'Transformieren der source Koordinaten
         csc = mx + (cs - ox) * ASin + (rs - oy) * ACos
         rsc = my + (rs - oy) * ASin - (cs - ox) * ACos
         
         'Überprüfen ob es im Bereich liegt
         If (csc >= sx1 And csc <= sx2) Then
           If (rsc >= sy1 And rsc <= sy2) Then
             
             'Pixelwerte aus source-bitmap erfassen
             srcR = PictS(csc * 3, rsc)
             srcG = PictS(csc * 3 + 1, rsc)
             srcB = PictS(csc * 3 + 2, rsc)
             If srcR <> TransR And srcG <> transG And _
                srcB <> transG Then
               'nicht transparent also COPY!!
               PictD(cd * 3, rd) = srcR
               PictD(cd * 3 + 1, rd) = srcG
               PictD(cd * 3 + 2, rd) = srcB
              End If
            End If
          End If
        rs = rs + 1
      Next rd
    Next cd
    
    Call CopyMemory(ByVal VarPtrArray(PictD), 0&, 4)
    Call CopyMemory(ByVal VarPtrArray(PictS), 0&, 4)
End Sub

Private Sub Form_Load()
  po.Picture = LoadPicture(App.Path & "\back.jpg")
  pd.Picture = LoadPicture(App.Path & "\back.jpg")
  
  'Mittelwert als Voreinstellung
  HScroll1.Value = 15707
  HScroll1_Scroll
End Sub

Private Sub HScroll1_Change()
  HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
  'Hintergrund löschen
  Call CopyPicture(pd, po)
  
  'Grafik rotieren und zoomen
  Call RotateZoom(pd, 70, 20, (CDbl(HScroll1.Value) / 5000), _
                  (CDbl(VScroll1.Value) / 100), ps, 0, 0, _
                  ps.ScaleWidth - 10, ps.ScaleHeight - 10)
  'Anzeigen
  pd.Refresh
End Sub

Sub CopyPicture(DestP As PictureBox, DestO As PictureBox)
  Dim PictD() As Byte, PictO() As Byte
  Dim sD As SAFEARRAY2D, BmpD As BITMAP
  Dim sO As SAFEARRAY2D, bmpO As BITMAP
  Dim r%, c%
  
    'Bitmap info
    Call GetObject(DestP.Picture, Len(BmpD), BmpD)
    Call GetObject(DestO.Picture, Len(bmpO), bmpO)
    
    'Es werden nur 24-Bit Bitmaps unterstützt
    If bmpO.bmBitsPixel <> 24 Then
      MsgBox "Es werden nur 24-Bit Bitmaps unterstützt"
      Exit Sub
    End If
    
    With sD
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = BmpD.bmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = BmpD.bmWidthBytes
      .pvData = BmpD.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(PictD), VarPtr(sD), 4)
    
    With sO
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = bmpO.bmHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = bmpO.bmWidthBytes
      .pvData = bmpO.bmBits
    End With
    
    Call CopyMemory(ByVal VarPtrArray(PictO), VarPtr(sO), 4)
    
    'Einfaches kopieren der Pixel (Spiegelverkehrt)
    For c = 0 To UBound(PictO, 1) - 1
      For r = 0 To UBound(PictO, 2) - 1
        PictD(c, r) = PictO(c, r)
      Next r
    Next c
    
    'Freigeben
    Call CopyMemory(ByVal VarPtrArray(PictD), 0&, 4)
    Call CopyMemory(ByVal VarPtrArray(PictO), 0&, 4)
End Sub

Private Sub VScroll1_Change()
  HScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
  HScroll1_Scroll
End Sub
