VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "www.activevb.de"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows-Standard
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   1
      LargeChange     =   10
      Left            =   2160
      Max             =   167
      Min             =   20
      SmallChange     =   2
      TabIndex        =   4
      Top             =   2160
      Value           =   20
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Index           =   0
      LargeChange     =   10
      Left            =   360
      Max             =   350
      Min             =   20
      SmallChange     =   2
      TabIndex        =   3
      Top             =   2160
      Value           =   20
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3600
      ScaleHeight     =   420
      ScaleWidth      =   375
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1620
      Left            =   120
      ScaleHeight     =   1560
      ScaleWidth      =   3270
      TabIndex        =   0
      Top             =   360
      Width           =   3330
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   3585
      Top             =   105
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Original"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3330
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

Private Sub Form_Load()
    Picture1.Picture = LoadPicture(App.Path & "\Bild.jpg")
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
End Sub

Private Sub StretchPicture(ByVal fX As Long, ByVal fY As Long)
    
    Dim ax As Double: ax = fX / Picture1.ScaleWidth
    Dim ay As Double: ay = fY / Picture1.ScaleHeight
    
    Dim a As Double
    If Picture1.ScaleWidth * ay > fX Then
        a = ax
    Else
        a = ay
    End If
    
    Dim x As Long: x = Picture1.ScaleWidth * a
    Dim y As Long: y = Picture1.ScaleHeight * a
    
    Picture2.AutoRedraw = True
    Picture2.Width = x * Screen.TwipsPerPixelX
    Picture2.Height = y * Screen.TwipsPerPixelY
    Picture2.Refresh
    
    Picture2.PaintPicture Picture1.Picture, 0, 0, x, y
    Picture2.AutoRedraw = False
End Sub

Private Sub HScroll1_Change(Index As Integer)
    Dim x As Long: x = HScroll1(0).Value
    Dim y As Long: y = HScroll1(1).Value
    StretchPicture x, y
    Label4.Caption = "Anpassung auf " & x & " x " & y
    Shape1.Width = (x + 2) * Screen.TwipsPerPixelX
    Shape1.Height = (y + 2) * Screen.TwipsPerPixelY
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
  Call HScroll1_Change(Index)
End Sub
