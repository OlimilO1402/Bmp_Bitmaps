VERSION 5.00
Begin VB.Form FPalette 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   10320
   ClientTop       =   4785
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   4320
      Width           =   1695
   End
   Begin VB.PictureBox PanelPalette 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Shape ShPalette 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Undurchsichtig
         BorderColor     =   &H8000000D&
         BorderWidth     =   2
         Height          =   255
         Index           =   0
         Left            =   0
         Shape           =   1  'Quadrat
         Top             =   0
         Width           =   255
      End
   End
End
Attribute VB_Name = "FPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Result    As VbMsgBoxResult
Private m_Bmp       As Bitmap
Private m_Palette() As Long
Private m_Index     As Long

Private Sub Form_Load()
    '
End Sub

Public Function ShowDialog(Owner As Form, Bmp As Bitmap) As VbMsgBoxResult
    Set m_Bmp = Bmp
    If Not m_Bmp.IsIndexed Then Exit Function
    SavePalette m_Bmp
    LoadSHPalette m_Bmp.PaletteCount
    Me.Show vbModal, Owner
    ShowDialog = m_Result
End Function

Private Sub BtnOK_Click()
    m_Result = vbOK
    'Take all the changes and write it to the bitmap-palette
    Dim i As Long
    For i = 0 To UBound(m_Palette)
        m_Bmp.PaletteColor(i) = m_Palette(i)
    Next
    Unload Me
End Sub
Private Sub BtnCancel_Click()
    m_Result = vbCancel
    Unload Me
End Sub

Sub SavePalette(Bmp As Bitmap)
    Dim u As Long: u = Bmp.PaletteCount - 1
    ReDim m_Palette(0 To u)
    Dim i As Long
    For i = 0 To u
        m_Palette(i) = Bmp.PaletteColor(i)
    Next
End Sub

Sub LoadSHPalette(ByVal n As Long)
    Dim L0 As Single: L0 = ShPalette(0).Left
    Dim T0 As Single: T0 = ShPalette(0).Top
    Dim W0 As Single: W0 = ShPalette(0).Width
    Dim H0 As Single: H0 = ShPalette(0).Height
    Dim L As Single: L = L0
    Dim T As Single: T = T0
    Dim i As Long
    For i = 0 To n - 1
        If i > 0 Then Load ShPalette(i)
        ShPalette(i).Move L, T, W0, H0
        ShPalette(i).Visible = True
        ShPalette(i).BorderStyle = 0
        ShPalette(i).BackColor = m_Palette(i)
        ShPalette(i).BorderWidth = 3
        L = L + W0
        If ((i + 1) Mod 16) = 0 Then
            L = L0
            T = T + H0
        End If
    Next
    Dim PH As Single: PH = IIf(n < 255, 1, 16) * H0
    PanelPalette.Height = PH
    BtnOK.Top = PanelPalette.Top + PH + 8
    BtnCancel.Top = BtnOK.Top
    Dim borders As Single: borders = Me.Height - (Me.ScaleHeight * Screen.TwipsPerPixelY)
    Dim SH As Single: SH = BtnOK.Top + BtnOK.Height + 8
    Me.Height = borders + (SH * Screen.TwipsPerPixelY) '5385
End Sub

Private Sub PanelPalette_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_Index = GetShapeIndex(X, Y)
    SetBorderStyleTransparent
    If m_Index < 0 Then Exit Sub
    ShPalette(m_Index).BorderStyle = 1
    ShPalette(m_Index).BorderColor = &H8000000D
End Sub

Private Sub PanelPalette_DblClick()
    If m_Index < 0 Then Exit Sub
    Dim CD As ColorDialog: Set CD = New ColorDialog
    CD.Color = m_Palette(m_Index)
    If CD.ShowDialog(Me) = vbCancel Then Exit Sub
    m_Palette(m_Index) = CD.Color
    ShPalette(m_Index).BackColor = m_Palette(m_Index)
End Sub

Private Sub SetBorderStyleTransparent()
    Dim i As Long
    For i = 0 To ShPalette.UBound
        If ShPalette(i).BorderStyle = 1 Then
            ShPalette(i).BorderStyle = 0
        End If
    Next
End Sub

Private Function GetShapeIndex(ByVal X As Long, ByVal Y As Long) As Long
    Dim i As Long: GetShapeIndex = -1
    For i = 0 To ShPalette.UBound
        If (ShPalette(i).Left < X) And (X < ShPalette(i).Left + ShPalette(i).Width) And _
           (ShPalette(i).Top < Y) And (Y < ShPalette(i).Top + ShPalette(i).Height) Then
           GetShapeIndex = i
           Exit Function
        End If
    Next
End Function

Private Sub SetDefaultColorPalette()
    ShPalette(0).BackColor = &HFFFFFF
    ShPalette(1).BackColor = &HC0C0FF
    ShPalette(2).BackColor = &HC0E0FF
    ShPalette(3).BackColor = &HC0FFFF
    ShPalette(4).BackColor = &HC0FFC0
    ShPalette(5).BackColor = &HFFFFC0
    ShPalette(6).BackColor = &HFFC0C0
    ShPalette(7).BackColor = &HFFC0FF
    ShPalette(16).BackColor = &HE0E0E0
    ShPalette(17).BackColor = &H8080FF
    ShPalette(18).BackColor = &H80C0FF
    ShPalette(19).BackColor = &H80FFFF
    ShPalette(20).BackColor = &H80FF80
    ShPalette(21).BackColor = &HFFFF80
    ShPalette(22).BackColor = &HFF8080
    ShPalette(23).BackColor = &HFF80FF
    ShPalette(32).BackColor = &HC0C0C0
    ShPalette(33).BackColor = &HFF&
    ShPalette(34).BackColor = &H80FF&
    ShPalette(35).BackColor = &HFFFF&
    ShPalette(36).BackColor = &HFF00&
    ShPalette(37).BackColor = &HFFFF00
    ShPalette(38).BackColor = &HFF0000
    ShPalette(39).BackColor = &HFF00FF
    ShPalette(48).BackColor = &H808080
    ShPalette(49).BackColor = &HC0&
    ShPalette(50).BackColor = &H40C0&
    ShPalette(51).BackColor = &HC0C0&
    ShPalette(52).BackColor = &HC000&
    ShPalette(53).BackColor = &HC0C000
    ShPalette(54).BackColor = &HC00000
    ShPalette(55).BackColor = &HC000C0
    ShPalette(64).BackColor = &H404040
    ShPalette(65).BackColor = &H80&
    ShPalette(66).BackColor = &H4080&
    ShPalette(67).BackColor = &H8080&
    ShPalette(68).BackColor = &H8000&
    ShPalette(69).BackColor = &H808000
    ShPalette(70).BackColor = &H800000
    ShPalette(71).BackColor = &H800080
    ShPalette(80).BackColor = &H0&
    ShPalette(81).BackColor = &H40&
    ShPalette(82).BackColor = &H404080
    ShPalette(83).BackColor = &H4040&
    ShPalette(84).BackColor = &H4000&
    ShPalette(85).BackColor = &H404000
    ShPalette(86).BackColor = &H400000
    ShPalette(87).BackColor = &H400040
End Sub

