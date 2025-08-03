VERSION 5.00
Begin VB.Form FDlgNewPicture 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "New Picture"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox TxtWidth 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox CmbPixel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "FDlgNewPicture.frx":0000
      Left            =   3360
      List            =   "FDlgNewPicture.frx":000D
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox TxtHeight 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox TxtResolution 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
      Width           =   1815
   End
   Begin VB.ComboBox CmbResolution 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "FDlgNewPicture.frx":002A
      Left            =   3360
      List            =   "FDlgNewPicture.frx":0034
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.PictureBox PBBackColor 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox CmbPixelFormat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "FDlgNewPicture.frx":005C
      Left            =   1440
      List            =   "FDlgNewPicture.frx":0066
      TabIndex        =   2
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Picture dimensions and properties:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   240
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Resolution:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Background Color:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Pixelformat:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   1035
   End
   Begin VB.Label LblMem 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Memory in Use: xx.xxx MByte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1365
      TabIndex        =   9
      Top             =   3600
      Width           =   2505
   End
End
Attribute VB_Name = "FDlgNewPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_DlgResult  As VbMsgBoxResult
Private m_Width      As Double
Private m_Height     As Double
Private m_DimUnit    As Long
Private m_Resolution As Double
Private m_ResUnit    As Long
Private m_PixFmt     As EPixelFormat
Private m_bInit      As Boolean
Private m_bUpVw      As Boolean
Private m_SelColor   As Long

Private Sub BtnOK_Click()
    m_DlgResult = vbOK
    Unload Me
End Sub

Private Sub BtnCancel_Click()
    m_DlgResult = vbCancel
    Unload Me
End Sub

Private Sub Form_Load()
    Pixelformat_ToListBox CmbPixelFormat
    'Init
End Sub

Private Sub Init(Optional bmp As Bitmap = Nothing)
    m_bInit = True
    If bmp Is Nothing Then
        m_Width = Screen.Width / Screen.TwipsPerPixelX / 5
        m_Height = m_Width * 9 / 16
        m_DimUnit = 0 '0=Pixel, 1=Inch, 2=Centimeter
        m_ResUnit = 0 'dpi
        m_Resolution = 96
        m_PixFmt = EPixelFormat.Format32bppArgb
    Else
        m_Width = bmp.Width
        m_Height = bmp.Height
        m_DimUnit = 0 '0=Pixel, 1=Inch, 2=Centimeter
        m_ResUnit = 0 'dpi
        m_Resolution = CLng(bmp.PixelPerMeterX * 0.025401)
        m_PixFmt = bmp.PixelFormat
    End If
    m_bInit = False
    UpdateView
End Sub

Public Function ShowDialog(Owner As FMain, bmp_inout As Bitmap) As VbMsgBoxResult
    Init bmp_inout
    Me.Show vbModal, Owner
    ShowDialog = m_DlgResult
    If m_DlgResult <> vbOK Then Exit Function
    If bmp_inout Is Nothing Then
        Set bmp_inout = MNew.BitmapWH(CLng(m_Width), CLng(m_Height), m_PixFmt)
    Else
        If bmp_inout.Width <> m_Width Or bmp_inout.Height <> m_Width Then
            bmp_inout.NewWH m_Width, m_Height, m_PixFmt
        End If
    End If
    Dim ppm As Long: ppm = CLng(PixelPerMeter)
    If bmp_inout.PixelPerMeterX <> ppm Then
        bmp_inout.PixelPerMeterX = ppm
    End If
    If bmp_inout.PixelPerMeterY <> ppm Then
        bmp_inout.PixelPerMeterY = ppm
    End If
    If m_SelColor <> 0 Then
        bmp_inout.Fill m_SelColor 'PBBackColor.BackColor
    End If
End Function

Private Property Get PixelPerMeter() As Double
    PixelPerMeter = m_Resolution / 0.025401
End Property

Private Property Get BytesPerPixel() As Single
    'for 1 bit this is 1/8
    Select Case m_PixFmt
    Case EPixelFormat.Format1bppIndexed:    BytesPerPixel = 1 / 8
    Case EPixelFormat.Format4bppIndexed:    BytesPerPixel = 4 / 8
    Case EPixelFormat.Format8bppIndexed:    BytesPerPixel = 8 / 8
    Case EPixelFormat.Format16bppArgb1555, EPixelFormat.Format16bppGrayScale, EPixelFormat.Format16bppRgb555, EPixelFormat.Format16bppRgb565
                                            BytesPerPixel = 16 / 8
    Case EPixelFormat.Format24bppRgb:       BytesPerPixel = 24 / 8
    Case EPixelFormat.Format32bppRgb, EPixelFormat.Format32bppArgb, EPixelFormat.Format32bppPArgb
                                            BytesPerPixel = 32 / 8
    Case EPixelFormat.Format48bppRgb:       BytesPerPixel = 48 / 8
    Case EPixelFormat.Format64bppArgb, EPixelFormat.Format64bppPArgb
                                            BytesPerPixel = 64 / 8
    End Select
End Property
Private Sub UpdateView()
    If m_bInit Then Exit Sub
    m_bUpVw = True
    Dim W As Double: W = m_Width
    Dim H As Double: H = m_Height
    Select Case m_DimUnit
    Case 0: 'Pixel ' OK nothing todo here
    Case 1: 'Inches
            W = W / m_Resolution: H = H / m_Resolution
    Case 2: Dim ppm As Double: ppm = PixelPerMeter
            W = W / (ppm / 100):  H = H / (ppm / 100)
    End Select
    TxtWidth.Text = Format(W, "0.0000")
    TxtHeight.Text = Format(H, "0.0000")
    CmbPixel.ListIndex = m_DimUnit
    Dim R As Double: R = m_Resolution
    Select Case m_ResUnit
    Case 0: ' dpi ' OK nothing todo here
    Case 1: ' PixelPerCentimeter
            R = R / 2.54
    End Select
    TxtResolution.Text = Format(R, "0.0000")
    CmbResolution.ListIndex = m_ResUnit
    CmbPixelFormat.ListIndex = PixelFormat_ToIndex(m_PixFmt)
    'Now we calculate the amount of memory in use by the Picture
    'Memory in Use: xx.xxx MByte
    Dim m As Double: m = m_Width * m_Height * BytesPerPixel '
    Dim un As String: un = " Byte"
    If m > 1024 Then m = m / 1024: un = " KByte"
    If m > 1024 Then m = m / 1024: un = " MByte"
    LblMem.Caption = "Memory in use: " & Format(m, "0.00") & un
    m_bUpVw = False
End Sub

Private Function NewEPixelFormat(aCBLB As ComboBox) As EPixelFormat
    Dim pf As EPixelFormat
    Select Case aCBLB.ListIndex
    Case 0:  pf = EPixelFormat.Format1bppIndexed
    Case 1:  pf = EPixelFormat.Format4bppIndexed
    Case 2:  pf = EPixelFormat.Format8bppIndexed
    Case 3:  pf = EPixelFormat.Format16bppRgb555
    Case 4:  pf = EPixelFormat.Format16bppRgb565
    Case 5:  pf = EPixelFormat.Format16bppArgb1555
    Case 6:  pf = EPixelFormat.Format16bppGrayScale
    Case 7:  pf = EPixelFormat.Format24bppRgb
    Case 8:  pf = EPixelFormat.Format32bppRgb
    Case 9:  pf = EPixelFormat.Format32bppArgb
    Case 10: pf = EPixelFormat.Format32bppPArgb
    'Case 11
    End Select
    NewEPixelFormat = pf
End Function

Private Sub Pixelformat_ToListBox(aCBLB)
    aCBLB.Clear
    With aCBLB
        .AddItem Pixelformat_ToStr(EPixelFormat.Format1bppIndexed)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format4bppIndexed)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format8bppIndexed)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format16bppRgb555)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format16bppRgb565)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format16bppArgb1555)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format16bppGrayScale)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format24bppRgb)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format32bppRgb)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format32bppArgb)
        .AddItem Pixelformat_ToStr(EPixelFormat.Format32bppPArgb)
    End With
End Sub

Function Pixelformat_ToStr(this As EPixelFormat) As String
    Dim s As String
    Select Case this
    Case EPixelFormat.Format1bppIndexed:    s = "1-Bpp 2 colors+palette"
    Case EPixelFormat.Format4bppIndexed:    s = "4-Bpp 16 colors+palette"
    Case EPixelFormat.Format8bppIndexed:    s = "8-Bpp 256 colors+palette"
    Case EPixelFormat.Format16bppRgb555:    s = "15-Bpp 32768 colors"
    Case EPixelFormat.Format16bppRgb565:    s = "16-Bpp 65536 colors"
    Case EPixelFormat.Format16bppArgb1555:  s = "15-Bpp 32768 colors+transparency-bit"
    Case EPixelFormat.Format16bppGrayScale: s = "16-Bpp 65536 greyscale"
    Case EPixelFormat.Format24bppRgb:       s = "24-Bpp RGB 16.77mio.colors"
    Case EPixelFormat.Format32bppRgb:       s = "32-Bpp RGB 16.77mio.colors opaque"
    Case EPixelFormat.Format32bppArgb:      s = "32-Bpp ARGB 16.77mio.colors transp."
    Case EPixelFormat.Format32bppPArgb:     s = "32-Bpp PARGB 16.77mio.colors"
    'Case Else: s = ""
    End Select
    Pixelformat_ToStr = s
End Function

Function PixelFormat_ToIndex(this As EPixelFormat) As Long
    Dim i As Long
    Select Case this
    Case EPixelFormat.Format1bppIndexed:    i = 0  '"1-Bpp 2 colors+palette"
    Case EPixelFormat.Format4bppIndexed:    i = 1  '"4-Bpp 16 colors+palette"
    Case EPixelFormat.Format8bppIndexed:    i = 2  '"8-Bpp 256 colors+palette"
    Case EPixelFormat.Format16bppRgb555:    i = 3  '"15-Bpp 32768 colors"
    Case EPixelFormat.Format16bppRgb565:    i = 4  '"16-Bpp 65536 colors"
    Case EPixelFormat.Format16bppArgb1555:  i = 5  '"15-Bpp 32768 colors+transparency-bit"
    Case EPixelFormat.Format16bppGrayScale: i = 6  '"16-Bpp 65536 greyscale"
    Case EPixelFormat.Format24bppRgb:       i = 7  '"24-Bpp RGB 16.77mio.colors"
    Case EPixelFormat.Format32bppRgb:       i = 8  '"32-Bpp RGB 16.77mio.colors opaque"
    Case EPixelFormat.Format32bppArgb:      i = 9  '"32-Bpp ARGB 16.77mio.colors transp."
    Case EPixelFormat.Format32bppPArgb:     i = 10 '"32-Bpp PARGB 16.77mio.colors"
    'Case Else: s = ""
    End Select
    PixelFormat_ToIndex = i
End Function

Private Function Double_TryParse(ByVal s As String, ByRef d_out As Double) As Boolean
Try: On Error GoTo Catch
    s = Replace(s, ",", ".")
    If IsNumeric(s) Then
        d_out = Val(s)
        Double_TryParse = True
    End If
Catch:
End Function

Private Sub PBBackColor_Click()
    Dim ColorDlg As New ColorDialog
    ColorDlg.Color = PBBackColor.BackColor
    If ColorDlg.ShowDialog(Me) = vbCancel Then Exit Sub
    m_SelColor = ColorDlg.Color
    PBBackColor.BackColor = m_SelColor
End Sub

Private Sub TxtWidth_LostFocus()
    Dim s As String: s = TxtWidth.Text
    Dim W As Double
    If Double_TryParse(s, W) Then
        Select Case m_DimUnit
        Case 0: m_Width = W
        Case 1: m_Width = W * m_Resolution
        Case 2: m_Width = W * PixelPerMeter / 100
        End Select
        If Not m_bUpVw Then UpdateView
    Else
        MsgBox "Please give a valid numerical value: " & vbCrLf & s
    End If
End Sub

Private Sub TxtHeight_LostFocus()
    Dim s As String: s = TxtHeight.Text
    Dim H As Double
    If Double_TryParse(s, H) Then
        Select Case m_DimUnit
        Case 0: m_Height = H
        Case 1: m_Height = H * m_Resolution
        Case 2: m_Height = H * PixelPerMeter / 100
        End Select
        If Not m_bUpVw Then UpdateView
    Else
        MsgBox "Please give a valid numerical value: " & vbCrLf & s
    End If
End Sub

Private Sub TxtResolution_LostFocus()
    Dim s As String: s = TxtResolution.Text
    Dim R As Double
    If Double_TryParse(s, R) Then
        Select Case m_ResUnit
        Case 0: m_Resolution = R * 1    'Pixel/Inch (dpi)
        Case 1: m_Resolution = R * 2.54 'Pixel/Centimeter
        End Select
        If Not m_bUpVw Then UpdateView
    Else
        MsgBox "Please give a valid numerical value: " & vbCrLf & s
    End If
End Sub

Private Sub CmbPixel_Click()
    m_DimUnit = CmbPixel.ListIndex
    If Not m_bUpVw Then UpdateView
End Sub

Private Sub CmbResolution_Click()
    m_ResUnit = CmbResolution.ListIndex
    If Not m_bUpVw Then UpdateView
End Sub

Private Sub CmbPixelFormat_Click()
    m_PixFmt = NewEPixelFormat(CmbPixelFormat)
    If Not m_bUpVw Then UpdateView
End Sub
