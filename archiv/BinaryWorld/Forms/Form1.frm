VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "API Demo "
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7665
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows-Standard
   Begin VB.HScrollBar RGB 
      Height          =   195
      Index           =   2
      Left            =   5265
      TabIndex        =   20
      Top             =   4965
      Width           =   2325
   End
   Begin VB.HScrollBar RGB 
      Height          =   195
      Index           =   1
      Left            =   5265
      TabIndex        =   19
      Top             =   4725
      Width           =   2325
   End
   Begin VB.HScrollBar RGB 
      Height          =   195
      Index           =   0
      Left            =   5265
      TabIndex        =   18
      Top             =   4485
      Width           =   2325
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Fade Effect"
      Height          =   330
      Left            =   2685
      TabIndex        =   17
      Top             =   4455
      Width           =   1485
   End
   Begin VB.PictureBox Picture2 
      Height          =   2910
      Left            =   2685
      ScaleHeight     =   2850
      ScaleWidth      =   2250
      TabIndex        =   14
      Top             =   1440
      Width           =   2310
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3540
      Left            =   5235
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   13
      Top             =   825
      Width           =   2370
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2700
      Style           =   2  'Dropdown-Liste
      TabIndex        =   11
      Top             =   4890
      Width           =   1485
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save BMP >>"
      Height          =   330
      Left            =   75
      TabIndex        =   9
      Top             =   4890
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1050
      TabIndex        =   8
      Top             =   825
      Width           =   3945
   End
   Begin VB.PictureBox Picture1 
      Height          =   2895
      Left            =   90
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   2250
      TabIndex        =   7
      Top             =   1455
      Width           =   2310
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load BMP"
      Height          =   330
      Left            =   75
      TabIndex        =   2
      Top             =   4455
      Width           =   1485
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00FFFFFF&
      Height          =   1560
      Left            =   45
      ScaleHeight     =   1500
      ScaleWidth      =   7545
      TabIndex        =   1
      Top             =   5310
      Width           =   7605
      Begin VB.Label l6 
         BackStyle       =   0  'Transparent
         Caption         =   "API"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   795
         Width           =   1830
      End
      Begin VB.Label l7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form1.frx":15DDA
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   45
         TabIndex        =   5
         Top             =   1005
         Width           =   7470
      End
      Begin VB.Label l5 
         BackStyle       =   0  'Transparent
         Caption         =   "This demo will show you how to work with Bitmaps and DIBs. This code use very easy to implement Bitmap and DIB class"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   60
         TabIndex        =   3
         Top             =   315
         Width           =   7470
      End
      Begin VB.Label l4 
         BackStyle       =   0  'Transparent
         Caption         =   "Example Description"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   45
         Width           =   1830
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10000
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Picture         =   "Form1.frx":15E6B
         Top             =   120
         Width           =   4140
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   4695
      TabIndex        =   23
      Top             =   4965
      Width           =   510
   End
   Begin VB.Label Label4 
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   165
      Left            =   4680
      TabIndex        =   22
      Top             =   4725
      Width           =   570
   End
   Begin VB.Label Label3 
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   165
      Left            =   4695
      TabIndex        =   21
      Top             =   4485
      Width           =   450
   End
   Begin VB.Label Label9 
      Caption         =   "Save Image"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2685
      TabIndex        =   16
      Top             =   1215
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "Original Image"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   105
      TabIndex        =   15
      Top             =   1215
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Bits/Pixel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1800
      TabIndex        =   12
      Top             =   4965
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "File Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   885
      Width           =   1050
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oTestBitmap As BITMAP    ' sample bitmap object
Dim cDib As New DIBSection
Dim cDibBuffer As New DIBSection

Private Sub Command1_Click()
    Dim sFileName As String    ' path + name of bitmap file to be loaded
    Dim bResult As Boolean    ' result of loading bitmap into bitmap object
    ' use the LoadImage API to load in a bitmap image. The name of the
    ' image file will be retrieved from the textbox control sited on the form.

    sFileName = Text1.Text

    ' create a new bitmap object for testing purposes
    Set oTestBitmap = New BITMAP

    ' attempt to load the bitmap into to into the bitmap object using its
    ' handle
    
    
    '//Method 1 (Load Empty with specified height and width)
    bResult = oTestBitmap.LoadBMP(Picture1.Width / 15, Picture1.Height / 15)
    oTestBitmap.LoadPictureBlt Picture1.hdc
    oTestBitmap.PaintPicture Picture2.hdc
    'Picture2.Refresh
    
    '//Method 2 (Load from file)
    bResult = oTestBitmap.LoadBMPFromFile(sFileName)
    
    '//Method 2 (Load from StdPic object)
    'bResult = oTestBitmap.LoadBMPFromStdPic(Picture1.Picture)
    
    ' verify the image was successfully loaded into the bitmap object.
    ' SetBitmap returns true if successful.

    If (bResult = False) Then
        MsgBox "Error : Unable To Load Image Into Bitmap Object", vbOKOnly, "Bitmap Object Error"
        Set oTestBitmap = Nothing
        Exit Sub
    End If

    Picture1.Cls
    Call oTestBitmap.PaintPicture(Picture1.hdc)
    ' refresh the picturebox to show the blit
    Picture1.Refresh

    ' test the bitmap object with the loaded image
    Command2.Enabled = True
    Command3.Enabled = True

    Dim strmsg
    strmsg = "Source Image" & vbCrLf
    strmsg = strmsg & "-------------------" & vbCrLf
    strmsg = strmsg & "Width       :" & oTestBitmap.Width & vbCrLf
    strmsg = strmsg & "Height      :" & oTestBitmap.Height & vbCrLf
    strmsg = strmsg & "Bits/pix    :" & oTestBitmap.BitCount & vbCrLf
    strmsg = strmsg & "Compression :" & oTestBitmap.Compression & vbCrLf
    strmsg = strmsg & "Size (Bytes):" & oTestBitmap.Size & vbCrLf

    Text2.Text = strmsg

    '//Create DIBs so we can work with 2D pixel array of the bitmap
    '//Main DIB which will be modified when we give some effect to Bitmap
    cDib.Create oTestBitmap.Width, oTestBitmap.Height
    cDib.LoadPictureBlt oTestBitmap.ImageDC

    ' Create a copy of Main DIB so we can restore it back to its original state
    cDibBuffer.Create cDib.Width, cDib.Height
    cDib.PaintPicture cDibBuffer.hdc
End Sub

Private Sub Command2_Click()
    Dim ret As Boolean
    Dim fPath
    fPath = App.Path & "\save_" & Combo1.Text & "bit.bmp"

    '//Paint our modified DIB in to the save Bitmap
    cDib.PaintPicture oTestBitmap.ImageDC

    ret = oTestBitmap.SaveBMP(fPath, CInt(Combo1.Text))
    If ret = True Then
        Picture2.Cls
        Picture2.Picture = LoadPicture(fPath)

        Dim strmsg
        strmsg = "Source Image" & vbCrLf
        strmsg = strmsg & "-------------------" & vbCrLf
        strmsg = strmsg & "Width       :" & oTestBitmap.Width & vbCrLf
        strmsg = strmsg & "Height      :" & oTestBitmap.Height & vbCrLf
        strmsg = strmsg & "Bits/pix    :" & oTestBitmap.BitCount & vbCrLf
        strmsg = strmsg & "Compression :" & oTestBitmap.Compression & vbCrLf
        strmsg = strmsg & "Size (Bytes):" & oTestBitmap.Size & vbCrLf & vbCrLf

        strmsg = strmsg & "Save Image" & vbCrLf
        strmsg = strmsg & "-------------------" & vbCrLf
        strmsg = strmsg & "Width       :" & oTestBitmap.SaveBMPWidth & vbCrLf
        strmsg = strmsg & "Height      :" & oTestBitmap.SaveBMPHeight & vbCrLf
        strmsg = strmsg & "Bits/pix    :" & oTestBitmap.SaveBMPBitCount & vbCrLf
        strmsg = strmsg & "Compression :" & oTestBitmap.SaveBMPCompression & vbCrLf
        strmsg = strmsg & "Size (Bytes):" & oTestBitmap.SaveBMPSize & vbCrLf

        Text2.Text = strmsg
    Else
        MsgBox "Error during save"
    End If
End Sub

Private Sub Command3_Click()
    Dim i As Long, hDidRet As Long

    ' Fade Loop:
    For i = 0 To 255 Step 5
        ' Fade the dib by amount i:
        cDib.Fade i

        ' Draw it:
        cDib.PaintPicture Picture2.hdc

        ' Breathe a little. You may have to put a slowdown here:
        DoEvents

        ' Reset for next fade using original copy which is stored in cDibBuffer
        cDibBuffer.PaintPicture cDib.hdc
    Next i
End Sub

Private Sub Form_Load()
    'disable the test buttons
    Command2.Enabled = False
    Command3.Enabled = False

    Combo1.AddItem "1"
    Combo1.AddItem "4"
    Combo1.AddItem "8"
    Combo1.AddItem "16"
    Combo1.AddItem "24"
    Combo1.ListIndex = 4

    RGB(0).Max = 255
    RGB(1).Max = 255
    RGB(2).Max = 255

    ' get the path of the test bitmap
    Text1.Text = App.Path & "\24bit.bmp"

    ' set the picturebox to autoredraw so the blits will show up
    Picture1.AutoRedraw = True
End Sub

Private Sub RGB_Change1(Index As Integer)
    Dim R As Byte, G As Byte, B As Byte, Row As Long, Col As Long

    cDib.LoadPictureBlt cDibBuffer.hdc

    For Row = 1 To cDib.Height
        For Col = 1 To cDib.Width
            Call cDib.GetPixelRGB(Row, Col, R, G, B)    '//read RGB values into variables
            If Index = 0 Then
                Call cDib.SetPixelRGB(Row, Col, RGB(0).value)    ''//Modify RGB values
            ElseIf Index = 1 Then
                Call cDib.SetPixelRGB(Row, Col, , RGB(1).value)    ''//Modify RGB values
            ElseIf Index = 2 Then
                Call cDib.SetPixelRGB(Row, Col, , , RGB(2).value)    ''//Modify RGB values
            End If
        Next
        'strmsg = strmsg & vbCrLf
        'Me.Caption = i
    Next

    'Call Effect1
    cDib.PaintPicture Picture2.hdc
End Sub
Private Sub RGB_Change(Index As Integer)
    Dim R As Byte, G As Byte, B As Byte, Row As Long, Col As Long

    cDib.LoadPictureBlt cDibBuffer.hdc
    'oTestBitmap.
    For Row = 1 To cDib.Height
        For Col = 1 To cDib.Width
            Call cDib.GetPixelRGB(Row, Col, R, G, B)    '//read RGB values into variables
            '//Just average value of R,G,B with selected values
            Call cDib.SetPixelRGB(Row, Col, (R + RGB(0).value) / 2, (G + RGB(1).value) / 2, (B + RGB(2).value) / 2) ''//Modify RGB values
        Next
        'strmsg = strmsg & vbCrLf
        'Me.Caption = i
    Next

    cDib.PaintPicture Picture2.hdc
End Sub

