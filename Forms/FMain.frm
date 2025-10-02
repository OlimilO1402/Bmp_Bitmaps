VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Bitmaps"
   ClientHeight    =   8175
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15510
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1034
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.ComboBox CmbZoom 
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   0
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   360
      Width           =   4215
   End
   Begin VB.PictureBox PanelBmp 
      BackColor       =   &H00400040&
      Height          =   6735
      Left            =   4200
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   589
      TabIndex        =   11
      Top             =   360
      Width           =   8895
      Begin VB.PictureBox PBBitmap 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400040&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6495
         Left            =   0
         OLEDropMode     =   1  'Manuell
         ScaleHeight     =   433
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   577
         TabIndex        =   12
         ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
         Top             =   0
         Width           =   8655
      End
   End
   Begin VB.PictureBox PnlSideRight 
      Align           =   4  'Rechts ausrichten
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8175
      Left            =   14295
      ScaleHeight     =   545
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton BtnSelColorChangeForeBack 
         Caption         =   "^>"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   9
         Top             =   4590
         Width           =   420
      End
      Begin VB.PictureBox PBTest16bpp2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   15
         Top             =   6960
         Width           =   375
      End
      Begin VB.PictureBox PBTest16bpp1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   14
         Top             =   6960
         Width           =   375
      End
      Begin VB.CommandButton BtnTest16bpp 
         Caption         =   "Test 16bpp"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   6240
         Width           =   855
      End
      Begin VB.CommandButton BtnTestFileSave 
         Caption         =   "Test Save"
         Height          =   735
         Left            =   240
         TabIndex        =   10
         Top             =   5400
         Width           =   735
      End
      Begin VB.PictureBox PbSelColorFore 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   585
         TabIndex        =   7
         Top             =   4200
         Width           =   615
      End
      Begin VB.PictureBox PbSelColorBack 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   540
         ScaleHeight     =   585
         ScaleWidth      =   405
         TabIndex        =   8
         Top             =   4395
         Width           =   435
      End
      Begin VB.CommandButton BtnPickAColor 
         Caption         =   "Pick a Color"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton BtnClone 
         Caption         =   "Clone >>"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox PbColorSelect 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   240
         Picture         =   "FMain.frx":1782
         ScaleHeight     =   1485
         ScaleWidth      =   720
         TabIndex        =   6
         Top             =   2640
         Width           =   720
      End
      Begin VB.PictureBox PBCurColor 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   300
         ScaleHeight     =   405
         ScaleWidth      =   585
         TabIndex        =   4
         Top             =   765
         Width           =   615
      End
      Begin VB.Label LblCurColor 
         Alignment       =   2  'Zentriert
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   60
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuOpenBmpFolder 
         Caption         =   "Open bmp subfolder"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuFileImportTwain 
            Caption         =   "&Twain"
            Begin VB.Menu mnuFileImportTwainSelectSource 
               Caption         =   "&Select Source..."
            End
            Begin VB.Menu mnuFileImportTwainRead 
               Caption         =   "&Read..."
            End
         End
         Begin VB.Menu mnuFileImportWIA 
            Caption         =   "&WIA"
            Begin VB.Menu mnuFileImportWIASelectSource 
               Caption         =   "&Select Source..."
            End
            Begin VB.Menu mnuFileImportWIARead 
               Caption         =   "&Read"
            End
         End
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditResize 
         Caption         =   "Resize"
      End
      Begin VB.Menu mnuEditPalette 
         Caption         =   "Palette"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewZoomNormal 
         Caption         =   "Normal 1:1"
      End
      Begin VB.Menu mnuViewZoomIn_ 
         Caption         =   "Zoom In"
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "2:1"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "3:1"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "4:1"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "5:1"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "6:1"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "7:1"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "8:1"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "9:1"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "10:1"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "11:1"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "12:1"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "13:1"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "14:1"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "15:1"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "16:1"
            Index           =   16
         End
      End
      Begin VB.Menu mnuViewZoomOut_ 
         Caption         =   "Zoom Out"
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:2"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:3"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:4"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:5"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:6"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:7"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:8"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:9"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:10"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:11"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:12"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:13"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:14"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:15"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:16"
            Index           =   16
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " &? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type

'Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
'Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As LongPtr, ByVal X As Long, ByVal Y As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
'Private CurMousePos As POINTAPI

Private m_ScanTwain As ScannerTwain
Private m_ScanWIA   As ScannerWIA

Private m_Bmp      As Bitmap
Private m_PBZoom   As PictureBoxZoom
Private mColorSel  As ColorSelector

Private m_PFNTests As Collection

Private Sub BtnTest16bpp_Click()
    Dim cdlg As ColorDialog: Set cdlg = New ColorDialog
    If cdlg.ShowDialog(Me) = vbCancel Then Exit Sub
    Dim Col As Long: Col = cdlg.Color
    PBTest16bpp1.BackColor = Col
    Dim RGB555 As RGB555: RGB555 = LngColor_ToRGB555(LngColor(Col))
    PBTest16bpp2.BackColor = MColor.RGB555_ToLngColor(RGB555).Value
End Sub

Private Sub Form_Load()
    Set m_ScanTwain = MNew.ScannerTwain(Me)
    Set m_ScanWIA = New ScannerWIA

    PFNTests_AddFiles
    mnuEditPalette.Enabled = False
    BtnPickAColor.Enabled = False
    BtnClone.Enabled = False
    UpdateFormCaption
    InitZoom
    Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.PBBitmap, Nothing)
    Set mColorSel = MNew.ColorSelector(Timer1, BtnPickAColor, PBCurColor, LblCurColor)
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim W As Single: W = Text1.Width - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
    L = W:    W = Me.ScaleWidth - W - PnlSideRight.Width
    If W > 0 And H > 0 Then PanelBmp.Move L, T, W, H
    If W > 0 And H > 0 Then PBBitmap.Move 0, 0, W, H
End Sub

Private Sub BtnSelColorChangeForeBack_Click()
    Dim Color As Long: Color = Me.PbSelColorBack.BackColor
    Me.PbSelColorBack.BackColor = Me.PbSelColorFore.BackColor
    Me.PbSelColorFore.BackColor = Color
End Sub

Private Sub BtnTestFileSave_Click()
    Dim v, PFN As String
    Debug.Print m_PFNTests.Count
    For Each v In m_PFNTests
        PFN = v
        If FileExists(PFN) Then
            Debug.Print PFN
            TestBmp PFN
        Else
            Debug.Print "File does not exists: " & vbCrLf & PFN
        End If
    Next
End Sub

Private Sub TestBmp(PFN As String)
Try: On Error GoTo Catch
    If Len(PFN) = 0 Then Exit Sub
    Dim bmp0 As Bitmap: Set bmp0 = MNew.Bitmap(PFN)
    Dim tmpPFN As String: tmpPFN = Environ("tmp") & "\test.bmp"
    If FileExists(tmpPFN) Then Kill tmpPFN
    bmp0.Save tmpPFN
    Dim data0() As Byte: ReadFileContentBuffer PFN, data0
    Dim Data1() As Byte: ReadFileContentBuffer tmpPFN, Data1
    Dim L0 As Long: L0 = UBound(data0) + 1
    Dim l1 As Long: l1 = UBound(Data1) + 1
    If L0 <> l1 Then
        MsgBox "The length ist not equal: l0=" & L0 & " <> l1=" & l1
    End If
    Dim c As Long: c = RtlCompareMemory(data0(0), Data1(0), L0)
    Dim diff As Long: diff = Abs(L0 - c)
    If diff = 0 Then Debug.Print "diff=0 OK data0 and data1 is identical"
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub

Private Function FileExists(ByVal PFN As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(PFN) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Private Sub ReadFileContentBuffer(PFN As String, Buffer() As Byte)
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Read As FNr
    ReDim Buffer(0 To LOF(FNr) - 1)
    Get FNr, , Buffer
    GoTo Finally
Catch: MsgBox Err.Description
Finally: Close FNr
End Sub

Private Sub PFNTests_AddFiles()
    Set m_PFNTests = New Collection
    Dim FNm As String, Path0 As String: Path0 = App.Path & "\bmps\"
    Dim PFN As String, Path1 As String, Path As String
    
    Path1 = "OS2\":    Path = Path0 & Path1

    FNm = "PSPColors_OS2_01bpp.bmp":         m_PFNTests.Add Path & FNm
    FNm = "PSPColors_OS2_04bpp.bmp":         m_PFNTests.Add Path & FNm
    FNm = "PSPColors_OS2_08bpp.bmp":         m_PFNTests.Add Path & FNm
    FNm = "PSPColors_OS2_24bpp.bmp":         m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_OS2_01bpp.bmp":    m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_OS2_04bpp.bmp":    m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_OS2_08bpp.bmp":    m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_OS2_24bpp.bmp":    m_PFNTests.Add Path & FNm
    
    Path1 = "Win\RGB\":    Path = Path0 & Path1

    FNm = "PSPColors_Win_01bpp.bmp":                      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_04bpp.bmp":                      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_08bpp.bmp":                      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_ARGB1555.bmp":             m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_ARGB1555_woCSType.bmp":    m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_RGB555.bmp":               m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_RGB555_woCSType.bmp":      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_RGB565.bmp":               m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_16bpp_RGB565_woCSType.bmp":      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_24bpp.bmp":                      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_32bpp.bmp":                      m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_32bpp_ARGB.bmp":                 m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_32bpp_ARGB_woCSType.bmp":        m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_32bpp_RGB.bmp":                  m_PFNTests.Add Path & FNm
    FNm = "PSPColors_Win_32bpp_RGB_woCSType.bmp":         m_PFNTests.Add Path & FNm

    FNm = "SleepPolarBear_Win_01bpp.bmp":                 m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_04bpp.bmp":                 m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_08bpp.bmp":                 m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_16bpp_ARGB1555.bmp":        m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_16bpp_RGB555.bmp":          m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_16bpp_RGB565.bmp":          m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_24bpp.bmp":                 m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_32bpp_ARGB.bmp":            m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_32bpp_XRGB.bmp":            m_PFNTests.Add Path & FNm
    FNm = "SleepPolarBear_Win_32bpp_XRGB_woCSInfo.bmp":   m_PFNTests.Add Path & FNm

End Sub

'Private Function Hex2(b As Byte) As String
'    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
'End Function

Public Function Clone() As FMain
    Set Clone = New FMain
    Clone.NewC m_Bmp
End Function

Friend Sub NewC(other As Bitmap)
    Set m_Bmp = other.Clone
    Me.Show
End Sub

Private Sub BtnClone_Click()
    If m_Bmp Is Nothing Then Exit Sub
    Dim NewForm As FMain: Set NewForm = Me.Clone
    NewForm.UpdateView
End Sub

Private Sub UpdateFormCaption()
    Dim PFN As String
    If Not m_Bmp Is Nothing Then PFN = m_Bmp.FileName
    Me.Caption = "Bitmaps" & " v" & App.Major & "." & App.Minor & "." & App.Revision & IIf(Len(PFN), " - " & PFN, "")
End Sub

'Private Sub BtnPickAColor_Click()
'    m_bPickAColor = True
'End Sub

' v ############################## v '    mnuFile    ' v ############################## v '
Private Sub mnuFileNew_Click()
    MiddlePosDlg FDlgNewPicture
    Dim bmp As Bitmap
    'If Not m_Bmp Is Nothing Then Set Bmp = m_Bmp.Clone
    If FDlgNewPicture.ShowDialog(Me, bmp) = vbCancel Then Exit Sub
    Set m_Bmp = bmp
    
    UpdateView
End Sub

Private Sub MiddlePosDlg(Frm As Form)
    Dim W As Single: W = Frm.Width
    Dim H As Single: H = Frm.Height
    Dim L As Single: L = Me.Left + (Me.Width - W) / 2
    Dim T As Single: T = Me.Top + (Me.Height - H) / 2
    Frm.Move L, T
End Sub

Private Sub mnuFileOpen_Click()
Try: On Error GoTo Catch
    'Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    'OFD.Filter = "Bitmaps (*.bmp)|*.bmp|All files (*.*)|*.*"
    'If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    'Dim PFN As String: PFN = OFD.FileName
    'Dim FD As New OpenFileDialog
    'If Not m_Bmp Is Nothing Then FD.FileName = m_Bmp.FileName
    Dim aPFN As String: If Not m_Bmp Is Nothing Then aPFN = m_Bmp.FileName
    If aPFN = "<Clipboard>" Then aPFN = vbNullString
    aPFN = MMain.GetOpenFileName(Me, aPFN)
    If Len(aPFN) = 0 Then Exit Sub
    Dim pos As Long:   pos = InStrRev(aPFN, ".")
    Dim ext As String: ext = LCase(Right(aPFN, Len(aPFN) - pos))
    Dim pic As StdPicture
    If ext = "bmp" Then
        Set m_Bmp = MNew.Bitmap(aPFN)
    Else
        Select Case ext
        Case "png": Set pic = MLoadPng.LoadPictureGDIp(aPFN)
        Case "gif"
                    'Set PBBitmap.Picture = LoadPicture(aPFN)
                    'Dim ipd As IPictureDisp: Set ipd = LoadPicture(aPFN)
                    'Set PBBitmap.Picture = ipd
                    'Dim sdp As StdPicture: Set sdp = LoadPicture(aPFN)
                    Set pic = LoadPicture(aPFN)
                    'Set PBBitmap.Picture = sdp
                    'UpdateView
                    'Exit Sub
        ', "jpg": Set pic = LoadPicture(aPFN)
        Case Else 'Just give it a try
                    Set pic = LoadPicture(aPFN)
        End Select
        Set m_Bmp = MNew.BitmapSP(pic, aPFN)
    End If
    'Set m_PBZoom.Image = m_Bmp.ToPicture
    UpdateView
    Exit Sub
Catch:
    MsgBox Err.Description & vbCrLf & _
           aPFN
End Sub

Private Sub mnuFileSave_Click()
Try: On Error GoTo Catch
    If m_Bmp Is Nothing Then Exit Sub
    m_Bmp.Save
    GoTo Finally
Catch:
    MsgBox Err.Description
Finally:
End Sub

Private Sub mnuFileSaveAs_Click()
    'Dim FD As New SaveFileDialog
    'If Not m_Bmp Is Nothing Then FD.FileName = m_Bmp.FileName
    Dim PFN As String: If Not m_Bmp Is Nothing Then PFN = m_Bmp.FileName
    PFN = MMain.GetSaveFileName(Me, PFN)
    If Len(PFN) = 0 Then Exit Sub
    m_Bmp.Save PFN
    UpdateView
End Sub

Private Sub mnuOpenBmpFolder_Click()
    Dim p As String: p = App.Path & "\bmps\"
    If MsgBox("Open folder?" & vbCrLf & p, vbOKCancel) = vbCancel Then Exit Sub
    Shell "Explorer.exe " & p, vbNormalFocus
End Sub

' FileImport
Private Sub mnuFileImportTwainSelectSource_Click()
    If Not m_ScanTwain.SelectDevice Then
        MsgBox "An error occured, maybe EZTW32.DLL not found, make sure this file can be found in the search-path."
    End If
End Sub

Private Sub mnuFileImportTwainRead_Click()
    GetScannedImage m_ScanTwain
End Sub

Private Sub mnuFileImportWIASelectSource_Click()
    Dim DeviceNames() As String
    If Not m_ScanWIA.TryGetDeviceNames(DeviceNames) Then
        MsgBox "No devices found!"
        Exit Sub
    End If
    Dim SelDevice As String
    If FSelect.ShowDialog(Me, "Select Source", DeviceNames, SelDevice) = vbCancel Then Exit Sub
    m_ScanWIA.SelectDevice SelDevice
End Sub

Private Sub mnuFileImportWIAProperties_Click()
    m_ScanWIA.ShowDevicePropertiesDialog
End Sub

Private Sub mnuFileImportWIARead_Click()
    GetScannedImage m_ScanWIA
End Sub


Private Sub mnuFileExit_Click()
    Unload Me
End Sub
' ^ ############################## ^ '    mnuFile    ' ^ ############################## ^ '

' v ############################## v '    mnuEdit    ' v ############################## v '
Private Sub mnuEditCut_Click()
    'copy all to clipboard and remove
    If m_Bmp Is Nothing Then Exit Sub
    Clipboard.SetData m_Bmp.ToPicture, ClipBoardConstants.vbCFBitmap
    'Clipboard.SetData m_Bmp.ToPicture, ClipBoardConstants.vbCFDIB 'which one is correct
    Set m_Bmp = Nothing
    UpdateView
End Sub

Private Sub mnuEditCopy_Click()
    'copy all to clipboard
    If m_Bmp Is Nothing Then Exit Sub
    Clipboard.SetData m_Bmp.ToPicture, ClipBoardConstants.vbCFBitmap
End Sub

Private Sub mnuEditPaste_Click()
    'paste from clipboard and create new
    
    Dim bBmp As Boolean: bBmp = Clipboard.GetFormat(ClipBoardConstants.vbCFBitmap)
    Dim bDIB As Boolean: bDIB = Clipboard.GetFormat(ClipBoardConstants.vbCFDIB)
        
    If (Not bBmp) And (Not bDIB) Then
        MsgBox "Neither bitmap- nor dib-data in clipboard"
        Exit Sub
    End If
    
    Dim pic As StdPicture
    If bBmp Then
        'MsgBox "Trying to read bitmap from clipboard"
        Set pic = Clipboard.GetData(ClipBoardConstants.vbCFBitmap)
        If pic Is Nothing Then
            MsgBox "Could not read bmp from clipboard"
            Exit Sub
        End If
    ElseIf bDIB Then
        'MsgBox "Trying to read dib from clipboard"
        Set pic = Clipboard.GetData(ClipBoardConstants.vbCFDIB)
        If pic Is Nothing Then
            MsgBox "Could not read dib from clipboard"
            Exit Sub
        End If
    End If
    
    If m_Bmp Is Nothing Then
        Set m_Bmp = MNew.BitmapSP(pic, "<Clipboard>")
    Else
        m_Bmp.NewSP pic, "<Clipboard>"
    End If
    UpdateView
End Sub

Private Sub mnuEditPalette_Click()
    If m_Bmp Is Nothing Then Exit Sub
    FPalette.Move Me.Left + Me.Width / 2 - FPalette.Width / 2, Me.Top + Me.Height / 2 - FPalette.Height / 2
    If FPalette.ShowDialog(Me, m_Bmp) = vbCancel Then Exit Sub
    UpdateView
End Sub

Private Sub mnuEditResize_Click()
    MiddlePosDlg FDlgNewPicture
    Dim bmp As Bitmap
    If Not m_Bmp Is Nothing Then Set bmp = m_Bmp.Clone
    If FDlgNewPicture.ShowDialog(Me, bmp) = vbCancel Then Exit Sub
    Set m_Bmp = bmp
    UpdateView
End Sub
' ^ ############################## ^ '    mnuEdit    ' ^ ############################## ^ '

' v ' ############################## ' v '    mnuView    ' v ' ############################## ' v '
' v ' #################### ' v '    Controls for PictureBoxZoom    ' v ' #################### ' v '
Sub InitZoom()
    Dim i As Long
    For i = 16 To 1 Step -1: CmbZoom.AddItem CStr(i) & ":1": Next
    For i = 2 To 16:         CmbZoom.AddItem "1:" & CStr(i): Next
    CmbZoom.ListIndex = 15
End Sub

Private Sub mnuViewZoomNormal_Click()
    CmbZoom.ListIndex = 15
End Sub

Private Sub mnuViewZoomIn_Click(Index As Integer)
    CmbZoom.ListIndex = 16 - Index
End Sub

Private Sub mnuViewZoomOut_Click(Index As Integer)
    CmbZoom.ListIndex = Index + 14
End Sub

Private Sub CmbZoom_Click()
    If m_PBZoom Is Nothing Then Exit Sub
    Dim li As Long:     li = CmbZoom.ListIndex - 15
    Dim z As Double: z = IIf(li = 0, 1, IIf(li < 1, Abs(li) + 1, 1))
    Dim n As Double: n = IIf(li = 0, 1, IIf(li < 1, 1, Abs(li) + 1))
    m_PBZoom.ZoomFactor = z / n
    mnuViewZoom_UnCheckAll
    If z = 1 And n = 1 Then
        mnuViewZoomNormal.Checked = True
    ElseIf z = 1 Then
        mnuViewZoomOut(CLng(n)).Checked = True
    ElseIf n = 1 Then
        mnuViewZoomIn(CLng(z)).Checked = True
    End If
End Sub

Private Sub mnuViewZoom_UnCheckAll()
    Dim i As Integer
    For i = mnuViewZoomIn.LBound To mnuViewZoomIn.UBound:   mnuViewZoomIn(i).Checked = False:  Next
    For i = mnuViewZoomOut.LBound To mnuViewZoomOut.UBound: mnuViewZoomOut(i).Checked = False: Next
    mnuViewZoomNormal.Checked = False
End Sub

' ^ ############################## ^ '    mnuView    ' ^ ############################## ^ '

' v ############################## v '    mnuHelp    ' v ############################## v '
Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub
' ^ ############################## ^ '    mnuHelp    ' ^ ############################## ^ '

Private Sub PbColorSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case MouseButtonConstants.vbLeftButton:  PbSelColorFore.BackColor = PBCurColor.BackColor
    Case MouseButtonConstants.vbRightButton: PbSelColorBack.BackColor = PBCurColor.BackColor
    End Select
End Sub

'Private Sub PbColorSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Dim Color As Long: Color = PbColorSelect.Point(X, Y)
'    If Color < 0 Then Exit Sub
'    PBCurColor.BackColor = Color
'    LblCurColor.Caption = MouseCoordsNColor_ToStr(X, Y, Color)
'End Sub

'Private Sub PBBitmap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If m_Bmp Is Nothing Then Exit Sub
'    If Not m_bPickAColor Then Exit Sub
'    Dim Color As Long: Color = m_Bmp.Pixel(x, y)
'    PBCurColor.BackColor = Color
'    Dim s As String
'    If m_Bmp.IsIndexed Then
'        s = "Index: " & m_Bmp.PalettePixelIndex(x, y) & vbCrLf
'    End If
'    s = s & MouseCoordsNColor_ToStr(x, y, Color)
'    LblCurColor.Caption = s
'End Sub
'Private Function MouseCoordsNColor_ToStr(X As Single, Y As Single, ByVal Color As Long) As String
'    MouseCoordsNColor_ToStr = "X;Y: " & X & ";" & Y & vbCrLf & Color_ToStr(Color)
'End Function
    
Private Sub PBBitmap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If Not m_bPickAColor Then Exit Sub
    Select Case Button
    Case MouseButtonConstants.vbLeftButton:  PbSelColorFore.BackColor = PBCurColor.BackColor
    Case MouseButtonConstants.vbRightButton: PbSelColorBack.BackColor = PBCurColor.BackColor
    End Select
End Sub

'Private Function Color_ToStr(ByVal this As Long) As String
'    Dim r As Long: r = (this And &HFF&)
'    Dim G As Long: G = (this And &HFF00&) \ &H100&
'    Dim b As Long: b = (this And &HFF0000) \ &H10000
'    Dim hexprefix As String: hexprefix = "&&H"
'    Dim sr As String: sr = CStr(r): sr = Space$(3 - Len(sr)) & sr
'    Dim sG As String: sG = CStr(G): sG = Space$(3 - Len(sG)) & sG
'    Dim sB As String: sB = CStr(b): sB = Space$(3 - Len(sB)) & sB
'    Color_ToStr = "R=" & sr & " (" & hexprefix & MString.Hex2(CByte(r)) & ")" & vbCrLf & _
'                  "G=" & sG & " (" & hexprefix & MString.Hex2(CByte(G)) & ")" & vbCrLf & _
'                  "B=" & sB & " (" & hexprefix & MString.Hex2(CByte(b)) & ")"
'End Function

'Private Sub PBBitmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If m_bPickAColor Then m_bPickAColor = False
'End Sub

Private Sub PBBitmap_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub PbSelColorBack_Click()
    Dim Color As Long: Color = PbSelColorBack.BackColor
    PbSelColorBack.BackColor = ColorDlg(Color)
End Sub

Private Sub PbSelColorFore_Click()
    Dim Color As Long: Color = PbSelColorFore.BackColor
    PbSelColorFore.BackColor = ColorDlg(Color)
End Sub

Private Function ColorDlg(ByVal CurColor As Long) As Long
    Dim CD As ColorDialog: Set CD = New ColorDialog: CD.Color = CurColor
    ColorDlg = IIf(CD.ShowDialog(Me) = vbOK, CD.Color, CurColor)
End Function

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub PanelBmp_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub AllOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim PFN As String: PFN = Data.Files(1)
    Dim ext As String: ext = LCase(Right(PFN, 3))
    Dim pic As StdPicture
    If ext = "bmp" Then
        Set m_Bmp = MNew.Bitmap(PFN)
        'Set pic = m_Bmp.ToPicture
    ElseIf ext = "png" Then
        Set pic = MLoadPng.LoadPictureGDIp(PFN)
        Set m_Bmp = MNew.BitmapSP(pic, PFN)
        'Set PBBitmap.Picture = MLoadPng.LoadPictureGDIp(PFN)
    ElseIf ext = "jpg" Then
        Set pic = MLoadPng.LoadPictureGDIp(PFN)
        Set m_Bmp = MNew.BitmapSP(pic, PFN)
        'Set PBBitmap.Picture = MLoadPng.LoadPictureGDIp(PFN)
    ElseIf ext = "gif" Then
        Set pic = MLoadPng.LoadPictureGDIp(PFN)
        Set m_Bmp = MNew.BitmapSP(pic, PFN)
        'Set PBBitmap.Picture = MLoadPng.LoadPictureGDIp(PFN)
    End If
    'Set m_PBZoom.Image = pic
    UpdateView
End Sub

Public Sub UpdateView()
    Dim dt As Single: dt = Timer
    dt = Timer - dt
    BtnClone.Enabled = Not m_Bmp Is Nothing
    If m_Bmp Is Nothing Then Exit Sub
    
    Set m_PBZoom.Image = m_Bmp.ToPicture

    'Set PBBitmap.Picture = m_Bmp.ToPicture
    'Label1.Caption = "File loading time t: " & dt & "sec;"
    UpdateFormCaption
    Text1.Text = m_Bmp.ToStr
    mnuEditPalette.Enabled = m_Bmp.IsIndexed
    BtnPickAColor.Enabled = True
    'BtnClone.Enabled = True
End Sub

Public Sub GetScannedImage(ImageScanner)
    Dim img As IPictureDisp: Set img = ImageScanner.Scan
    If img Is Nothing Then MsgBox "Image not found!": Exit Sub
    If img = 0 Then MsgBox "Image not found!": Exit Sub
    'Set m_Image = img
    Set m_Bmp = MNew.BitmapSP(img, "<Scanned Image>")
    UpdateView
    'If m_PBZoom Is Nothing Then
    '    Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.Picture1, m_Image)
    'Else
    '    Set m_PBZoom.Image = m_Image
    'End If
End Sub

' v ' ############################## ' v '    Pick a Color    ' v ' ############################## ' v '
'Private Sub Timer1_Timer()
'    GetCursorPos CurMousePos
'    Dim Color As Long: Color = ColorUnderMouse(CurMousePos.X, CurMousePos.Y)
'    PBCurColor.BackColor = Color
'    LblCurColor.Caption = MouseCoordsNColor_ToStr(X, Y, Color)
'    'UpdateView
'End Sub

