VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Bitmaps"
   ClientHeight    =   8670
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14895
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   993
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PnlSideRight 
      Align           =   4  'Rechts ausrichten
      BorderStyle     =   0  'Kein
      Height          =   8670
      Left            =   13680
      ScaleHeight     =   578
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   2
      Top             =   0
      Width           =   1215
      Begin VB.PictureBox PBSelectForeBackColor 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   300
         ScaleHeight     =   540
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   5415
         Width           =   615
      End
      Begin VB.CommandButton BtnSelColorChangeForeBack 
         Caption         =   "^>"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   5010
         Width           =   375
      End
      Begin VB.PictureBox PbSelColorFore 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   120
         ScaleHeight     =   540
         ScaleWidth      =   585
         TabIndex        =   8
         Top             =   4440
         Width           =   615
      End
      Begin VB.PictureBox PbSelColorBack 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   480
         ScaleHeight     =   540
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   4800
         Width           =   615
      End
      Begin VB.CommandButton BtnPickAColor 
         Caption         =   "Pick a Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton BtnClone 
         Caption         =   "Clone >>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
      Begin VB.PictureBox PbColorSelect 
         Appearance      =   0  '2D
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   1485
         Left            =   240
         Picture         =   "FMain.frx":1782
         ScaleHeight     =   1485
         ScaleWidth      =   720
         TabIndex        =   7
         Top             =   2880
         Width           =   720
      End
      Begin VB.PictureBox PBCurColor 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   300
         ScaleHeight     =   540
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   780
         Width           =   615
      End
      Begin VB.Label LblSelColor 
         Alignment       =   2  'Zentriert
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   0
         TabIndex        =   12
         Top             =   6075
         Width           =   1215
      End
      Begin VB.Label LblCurColor 
         Alignment       =   2  'Zentriert
         Caption         =   ". . ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   0
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
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
      TabIndex        =   1
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   0
      Width           =   4215
   End
   Begin VB.PictureBox PBBitmap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      BorderStyle     =   0  'Kein
      Height          =   6735
      Left            =   4200
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   0
      Width           =   7695
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditResize 
         Caption         =   "Resize"
      End
      Begin VB.Menu mnuEditPalette 
         Caption         =   "Palette"
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
'Private m_PFN As String
Private m_Bmp As Bitmap
Private m_bPickAColor As Boolean

Private Sub Form_Load()
    mnuEditPalette.Enabled = False
    BtnPickAColor.Enabled = False
    BtnClone.Enabled = False
    UpdateFormCaption
End Sub

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

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = Text1.Top
    Dim w As Single: w = Text1.Width - L
    Dim h As Single: h = Me.ScaleHeight - T
    If w > 0 And h > 0 Then Text1.Move L, T, w, h
    L = w:    w = Me.ScaleWidth - w
    If w > 0 And h > 0 Then PBBitmap.Move L, T, w, h
End Sub

Private Sub BtnPickAColor_Click()
    m_bPickAColor = True
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

Private Sub mnuFileNew_Click()
    MiddlePosDlg FDlgNewPicture
    Dim bmp As Bitmap
    'If Not m_Bmp Is Nothing Then Set Bmp = m_Bmp.Clone
    If FDlgNewPicture.ShowDialog(Me, bmp) = vbCancel Then Exit Sub
    Set m_Bmp = bmp
    UpdateView
End Sub

Private Sub MiddlePosDlg(Frm As Form)
    Dim w As Single: w = Frm.Width
    Dim h As Single: h = Frm.Height
    Dim L As Single: L = Me.Left + (Me.Width - w) / 2
    Dim T As Single: T = Me.Top + (Me.Height - h) / 2
    Frm.Move L, T
End Sub
Private Sub mnuFileOpen_Click()
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    OFD.Filter = "Bitmaps (*.bmp)|*.bmp|All files (*.*)|*.*"
    If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    Dim PFN As String: PFN = OFD.FileName
    Set m_Bmp = MNew.Bitmap(PFN)
    UpdateView
End Sub

Private Sub mnuOpenBmpFolder_Click()
    Dim p As String: p = App.Path & "\bmps\"
    If MsgBox("Open folder?" & vbCrLf & p, vbOKCancel) = vbCancel Then Exit Sub
    Shell "Explorer.exe " & p, vbNormalFocus
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub PbColorSelect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case MouseButtonConstants.vbLeftButton:  PbSelColorFore.BackColor = PBSelectForeBackColor.BackColor
    Case MouseButtonConstants.vbRightButton: PbSelColorBack.BackColor = PBSelectForeBackColor.BackColor
    End Select
End Sub

Private Sub PbColorSelect_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Color As Long: Color = PbColorSelect.Point(X, Y)
    PBSelectForeBackColor.BackColor = Color
    LblSelColor.Caption = "X: " & X & "; Y: " & Y & vbCrLf & Color_ToStr(Color)
End Sub

Private Sub PBBitmap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Bmp Is Nothing Then Exit Sub
    If m_bPickAColor Then
        Dim Color As Long: Color = m_Bmp.Pixel(X, Y)
        PBCurColor.BackColor = Color
        Dim s As String
        If m_Bmp.IsIndexed Then
            s = "Index: " & m_Bmp.PalettePixelIndex(X, Y)
        End If
        s = s & "X: " & X & "; Y: " & Y & vbCrLf & Color_ToStr(Color)
        LblCurColor.Caption = s
    End If
End Sub

Private Function Color_ToStr(ByVal this As Long) As String
    Dim R As Long: R = (this And &HFF&)
    Dim G As Long: G = (this And &HFF00&) \ &H100&
    Dim b As Long: b = (this And &HFF0000) \ &H10000
    Dim hexprefix As String: hexprefix = "&&H"
    Color_ToStr = "R=" & R & " (" & hexprefix & Hex(R) & ")" & vbCrLf & _
                  "G=" & G & " (" & hexprefix & Hex(G) & ")" & vbCrLf & _
                  "B=" & b & " (" & hexprefix & Hex(b) & ")"
End Function


Private Sub PBBitmap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bPickAColor Then m_bPickAColor = False
End Sub

Private Sub PBBitmap_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Picture4_Click()

End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub AllOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim PFN As String: PFN = Data.Files(1)
    Dim ext As String: ext = LCase(Right(PFN, 3))
    If ext = "bmp" Then
        Set m_Bmp = MNew.Bitmap(PFN)
        UpdateView
    ElseIf ext = "png" Then
        Set PBBitmap.Picture = MLoadPng.LoadPictureGDIp(PFN)
    End If
End Sub

Public Sub UpdateView()
    Dim dt As Single: dt = Timer
    dt = Timer - dt
    BtnClone.Enabled = Not m_Bmp Is Nothing
    If m_Bmp Is Nothing Then Exit Sub
    Set PBBitmap.Picture = m_Bmp.ToPicture
    'Label1.Caption = "File loading time t: " & dt & "sec;"
    UpdateFormCaption
    Text1.Text = m_Bmp.ToStr
    mnuEditPalette.Enabled = m_Bmp.IsIndexed
    BtnPickAColor.Enabled = True
    'BtnClone.Enabled = True
    
End Sub
