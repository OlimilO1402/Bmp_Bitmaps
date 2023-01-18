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
   Begin VB.CommandButton BtnClone 
      Caption         =   "Clone >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      TabIndex        =   8
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton BtnPalette 
      Caption         =   "Palette"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      TabIndex        =   6
      Top             =   0
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10080
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton BtnPickAColor 
      Caption         =   "Pick a Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   0
      Width           =   1335
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
      TabIndex        =   3
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   390
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      BorderStyle     =   0  'Kein
      Height          =   6735
      Left            =   4080
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   1
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   390
      Width           =   7695
   End
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open bmps-subfolder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "        "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   45
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Drag'n'drop pictures of filetype *.bmp onto the window."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   45
      Width           =   4680
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
    BtnPalette.Enabled = False
    BtnPickAColor.Enabled = False
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
    Dim W As Single: W = Text1.Width - L
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
    L = W:    W = Me.ScaleWidth - W
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub

Private Sub BtnOpenFolder_Click()
    Dim p As String: p = App.Path & "\bmps\"
    If MsgBox("Open folder?" & vbCrLf & p, vbOKCancel) = vbCancel Then Exit Sub
    Shell "Explorer.exe " & p, vbNormalFocus
End Sub

Private Sub BtnPickAColor_Click()
    m_bPickAColor = True
End Sub

Private Sub BtnPalette_Click()
    FPalette.Move Me.Left + Me.Width / 2 - FPalette.Width / 2, Me.Top + Me.Height / 2 - FPalette.Height / 2
    If FPalette.ShowDialog(Me, m_Bmp) = vbCancel Then Exit Sub
    UpdateView
End Sub

Private Sub mnuEditResize_Click()
    MiddlePosDlg FDlgNewPicture
    Dim Bmp As Bitmap
    If Not m_Bmp Is Nothing Then Set Bmp = m_Bmp.Clone
    If FDlgNewPicture.ShowDialog(Me, Bmp) = vbCancel Then Exit Sub
    Set m_Bmp = Bmp
    UpdateView
End Sub

Private Sub mnuFileNew_Click()
    MiddlePosDlg FDlgNewPicture
    Dim Bmp As Bitmap
    'If Not m_Bmp Is Nothing Then Set Bmp = m_Bmp.Clone
    If FDlgNewPicture.ShowDialog(Me, Bmp) = vbCancel Then Exit Sub
    Set m_Bmp = Bmp
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
    Dim OFD As OpenFileDialog: Set OFD = New OpenFileDialog
    OFD.Filter = "Bitmaps (*.bmp)|*.bmp|All files (*.*)|*.*"
    If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    Dim PFN As String: PFN = OFD.FileName
    Set m_Bmp = MNew.Bitmap(PFN)
    UpdateView
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_Bmp Is Nothing Then Exit Sub
    If m_bPickAColor Then
        Picture2.BackColor = m_Bmp.Pixel(X, Y)
        Label2.Caption = "X: " & X & "; Y: " & Y
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_bPickAColor Then m_bPickAColor = False
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
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
        Set Picture1.Picture = MLoadPng.LoadPictureGDIp(PFN)
    End If
End Sub

Public Sub UpdateView()
    Dim dt As Single: dt = Timer
    dt = Timer - dt
    If m_Bmp Is Nothing Then Exit Sub
    Set Picture1.Picture = m_Bmp.ToPicture
    Label1.Caption = "File loading time t: " & dt & "sec;"
    UpdateFormCaption
    Text1.Text = m_Bmp.ToStr
    BtnPalette.Enabled = m_Bmp.IsIndexed
    BtnPickAColor.Enabled = True
    BtnClone.Enabled = True
End Sub
