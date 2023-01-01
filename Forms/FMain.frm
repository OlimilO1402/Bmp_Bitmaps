VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Bitmaps"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14895
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   993
   StartUpPosition =   3  'Windows-Standard
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
      TabIndex        =   4
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00400040&
      Height          =   6735
      Left            =   4080
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   1
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   600
      Width           =   7695
   End
   Begin VB.CommandButton BtnOpenFolder 
      Caption         =   "Open bmps-subfolder"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      ToolTipText     =   "Drag'n'drop pictures of filetype *.bmp to the window."
      Top             =   240
      Width           =   7215
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As String
Private m_bmp As Bitmap

Private Sub Form_Load()
    Me.Caption = "Bitmaps" & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub FormCaptionAddFilename()
    Me.Caption = "Bitmaps" & " v" & App.Major & "." & App.Minor & "." & App.Revision & " - " & m_PFN
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

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub Picture1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    AllOLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub AllOLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    m_PFN = Data.Files(1)
    UpdateView
End Sub

Private Sub UpdateView()
    Dim dt As Single: dt = Timer
    Set m_bmp = MNew.Bitmap(m_PFN)
    dt = Timer - dt
    If m_bmp Is Nothing Then Exit Sub
    Set Picture1.Picture = m_bmp.ToPicture
    Label1.Caption = "File loading time t: " & dt & "sec;"
    FormCaptionAddFilename
    Text1.Text = m_bmp.ToStr
End Sub
