VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Bitmaps"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10980
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "FMain"
   ScaleHeight     =   516
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   3  'Windows-Standard
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
      Left            =   0
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   445
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   701
      TabIndex        =   1
      Top             =   600
      Width           =   10575
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
      Caption         =   "Drag'n'drop pictures of filetype bmp into the box."
      Height          =   255
      Left            =   3240
      TabIndex        =   2
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
Private m_bmp As Bitmap

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim t As Single: t = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth - L
    Dim H As Single: H = Me.ScaleHeight - t
    If W > 0 And H > 0 Then Picture1.Move L, t, W, H
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
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim PFN As String: PFN = Data.Files(1)
    Dim dt As Single: dt = Timer
    Set m_bmp = MNew.Bitmap(PFN)
    dt = Timer - dt
    If m_bmp Is Nothing Then Exit Sub
    Set Picture1.Picture = m_bmp.ToPicture
    Label1.Caption = m_bmp.ToStr & " dt: " & dt & "sec;"
    Me.Caption = "Bitmap - " & PFN
End Sub
