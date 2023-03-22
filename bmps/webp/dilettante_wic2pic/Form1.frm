VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   ScaleHeight     =   444
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   979
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10440
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   6960
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   1
      Top             =   600
      Width           =   6735
   End
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   120
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Picture1.BorderStyle = 0
    Picture1.AutoSize = True
    Picture2.BorderStyle = 0
    Picture2.AutoSize = True
End Sub

Private Sub Command1_Click()
    With New Wic2Pic
        Set Picture1.Picture = .LoadFile("sample.webp", hDC)
    End With
End Sub
Private Sub Command2_Click()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static counter As Long
    With New cWICImage
        If .OpenFile("sample.webp", counter) Then
            .Render Picture2.hDC, 0, 0, Picture1.Width, Picture1.Height
        'Else
        '    MsgBox "could not load file"
        End If
        counter = counter + 1
        If counter = .FrameCount Then counter = 0
        Debug.Print counter & " v " & .FrameCount
    End With
End Sub
