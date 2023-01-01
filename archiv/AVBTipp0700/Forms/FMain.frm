VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FMain 
   Caption         =   "PixelFormat"
   ClientHeight    =   6015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBConv 
      Height          =   5535
      Left            =   6000
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   5475
      ScaleWidth      =   5955
      TabIndex        =   5
      Top             =   360
      Width           =   6015
   End
   Begin VB.PictureBox PBOrig 
      Height          =   5535
      Left            =   0
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   5475
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   360
      Width           =   6015
   End
   Begin VB.ComboBox CmbBPP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "FMain.frx":0000
      Left            =   8880
      List            =   "FMain.frx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton BtnConvert 
      Caption         =   "Convert"
      Height          =   375
      Left            =   10920
      TabIndex        =   0
      Top             =   0
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PixelFormat:"
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
      Left            =   7680
      TabIndex        =   8
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Converted Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   7
      Top             =   60
      Width           =   1515
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selected Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   1380
   End
   Begin VB.Label LblPixelFormatOrig 
      AutoSize        =   -1  'True
      Caption         =   "-----------------"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PixelFormat:"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   60
      Width           =   1080
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFNOrig As String
Private m_PicOrig As StdPicture
Private m_PicConv As StdPicture
Private m_PFNConv As String
Private m_PFNCurr As String

Private Sub Form_Load()
    MBitmap.BPP_ToCBLB CmbBPP
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = PBOrig.Top
    Dim W As Single: W = Me.ScaleWidth / 2
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then PBOrig.Move L, T, W, H
    L = W
    If W > 0 And H > 0 Then PBConv.Move L, T, W, H
    Label3.Left = L
    Label4.Left = Label3.Left + 1680
    CmbBPP.Left = Label4.Left + 1200
    BtnConvert.Left = CmbBPP.Left + CmbBPP.Width
End Sub

Private Sub mnuFileOpen_Click()
Try: On Error GoTo Catch
    Set PBOrig.Picture = Nothing
    Set PBConv.Picture = Nothing
    ' set dialog parameters
    With CommonDialog1
        .Filter = "All Picture Files (*.BMP;*.DIB;*.JPG;*.GIF;*.EMF;*.WMF;*.ICO;*.CUR)|*.BMP;*.DIB;*.JPG;*.GIF;*.EMF;*.WMF;*.ICO;*.CUR"
        .CancelError = True
        .ShowOpen
    End With
    LoadPictureFile CommonDialog1.FileName
    Exit Sub
Catch:
    If Err.Number = 32755 Then Exit Sub 'Cancel
    MsgBox Err.Number & " " & Err.Description
End Sub

Private Sub LoadPictureFile(PFN As String)
Try: On Error GoTo Catch
    m_PFNOrig = PFN
    m_PFNCurr = m_PFNOrig
    Set m_PicOrig = LoadPicture(m_PFNOrig)
    Set m_PicConv = m_PicOrig
    UpdateView
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub

Private Sub UpdateView()
    Set PBOrig.Picture = m_PicOrig
    Set PBConv.Picture = m_PicConv
    Me.Caption = "PixelFormat " & m_PFNCurr
End Sub

Private Function GetSelectedConvertToPixelFormat() As BPP
    Dim i As Long: i = CmbBPP.ListIndex
    GetSelectedConvertToPixelFormat = i
End Function

Private Sub BtnConvert_Click()
    Dim b As BPP
    b = GetSelectedConvertToPixelFormat
    Set m_PicConv = ConvertBitmapAllRes(m_PicOrig, b)
    UpdateView
End Sub

Private Sub mnuFileSave_Click()
Try: On Error GoTo Catch
    ' Dialogparameter setzen
    With CommonDialog1
        .Filter = "Bitmap Files (*.BMP)|*.BMP"
        .FileName = "*.bmp"
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
    End With
    Dim pf As BPP: pf = GetSelectedConvertToPixelFormat
    ' Bild konvertieren und speichern
    m_PFNConv = CommonDialog1.FileName
    m_PFNCurr = m_PFNConv
    Dim s  As String:         s = "Das Speichern der Bitmap war "
    Dim ms As VbMsgBoxStyle: ms = vbOKOnly Or vbInformation
    If Not SaveBitmapAllRes(PBConv.Picture, m_PFNConv, pf) Then
        s = s & "nicht ":    ms = ms Or vbCritical
    End If
    s = s & "erfolgreich"
    MsgBox s, ms
    Exit Sub
Catch:
    MsgBox Err.Description
End Sub

Private Sub PBOrig_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub PBConv_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    OLEDragDrop Data, Effect, Button, Shift, X, Y
End Sub

Private Sub OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(vbCFFiles) Then Exit Sub
    Dim PFN As String: PFN = Data.Files(1)
    LoadPictureFile PFN
End Sub
