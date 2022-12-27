VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows-Standard
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   4530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picOrg 
      Height          =   3825
      Left            =   60
      ScaleHeight     =   3765
      ScaleWidth      =   4695
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   30
      Width           =   4755
   End
   Begin VB.PictureBox picConv 
      Height          =   3825
      Left            =   4890
      ScaleHeight     =   3765
      ScaleWidth      =   4695
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   30
      Width           =   4755
   End
   Begin VB.CommandButton cmdLoadPicture 
      Caption         =   "&Load Picture"
      Height          =   495
      Left            =   60
      TabIndex        =   8
      Top             =   3990
      Width           =   3345
   End
   Begin VB.CommandButton cmdSavePicture 
      Caption         =   "&Save Picture"
      Height          =   495
      Left            =   6240
      TabIndex        =   7
      Top             =   3990
      Width           =   3405
   End
   Begin VB.Frame frPixelFormat 
      Caption         =   "Convert to"
      Height          =   2925
      Left            =   3480
      TabIndex        =   0
      Top             =   3900
      Width           =   2655
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat32bppRGB"
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   12
         Top             =   2550
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat24bppRGB"
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   11
         Top             =   2220
         Value           =   -1  'True
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat1bppIndexed"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat4bppIndexed"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   570
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat4bppIndexed RLE"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Top             =   900
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat8bppIndexed"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   3
         Top             =   1230
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat8bppIndexed RLE"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   2
         Top             =   1560
         Width           =   2445
      End
      Begin VB.OptionButton obPixelFormat 
         Caption         =   "PixelFormat16bppRGB"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   1
         Top             =   1890
         Width           =   2445
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Dieser Source stammt von http://www.activevb.de
' und kann frei verwendet werden. Für eventuelle Schäden
' wird nicht gehaftet.
'
' Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
' Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

Private LastPixelFormat As BPP

Private Sub cmdLoadPicture_Click()

    ' Fehlerbehandlung
    On Error GoTo errorhandler
    
    ' Dialogparameter setzen
    With CommonDialog1
    
        .Filter = "All Picture Files (*.BMP;*.DIB;*.JPG;*.GIF;*.EMF;*." & _
            "WMF;*.ICO;*.CUR)|*.BMP;*.DIB;*.JPG;*.GIF;*.EMF;*.WMF;*.IC" & _
            "O;*.CUR"
            
        .CancelError = True
        
        .ShowOpen
        
    End With
    
    ' Frame und Button aktivieren
    frPixelFormat.Enabled = True
    cmdSavePicture.Enabled = True
    
    ' Bild laden
    picOrg.Picture = LoadPicture(CommonDialog1.FileName)
    
    ' Bild konvertieren
    picConv.Picture = ConvertBitmapAllRes(picOrg.Picture, LastPixelFormat)
    
    Exit Sub
    
errorhandler:

End Sub

Private Sub cmdSavePicture_Click()

    ' Fehlerbehandlung
    On Error GoTo errorhandler
    
    ' Dialogparameter setzen
    With CommonDialog1
        
        .Filter = "Bitmap Files (*.BMP|*.BMP"
        .FileName = "*.bmp"
        .CancelError = True
        .ShowSave
        .Flags = cdlOFNOverwritePrompt
        
    End With
    
    ' Bild konvertieren und speichern
    If SaveBitmapAllRes(picConv.Picture, CommonDialog1.FileName, _
        LastPixelFormat) Then
        
        MsgBox "Das speichern der Bitmap war erfolgreich.", vbOKOnly Or _
            vbInformation
            
    Else
    
        MsgBox "Das speichern der Bitmap war nicht erfolgreich.", _
            vbOKOnly Or vbCritical
            
    End If
    
    Exit Sub
    
errorhandler:

End Sub

Private Sub Form_Load()

    LastPixelFormat = PixelFormat24bppRGB
    cmdSavePicture.Enabled = False
    frPixelFormat.Enabled = False
    
End Sub

Private Sub obPixelFormat_Click(Index As Integer)

    Select Case Index
    
    Case 0
        LastPixelFormat = PixelFormat1bppIndexed
        
    Case 1
        LastPixelFormat = PixelFormat4bppIndexed
        
    Case 2
        LastPixelFormat = PixelFormat4bppIndexed_RLE
        
    Case 3
        LastPixelFormat = PixelFormat8bppIndexed
        
    Case 4
        LastPixelFormat = PixelFormat8bppIndexed_RLE
        
    Case 5
        LastPixelFormat = PixelFormat16bppRGB
        
    Case 6
        LastPixelFormat = PixelFormat24bppRGB
        
    Case 7
        LastPixelFormat = PixelFormat32bppRGB
        
    End Select
    
    picConv.Picture = ConvertBitmapAllRes(picOrg.Picture, LastPixelFormat)
    
End Sub
