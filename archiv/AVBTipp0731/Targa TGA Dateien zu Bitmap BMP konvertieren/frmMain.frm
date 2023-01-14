VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Tga2Bmp"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save as Bitmap"
      Height          =   435
      Left            =   2010
      TabIndex        =   2
      Top             =   60
      Width           =   1905
   End
   Begin MSComDlg.CommonDialog cdOpen 
      Left            =   6000
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Height          =   4965
      Left            =   60
      ScaleHeight     =   4905
      ScaleWidth      =   6375
      TabIndex        =   1
      Top             =   570
      Width           =   6435
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open Targa"
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1905
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOpen_Click()

    ' Fehlerbehandlung
    On Error GoTo errorhandler
    
    ' div. Parameter für den Dialog
    With cdOpen
    
        .DialogTitle = "Load Targafile"
        .Filter = "Targa Files *.tga | *.tga"
        .InitDir = App.Path
        .CancelError = True
        .ShowOpen
        
    End With
    
    ' TGA laden, konvertieren und anzeigen
    Picture1.Picture = ConvertTga2Bmp(cdOpen.FileName)
    
    ' ist ein Bild in der PictureBox vorhanden
    If Not (Picture1.Picture Is Nothing) Then
    
        ' Button zum speichern aktivieren
        cmdSave.Enabled = True
        
    End If
    
    Exit Sub
    
errorhandler:

End Sub

Private Sub cmdSave_Click()

    ' Fehlerbehandlung
    On Error GoTo errorhandler
    
    ' div. Parameter für den Dialog
    With cdOpen
    
        .DialogTitle = "Save as Bitmap"
        .Filter = "Bitmp Files *.bmp | *.bmp"
        .FileName = "ConvTarga"
        .DefaultExt = "bmp"
        .InitDir = App.Path
        .CancelError = True
        .Flags = cdlOFNOverwritePrompt
        .ShowSave
        
    End With
    
    ' Bild als Bitmap speichern
    Call SavePicture(Picture1.Picture, cdOpen.FileName)
    
    Exit Sub
    
errorhandler:

End Sub

Private Sub Form_Load()

    ' Button zum speichern deaktivieren
    cmdSave.Enabled = False
    
End Sub
