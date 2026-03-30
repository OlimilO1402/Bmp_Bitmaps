VERSION 5.00
Begin VB.Form frmDIB_Test 
   Caption         =   "frmDIB_Test"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   ScaleHeight     =   621
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   619
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame4 
      Height          =   9225
      Left            =   7740
      TabIndex        =   12
      Top             =   0
      Width           =   1470
      Begin VB.CommandButton cmdMirrow 
         Caption         =   "Mirrow"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   20
         Top             =   2085
         Width           =   1200
      End
      Begin VB.CommandButton cmdPaste 
         Caption         =   "Copy / Paste"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   19
         Top             =   1245
         Width           =   1200
      End
      Begin VB.CommandButton cmdInvert 
         Caption         =   "Invert"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   18
         Top             =   1665
         Width           =   1200
      End
      Begin VB.CommandButton cmdCopyPicII 
         Caption         =   "Copy Pic II"
         Height          =   390
         Left            =   135
         TabIndex        =   17
         Top             =   675
         Width           =   1200
      End
      Begin VB.CommandButton cmdZoom 
         Caption         =   "Zoom"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   16
         Top             =   2505
         Width           =   1200
      End
      Begin VB.CommandButton cmdAnimate 
         Caption         =   "Animate"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   15
         Top             =   3060
         Width           =   1200
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   390
         Left            =   135
         TabIndex        =   14
         Top             =   3480
         Width           =   1200
      End
      Begin VB.CommandButton cmdCreatePicI 
         Caption         =   "Create Pic I"
         Height          =   390
         Left            =   135
         TabIndex        =   13
         Top             =   255
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4020
      Left            =   90
      TabIndex        =   10
      Top             =   0
      Width           =   7560
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         Height          =   3675
         Left            =   120
         Picture         =   "frmDIB_Test.frx":0000
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   483
         TabIndex        =   11
         Top             =   210
         Width           =   7305
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4020
      Left            =   90
      TabIndex        =   8
      Top             =   5205
      Width           =   7560
      Begin VB.PictureBox picScreen 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   165
         ScaleHeight     =   241
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   482
         TabIndex        =   9
         Top             =   255
         Width           =   7230
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1185
      Left            =   90
      TabIndex        =   0
      Top             =   4020
      Width           =   7560
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   390
         Left            =   6315
         TabIndex        =   4
         Top             =   225
         Width           =   1110
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   180
         Index           =   2
         LargeChange     =   20
         Left            =   360
         Max             =   255
         TabIndex        =   3
         Top             =   255
         Value           =   127
         Width           =   5550
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   180
         Index           =   1
         LargeChange     =   20
         Left            =   360
         Max             =   255
         TabIndex        =   2
         Top             =   540
         Value           =   127
         Width           =   5550
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   180
         Index           =   0
         LargeChange     =   20
         Left            =   360
         Max             =   255
         TabIndex        =   1
         Top             =   825
         Value           =   127
         Width           =   5550
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fest Einfach
         Height          =   180
         Left            =   135
         TabIndex        =   7
         Top             =   825
         Width           =   180
      End
      Begin VB.Label Label2 
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fest Einfach
         Height          =   180
         Left            =   135
         TabIndex        =   6
         Top             =   540
         Width           =   180
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fest Einfach
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   255
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmDIB_Test"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents DIB As clsDIB
Attribute DIB.VB_VarHelpID = -1

Private bArray1() As Byte
Private bArray2() As Byte

Private NoScroll As Boolean
Private Animating As Boolean

Private Sub Form_Load()
    On Error GoTo ErrLoad
    
        Set DIB = New clsDIB
    
        DIB.hWndSource = Picture2.hWnd
        DIB.hWndDestination = picScreen.hWnd
        
        Debug.Print 1 / 0
        
        Exit Sub

ErrLoad:
        Call MsgBox("Best viewn as exe-file!", vbInformation, "DIB")
        Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Animating = False
    DoEvents
        
    DIB.Release
    Set DIB = Nothing
End Sub

Private Sub cmdCreatePicI_Click()
    Call CreatePicI
    DIB.BitBlit bArray2
    Call Reset
    Call cLock(True)
End Sub

Private Sub cmdCopyPicII_Click()
    Call CopyPicII
    Call cLock(True)
End Sub

Private Sub cLock(Flag As Boolean)
    Frame1.Enabled = Flag
    cmdAnimate.Enabled = Flag
    cmdPaste.Enabled = Flag
    cmdInvert.Enabled = Flag
    cmdMirrow.Enabled = Flag
    cmdZoom.Enabled = Flag
End Sub

Private Sub cmdPaste_Click()
    DIB.BitBlit bArray1, 340, 100, 120, 70, 20, 50
End Sub

Private Sub cmdInvert_Click()
    DIB.StretchBitBlit bArray2, , , , , , , , , vbNotSrcCopy
End Sub

Private Sub cmdMirrow_Click()
    Dim h As Long, w As Long
    Static z As Long
    Dim p1 As Long, p2 As Long, p3 As Long, p4 As Long
    
        w = picScreen.ScaleWidth
        h = picScreen.ScaleHeight
        
        z = z + 1
        
        Select Case z
            Case 1: p1 = 0: p2 = h: p3 = 0: p4 = -h
            Case 2: p1 = w: p2 = h: p3 = -w: p4 = -h
            Case 3: p1 = w: p2 = 0: p3 = -w: p4 = 0
            Case 4: p1 = 0: p2 = 0: p3 = w: p4 = h: z = 0
        End Select
        
        DIB.StretchBitBlit bArray2, 0, 0, w, h, p1, p2, p3, p4, vbSrcCopy
End Sub

Private Sub cmdZoom_Click()
    Const Limit As Long = 16
    Static z As Long, zf As Double, x2 As Long, y2 As Long, x1 As Long, y1 As Long
    Static s As Long, w As Long, h As Long, wf As Long, hf As Long
    Dim scH As Long, scW As Long, x As Long, y As Long, xf  As Long, yf As Long
    
        If z = 0 Then
            z = 1
            s = 1
        ElseIf z = Limit Then
            s = -1
        ElseIf z < 1 Then
            s = 1
        End If

        z = z + s
        If z = 0 Then z = z + s

        If z > 0 Then
           zf = z
        ElseIf z < 0 Then
            zf = 1 / Abs(z)
        End If
        
        scW = picScreen.ScaleWidth
        scH = picScreen.ScaleHeight
        
        x1 = (scW / 2) - (scW / 2) / zf
        x2 = (scW / 2) + (scW / 2) / zf
        
        x = x1
        If x < 0 Then x = 0
        
        w = (x2 - x1)
        If w > scW Then w = scW
        
        wf = Int(w * zf)
        If wf > scW Then wf = scW
        
        xf = Int((scW - wf) / 2)
        
        
        y1 = (scH / 2) - (scH / 2) / zf
        y2 = (scH / 2) + (scH / 2) / zf
        
        y = y1
        If y < 0 Then y = 0
        
        h = (y2 - y1)
        If h > scH Then h = scH
        
        hf = Int(h * zf)
        If hf > scH Then hf = scH
        
        yf = Int((scH - hf) / 2)
        
        DIB.StretchBitBlit bArray2, x - 1, y - 1, w + 2, h + 2, xf, yf, wf, hf, vbSrcCopy
End Sub

Private Sub cmdAnimate_Click()
    If DIB.IsCreated Then
        Call cLock(False)
        Call Reset
        bArray2 = bArray1
        DIB.BitBlit bArray2
        
        cmdAnimate.Enabled = False
        cmdCreatePicI.Enabled = False
        cmdCopyPicII.Enabled = False
    
        cmdStop.Enabled = True
        Animating = True
        Call Animate
    End If
End Sub

Private Sub cmdStop_Click()
    Animating = False
    cmdAnimate.Enabled = True
    cmdStop.Enabled = False
    Call cLock(True)
End Sub

Private Sub cmdReset_Click()
    Call Reset
    bArray2 = bArray1
    DIB.BitBlit bArray2
End Sub

Private Sub DIB_Error(ErrorNumber As Long, ErrorText As String)
    Call MsgBox(ErrorText, vbCritical, "DIB")
End Sub

Private Sub picScreen_Paint()
   If DIB.IsCreated Then DIB.BitBlit bArray2
End Sub

Private Sub HScroll1_Change(Index As Integer)
     Call ColorPic(Index)
End Sub

Private Sub HScroll1_GotFocus(Index As Integer)
    picScreen.SetFocus
End Sub

Private Sub HScroll1_Scroll(Index As Integer)
    Call ColorPic(Index)
End Sub

Private Sub CreatePicI()
    If DIB.IsCreated Then DIB.Release

    DIB.Width = 480
    DIB.Height = 240
    picScreen.Width = DIB.Width * Screen.TwipsPerPixelX
    picScreen.Height = DIB.Height * Screen.TwipsPerPixelY
    
    If DIB.SizeArray(bArray1, dibNew) Then
        DIB.SizeArray bArray2, dibNew
        Call FillArray
        DIB.Create
    End If
End Sub

Private Sub CopyPicII()
    If DIB.IsCreated Then DIB.Release
    
    DIB.Width = Picture2.ScaleWidth
    DIB.Height = Picture2.ScaleHeight

    DIB.SetRedraw = False
    picScreen.Width = DIB.Width * Screen.TwipsPerPixelX
    picScreen.Height = DIB.Height * Screen.TwipsPerPixelY
    picScreen.Refresh
    DIB.SetRedraw = True
    
    If DIB.SizeArray(bArray1, dibNew) And DIB.SizeArray(bArray2, dibNew) Then
        DIB.Create
        DIB.CopyBitmapToArray bArray1, , , , , , , vbSrcCopy
    
        bArray2 = bArray1
        DIB.BitBlit bArray2
    End If
    
    Call Reset
End Sub

Private Sub FillArray()
    Dim x As Long, y As Long
    Dim pR As Double, pB As Double
    Dim uX As Long, uY As Long

        uX = DIB.SizeX
        uY = DIB.SizeY
        
        pR = 255 / uY
        pB = 255 / uX
        
        For y = 0 To uY
            For x = 0 To uX Step 3
                bArray1(x + 0, y) = y * pR
                bArray1(x + 1, y) = 255 - x * pB
                bArray1(x + 2, y) = x * pB
            Next x
        Next y
        
        bArray2 = bArray1
End Sub

Private Sub ColorPic(Index As Integer)
    Dim bA() As Byte, x As Long, y As Long, b As Long, v As Long
        
        If NoScroll Then Exit Sub
        v = HScroll1(Index).Value - 127
        
        For y = 0 To DIB.SizeY
            For x = 0 To DIB.SizeX - 2 Step 3
                b = bArray1(x + Index, y) + v
                If b > 255 Then
                    b = 255
                ElseIf b < 0 Then
                    b = 0
                End If
            
                bArray2(x + Index, y) = b
            Next x
        Next y
        
        DIB.BitBlit bArray2
End Sub

Private Sub Reset()
    Dim z As Long
    
        NoScroll = True
        For z = 0 To 2
            HScroll1(z).Value = 127
        Next z
        NoScroll = False
End Sub

Private Sub Animate()
    Static x As Long, b As Long, y As Long, zx As Long
    Dim x1 As Long, x2 As Long, c As Long
    Dim z As Long, w3 As Long
    Const w As Long = 120
    
        Animating = True
        w3 = w / 3
        
        Do
            x = x + 12
            If x - w > DIB.SizeX Then x = -w
            
            bArray2 = bArray1
            
            For y = 0 To DIB.SizeY
                
                b = 0
                For zx = x To x + w Step 3
                    
                    If zx > 0 And zx <= DIB.SizeX Then
                        c = bArray1(zx, y) + w - b
                        If c > 255 Then c = 255
                        bArray2(zx, y) = c
                    End If
                    
                    z = x - b
                    If z > 0 And z < DIB.SizeX Then
                        c = bArray1(z, y) + w - b
                        If c > 255 Then c = 255
                        bArray2(z, y) = c
                    End If
                    
                    b = b + 3
                Next zx
    
            Next y
            
            DIB.RuntimeMode = dibSpeed
            DIB.BitBlit bArray2, (x - w) / 3, 0, w * 2 / 3, DIB.Height, (x - w) / 3, 0
            DIB.RuntimeMode = dibSecure
        
            If DIB.MessagesQueued(qsKey Or qsMouse Or qsPaint, qsAll) <> 0 Then DoEvents
          
        Loop While Animating
End Sub
