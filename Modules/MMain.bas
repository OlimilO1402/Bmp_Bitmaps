Attribute VB_Name = "MMain"
Option Explicit
Private m_Filter As String

Sub Main()
    m_Filter = "All files (*.*)|*.*|Windows & OS/2 Bitmap [*.bmp]|*.bmp|Portable Network Graphic [*.png]|*.png|JPEG-JFIF Compliant [*.jpg;*.jif;*.jpeg]|*.jpg;*.jif;*.jpeg|Graphics Interchange Format [*.gif]|*.gif" '""
    FMain.Show
End Sub
'
Public Function GetOpenFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New OpenFileDialog
    GetOpenFileName = GetFileName(Owner, FD, PFN)
'    With FD
'        .Filter = m_Filter
'        '.InitialDirectory
'        If .ShowDialog(FMain) = vbCancel Then Exit Function
'        GetOpenFileName = .FileName
'    End With
End Function
'
Public Function GetSaveFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New SaveFileDialog
    GetSaveFileName = GetFileName(Owner, FD, PFN)
'    GetSaveFileName
'    With FD
'        .Filter = m_Filter
'        '.InitialDirectory
'        If .ShowDialog(FMain) = vbCancel Then Exit Function
'        GetSaveFileName = .FileName
'    End With
End Function

Public Function GetFileName(Owner As Form, FD, Optional ByVal PFN As String) As String
    If FD Is Nothing Then Exit Function
    With FD
        .Filter = m_Filter
        .FileName = PFN
        '.InitialDirectory
        If .ShowDialog(Owner) = vbCancel Then Exit Function
        GetFileName = .FileName
    End With
End Function

' All Files                       (*.*)
' Amiga                           (*.iff)
' Autodeks Drawing Interchange    (*.dxf)
' CompuServe Graphics Interchange (*.gif)
' computer Graphics Metafile      (*.cgm)
' Corel Clipart                   (*.cmx)
' CorelDraw Drawing               (*.cdr)
' Deluxe Paint                    (*.lbm)
' Dr. Halo                        (*.cut)
' GEM Paint                       (*.img)
' HP Graphics Language            (*.hgl)
' JPEG - JFIF Compliant           (*.jpg; *.jif; *.jpeg)
' Kodak Digital Camera File       (*.kdc)
' Kodak FlashPix                  (*.fpx)
' Kodak Photo CD                  (*.pcd)
' Lotus PIC                       (*.pic)
' Macintosh PICT                  (*.pct)
' MacPaint                        (*.mac)
' Micrografx Draw                 (*.drw)
' Microsoft Paint                 (*.msp)
' Paint Shop Pro Image            (*.psp)
' PC Paint                        (*.pic)
' Photoshop                       (*.psd)
' Portable Bitmap                 (*.pbm)
' Portable Greymap                (*.pgm)
' Portable Network Graphics       (*.png)
' Portable Pixelmap               (*.ppm)
' Raw File Format                 (*.raw; *.*)
' SciTex Continuous Tone          (*.sct; *.ct)
' Sun Raster Image                (*.ras)
' Tagged Image File Format        (*.tif; *.tiff)
' Truevision Targa                (*.tga)
' Ventura/GEM Drawing             (*.gem)
' Windows Clipboard               (*.clp)
' Windows Enhanced Meta File      (*.emf)
' Windows Meta File               (*.wmf)
' Windows or Compuserve RLE       (*.rle)
' Windows or OS/2 Bitmap          (*.bmp)
' Windows of OS/2 DIB             (*.dib)
' WordPerfect Bitmap              (*.wpg)
' WordPerfect Vector              (*.wpg)
' Zsoft Multipage Paintbrush      (*.dcx)
' Zsoft Paintbrush                (*.pcx)
