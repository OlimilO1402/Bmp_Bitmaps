Attribute VB_Name = "MMain"
Option Explicit
Private m_Filter As String

Sub Main()
    SetFileFilters
    WriteEZTW32IfNotFound
    FMain.Show
End Sub

Sub WriteEZTW32IfNotFound()
    Dim aPFN As String: aPFN = App.Path & "\eztw32.dll"
    Dim s As String
    If Dir(aPFN) = "" Then
        s = "File not found, trying to write the file: " & vbCrLf & aPFN
        MsgBox s
        Dim bin() As Byte: bin = LoadResData(101, "CUSTOM")
        s = IIf(WriteFile(bin, aPFN), "Successfully written " & UBound(bin) + 1 & " bytes to file: ", _
                                      "Could not write the file: ")
        MsgBox s & vbCrLf & aPFN
        'If Not WriteFile(bin, aPFN) Then
        '    MsgBox "Could not write the file: " & vbCrLf & aPFN
        'End If
    End If
End Sub

Function WriteFile(bytes() As Byte, PFN As String) As Boolean
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , bytes
    WriteFile = True
    GoTo Finally
Catch:
    MsgBox "Error writing the file: " & vbCrLf & PFN
Finally:
    Close FNr
End Function

Public Function GetOpenFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New OpenFileDialog
    GetOpenFileName = GetFileName(Owner, FD, PFN)
End Function
Public Function GetSaveFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New SaveFileDialog
    GetSaveFileName = GetFileName(Owner, FD, PFN)
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

Sub SetFileFilters()
'    m_Filter = "All files [*.*]|*.*|" & _
'               "Windows & OS/2 Bitmap [*.bmp]|*.bmp|" & _
'               "Portable Network Graphic [*.png]|*.png|" & _
'               "JPEG-JFIF Compliant [*.jpg;*.jif;*.jpeg]|*.jpg;*.jif;*.jpeg|" & _
'               "Graphics Interchange Format [*.gif]|*.gif" '""
    Dim i As Long
    ReDim sa(0 To 60) As String
    
    AddFlt sa, "All Files", "*", i
    'AddFlt sa, "Amiga", "iff", i
    'AddFlt sa, "Autodesk Drawing Interchange", "dxf", i
    AddFlt sa, "CompuServe Graphics Interchange", "gif", i
    'AddFlt sa, "Computer Graphics Metafile", "cgm", i
    'AddFlt sa, "Corel Clipart", "cmx", i
    'AddFlt sa, "CorelDraw Drawing", "cdr", i
    'AddFlt sa, "Deluxe Paint", "lbm", i
    'AddFlt sa, "Dr. Halo", "cut", i
    'AddFlt sa, "GEM Paint", "img", i
    'AddFlt sa, "HP Graphics Language", "hgl", i
    AddFlt sa, "JPEG-JFIF Compliant", Array("jpg", "jif", "jpeg"), i
    'AddFlt sa, "Kodak Digital Camera File", "kdc", i
    'AddFlt sa, "Kodak FlashPix", "fpx", i
    'AddFlt sa, "Kodak Photo CD", "pcd", i
    'AddFlt sa, "Lotus PIC", "pic", i
    'AddFlt sa, "Macintosh PICT", "pct", i
    'AddFlt sa, "MacPaint", "mac", i
    'AddFlt sa, "Micrografx Draw", "drw", i
    'AddFlt sa, "Microsoft Paint", "msp", i
    'AddFlt sa, "Paint Shop Pro Image", "psp", i
    'AddFlt sa, "PC Paint", "pic", i
    'AddFlt sa, "Photoshop", "psd", i
    'AddFlt sa, "Portable Bitmap", "pbm", i
    'AddFlt sa, "Portable Greymap", "pgm", i
    AddFlt sa, "Portable Network Graphics", "png", i
    'AddFlt sa, "Portable Pixelmap", "ppm", i
    'AddFlt sa, "Raw File Format", "raw", i
    'AddFlt sa, "SciTex Continuous Tone", Array("sct, ct"), i
    'AddFlt sa, "Sun Raster Image", "ras", i
    'AddFlt sa, "Tagged Image File Format", Array("tif", "tiff"), i
    'AddFlt sa, "Truevision Targa", "tga", i
    'AddFlt sa, "Ventura/GEM Drawing", "gem", i
    'AddFlt sa, "Windows Clipboard", "clp", i
    'AddFlt sa, "Windows Enhanced Meta File", "emf", i
    'AddFlt sa, "Windows Meta File", "wmf", i
    'AddFlt sa, "Windows or Compuserve RLE", "rle", i
    AddFlt sa, "Windows or OS/2 Bitmap", "bmp", i
    'AddFlt sa, "Windows of OS/2 DIB", "dib", i
    'AddFlt sa, "WordPerfect Bitmap", "wpg", i
    'AddFlt sa, "WordPerfect Vector", "wpg", i
    'AddFlt sa, "Zsoft Multipage Paintbrush", "dcx", i
    'AddFlt sa, "Zsoft Paintbrush", "pcx", i
    m_Filter = Join(sa, "")
End Sub

Sub AddFlt(sa() As String, ByVal Filtername As String, Extensions, ByRef i_inout As Long)
    Dim e As String: e = GetExt(Extensions)
    sa(i_inout) = Filtername & " [" & e & "]|" & e & "|"
    i_inout = i_inout + 1
End Sub
Function GetExt(Extensions) As String
    Dim s As String, vt As VbVarType: vt = VarType(Extensions)
    Dim vbArrVar As Long: vbArrVar = VbVarType.vbArray Or VbVarType.vbVariant 'vbString
    Select Case vt
    Case VbVarType.vbString: s = "*." & Extensions
    Case vbArrVar
        Dim i As Long, u As Long: u = UBound(Extensions)
        For i = LBound(Extensions) To u
            s = s & "*." & Extensions(i) & IIf(i < u, ";", "")
        Next
    End Select
    GetExt = s
End Function

'Paint Shop Pro
'File-Open: File Type dropdown-box
' All Files                       (*.*)
' Amiga                           (*.iff)
' Autodeks Drawing Interchange    (*.dxf)
' CompuServe Graphics Interchange (*.gif)
' Computer Graphics Metafile      (*.cgm)
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
