VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://de.wikipedia.org/wiki/Windows_Bitmap
'https://en.wikipedia.org/wiki/BMP_file_format
'
'

'Private Const BITSPIXEL    As Long = 12 ' ???

Public Enum EPixelFormat
    'was soll der ganze Mist????
    'Undefined = 0&                  '      &H0       0  Das Pixelformat ist nicht definiert oder
    'DontCare = 0&                   '      &H0       0  Es wurde kein Pixelformat angegeben.
    'Indexed = &H10000               '  &H10000   65536  Die Pixeldaten enthalten farbindizierte Werte, d. h., diese sind ein Index für die Farben in der Systemfarbtabelle, im Gegensatz zu einzelnen Farbwerten.
    'Gdi = &H20000                   '  &H20000  131072  Die Pixeldaten enthalten GDI-Farben.
    'Alpha = &H40000                 '  &H40000  262144  Die Pixeldaten enthalten Alphawerte, die nicht zuvor multipliziert wurden.
    'PAlpha = &H80000                '  &H80000  524288  Das Pixelformat enthält zuvor multiplizierte Alphawerte.
    'Extended = &H100000             ' &H100000 1048576  Reserviert.
    'Canonical = &H200000            ' &H200000 2097152  Das Standardpixelformat ist 32 Bits pro Pixel.Das Format gibt 24-Bit Farbtiefe und einen 8-Bit-Alphakanal an.
    
    Format1bppIndexed = &H30101     '  &H30101  196865  Gibt an, dass das Format  1 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält   2 Farben.
    Format4bppIndexed = &H30402     '  &H30402  197634  Gibt an, dass das Format  4 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält  16 Farben.
    Format8bppIndexed = &H30803     '  &H30803  198659  Gibt an, dass das Format  8 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält 256 Farben.
    
    'Graustufen? es gibt kein Format das darauf festgelegt werden kann was soll das?
    Format16bppGrayScale = &H101004 ' &H101004 1052676  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen geben 65.536 Grauschattierungen an.
    Format16bppRgb555 = &H21005     '  &H21005  135173  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot-, Blau- und Grünanteil werden jeweils 5 Bit verwendet. Das verbleibende Bit wird nicht verwendet.
    Format16bppRgb565 = &H21006     '  &H21006  135174  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot- und Blauanteil werden jeweils 5 Bit und für den Grünanteil 6 Bit verwendet.
    Format16bppArgb1555 = &H61007   '  &H61007  397319  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen liefern 32.768 Farbschattierungen, wobei der Rot-, Grün- und Blauanteil jeweils von 5 Bits und die Alphakomponente von 1 Bit wiedergegeben wird.
    Format24bppRgb = &H21808        '  &H21808  137224  Gibt an, dass das Format 24 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Format32bppRgb = &H22009        '  &H22009  139273  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die verbleibendenden 8 Bits werden nicht verwendet.
    Format32bppArgb = &H26200A      ' &H26200A 2498570  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Format32bppPArgb = &HE200B      '  &HE200B  925707  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    Format48bppRgb = &H10300C       ' &H10300C 1060876  Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    Format64bppArgb = &H34400D      ' &H34400D 3424269  Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.
    Format64bppPArgb = &H1C400E     ' &H1C400E 1851406  Gibt an, dass das Format 64 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 16 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    Max = &HF                       '      &HF      15  Der Höchstwert für diese Enumeration. ??????
End Enum

'#If VBA7 = 0 Then
'    Private Enum LongPtr
'        [_]
'    End Enum
'#End If

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
End Type

Private Type TLng
    Value As Long
End Type

Private Type BGRA '
    b As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapfileheader
Private Type BITMAPFILEHEADER
'do not read at once read every single variable separately instead
'is only used to determine type not for the bitmap data
    bfType      As Integer ' 2 ' "BM" = 66, 77 = &H424D = 16973
    bfSize      As Long    ' 4 ' Größe der BMP-Datei in Byte. (unzuverlässig)
    bfReserved1 As Integer ' 2 '
    bfReserved2 As Integer ' 2 '
    bfOffBits   As Long    ' 4 ' Offset to beginning of data
End Type             ' Sum: 14

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
'aka BITMAPV3HEADER aka DIBHEADER the old and first bitmap-header
Private Type BITMAPINFOHEADER
    'biSize          As Long    ' 4
    biWidth         As Long    ' 4
    biHeight        As Long    ' 4
    
    biPlanes        As Integer ' 2
    biBitCount      As Integer ' 2
    
    biCompression   As Long    ' 4 ' the BI_ - constants
    biSizeImage     As Long    ' 4
    biXPelsPerMeter As Long    ' 4
    biYPelsPerMeter As Long    ' 4
    biClrUsed       As Long    ' 4
    biClrImportant  As Long    ' 4
    
End Type                 ' Sum: 40

'BI_-constants
Private Const BI_RGB       As Long = 0& 'An uncompressed format.
Private Const BI_RLE8      As Long = 1& 'A run-length encoded (RLE) format for bitmaps with 8 bpp. The compression format is a 2-byte format consisting of a count byte followed by a byte containing a color index. For more information, see Bitmap Compression.
Private Const BI_RLE4      As Long = 2& 'An RLE format for bitmaps with 4 bpp. The compression format is a 2-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
Private Const BI_BITFIELDS As Long = 3& 'Specifies that the bitmap is not compressed. The members bV4RedMask, bV4GreenMask, and bV4BlueMask specify the red, green, and blue components for each pixel. This is valid when used with 16- and 32-bpp bitmaps
                                        'If biCompression = BI_BITFIELDS -> the format is either RGB-555 or RGB-565.
                                        'Use the subtype GUID in the AM_MEDIA_TYPE structure to determine the specific RGB type.
                                        
                                        'https://learn.microsoft.com/en-us/windows/win32/api/strmif/ns-strmif-am_media_type
                                        
Private Const BI_JPEG      As Long = 4& 'Specifies that the image is compressed using the JPEG file interchange format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
Private Const BI_PNG       As Long = 5& 'Specifies that the image is compressed using the PNG  file interchange format.
Private Const BI_1632      As Long = &H32333631 ' "1632"
Private Const LCS_CALIBRATED_RGB As Long = &H0&

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-ciexyz
Private Type CIEXYZ
    ciexyzX As Long     ' 4
    ciexyzY As Long     ' 4
    ciexyzZ As Long     ' 4
End Type          ' Sum: 12

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-ciexyztriple
Private Type CIEXYZTRIPLE
    ciexyzRed   As CIEXYZ ' 12
    ciexyzGreen As CIEXYZ ' 12
    ciexyBlue   As CIEXYZ ' 12
End Type             ' Sum: 36

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapv4header
Private Type AddBITMAPV4HEADER
    'bV4Size          As Long     ' 4
    'bV4Width         As Long     ' 4
    'bV4Height        As Long     ' 4
    
    'bV4Planes        As Integer  ' 2
    'bV4BitCount      As Integer  ' 2
    
    'bV4Compression   As Long     ' 4
    'bV4SizeImage     As Long     ' 4
    'bV4XPelsPerMeter As Long     ' 4
    'bV4YPelsPerMeter As Long     ' 4
    'bV4ClrUsed       As Long     ' 4
    'bV4ClrImportant  As Long     ' 4
    
    bV4RedMask       As Long     ' 4
    bV4GreenMask     As Long     ' 4
    bV4BlueMask      As Long     ' 4
    bV4AlphaMask     As Long     ' 4
    bV4CSType        As Long     ' 4
    
    bV4Endpoints     As CIEXYZTRIPLE ' 36
    
    bV4GammaRed      As Long     ' 4
    bV4GammaGreen    As Long     ' 4
    bV4GammaBlue     As Long     ' 4
    
End Type                  ' Sum: 108

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapv5header
Private Type AddBITMAPV5HEADER
    'bV5Size          As Long     ' 4
    'bV5Width         As Long     ' 4
    'bV5Height        As Long     ' 4
    
    'bV5Planes        As Integer  ' 2
    'bV5BitCount      As Integer  ' 2
    
    'bV5Compression   As Long     ' 4
    'bV5SizeImage     As Long     ' 4
    'bV5XPelsPerMeter As Long     ' 4
    'bV5YPelsPerMeter As Long     ' 4
    'bV5ClrUsed       As Long     ' 4
    'bV5ClrImportant  As Long     ' 4
    
    'bV5RedMask       As Long     ' 4
    'bV5GreenMask     As Long     ' 4
    'bV5BlueMask      As Long     ' 4
    'bV5AlphaMask     As Long     ' 4
    'bV5CSType        As Long     ' 4
    
    'bV5Endpoints     As CIEXYZTRIPLE
    
    'bV5GammaRed      As Long     ' 4
    'bV5GammaGreen    As Long     ' 4
    'bV5GammaBlue     As Long     ' 4
    
    bV5Intent        As Long     ' 4
    bV5ProfileData   As Long     ' 4
    bV5ProfileSize   As Long     ' 4
    bV5Reserved      As Long     ' 4
End Type

Private Type BITMAPINFOV3
    biSize      As Long
    bmiHeader   As BITMAPINFOHEADER
    bmiColors(0 To 255) As BGRA
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfo
Private Type BITMAPINFO
    biSize      As Long
    bmiHeader   As BITMAPINFOHEADER
    bmiHeaderV4 As AddBITMAPV4HEADER
    bmiHeaderV5 As AddBITMAPV5HEADER
    bmiColors(0 To 255) As BGRA
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/olectl/ns-olectl-pictdesc
Private Type PICTDESC
    cbSize  As Long    ' 4
    picType As Long    ' 4    ' PictureTypeConstants
    hbitmap As LongPtr ' 4, 8 ' HBITMAP
    hpal    As LongPtr ' 4, 8 ' HPALETTE
End Type         ' Sum: 16/24

#If VBA7 Then
    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal pbmi As LongPtr, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal offset As Long) As Long
    
    Private Declare PtrSafe Function SetDIBits Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal hbitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
        ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
#Else
    Private Declare Function CreateDIBSection Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal pbmi As LongPtr, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal Offset As Long) As Long
    
    Private Declare Function SetDIBits Lib "gdi32" ( _
        ByVal hdc As LongPtr, ByVal hbitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
        ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
#End If
Private m_PFN     As String
'Private m_PNMType As EPNMType ' P2: Daten als Text gespeichert, P5: Daten binär gespeichert
Private m_BmpHead As BITMAPFILEHEADER
Private m_BmpInfo As BITMAPINFO
Private m_PicGUID As GUID
Private m_PicDesc As PICTDESC
Private m_data()  As Byte 'the pixel-data

Private Sub Class_Initialize()
    Dim i As Integer
    With m_BmpInfo
        .biSize = 40  'LenB(.bmiHeader) ' nope not LenB anymore ' a default value for V3-header as v4 or v5 is not so often
        .bmiHeader.biPlanes = 1
    End With
    With m_PicDesc
        .cbSize = LenB(m_PicDesc)
        .picType = 1 'PictureTypeConstants.vbPicTypeBitmap
    End With
    With m_PicGUID
        .Data1 = &H7BF80981
        .Data2 = &HBF32:  .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(3) = &HAA: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With
End Sub

Friend Sub New_(aPFN As String)
    Read aPFN
End Sub

Public Sub Read(aPFN As String)
    m_PFN = aPFN
Try: On Error GoTo Catch
    Dim sErr As String
    Dim FNr As Integer: FNr = FreeFile
    Open m_PFN For Binary Access Read As FNr
    If Not IsBitmap(FNr) Then sErr = "Unknown FileHeader: " & m_BmpHead.bfType: GoTo Catch
    ReadBmpInfo FNr
    If Not ReadData(FNr) Then sErr = "Reading pixel data failed": GoTo Catch
    GoTo Finally
Catch:
    MsgBox "Error reading file: " & vbCrLf & m_PFN & vbCrLf & sErr
Finally:
    Close FNr
End Sub

Private Function IsBitmap(ByVal FNr As Integer) As Boolean
    'reading fileheader
Try: On Error GoTo Catch
    With m_BmpHead
        Get FNr, , .bfType
        Get FNr, , .bfSize
        Get FNr, , .bfReserved1
        Get FNr, , .bfReserved2
        Get FNr, , .bfOffBits
        Dim T As Integer: T = .bfType
        IsBitmap = T = &H4D42 Or T = &H4142 Or T = &H4943 Or T = &H5043 Or T = &H4349 Or T = &H5450
        '&H4D42 = 19778 "BM" ' Windows 3.1x, 95, NT, . . .
        '&H4142 = 16706 "BA" ' OS/2 struct bitmap array
        '&H4943 = 18755 "CI" ' OS/2 struct color icon
        '&H5043 = 19778 "CP" ' OS/2 const color pointer
        '&H4349 = 20547 "IC" ' OS/2 struct icon
        '&H5450 = 21584 "PT" ' OS/2 pointer
    End With
Catch:
End Function

Private Function ReadBmpInfo(ByVal FNr As Integer) As Boolean
Try: On Error GoTo Catch
    Dim sErr As String
    With m_BmpInfo
        Get FNr, , .biSize
        ReadBmpInfo = .biSize = 40 Or .biSize = 108 Or .biSize = 124
        If Not ReadBmpInfo Then sErr = "Header-size not in the range (40, 108, 124): " & .biSize: GoTo Catch
        Get FNr, , .bmiHeader
        If .biSize = 40 Then
            Dim bic As Long: bic = .bmiHeader.biBitCount
            Dim bcu As Long: bcu = .bmiHeader.biClrUsed
            If (bic = 1 Or bic = 4 Or bic = 8) Then If bcu = 0 Then .bmiHeader.biClrUsed = 2 ^ bic
            Dim i As Long, aColor As BGRA
            For i = 0 To .bmiHeader.biClrUsed - 1
                Get FNr, , aColor
                'Debug.Print BGRA_ToStr(aColor)
                .bmiColors(i) = aColor
            Next
        ElseIf .biSize > 40 Then ' = 108
            Get FNr, , .bmiHeaderV4
            If .biSize > 108 Then ' = 124
                Get FNr, , .bmiHeaderV5
            End If
        End If
    End With
    Exit Function
Catch: MsgBox "ReadBmpInfo " & sErr
End Function

Private Function New_BGRA(ByVal R As Byte, ByVal G As Byte, ByVal b As Byte, ByVal A As Byte) As BGRA
    With New_BGRA
        .R = R: .G = G: .b = b: .A = A
    End With
End Function

Private Function BGRA_ToStr(this As BGRA) As String
    Dim s As String
    With this
        s = "  BGRA{&H" & Hex2(.b) & Hex2(.G) & Hex2(.R) & Hex2(.A) & "}"
    End With
    BGRA_ToStr = s
End Function

Private Function Hex2(b As Byte) As String
    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Private Function ReadData(FNr As Integer) As Boolean
Try: On Error GoTo Catch
    Dim Offset As Long: Offset = Seek(FNr) - 1
    'der Offset muss 4 aligned sein?hmmm ist gar nicht der Fall????
    ReadData = Offset = m_BmpHead.bfOffBits
    If Not ReadData Then GoTo Catch
    Dim u As Long: u = LOF(FNr) - Offset
    ReDim m_data(0 To u) As Byte
    Get FNr, , m_data
    ReadData = True
    Exit Function
Catch:
    MsgBox Err.Number & " " & "Error reading data" & vbCrLf & Err.Description
End Function

Friend Function ToPicture() As StdPicture
Try: On Error GoTo Catch
    Dim sErr As String, hr As Long
    With m_PicDesc
        Dim pBI As LongPtr
        If m_BmpInfo.biSize = 40 Then
            Dim bi As BITMAPINFOV3
            bi.biSize = m_BmpInfo.biSize
            bi.bmiHeader = m_BmpInfo.bmiHeader
            Dim i As Long
            For i = 0 To m_BmpInfo.bmiHeader.biClrUsed - 1
                bi.bmiColors(i) = m_BmpInfo.bmiColors(i)
            Next
            pBI = VarPtr(bi)
        Else
            pBI = VarPtr(m_BmpInfo)
        End If
        .hbitmap = CreateDIBSection(0&, ByVal pBI, 0&, 0&, 0&, 0&)
        If .hbitmap = 0 Then sErr = "hBitmap=0, CreateDIBSection failed": GoTo Catch
        hr = SetDIBits(0&, .hbitmap, 0&, Me.Height, m_data(0), ByVal pBI, 0&)
        If hr = 0 Then sErr = "SetDIBits failed": GoTo Catch
    End With
    hr = OleCreatePictureIndirect(m_PicDesc, m_PicGUID, 1&, ToPicture)
    If hr <> 0 Then sErr = "OleCreatePictureIndirect failed": GoTo Catch
    Exit Function
Catch:
    Dim mess As String: mess = "Could not convert to Picture"
    Dim lde As Long: lde = Err.LastDllError
    If lde Then mess = mess & vbCrLf & lde & " " & Err.Description & IIf(Len(sErr), vbCrLf & sErr, "")
    MsgBox mess
End Function

Public Function ToStr() As String
    ToStr = BITMAPFILEHEADER_ToStr(m_BmpHead) & vbCrLf & BITMAPINFO_ToStr(m_BmpInfo)
End Function

Public Property Get BitmapInfoVersion() As Integer
    With m_BmpInfo
        BitmapInfoVersion = IIf(.biSize = 40, 3, IIf(.biSize = 108, 4, IIf(.biSize = 124, 5, -1)))
    End With
End Property
'Private m_BmpHead As BITMAPFILEHEADER
'Private m_BmpInfo As BITMAPINFO

'    bfType      As Integer ' 2 ' "BM" = 66, 77 = &H424D = 16973
'    bfSize      As Long    ' 4 ' Größe der BMP-Datei in Byte. (unzuverlässig)
'    bfReserved1 As Integer ' 2 '
'    bfReserved2 As Integer ' 2 '
'    bfOffBits   As Long    ' 4 ' Offset to beginning of data
'Private m_BmpHead As BITMAPFILEHEADER
'Private m_BmpInfo As BITMAPINFO

Private Function BITMAPFILEHEADER_ToStr(this As BITMAPFILEHEADER) As String
    Dim s As String: s = "Fileheader" & vbCrLf
    With this
        s = s & "Type  : "
        Select Case .bfType
        Case &H4D42: s = s & "Windows bitmap"
        Case &H4142: s = s & "OS/2 bitmap"
        Case &H4943: s = s & "OS/2 color icon"
        Case &H5043: s = s & "OS/2 color pointer"
        Case &H4349: s = s & "OS/2 struct icon"
        Case &H5450: s = s & "OS/2 pointer"
        End Select
        s = s & vbCrLf
        s = s & "Size  : " & .bfSize & " Bytes" & vbCrLf
        s = s & "Offset: " & .bfOffBits & vbCrLf
    End With
    BITMAPFILEHEADER_ToStr = s
End Function

Private Function BITMAPINFO_ToStr(this As BITMAPINFO) As String
    Dim s As String:  s = "BitmapInfo" & vbCrLf
    Dim v As Integer: v = Me.BitmapInfoVersion
    With this
        s = s & "Version       : " & v & vbCrLf
        s = s & "PixelFormat   : " & PixelFormat_ToStr & vbCrLf
        With .bmiHeader
            s = s & "Width         : " & .biWidth & vbCrLf
            s = s & "Height        : " & .biHeight & vbCrLf
            s = s & "Planes        : " & .biPlanes & vbCrLf
            s = s & "BitCount      : " & .biBitCount & vbCrLf
            s = s & "Compression   : " & .biCompression & vbCrLf
            s = s & "SizeImage     : " & .biSizeImage & vbCrLf
            s = s & "XPelsPerMeter : " & .biXPelsPerMeter & "=" & GetDPI(.biXPelsPerMeter) & "dpi" & vbCrLf
            s = s & "YPelsPerMeter : " & .biYPelsPerMeter & "=" & GetDPI(.biYPelsPerMeter) & "dpi" & vbCrLf
            s = s & "ColorsUsed    : " & .biClrUsed & vbCrLf
            s = s & "ClrImportant  : " & .biClrImportant & vbCrLf
        End With
        If v = 3 Then
            Dim i As Long
            For i = 0 To .bmiHeader.biClrUsed - 1
                s = s & BGRA_ToStr(.bmiColors(i)) & vbCrLf
            Next
            BITMAPINFO_ToStr = s:            Exit Function
        End If
        With .bmiHeaderV4
            s = s & "RedMask       : " & .bV4RedMask & vbCrLf
            s = s & "GreenMask     : " & .bV4GreenMask & vbCrLf
            s = s & "BlueMask      : " & .bV4BlueMask & vbCrLf
            s = s & "AlphaMask     : " & .bV4AlphaMask & vbCrLf
            s = s & "CSType        : " & Hex(.bV4CSType) & vbCrLf
            s = s & "Endpoints     : " & CIEXYZTRIPLE_ToStr(.bV4Endpoints) & vbCrLf
            s = s & "GammaRed      : " & .bV4GammaRed & vbCrLf
            s = s & "GammaGreen    : " & .bV4GammaGreen & vbCrLf
            s = s & "GammaBlue     : " & .bV4GammaBlue & vbCrLf
            If v = 4 Then
                BITMAPINFO_ToStr = s:        Exit Function
            End If
        End With
        With .bmiHeaderV5
            s = s & "Intent        : " & .bV5Intent & vbCrLf
            s = s & "ProfileData   : " & .bV5ProfileData & vbCrLf
            s = s & "ProfileSize   : " & .bV5ProfileSize & vbCrLf
            s = s & "Reserved      : " & .bV5Reserved & vbCrLf
            If v = 5 Then
                BITMAPINFO_ToStr = s:        Exit Function
            End If
        End With
    End With
End Function
Private Function GetDPI(ByVal PixelsPerMeter As Long) As Long
    GetDPI = PixelsPerMeter * 2.54 / 100
End Function
'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-ciexyztriple
Private Function CIEXYZTRIPLE_ToStr(this As CIEXYZTRIPLE) As String
    Dim s As String: s = vbCrLf
    With this
        s = s & "  ciexyzRed  : " & CIEXYZ_ToStr(.ciexyzRed) & vbCrLf
        s = s & "  ciexyzGreen: " & CIEXYZ_ToStr(.ciexyzGreen) & vbCrLf
        s = s & "  ciexyBlue  : " & CIEXYZ_ToStr(.ciexyBlue) & vbCrLf
    End With
    CIEXYZTRIPLE_ToStr = s
End Function

Private Function CIEXYZ_ToStr(this As CIEXYZ) As String
    Dim s As String: s = vbCrLf
    With this
        s = s & "    ciexyzX: " & .ciexyzX & vbCrLf
        s = s & "    ciexyzY: " & .ciexyzY & vbCrLf
        s = s & "    ciexyzZ: " & .ciexyzZ & vbCrLf
    End With
    CIEXYZ_ToStr = s
End Function

Private Function PixelFormat_ToStr() As String
    Dim pf As EPixelFormat: pf = Me.PixelFormat
    Dim s As String
    Select Case pf
    Case 0:        s = "DontCare" 'Es wurde kein Pixelformat angegeben.
    Case &H10000:  s = "Indexed"               '  &H10000   65536  Die Pixeldaten enthalten farbindizierte Werte, d. h., diese sind ein Index für die Farben in der Systemfarbtabelle, im Gegensatz zu einzelnen Farbwerten.
    Case &H20000:  s = "Gdi"                   '  &H20000  131072  Die Pixeldaten enthalten GDI-Farben.
    Case &H40000:  s = "Alpha "                '  &H40000  262144  Die Pixeldaten enthalten Alphawerte, die nicht zuvor multipliziert wurden.
    Case &H80000:  s = "PAlpha"                '  &H80000  524288  Das Pixelformat enthält zuvor multiplizierte Alphawerte.
    Case &H100000: s = "Extended"              ' &H100000 1048576  Reserviert.
    Case &H200000: s = "Canonical"             ' &H200000 2097152  Das Standardpixelformat ist 32 Bits pro Pixel.Das Format gibt 24-Bit Farbtiefe und einen 8-Bit-Alphakanal an.
    
    Case &H30101:  s = "Format1bppIndexed"     '  &H30101  196865  Gibt an, dass das Format  1 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält   2 Farben.
    Case &H30402:  s = "Format4bppIndexed"     '  &H30402  197634  Gibt an, dass das Format  4 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält  16 Farben.
    Case &H30803:  s = "Format8bppIndexed"     '  &H30803  198659  Gibt an, dass das Format  8 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält 256 Farben.
    
    'Graustufen? es gibt kein Format das darauf festgelegt werden kann was soll das?
    Case &H101004: s = "Format16bppGrayScale"  ' &H101004 1052676  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen geben 65.536 Grauschattierungen an.
    Case &H21005:  s = "Format16bppRgb555"     '  &H21005  135173  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot-, Blau- und Grünanteil werden jeweils 5 Bit verwendet. Das verbleibende Bit wird nicht verwendet.
    Case &H21006:  s = "Format16bppRgb565"     '  &H21006  135174  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot- und Blauanteil werden jeweils 5 Bit und für den Grünanteil 6 Bit verwendet.
    Case &H61007:  s = "Format16bppArgb1555"   '  &H61007  397319  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen liefern 32.768 Farbschattierungen, wobei der Rot-, Grün- und Blauanteil jeweils von 5 Bits und die Alphakomponente von 1 Bit wiedergegeben wird.
    Case &H21808:  s = "Format24bppRgb"        '  &H21808  137224  Gibt an, dass das Format 24 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Case &H22009:  s = "Format32bppRgb"        '  &H22009  139273  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die verbleibendenden 8 Bits werden nicht verwendet.
    Case &H26200A: s = "Format32bppArgb"       ' &H26200A 2498570  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Case &HE200B:  s = "Format32bppPArgb"      '  &HE200B  925707  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    Case &H10300C: s = "Format48bppRgb"        ' &H10300C 1060876  Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    Case &H34400D: s = "Format64bppArgb"       ' &H34400D 3424269  Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.
    Case &H1C400E: s = "Format64bppPArgb"      ' &H1C400E 1851406  Gibt an, dass das Format 64 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 16 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    End Select
    PixelFormat_ToStr = s
End Function

Public Property Get PixelFormat() As EPixelFormat
    Dim epf As EPixelFormat
    With m_BmpInfo.bmiHeader
        Dim bic As Integer: bic = .biBitCount
        Select Case bic
        Case 1:  If .biClrUsed > 0 Then epf = Format1bppIndexed
        Case 4:  If .biClrUsed > 0 Then epf = Format4bppIndexed
        Case 8:  If .biClrUsed > 0 Then epf = Format8bppIndexed
        Case 16:
                 'How to distinguish the different 16-bit pixel-color-formats? -> by using the GUIDs of the AM_MEDIA-Type
                 'epf = Format16bppArgb1555
                 'epf = Format16bppGrayScale
                 'epf = Format16bppRgb555
                 'epf = Format16bppRgb555
                 epf = Format16bppRgb565
                 
        Case 24: epf = Format24bppRgb
        Case 32: epf = Format32bppArgb
        End Select
'        If bic = 1 Or bic = 4 Or bic = 8 Then
'            If .biClrUsed > 0 Then
'                epf =
'            End If
'        End If
    End With
    PixelFormat = epf
End Property

Public Property Get Width() As Long
    Width = m_BmpInfo.bmiHeader.biWidth
End Property
Public Property Let Width(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biWidth = Value
End Property

Public Property Get Height() As Long
    Height = m_BmpInfo.bmiHeader.biHeight
End Property
Public Property Let Height(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biHeight = Value
End Property

'Public Property Get Pixel(ByVal x As Long, ByVal y As Long) As Long
'    Dim c As Byte: c = m_data(y * m_Width + x)
'    Pixel = RGB(c, c, c)
'End Property
'Public Property Let Pixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
'    Dim l As TLng: l.Value = Color
'    Dim c As RGBA: LSet c = l
'    With c
'        m_data(y * m_Width + x) = (.R + .G + .B) \ 3
'    End With
'End Property

Public Property Get Pixel(ByVal X As Long, ByVal Y As Long) As Long
'    Dim Index As Long: Index = Y * Me.Width + X
'    If m_PNMType = EPNMType.PortableBitmapBin Then Index = Index / 8
'    If Index > UBound(m_data) Then Index = UBound(m_data)
'    Data = m_data(Index)
End Property

Public Property Let Pixel(ByVal X As Long, ByVal Y As Long, Value As Long)
'    Dim Index As Long
'    Index = Y * Me.Width + X
'    m_data(Index) = Value
End Property

