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
    'Undefined =        0&          '      &H0       0  Das Pixelformat ist nicht definiert oder
    'DontCare  =        0&          '      &H0       0  Es wurde kein Pixelformat angegeben.
    'Indexed   =  &H10000           '  &H10000   65536  Die Pixeldaten enthalten farbindizierte Werte, d. h., diese sind ein Index für die Farben in der Systemfarbtabelle, im Gegensatz zu einzelnen Farbwerten.
    'Gdi       =  &H20000           '  &H20000  131072  Die Pixeldaten enthalten GDI-Farben.
    'Alpha     =  &H40000           '  &H40000  262144  Die Pixeldaten enthalten Alphawerte, die nicht zuvor multipliziert wurden.
    'PAlpha    =  &H80000           '  &H80000  524288  Das Pixelformat enthält zuvor multiplizierte Alphawerte.
    'Extended  = &H100000           ' &H100000 1048576  Reserviert.
    'Canonical = &H200000           ' &H200000 2097152  Das Standardpixelformat ist 32 Bits pro Pixel.Das Format gibt 24-Bit Farbtiefe und einen 8-Bit-Alphakanal an.
    
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
    B As Byte
    G As Byte
    R As Byte
    A As Byte
End Type

Private Type RGBA
    R As Byte
    G As Byte
    B As Byte
    A As Byte
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmap
Private Type BitmapHeader
    bmType       As Long    ' The bitmap type. This member must be zero.
    bmWidth      As Long    ' The width, in pixels, of the bitmap. The width must be greater than zero.
    bmHeight     As Long    ' The height, in pixels, of the bitmap. The height must be greater than zero.
    bmWidthBytes As Long    ' The number of bytes in each scan line. This value must be divisible by 2, because the system assumes that the bit values of a bitmap form an array that is word aligned.
    bmPlanes     As Integer ' The count of color planes.
    bmBitsPixel  As Integer ' The number of bits required to indicate the color of a pixel.
    bmBits       As Long    ' A pointer to the location of the bit values for the bitmap. The bmBits member must be a pointer to an array of character (1-byte) values.
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapfileheader
Private Type BITMAPFILEHEADER
'do not read at once read every single variable separately instead
'is only used to determine type not for the bitmap data
    bfType      As Integer ' 2 ' The file type; must be BM. "BM" = 66, 77 = &H424D = 16973
    bfSize      As Long    ' 4 ' The size, in bytes, of the bitmap file. Größe der BMP-Datei in Byte. (unzuverlässig)
    bfReserved1 As Integer ' 2 ' Reserved; must be zero.
    bfReserved2 As Integer ' 2 ' Reserved; must be zero.
    bfOffBits   As Long    ' 4 ' The offset, in bytes, from the beginning of the BITMAPFILEHEADER structure to the bitmap bits. Offset to beginning of data
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
Private Const BI_RGB       As Long = 0&  ' An uncompressed format.
Private Const BI_RLE8      As Long = 1&  ' A run-length encoded (RLE) format for bitmaps with 8 bpp. The compression format is a 2-byte format consisting of a count byte followed by a byte containing a color index. For more information, see Bitmap Compression.
Private Const BI_RLE4      As Long = 2&  ' An RLE format for bitmaps with 4 bpp. The compression format is a 2-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
Private Const BI_BITFIELDS As Long = 3&  ' Specifies that the bitmap is not compressed. The members bV4RedMask, bV4GreenMask, and bV4BlueMask specify the red, green, and blue components for each pixel. This is valid when used with 16- and 32-bpp bitmaps
                                         ' If biCompression = BI_BITFIELDS -> the format is either RGB-555 or RGB-565.
                                         ' Use the subtype GUID in the AM_MEDIA_TYPE structure to determine the specific RGB type.
                                         ' https://learn.microsoft.com/en-us/windows/win32/api/strmif/ns-strmif-am_media_type
                                        
Private Const BI_JPEG      As Long = 4&  ' Specifies that the image is compressed using the JPEG file interchange format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
Private Const BI_PNG       As Long = 5&  ' Specifies that the image is compressed using the PNG  file interchange format.
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
    With m_BmpHead
        .bfType = &H4D42
        '.bfSize
        '.bfOffBits
    End With
    With m_BmpInfo
        .biSize = 40  'LenB(.bmiHeader) ' nope not LenB anymore ' a default value for V3-header because v4 or v5 is not so often
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

Friend Sub NewWH(ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As EPixelFormat)
    With m_BmpHead
        '.bfType = &H4D42
        '.bfReserved1 = 0
        '.bfReserved2 = 0
        '.bfSize =
        '.bfOffBits
    End With
    Dim pf As EPixelFormat
    Select Case Me.PixelFormat
    Case pf = Format1bppIndexed
    Case pf = Format4bppIndexed
    Case pf = Format8bppIndexed
    Case pf = Format16bppArgb1555
    Case pf = Format16bppGrayScale
    Case pf = Format16bppRgb565
    End Select
    With m_BmpInfo
        '
    End With
End Sub

Friend Sub NewC(other As Bitmap)
    Me.NewWH other.Width, other.Height, other.PixelFormat
End Sub

Public Function Clone() As Bitmap
    Set Clone = New Bitmap
    Clone.NewC Me
End Function

Public Property Get DataPtr()
    DataPtr = VarPtr(m_data(0))
End Property

Public Property Get DataSize() As Long
    DataSize = UBound(m_data) + 1
End Property

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
        IsBitmap = (T = &H4D42) Or (T = &H4142) Or (T = &H4943) Or (T = &H5043) Or (T = &H4349) Or (T = &H5450)
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

Private Function New_BGRA(ByVal R As Byte, ByVal G As Byte, ByVal B As Byte, ByVal A As Byte) As BGRA
    With New_BGRA
        .R = R: .G = G: .B = B: .A = A
    End With
End Function

Private Function BGRA_ToStr(this As BGRA) As String
    Dim s As String
    With this
        s = "  BGRA{&H" & Hex2(.B) & Hex2(.G) & Hex2(.R) & Hex2(.A) & "}"
    End With
    BGRA_ToStr = s
End Function

Private Function Hex2(B As Byte) As String
    Hex2 = Hex(B): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Private Function BGRA_ToRGBA(this As BGRA) As RGBA
    With this
        BGRA_ToRGBA.A = .A
        BGRA_ToRGBA.B = .B
        BGRA_ToRGBA.G = .G
        BGRA_ToRGBA.R = .R
    End With
End Function

Private Function RGBA_ToBGRA(this As RGBA) As BGRA
    With this
        RGBA_ToBGRA.A = .A
        RGBA_ToBGRA.B = .B
        RGBA_ToBGRA.G = .G
        RGBA_ToBGRA.R = .R
    End With
End Function

Private Function ReadData(FNr As Integer) As Boolean
Try: On Error GoTo Catch
    Dim Offset As Long: Offset = Seek(FNr) - 1
    'der Offset muss 4 aligned sein?hmmm ist gar nicht der Fall????
    ReadData = Offset = m_BmpHead.bfOffBits
    If Not ReadData Then GoTo Catch
    Dim u As Long: u = LOF(FNr) - Offset - 1
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
            s = s & "XPelsPerMeter : " & .biXPelsPerMeter & " = " & GetDPI(.biXPelsPerMeter) & " dpi" & vbCrLf
            s = s & "YPelsPerMeter : " & .biYPelsPerMeter & " = " & GetDPI(.biYPelsPerMeter) & " dpi" & vbCrLf
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
            s = s & "RedMask       : &H" & Hex(.bV4RedMask) & vbCrLf
            s = s & "GreenMask     : &H" & Hex(.bV4GreenMask) & vbCrLf
            s = s & "BlueMask      : &H" & Hex(.bV4BlueMask) & vbCrLf
            s = s & "AlphaMask     : &H" & Hex(.bV4AlphaMask) & vbCrLf
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
    Case &H10000:  s = "Indexed"           '  &H10000   65536  Die Pixeldaten enthalten farbindizierte Werte, d. h., diese sind ein Index für die Farben in der Systemfarbtabelle, im Gegensatz zu einzelnen Farbwerten.
    Case &H20000:  s = "Gdi"               '  &H20000  131072  Die Pixeldaten enthalten GDI-Farben.
    Case &H40000:  s = "Alpha "            '  &H40000  262144  Die Pixeldaten enthalten Alphawerte, die nicht zuvor multipliziert wurden.
    Case &H80000:  s = "PAlpha"            '  &H80000  524288  Das Pixelformat enthält zuvor multiplizierte Alphawerte.
    Case &H100000: s = "Extended"          ' &H100000 1048576  Reserviert.
    Case &H200000: s = "Canonical"         ' &H200000 2097152  Das Standardpixelformat ist 32 Bits pro Pixel.Das Format gibt 24-Bit Farbtiefe und einen 8-Bit-Alphakanal an.
    
    Case &H30101:  s = "1 bpp Indexed"     '  &H30101  196865  Gibt an, dass das Format  1 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält   2 Farben.
    Case &H30402:  s = "4 bpp Indexed"     '  &H30402  197634  Gibt an, dass das Format  4 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält  16 Farben.
    Case &H30803:  s = "8 bpp Indexed"     '  &H30803  198659  Gibt an, dass das Format  8 Bit pro Pixel ist. Es werden indizierte Farben verwendet. Die Farbtabelle enthält 256 Farben.
    
    Case &H101004: s = "16 bpp GrayScale"  ' &H101004 1052676  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen geben 65.536 Grauschattierungen an.
    Case &H21005:  s = "16 bpp RGB 555"    '  &H21005  135173  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot-, Blau- und Grünanteil werden jeweils 5 Bit verwendet. Das verbleibende Bit wird nicht verwendet.
    Case &H21006:  s = "16 bpp RGB 565"    '  &H21006  135174  Gibt an, dass das Format 16 Bit pro Pixel ist. Für den Rot- und Blauanteil werden jeweils 5 Bit und für den Grünanteil 6 Bit verwendet.
    Case &H61007:  s = "16 bpp ARGB 1555"  '  &H61007  397319  Gibt an, dass das Format 16 Bit pro Pixel ist. Die Farbinformationen liefern 32.768 Farbschattierungen, wobei der Rot-, Grün- und Blauanteil jeweils von 5 Bits und die Alphakomponente von 1 Bit wiedergegeben wird.
    Case &H21808:  s = "24 bpp RGB"        '  &H21808  137224  Gibt an, dass das Format 24 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Case &H22009:  s = "32 bpp RGB"        '  &H22009  139273  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die verbleibendenden 8 Bits werden nicht verwendet.
    Case &H26200A: s = "32 bpp ARGB"       ' &H26200A 2498570  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.
    Case &HE200B:  s = "32 bpp PARGB"      '  &HE200B  925707  Gibt an, dass das Format 32 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 8 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    Case &H10300C: s = "48 bpp RGB"        ' &H10300C 1060876  Specifies that the format is 48 bits per pixel; 16 bits each are used for the red, green, and blue components.
    Case &H34400D: s = "64 bpp ARGB"       ' &H34400D 3424269  Specifies that the format is 64 bits per pixel; 16 bits each are used for the alpha, red, green, and blue components.
    Case &H1C400E: s = "64 bpp PARGB"      ' &H1C400E 1851406  Gibt an, dass das Format 64 Bits pro Pixel ist, wobei für den Alpha-, Rot-, Grün- und Blauanteil jeweils 16 Bits verwendet werden.Die Rot-, Grün- und Blaukomponente wird entsprechend der Alphakomponente im Voraus multipliziert.
    End Select
    PixelFormat_ToStr = s
End Function

Public Property Get PixelFormat() As EPixelFormat
    Dim epf As EPixelFormat
    Dim bic As Integer: bic = m_BmpInfo.bmiHeader.biBitCount
    Dim bcu As Integer: bcu = m_BmpInfo.bmiHeader.biClrUsed
    Select Case bic
    Case 1:     If bcu > 0 Then epf = EPixelFormat.Format1bppIndexed
    Case 4:     If bcu > 0 Then epf = EPixelFormat.Format4bppIndexed
    Case 8:     If bcu > 0 Then epf = EPixelFormat.Format8bppIndexed
    Case 16:
                
                'use the ColorMasks of v4-header to distinguish the different 16-bit-color-formats
                If m_BmpInfo.bmiHeaderV4.bV4AlphaMask = &H8000& Then
                    epf = EPixelFormat.Format16bppArgb1555
                Else
                    If m_BmpInfo.bmiHeaderV4.bV4GreenMask = &H3E0 Then
                        epf = EPixelFormat.Format16bppRgb555
                    Else
                        epf = EPixelFormat.Format16bppRgb565
                    End If
                End If
    Case 24:    epf = EPixelFormat.Format24bppRgb
    Case 32:    epf = EPixelFormat.Format32bppArgb
                'How to distinguish the different 32bpp formats?
                'epf = EPixelFormat.Format32bppPArgb
                'epf = EPixelFormat.Format32bppRgb
    End Select
    PixelFormat = epf
End Property

Public Property Get Width() As Long
    Width = m_BmpInfo.bmiHeader.biWidth
End Property

Public Property Let Width(ByVal Value As Long)
    If Value < 0 Then Exit Property
    If m_BmpInfo.bmiHeader.biWidth = Value Then Exit Property
    Resize Value, Me.Height
End Property

Public Property Get Height() As Long
    Height = m_BmpInfo.bmiHeader.biHeight
End Property

Public Property Let Height(ByVal Value As Long)
    If Value < 0 Then Exit Property
    If m_BmpInfo.bmiHeader.biHeight = Value Then Exit Property
    Resize Me.Width, Value
End Property

Public Function Resize(ByVal NewWidth As Long, ByVal NewHeight As Long) As Boolean
    Dim W As Long: W = Me.Width
    Dim H As Long: H = Me.Height
    'return if nothing has changed
    If W = NewWidth And H = NewHeight Then Exit Function
    If W <> NewWidth Then m_BmpInfo.bmiHeader.biWidth = NewWidth
    If H <> NewHeight Then m_BmpInfo.bmiHeader.biHeight = NewHeight
    Dim factor As Single
    Select Case Me.PixelFormat
    Case EPixelFormat.Format1bppIndexed: factor = 1 / 8
    Case EPixelFormat.Format4bppIndexed: factor = 1 / 2
    Case EPixelFormat.Format8bppIndexed: 'Then OK, do nothing, leave it as it is
    Case EPixelFormat.Format16bppArgb1555, EPixelFormat.Format16bppGrayScale, EPixelFormat.Format16bppRgb555, EPixelFormat.Format16bppRgb565
                                         factor = 2
    Case EPixelFormat.Format24bppRgb:    factor = 3
    Case EPixelFormat.Format32bppArgb, EPixelFormat.Format32bppPArgb, EPixelFormat.Format32bppRgb
                                         factor = 4
    Case EPixelFormat.Format48bppRgb:    factor = 6
    Case EPixelFormat.Format64bppArgb, EPixelFormat.Format64bppPArgb
                                         factor = 8
    End Select
    Dim NewDataSize As Long: NewDataSize = NewWidth * NewHeight * factor 'in Bytes
    'we do not preserve here
    ReDim m_data(0 To NewDataSize - 1)
End Function

Public Property Get IsIndexed() As Boolean
    IsIndexed = Me.PaletteCount > 0 'm_BmpInfo.bmiHeader.biClrUsed > 0
End Property

Public Property Get PaletteCount() As Long
    PaletteCount = m_BmpInfo.bmiHeader.biClrUsed
End Property

Public Property Get PaletteColor(ByVal Index As Long) As Long
    If Index < 0 Or Me.PaletteCount <= Index Then Exit Property
    Dim c As TLng: LSet c = m_BmpInfo.bmiColors(Index)
    PaletteColor = c.Value
End Property
Public Property Let PaletteColor(ByVal Index As Long, ByVal Value As Long)
    If Index < 0 Or Me.PaletteCount <= Index Then Exit Property
    Dim c As TLng: c.Value = Value
    LSet m_BmpInfo.bmiColors(Index) = c
End Property

Public Property Get Pixel(ByVal X As Long, ByVal Y As Long) As Long
    Select Case Me.PixelFormat
    Case EPixelFormat.Format1bppIndexed:    Pixel = Me.Pixel1(X, Y)
    Case EPixelFormat.Format4bppIndexed:    Pixel = Me.Pixel4(X, Y)
    Case EPixelFormat.Format8bppIndexed:    Pixel = Me.Pixel8(X, Y)
    Case EPixelFormat.Format16bppArgb1555, EPixelFormat.Format16bppGrayScale, EPixelFormat.Format16bppRgb555, EPixelFormat.Format16bppRgb565
                                            Pixel = Me.Pixel16(X, Y)
    Case EPixelFormat.Format24bppRgb:       Pixel = Me.Pixel24(X, Y)
    Case EPixelFormat.Format32bppArgb, EPixelFormat.Format32bppPArgb, EPixelFormat.Format32bppRgb
                                            Pixel = Me.Pixel32(X, Y)
    End Select
End Property

'Public Property Let Pixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
'    Dim l As TLng: l.Value = Color
'    Dim c As RGBA: LSet c = l
'    With c
'        m_data(y * m_Width + x) = (.R + .G + .B) \ 3
'    End With
'End Property
'
Public Property Get Pixel1(ByVal X As Long, ByVal Y As Long) As Long
    Pixel1 = 0
End Property
'Public Property Let Pixel1(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property
'
Public Property Get Pixel4(ByVal X As Long, ByVal Y As Long) As Long
    Dim W As Long: W = Me.Width:     If X < 0 Or W <= X Then Exit Property
    Dim H As Long: H = Me.Height:    If Y < 0 Or H <= Y Then Exit Property
    If Not Me.IsIndexed Then Exit Property
    Dim si As Single: si = ((H - 1 - Y) * W + X) / 2
    Dim i  As Long:    i = Int(si)
    'Debug.Print UBound(m_data)
    If si = CSng(i) Then
        'the lower nibble
        i = &HF And m_data(i)
    Else
        'the higher nibble
        i = &HF0 And m_data(i)
        i = i / &H10
    End If
    Dim BGRA As BGRA: BGRA = m_BmpInfo.bmiColors(i)
    Dim RGBA As RGBA: RGBA = BGRA_ToRGBA(BGRA)
    Dim l As TLng: LSet l = RGBA
    Pixel4 = l.Value
End Property
'Public Property Let Pixel4(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel8(ByVal X As Long, ByVal Y As Long) As Long
    Dim W As Long: W = Me.Width:     If X < 0 Or W <= X Then Exit Property
    Dim H As Long: H = Me.Height:    If Y < 0 Or H <= Y Then Exit Property
    If Not Me.IsIndexed Then Exit Property
    Dim i As Long: i = (H - 1 - Y) * W + X: i = m_data(i)
    Dim BGRA As BGRA: BGRA = m_BmpInfo.bmiColors(i)
    Dim RGBA As RGBA: RGBA = BGRA_ToRGBA(BGRA)
    Dim l As TLng: LSet l = RGBA
    Pixel8 = l.Value
    'Pixel8 = RGB(BGRA.R, BGRA.G, BGRA.B)
End Property

'Public Property Let Pixel8(ByVal X As Long, ByVal Y As Long, ByVal Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel16(ByVal X As Long, ByVal Y As Long) As Long
    Dim W As Long: W = Me.Width:     If X < 0 Or W <= X Then Exit Property
    Dim H As Long: H = Me.Height:    If Y < 0 Or H <= Y Then Exit Property
    Dim i As Long: i = ((H - 1 - Y) * W + X) * 2
    Dim b1 As Byte: b1 = m_data(i)
    Dim b2 As Byte: b2 = m_data(i + 1)
    Dim B As BGRA: B.B = b1: B.G = b2
    Dim l As TLng: LSet l = B
    'l.Value=rgb(
    Pixel16 = l.Value
    'Pixel16=
End Property
'Public Property Let Pixel16(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel24(ByVal X As Long, ByVal Y As Long) As Long
    Dim W As Long: W = Me.Width:    If X < 0 Or W <= X Then Exit Property
    Dim H As Long: H = Me.Height:   If Y < 0 Or H <= Y Then Exit Property
    Dim i As Long: i = ((H - 1 - Y) * W + X) * 3
    'Dim BGRA As BGRA
    Dim RGBA As RGBA
    With RGBA
        .R = m_data(i + 2)
        .G = m_data(i + 1)
        .B = m_data(i)
    End With
    Dim l As TLng: LSet l = RGBA 'BGRA
    Pixel24 = l.Value
    'Pixel24 = RGB(RGBA.R, RGBA.G, RGBA.B)
    'Pixel24 = RGB(m_data(i + 2), m_data(i + 1), m_data(i))
End Property
'Public Property Let Pixel24(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel32(ByVal X As Long, ByVal Y As Long) As Long
    If X < 0 Or Me.Width <= X Then Exit Property
    If Y < 0 Or Me.Height <= Y Then Exit Property
    Dim W As Long: W = Me.Width
    Dim H As Long: H = Me.Height
    Dim i As Long: i = ((H - Y) * W + X) * 4
    Dim BGRA As BGRA
    With BGRA
        .B = m_data(i)
        .G = m_data(i + 1)
        .R = m_data(i + 2)
        .A = m_data(i + 3)
    End With
    Dim l As TLng: LSet l = BGRA
    Pixel32 = l.Value
End Property
Public Property Let Pixel32(ByVal X As Long, ByVal Y As Long, Value As Long)
    'm_data=0
End Property

'Public Property Get Pixel48(ByVal X As Long, ByVal Y As Long) As Long
'    'Pixel48 = 0
'End Property
'Public Property Let Pixel48(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property
'
'Public Property Get Pixel64(ByVal X As Long, ByVal Y As Long) As Long
'    Pixel64 = 0
'End Property
'Public Property Let Pixel64(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property
