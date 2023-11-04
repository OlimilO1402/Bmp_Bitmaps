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
'https://www.fileformat.info/format/os2bmp/egff.htm
'https://stackoverflow.com/questions/30883102/why-does-windows-gdi-use-rgba-format-for-colorref-instead-of-bgra/48576177#48576177

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

Private Type TByt2
    Value0 As Byte
    Value1 As Byte
End Type

Private Type TByt4 'used to convert 16bpp-color to 32bpp-color
    Value0 As Byte
    Value1 As Byte
    Reserv As Integer
End Type

Private Type TInt
    Value As Integer
End Type

Private Type TInt2
    Value0 As Integer
    Value1 As Integer
End Type

Private Type TLng
    Value As Long
End Type

Private Type BGRA '
    b As Byte
    G As Byte
    R As Byte
    a As Byte
End Type

Private Type RGBA
    R As Byte
    G As Byte
    b As Byte
    a As Byte
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmap
'The name "BITMAP" is not suitable here because the class's name is "Bitmap"
'so we call it BitmapHeader, maybe too much "Header" but anyway deal with it!
Private Type BitmapHeader
    bmType       As Long    ' 4 ' The Bitmap type. This member must be zero.
    bmWidth      As Long    ' 4 ' The width, in pixels, of the Bitmap. The width must be greater than zero.
    bmHeight     As Long    ' 4 ' The height, in pixels, of the Bitmap. The height must be greater than zero.
    bmWidthBytes As Long    ' 4 ' The number of bytes in each scan line. This value must be divisible by 2, because the system assumes that the bit values of a Bitmap form an array that is word aligned.
    bmPlanes     As Integer ' 2 ' The count of color planes.
    bmBitsPixel  As Integer ' 2 ' The number of bits required to indicate the color of a pixel.
    bmBits       As Long    ' 4 ' A pointer to the location of the bit values for the Bitmap. The bmBits member must be a pointer to an array of character (1-byte) values.
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmapfileheader
Private Type BITMAPFILEHEADER
'do not read at once read every single variable separately instead
'is only used to determine type not for the Bitmap data
    bfType      As Integer ' 2 ' The file type; must be BM. "BM" = 66, 77 = &H424D = 16973
    bfSize      As Long    ' 4 ' The size, in bytes, of the Bitmap file. Größe der BMP-Datei in Byte. (unzuverlässig)
    bfReserved1 As Integer ' 2 ' Reserved; must be zero.
    bfReserved2 As Integer ' 2 ' Reserved; must be zero.
    bfOffBits   As Long    ' 4 ' The offset, in bytes, from the beginning of the BITMAPFILEHEADER structure to the Bitmap bits. Offset to beginning of data
End Type             ' Sum: 14

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmapinfoheader
'aka BITMAPV3HEADER aka DIBHEADER the old and first Bitmap-header
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
Private Const BI_RLE8      As Long = 1&  ' A run-length encoded (RLE) format for Bitmaps with 8 bpp. The compression format is a 2-byte format consisting of a count byte followed by a byte containing a color index. For more information, see Bitmap Compression.
Private Const BI_RLE4      As Long = 2&  ' An RLE format for Bitmaps with 4 bpp. The compression format is a 2-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
Private Const BI_BITFIELDS As Long = 3&  ' Specifies that the Bitmap is not compressed. The members bV4RedMask, bV4GreenMask, and bV4BlueMask specify the red, green, and blue components for each pixel. This is valid when used with 16- and 32-bpp Bitmaps
                                         ' If biCompression = BI_BITFIELDS -> the format is either RGB-555 or RGB-565.
                                         ' Use the subtype GUID in the AM_MEDIA_TYPE structure to determine the specific RGB type.
                                         ' https://learn.microsoft.com/en-us/windows/win32/api/strmif/ns-strmif-am_media_type
                                        
Private Const BI_JPEG      As Long = 4&  ' Specifies that the image is compressed using the JPEG file interchange format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
Private Const BI_PNG       As Long = 5&  ' Specifies that the image is compressed using the PNG  file interchange format.
Private Const BI_1632      As Long = &H32333631 ' "1632"
Private Const LCS_CALIBRATED_RGB As Long = &H0&

'DIB_-constants
Private Enum DIBColors
    DIB_RGB_COLORS = &H0&  ' The color table contains literal RGB values.
    DIB_PAL_COLORS = &H1&  ' The color table consists of an array of 16-bit indexes into the LogPalette object (section 2.2.17) that is currently defined in the playback device context.
    DIB_PAL_INDICES = &H2& ' No color table exists. The pixels in the DIB are indices into the current logical palette in the playback device context.
End Enum
'Private Const DIB_RGB_COLORS As Long = 0
'Private Const DIB_PAL_COLORS As Long = 1
'Private Const DIB_PAL_INDICES As Long = 2    'Private Const DIB_PAL_PHYSINDICES As Long = 2
'Private Const DIB_PAL_LOGINDICES As Long = 4

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

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmapv4header
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

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmapv5header
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
    
End Type                  ' Sum: 124

Private Type TPalette
    bmiColors(0 To 255) As BGRA
End Type
Private Type BITMAPINFOV3
    biSize      As Long
    bmiHeader   As BITMAPINFOHEADER
    Palette     As TPalette
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-Bitmapinfo
Private Type BITMAPINFO
    biSize      As Long
    bmiHeader   As BITMAPINFOHEADER
    bmiHeaderV4 As AddBITMAPV4HEADER
    bmiHeaderV5 As AddBITMAPV5HEADER
    'bmiColors(0 To 255) As BGRA
    Palette     As TPalette
End Type

'https://docs.microsoft.com/en-us/windows/win32/api/olectl/ns-olectl-pictdesc
Private Type PICTDESC
    cbSize  As Long    ' 4
    picType As Long    ' 4    ' PictureTypeConstants
    hBitmap As LongPtr ' 4, 8 ' HBITMAP
    hPal    As LongPtr ' 4, 8 ' HPALETTE
End Type         ' Sum: 16/24

#If VBA7 Then
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createdibsection
    Private Declare PtrSafe Function CreateDIBSection Lib "gdi32" (ByVal hdc As LongPtr, ByVal pbmi As LongPtr, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal offset As Long) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdibits
    Private Declare PtrSafe Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-setdibits
    Private Declare PtrSafe Function SetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getobjecta
    Private Declare PtrSafe Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLength As Long)
    Private Declare PtrSafe Sub RtlFillMemory Lib "kernel32" (pDst As Any, ByVal bytLength As Long, ByVal fill As Byte)
    
    ' ----==== USER32 API Deklarationen ====----
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    
#Else
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createdibsection
    Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, ByVal pbmi As LongPtr, ByVal usage As Long, ByVal ppvBits As LongPtr, ByVal hSection As LongPtr, ByVal Offset As Long) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdibits
    Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-setdibits
    Private Declare Function SetDIBits Lib "gdi32" (ByVal aHDC As LongPtr, ByVal hBitmap As LongPtr, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByVal lpBI As LongPtr, ByVal wUsage As Long) As Long
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef riid As GUID, ByVal fOwn As Long, ByRef lplpvObj As Object) As Long
    'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getobjecta
    Private Declare Function GetObjectW Lib "gdi32" (ByVal hObject As LongPtr, ByVal nCount As Long, ByRef lpObject As Any) As Long
    Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytLength As Long)
    Private Declare Sub RtlFillMemory Lib "kernel32" (pDst As Any, ByVal bytLength As Long, ByVal fill As Byte)
    
    ' ----==== USER32 API Deklarationen ====----
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
    
#End If
Private m_PFN     As String
Private m_BmpHead As BITMAPFILEHEADER
Private m_BmpInfo As BITMAPINFO
Private m_PicGUID As GUID
Private m_PicDesc As PICTDESC
Private m_data()  As Byte 'the pixel-data
'private membervariables related to property Pixel:
'Private m_Stride  As Long 'Width in bytes of one picture-row including padbytes

Private Sub Class_Initialize()
    'Debug.Print LenB(m_BmpHead) '  16
    'Debug.Print LenB(m_BmpInfo) '1148
    Dim i As Integer
    With m_BmpHead
        .bfType = &H4D42
        '.bfSize = LenB(m_BmpHead)
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
'    With m_BmpHead
'        '.bfReserved1 = 0
'        '.bfReserved2 = 0
'        '.bfSize =
'        '.bfOffBits
'    End With
    Dim bpp As Byte: bpp = CByte((PixelFormat And &HFF00&) / &H100)
    Dim cuci As Integer, bm As Long, gm As Long, rm As Long, am As Long
    Dim v As Single: v = 3
    Select Case PixelFormat
    Case EPixelFormat.Format1bppIndexed:    cuci = 2
    Case EPixelFormat.Format4bppIndexed:    cuci = 16
    Case EPixelFormat.Format8bppIndexed:    cuci = 256
    Case EPixelFormat.Format16bppArgb1555:  bm = &H1F&: gm = &H3E0&: rm = &H7C00&: am = &H8000&: v = 4
    Case EPixelFormat.Format16bppGrayScale: bm = &H1F&: gm = &H3E0&: rm = &H7C00&: am = &H8000&: v = 4
    Case EPixelFormat.Format16bppRgb565:    bm = &H1F&: gm = &H7E0&: rm = &HF800&: am = &H0&:    v = 4
    Case EPixelFormat.Format16bppRgb555:    bm = &H1F6: gm = &H3E0&: rm = &H7C00&: am = &H0&:    v = 4
    Case EPixelFormat.Format24bppRgb: v = 4
    Case Else: v = 5 'bpp = 32
    'Case EPixelFormat.Format32bppArgb:    '    'bpp = 32
    'Case EPixelFormat.Format32bppPArgb:    '    'bpp = 32
    'Case EPixelFormat.Format32bppRgb:    '    'bpp = 32
    End Select
    
    With m_BmpInfo
        .biSize = IIf(v = 3, 40, IIf(v = 4, 108, 124))
        With .bmiHeader
            .biHeight = Height
            .biWidth = Width
            .biXPelsPerMeter = 3779 '= 96dpi
            .biYPelsPerMeter = 3779 '= 96dpi
            .biPlanes = 1
            .biCompression = BI_RGB
            .biBitCount = bpp
            .biClrUsed = cuci
            .biClrImportant = cuci
            'm_Stride = CalcStride(Width, bpp)
            .biSizeImage = CalcStride(Width, bpp) * Height
            ReDim m_data(0 To .biSizeImage - 1)
        End With
        If bpp = 16 Then
            With .bmiHeaderV4
                .bV4BlueMask = bm
                .bV4GreenMask = gm
                .bV4RedMask = rm
                .bV4AlphaMask = am
            End With
        End If
    End With
    With m_BmpHead
        .bfSize = LenB(m_BmpHead) + m_BmpInfo.biSize + m_BmpInfo.bmiHeader.biSizeImage
        .bfOffBits = LenB(m_BmpHead) + m_BmpInfo.biSize
    End With
    Dim i As Long, j As Long, k As Long, c As Long
    With m_BmpInfo.Palette
        Select Case bpp
        Case 1: .bmiColors(i) = New_BGRA(0, 0, 0, 0):       i = i + 1
                .bmiColors(i) = New_BGRA(255, 255, 255, 0): i = i + 1
        Case 4: .bmiColors(i) = New_BGRA(0, 0, 0, 0):       i = i + 1:    .bmiColors(i) = New_BGRA(128, 0, 0, 0):     i = i + 1
                .bmiColors(i) = New_BGRA(0, 128, 0, 0):     i = i + 1:    .bmiColors(i) = New_BGRA(128, 128, 0, 0):   i = i + 1
                .bmiColors(i) = New_BGRA(0, 0, 128, 0):     i = i + 1:    .bmiColors(i) = New_BGRA(128, 0, 128, 0):   i = i + 1
                .bmiColors(i) = New_BGRA(0, 128, 128, 0):   i = i + 1:    .bmiColors(i) = New_BGRA(192, 192, 192, 0): i = i + 1
                .bmiColors(i) = New_BGRA(128, 128, 128, 0): i = i + 1:    .bmiColors(i) = New_BGRA(255, 0, 0, 0):     i = i + 1
                .bmiColors(i) = New_BGRA(0, 255, 0, 0):     i = i + 1:    .bmiColors(i) = New_BGRA(255, 255, 0, 0):   i = i + 1
                .bmiColors(i) = New_BGRA(0, 0, 255, 0):     i = i + 1:    .bmiColors(i) = New_BGRA(255, 0, 255, 0):   i = i + 1
                .bmiColors(i) = New_BGRA(0, 255, 255, 0):   i = i + 1:    .bmiColors(i) = New_BGRA(255, 255, 255, 0): 'i = i + 1
        Case 8
                Dim R As Integer, G As Integer, b As Integer
                For i = 0 To 5
                    For j = 0 To 5
                        For k = 0 To 5
                            .bmiColors(c) = New_BGRA(CByte(R), CByte(G), CByte(b), 0)
                            c = c + 1: b = b + 51
                        Next
                        b = 0: R = R + 51
                    Next
                    b = 0: R = 0: G = G + 51
                Next
        End Select
    End With
End Sub

Friend Sub NewSP(aStdPicture As StdPicture)
    Dim hr As Long
    Dim bmpHead As BitmapHeader
    hr = GetObjectW(aStdPicture.Handle, LenB(bmpHead), bmpHead)
    If hr <> 24 Then
        MsgBox "Could not get object for handle: " & aStdPicture.Handle
        Exit Sub
    End If
    Dim bH As Long:    bH = bmpHead.bmHeight
    Dim bc As Integer: bc = bmpHead.bmBitsPixel
    With m_BmpInfo
        '.biSize = 126 'LenB(m_BmpInfo)
        With .bmiHeader
            .biWidth = bmpHead.bmWidth
            .biHeight = bH
            .biPlanes = bmpHead.bmPlanes
            .biBitCount = bc
            .biCompression = BI_RGB
            '.biSizeImage = ?
            .biXPelsPerMeter = 3779
            .biYPelsPerMeter = 3779
            '.biClrUsed
            '.biClrImportant
        End With
    End With
    'm_BmpInfo
    Dim aHDC As LongPtr: aHDC = GetDC(0&)
    Dim pBI  As LongPtr
    If bc = 1 Or bc = 4 Or bc = 8 Then
        Dim bi As BITMAPINFOV3: bi.biSize = 40
        pBI = VarPtr(bi)
        hr = GetDIBits(aHDC, aStdPicture.Handle, 0&, bH, ByVal 0&, ByVal pBI, DIBColors.DIB_RGB_COLORS)
        If hr = 0 Then
            MsgBox "Could not get DIBits"
            'GoTo Finally
        End If
        With m_BmpInfo
            .biSize = bi.biSize '40
            .bmiHeader = bi.bmiHeader
            Debug.Print bi.bmiHeader.biClrUsed ' .bmiHeader.biClrUsed
            Debug.Print bi.bmiHeader.biClrImportant '  .bmiHeader.biClrImportant
            .bmiHeader.biXPelsPerMeter = 3779
            .bmiHeader.biYPelsPerMeter = 3779
            '.Palette = bi.Palette
            'Debug.Print BGRA_ToStr(bi.Palette.bmiColors(128))
        End With
    Else
        pBI = VarPtr(m_BmpInfo)
        m_BmpInfo.biSize = 124
    End If
    Dim bu As Long: bu = bc / 8 * Abs(bmpHead.bmHeight) * bmpHead.bmWidthBytes - 1
    ReDim m_data(0 To bu)
    hr = GetDIBits(aHDC, aStdPicture.Handle, 0&, bH, m_data(0), ByVal pBI, DIBColors.DIB_RGB_COLORS)
    If bc < 16 Then
        m_BmpInfo.Palette = bi.Palette
    End If
    If hr <> bH Then
        MsgBox "Could not get di-bitmap data"
    End If
    ReleaseDC 0&, aHDC
    'm_BmpInfo As BITMAPINFO
End Sub

Friend Sub NewC(other As Bitmap)
    m_PFN = other.FileName
    '1. copy the complete Bitmap-header-stuff
    RtlMoveMemory m_BmpHead, ByVal other.HeaderPtr, 1164
    '2. copy the complete data-array
    Dim bytSize As Long: bytSize = other.DataSize
    ReDim m_data(0 To bytSize - 1)
    RtlMoveMemory m_data(0), ByVal other.DataPtr, bytSize
End Sub

Public Function Clone() As Bitmap
    Set Clone = New Bitmap
    Clone.NewC Me
End Function

Public Sub SavePixelData()
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Dim FNm As String: FNm = m_PFN & ".data"
    Open FNm For Binary Access Write As FNr
    Put FNr, , m_data
    GoTo Finally
Catch:
    MsgBox Err.Description
Finally:
    Close FNr
End Sub

Public Property Get FileName() As String
    FileName = m_PFN
End Property

Public Property Get HeaderPtr() As LongPtr
    HeaderPtr = VarPtr(m_BmpHead)
End Property
Public Property Get DataPtr() As LongPtr
    DataPtr = VarPtr(m_data(0))
End Property

Public Property Get DataSize() As Long
    DataSize = UBound(m_data) + 1
End Property

Public Property Get BitsPerPixel() As Byte
    BitsPerPixel = m_BmpInfo.bmiHeader.biBitCount
End Property
'Public Property Let BitsPerPixel(ByVal Value As Byte)
'    m_BmpInfo.bmiHeader.biBitCount = Value
'End Property

Public Property Get PixelPerMeterX() As Long
    PixelPerMeterX = m_BmpInfo.bmiHeader.biXPelsPerMeter
End Property
Public Property Let PixelPerMeterX(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biXPelsPerMeter = Value
End Property

Public Property Get PixelPerMeterY() As Long
    PixelPerMeterY = m_BmpInfo.bmiHeader.biYPelsPerMeter
End Property
Public Property Let PixelPerMeterY(ByVal Value As Long)
    m_BmpInfo.bmiHeader.biYPelsPerMeter = Value
End Property

Public Sub Read(aPFN As String)
    m_PFN = aPFN
Try: On Error GoTo Catch
    Dim sErr As String
    Dim FNr As Integer: FNr = FreeFile
    Open m_PFN For Binary Access Read As FNr
    'Debug.Print Seek(FNr) & " " & LOF(FNr)
    If Not IsBitmap(FNr) Then sErr = "Unknown FileHeader: " & m_BmpHead.bfType: GoTo Catch
    ReadBmpInfo FNr
    If ReadData(FNr) Then
        If Me.PixelFormat = EPixelFormat.Format8bppIndexed Then
            If m_BmpInfo.bmiHeader.biCompression = BI_RLE8 Then
                Dim buffer_RLEDecoded() As Byte
                If RLE8_Decode(Me.Width, Me.Height, m_data, buffer_RLEDecoded) Then
                    m_data = buffer_RLEDecoded
                End If
            End If
        End If
    Else
        sErr = "Reading pixel data failed": GoTo Catch
    End If
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
        '&H4D42 = 19778 "BM" ' Windows 3.1x, 95, NT, . . . also for OS/2
        '&H4142 = 16706 "BA" ' OS/2 struct Bitmap array
        '&H4943 = 18755 "CI" ' OS/2 struct color icon
        '&H5043 = 19778 "CP" ' OS/2 const color pointer
        '&H4349 = 20547 "IC" ' OS/2 struct icon
        '&H5450 = 21584 "PT" ' OS/2 pointer
    End With
Catch:
End Function

Private Function ReadBmpInfo(ByVal FNr As Integer) As Boolean
Try: On Error GoTo Catch
    'Debug.Print Seek(FNr) & " " & LOF(FNr)
    Dim sErr As String
    With m_BmpInfo
        Get FNr, , .biSize
        'Debug.Print .biSize
        ReadBmpInfo = .biSize = 12 Or .biSize = 40 Or .biSize = 56 Or .biSize = 108 Or .biSize = 124
        'biSize=12  => BitmapCoreHeader    -> OS/2-Bitmap
        'biSize=40  => BitmapHeader_v3     -> Windows-Bitmap
        'biSize=56  => BitmapInfoHeader_v4 -> without CSType-Info
        'biSize=108 => BitmapInfoHeader_v4 -> with CSType-Info
        'biSize=124 => BitmapInfoHeader_v5
        If Not ReadBmpInfo Then sErr = "Header-size not in the range (12, 40, 56, 108, 124): " & .biSize: GoTo Catch
        Dim i As Long, aColor As BGRA
        If .biSize = 12 Then
            Dim W As Integer: Get FNr, , W
            Dim H As Integer: Get FNr, , H
            .bmiHeader.biWidth = CLng(W)
            .bmiHeader.biHeight = CLng(H)
            Dim pl As Integer: Get FNr, , pl
            Dim bc As Integer: Get FNr, , bc
            .bmiHeader.biPlanes = pl
            .bmiHeader.biBitCount = bc
            .bmiHeader.biXPelsPerMeter = 3779
            .bmiHeader.biYPelsPerMeter = 3779
            Dim n As Long
            If bc < 24 Then
                n = 2 ^ bc
                .bmiHeader.biClrUsed = n
                .bmiHeader.biClrImportant = n
                For i = 0 To n - 1
                    Get FNr, , aColor.b
                    Get FNr, , aColor.G
                    Get FNr, , aColor.R
                    .Palette.bmiColors(i) = aColor
                Next
            'Else
                'm_BmpHead.bfOffBits = Seek(FNr) - 1
                'n = LOF(FNr) - Seek(FNr)
                'ReDim m_data(0 To n - 1)
                'Get FNr, , m_data
            End If
            
        Else
            Get FNr, , .bmiHeader
            If .biSize = 40 Then
                Dim bic As Long: bic = .bmiHeader.biBitCount
                Dim bcu As Long: bcu = .bmiHeader.biClrUsed
                If (bic = 1 Or bic = 4 Or bic = 8) Then If bcu = 0 Then .bmiHeader.biClrUsed = 2 ^ bic
                For i = 0 To .bmiHeader.biClrUsed - 1
                    Get FNr, , aColor
                    'Debug.Print BGRA_ToStr(aColor)
                    .Palette.bmiColors(i) = aColor
                Next
            ElseIf .biSize > 40 Then ' = 108
                If .biSize = 56 Then
                    With .bmiHeaderV4
                        Get FNr, , .bV4RedMask
                        Get FNr, , .bV4GreenMask
                        Get FNr, , .bV4BlueMask
                        Get FNr, , .bV4AlphaMask
                    End With
                Else
                    Get FNr, , .bmiHeaderV4
                    If .biSize > 108 Then ' = 124
                        Get FNr, , .bmiHeaderV5
                    End If
                End If
            End If
        End If
    End With
    'Debug.Print Seek(FNr) & " " & LOF(FNr)
    Exit Function
Catch: MsgBox "ReadBmpInfo " & sErr
End Function

Private Function New_BGRA(ByVal R As Byte, ByVal G As Byte, ByVal b As Byte, ByVal a As Byte) As BGRA
    With New_BGRA
        .R = R: .G = G: .b = b: .a = a
    End With
End Function

Private Function BGRA_ToStr(this As BGRA, Optional ByVal skipAlpha As Boolean = False) As String
    Dim s As String: s = "  BGR" & IIf(skipAlpha, "", "A")
    With this
        s = s & "{&H" & Hex2(.b) & Hex2(.G) & Hex2(.R) & IIf(skipAlpha, "", Hex2(.a)) & "}"
    End With
    BGRA_ToStr = s
End Function

Private Function Hex2(ByVal b As Byte) As String
    Hex2 = Hex(b): If Len(Hex2) < 2 Then Hex2 = "0" & Hex2
End Function

Private Function BGRA_ToRGBA(this As BGRA) As RGBA
    With this
        BGRA_ToRGBA.a = .a
        BGRA_ToRGBA.b = .b
        BGRA_ToRGBA.G = .G
        BGRA_ToRGBA.R = .R
    End With
End Function

Private Function RGBA_ToBGRA(this As RGBA) As BGRA
    With this
        RGBA_ToBGRA.a = .a
        RGBA_ToBGRA.b = .b
        RGBA_ToBGRA.G = .G
        RGBA_ToBGRA.R = .R
    End With
End Function

Private Function ReadData(FNr As Integer) As Boolean
Try: On Error GoTo Catch
    'Debug.Print Seek(FNr) & " " & LOF(FNr)
    Dim Offset As Long: Offset = Seek(FNr) - 1
    'der Offset muss 4 aligned sein?hmmm ist gar nicht der Fall????
    'Debug.Print "m_BmpHead.bfOffBits: " & m_BmpHead.bfOffBits & " == Offset=" & Offset
    ReadData = Offset = m_BmpHead.bfOffBits
    If Not ReadData Then GoTo Catch
    Dim u As Long: u = LOF(FNr) - Offset - 1
    ReDim m_data(0 To u) As Byte
    Get FNr, , m_data
    ReadData = True
    'Debug.Print Seek(FNr) & " " & LOF(FNr)
    Exit Function
Catch:
    MsgBox Err.Number & " " & "Error reading data" & vbCrLf & Err.Description
End Function

Private Function RLE8_Decode(ByVal W As Long, ByVal H As Long, buffer_in() As Byte, buffer_out() As Byte) As Boolean
'https://learn.microsoft.com/en-us/windows/win32/gdi/Bitmap-compression
Try: On Error GoTo Catch
    Dim Stride As Long: Stride = CalcStride(W, 8)
    'Dim Stride As Long: Stride = GetStride(W, 8)
    Dim Size_of_buffer_out As Long: Size_of_buffer_out = Stride * H
    Dim i As Long, ui As Long: ui = UBound(buffer_in)
    Dim o As Long, uo As Long: uo = Size_of_buffer_out - 1: ReDim buffer_out(0 To uo)
    Dim d As Byte: d = Stride - W
    Dim Byte1 As Byte, Byte2 As Byte, Size As Byte, fill As Byte
    Dim dRi As Byte, dUp As Byte
    Do While i <= ui
        'read 2 bytes
        Byte1 = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
        Byte2 = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
        Select Case Byte2
        Case 0: 'End of line
                'insert padbytes -> just set the output-Index
                o = o + d
        Case 1: 'End of Bitmap
                Exit Do
        Case 2: 'delta mode
            dRi = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
            dUp = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
            'jump to the position
            'todo
        Case Else
            If Byte1 = 0 Then
                'Absolute mode
                Size = Byte2
                RtlMoveMemory buffer_out(o), buffer_in(i), Size
                i = i + Size: If ui <= i Then Exit Do
                o = o + Size: If uo <= o Then Exit Do
                If IsOdd(Size) Then i = i + 1
                'do the padding
                '
            Else
                'Encoded mode
                'now fill the outputbuffer size-times with the same pixel
                Size = Byte1
                fill = Byte2
                RtlFillMemory buffer_out(o), Size, fill
                o = o + Size
            End If
        End Select
    Loop
    RLE8_Decode = True
    Exit Function
Catch:
    MsgBox Err.Description
End Function

Private Function IsOdd(ByVal n As Long) As Boolean
    'gibt zurück ob die Zahl n ungerade ist
    IsOdd = n Mod 2
End Function

'Private Function CalcOffset() As Long
'    Dim v As Single: v = Me.BitmapInfoVersion
'    Select Case v
'    Case 2:
'    Case 3:
'    Case 3.5:
'    Case 4:
'    Case 5:
'    End Select
'End Function
'
'Private Function CalcSize() As Long
'    CalcSize = CalcOffset + UBound(m_data) + 1
'End Function

Public Function Save(Optional ByVal aPFN As String = "") As Boolean
Try: On Error GoTo Catch
    Dim sErr As String
    If Len(aPFN) = 0 Then aPFN = m_PFN
    Dim FNr As Integer: FNr = FreeFile
    Open aPFN For Binary Access Write As FNr
    'now write all stuff
    'calculate filelen and biOffset
    Put FNr, , m_BmpHead '.bfType
    'Put FNr, , m_BmpHead.bfSize
    With m_BmpInfo
        Put FNr, , .biSize
        Dim pl As Long: pl = .bmiHeader.biPlanes
        Dim bc As Long: bc = .bmiHeader.biBitCount
        If .biSize < 56 And bc < 16 Then
            Dim i As Long, u As Long: u = 2 ^ bc - 1
        End If
        If 12 = .biSize Then
            If .bmiHeader.biWidth >= 65536 Then
                sErr = "For OS/2-bitmap width is too big: " & .bmiHeader.biWidth & vbCrLf & ", max width is 65535 pixel"
                GoTo Catch
            End If
            If .bmiHeader.biHeight >= 65536 Then
                sErr = "For OS/2-bitmap height is too big: " & .bmiHeader.biHeight & vbCrLf & ", max height is 65535 pixel"
                GoTo Catch
            End If
            Dim tl As TLng, ti2 As TInt2
            tl.Value = .bmiHeader.biWidth:  LSet ti2 = tl: Put FNr, , ti2.Value0
            tl.Value = .bmiHeader.biHeight: LSet ti2 = tl: Put FNr, , ti2.Value0
            Put FNr, , .bmiHeader.biPlanes    'Int16
            Put FNr, , .bmiHeader.biBitCount  'Int16
            If bc < 16 Then
                For i = 0 To u
                    Put FNr, , .Palette.bmiColors(i).b
                    Put FNr, , .Palette.bmiColors(i).G
                    Put FNr, , .Palette.bmiColors(i).R
                Next
            End If
        ElseIf 12 < .biSize Then
            With .bmiHeader
                Put FNr, , .biWidth
                Put FNr, , .biHeight
                Put FNr, , .biPlanes
                Put FNr, , .biBitCount
                Dim comp As Long: comp = .biCompression
                If comp = BI_RLE8 Then
                    MsgBox "RLE encoding is not yet implemented!"
                    comp = BI_RGB
                End If
                Put FNr, , comp
                Put FNr, , .biSizeImage
                Put FNr, , .biXPelsPerMeter
                Put FNr, , .biYPelsPerMeter
                Put FNr, , .biClrUsed
                Put FNr, , .biClrImportant
            End With
            If bc < 16 Then
                With .Palette
                    For i = 0 To u
                        Put FNr, , .bmiColors(i)
                    Next
                End With
            End If
            If 40 < .biSize Then
                With .bmiHeaderV4
                    Put FNr, , .bV4RedMask
                    Put FNr, , .bV4GreenMask
                    Put FNr, , .bV4BlueMask
                    Put FNr, , .bV4AlphaMask
                End With
                If 56 < .biSize Then
                    With .bmiHeaderV4
                        Put FNr, , .bV4CSType
                        Put FNr, , .bV4Endpoints
                        Put FNr, , .bV4GammaRed
                        Put FNr, , .bV4GammaGreen
                        Put FNr, , .bV4GammaBlue
                    End With
                    If 108 < .biSize Then
                        Put FNr, , .bmiHeaderV5
                    End If
                End If
            End If
        End If
    End With
    Put FNr, , m_data
    GoTo Finally
Catch:
    MsgBox Err.Description & vbCrLf & sErr
Finally:
    Close FNr
End Function

Friend Function ToPicture() As StdPicture
Try: On Error GoTo Catch
    Dim sErr As String, hr As Long
    With m_PicDesc
        Dim pBI As LongPtr
        'If m_BmpInfo.biSize = 40 Then
        If m_BmpInfo.biSize <= 40 Then
            Dim bi As BITMAPINFOV3
            'if m_BmpInfo.biSize = 12 then
            'bi.biSize = 40 'm_BmpInfo.biSize
            bi.biSize = 40 'm_BmpInfo.biSize
            bi.bmiHeader = m_BmpInfo.bmiHeader
            bi.bmiHeader.biCompression = BI_RGB
            bi.Palette = m_BmpInfo.Palette
            pBI = VarPtr(bi)
        Else
            pBI = VarPtr(m_BmpInfo)
        End If
        .hBitmap = CreateDIBSection(0&, ByVal pBI, 0&, 0&, 0&, 0&)
        If .hBitmap = 0 Then sErr = "hBitmap=0, CreateDIBSection failed": GoTo Catch
        hr = SetDIBits(0&, .hBitmap, 0&, Me.Height, m_data(0), ByVal pBI, 0&)
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


'| %: integer  || &: long     || @: currency || !: single   || #: double   || $: string   |
Public Property Get BitmapInfoVersion() As Single
    With m_BmpInfo
        BitmapInfoVersion = IIf(.biSize = 12, 2!, IIf(.biSize = 40, 3!, IIf(.biSize = 56, 3.5!, IIf(.biSize = 108, 4!, IIf(.biSize = 124, 5, -1)))))
    End With
End Property
Public Property Let BitmapInfoVersion(ByVal Value As Single)
    Select Case Value
    Case 2!:   m_BmpInfo.biSize = 12   'OS/2
    Case 3!:   m_BmpInfo.biSize = 40   '24bpp or indexed
    Case 3.5!: m_BmpInfo.biSize = 56   '32bpp wo cstype
    Case 4!:   m_BmpInfo.biSize = 108  '16bpp
    Case 5!:   m_BmpInfo.biSize = 124  '32,48,64bpp
    End Select
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
        Case &H4D42: s = s & IIf(Me.IsOS2, "OS/2 Bitmap", "Windows Bitmap")
        Case &H4142: s = s & "OS/2 Bitmap"
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
    Dim v As Single: v = Me.BitmapInfoVersion
    With this
        s = s & "Version       : " & v & vbCrLf
        s = s & "PixelFormat   : " & Pixelformat_ToStr & vbCrLf
        With .bmiHeader
            s = s & "Width         : " & .biWidth & vbCrLf
            s = s & "Height        : " & .biHeight & vbCrLf
            s = s & "Planes        : " & .biPlanes & vbCrLf
            s = s & "BitCount      : " & .biBitCount & vbCrLf
            s = s & "Compression   : " & Compression_ToStr(.biCompression) & vbCrLf
            s = s & "SizeImage     : " & .biSizeImage & vbCrLf
            s = s & "XPelsPerMeter : " & .biXPelsPerMeter & " = " & GetDPI(.biXPelsPerMeter) & " dpi" & vbCrLf
            s = s & "YPelsPerMeter : " & .biYPelsPerMeter & " = " & GetDPI(.biYPelsPerMeter) & " dpi" & vbCrLf
            s = s & "ColorsUsed    : " & .biClrUsed & vbCrLf
            s = s & "ClrImportant  : " & .biClrImportant & vbCrLf
        End With
        If v >= 2 Then
            Dim bOS2 As Boolean: bOS2 = Me.IsOS2
            Dim i As Long
            For i = 0 To .bmiHeader.biClrUsed - 1
                s = s & BGRA_ToStr(.Palette.bmiColors(i), bOS2) & vbCrLf
            Next
        End If
        If v > 3 Then
            With .bmiHeaderV4
                s = s & "RedMask       : &H" & Hex(.bV4RedMask) & vbCrLf
                s = s & "GreenMask     : &H" & Hex(.bV4GreenMask) & vbCrLf
                s = s & "BlueMask      : &H" & Hex(.bV4BlueMask) & vbCrLf
                s = s & "AlphaMask     : &H" & Hex(.bV4AlphaMask) & vbCrLf
            End With
        End If
        If v > 3.5 Then
            With .bmiHeaderV4
                s = s & "CSType        : " & CSType_ToStr(.bV4CSType) & vbCrLf
                s = s & "Endpoints     : " & CIEXYZTRIPLE_ToStr(.bV4Endpoints) & vbCrLf
                s = s & "GammaRed      : " & .bV4GammaRed & vbCrLf
                s = s & "GammaGreen    : " & .bV4GammaGreen & vbCrLf
                s = s & "GammaBlue     : " & .bV4GammaBlue & vbCrLf
            End With
        End If
        If v > 4 Then
            With .bmiHeaderV5
                s = s & "Intent        : " & .bV5Intent & vbCrLf
                s = s & "ProfileData   : " & .bV5ProfileData & vbCrLf
                s = s & "ProfileSize   : " & .bV5ProfileSize & vbCrLf
                s = s & "Reserved      : " & .bV5Reserved & vbCrLf
            End With
        End If
    End With
    BITMAPINFO_ToStr = s
End Function

Private Function Compression_ToStr(ByVal biComp As Long) As String
    Dim s As String: s = "&H" & Hex(biComp) & " : "
    Select Case biComp
    'BI_-constants
    Case BI_RGB:       s = s & "BI_RGB"       ' = 0&  ' An uncompressed format.
    Case BI_RLE8:      s = s & "BI_RLE8"      ' = 1&  ' A run-length encoded (RLE) format for Bitmaps with 8 bpp. The compression format is a 2-byte format consisting of a count byte followed by a byte containing a color index. For more information, see Bitmap Compression.
    Case BI_RLE4:      s = s & "BI_RLE4"      ' = 2&  ' An RLE format for Bitmaps with 4 bpp. The compression format is a 2-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
    Case BI_BITFIELDS: s = s & "BI_BITFIELDS" ' = 3&  ' Specifies that the Bitmap is not compressed. The members bV4RedMask, bV4GreenMask, and bV4BlueMask specify the red, green, and blue components for each pixel. This is valid when used with 16- and 32-bpp Bitmaps
                                          '       ' If biCompression = BI_BITFIELDS -> the format is either RGB-555 or RGB-565.
                                          '       ' Use the subtype GUID in the AM_MEDIA_TYPE structure to determine the specific RGB type.
                                          '       ' https://learn.microsoft.com/en-us/windows/win32/api/strmif/ns-strmif-am_media_type
                                        
    Case BI_JPEG:      s = s & "BI_JPEG"      ' = 4&  ' Specifies that the image is compressed using the JPEG file interchange format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
    Case BI_PNG:       s = s & "BI_PNG"       ' = 5&  ' Specifies that the image is compressed using the PNG  file interchange format.
    Case BI_1632:      s = s & "BI_1632"      ' = &H32333631 ' "1632"
    Case &H56595559:   s = s & "YUYV video"   ' = &H56595559 ' "YUYV" video has the FOURCC 'VYUY' or 0x56595559
    Case Else:         s = s & "unkown"
    End Select
    Compression_ToStr = s
End Function

Private Function CSType_ToStr(ByVal CSTyp As Long) As String
    Dim s As String: s = "&H" & Hex(CSTyp) & " : "
    Select Case CSTyp
    Case 0:          s = s & "CALIBRATED_RGB"
    Case &H73524742: s = s & "sRGB"
    Case &H206E6957: s = s & "WINDOWS_COLOR_SPACE" ' &H206E6957 = "Win "
    End Select
    CSType_ToStr = s
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

Private Function Pixelformat_ToStr() As String
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
    Pixelformat_ToStr = s
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
                'how to distinguish the 16bpp greyscale format?
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
    Case 32:
                If m_BmpInfo.bmiHeaderV4.bV4AlphaMask = &H0& Then
                    epf = EPixelFormat.Format32bppRgb
                Else
                    epf = EPixelFormat.Format32bppArgb
                End If
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

Public Property Get Stride() As Long
    Stride = CalcStride(Me.Width, Me.BitsPerPixel)
End Property

Private Function CalcStride(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
    Dim BytesPerPixel As Single: BytesPerPixel = BitsPerPixel / 8
    Select Case BitsPerPixel
    Case 1, 4:      CalcStride = W * BytesPerPixel
                    Dim m As Long: m = CalcStride Mod 4: If m > 0 Then m = 4 - m
                    CalcStride = CalcStride + m
    'Case 4:
    Case 8:         CalcStride = (W + 3) And Not 3
    Case 16, 24:    CalcStride = ((W * BytesPerPixel) + BytesPerPixel) And Not BytesPerPixel
    Case 32:        CalcStride = W * BytesPerPixel
    End Select
End Function

'Private Function CalcStride(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
'    'calculates the number of bytes in one horizontal line of pixels,
'    'including the number of pad-bytes for the 4-aligned result
'    CalcStride = W * BitsPerPixel / 8
'    Dim m As Long: m = CalcStride Mod 4: If m > 0 Then m = 4 - m
'    CalcStride = CalcStride + m
'End Function
'Private Function GetStrideFS() As Long
'    Dim W As Long: W = Me.Width
'    Dim BytesPerPixel As Long: BytesPerPixel = m_bpp / 8
'    Select Case m_bpp
'    Case 8:        GetStrideFS = (W + 3) And Not 3
'    Case 16, 24:   GetStrideFS = ((W * BytesPerPixel) + BytesPerPixel) And Not BytesPerPixel
'    Case 32:       GetStrideFS = W * BytesPerPixel
'    End Select
'End Function

'Private m_bpp As Byte
'Private m_Width As Integer
'
'Private Sub TestStride()
'    m_bpp = 8
'    m_Width = 100
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 101
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 102
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 103
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 104
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_bpp = 16
'    m_Width = 100
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 101
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 102
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 103
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 104
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_bpp = 24
'    m_Width = 100
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 101
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 102
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 103
'    Debug.Print GetStrideFS & " " & GetStrideOM
'    m_Width = 104
'    Debug.Print GetStrideFS & " " & GetStrideOM
'End Sub
'
'Public Function GetStrideFS() As Long 'Anzahl an Bytes pro Zeile
'    Dim BytesPerPixel As Long: BytesPerPixel = m_bpp / 8
'    Select Case m_bpp
'    Case 8:        GetStrideFS = (m_Width + 3) And Not 3
'    Case 16, 24:   GetStrideFS = ((m_Width * BytesPerPixel) + BytesPerPixel) And Not BytesPerPixel
'    Case 32:       GetStrideFS = m_Width * BytesPerPixel
'    End Select
'End Function
'
'Public Function GetStrideOM() As Long 'Anzahl an Bytes pro Zeile inkl Padbytes
'    GetStrideOM = m_Width * m_bpp / 8
'    Dim m As Byte: m = GetStrideOM Mod 4: If m > 0 Then m = 4 - m
'    GetStrideOM = GetStrideOM + m
'End Function
'



Public Property Get Height() As Long
    Height = Abs(m_BmpInfo.bmiHeader.biHeight)
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

Public Property Get IsOS2() As Boolean
    Dim T As Long: T = m_BmpHead.bfType
    IsOS2 = (T = &H4142&) Or (T = &H4943&) Or (T = &H5043&) Or (T = &H4349&) Or (T = &H5450&) Or ((T = &H4D42&) And (m_BmpInfo.biSize < 40))
    'If T = &H4D42& Then
    '    'If BM then only if bfSize is less than 40
    '    If m_BmpHead.bfSize < 40 Then IsOS2 = True
    'End If
    '&H4D42 = 19778 "BM" ' Windows 3.1x, 95, NT, . . . also for OS/2
    '&H4142 = 16706 "BA" ' OS/2 struct Bitmap array
    '&H4943 = 18755 "CI" ' OS/2 struct color icon
    '&H5043 = 19778 "CP" ' OS/2 const color pointer
    '&H4349 = 20547 "IC" ' OS/2 struct icon
    '&H5450 = 21584 "PT" ' OS/2 pointer
End Property

Public Property Get IsIndexed() As Boolean
    IsIndexed = Me.PaletteCount > 0 'm_BmpInfo.bmiHeader.biClrUsed > 0
End Property

Public Property Get PaletteCount() As Long
    PaletteCount = m_BmpInfo.bmiHeader.biClrUsed
End Property

Public Property Get PaletteColor(ByVal Index As Long) As Long
    If Index < 0 Or Me.PaletteCount <= Index Then Exit Property
    Dim BGRA As BGRA: BGRA = m_BmpInfo.Palette.bmiColors(Index)
    Dim RGBA As RGBA: RGBA = BGRA_ToRGBA(BGRA)
    Dim c As TLng: LSet c = RGBA 'm_BmpInfo.Palette.bmiColors(Index)
    PaletteColor = c.Value
End Property

Public Property Let PaletteColor(ByVal Index As Long, ByVal Value As Long)
    If Index < 0 Or Me.PaletteCount <= Index Then Exit Property
    Dim c As TLng: c.Value = Value
    Dim RGBA As RGBA: LSet RGBA = c
    Dim BGRA As BGRA: BGRA = RGBA_ToBGRA(RGBA)
    m_BmpInfo.Palette.bmiColors(Index) = BGRA
End Property

Public Property Get PalettePixelIndex(ByVal x As Long, ByVal y As Long) As Long
    If Not Me.IsIndexed Then Exit Property
    Dim W As Long: W = Me.Width:     If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:    If y < 0 Or H <= y Then Exit Property
    Select Case Me.PixelFormat
    Case EPixelFormat.Format1bppIndexed:    PalettePixelIndex = PalettePixelIndex1(x, y)
    Case EPixelFormat.Format4bppIndexed:    PalettePixelIndex = PalettePixelIndex4(x, y)
    Case EPixelFormat.Format8bppIndexed:    PalettePixelIndex = PalettePixelIndex8(x, y)
    End Select
End Property

Private Property Get PalettePixelIndex1(ByVal x As Long, ByVal y As Long) As Long
    'n.y.i.
End Property

Private Property Get PalettePixelIndex4(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = m_BmpInfo.bmiHeader.biWidth
    'Dim S As Long: S = Me.Stride
    Dim H As Long: H = m_BmpInfo.bmiHeader.biHeight
    Dim si As Single: si = ((H - 1 - y) * W + x) / 2
    'Dim si As Single: si = ((H - 1 - Y) * S + X) / 2
    Dim i  As Long:    i = Int(si)
    If si = CSng(i) Then
        'the lower nibble
        i = &HF And m_data(i)
    Else
        'the higher nibble
        i = &HF0 And m_data(i)
        i = i / &H10
    End If
    PalettePixelIndex4 = i
End Property

Private Property Get PalettePixelIndex8(ByVal x As Long, ByVal y As Long) As Long
    'Dim W As Long: W = m_BmpInfo.bmiHeader.biWidth
    Dim s As Long: s = Me.Stride
    Dim H As Long: H = m_BmpInfo.bmiHeader.biHeight
    'Dim i As Long: i = (H - 1 - Y) * W + X: PixelPaletteIndex8 = m_data(i)
    Dim i As Long: i = (H - 1 - y) * s + x: PalettePixelIndex8 = m_data(i)
End Property

Public Property Get Pixel(ByVal x As Long, ByVal y As Long) As Long
    Select Case Me.PixelFormat
    Case EPixelFormat.Format1bppIndexed:    Pixel = Me.Pixel1(x, y)
    Case EPixelFormat.Format4bppIndexed:    Pixel = Me.Pixel4(x, y)
    Case EPixelFormat.Format8bppIndexed:    Pixel = Me.Pixel8(x, y)
    Case EPixelFormat.Format16bppArgb1555:  Pixel = Me.Pixel16Argb1555(x, y)
    Case EPixelFormat.Format16bppGrayScale: Pixel = Me.Pixel16Grayscale(x, y)
    Case EPixelFormat.Format16bppRgb555:    Pixel = Me.Pixel16Rgb555(x, y)
    Case EPixelFormat.Format16bppRgb565:    Pixel = Me.Pixel16Rgb565(x, y)
    Case EPixelFormat.Format24bppRgb:       Pixel = Me.Pixel24(x, y)
    Case EPixelFormat.Format32bppArgb, EPixelFormat.Format32bppPArgb, EPixelFormat.Format32bppRgb
                                            Pixel = Me.Pixel32(x, y)
    End Select
End Property

Public Property Let Pixel(ByVal x As Long, ByVal y As Long, ByVal Color As Long)
    Select Case Me.PixelFormat
    Case EPixelFormat.Format1bppIndexed:    Pixel = Me.Pixel1(x, y) = Color
    Case EPixelFormat.Format4bppIndexed:    Pixel = Me.Pixel4(x, y) = Color
    Case EPixelFormat.Format8bppIndexed:    Pixel = Me.Pixel8(x, y) = Color
    Case EPixelFormat.Format16bppArgb1555:  Pixel = Me.Pixel16Argb1555(x, y) = Color
    Case EPixelFormat.Format16bppGrayScale: Pixel = Me.Pixel16Grayscale(x, y) = Color
    Case EPixelFormat.Format16bppRgb555:    Pixel = Me.Pixel16Rgb555(x, y) = Color
    Case EPixelFormat.Format16bppRgb565:    Pixel = Me.Pixel16Rgb565(x, y) = Color
    Case EPixelFormat.Format24bppRgb:       Pixel = Me.Pixel24(x, y) = Color
    Case EPixelFormat.Format32bppArgb, EPixelFormat.Format32bppPArgb, EPixelFormat.Format32bppRgb
                                            Pixel = Me.Pixel32(x, y) = Color
    End Select
        
    Dim l As TLng: l.Value = Color
    Dim c As RGBA: LSet c = l
    With c
        m_data(y * m_Width + x) = (.R + .G + .b) \ 3
    End With
End Property
'
Public Property Get Pixel1(ByVal x As Long, ByVal y As Long) As Long
    Pixel1 = 0
End Property
'Public Property Let Pixel1(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property
'
Public Property Get Pixel4(ByVal x As Long, ByVal y As Long) As Long
    If Not Me.IsIndexed Then Exit Property
    Dim W As Long: W = Me.Width:     If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:    If y < 0 Or H <= y Then Exit Property
    'Dim si As Single: si = ((H - 1 - Y) * W + X) / 2
    'Dim i  As Long:    i = Int(si)
    Dim i  As Long:    i = PalettePixelIndex4(x, y)
    'Debug.Print UBound(m_data)
    'If si = CSng(i) Then
    '    'the lower nibble
    '    i = &HF And m_data(i)
    'Else
    '    'the higher nibble
    '    i = &HF0 And m_data(i)
    '    i = i / &H10
    'End If
    Dim BGRA As BGRA: BGRA = m_BmpInfo.Palette.bmiColors(i)
    Dim RGBA As RGBA: RGBA = BGRA_ToRGBA(BGRA)
    Dim l As TLng: LSet l = RGBA
    Pixel4 = l.Value
End Property
'Public Property Let Pixel4(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel8(ByVal x As Long, ByVal y As Long) As Long
    If Not Me.IsIndexed Then Exit Property
    Dim W As Long: W = Me.Width:     If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:    If y < 0 Or H <= y Then Exit Property
    Dim s As Long: s = Me.Stride
    'Dim i As Long: i = (H - 1 - Y) * W + X: i = m_data(i)
    Dim i As Long: i = (H - 1 - y) * s + x: i = m_data(i)
    Dim BGRA As BGRA: BGRA = m_BmpInfo.Palette.bmiColors(i)
    Dim RGBA As RGBA: RGBA = BGRA_ToRGBA(BGRA)
    Dim l As TLng: LSet l = RGBA
    Pixel8 = l.Value
    'Pixel8 = RGB(BGRA.R, BGRA.G, BGRA.B)
End Property

'Public Property Let Pixel8(ByVal X As Long, ByVal Y As Long, ByVal Value As Long)
'    'm_data=0
'End Property

'Public Property Get Pixel16(ByVal X As Long, ByVal Y As Long) As Long
'    Dim W As Long: W = Me.Width:     If X < 0 Or W <= X Then Exit Property
'    Dim H As Long: H = Me.Height:    If Y < 0 Or H <= Y Then Exit Property
'    Dim s As Long: s = Me.Stride
'    'Dim i As Long: i = ((H - 1 - Y) * W + X) * 2
'    Dim i As Long: i = ((H - 1 - Y) * s + X * 2) ' * 2
'    'Dim b1 As Byte: b1 = m_data(i)
'    'Dim b2 As Byte: b2 = m_data(i + 1)
'    Dim tb As TByt2
'    tb.Value1 = m_data(i)
'    tb.Value0 = m_data(i + 1)
'    Dim ti As TInt: LSet ti = tb
'    'now we have to extract the bits and expand it to 255
'    'Dim r As Byte, g As Byte, b As Byte, a As Byte
'    'B = b2 And &H1F * 255 \ &H1F
'    'Dim bb As BGRA
'    'bb.B = b2 And &H1F * 255 \ &H1F
'
'    'bb.G = b2
'    'Dim L As TLng: LSet L = bb
'    'l.Value=rgb(
'    'Pixel16 = L.Value
'    'Pixel16=
'End Property
''Public Property Let Pixel16(ByVal X As Long, ByVal Y As Long, Value As Long)
''    'm_data=0
''End Property

Private Function Color16bppBGRA1555_ToColor32(ByVal Value As Long) As Long
'     BlueMask = &H1F    ' 5 bit
'    GreenMask = &H3E0   ' 5 bit
'      RedMask = &H7C00  ' 5 bit
'    AlphaMask = &H8000  ' 1 bit
    Dim R As Long, G As Long, b As Long, a As Long
    b = ((Value And &H1F&) * 256) \ &H1F&
    G = (((Value And &H3E0&) \ &H20&) * 256) \ &H1F&
    R = (((Value And &H7C00&) \ &H400&) * 256) \ &H1F&
    a = (((Value And &H8000&) \ &H8000&) * 256) ' alpha is only 1 bit, so it is 0 or 1, resp 0 or 255
    Color16bppBGRA1555_ToColor32 = RGB(R, G, b)
End Function

Private Function Color16bppBGR565_ToColor32(ByVal Value As Long) As Long
'     BlueMask = &H1F    ' 5 bit
'    GreenMask = &H7E0   ' 6 bit
'      RedMask = &HF800  ' 5 bit
'    AlphaMask = &H0     ' ----
    Dim R As Long, G As Long, b As Long, a As Long
    b = ((Value And &H1F&) * 256) \ &H1F&
    G = (((Value And &H7E0&) \ &H20&) * 256) \ &H3F&
    R = (((Value And &HF800&) \ &H800&) * 256) \ &H1F&
    'a = ' alpha is not included
    Color16bppBGR565_ToColor32 = RGB(R, G, b)
End Function

Public Property Get Pixel16Argb1555(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = Me.Width:    If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:   If y < 0 Or H <= y Then Exit Property
    'Dim BitsPerPixel  As Long: BitsPerPixel = 16
    'Dim BitsPerByte   As Long: BitsPerByte = 8
    Dim BytesPerPixel As Long: BytesPerPixel = 2 'BitsPerPixel / 8 ' 2
    Dim s As Long: s = Me.Stride
    Dim i As Long: i = (H - 1 - y) * s + x * BytesPerPixel
    Dim tb As TByt4: tb.Value0 = m_data(i): tb.Value1 = m_data(i + 1)
    Dim tl As TLng: LSet tl = tb
    Pixel16Argb1555 = Color16bppBGRA1555_ToColor32(tl.Value)
End Property

Public Property Get Pixel16Grayscale(ByVal x As Long, ByVal y As Long) As Long
    '
End Property

Public Property Get Pixel16Rgb555(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = Me.Width:    If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:   If y < 0 Or H <= y Then Exit Property
    'Dim BitsPerPixel  As Long: BitsPerPixel = 16
    'Dim BitsPerByte   As Long: BitsPerByte = 8
    Dim BytesPerPixel As Long: BytesPerPixel = 2 'BitsPerPixel / 8 ' 2
    Dim s As Long: s = Me.Stride
    Dim i As Long: i = (H - 1 - y) * s + x * BytesPerPixel
    Dim tb As TByt4: tb.Value0 = m_data(i): tb.Value1 = m_data(i + 1)
    Dim tl As TLng: LSet tl = tb
    Pixel16Rgb555 = Color16bppBGRA1555_ToColor32(tl.Value)
End Property

Public Property Get Pixel16Rgb565(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = Me.Width:    If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:   If y < 0 Or H <= y Then Exit Property
    'Dim BitsPerPixel  As Long: BitsPerPixel = 16
    'Dim BitsPerByte   As Long: BitsPerByte = 8
    Dim BytesPerPixel As Long: BytesPerPixel = 2 'BitsPerPixel / 8 ' 2
    Dim s As Long: s = Me.Stride
    Dim i As Long: i = (H - 1 - y) * s + x * BytesPerPixel
    Dim tb As TByt4: tb.Value0 = m_data(i): tb.Value1 = m_data(i + 1)
    Dim tl As TLng: LSet tl = tb
    Pixel16Rgb565 = Color16bppBGR565_ToColor32(tl.Value)
End Property


Public Property Get Pixel24(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = Me.Width:    If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:   If y < 0 Or H <= y Then Exit Property
    'W = W + (W Mod 4)
    'Dim BitsPerPixel  As Long: BitsPerPixel = 24
    'Dim BitsPerByte   As Long: BitsPerByte = 8
    Dim BytesPerPixel As Long: BytesPerPixel = 3 'BitsPerPixel / 8 'BitsPerByte ' 3
    'Dim W1 As Long: W1 = W * BytesPerPixel
    Dim s As Long: s = Me.Stride
    Dim i As Long: i = (H - 1 - y) * s + x * BytesPerPixel
    'Dim BGRA As BGRA
    Dim RGBA As RGBA
    With RGBA
        .R = m_data(i + 2)
        .G = m_data(i + 1)
        .b = m_data(i + 0)
    End With
    Dim l As TLng: LSet l = RGBA 'BGRA
    Pixel24 = l.Value
    'Pixel24 = RGB(RGBA.R, RGBA.G, RGBA.B)
    'Pixel24 = RGB(m_data(i + 2), m_data(i + 1), m_data(i))
End Property
'Public Property Let Pixel24(ByVal X As Long, ByVal Y As Long, Value As Long)
'    'm_data=0
'End Property

Public Property Get Pixel32(ByVal x As Long, ByVal y As Long) As Long
    Dim W As Long: W = Me.Width:    If x < 0 Or W <= x Then Exit Property
    Dim H As Long: H = Me.Height:   If y < 0 Or H <= y Then Exit Property
    'Dim S As Long: S = Me.Stride 'nope not needed here
    Dim i As Long: i = ((H - 1 - y) * W + x) * 4
    Dim RGBA As RGBA
    'Debug.Print UBound(m_data)
    With RGBA
        '.A = m_data(i + 3)
        .R = m_data(i + 2)
        .G = m_data(i + 1)
        .b = m_data(i + 0)
    End With
    Dim l As TLng: LSet l = RGBA
    Pixel32 = l.Value
End Property
Public Property Let Pixel32(ByVal x As Long, ByVal y As Long, Value As Long)
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
