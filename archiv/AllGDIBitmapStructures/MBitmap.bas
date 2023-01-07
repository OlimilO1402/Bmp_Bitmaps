Attribute VB_Name = "MBitmap"
Option Explicit

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmap
'typedef struct tagBITMAP {
'  LONG   bmType;
'  LONG   bmWidth;
'  LONG   bmHeight;
'  LONG   bmWidthBytes;
'  WORD   bmPlanes;
'  WORD   bmBitsPixel;
'  LPVOID bmBits;
'} BITMAP, *PBITMAP, *NPBITMAP, *LPBITMAP;
Public Type TBITMAP
    bmType       As Long    ' The bitmap type. This member must be zero.
    bmWidth      As Long    ' The width, in pixels, of the bitmap. The width must be greater than zero.
    bmHeight     As Long    ' The height, in pixels, of the bitmap. The height must be greater than zero.
    bmWidthBytes As Long    ' The number of bytes in each scan line. This value must be divisible by 2, because the system assumes that the bit values of a bitmap form an array that is word aligned.
    bmPlanes     As Integer ' The count of color planes.
    bmBitsPixel  As Integer ' The number of bits required to indicate the color of a pixel.
    bmBits       As LongPtr ' A pointer to the location of the bit values for the bitmap. The bmBits member must be a pointer to an array of character (1-byte) values.
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapcoreheader
'typedef struct tagBITMAPCOREHEADER {
'  DWORD bcSize;
'  WORD  bcWidth;
'  WORD  bcHeight;
'  WORD  bcPlanes;
'  WORD  bcBitCount;
'} BITMAPCOREHEADER, *LPBITMAPCOREHEADER, *PBITMAPCOREHEADER;
Public Type BITMAPCOREHEADER
    bcSize       As Long    ' The number of bytes required by the structure.
    bcWidth      As Integer ' The width of the bitmap, in pixels.
    bcHeight     As Integer ' The height of the bitmap, in pixels.
    bcPlanes     As Integer ' The number of planes for the target device. This value must be 1.
    bcBitCount   As Integer ' The number of bits-per-pixel. This value must be 1, 4, 8, or 24.
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-rgbtriple
'typedef struct tagRGBTRIPLE {
'  BYTE rgbtBlue;
'  BYTE rgbtGreen;
'  BYTE rgbtRed;
'} RGBTRIPLE, *PRGBTRIPLE, *NPRGBTRIPLE, *LPRGBTRIPLE;
Public Type RGBTRIPLE
    rgbtBlue  As Byte ' The intensity of blue in the color.
    rgbtGreen As Byte ' The intensity of green in the color.
    rgbtRed   As Byte ' The intensity of red in the color.
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapcoreinfo
'The BITMAPCOREINFO structure combines the BITMAPCOREHEADER structure and a color table to provide a complete definition of the dimensions and colors of a DIB. For more information about specifying a DIB, see BITMAPCOREINFO.
'An application should use the information stored in the bcSize member to locate the color table in a BITMAPCOREINFO structure, using a method such as the following:
'pColor = ((LPBYTE) pBitmapCoreInfo + (WORD) (pBitmapCoreInfo -> bcSize))
'typedef struct tagBITMAPCOREINFO {
'  BITMAPCOREHEADER bmciHeader;
'  RGBTRIPLE        bmciColors[1];
'} BITMAPCOREINFO, *LPBITMAPCOREINFO, *PBITMAPCOREINFO;
Public Type BITMAPCOREINFO
    bmciHeader    As BITMAPCOREHEADER ' A BITMAPCOREHEADER structure that contains information about the dimensions and color format of a DIB.
    bmciColors(1) As RGBTRIPLE        ' Specifies an array of RGBTRIPLE structures that define the colors in the bitmap.
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapfileheader
'typedef struct tagBITMAPFILEHEADER {
'  WORD  bfType;
'  DWORD bfSize;
'  WORD  bfReserved1;
'  WORD  bfReserved2;
'  DWORD bfOffBits;
'} BITMAPFILEHEADER, *LPBITMAPFILEHEADER, *PBITMAPFILEHEADER;
Public Type BITMAPFILEHEADER
    bfType      As Integer ' The file type; must be BM.
    bfSize      As Long    ' The size, in bytes, of the bitmap file.
    bfReserved1 As Integer ' Reserved; must be zero.
    bfReserved2 As Integer ' Reserved; must be zero.
    bfOffBits   As Long    ' The offset, in bytes, from the beginning of the BITMAPFILEHEADER structure to the bitmap bits.
End Type
'A BITMAPINFO or BITMAPCOREINFO structure immediately follows the BITMAPFILEHEADER structure in the DIB file. For more information, see Bitmap Storage.

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-rgbquad
'typedef struct tagRGBQUAD {
'  BYTE rgbBlue;
'  BYTE rgbGreen;
'  BYTE rgbRed;
'  BYTE rgbReserved;
'} RGBQUAD;
Public Type RGBQUAD
    rgbBlue     As Byte ' The intensity of blue in the color.
    rgbGreen    As Byte ' The intensity of green in the color.
    rgbRed      As Byte ' The intensity of red in the color.
    rgbReserved As Byte ' This member is reserved and must be zero.
End Type
'The bmiColors member of the BITMAPINFO structure consists of an array of RGBQUAD structures.

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfo
'typedef struct tagBITMAPINFO {
'  BITMAPINFOHEADER bmiHeader;
'  RGBQUAD          bmiColors[1];
'} BITMAPINFO, *LPBITMAPINFO, *PBITMAPINFO;
Public Type BITMAPINFO
    bmiHeader    As BITMAPINFOHEADER
    bmiColors(1) As RGBQUAD
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapinfoheader
'typedef struct tagBITMAPINFOHEADER {
'  DWORD biSize;
'  LONG  biWidth;
'  LONG  biHeight;
'  WORD  biPlanes;
'  WORD  biBitCount;
'  DWORD biCompression;
'  DWORD biSizeImage;
'  LONG  biXPelsPerMeter;
'  LONG  biYPelsPerMeter;
'  DWORD biClrUsed;
'  DWORD biClrImportant;
'} BITMAPINFOHEADER, *LPBITMAPINFOHEADER, *PBITMAPINFOHEADER;
Public Type BITMAPINFOHEADER
    biSize          As Long    ' Specifies the number of bytes required by the structure. This value does not include the size of the color table or the size of the color masks, if they are appended to the end of structure. See Remarks.
    biWidth         As Long    ' Specifies the width of the bitmap, in pixels. For information about calculating the stride of the bitmap, see Remarks.
    biHeight        As Long    ' Specifies the height of the bitmap, in pixels.
                               ' * For uncompressed RGB bitmaps, if biHeight is positive, the bitmap is a bottom-up DIB with the origin at the lower left corner. If biHeight is negative, the bitmap is a top-down DIB with the origin at the upper left corner.
                               ' * For YUV bitmaps, the bitmap is always top-down, regardless of the sign of biHeight. Decoders should offer YUV formats with positive biHeight, but for backward compatibility they should accept YUV formats with either positive or negative biHeight.
                               ' * For compressed formats, biHeight must be positive, regardless of image orientation.
    
    biPlanes        As Integer ' Specifies the number of planes for the target device. This value must be set to 1.
    biBitCount      As Integer ' Specifies the number of bits per pixel (bpp). For uncompressed formats, this value is the average number of bits per pixel. For compressed formats, this value is the implied bit depth of the uncompressed image, after the image has been decoded.
    biCompression   As Long    ' For compressed video and YUV formats, this member is a FOURCC code, specified as a DWORD in little-endian order. For example, YUYV video has the FOURCC 'VYUY' or 0x56595559. For more information, see FOURCC Codes.
                               ' For uncompressed RGB formats, the following values are possible:
                               ' Value         Meaning
                               ' BI_RGB        Uncompressed RGB.
                               ' BI_BITFIELDS  Uncompressed RGB with color masks. Valid for 16-bpp and 32-bpp bitmaps.
                               ' See Remarks for more information. Note that BI_JPG and BI_PNG are not valid video formats.
                               ' For 16-bpp bitmaps, if biCompression equals BI_RGB, the format is always RGB 555.
                               ' If biCompression equals BI_BITFIELDS, the format is either RGB 555 or RGB 565.
                               ' Use the subtype GUID in the AM_MEDIA_TYPE structure to determine the specific RGB type.
                               
    biSizeImage     As Long    ' Specifies the size, in bytes, of the image. This can be set to 0 for uncompressed RGB bitmaps.
    biXPelsPerMeter As Long    ' Specifies the horizontal resolution, in pixels per meter, of the target device for the bitmap.
    biYPelsPerMeter As Long    ' Specifies the vertical resolution, in pixels per meter, of the target device for the bitmap.
    biClrUsed       As Long    ' Specifies the number of color indices in the color table that are actually used by the bitmap. See Remarks for more information.
    biClrImportant  As Long    ' Specifies the number of color indices that are considered important for displaying the bitmap. If this value is zero, all colors are important.
End Type
