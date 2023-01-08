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

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapv4header
'typedef struct {
'  DWORD        bV4Size;
'  LONG         bV4Width;
'  LONG         bV4Height;
'  WORD         bV4Planes;
'  WORD         bV4BitCount;
'  DWORD        bV4V4Compression;
'  DWORD        bV4SizeImage;
'  LONG         bV4XPelsPerMeter;
'  LONG         bV4YPelsPerMeter;
'  DWORD        bV4ClrUsed;
'  DWORD        bV4ClrImportant;
'  DWORD        bV4RedMask;
'  DWORD        bV4GreenMask;
'  DWORD        bV4BlueMask;
'  DWORD        bV4AlphaMask;
'  DWORD        bV4CSType;
'  CIEXYZTRIPLE bV4Endpoints;
'  DWORD        bV4GammaRed;
'  DWORD        bV4GammaGreen;
'  DWORD        bV4GammaBlue;
'} BITMAPV4HEADER, *LPBITMAPV4HEADER, *PBITMAPV4HEADER;
Public Type BITMAPV4HEADER
    ': v same datatypes as above v :
    bV4Size          As Long    ' The number of bytes required by the structure. Applications should use this member to determine which bitmap information header structure is being used.
    bV4Width         As Long    ' The width of the bitmap, in pixels. If bV4Compression is BI_JPEG or BI_PNG, bV4Width specifies the width of the JPEG or PNG image in pixels.
    bV4Height        As Long    ' The height of the bitmap, in pixels. If bV4Height is positive, the bitmap is a bottom-up DIB and its origin is the lower-left corner. If bV4Height is negative, the bitmap is a top-down DIB and its origin is the upper-left corner.
                                ' If bV4Height is negative, indicating a top-down DIB, bV4Compression must be either BI_RGB or BI_BITFIELDS. Top-down DIBs cannot be compressed. If bV4Compression is BI_JPEG or BI_PNG, bV4Height specifies the height of the JPEG or PNG image in pixels.
    bV4Planes        As Integer ' The number of planes for the target device. This value must be set to 1.
    bV4BitCount      As Integer ' The number of bits-per-pixel. The bV4BitCount member of the BITMAPV4HEADER structure determines the number of bits that define each pixel and the maximum number of colors in the bitmap. This member must be one of the following values.
'Value Meaning
' 0    The number of bits-per-pixel is specified or is implied by the JPEG or PNG file format.
' 1    The bitmap is monochrome, and the bmiColors member of BITMAPINFO contains two entries. Each bit in the bitmap array represents a pixel. If the bit is clear, the pixel is displayed with the color of the first entry in the bmiColors table; if the bit is set, the pixel has the color of the second entry in the table.
' 4    The bitmap has a maximum of 16 colors, and the bmiColors member of BITMAPINFO contains up to 16 entries. Each pixel in the bitmap is represented by a 4-bit index into the color table. For example, if the first byte in the bitmap is 0x1F, the byte represents two pixels. The first pixel contains the color in the second table entry, and the second pixel contains the color in the sixteenth table entry.
' 8    The bitmap has a maximum of 256 colors, and the bmiColors member of BITMAPINFO contains up to 256 entries. In this case, each byte in the array represents a single pixel.
'16    The bitmap has a maximum of 2^16 colors. If the bV4Compression member of the BITMAPV4HEADER structure is BI_RGB, the bmiColors member of BITMAPINFO is NULL. Each WORD in the bitmap array represents a single pixel. The relative intensities of red, green, and blue are represented with five bits for each color component. The value for blue is in the least significant five bits, followed by five bits each for green and red, respectively. The most significant bit is not used. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV4ClrUsed member of the BITMAPV4HEADER.If the bV4Compression member of the BITMAPV4HEADER is BI_BITFIELDS, the bmiColors member contains three DWORD color masks that specify the red, green, and blue components of each pixel. Each WORD in the bitmap array represents a single pixel.
'24    The bitmap has a maximum of 2^24 colors, and the bmiColors member of BITMAPINFO is NULL. Each 3-byte triplet in the bitmap array represents the relative intensities of blue, green, and red for a pixel. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV4ClrUsed member of the BITMAPV4HEADER.
'32    The bitmap has a maximum of 2^32 colors. If the bV4Compression member of the BITMAPV4HEADER is BI_RGB, the bmiColors member of BITMAPINFO is NULL. Each DWORD in the bitmap array represents the relative intensities of blue, green, and red for a pixel. The value for blue is in the least significant 8 bits, followed by 8 bits each for green and red. The high byte in each DWORD is not used. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV4ClrUsed member of the BITMAPV4HEADER.If the bV4Compression member of the BITMAPV4HEADER is BI_BITFIELDS, the bmiColors member contains three DWORD color masks that specify the red, green, and blue components of each pixel. Each DWORD in the bitmap array represents a single pixel.
    
    bV4V4Compression As Long    ' The type of compression for a compressed bottom-up bitmap (top-down DIBs cannot be compressed). This member can be one of the following values.
'Value         Description
'BI_RGB        An uncompressed format.
'BI_RLE8       A run-length encoded (RLE) format for bitmaps with 8 bpp. The compression format is a 2-byte format consisting of a count byte followed by a byte containing a color index. For more information, see Bitmap Compression.
'BI_RLE4       An RLE format for bitmaps with 4 bpp. The compression format is a 2-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
'BI_BITFIELDS  Specifies that the bitmap is not compressed. The members bV4RedMask, bV4GreenMask, and bV4BlueMask specify the red, green, and blue components for each pixel. This is valid when used with 16- and 32-bpp bitmaps.
'BI_JPEG       Specifies that the image is compressed using the JPEG file interchange format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
'BI_PNG        Specifies that the image is compressed using the PNG file interchange format.
    
    bV4SizeImage     As Long    ' The size, in bytes, of the image. This may be set to zero for BI_RGB bitmaps. If bV4Compression is BI_JPEG or BI_PNG, bV4SizeImage is the size of the JPEG or PNG image buffer.
    bV4XPelsPerMeter As Long    ' The horizontal resolution, in pixels-per-meter, of the target device for the bitmap. An application can use this value to select a bitmap from a resource group that best matches the characteristics of the current device.
    bV4YPelsPerMeter As Long    ' The vertical resolution, in pixels-per-meter, of the target device for the bitmap.
    bV4ClrUsed       As Long    ' The number of color indexes in the color table that are actually used by the bitmap. If this value is zero, the bitmap uses the maximum number of colors corresponding to the value of the bV4BitCount member for the compression mode specified by bV4Compression. If bV4ClrUsed is nonzero and the bV4BitCount member is less than 16, the bV4ClrUsed member specifies the actual number of colors the graphics engine or device driver accesses. If bV4BitCount is 16 or greater, the bV4ClrUsed member specifies the size of the color table used to optimize performance of the system color palettes. If bV4BitCount equals 16 or 32, the optimal color palette starts immediately following the BITMAPV4HEADER.
    bV4ClrImportant  As Long    ' The number of color indexes that are required for displaying the bitmap. If this value is zero, all colors are important.
    ': ^ same datatypes as above ^ :
    bV4RedMask       As Long    ' Color mask that specifies the red component of each pixel, valid only if bV4Compression is set to BI_BITFIELDS.
    bV4GreenMask     As Long    ' Color mask that specifies the green component of each pixel, valid only if bV4Compression is set to BI_BITFIELDS.
    bV4BlueMask      As Long    ' Color mask that specifies the blue component of each pixel, valid only if bV4Compression is set to BI_BITFIELDS.
    bV4AlphaMask     As Long    ' Color mask that specifies the alpha component of each pixel.
    bV4CSType        As Long    ' The color space of the DIB. The following table lists the value for bV4CSType.
'Value               Meaning
'LCS_CALIBRATED_RGB  This value indicates that endpoints and gamma values are given in the appropriate fields.
'See the LOGCOLORSPACE structure for information that defines a logical color space.
'Value                     Meaning
'LCS_CALIBRATED_RGB        Color values are calibrated RGB values. The values are translated using the endpoints specified by the lcsEndpoints member before being passed to the device.
'LCS_sRGB                  Color values are values are sRGB values.
'LCS_WINDOWS_COLOR_SPACE   Color values are Windows default color space color values.

    bV4Endpoints     As CIEXYZTRIPLE ' A CIEXYZTRIPLE structure that specifies the x, y, and z coordinates of the three colors that correspond to the red, green, and blue endpoints for the logical color space associated with the bitmap. This member is ignored unless the bV4CSType member specifies LCS_CALIBRATED_RGB.
                                     ' Note  A color space is a model for representing color numerically in terms of three or more coordinates. For example, the RGB color space represents colors in terms of the red, green, and blue coordinates.
    bV4GammaRed      As Long    ' Tone response curve for red. This member is ignored unless color values are calibrated RGB values and bV4CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
    bV4GammaGreen    As Long    ' Tone response curve for green. Used if bV4CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
    bV4GammaBlue     As Long    ' Tone response curve for blue. Used if bV4CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
End Type
'Remarks: The BITMAPV4HEADER structure is extended to allow a JPEG or PNG image to be passed as the source image to StretchDIBits.

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-bitmapv5header
'typedef struct {
'  DWORD        bV5Size;
'  LONG         bV5Width;
'  LONG         bV5Height;
'  WORD         bV5Planes;
'  WORD         bV5BitCount;
'  DWORD        bV5Compression;
'  DWORD        bV5SizeImage;
'  LONG         bV5XPelsPerMeter;
'  LONG         bV5YPelsPerMeter;
'  DWORD        bV5ClrUsed;
'  DWORD        bV5ClrImportant;
'  DWORD        bV5RedMask;
'  DWORD        bV5GreenMask;
'  DWORD        bV5BlueMask;
'  DWORD        bV5AlphaMask;
'  DWORD        bV5CSType;
'  CIEXYZTRIPLE bV5Endpoints;
'  DWORD        bV5GammaRed;
'  DWORD        bV5GammaGreen;
'  DWORD        bV5GammaBlue;
'  DWORD        bV5Intent;
'  DWORD        bV5ProfileData;
'  DWORD        bV5ProfileSize;
'  DWORD        bV5Reserved;
'} BITMAPV5HEADER, *LPBITMAPV5HEADER, *PBITMAPV5HEADER;

Public Type BITMAPV5HEADER
 ': v same datatypes as above v :
    bV5Size          As Long    ' The number of bytes required by the structure. Applications should use this member to determine which bitmap information header structure is being used.
    bV5Width         As Long    ' The width of the bitmap, in pixels. If bV5Compression is BI_JPEG or BI_PNG, the bV5Width member specifies the width of the decompressed JPEG or PNG image in pixels.
    bV5Height        As Long    ' The height of the bitmap, in pixels. If the value of bV5Height is positive, the bitmap is a bottom-up DIB and its origin is the lower-left corner. If bV5Height value is negative, the bitmap is a top-down DIB and its origin is the upper-left corner. If bV5Height is negative, indicating a top-down DIB, bV5Compression must be either BI_RGB or BI_BITFIELDS. Top-down DIBs cannot be compressed. If bV5Compression is BI_JPEG or BI_PNG, the bV5Height member specifies the height of the decompressed JPEG or PNG image in pixels.
    bV5Planes        As Integer ' The number of planes for the target device. This value must be set to 1.
    bV5BitCount      As Integer ' The number of bits that define each pixel and the maximum number of colors in the bitmap. This member can be one of the following values.
'Value Meaning
' 0    The number of bits per pixel is specified or is implied by the JPEG or PNG file format.
' 1    The bitmap is monochrome, and the bmiColors member of BITMAPINFO contains two entries. Each bit in the bitmap array represents a pixel. If the bit is clear, the pixel is displayed with the color of the first entry in the bmiColors color table. If the bit is set, the pixel has the color of the second entry in the table.
' 4    The bitmap has a maximum of 16 colors, and the bmiColors member of BITMAPINFO contains up to 16 entries. Each pixel in the bitmap is represented by a 4-bit index into the color table. For example, if the first byte in the bitmap is 0x1F, the byte represents two pixels. The first pixel contains the color in the second table entry, and the second pixel contains the color in the sixteenth table entry.
' 8    The bitmap has a maximum of 256 colors, and the bmiColors member of BITMAPINFO contains up to 256 entries. In this case, each byte in the array represents a single pixel.
'16    The bitmap has a maximum of 2^16 colors. If the bV5Compression member of the BITMAPV5HEADER structure is BI_RGB, the bmiColors member of BITMAPINFO is NULL. Each WORD in the bitmap array represents a single pixel. The relative intensities of red, green, and blue are represented with five bits for each color component. The value for blue is in the least significant five bits, followed by five bits each for green and red. The most significant bit is not used. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV5ClrUsed member of the BITMAPV5HEADER.If the bV5Compression member of the BITMAPV5HEADER is BI_BITFIELDS, the bmiColors member contains three DWORD color masks that specify the red, green, and blue components, respectively, of each pixel. Each WORD in the bitmap array represents a single pixel.
'      When the bV5Compression member is BI_BITFIELDS, bits set in each DWORD mask must be contiguous and should not overlap the bits of another mask. All the bits in the pixel do not need to be used.
'24    The bitmap has a maximum of 2^24 colors, and the bmiColors member of BITMAPINFO is NULL. Each 3-byte triplet in the bitmap array represents the relative intensities of blue, green, and red, respectively, for a pixel. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV5ClrUsed member of the BITMAPV5HEADER structure.
'32    The bitmap has a maximum of 2^32 colors. If the bV5Compression member of the BITMAPV5HEADER is BI_RGB, the bmiColors member of BITMAPINFO is NULL. Each DWORD in the bitmap array represents the relative intensities of blue, green, and red for a pixel. The value for blue is in the least significant 8 bits, followed by 8 bits each for green and red. The high byte in each DWORD is not used. The bmiColors color table is used for optimizing colors used on palette-based devices, and must contain the number of entries specified by the bV5ClrUsed member of the BITMAPV5HEADER.If the bV5Compression member of the BITMAPV5HEADER is BI_BITFIELDS, the bmiColors member contains three DWORD color masks that specify the red, green, and blue components of each pixel. Each DWORD in the bitmap array represents a single pixel.
    
    bV5Compression   As Long    'Specifies that the bitmap is not compressed. The bV5RedMask, bV5GreenMask, and bV5BlueMask members specify the red, green, and blue components of each pixel. This is valid when used with 16- and 32-bpp bitmaps. This member can be one of the following values.
'Value         Meaning
'BI_RGB        An uncompressed format.
'BI_RLE8       A run-length encoded (RLE) format for bitmaps with 8 bpp. The compression format is a two-byte format consisting of a count byte followed by a byte containing a color index. If bV5Compression is BI_RGB and the bV5BitCount member is 16, 24, or 32, the bitmap array specifies the actual intensities of blue, green, and red rather than using color table indexes. For more information, see Bitmap Compression.
'BI_RLE4       An RLE format for bitmaps with 4 bpp. The compression format is a two-byte format consisting of a count byte followed by two word-length color indexes. For more information, see Bitmap Compression.
'BI_BITFIELDS  Specifies that the bitmap is not compressed and that the color masks for the red, green, and blue components of each pixel are specified in the bV5RedMask, bV5GreenMask, and bV5BlueMask members. This is valid when used with 16- and 32-bpp bitmaps.
'BI_JPEG       Specifies that the image is compressed using the JPEG file Interchange Format. JPEG compression trades off compression against loss; it can achieve a compression ratio of 20:1 with little noticeable loss.
'BI_PNG        Specifies that the image is compressed using the PNG file Interchange Format.
    
    bV5SizeImage     As Long    ' The size, in bytes, of the image. This may be set to zero for BI_RGB bitmaps. If bV5Compression is BI_JPEG or BI_PNG, bV5SizeImage is the size of the JPEG or PNG image buffer.
    bV5XPelsPerMeter As Long    ' The horizontal resolution, in pixels-per-meter, of the target device for the bitmap. An application can use this value to select a bitmap from a resource group that best matches the characteristics of the current device.
    bV5YPelsPerMeter As Long    ' The vertical resolution, in pixels-per-meter, of the target device for the bitmap.
    bV5ClrUsed       As Long    ' The number of color indexes in the color table that are actually used by the bitmap. If this value is zero, the bitmap uses the maximum number of colors corresponding to the value of the bV5BitCount member for the compression mode specified by bV5Compression.
                                ' If bV5ClrUsed is nonzero and bV5BitCount is less than 16, the bV5ClrUsed member specifies the actual number of colors the graphics engine or device driver accesses. If bV5BitCount is 16 or greater, the bV5ClrUsed member specifies the size of the color table used to optimize performance of the system color palettes. If bV5BitCount equals 16 or 32, the optimal color palette starts immediately following the BITMAPV5HEADER. If bV5ClrUsed is nonzero, the color table is used on palettized devices, and bV5ClrUsed specifies the number of entries.
    bV5ClrImportant  As Long    ' The number of color indexes that are required for displaying the bitmap. If this value is zero, all colors are required.
    bV5RedMask       As Long    ' Color mask that specifies the red component of each pixel, valid only if bV5Compression is set to BI_BITFIELDS.
    bV5GreenMask     As Long    ' Color mask that specifies the green component of each pixel, valid only if bV5Compression is set to BI_BITFIELDS.
    bV5BlueMask      As Long    ' Color mask that specifies the blue component of each pixel, valid only if bV5Compression is set to BI_BITFIELDS.
    bV5AlphaMask     As Long    ' Color mask that specifies the alpha component of each pixel.
    bV5CSType        As Long    ' The color space of the DIB. The following table specifies the values for bV5CSType.
'Value                    Meaning
'LCS_CALIBRATED_RGB       This value implies that endpoints and gamma values are given in the appropriate fields.
'LCS_sRGB                 Specifies that the bitmap is in sRGB color space.
'LCS_WINDOWS_COLOR_SPACE  This value indicates that the bitmap is in the system default color space, sRGB.
'PROFILE_LINKED           This value indicates that bV5ProfileData points to the file name of the profile to use (gamma and endpoints values are ignored).
'PROFILE_EMBEDDED         This value indicates that bV5ProfileData points to a memory buffer that contains the profile to be used (gamma and endpoints values are ignored).
    
    bV5Endpoints     As CIEXYZTRIPLE 'A CIEXYZTRIPLE structure that specifies the x, y, and z coordinates of the three colors that correspond to the red, green, and blue endpoints for the logical color space associated with the bitmap. This member is ignored unless the bV5CSType member specifies LCS_CALIBRATED_RGB.
    bV5GammaRed      As Long    ' Toned response curve for red. Used if bV5CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
    bV5GammaGreen    As Long    ' Toned response curve for green. Used if bV5CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
    bV5GammaBlue     As Long    ' Toned response curve for blue. Used if bV5CSType is set to LCS_CALIBRATED_RGB. Specify in unsigned fixed 16.16 format. The upper 16 bits are the unsigned integer value. The lower 16 bits are the fractional part.
 ': ^ same datatypes as above ^ :
    
    bV5Intent        As Long    ' Rendering intent for bitmap. This can be one of the following values.
'Value                       Intent   ICC name                Meaning
'LCS_GM_ABS_COLORIMETRIC     Match    Absolute Colorimetric   Maintains the white point. Matches the colors to their nearest color in the destination gamut.
'LCS_GM_BUSINESS             Graphic  Saturation              Maintains saturation. Used for business charts and other situations in which undithered colors are required.
'LCS_GM_GRAPHICS             Proof    Relative Colorimetric   Maintains colorimetric match. Used for graphic designs and named colors.
'LCS_GM_IMAGES               Picture  Perceptual              Maintains contrast. Used for photographs and natural images.
    
    bV5ProfileData   As Long    ' The offset, in bytes, from the beginning of the BITMAPV5HEADER structure to the start of the profile data. If the profile is embedded, profile data is the actual profile, and it is linked. (The profile data is the null-terminated file name of the profile.) This cannot be a Unicode string. It must be composed exclusively of characters from the Windows character set (code page 1252). These profile members are ignored unless the bV5CSType member specifies PROFILE_LINKED or PROFILE_EMBEDDED.
    bV5ProfileSize   As Long    ' Size, in bytes, of embedded profile data.
    bV5Reserved      As Long    ' This member has been reserved. Its value should be set to zero.
End Type
'Remarks
'If bV5Height is negative, indicating a top-down DIB, bV5Compression must be either BI_RGB or BI_BITFIELDS. Top-down DIBs cannot be compressed.
'The Independent Color Management interface (ICM) 2.0 allows International Color Consortium (ICC) color profiles to be linked or embedded in DIBs (DIBs). See Using Structures for more information.
'When a DIB is loaded into memory, the profile data (if present) should follow the color table, and the bV5ProfileData should provide the offset of the profile data from the beginning of the BITMAPV5HEADER structure. The value stored in bV5ProfileData will be different from the value returned by the sizeof operator given the BITMAPV5HEADER argument, because bV5ProfileData is the offset in bytes from the beginning of the BITMAPV5HEADER structure to the start of the profile data. (Bitmap bits do not follow the color table in memory). Applications should modify the bV5ProfileData member after loading the DIB into memory.
'For packed DIBs, the profile data should follow the bitmap bits similar to the file format. The bV5ProfileData member should still give the offset of the profile data from the beginning of the BITMAPV5HEADER.
'Applications should access the profile data only when bV5Size equals the size of the BITMAPV5HEADER and bV5CSType equals PROFILE_EMBEDDED or PROFILE_LINKED.
'If a profile is linked, the path of the profile can be any fully qualified name (including a network path) that can be opened using the CreateFile function.


'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logcolorspacea
'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/ns-wingdi-logcolorspacew
'typedef struct tagLOGCOLORSPACEA {
'  DWORD         lcsSignature;
'  DWORD         lcsVersion;
'  DWORD         lcsSize;
'  LCSCSTYPE     lcsCSType;
'  LCSGAMUTMATCH lcsIntent;
'  CIEXYZTRIPLE  lcsEndpoints;
'  DWORD         lcsGammaRed;
'  DWORD         lcsGammaGreen;
'  DWORD         lcsGammaBlue;
'  CHAR          lcsFilename[MAX_PATH];
'} LOGCOLORSPACEA, *LPLOGCOLORSPACEA;
Public Type LOGCOLORSPACEA
    lcsSignature  As Long          ' Color space signature. At present, this member should always be set to LCS_SIGNATURE.
    lcsVersion    As Long          ' Version number; must be 0x400.
    lcsSize       As Long          ' Size of this structure, in bytes.
    LCSCSTYPE     As LCSCSTYPE     ' Color space type. The member can be one of the following values.
'Value                    Meaning
'LCS_CALIBRATED_RGB       Color values are calibrated RGB values. The values are translated using the endpoints specified by the lcsEndpoints member before being passed to the device.
'LCS_sRGB                 Color values are values are sRGB values.
'LCS_WINDOWS_COLOR_SPACE  Color values are Windows default color space color values.

    lcsIntent     As LCSGAMUTMATCH ' The gamut mapping method. This member can be one of the following values.
'Value                       Intent   ICC name                Meaning
'LCS_GM_ABS_COLORIMETRIC     Match    Absolute Colorimetric   Maintains the white point. Matches the colors to their nearest color in the destination gamut.
'LCS_GM_BUSINESS             Graphic  Saturation              Maintains saturation. Used for business charts and other situations in which undithered colors are required.
'LCS_GM_GRAPHICS             Proof    Relative Colorimetric   Maintains colorimetric match. Used for graphic designs and named colors.
'LCS_GM_IMAGES               Picture  Perceptual              Maintains contrast. Used for photographs and natural images.
    
    lcsEndpoints  As CIEXYZTRIPLE  ' Red, green, blue endpoints.
    lcsGammaRed   As Long          ' Scale of the red coordinate.
    lcsGammaGreen As Long          ' Scale of the green coordinate.
    lcsGammaBlue  As Long          ' Scale of the blue coordinate.
    lcsFilename(MAX_PATH) As Byte  'or WCHAR / VB.Integer for the W-version
                                   'A null-terminated string that names a color profile file. This member is typically set to zero, but may be used to set the color space to be exactly as specified by the color profile. This is useful for devices that input color values for a specific printer, or when using an installable image color matcher. If a color profile is specified, all other members of this structure should be set to reasonable values, even if the values are not completely accurate.
End Type
'Remarks
'Like palettes, but unlike pens and brushes, a pointer must be passed when creating a LogColorSpace.
'If the lcsCSType member is set to LCS_sRGB or LCS_WINDOWS_COLOR_SPACE, the other members of this structure are ignored, and WCS uses the sRGB color space. The lcsEndpoints,lcsGammaRed, lcsGammaGreen, and lcsGammaBlue members are used to describe the logical color space. The lcsEndpoints member is a CIEXYZTRIPLE that contains the x, y, and z values of the color space's RGB endpoint.
'The required DWORD bit format for the lcsGammaRed, lcsGammaGreen, and lcsGammaBlue is an 8.8 fixed point integer left-shifted by 8 bits. This means 8 integer bits are followed by 8 fraction bits. Taking the bit shift into account, the required format of the 32-bit DWORD is:
'0       nnnnnnnnffffffff00000000
'Whenever the lcsFilename member contains a file name and the lcsCSType member is set to LCS_CALIBRATED_RGB, WCS ignores the other members of this structure. It uses the color space in the file as the color space to which this LOGCOLORSPACE structure refers.
'The relation between tri-stimulus values X,Y,Z and chromaticity values x,y,z is as follows:
'X = X / (X + Y + Z)
'Y = Y / (X + Y + Z)
'Z = Z / (X + Y + Z)
'If the lcsCSType member is set to LCS_sRGB or LCS_WINDOWS_COLOR_SPACE, the other members of this structure are ignored, and ICM uses the sRGB color space. Applications should still initialize the rest of the structure since CreateProfileFromLogColorSpace ignores lcsCSType member and uses lcsEndpoints, lcsGammaRed, lcsGammaGreen, lcsGammaBlue members to create a profile, which may not be initialized in case of LCS_sRGB or LCS_WINDOWS_COLOR_SPACE color spaces.
