Attribute VB_Name = "Module1"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'
'
'Option Explicit
'
'' refer to the MSDN for more detailed information regarding the
'' constants used in this sample.
'Public Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
'Public Const LR_LOADFROMFILE = 16    ' used with LoadImage
'Public Const LR_CREATEDIBSECTION = 8192    ' used with LoadImage
'
'' constants used in this example are declared here
'Public Const SRCAND = &H8800C6    ' used to determine how a blit will  turn out
'Public Const SRCCOPY = &HCC0020    ' used to determine how a blit will turn out
'Public Const SRCERASE = &H440328    ' used to determine how a blit will turn out
'Public Const SRCINVERT = &H660046    ' used to determine how a blit will turn out
'Public Const SRCPAINT = &HEE0086    ' used to determine how a blit will turn out
'
'Public Const DIB_RGB_COLORS    As Long = 0
'
'Private Const BI_JPEG = 4&
'Private Const BI_PNG = 5&
'Private Const BI_RGB = 0&
'Private Const BI_RLE4 = 2&
'Private Const BI_RLE8 = 1&
'
'' structures used in this example
'Type BITMAP
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type
'
'Type BITMAPFILEHEADER    '14 bytes
'   bfType As Integer
'   bfSize As Long
'   bfReserved1 As Integer
'   bfReserved2 As Integer
'   bfOffBits As Long
'End Type
'
'Type BITMAPINFOHEADER   '40 bytes
'   biSize As Long
'   biWidth As Long
'   biHeight As Long
'   biPlanes As Integer
'   biBitCount As Integer
'   biCompression As Long
'   biSizeImage As Long
'   biXPelsPerMeter As Long
'   biYPelsPerMeter As Long
'   biClrUsed As Long
'   biClrImportant As Long
'End Type
'
'Type RGBQUAD
'   rgbBlue As Byte
'   rgbGreen As Byte
'   rgbRed As Byte
'   rgbReserved As Byte
'End Type
'
'Type BITMAPINFO_256
'   bmiHeader As BITMAPINFOHEADER
'   bmiColors(255) As RGBQUAD
'End Type
'
'' API's used in this example
'Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
'Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_256, ByVal wUsage As Long) As Long

