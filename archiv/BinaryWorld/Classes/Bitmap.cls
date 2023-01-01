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

' refer to the MSDN for more detailed information regarding the
' constants used in this sample.
Const IMAGE_BITMAP = &O0    ' used with LoadImage to load a bitmap
Const LR_LOADFROMFILE = 16    ' used with LoadImage
Const LR_CREATEDIBSECTION = 8192    ' used with LoadImage

'/////////////////////////////////////////////////////////////////////

' constants used in this example are declared here
Const SRCAND = &H8800C6    ' used to determine how a blit will  turn out
Const SRCCOPY = &HCC0020    ' used to determine how a blit will turn out
Const SRCERASE = &H440328    ' used to determine how a blit will turn out
Const SRCINVERT = &H660046    ' used to determine how a blit will turn out
Const SRCPAINT = &HEE0086    ' used to determine how a blit will turn out

Const DIB_RGB_COLORS    As Long = 0

Private Const BI_JPEG = 4&
Private Const BI_PNG = 5&
Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&

'' structures used in this example
'Private Type BITMAP
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type

Private Type BITMAPFILEHEADER    '14 bytes
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER   '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO_256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

' API's used in this example
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDIBits256 Lib "gdi32" Alias "GetDIBits" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO_256, ByVal wUsage As Long) As Long
Private Declare Function GetObject Lib "GDI32.dll" Alias "GetObjectA" ( _
        ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Const CLR_INVALID = &HFFFF    ' used to test results of color
' API calls

Private m_hImageDC As Long    ' DC to the Image bitmap
Private m_hbmImageOld As Long    'used to remove bitmaps

Private iWidth As Long    ' Size of Bitmap - X
Private iHeight As Long    ' Size of Bitmap - Y

'Private OpenBitmapHeader As BITMAPINFOHEADER
Private SaveBitmapHeader As BITMAPINFOHEADER

Private m_tBI As BITMAPINFO

Private m_lPtrBitmap As Long    '//Pointer to source bitmap returned by CreateDIBSection

'///////////////////////////////////////////////////////
'//Original BMP properties
'///////////////////////////////////////////////////////
Property Get ImageDC() As Long
    ImageDC = m_hImageDC    ' return the DC that contains the regular image
End Property

Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth    ' return the width of the bitmap
End Property

Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight    ' return the height of the bitmap
End Property

Property Get BitCount() As Long
    BitCount = m_tBI.bmiHeader.biBitCount
End Property

Property Get Size() As Long
    Size = m_tBI.bmiHeader.biSizeImage
End Property

Property Get Compression() As Long
    Compression = m_tBI.bmiHeader.biCompression
End Property

'///////////////////////////////////////////////////////
'//Modified BMP properties
'///////////////////////////////////////////////////////
Property Get SaveBMPWidth() As Long
    SaveBMPWidth = SaveBitmapHeader.biWidth    ' return the width of the bitmap
End Property

Property Get SaveBMPHeight() As Long
    SaveBMPHeight = SaveBitmapHeader.biHeight    ' return the height of the bitmap
End Property

Property Get SaveBMPBitCount() As Long
    SaveBMPBitCount = SaveBitmapHeader.biBitCount
End Property

Property Get SaveBMPSize() As Long
    SaveBMPSize = SaveBitmapHeader.biSizeImage
End Property

Property Get SaveBMPCompression() As Long
    SaveBMPCompression = SaveBitmapHeader.biCompression
End Property

Private Sub Class_Terminate()
    CleanUp
End Sub
Private Sub CleanUp()
    Dim lresult As Long

    lresult = DeleteObject(SelectObject(m_hImageDC, m_hbmImageOld))

    ' When this object is destroyed, destroy the Device Contexts that
    ' are being used.
    Call DeleteDC(ImageDC)    ' delete the image DC
End Sub

'//Returns number of bytes required for Image data (excluding headers)
Private Function ImageBufferSize(Optional BitsPerPix As Integer = 0) As Long
    If BitsPerPix = 0 Then BitsPerPix = m_tBI.bmiHeader.biBitCount
     ImageBufferSize = (((m_tBI.bmiHeader.biWidth * BitsPerPix) + 31) / 32) * 4 * m_tBI.bmiHeader.biHeight
End Function

'//Create a blank Bitmap
Public Function LoadBMP(bmpWidth As Integer, bmpHeight As Integer) As Boolean
    Dim hBmp As Long
        
    CleanUp    '//Clenup all previous resources
    
    ' Create the device context.
    m_hImageDC = CreateCompatibleDC(0)

    ' Create the bitmap.
    hBmp = CreateCompatibleBitmap(GetDC(0), bmpWidth, bmpHeight)
    If hBmp = 0 Then Exit Function
    
    ' Make the device context use the bitmap.
    m_hbmImageOld = SelectObject(m_hImageDC, hBmp)
    
    If (GetObject(hBmp, Len(m_tBI.bmiHeader), m_tBI.bmiHeader) = 0) Then Exit Function
    m_tBI.bmiHeader.biBitCount = 24
        
    LoadBMP = True
End Function
Public Function LoadBMPFromStdPic(Pic As StdPicture) As Boolean
    Dim hBmp As Long
    Dim lresult As Long    ' lResults of our API calls

    CleanUp    '//Clenup all previous resources
    
    hBmp = Pic.handle
    
    ''''''''''''''''''''''''''''''''''''''''''''''''
    '  open file and get some info
    ''''''''''''''''''''''''''''''''''''''''''''''''
    If (GetObject(hBmp, Len(m_tBI.bmiHeader), m_tBI.bmiHeader) = 0) Then Exit Function
    '//Picture object has bitcount =0
    m_tBI.bmiHeader.biBitCount = 24

    m_hImageDC = CreateCompatibleDC(0)
    ' Make sure that the Memory Dc is created.
    If (m_hImageDC = 0) Then Exit Function
    
    ' Save the actual bitmap within this bitmap object.
    m_hbmImageOld = SelectObject(m_hImageDC, hBmp)
    ' Make sure that the object is selected.
    If (m_hbmImageOld = 0) Then Exit Function
        
    m_tBI.bmiHeader = m_tBI.bmiHeader
    
    LoadBMPFromStdPic = True

End Function
Public Function LoadBMPFromFile(OpenFileName As String) As Boolean
    Dim hBmp As Long
    Dim lresult As Long    ' lResults of our API calls

    CleanUp    '//Clenup all previous resources

    LoadBMPFromFile = True

    ' load the bitmap into memory
    hBmp = LoadImage(0, OpenFileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

    ' make sure the call succeeded
    If (hBmp = 0) Then
        MsgBox "Error, Unable To Load Bitmap", vbOKOnly, "Bitmap Load Error"
        LoadBMPFromFile = False
        Exit Function
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''
    '  open file and get some info
    ''''''''''''''''''''''''''''''''''''''''''''''''
    Dim fNum As Long

    fNum = FreeFile
    Open OpenFileName For Binary As fNum
    Get fNum, 15, m_tBI.bmiHeader    '//First 14 bytes are for BMP file header
                                     '//so skip that and start at 15th byte
    Close fNum

    m_hImageDC = CreateCompatibleDC(0)

    ' Save the actual bitmap within this bitmap object.
    m_hbmImageOld = SelectObject(m_hImageDC, hBmp)

    ' Make sure that the object is selected.
    If (m_hbmImageOld = 0) Then LoadBMPFromFile = False
    
    lresult = DeleteObject(hBmp)
    ' Make sure that the temporary bitmap is deleted.
    If (lresult = 0) Then LoadBMPFromFile = False

    m_tBI.bmiHeader = m_tBI.bmiHeader
End Function

Public Function SaveBMP(ByVal SaveFileName As String, BitsPerPix As Integer) As Boolean
    Dim SaveBitmapInfo As BITMAPINFO_256
    Dim SaveFileHeader As BITMAPFILEHEADER

    Dim SaveBits() As Byte
    Dim BufferSize As Long
    Dim hNewDc As Long, hNewBm As Long, hOldBm As Long
    Dim fNum As Integer
    Dim RetVal As Long

    SaveBMP = True

    '''''''''''''''''''''''''''''''''
    'create compatible DC and bitmap
    '''''''''''''''''''''''''''''''''
    hNewDc = CreateCompatibleDC(Me.ImageDC)
    hNewBm = CreateCompatibleBitmap(Me.ImageDC, m_tBI.bmiHeader.biWidth, m_tBI.bmiHeader.biHeight)
    hOldBm = SelectObject(hNewDc, hNewBm)

    RetVal = BitBlt(hNewDc, 0, 0, m_tBI.bmiHeader.biWidth, m_tBI.bmiHeader.biHeight, Me.ImageDC, 0, 0, SRCCOPY)

    If (RetVal = 0) Or (hOldBm = 0) Or (hNewBm = 0) Or (hNewDc = 0) Then
        SaveBMP = False
        Exit Function
    End If

    '''''''''''''''''''''''''''''''
    'size buffer to hold DIB bits
    '''''''''''''''''''''''''''''''
    'BufferSize = ((m_tBI.bmiHeader.biWidth + 3) And &HFFFC) * m_tBI.bmiHeader.biHeight
    BufferSize = ImageBufferSize(BitsPerPix)

    ReDim SaveBits(0 To BufferSize - 1)

    ''''''''''''''''''''''''''''''
    'create headers
    ''''''''''''''''''''''''''''''
    With SaveFileHeader
        .bfType = &H4D42
        .bfSize = Len(SaveFileHeader) + Len(SaveBitmapInfo) + BufferSize
        .bfOffBits = Len(SaveFileHeader) + Len(SaveBitmapInfo)
    End With

    With SaveBitmapInfo.bmiHeader
        .biSize = 40
        .biWidth = m_tBI.bmiHeader.biWidth
        .biHeight = m_tBI.bmiHeader.biHeight
        .biPlanes = 1
        .biBitCount = BitsPerPix
        '.biCompression = m_tBI.bmiHeader.biCompression
        '.biClrUsed = m_tBI.bmiHeader.biClrUsed
        .biSizeImage = BufferSize
    End With

    ''''''''''''''''''''''''''''''''''''''''''''''''
    'get actual DIB
    ''''''''''''''''''''''''''''''
    hNewBm = SelectObject(hNewDc, hOldBm)
    If (RetVal = 0) Then
        SaveBMP = False
        Exit Function
    End If

    '//Get Byte array of save bitmap,
    '//GetDIBits256 also fills color array for 8,4,1 BPP bitmap defined as RGBQUAD array
    RetVal = GetDIBits256(Me.ImageDC, hNewBm, 0, m_tBI.bmiHeader.biHeight, SaveBits(0), SaveBitmapInfo, DIB_RGB_COLORS)

    If (RetVal = 0) Then
        SaveBMP = False
        Exit Function
    End If

    RetVal = DeleteObject(hNewBm)
    RetVal = DeleteDC(hNewDc)

    ''''''''''''''''''''''''''''''''''''''''''''''''
    'save it to disk
    ''''''''''''''''''''''''''''''
    fNum = FreeFile

    Open SaveFileName For Binary As fNum
    Put fNum, , SaveFileHeader    'Write Fileheader
    Put fNum, , SaveBitmapInfo    'Write BMPheader
    Put fNum, , SaveBits()    'Write Actual Image Bytes
    Close fNum

    SaveBitmapHeader = SaveBitmapInfo.bmiHeader
End Function

'//Paint picture to destination dc
Public Sub PaintPicture( _
        ByVal hDestDC As Long, _
        Optional ByVal lDestLeft As Long = 0, _
        Optional ByVal lDestTop As Long = 0, _
        Optional ByVal lDestWidth As Long = -1, _
        Optional ByVal lDestHeight As Long = -1, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt hDestDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hImageDC, lSrcLeft, lSrcTop, eRop
End Sub

'//Paint picture from source
Public Sub LoadPictureBlt( _
        ByVal hSrcDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hImageDC, 0, 0, lSrcWidth, lSrcHeight, hSrcDC, lSrcLeft, lSrcTop, eRop
End Sub
