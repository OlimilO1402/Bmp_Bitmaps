VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DIBSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER '40 bytes
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

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" ( _
        ByVal hdc As Long) As Long

Private Declare Function GetDC Lib "user32" ( _
        ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
        (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, _
        lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Declare Function BitBlt Lib "gdi32" ( _
        ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, _
        ByVal hSrcDC As Long, ByVal xSrc As Long, _
        ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" ( _
        ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" ( _
        ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" ( _
        ByVal hdc As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
        ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, _
        ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0    '  color table in RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Const CF_BITMAP = 2
Private Const CF_DIB = 8

' Handle to the current DIBSection:
Private m_hDIb As Long
' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long
' Handle to the Device context holding the DIBSection:
Private m_hDC As Long
' Address of memory pointing to the DIBSection's bits:
Private m_lPtr As Long
' Type containing the Bitmap information:
Private m_tBI As BITMAPINFO

'//for bit by bit operation with DIB

Dim bDib() As Byte    '//2 dimentional dynamic array for Bitmap
Dim tSA As SAFEARRAY2D
Dim m_xBound As Long, m_yBound As Long    '//Pre calculated values for pixel array dimentions

Dim bTable(0 To 100, 0 To 255) As Byte    'Brightness table
Dim cTable(0 To 100, 0 To 255) As Byte    'Contrast table
Dim gTable(0 To 100, 0 To 255) As Byte    'Gamma table

Public Function CreateFromPicture( _
        ByRef picThis As StdPicture _
    )
    Dim lHDC As Long
    Dim lhDCDesktop As Long
    Dim lhBmpOld As Long
    Dim tBMP As BITMAP

    GetObjectAPI picThis.handle, Len(tBMP), tBMP
    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhDCDesktop = GetDC(GetDesktopWindow())
        If (lhDCDesktop <> 0) Then
            lHDC = CreateCompatibleDC(lhDCDesktop)
            DeleteDC lhDCDesktop
            If (lHDC <> 0) Then
                lhBmpOld = SelectObject(lHDC, picThis.handle)
                LoadPictureBlt lHDC
                SelectObject lHDC, lhBmpOld
                DeleteObject lHDC
            End If
        End If
    End If
End Function

Public Function Create( _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
    
    ClearUp
    m_hDC = CreateCompatibleDC(0)
    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If
End Function

Public Function CreateDIB( _
        ByVal lHDC As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long, _
        ByRef hDib As Long _
    ) As Boolean

    '//Create 24 Bit DIB

    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = ImageBufferSize * .biHeight
    End With

    hDib = CreateDIBSection( _
            lHDC, _
            m_tBI, _
            DIB_RGB_COLORS, _
            m_lPtr, _
            0, 0)

    CreateDIB = (hDib <> 0)

    If CreateDIB = True Then
        Call CreateSafeArray
    Else
        m_xBound = -1
        m_yBound = -1
    End If
End Function

Public Property Get ImageBufferSize() As Long
    ' Scans must align on dword boundaries: '//For 24 bit
    ImageBufferSize = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC
End Property

Public Property Get Width() As Long
    Width = m_tBI.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = m_tBI.bmiHeader.biHeight
End Property
Public Property Get XBound() As Long
    XBound = m_xBound
End Property
Public Property Get YBound() As Long
    YBound = m_yBound
End Property

Public Sub LoadPictureBlt( _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Sub PaintPicture( _
        ByVal lHDC As Long, _
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
    BitBlt lHDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop
End Sub

Public Property Get hdc() As Long
    hdc = m_hDC
End Property

Public Property Get hDib() As Long
    hDib = m_hDIb
End Property

Public Property Get DIBSectionBitsPtr() As Long
    DIBSectionBitsPtr = m_lPtr
End Property

Public Sub ClearUp()
    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteObject m_hDC
    End If
    m_hDC = 0: m_hDIb = 0: m_hBmpOld = 0: m_lPtr = 0

    CopyMemory ByVal VarPtrArray(bDib), 0&, 4    '//clear DIB array pointer
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub
Public Sub Fade(ByVal lAmount As Long)
    Dim x As Long, y As Long
    Dim xMax As Long, yMax As Long
    Dim lB As Long, lG As Long, lR As Long

    yMax = Height - 1
    xMax = Width - 1

    For x = 0 To (xMax * 3) Step 3
        For y = 0 To yMax
            lB = lAmount * bDib(x, y) \ 255
            lG = lAmount * bDib(x + 1, y) \ 255
            lR = lAmount * bDib(x + 2, y) \ 255
            bDib(x, y) = lB
            bDib(x + 1, y) = lG
            bDib(x + 2, y) = lR

        Next y
    Next x
End Sub

'//0 based pixel array (0,0) is first element
Public Function SetBitValue(xPos As Long, yPos As Long, value As Byte) As Boolean
    If (xPos <= m_xBound And yPos <= m_yBound) And (yPos >= 0 And yPos >= 0) Then
        bDib(xPos, yPos) = value
    End If
End Function

'//////////////////////////////////////////////////////////////////
'Now the bDib() array is a two dimensional array with
'the first dimension being the x values and the second
'being the y values. A 24 bit DIB section is arranged
'so the bytes run Blue, Green, Red and remember that since
'the array is padded to a DWORD boundary there may be up to
'three unused bytes at the end of each row. So, for example,
'to set the top left pixel to purple you would write this:
'
'        bDib(0, 0) = 255 ' Blue
'        bDib(1, 0) = 0 ' Green
'        bDib(2, y) = 255 ' Red
'
'first pixel at (PixelCol=1,PixelRow=1)
'
'Description : This function returns an array of R,G,B value at
'              specified pixel (e.g PixelCol=1 and PixelRow=1 will
'              return R,G,B value of top-left pixel. 1 Pixel requires 3 bytes
'              to store RGB value
'Input  : PixelCol and PixelRow => X starts at 1, Y starts at 1
'         R,G,B are return values (R=red,G=green,B=blue)
'Return : Red,Green and Blue values
'/////////////////////////////////////////////////////////////////
Public Function GetPixelRGB(PixelRow As Long, PixelCol As Long, ByRef R As Byte, ByRef G As Byte, ByRef B As Byte)
    If (PixelCol <= Width And PixelRow <= Height) And (PixelCol > 0 And PixelRow > 0) Then
        R = bDib((PixelCol - 1) * 3 + 2, PixelRow - 1)    '//Extract value for Red color of Pixel
        G = bDib((PixelCol - 1) * 3 + 1, PixelRow - 1)    '//Extract value for Green color of Pixel
        B = bDib((PixelCol - 1) * 3, PixelRow - 1)    '//Extract value for Blue color of Pixel
    End If
End Function
Public Function SetPixelRGB(Optional PixelRow As Long = 1, Optional PixelCol As Long = 1, Optional R, Optional G, Optional B)
    If (PixelCol <= Width And PixelRow <= Height) And (PixelCol > 0 And PixelRow > 0) Then
        If IsMissing(R) = False Then bDib((PixelCol - 1) * 3 + 2, PixelRow - 1) = R    '//Set value for Red color of Pixel
        If IsMissing(G) = False Then bDib((PixelCol - 1) * 3 + 1, PixelRow - 1) = G    '//Set value for Green color of Pixel
        If IsMissing(B) = False Then bDib((PixelCol - 1) * 3, PixelRow - 1) = B    '//Set value for Blue color of Pixel
    End If
End Function

Private Sub CreateSafeArray()
    ' have the local matrix point to bitmap pixels
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = Me.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = Me.ImageBufferSize
        .pvData = DIBSectionBitsPtr
    End With
    CopyMemory ByVal VarPtrArray(bDib), VarPtr(tSA), 4

    m_xBound = UBound(bDib, 1)
    m_yBound = UBound(bDib, 2)
End Sub
