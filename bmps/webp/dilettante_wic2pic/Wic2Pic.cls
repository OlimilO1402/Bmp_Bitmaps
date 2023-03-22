VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Wic2Pic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'NOTE:
'
'   Most Enums here are stripped to the minimum.
'

Private Const WIN32_NULL As Long = 0
Private Const WIN32_TRUE As Long = 1

Private Enum HRESULT
    S_OK = 0
End Enum

Private Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal Width As Long, _
    ByVal Height As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function IIDFromString Lib "ole32" ( _
    ByVal lpsz As Long, _
    ByRef IID As IID) As Long

Private Declare Function IWICBitmapDecoder_GetFrame_Proxy Lib "windowscodecs" ( _
    ByVal THIS_PTR As IUnknown, _
    ByVal Index As Long, _
    ByRef pIBitmapFrame As IUnknown) As HRESULT

Private Enum WICBitmapInterpolationMode
    WICBitmapInterpolationModeNearestNeighbor
    WICBitmapInterpolationModeLinear
    WICBitmapInterpolationModeCubic
    WICBitmapInterpolationModeFant
    WICBitmapInterpolationModeHighQualityCubic
End Enum

Private Declare Function IWICBitmapScaler_Initialize_Proxy Lib "windowscodecs" ( _
    ByVal THIS_PTR As IUnknown, _
    ByVal pISource As IUnknown, _
    ByVal uiWidth As Long, _
    ByVal uiHeight As Long, _
    ByVal Mode As WICBitmapInterpolationMode) As HRESULT

Private Declare Function IWICBitmapSource_CopyPixels_Proxy Lib "windowscodecs" ( _
    ByVal THIS_PTR As IUnknown, _
    ByRef rc As Any, _
    ByVal cbStride As Long, _
    ByVal cbBufferSize As Long, _
    ByRef bBuffer As Byte) As HRESULT

Private Declare Function IWICBitmapSource_GetSize_Proxy Lib "windowscodecs" ( _
    ByVal THIS_PTR As IUnknown, _
    ByRef uiWidth As Long, _
    ByRef uiHeight As Long) As HRESULT

Private Enum WICBitmapDitherType
    WICBitmapDitherTypeNone
End Enum

Private Enum WICBitmapPaletteType
    WICBitmapPaletteTypeCustom
End Enum

Private Declare Function IWICFormatConverter_Initialize_Proxy Lib "windowscodecs" ( _
    ByVal THIS_PTR As IUnknown, _
    ByVal pISource As IUnknown, _
    ByRef dstFormat As IID, _
    Optional ByVal Dither As WICBitmapDitherType = WICBitmapDitherTypeNone, _
    Optional ByVal pIPalette As IUnknown = Nothing, _
    Optional ByVal alphaThresholdPercent As Double = 0#, _
    Optional ByVal paletteTranslate As WICBitmapPaletteType = WICBitmapPaletteTypeCustom) _
    As HRESULT

Private Declare Function IWICImagingFactory_CreateBitmapScaler_Proxy Lib "windowscodecs" ( _
    ByVal pFactory As IUnknown, _
    ByRef pIBitmapScaler As IUnknown) As HRESULT

Private Enum ACCESS_TYPES
    GENERIC_READ = &H80000000
End Enum

Private Enum WICDecodeOptions
    WICDecodeMetadataCacheOnDemand
End Enum

Private Declare Function IWICImagingFactory_CreateDecoderFromFilename_Proxy _
    Lib "windowscodecs" ( _
    ByVal pFactory As IUnknown, _
    ByVal wzFilename As Long, _
    ByRef guidVendor As IID, _
    ByVal dwDesiredAccess As ACCESS_TYPES, _
    ByVal metadataOptions As WICDecodeOptions, _
    ByRef pIDecoder As IUnknown) As HRESULT

Private Declare Function IWICImagingFactory_CreateFormatConverter_Proxy _
    Lib "windowscodecs" ( _
    ByVal pFactory As IUnknown, _
    ByRef pIFormatConverter As IUnknown) As HRESULT

Private Enum WINCODEC_SDK_VERSIONS
    WINCODEC_SDK_VERSION1 = &H236&
    WINCODEC_SDK_VERSION2 = &H237&
End Enum

Private Declare Function WICCreateImagingFactory_Proxy Lib "windowscodecs" ( _
    ByVal SDKVersion As WINCODEC_SDK_VERSIONS, _
    ByRef pIImagingFactory As IUnknown) As HRESULT

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As PictureTypeConstants
    hBitmap As Long
    hpal As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef PICTDESC As PICTDESC, _
    ByRef IID As IID, _
    ByVal fOwn As Long, _
    ByRef IPicture As IPicture) As HRESULT

Private Declare Function ReleaseDC Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hDC As Long) As Long

Private Const BI_RGB As Long = 0&

Private Type BITMAPINFOHEADER
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
End Type

Private Const DIB_RGB_COLORS As Long = 0&

Private Declare Function SetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef Bits As Byte, _
    ByRef bmi As BITMAPINFO, _
    Optional ByVal wUsage As Long = DIB_RGB_COLORS) As Long

Private FormatConverter As IUnknown
Private GUID_NULL As IID
Private GUID_WICPixelFormat24bppBGR As IID
Private IID_IPicture As IID
Private ImagingFactory As IUnknown
Private ImagingFactoryVersion As Byte

Public Function LoadFile( _
    ByRef File As String, _
    Optional ByVal hDC As Long = WIN32_NULL) As IPicture
    'Returns Nothing on any detected errors.

    Dim BitmapDecoder As IUnknown
    Dim BitmapFrame As IUnknown
    Dim BitmapScaler As IUnknown
    Dim Width As Long
    Dim Height As Long
    Dim Stride As Long
    Dim Bits() As Byte
    Dim hBitmap As Long
    Dim BITMAPINFO As BITMAPINFO
    Dim PICTDESC As PICTDESC

    If FormatConverter Is Nothing Then Exit Function
    If IWICImagingFactory_CreateDecoderFromFilename_Proxy(ImagingFactory, _
                                                          StrPtr(File), _
                                                          GUID_NULL, _
                                                          GENERIC_READ, _
                                                          WICDecodeMetadataCacheOnDemand, _
                                                          BitmapDecoder) <> S_OK Then
        Exit Function
    End If
    'Frame indices start at 0:
    If IWICBitmapDecoder_GetFrame_Proxy(BitmapDecoder, 0, BitmapFrame) <> S_OK Then
        Exit Function
    End If
    If IWICBitmapSource_GetSize_Proxy(BitmapFrame, Width, Height) <> S_OK Then Exit Function
    If IWICFormatConverter_Initialize_Proxy(FormatConverter, _
                                            BitmapFrame, _
                                            GUID_WICPixelFormat24bppBGR) <> S_OK Then
        Exit Function
    End If
    If IWICImagingFactory_CreateBitmapScaler_Proxy(ImagingFactory, BitmapScaler) <> S_OK Then
        Exit Function
    End If
    If IWICBitmapScaler_Initialize_Proxy(BitmapScaler, _
                                         FormatConverter, _
                                         Width, _
                                         Height, _
                                         WICBitmapInterpolationModeNearestNeighbor) <> S_OK Then
        Exit Function
    End If
    Stride = 4& * ((24 * Width + 31&) \ 32&)
    ReDim Bits(Stride * Height - 1)
    If IWICBitmapSource_CopyPixels_Proxy(BitmapScaler, _
                                         ByVal WIN32_NULL, _
                                         Stride, _
                                         UBound(Bits) + 1, _
                                         Bits(0)) <> S_OK Then Exit Function
    
    Set BitmapScaler = Nothing
    Set BitmapFrame = Nothing
    Set BitmapDecoder = Nothing
    With BITMAPINFO.bmiHeader
        .biSize = LenB(BITMAPINFO.bmiHeader)
        .biWidth = Width
        .biHeight = -Height
        .biPlanes = 1
        .biCompression = BI_RGB
        .biBitCount = 24
    End With
    If hDC = WIN32_NULL Then
        hDC = GetDC(WIN32_NULL)
        hBitmap = CreateCompatibleBitmap(hDC, Width, Height)
        If SetDIBits(hDC, hBitmap, 0, Height, Bits(0), BITMAPINFO) = 0 Then Exit Function
        ReleaseDC WIN32_NULL, hDC
    Else
        hBitmap = CreateCompatibleBitmap(hDC, Width, Height)
        If SetDIBits(hDC, hBitmap, 0, Height, Bits(0), BITMAPINFO) = 0 Then Exit Function
    End If
    Erase Bits
    With PICTDESC
        .cbSizeOfStruct = LenB(PICTDESC)
        .picType = vbPicTypeBitmap
        .hBitmap = hBitmap
    End With
    OleCreatePictureIndirect PICTDESC, IID_IPicture, WIN32_TRUE, LoadFile
End Function

Private Sub Class_Initialize()
    IIDFromString StrPtr("{6fddc324-4e03-4bfe-b185-3d77768dc90c}"), _
                  GUID_WICPixelFormat24bppBGR
    IIDFromString StrPtr("{7bf80980-bf32-101a-8bbb-00aa00300cab}"), IID_IPicture
    If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION2, ImagingFactory) = S_OK Then
        ImagingFactoryVersion = 2
    Else
        If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION1, ImagingFactory) = S_OK Then
            ImagingFactoryVersion = 1
        End If
    End If
    If ImagingFactoryVersion <> 0 Then
        IWICImagingFactory_CreateFormatConverter_Proxy ImagingFactory, FormatConverter
    End If
End Sub
