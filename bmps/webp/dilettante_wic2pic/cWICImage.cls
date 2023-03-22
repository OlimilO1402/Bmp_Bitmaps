VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWICImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const AC_SRC_OVER    As Byte = &H0
Private Const AC_SRC_ALPHA   As Byte = &H1
Private Const BI_RGB         As Long = 0
Private Const DIB_RGB_COLORS As Long = 0
Private Const GENERIC_READ   As Long = &H80000000
Private Const S_OK           As Long = 0

Private Type ARGB
    Blue  As Byte
    Green As Byte
    Red   As Byte
    Alpha As Byte
End Type

Private Type BLENDFUNCTION
    BlendOp             As Byte
    BlendFlags          As Byte
    SourceConstantAlpha As Byte
    AlphaFormat         As Byte
End Type

Private Type BLENDFUNCTION_Long
    Value As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As ARGB
End Type

Private Type Guid
    Data1    As Long
    Data2    As Integer
    Data3    As Integer
    Data4(7) As Byte
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal uiUsage As Long, ByRef ppvBits As Any, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GdiAlphaBlend Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal xoriginDest As Long, ByVal yoriginDest As Long, ByVal wDest As Long, ByVal hDest As Long, ByVal hdcSrc As Long, ByVal xoriginSrc As Long, ByVal yoriginSrc As Long, ByVal wSrc As Long, ByVal hSrc As Long, ByVal ftn As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hgdiobj As Long) As Long

Private Declare Function IWICBitmapDecoder_GetFrameCount_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByRef pCount As Long) As Long
Private Declare Function IWICBitmapDecoder_GetFrame_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByVal Index As Long, ByRef ppIBitmapFrame As stdole.IUnknown) As Long
Private Declare Function IWICBitmapScaler_Initialize_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByVal pISource As stdole.IUnknown, ByVal uiWidth As Long, ByVal uiHeight As Long, ByVal Mode As Long) As Long
Private Declare Function IWICBitmapSource_CopyPixels_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByRef prc As Any, ByVal cbStride As Long, ByVal cbBufferSize As Long, ByRef pbBuffer As Any) As Long
Private Declare Function IWICBitmapSource_GetSize_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByRef puiWidth As Long, ByRef puiHeight As Long) As Long
Private Declare Function IWICFormatConverter_Initialize_Proxy Lib "windowscodecs.dll" (ByVal This_Ptr As stdole.IUnknown, ByVal pISource As stdole.IUnknown, ByRef dstFormat As Guid, ByVal Dither As Long, ByVal pIPalette As stdole.IUnknown, ByVal alphaThresholdPercent As Double, ByVal paletteTranslate As Long) As Long
Private Declare Function IWICImagingFactory_CreateBitmapScaler_Proxy Lib "windowscodecs.dll" (ByVal pFactory As stdole.IUnknown, ByRef ppIBitmapScaler As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateDecoderFromFilename_Proxy Lib "windowscodecs.dll" (ByVal pFactory As stdole.IUnknown, ByVal wzFilename As Long, ByRef pguidVendor As Guid, ByVal dwDesiredAccess As Long, ByVal metadataOptions As Long, ByRef ppIDecoder As stdole.IUnknown) As Long
Private Declare Function IWICImagingFactory_CreateFormatConverter_Proxy Lib "windowscodecs.dll" (ByVal pFactory As stdole.IUnknown, ByRef ppIFormatConverter As stdole.IUnknown) As Long
Private Declare Function WICCreateImagingFactory_Proxy Lib "windowscodecs.dll" (ByVal SDKVersion As Long, ByRef ppIImagingFactory As stdole.IUnknown) As Long

Private m_Loaded    As Boolean
Private m_Width     As Long
Private m_Height    As Long
Private m_Frame     As Long
Private m_FrameCnt  As Long
Private m_File      As String
Private m_ImgFact   As stdole.IUnknown 'WICImagingFactory
Private m_Converter As stdole.IUnknown 'IWICFormatConverter

Public Property Get ImageWidth() As Long
    ImageWidth = m_Width
End Property

Public Property Get ImageHeight() As Long
    ImageHeight = m_Height
End Property

Public Property Get FrameCount() As Long
    FrameCount = m_FrameCnt
End Property

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_Loaded
End Property

Public Sub Render(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long)
    Const WICBitmapInterpolationMode = 4&
    Dim hDCScr As Long, hDIBBitmap As Long, nImage As Long, nStride As Long, pvImageBits As Long, RV As Long, tBMI As BITMAPINFO
    Dim hDCMem As Long, hbmOrig As Long, BF As BLENDFUNCTION, BFL As BLENDFUNCTION_Long, oBmpScaler As stdole.IUnknown 'IWICBitmapScaler

    If Not m_Loaded Then Exit Sub
    If IWICImagingFactory_CreateBitmapScaler_Proxy(m_ImgFact, oBmpScaler) <> S_OK Then Exit Sub
    If IWICBitmapScaler_Initialize_Proxy(oBmpScaler, m_Converter, cX, cY, WICBitmapInterpolationMode) <> S_OK Then Exit Sub

    hDCScr = GetDC(0&)

    If hDCScr Then
        hDCMem = CreateCompatibleDC(hDCScr)

        If hDCMem Then
            With tBMI.bmiHeader
                .biSize = LenB(tBMI.bmiHeader)
                .biWidth = cX
                .biHeight = -cY
                .biPlanes = 1
                .biBitCount = 32
                .biCompression = BI_RGB
            End With

            hDIBBitmap = CreateDIBSection(hDCScr, tBMI, DIB_RGB_COLORS, pvImageBits, 0&, 0&)

            If hDIBBitmap Then
                hbmOrig = SelectObject(hDCMem, hDIBBitmap)

                If hbmOrig Then
                    nStride = 4& * ((cX * 32& + 31&) \ 32&)
                    nImage = nStride * cY

                    If IWICBitmapSource_CopyPixels_Proxy(oBmpScaler, ByVal 0&, nStride, nImage, ByVal pvImageBits) = S_OK Then
                        BF.BlendOp = AC_SRC_OVER
                        BF.BlendFlags = 0
                        BF.SourceConstantAlpha = 255
                        BF.AlphaFormat = AC_SRC_ALPHA
                        LSet BFL = BF

                        RV = GdiAlphaBlend(hDC, X, Y, cX, cY, hDCMem, 0&, 0&, cX, cY, BFL.Value)
                    End If

                    RV = SelectObject(hDCMem, hbmOrig)
                End If

                RV = DeleteObject(hDIBBitmap)
            End If

            RV = DeleteDC(hDCMem)
        End If

        RV = ReleaseDC(0&, hDCScr)
    End If
End Sub

Public Function OpenFile(ByRef sFile As String, Optional ByVal nFrame As Long) As Boolean
    Const WICDecodeMetadataCacheOnDemand = 0&, WICBitmapDitherTypeNone = 0&, WICBitmapPaletteTypeCustom = 0&
    Const WINCODEC_SDK_VERSION1 = &H236&, WINCODEC_SDK_VERSION2 = &H237&
    Dim oBmpDecoder As stdole.IUnknown 'IWICBitmapDecoder
    Dim oFrame      As stdole.IUnknown 'IWICBitmapFrameDecode

    m_Loaded = False
    
    If m_ImgFact Is Nothing Then
        If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION2, m_ImgFact) <> S_OK Then
            If WICCreateImagingFactory_Proxy(WINCODEC_SDK_VERSION1, m_ImgFact) <> S_OK Then Exit Function
        End If
    End If

    m_File = sFile
    m_Frame = nFrame
    
    If IWICImagingFactory_CreateDecoderFromFilename_Proxy(m_ImgFact, StrPtr(sFile), UUID_NULL, GENERIC_READ, WICDecodeMetadataCacheOnDemand, oBmpDecoder) <> S_OK Then Exit Function
    If IWICBitmapDecoder_GetFrameCount_Proxy(oBmpDecoder, m_FrameCnt) <> S_OK Then Exit Function
    If IWICBitmapDecoder_GetFrame_Proxy(oBmpDecoder, m_Frame, oFrame) <> S_OK Then Exit Function
    If IWICBitmapSource_GetSize_Proxy(oFrame, m_Width, m_Height) <> S_OK Then Exit Function
    If IWICImagingFactory_CreateFormatConverter_Proxy(m_ImgFact, m_Converter) <> S_OK Then Exit Function
    If IWICFormatConverter_Initialize_Proxy(m_Converter, oFrame, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, Nothing, 0#, WICBitmapPaletteTypeCustom) <> S_OK Then Exit Function

    m_Loaded = True
    OpenFile = True
End Function

Private Function UUID_NULL() As Guid: End Function

Private Function GUID_WICPixelFormat32bppPBGRA() As Guid
    Static iid As Guid
    If iid.Data1 = 0& Then DEFINE_UUID iid, &H6FDDC324, &H4E03, &H4BFE, &HB1, &H85, &H3D, &H77, &H76, &H8D, &HC9, &H10
    GUID_WICPixelFormat32bppPBGRA = iid
End Function

Private Sub DEFINE_UUID(ByRef U As Guid, ByVal D1 As Long, ByVal D2 As Integer, ByVal D3 As Integer, ByVal D4_0 As Byte, ByVal D4_1 As Byte, ByVal D4_2 As Byte, _
    ByVal D4_3 As Byte, ByVal D4_4 As Byte, ByVal D4_5 As Byte, ByVal D4_6 As Byte, ByVal D4_7 As Byte)
    U.Data1 = D1: U.Data2 = D2: U.Data3 = D3
    U.Data4(0) = D4_0: U.Data4(1) = D4_1: U.Data4(2) = D4_2: U.Data4(3) = D4_3: U.Data4(4) = D4_4: U.Data4(5) = D4_5: U.Data4(6) = D4_6: U.Data4(7) = D4_7
End Sub
