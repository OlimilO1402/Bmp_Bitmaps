Attribute VB_Name = "MLoadPng"
Option Explicit

'This module provides a LoadPictureGDI function, which can
'be used instead of VBA's LoadPicture, to load a wide variety
'of image types from disk - including png.
'
'The png format is used in Office 2007-2010 to provide images that
'include an alpha channel for each pixel's transparency
'
'Author:    Stephen Bullen
'Date:      31 October, 2006
'Email:     stephen@oaltd.co.uk

'Updated :  30 December, 2010
'By :       Rob Bovey
'Reason :   Also working now in the 64 bit version of Office 2010

Private Const ERROR_SUCCESS As Long = 0

'Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1         As Long
    Data2         As Integer
    Data3         As Integer
    Data4(0 To 7) As Byte
End Type

'Declare a UDT to store the bitmap information
Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

'Declare a UDT to store the GDI+ Startup information
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'Windows API calls into the GDI+ library
#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As LongPtr, bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#Else
    Private Declare Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
    Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal FileName As LongPtr, bitmap As LongPtr) As Long
    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
    Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As Long
    Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr)
    Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#End If

' Procedure:    LoadPictureGDI
' Purpose:      Loads an image using GDI+
' Returns:      The image as an IPicture Object
Public Function LoadPictureGDIp(ByVal sFilename As String) As StdPicture
    
    'Initialize GDI+
    Dim gsi As GdiplusStartupInput: gsi.GdiplusVersion = 1
    Dim hGDIP As LongPtr
    If GdiplusStartup(hGDIP, gsi) <> ERROR_SUCCESS Then
        MsgBox "Could not start GDI+"
        Exit Function
    End If
    
    'Load the image
    Dim hImage As LongPtr
    If GdipCreateBitmapFromFile(StrPtr(sFilename), hImage) <> ERROR_SUCCESS Then
        MsgBox "Could not create bitmap from file: " & vbCrLf & sFilename
        Exit Function
    End If
    
    'Create a bitmap handle from the GDI image
    Dim hBitmap   As LongPtr
    'lResult = GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0)
    If GdipCreateHBITMAPFromBitmap(hImage, hBitmap, 0) <> ERROR_SUCCESS Then
        MsgBox "Could not create handle from bitmap: " & hBitmap
        Exit Function
    End If
    
    'Create the IPicture object from the bitmap handle
    Set LoadPictureGDIp = CreateStdPicture(hBitmap)
    
    'Tidy up
    GdipDisposeImage hImage
    
    'Shutdown GDI+
    GdiplusShutdown hGDIP
    
End Function

' Procedure:    CreateIPicture
' Purpose:      Converts a image handle into an IPicture object.
' Returns:      The IPicture object
Private Function CreateStdPicture(ByVal hPic As LongPtr) As StdPicture
    
    ' Create the Interface GUID (for the IPicture interface)
    Dim IID_IPicture As GUID
    With IID_IPicture
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With

    ' Fill uPicInfo with necessary parts.
    'OLE Picture types
    Const PICTYPE_BITMAP As Long = 1
    Dim uPicInfo As PICTDESC
    With uPicInfo
        .Size = Len(uPicInfo)
        .Type = PICTYPE_BITMAP
        .hPic = hPic
        .hPal = 0
    End With

    ' Create the Picture object.
    Dim lResult As Long, StdPic As StdPicture
    lResult = OleCreatePictureIndirect(uPicInfo, IID_IPicture, True, StdPic)

    ' Return the new Picture object.
    Set CreateStdPicture = StdPic

End Function
