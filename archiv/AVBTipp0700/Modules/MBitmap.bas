Attribute VB_Name = "MBitmap"
Option Explicit
' Dieser Source stammt von http://www.activevb.de
' und kann frei verwendet werden. Für eventuelle Schäden wird nicht gehaftet.
' ----==== Const ====----
Private Const BI_RGB  As Long = 0&
Private Const BI_RLE8 As Long = 1&
Private Const BI_RLE4 As Long = 2&

Private Const DIB_RGB_COLORS As Long = 0&
Private Const DIB_PAL_COLORS As Long = 1&

Private Const IPictureCLSID As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
    
Private Const S_OK As Long = &H0

' ----==== Enum ====----
Public Enum BPP
    PixelFormat1bppIndexed = 0      ' 1
    PixelFormat4bppIndexed = 1      ' 4
    PixelFormat4bppIndexed_RLE = 2  ' 5
    PixelFormat8bppIndexed = 3      ' 8
    PixelFormat8bppIndexed_RLE = 4  ' 9
    PixelFormat16bppRGB = 5         '16
    PixelFormat24bppRGB = 6         '24
    PixelFormat32bppRGB = 7         '32
End Enum

' ----==== Type ====----
Private Type BITMAP
    bmType       As Long
    bmWidth      As Long
    bmHeight     As Long
    bmWidthBytes As Long
    bmPlanes     As Integer
    bmBitsPixel  As Integer
    bmBits       As Long
End Type

Private Type BITMAPFILEHEADER
    bfType      As Integer
    bfSize      As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits   As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long  ' the BI_-constants
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType        As Long
    hgdiObj        As Long
    hPalOrXYExt    As Long
End Type

' ----==== GDI32 API Deklarationen ====----

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-createdibsection
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, ByRef pBitmapInfo As BITMAPINFO256, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

'GetDIBits
'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getdibits
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long
'SetDIBits
'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-setdibits
Private Declare Function SetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO256, ByVal wUsage As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/wingdi/nf-wingdi-getobjecta
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long


' ----==== OLE32 API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef pclsid As Guid) As Long

' ----==== OLEAUT32 API Declarations ====----
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (ByRef lpPictDesc As PICTDESC, ByRef riid As Guid, ByVal fOwn As Boolean, ByRef lplpvObj As Object) As Long

' ----==== USER32 API Deklarationen ====----
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
                         
Public Function BPP_ToStr(ByVal Value As BPP) As String
    Dim s As String
    Select Case Value
    Case PixelFormat1bppIndexed:      s = "1 bpp Indexed"      ' 0
    Case PixelFormat4bppIndexed:      s = "4 bpp Indexed"      ' 1
    Case PixelFormat4bppIndexed_RLE:  s = "4 bpp Indexed RLE"  ' 2
    Case PixelFormat8bppIndexed:      s = "8 bpp Indexed"      ' 3
    Case PixelFormat8bppIndexed_RLE:  s = "8 bpp Indexed RLE"  ' 4
    Case PixelFormat16bppRGB:         s = "16 bpp RGB"         ' 5
    Case PixelFormat24bppRGB:         s = "24 bpp RGB"         ' 6
    Case PixelFormat32bppRGB:         s = "32 bpp RGB"         ' 7
    End Select
    BPP_ToStr = s
End Function

Public Sub BPP_ToCBLB(aCBLB)
    With aCBLB
        .Clear
        Dim i As Long, s As String
        For i = 0 To 7
            s = BPP_ToStr(i)
            If Len(s) Then .AddItem s
        Next
    End With
End Sub

Public Function BPP_Parse(s As String) As BPP
    Dim b As Byte: b = CByte(Trim(Left(s, 2)))
    Dim r As Boolean: r = Right(s, 3) <> "RLE"
    Select Case b
    Case 1:  b = 0
    Case 4:  b = IIf(r, 1, 2)
    Case 8:  b = IIf(r, 3, 4)
    Case 16: b = 5
    Case 24: b = 6
    Case 32: b = 7
    End Select
    BPP_Parse = b
End Function

' ------------------------------------------------------
' Funktion     : ConvertBitmapAllRes
' Beschreibung : Konvertiert ein StdPicture mit eingestellter
'                Farbtiefe in ein StdPicture
' Übergabewert : Pic = StdPicture
'                BitsPerPixel = Enum BPP
' Rückgabewert : StdPicture
' ------------------------------------------------------
Public Function ConvertBitmapAllRes(ByVal Pic As StdPicture, Optional ByVal BitsPerPixel As BPP = PixelFormat24bppRGB) As StdPicture
    
    ' Fehlerbehandlung
Try: On Error GoTo Catch
    
    Dim hConvBmp As Long
    Dim tBITMAPINFO As BITMAPINFO256
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As Guid
    Dim oPicture As IPicture
    
    Dim lngBitsPerPixel As Long
    Dim lngPalCount     As Long
    Dim bPalette        As Boolean
    Dim biCompr         As Long: biCompr = BI_RGB
    
    ' div. Standardeinstellungen für
    ' die entsprechenden Pixelformate
    Select Case BitsPerPixel
    Case BPP.PixelFormat1bppIndexed ' 2 Farben
        lngBitsPerPixel = 1
        lngPalCount = 2
        bPalette = True
    Case BPP.PixelFormat4bppIndexed, BPP.PixelFormat4bppIndexed_RLE ' 16 Farben
        lngBitsPerPixel = 4
        lngPalCount = 16
        bPalette = True
        If BitsPerPixel = BPP.PixelFormat4bppIndexed_RLE Then
            biCompr = BI_RLE4
        End If
    Case BPP.PixelFormat8bppIndexed, PixelFormat8bppIndexed_RLE ' 256 Farben
        lngBitsPerPixel = 8
        lngPalCount = 256
        bPalette = True
        'If BitsPerPixel = BPP.PixelFormat8bppIndexed_RLE Then
        '    biCompr = BI_RLE8
        'nd If
    Case BPP.PixelFormat16bppRGB ' 16Bit
        lngBitsPerPixel = 16
    Case BPP.PixelFormat24bppRGB ' 24Bit
        lngBitsPerPixel = 24
    Case BPP.PixelFormat32bppRGB ' 32Bit
        lngBitsPerPixel = 32
    End Select
    
    ' Bitmapinfos vom StdPicture auslesen -> tBITMAP
    Dim tBITMAP As BITMAP
    If GetObjectA(Pic.handle, Len(tBITMAP), tBITMAP) = 0 Then
        MsgBox "Could not get object for handle: " & Pic.handle
        Exit Function
    End If
    
    ' ausgelesene Bitmapinfos + Standardeinstellungen für
    ' das entsprechende Pixelformat übertragen
    With tBITMAPINFO.bmiHeader
        .biSize = Len(tBITMAPINFO.bmiHeader)
        .biWidth = tBITMAP.bmWidth
        .biHeight = tBITMAP.bmHeight
        .biPlanes = tBITMAP.bmPlanes
        .biBitCount = lngBitsPerPixel
        .biCompression = biCompr
    End With
    
    ' DC ermitteln
    Dim hDC As Long: hDC = GetDC(0&)
    
    ' ist ein DC vorhanden
    If hDC = 0 Then
        MsgBox "Could not create devicecontext"
        Exit Function
    End If
    
    ' Der 1. Aufruf ohne Übergabe von bytArray, dient dazu die
    ' Größe des benötigten Feldes festzustellen.
    ' Die Palette wird hier auch schon übertragen.
    If GetDIBits(hDC, Pic.handle, 0&, tBITMAP.bmHeight, ByVal 0&, tBITMAPINFO, DIB_RGB_COLORS) = 0 Then
        MsgBox "Could not get DIBits"
        GoTo Finally
    End If
    
    ' Array zur Aufnahme der Bitmapdaten dimensionieren.
    ReDim bytArray(tBITMAPINFO.bmiHeader.biSizeImage - 1) As Byte
    
    ' Jetzt wird gelesen. Die Bitmapdaten befinden sich anschließend im Byte-Array.
    If GetDIBits(hDC, Pic.handle, 0&, tBITMAP.bmHeight, bytArray(0), tBITMAPINFO, DIB_RGB_COLORS) = 0 Then
        MsgBox "Could not get bytearray from DIBits "
        GoTo Finally
    End If
        
    ' ist es eine Palettenbitmap
    ' (1bpp, 4bpp und 8bpp)
    If bPalette Then
        
        ' Anzahl der verwendeten Farben in der Palette
        tBITMAPINFO.bmiHeader.biClrUsed = lngPalCount
        
        ' Anzahl der verwendeten Farben in der Palette
        tBITMAPINFO.bmiHeader.biClrImportant = lngPalCount
        
    End If
    
    ' Neue DIB-Bitmap erstellen
    'hConvBmp = CreateDIBSection(hDC, tBITMAPINFO, DIB_RGB_COLORS, 0&, 0&, 0&)
    hConvBmp = CreateDIBSection(hDC, tBITMAPINFO, DIB_RGB_COLORS, 0&, 0&, 0&)
    'biCompr
    ' ist ein DIB-Bitmap vorhanden
    If hConvBmp = 0 Then
        MsgBox "Could not create DIB-Section"
        GoTo Finally
    End If
    
    ' Bitmapdaten in das DIB-Bitmap schreiben
    If SetDIBits(hDC, hConvBmp, 0&, tBITMAP.bmHeight, bytArray(0), tBITMAPINFO, DIB_RGB_COLORS) = 0 Then
        MsgBox "Could not set DIBits-bytearray"
        GoTo Finally
    End If
        
    ' Initialisiert die PICTDESC Struktur
    With tPictDesc
    
        .cbSizeOfStruct = Len(tPictDesc)
        .picType = vbPicTypeBitmap
        .hgdiObj = hConvBmp
        .hPalOrXYExt = 0&
        
    End With
    
    ' IPictureCLSID -> IID_IPicture
    If CLSIDFromString(StrPtr(IPictureCLSID), IID_IPicture) = S_OK Then
        
        ' Erzeugen des Ipicture-Objektes
        If OleCreatePictureIndirect(tPictDesc, IID_IPicture, True, oPicture) = S_OK Then
            
            ' DIB-Bitmap in ein StdPicture
            ' konvertieren
            Set ConvertBitmapAllRes = oPicture
        End If
    End If
             
    GoTo Finally
Catch:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , "ConvertBitmapAllRes"
    'Resume Finally
Finally:
    Call ReleaseDC(0&, hDC)
End Function

' ------------------------------------------------------
' Funktion     : FileExists
' Beschreibung : Prüft, ob eine Datei schon vorhanden ist
' Übergabewert : FileName = Pfad\Datei.ext
' Rückgabewert : True = Datei ist vorhanden
'                False = Datei ist nicht vorhanden
' ------------------------------------------------------
Private Function FileExists(ByVal FileName As String) As Boolean

    ' Fehlerbehandlung
    On Error Resume Next
    
    Dim ret As Long
    
    ret = Len(Dir$(FileName))
    
    If Err Or ret = 0 Then FileExists = False Else FileExists = True
    
End Function

' ------------------------------------------------------
' Funktion     : SaveBitmapAllRes
' Beschreibung : Speichert ein StdPicture mit eingestellter
'                Farbtiefe als Bitmap
' Übergabewert : Pic = StdPicture
'                FileName = Pfad\Datei.bmp
'                BitsPerPixel = Enum BPP
' Rückgabewert : True = Speichern war erfolgreich
'                False = Speichern war nicht erfolgreich
' ------------------------------------------------------
Public Function SaveBitmapAllRes(ByVal Pic As StdPicture, ByVal FileName As String, Optional ByVal BitsPerPixel As BPP = PixelFormat24bppRGB) As Boolean
    
    ' Fehlerbehandlung
    On Error GoTo PROC_ERR
    
    Dim lngDC As Long
    Dim lngFNr As Long
    Dim lngPalCount As Long
    Dim lngPalItem As Long
    Dim lngPalette() As Long
    Dim lngBitsPerPixel As Long
    Dim lngCompression As Long
    Dim bytArray() As Byte
    Dim bolPalette As Boolean
    Dim tBITMAP As BITMAP
    Dim tBITMAPINFO As BITMAPINFO256
    Dim tBITMAPFILEHEADER As BITMAPFILEHEADER
    
    ' div. Standardeinstellungen
    lngPalCount = 0
    lngCompression = BI_RGB
    bolPalette = False
    
    ' div. Standardeinstellungen für
    ' die entsprechenden Pixelformate
    Select Case BitsPerPixel
    
        ' 2 Farben unkomprimiert
    Case BPP.PixelFormat1bppIndexed
        lngBitsPerPixel = 1
        lngPalCount = 2
        bolPalette = True
        
        ' 16 Farben unkomprimiert
    Case BPP.PixelFormat4bppIndexed
        lngBitsPerPixel = 4
        lngPalCount = 16
        bolPalette = True
        
        ' 16 Farben (Run Length Encoded)
    Case BPP.PixelFormat4bppIndexed_RLE
        lngBitsPerPixel = 4
        lngCompression = BI_RLE4
        lngPalCount = 16
        bolPalette = True
        
        ' 256 Farben unkomprimiert
    Case BPP.PixelFormat8bppIndexed
        lngBitsPerPixel = 8
        lngPalCount = 256
        bolPalette = True
        
        ' 256 Farben (Run Length Encoded)
    Case BPP.PixelFormat8bppIndexed_RLE
        lngBitsPerPixel = 8
        lngCompression = BI_RLE8
        lngPalCount = 256
        bolPalette = True
        
        ' 16Bit unkomprimiert
    Case BPP.PixelFormat16bppRGB
        lngBitsPerPixel = 16
        
        ' 24Bit unkomprimiert
    Case BPP.PixelFormat24bppRGB
        lngBitsPerPixel = 24
        
        ' 32Bit unkomprimiert
    Case BPP.PixelFormat32bppRGB
        lngBitsPerPixel = 32
        
    End Select
    
    ' Bitmapinfos vom StdPicture auslesen -> tBITMAP
    If GetObjectA(Pic.handle, Len(tBITMAP), tBITMAP) <> 0 Then
    
        ' ausgelesene Bitmapinfos + Standardeinstellungen für
        ' das entsprechende Pixelformat übertragen
        tBITMAPINFO.bmiHeader.biSize = Len(tBITMAPINFO.bmiHeader)
        tBITMAPINFO.bmiHeader.biWidth = tBITMAP.bmWidth
        tBITMAPINFO.bmiHeader.biHeight = tBITMAP.bmHeight
        tBITMAPINFO.bmiHeader.biPlanes = tBITMAP.bmPlanes
        tBITMAPINFO.bmiHeader.biBitCount = lngBitsPerPixel
        tBITMAPINFO.bmiHeader.biCompression = lngCompression
        
        ' DC ermitteln
        lngDC = GetDC(0&)
        
        ' ist ein DC vorhanden
        If lngDC <> 0 Then
        
            ' Der 1. Aufruf ohne Übergabe von bytArray, dient dazu die
            ' Größe des benötigten Feldes festzustellen. Die Palette
            ' wird hier auch schon übertragen.
            If GetDIBits(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, ByVal 0&, tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                
                ' Array zur Aufnahme der Bitmapdaten dimensionieren.
                ReDim bytArray(tBITMAPINFO.bmiHeader.biSizeImage - 1)
                
                ' Jetzt wird tatsächlich gelesen. Die Bitmapdaten
                ' befinden sich anschließend in bytArray.
                If GetDIBits(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, bytArray(0), tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                    
                    ' ist es eine Palettenbitmap
                    ' (1bpp, 4bpp und 8bpp)
                    If bolPalette Then
                    
                        ' Anzahl der verwendeten Farben in der Palette
                        tBITMAPINFO.bmiHeader.biClrUsed = lngPalCount
                        
                        ' Anzahl der verwendeten Farben in der Palette
                        tBITMAPINFO.bmiHeader.biClrImportant = lngPalCount
                        
                        ' Array zur Aufnahme der Palettendaten
                        ' dimensionieren.
                        ReDim lngPalette(lngPalCount - 1)
                        
                        ' Palettendaten umkopieren, damit wir die
                        ' Palette einfach mit Put ausgeben können.
                        For lngPalItem = 0 To lngPalCount - 1
                        
                            lngPalette(lngPalItem) = tBITMAPINFO.bmiColors(lngPalItem)
                                
                        Next lngPalItem
                        
                    End If
                    
                    ' entspricht "BM"
                    tBITMAPFILEHEADER.bfType = 19778
                    
                    ' gesamte Größe der Bitmap
                    tBITMAPFILEHEADER.bfSize = Len(tBITMAPFILEHEADER) + Len(tBITMAPINFO.bmiHeader) + tBITMAPINFO.bmiHeader.biSizeImage
                        
                    ' Offset bis zu den Bitmapdaten
                    tBITMAPFILEHEADER.bfOffBits = Len(tBITMAPFILEHEADER) + Len(tBITMAPINFO.bmiHeader)
                    
                    ' ist es eine Palettenbitmap
                    ' (1bpp, 4bpp und 8bpp)
                    If bolPalette Then
                    
                        ' dann muss die größe der Palettendaten
                        ' hinzugerechnet werden
                        
                        ' größe der Palettendaten zur gesamten
                        ' größe der Bitmap hinzurechnen
                        tBITMAPFILEHEADER.bfSize = tBITMAPFILEHEADER.bfSize + (lngPalCount * 4)
                        
                        ' größe der Palettendaten zum Offset hinzurechnen
                        tBITMAPFILEHEADER.bfOffBits = tBITMAPFILEHEADER.bfOffBits + (lngPalCount * 4)
                            
                    End If
                    
                    ' eventuell vorhandene Datei löschen
                    If FileExists(FileName) Then Kill FileName
                    
                    ' freie Dateinummer ermitteln
                    lngFNr = FreeFile
                    
                    ' Datei öffnen
                    Open FileName For Binary As #lngFNr
                    
                    ' BITMAPFILEHEADER in die Datei schreiben
                    Put #lngFNr, , tBITMAPFILEHEADER
                    
                    ' BITMAPINFOHEADER in die Datei schreiben
                    Put #lngFNr, , tBITMAPINFO.bmiHeader
                    
                    ' ist es eine Palettenbitmap, dann müssen auch die
                    ' Palettendaten in die Datei geschrieben werden
                    ' (1bpp, 4bpp und 8bpp)
                    If bolPalette Then Put #lngFNr, , lngPalette()
                    
                    ' Bitmapdaten in die Datei schreiben
                    Put #lngFNr, , bytArray()
                    
                    ' Datei schließen
                    Close #lngFNr
                    
                    ' Speichern war erfolgreich
                    SaveBitmapAllRes = True
                    
                End If
                
            End If
            
            ' DC freigeben
            Call ReleaseDC(0&, lngDC)
            
        End If
        
    End If
    
PROC_EXIT:

    ' Funktion verlassen
    Exit Function
    
    ' bei Fehler
PROC_ERR:

    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "SaveBitmapAllRes"
        
    Resume PROC_EXIT
    
End Function
