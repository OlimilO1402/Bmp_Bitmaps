Attribute VB_Name = "modBitmap"
' Dieser Source stammt von http://www.activevb.de
' und kann frei verwendet werden. Für eventuelle Schäden
' wird nicht gehaftet.
'
' Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
' Ansonsten viel Spaß und Erfolg mit diesem Source!

Option Explicit

' ----==== Const ====----
Private Const BI_RGB As Long = 0&
Private Const BI_RLE4 As Long = 2&
Private Const BI_RLE8 As Long = 1&

Private Const DIB_RGB_COLORS As Long = 0&

Private Const IPictureCLSID As String = _
    "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"
    
Private Const S_OK As Long = &H0

' ----==== Enum ====----
Public Enum BPP
    PixelFormat1bppIndexed = 0
    PixelFormat4bppIndexed = 1
    PixelFormat4bppIndexed_RLE = 2
    PixelFormat8bppIndexed = 3
    PixelFormat8bppIndexed_RLE = 4
    PixelFormat16bppRGB = 5
    PixelFormat24bppRGB = 6
    PixelFormat32bppRGB = 7
End Enum

' ----==== Type ====----
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

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

Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

' ----==== GDI32 API Deklarationen ====----
Private Declare Function CreateDIBSection256 Lib "gdi32.dll" _
                         Alias "CreateDIBSection" ( _
                         ByVal hdc As Long, _
                         ByRef pBitmapInfo As BITMAPINFO256, _
                         ByVal un As Long, _
                         ByVal lplpVoid As Long, _
                         ByVal handle As Long, _
                         ByVal dw As Long) As Long
                         
Private Declare Function GetDIBits256 Lib "gdi32.dll" _
                         Alias "GetDIBits" ( _
                         ByVal aHDC As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO256, _
                         ByVal wUsage As Long) As Long
                         
Private Declare Function GetObject Lib "gdi32.dll" _
                         Alias "GetObjectA" ( _
                         ByVal hObject As Long, _
                         ByVal nCount As Long, _
                         ByRef lpObject As Any) As Long
                         
Private Declare Function SetDIBits256 Lib "gdi32.dll" _
                         Alias "SetDIBits" ( _
                         ByVal hdc As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO256, _
                         ByVal wUsage As Long) As Long
                         
' ----==== OLE32 API Declarations ====----
Private Declare Function CLSIDFromString Lib "ole32.dll" ( _
                         ByVal lpsz As Long, _
                         ByRef pclsid As Guid) As Long
                         
' ----==== OLEOUT32 API Declarations ====----
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
                         ByRef lpPictDesc As PICTDESC, _
                         ByRef riid As Guid, _
                         ByVal fOwn As Boolean, _
                         ByRef lplpvObj As Object) As Long
                         
' ----==== USER32 API Deklarationen ====----
Private Declare Function GetDC Lib "user32.dll" ( _
                         ByVal hwnd As Long) As Long
                         
Private Declare Function ReleaseDC Lib "user32.dll" ( _
                         ByVal hwnd As Long, _
                         ByVal hdc As Long) As Long
                         
' ------------------------------------------------------
' Funktion     : ConvertBitmapAllRes
' Beschreibung : Konvertiert ein StdPicture mit eingestellter
'                Farbtiefe in ein StdPicture
' Übergabewert : Pic = StdPicture
'                BitsPerPixel = Enum BPP
' Rückgabewert : StdPicture
' ------------------------------------------------------
Public Function ConvertBitmapAllRes(ByVal Pic As StdPicture, Optional _
    ByVal BitsPerPixel As BPP = PixelFormat24bppRGB) As StdPicture
    
    ' Fehlerbehandlung
    On Error GoTo PROC_ERR
    
    Dim lngDC As Long
    Dim hConvBmp As Long
    Dim lngPalCount As Long
    Dim lngBitsPerPixel As Long
    Dim bytArray() As Byte
    Dim bolPalette As Boolean
    Dim tBITMAP As BITMAP
    Dim tBITMAPINFO As BITMAPINFO256
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As Guid
    Dim oPicture As IPicture
    
    ' div. Standardeinstellungen
    lngPalCount = 0
    bolPalette = False
    
    ' div. Standardeinstellungen für
    ' die entsprechenden Pixelformate
    Select Case BitsPerPixel
    
        ' 2 Farben
    Case BPP.PixelFormat1bppIndexed
        lngBitsPerPixel = 1
        lngPalCount = 2
        bolPalette = True
        
        ' 16 Farben
    Case BPP.PixelFormat4bppIndexed, BPP.PixelFormat4bppIndexed_RLE
        lngBitsPerPixel = 4
        lngPalCount = 16
        bolPalette = True
        
        ' 256 Farben
    Case BPP.PixelFormat8bppIndexed, PixelFormat8bppIndexed_RLE
        lngBitsPerPixel = 8
        lngPalCount = 256
        bolPalette = True
        
        ' 16Bit
    Case BPP.PixelFormat16bppRGB
        lngBitsPerPixel = 16
        
        ' 24Bit
    Case BPP.PixelFormat24bppRGB
        lngBitsPerPixel = 24
        
        ' 32Bit
    Case BPP.PixelFormat32bppRGB
        lngBitsPerPixel = 32
        
    End Select
    
    ' Bitmapinfos vom StdPicture auslesen -> tBITMAP
    If GetObject(Pic.handle, Len(tBITMAP), tBITMAP) <> 0 Then
    
        ' ausgelesene Bitmapinfos + Standardeinstellungen für
        ' das entsprechende Pixelformat übertragen
        tBITMAPINFO.bmiHeader.biSize = Len(tBITMAPINFO.bmiHeader)
        tBITMAPINFO.bmiHeader.biWidth = tBITMAP.bmWidth
        tBITMAPINFO.bmiHeader.biHeight = tBITMAP.bmHeight
        tBITMAPINFO.bmiHeader.biPlanes = tBITMAP.bmPlanes
        tBITMAPINFO.bmiHeader.biBitCount = lngBitsPerPixel
        tBITMAPINFO.bmiHeader.biCompression = BI_RGB
        
        ' DC ermitteln
        lngDC = GetDC(0&)
        
        ' ist ein DC vorhanden
        If lngDC <> 0 Then
        
            ' Der 1. Aufruf ohne Übergabe von bytArray, dient dazu die
            ' Größe des benötigten Feldes festzustellen. Die Palette
            ' wird hier auch schon übertragen.
            If GetDIBits256(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, _
                ByVal 0&, tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                
                ' Array zur Aufnahme der Bitmapdaten dimensionieren.
                ReDim bytArray(tBITMAPINFO.bmiHeader.biSizeImage - 1)
                
                ' Jetzt wird tatsächlich gelesen. Die Bitmapdaten
                ' befinden sich anschließend in bytArray.
                If GetDIBits256(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, _
                    bytArray(0), tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                    
                    ' ist es eine Palettenbitmap
                    ' (1bpp, 4bpp und 8bpp)
                    If bolPalette Then
                    
                        ' Anzahl der verwendeten Farben in der Palette
                        tBITMAPINFO.bmiHeader.biClrUsed = lngPalCount
                        
                        ' Anzahl der verwendeten Farben in der Palette
                        tBITMAPINFO.bmiHeader.biClrImportant = lngPalCount
                        
                    End If
                    
                    ' DIB-Bitmap erstellen
                    hConvBmp = CreateDIBSection256(lngDC, tBITMAPINFO, _
                        DIB_RGB_COLORS, 0&, 0&, 0&)
                        
                    ' ist ein DIB-Bitmap vorhanden
                    If hConvBmp <> 0 Then
                    
                        ' Bitmapdaten in das DIB-Bitmap schreiben
                        If SetDIBits256(lngDC, hConvBmp, 0&, _
                            tBITMAP.bmHeight, bytArray(0), tBITMAPINFO, _
                            DIB_RGB_COLORS) <> 0 Then
                            
                            ' Initialisiert die PICTDESC Struktur
                            With tPictDesc
                            
                                .cbSizeOfStruct = Len(tPictDesc)
                                .picType = vbPicTypeBitmap
                                .hgdiObj = hConvBmp
                                .hPalOrXYExt = 0&
                                
                            End With
                            
                            ' IPictureCLSID -> IID_IPicture
                            If CLSIDFromString(StrPtr(IPictureCLSID), _
                                IID_IPicture) = S_OK Then
                                
                                ' Erzeugen des Ipicture-Objektes
                                If OleCreatePictureIndirect(tPictDesc, _
                                    IID_IPicture, True, oPicture) = S_OK _
                                    Then
                                    
                                    ' DIB-Bitmap in ein StdPicture
                                    ' konvertieren
                                    Set ConvertBitmapAllRes = oPicture
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
                    
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
        "ConvertBitmapAllRes"
        
    Resume PROC_EXIT
    
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
Public Function SaveBitmapAllRes(ByVal Pic As StdPicture, ByVal FileName _
    As String, Optional ByVal BitsPerPixel As BPP = PixelFormat24bppRGB) _
    As Boolean
    
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
    If GetObject(Pic.handle, Len(tBITMAP), tBITMAP) <> 0 Then
    
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
            If GetDIBits256(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, _
                ByVal 0&, tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                
                ' Array zur Aufnahme der Bitmapdaten dimensionieren.
                ReDim bytArray(tBITMAPINFO.bmiHeader.biSizeImage - 1)
                
                ' Jetzt wird tatsächlich gelesen. Die Bitmapdaten
                ' befinden sich anschließend in bytArray.
                If GetDIBits256(lngDC, Pic.handle, 0&, tBITMAP.bmHeight, _
                    bytArray(0), tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                    
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
                        
                            lngPalette(lngPalItem) = _
                                tBITMAPINFO.bmiColors(lngPalItem)
                                
                        Next lngPalItem
                        
                    End If
                    
                    ' entspricht "BM"
                    tBITMAPFILEHEADER.bfType = 19778
                    
                    ' gesamte Größe der Bitmap
                    tBITMAPFILEHEADER.bfSize = Len(tBITMAPFILEHEADER) + _
                        Len(tBITMAPINFO.bmiHeader) + _
                        tBITMAPINFO.bmiHeader.biSizeImage
                        
                    ' Offset bis zu den Bitmapdaten
                    tBITMAPFILEHEADER.bfOffBits = Len(tBITMAPFILEHEADER) _
                        + Len(tBITMAPINFO.bmiHeader)
                        
                    ' ist es eine Palettenbitmap
                    ' (1bpp, 4bpp und 8bpp)
                    If bolPalette Then
                    
                        ' dann muss die größe der Palettendaten
                        ' hinzugerechnet werden
                        
                        ' größe der Palettendaten zur gesamten
                        ' größe der Bitmap hinzurechnen
                        tBITMAPFILEHEADER.bfSize = _
                            tBITMAPFILEHEADER.bfSize + (lngPalCount * 4)
                            
                        ' größe der Palettendaten zum Offset hinzurechnen
                        tBITMAPFILEHEADER.bfOffBits = _
                            tBITMAPFILEHEADER.bfOffBits + (lngPalCount * _
                            4)
                            
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
