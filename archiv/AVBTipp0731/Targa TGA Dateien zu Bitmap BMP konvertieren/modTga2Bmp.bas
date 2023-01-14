Attribute VB_Name = "modTga2Bmp"
Option Explicit

' ----==== Const ====----
Private Const S_OK As Long = 0&
Private Const DIB_RGB_COLORS As Long = 0&
Private Const BI_RGB As Long = 0&
Private Const IID_IPicture As String = "{7BF80980-BF32-101A-8BBB-00AA00300CAB}"

' Flag zum erkennen, ob es sich bei den TGA-Typen 9, 10 und 11
' in den Bilddaten um RAW- oder RLE-kodierte Daten handelt
Private Const RleFlag As Long = &H80

' Vertikal- und Horizontalflag zum spiegeln des Bildes
Private Const VFlag As Long = &H10
Private Const HFlag As Long = &H20

' ----==== Type ====----
Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
    hgdiObj As Long
    hPalOrXYExt As Long
End Type

Private Type IID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7)  As Byte
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

Private Type ARGB
    Blue As Byte
    Green As Byte
    Red As Byte
    Alpha As Byte
End Type

Private Type BITMAPINFO256
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As ARGB
End Type

Private Type TgaHeader
    IdentSize As Byte
    ColorMapType As Byte
    ImageType As Byte
    ColorMapStart As Integer
    ColorMapLength As Integer
    ColorMapBits As Byte
    xStart As Integer
    yStart As Integer
    Width As Integer
    Height As Integer
    Bits As Byte
    Descriptor As Byte
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
                         
Private Declare Function SetDIBits256 Lib "gdi32.dll" _
                         Alias "SetDIBits" ( _
                         ByVal hdc As Long, _
                         ByVal hBitmap As Long, _
                         ByVal nStartScan As Long, _
                         ByVal nNumScans As Long, _
                         ByRef lpBits As Any, _
                         ByRef lpBI As BITMAPINFO256, _
                         ByVal wUsage As Long) As Long
                         
' ----==== KERNEL32 API Deklarationen ====----
Private Declare Sub CopyMemory Lib "kernel32.dll" _
                    Alias "RtlMoveMemory" ( _
                    ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
                    
' ----==== OLE32 API Declarationen ====----
Private Declare Function IIDFromString Lib "ole32.dll" ( _
                         ByVal lpsz As Long, _
                         ByRef lpIID As IID) As Long
                         
' ----==== OLEOUT32 API Declarations ====----
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
                         ByRef lpPictDesc As PICTDESC, _
                         ByRef riid As IID, _
                         ByVal fOwn As Boolean, _
                         ByRef lplpvObj As Object) As Long
                         
' ----==== USER32 API Deklarationen ====----
Private Declare Function GetDC Lib "user32.dll" ( _
                         ByVal hwnd As Long) As Long
                         
Private Declare Function ReleaseDC Lib "user32.dll" ( _
                         ByVal hwnd As Long, _
                         ByVal hdc As Long) As Long
                         
' ------------------------------------------------------
' Funktion     : ConvertTga2Bmp
' Beschreibung : konvertiert eine TGA-Datei in ein StdPicture
' Übergabewert : TgaFile = Pfad\Datei.ext
' Rückgabewert : StdPicture
' ------------------------------------------------------
' zur Zeit unterstützt die Funktion folgende Targa-Formate:
' Imagetyp: 1, 2, 3, 9, 10, 11
' Bits: 8, 16, 24, 32
' ColorMapBits: 24, 32
' ------------------------------------------------------
Public Function ConvertTga2Bmp(ByVal TgaFile As String) As StdPicture

    ' div. Variablen
    Dim X As Long
    Dim Y As Long
    Dim lngDC As Long
    Dim lngFNr As Long
    Dim hConvBmp As Long
    Dim PalIndex As Long
    Dim BmpWidth As Long
    Dim BmpHeight As Long
    Dim BmpStride As Long
    Dim TgaPixPos As Long
    Dim BmpPixPos As Long
    Dim BytePerPixel As Long
    Dim BmpPixelFormat As Long
    Dim RleID As Byte
    Dim TgaPal() As Byte
    Dim TgaData() As Byte
    Dim BmpData() As Byte
    Dim lngReadByte() As Byte
    Dim NoPadBytes As Boolean
    Dim tTgaHeader As TgaHeader
    Dim tBITMAPINFO As BITMAPINFO256
    
    ' ist die Datei nicht vorhanden
    If Not FileExists(TgaFile) Then
    
        ' dann aus der Funktion aussteigen
        Exit Function
        
    Else
    
        ' ist die Dateierweiterung <> .TGA
        If UCase$(Right$(TgaFile, 4)) <> ".TGA" Then
        
            ' dann aus der Funktion aussteigen
            Exit Function
            
        End If
    End If
    
    ' freie Dateinummer holen
    lngFNr = FreeFile
    
    ' Datei binär einlesen
    Open TgaFile For Binary Access Read As #lngFNr
    
    ' Header aus der TGA auslesen
    Get #lngFNr, , tTgaHeader
    
    ' Breite und Höhe des Bildes speichern
    BmpWidth = tTgaHeader.Width
    BmpHeight = tTgaHeader.Height
    
    ' ist die Breite des Bildes ohne Rest durch 4 teilbar
    ' oder ist es eine 32bpp-TGA
    If BmpWidth Mod 4 = 0 Or tTgaHeader.Bits = 32 Then
    
        ' dann gibt es keine PadBytes in der zu erstellenden Bitmap
        NoPadBytes = True
        
    End If
    
    ' nach TGA-ImageTyp selektieren
    Select Case tTgaHeader.ImageType
    
    Case 1, 2, 3, 9, 10, 11
    
        '  1 = Unkomprimiert, Indexed
        '  2 = Unkomprimiert, RGB
        '  3 = Unkomprimiert, Grauskale
        '  9 = RLE enkodiert, Indexed
        ' 10 = RLE enkodiert, RGB
        ' 11 = RLE enkodiert, Grauskale
        ' nach Anzahl der Bits per Pixel selektieren
        Select Case tTgaHeader.Bits
        
        Case 8
        
            ' Byte pro Pixel
            BytePerPixel = 1
            
            ' Breite einer Bildzeile inkl. PadBytes berechnen
            ' für die zu erstellende Bitmap
            BmpStride = (BmpWidth + 3) And Not 3
            
            ' Pixelformat für die zu erstellende Bitmap festlegen
            BmpPixelFormat = 8
            
        Case 16
            BytePerPixel = 2
            BmpStride = ((BmpWidth * 2) + 2) And Not 2
            BmpPixelFormat = 16
            
        Case 24
            BytePerPixel = 3
            BmpStride = ((BmpWidth * 3) + 3) And Not 3
            BmpPixelFormat = 24
            
        Case 32
            BytePerPixel = 4
            BmpStride = BmpWidth * 4
            BmpPixelFormat = 32
            
        Case Else
        
            ' andere
            BmpPixelFormat = 0
            
        End Select
        
        ' wenn die zu erstellende Bitmap PadBytes hat, dann brauchen
        ' wir BmpData zum späteren umkopieren von TgaData
        If Not NoPadBytes Then
        
            ' Größe des Arrays BmpData zur Aufnahme der Bilddaten
            ' für die zu erstellende Bitmap (OutBmp) berechnen
            ' und dimensionieren wenn PadBytes vorhanden sind
            ReDim BmpData((BmpHeight * BmpStride) - 1)
            
        End If
        
        ' Größe des Arrays TgaData zur Aufnahme der Bilddaten
        ' aus der TGA berechnen und dimensionieren
        ReDim TgaData((BmpHeight * (BmpWidth * BytePerPixel)) - 1)
        
    End Select
    
    ' Ist tTgaHeader.IdentSize > 0 dann folgt direkt nach dem Header
    ' ein Identblock in der Größe von tTgaHeader.IdentSize. Da wir
    ' diesen nicht benötigen, überspringen wir diesen Block.
    Seek #lngFNr, Seek(lngFNr) + tTgaHeader.IdentSize
    
    ' Direkt nach dem Header und/oder nach dem IdentBlock wenn
    ' vorhanden, kommen die Palettendaten wenn vorhanden.
    ' enthält die TGA Palettendaten
    If tTgaHeader.ColorMapType = 1 Then
    
        ' Größe des Arrays TgaPal zur Aufnahme der Palettendaten
        ' berechnen und dimensionieren
        ReDim TgaPal((tTgaHeader.ColorMapLength * (tTgaHeader.ColorMapBits / 8)) - 1)
        
        ' Palettendaten aus der TGA auslesen
        Get #lngFNr, , TgaPal
        
        ' alle Paletteneinträge aus der TGA, die wir zuvor in TgaPal
        ' eingelesen haben, in die Palette für die Bitmap umkopieren
        For PalIndex = tTgaHeader.ColorMapStart To tTgaHeader.ColorMapLength - 1
        
            ' nach Anzahl der Bits per Pixel in der
            ' Palette selektieren
            Select Case tTgaHeader.ColorMapBits
            
            Case 24
            
                ' Palettendaten umkopieren
                With tBITMAPINFO.bmiColors(PalIndex)
                
                    .Alpha = 255
                    .Red = TgaPal((PalIndex * 3) + 2)
                    .Green = TgaPal((PalIndex * 3) + 1)
                    .Blue = TgaPal((PalIndex * 3) + 0)
                    
                End With
                
            Case 32
            
                With tBITMAPINFO.bmiColors(PalIndex)
                
                    .Alpha = TgaPal((PalIndex * 4) + 3)
                    .Red = TgaPal((PalIndex * 4) + 2)
                    .Green = TgaPal((PalIndex * 4) + 1)
                    .Blue = TgaPal((PalIndex * 4) + 0)
                    
                End With
                
            End Select
            
        Next PalIndex
        
        ' Anzahl der verwendeten Farben in der Palette
        tBITMAPINFO.bmiHeader.biClrUsed = tTgaHeader.ColorMapLength
        
        ' Anzahl der verwendeten Farben in der Palette
        tBITMAPINFO.bmiHeader.biClrImportant = tTgaHeader.ColorMapLength
        
    Else
    
        ' nach TGA-ImageTyp selektieren
        Select Case tTgaHeader.ImageType
        
        Case 3, 11 ' nur Typ 3 und 11
        
            ' eine eigene Palette erstellen (Grauskale)
            For PalIndex = 0 To 255
            
                With tBITMAPINFO.bmiColors(PalIndex)
                
                    .Alpha = 255
                    .Red = PalIndex
                    .Green = PalIndex
                    .Blue = PalIndex
                    
                End With
                
            Next PalIndex
            
            ' Anzahl der verwendeten Farben in der Palette
            tBITMAPINFO.bmiHeader.biClrUsed = 256
            
            ' Anzahl der verwendeten Farben in der Palette
            tBITMAPINFO.bmiHeader.biClrImportant = 256
            
        End Select
        
    End If
    
    ' nach TGA-ImageTyp selektieren
    Select Case tTgaHeader.ImageType
    
    Case 1, 2, 3 ' nur die unkomprimierten TGA-Typen
    
        ' komplette Bilddaten aus der TGA auslesen
        Get #lngFNr, , TgaData
        
    Case 9, 10, 11 ' nur die komprimierten TGA-Typen (RAW/RLE)
    
        ' Da wir durch die RLE-Komprimierung nicht wissen wieviel
        ' Bytes an Bitmapdaten wir einlesen müssen, lesen wir
        ' solange die Daten ein bis UBound(TgaData)
        ' (unkomprimierte Größe) erreicht ist.
        For X = 0 To UBound(TgaData) - 1
        
            ' PacketHeader-Byte aus der TGA lesen
            Get #lngFNr, , RleID
            
            ' ist das Bit 8 von RleID = 1 dann liegt das folgende
            ' Datenpaket RLE-Komprimiert vor
            If CBool(RleID And RleFlag) Then
            
                ' In (RleID - RleFlag) steht die Anzahl der
                ' Wiederholungen - 1
                RleID = (RleID - RleFlag) + 1
                
                ' entsprechende Anzahl von Bytes aus der
                ' TGA auslesen und direkt nach TgaData an
                ' Offset X kopieren
                ReDim lngReadByte(BytePerPixel - 1)
                Get #lngFNr, , lngReadByte
                
                Call CopyMemory(TgaData(X), lngReadByte(0), BytePerPixel)
                
                ' nun kopieren wir die ausgelesenen Bytes
                ' entsprechend der Wiederholungen
                ' hintereinander
                For Y = 1 To RleID - 1
                
                    Call CopyMemory(TgaData(X + (Y * BytePerPixel)), TgaData(X), _
                        BytePerPixel)
                        
                Next Y
                
            Else
            
                ' ist das Bit 8 von RleID = 0 dann liegt das
                ' folgende Datenpaket unkomprimiert vor (RAW).
                ' RleID enthält die Anzahl der Pixel - 1.
                RleID = RleID + 1
                
                ' entsprechende Anzahl von Bytes aus der
                ' TGA auslesen und direkt nach TgaData an
                ' Offset X kopieren
                ReDim lngReadByte((RleID * BytePerPixel) - 1)
                Get #lngFNr, , lngReadByte
                
                Call CopyMemory(TgaData(X), lngReadByte(0), RleID * BytePerPixel)
                
            End If
            
            ' X = X + Offset
            X = X + (RleID * BytePerPixel) - 1
            
        Next X
        
    Case Else
    
        ' andere
    End Select
    
    ' Zugriff auf die Datei schließen
    Close #lngFNr
    
    ' Wurde kein entsprechendes Pixelformat festgelegt, dann haben wir es
    ' hier mit einem nicht implementierten TGA-Format zu tun. Also können
    ' wir auch aus dem Rest des Codes aussteigen.
    If BmpPixelFormat = 0 Then
    
        ' dann Nothing zurück geben
        Exit Function
        
    End If
    
    ' wenn die zu erstellende Bitmap PadBytes hat, dann müssen
    ' wir die Bilddaten pixelweise umkopieren.
    If Not NoPadBytes Then
    
        ' Da TGAs keine PadBytes haben aber Bitmaps schon, müssen
        ' wir die Bilddaten aus dem Array TgaData (TGA-Bilddaten)
        ' nach BmpData (BMP-Bilddaten) umkopieren.
        For Y = 0 To BmpHeight - 1
            For X = 0 To BmpWidth - 1
            
                ' Pixelposition für BmpData berechnen
                BmpPixPos = (Y * BmpStride) + (X * BytePerPixel)
                
                ' Pixelposition für TgaData berechnen
                TgaPixPos = (Y * (BmpWidth * BytePerPixel)) + (X * BytePerPixel)
                
                ' Pixeldaten von TgaData nach BmpData umkopieren
                Call CopyMemory(BmpData(BmpPixPos), TgaData(TgaPixPos), BytePerPixel)
                
            Next X
        Next Y
        
    End If
    
    ' DC ermitteln
    lngDC = GetDC(0&)
    
    ' ist ein DC vorhanden
    If lngDC <> 0 Then
    
        tBITMAPINFO.bmiHeader.biSize = Len(tBITMAPINFO.bmiHeader)
        
        ' Screen(destination)|Image(Origin)
        '  of first pixel    | bit 5 | bit 4
        ' -------------------|-------------
        ' Bottom(Left)       |   0   |   0
        ' Bottom(Right)      |   0   |   1
        ' Top(Left)          |   1   |   0
        ' Top(Right)         |   1   |   1
        ' ist das Bit 4 vom tTgaHeader.descriptor = 1
        If CBool(tTgaHeader.Descriptor And VFlag) Then
        
            ' dann vertikal spiegeln
            tBITMAPINFO.bmiHeader.biWidth = -BmpWidth
            
        Else
        
            ' nicht vertikal spiegeln
            tBITMAPINFO.bmiHeader.biWidth = BmpWidth
            
        End If
        
        ' ist das Bit 5 vom tTgaHeader.descriptor = 1
        If CBool(tTgaHeader.Descriptor And HFlag) Then
        
            ' dann horizontal spiegeln
            tBITMAPINFO.bmiHeader.biHeight = -BmpHeight
            
        Else
        
            ' nicht horizontal spiegeln
            tBITMAPINFO.bmiHeader.biHeight = BmpHeight
            
        End If
        
        tBITMAPINFO.bmiHeader.biPlanes = 1
        tBITMAPINFO.bmiHeader.biBitCount = BmpPixelFormat
        tBITMAPINFO.bmiHeader.biCompression = BI_RGB
        
        ' wenn keine PadBytes vorhanden sind
        If NoPadBytes Then
        
            ' TgaData verwenden
            tBITMAPINFO.bmiHeader.biSizeImage = UBound(TgaData) + 1
            
        Else
        
            ' wenn PadBytes vorhanden sind
            ' BmpData verwenden
            tBITMAPINFO.bmiHeader.biSizeImage = UBound(BmpData) + 1
            
        End If
        
        ' DIB-Bitmap erstellen
        hConvBmp = CreateDIBSection256(lngDC, tBITMAPINFO, DIB_RGB_COLORS, 0&, 0&, 0&)
            
        ' ist ein DIB-Bitmap vorhanden
        If hConvBmp <> 0 Then
        
            ' wenn keine PadBytes vorhanden sind
            If NoPadBytes Then
            
                ' TgaData in die DIB-Bitmap schreiben
                If SetDIBits256(lngDC, hConvBmp, 0&, BmpHeight, TgaData(0), _
                    tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                    
                    ' DIB-Bitmap in ein StdPicture konvertieren
                    Set ConvertTga2Bmp = HandleToPicture(hConvBmp)
                    
                End If
                
            Else
            
                ' wenn PadBytes vorhanden sind
                ' BmpData in die DIB-Bitmap schreiben
                If SetDIBits256(lngDC, hConvBmp, 0&, BmpHeight, BmpData(0), _
                    tBITMAPINFO, DIB_RGB_COLORS) <> 0 Then
                    
                    ' DIB-Bitmap in ein StdPicture konvertieren
                    Set ConvertTga2Bmp = HandleToPicture(hConvBmp)
                    
                End If
            End If
        End If
        
        ' DC freigeben
        Call ReleaseDC(0&, lngDC)
        
    End If
    
End Function

' ------------------------------------------------------
' Funktion     : FileExists
' Beschreibung : Ermittelt ob eine Datei vorhanden ist
' Übergabewert : FileName = Pfad\Dateiname.ext
' Rückgabewert : True = Datei vorhanden
'                False = Datei nicht vorhanden
' ------------------------------------------------------
Private Function FileExists(ByVal FileName As String) As Boolean

    On Error Resume Next
    
    Dim ret As Long
    
    ret = Len(Dir$(FileName))
    
    If Err Or ret = 0 Then FileExists = False Else FileExists = True
    
End Function

' ------------------------------------------------------
' Funktion     : HandleToPicture
' Beschreibung : Bitmap Handle in ein StdPicture Objekt umwandeln
' Übergabewert : hGDIHandle = Bitmap Handle
' Rückgabewert : StdPicture Objekt
' ------------------------------------------------------
Private Function HandleToPicture(ByVal hGDIHandle As Long) As StdPicture

    Dim tIID As IID
    Dim tPictDesc As PICTDESC
    Dim oPicture As IPicture
    
    ' IID_IPicture -> tIID
    If IIDFromString(StrPtr(IID_IPicture), tIID) = S_OK Then
    
        ' Initialisiert die PICTDESC Structur
        With tPictDesc
        
            .cbSizeOfStruct = Len(tPictDesc)
            .picType = vbPicTypeBitmap
            .hgdiObj = hGDIHandle
            
        End With
        
        ' StdPicture (Icon) aus dem Handle erstellen
        If OleCreatePictureIndirect(tPictDesc, tIID, True, oPicture) = S_OK Then
        
            ' Rückgabe des Pictureobjekts
            Set HandleToPicture = oPicture
            
        End If
    End If
    
End Function

