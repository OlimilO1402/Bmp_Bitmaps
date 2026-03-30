VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Error(ErrorNumber As Long, ErrorText As String)

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" _
        (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, _
        pccolorref As Long) As Long
        
Private Const CLR_INVALID = -1

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal _
        hDC As Long, ByVal x As Long, ByVal y As Long, ByVal _
        dx As Long, ByVal dy As Long, ByVal SrcX As Long, _
        ByVal SrcY As Long, ByVal Scan As Long, ByVal _
        NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, _
        ByVal wUsage As Long) As Long
        
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As _
        Long) As Long
        
Private Declare Function DeleteObject Lib "gdi32" (ByVal _
        hObject As Long) As Long
        
Private Declare Function CreateCompatibleDC Lib "gdi32" _
       (ByVal hDC As Long) As Long
       
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal _
        hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As _
        Long, ByVal lplpVoid As Long, ByVal handle As Long, _
        ByVal dw As Long) As Long
        
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC _
        As Long, ByVal hObject As Long) As Long
        
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As _
        Long, ByVal x As Long, ByVal y As Long, ByVal nWidth _
        As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop _
        As Long) As Long
        
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As _
        Long, ByVal hBitmap As Long, ByVal nStartScan As _
        Long, ByVal nNumScans As Long, lpBits As Any, lpBI As _
        BITMAPINFO, ByVal wUsage As Long) As Long
        
Private Declare Function GetDC Lib "user32" (ByVal hwnd As _
        Long) As Long
        
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As _
        Long, ByVal hDC As Long) As Long
        
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC _
        As Long, ByVal DestX As Long, ByVal DestY As Long, ByVal _
        DestWidth As Long, ByVal DestHeight As Long, ByVal SrcX _
        As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, _
        ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As _
        BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) _
        As Long

Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function GetQueueStatus Lib "user32" (ByVal _
        fuFlags As Long) As Long

Private Const WM_SETREDRAW = &HB
                     
'Modi zur Felddimensionierung
Public Enum DimConstants
    dibNew = 1
    dibPreserve
End Enum

'Sprachauswahl
Public Enum LanguageConstants
    dibGerman = 1
    dibEnglish
End Enum
       
'Modi, Fehlerbetrachtung oder Durchmarsch
Public Enum RunTimeConstants
    dibSecure = 1
    dibSpeed
End Enum

'Filter für die Nachrichtenbetrachtung
Public Enum QueueStatusConstants
    qsKey = &H1
    qsMouseMove = &H2
    qsMouseButton = &H4
    qsPostMessage = &H8
    qsTimer = &H10
    qsPaint = &H20
    qsSendMessage = &H40
    qsHotKey = &H80
    qsMouse = (qsMouseMove Or qsMouseButton)
    qsInput = (qsMouse Or qsKey)
    qsAllEvents = (qsInput Or qsPostMessage Or qsTimer Or qsPaint Or qsHotKey)
    qsAllInputs = (qsSendMessage Or qsPaint Or qsTimer Or qsPostMessage Or qsMouseButton Or qsMouseMove Or qsHotKey Or qsKey)
End Enum

Public Enum QueueFilterConstants
    qsAll = 1
    QsNew = 2
End Enum

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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const MCI_NOTIFY& = &H3B9
Private Const MCI_NOTIFY_SUCCESS& = &H1
Private Const GWL_WNDPROC = -4
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80

Private m_iBitmap As Long, m_iDC As Long, m_bi24BitInfo As BITMAPINFO

Private m_Width As Long, m_Height As Long, m_RunTime As RunTimeConstants

Private m_hWnd_Destination As Long, m_hWnd_Source As Long, m_Language As LanguageConstants
Private m_Initialized As Boolean, m_RedrawState As Boolean

Private m_LastUx As Long, m_LastUy As Long

'Weite der DIB lesen
Public Property Get Width() As Long
    On Error GoTo Err_GetWidth
    
    Width = m_Width
    
    Exit Property
    
Err_GetWidth:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetWidth")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Höhe der DIB lesen
Public Property Get Height() As Long
    On Error GoTo Err_GetHeight
    
    Height = m_Height
    
    Exit Property
    
Err_GetHeight:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetHeight")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Weite der DIB festlegen
Public Property Let Width(n_Width As Long)
    On Error GoTo Err_LetWidth
    
    m_Width = n_Width
    
    Exit Property
    
Err_LetWidth:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.LetWidth")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Höhe der DIB festlelgen
Public Property Let Height(n_Height As Long)
    On Error GoTo Err_LetHeight
    
    m_Height = n_Height
    
    Exit Property
    
Err_LetHeight:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.LetHeight")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'hWnd des Ziel-DCs festlegen
Public Property Let hWndDestination(n_hWnd As Long)
    On Error GoTo Err_LethWndDestination
    
    m_hWnd_Destination = n_hWnd
    
    Exit Property
    
Err_LethWndDestination:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.LethWndDestination")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'hWnd des Ziel-DCs auslesen
Public Property Get hWndDestination() As Long
    On Error GoTo Err_GethWndDestination
    
    hWndDestination = m_hWnd_Destination
    
    Exit Property
    
Err_GethWndDestination:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GethWndDestination")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'hWnd des Quell-DCs festlegen
Public Property Let hWndSource(n_hWnd As Long)
    On Error GoTo Err_LethWndSource
    
    m_hWnd_Source = n_hWnd
    
    Exit Property
    
Err_LethWndSource:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.hWndSource")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'hWnd des Quell-DCs lesen
Public Property Get hWndSource() As Long
    On Error GoTo Err_GethWndSource
    
    hWndSource = m_hWnd_Source
    
    Exit Property
    
Err_GethWndSource:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.hWndSource")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'RuntimeMode festlegen (tdSpeed und tdSecure)
Public Property Let RuntimeMode(n_RunTime As RunTimeConstants)
    On Error GoTo Err_LetRuntimeMode
    
    m_RunTime = n_RunTime
    
    Exit Property
    
Err_LetRuntimeMode:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.LetRuntimeMode")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'RuntimeMode auslesen (tdSpeed und tdSecure)
Public Property Get RuntimeMode() As RunTimeConstants
    On Error GoTo Err_GetRuntimeMode
    
    RuntimeMode = m_RunTime
    
    Exit Property
    
Err_GetRuntimeMode:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetRuntimeMode")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Sprachauswahl festlegen
Public Property Let Language(n_Language As LanguageConstants)
    On Error GoTo Err_LetLanguage
    
    m_RunTime = n_Language
    
    Exit Property
    
Err_LetLanguage:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.LetLanguage")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Sprachauswahl auslesen
Public Property Get Language() As LanguageConstants
    On Error GoTo Err_GetLanguage
    
    Language = m_Language
    
    Exit Property
    
Err_GetLanguage:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetLanguage")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Lock-Windows-Update ent- oder verriegeln
Public Property Let SetRedraw(State As Boolean)
    On Error GoTo Err_SetRedraw
    
    If m_hWnd_Destination <> 0 Then
        If State Then
            Call SendMessage(m_hWnd_Destination, WM_SETREDRAW, 1&, 0&)
        Else
            Call SendMessage(m_hWnd_Destination, WM_SETREDRAW, 0&, 0&)
        End If
        m_RedrawState = State
    Else
        Call HandleError(1)
    End If
    
    Exit Property
    
Err_SetRedraw:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.SetRedraw")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Lock-Windows-Update auslesen
Public Property Get SetRedraw() As Boolean
    On Error GoTo Err_GetSetRedraw
    
    SetRedraw = m_RedrawState
    
    Exit Property
    
Err_GetSetRedraw:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetSetRedraw")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Abfragemöglichkeit für in der WindowsMain stehende Nachrichten mit Filter
Public Function MessagesQueued(Messages As QueueStatusConstants, Mode As QueueFilterConstants) As Long
    On Error GoTo Err_MessagesQueued
    
    Dim x As Long, Hi As Long, Lo As Long
    
        x = GetQueueStatus(Messages)
        Lo = x And &HFF
        Hi = (x And &HFF00) / &HF00
    
        If Mode = qsAll Then
            MessagesQueued = Hi
        Else
            MessagesQueued = Lo
        End If
    
        Exit Function
    
Err_MessagesQueued:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.MessagesQueued")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function


'Schaun ob die DIB bereits initialisiert ist
Public Property Get IsCreated() As Boolean
    On Error GoTo Err_GetIsCreated
    
    IsCreated = m_Initialized
    
    Exit Property
    
Err_GetIsCreated:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetIsCreated")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'DIB initialisieren
Public Function Create() As Boolean
    On Error GoTo Err_Create
    
    Dim bWidth As Long, bHeight As Long, w As Long
    Dim n_hDC As Long, Result As Long
  
        If m_Initialized Then
            'DIB ist noch nicht intialisiert
            Call HandleError(4)
        Else
            If m_hWnd_Destination <> 0 Then
                'DC des Zielbildes festnageln
                n_hDC = GetDC(m_hWnd_Destination)
                If n_hDC <> 0 Then
                    
                    'Weite der DIB muß inter durch vier teilbar sein
                    w = m_Width - 1
                    If w Mod 4 <> 0 Then
                        bWidth = Int(w / 4) * 4 + 4
                    Else
                        bWidth = Width
                    End If
                    
                    bHeight = m_Height - 1
                    
                    'Bitmapdaten speichern
                    With m_bi24BitInfo.bmiHeader
                        .biBitCount = 24
                        .biCompression = BI_RGB
                        .biPlanes = 1
                        .biSize = Len(m_bi24BitInfo.bmiHeader)
                        .biWidth = bWidth
                        .biHeight = bHeight + 1
                    End With
            
                    'DC erstellen
                    m_iDC = CreateCompatibleDC(n_hDC)
                    
                    'DIB erstellen
                    m_iBitmap = CreateDIBSection(m_iDC, m_bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
            
                    'DIB in den DC wählen
                    Call SelectObject(m_iDC, m_iBitmap)
            
                    'DC des Zielbildes wieder freigeben
                    Result = ReleaseDC(m_hWnd_Destination, n_hDC)
                    If Result <> 0 Then
                        m_Initialized = True
                        Create = True
                    Else
                        Call HandleError(3)
                    End If
    
                Else
                    Call HandleError(2)
                End If
            Else
                Call HandleError(1)
            End If
        End If
    
        Exit Function
    
Err_Create:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.Create")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'Grafik eines sichtbaren DCs in ein Byte-Array kopieren
Public Function CopyBitmapToArray(bArray() As Byte, _
                                  Optional x1_Dst As Long, Optional y1_Dst As Long, _
                                  Optional w_Dst As Long, Optional h_Dst As Long, _
                                  Optional x1_Src As Long, Optional y1_Src As Long, _
                                  Optional RasterOperation As RasterOpConstants = vbSrcCopy) As Boolean
                                  
    On Error GoTo Err_CopyBitmapToArray
                                      
    Dim w As Long, h As Long, n_hDC As Long, Result As Long, x As Long, y As Long
    
        If m_RunTime = dibSpeed Then
            'Quell-DC festnageln
            n_hDC = GetDC(m_hWnd_Source)
            
            'Quellbild in den Speicher-DC kopieren, ROPs anwenden
            Call BitBlt(m_iDC, x1_Dst, y1_Dst, w_Dst, h_Dst, n_hDC, x1_Src, y1_Src, RasterOperation)
                                      
            'Bytes aus dem Speicher-DC in das Array holen
            Call GetDIBits(n_hDC, m_iBitmap, 0, m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), _
                           m_bi24BitInfo, DIB_RGB_COLORS)
            
            'Quell-DC wieder freigeben
            Call ReleaseDC(m_hWnd_Source, n_hDC)
            CopyBitmapToArray = True
        Else
            If Not m_Initialized Then
                Call HandleError(14)
            Else
            
                'Auf Plausibilität der Weite und Höhe der Ziel-DIB testen
                If w_Dst = 0 Then
                    w = m_Width
                Else
                    w = w_Dst
                    If w > m_Width Then w = m_Width
                End If
                
                
                If h_Dst = 0 Then
                    h = m_Height
                Else
                    h = h_Dst
                    If h > m_Height Then h = m_Height
                End If
                
                If w > 0 And h > 0 Then
                
                    'Auf Plausibilität der Auschnittsabmaße der Ziel-DIB testen
                    If x1_Dst > m_Width - 1 Then
                        Call HandleError(8)
                    ElseIf y1_Dst > m_Height - 1 Then
                        Call HandleError(9)
                    ElseIf x1_Dst + w > m_Width Then
                        Call HandleError(10)
                    ElseIf y1_Dst + h > m_Height Then
                        Call HandleError(11)
                    Else
                        If m_hWnd_Source <> 0 Then
                        
                            'Quell-DC festnageln
                            n_hDC = GetDC(m_hWnd_Source)
                            If n_hDC <> 0 Then
                            
                                'Quellbild in den Speicher-DC kopieren, ROPs anwenden
                                Call BitBlt(m_iDC, x1_Dst, y1_Dst, w, h, n_hDC, x1_Src, y1_Src, RasterOperation)
                                
                                'Bytes aus dem Speicher-DC in das Array holen
                                Call GetDIBits(n_hDC, m_iBitmap, 0, m_bi24BitInfo.bmiHeader.biHeight, _
                                               bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS)
                                               
                                'Quell-DC wieder freigeben
                                Result = ReleaseDC(m_hWnd_Source, n_hDC)
                                If Result <> 0 Then
                                    CopyBitmapToArray = True
                                Else
                                    Call HandleError(17)
                                End If
                            Else
                                Call HandleError(16)
                            End If
                        Else
                            Call HandleError(15)
                        End If
                    End If
                Else
                    If h < 1 Then
                        Call HandleError(6)
                    Else
                        Call HandleError(7)
                    End If
                End If
            End If
        End If
    
        Exit Function
    
Err_CopyBitmapToArray:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.CopyBitmapToArray")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'ByteArray auf einen DC kopieren, Weiten und Höhenanderungen können angepaßt werden
Public Function StretchBitBlit(bArray() As Byte, _
                        Optional ByVal x1_Src As Long, Optional ByVal y1_Src As Long, _
                        Optional ByVal w_Src As Long, Optional ByVal h_Src As Long, _
                        Optional ByVal x1_Dst As Long, Optional ByVal y1_Dst As Long, _
                        Optional ByVal w_Dst As Long, Optional ByVal h_Dst As Long, _
                        Optional ByVal RasterOperation As RasterOpConstants = vbSrcCopy) As Boolean
                
    On Error GoTo Err_StretchBitBlit
                    
    Dim h As Long, w As Long, n_hDC As Long, Result As Long
        
        If m_RunTime = dibSpeed Then
            'Ziel-DC festnageln
            n_hDC = GetDC(m_hWnd_Destination)
            
            'ByteArray auf DC ausgeben
            Call StretchDIBits(n_hDC, x1_Dst, y1_Dst, w_Dst, h_Dst, x1_Src, y1_Src, w_Src, h_Src, _
                               bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS, RasterOperation)
            
            'Ziel-DC wieder freigeben
            Call ReleaseDC(m_hWnd_Destination, n_hDC)
        Else
        
            If Not m_Initialized Then
                Call HandleError(14)
            Else
                'Auf Plausibilität der Weite und Höhe der Quell-DIB testen
                If w_Src = 0 Then
                    w = m_Width
                Else
                    w = w_Src
                    If w > m_Width Then w = m_Width
                End If
                
                
                If h_Src = 0 Then
                    h = m_Height
                Else
                    h = h_Src
                    If h > m_Height Then h = m_Height
                End If
                
                If w > 0 And h > 0 Then
                    'Auf Plausibilität der Auschnittsabmaße der Quell-DIB testen
                    If x1_Src > m_Width - 1 Then
                        Call HandleError(8)
                    ElseIf y1_Src > m_Height - 1 Then
                        Call HandleError(9)
                    ElseIf x1_Src + w > m_Width Then
                        Call HandleError(10)
                    ElseIf y1_Src + h > m_Height Then
                        Call HandleError(11)
                    Else
                        If m_hWnd_Destination <> 0 Then
                            'Ziel-DC festnageln
                            n_hDC = GetDC(m_hWnd_Destination)
                            If n_hDC <> 0 Then
                            
                                If w_Dst = 0 Then w_Dst = w
                                If h_Dst = 0 Then h_Dst = h
                                                                
                                'ByteArray auf DC ausgeben
                                Call StretchDIBits(n_hDC, x1_Dst, y1_Dst, w_Dst, h_Dst, x1_Src, y1_Src, w, h, bArray(0, 0), _
                                                   m_bi24BitInfo, DIB_RGB_COLORS, RasterOperation)
                            
                                'Ziel-DC wieder freigeben
                                Result = ReleaseDC(m_hWnd_Destination, n_hDC)
                                
                                If Result = 0 Then
                                    Call HandleError(3)
                                Else
                                    StretchBitBlit = True
                                End If
                            Else
                                Call HandleError(2)
                            End If
                        Else
                            Call HandleError(1)
                        End If
                    End If
                Else
                    If h < 1 Then
                        Call HandleError(6)
                    Else
                        Call HandleError(7)
                    End If
                End If
            End If
        End If
    
        Exit Function
    
Err_StretchBitBlit:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.StretchBitBlit")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

Public Function BitBlit(bArray() As Byte, _
                        Optional x1_Src As Long, Optional y1_Src As Long, _
                        Optional w_Src As Long, Optional h_Src As Long, _
                        Optional x1_Dst As Long, Optional y1_Dst As Long) As Boolean
                
    On Error GoTo Err_BitBlit
                                        
    Dim h As Long, w As Long, n_hDC As Long, Result As Long
        
        If m_RunTime = dibSpeed Then
            'Ziel-DC festnageln
            n_hDC = GetDC(m_hWnd_Destination)
            
            'ByteArray auf DC ausgeben
            Call SetDIBitsToDevice(n_hDC, x1_Dst, y1_Dst, w_Src, h_Src, x1_Src, y1_Src, 0, _
                                   m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), _
                                   m_bi24BitInfo, DIB_RGB_COLORS)
                                   
            'Ziel-DC freigeben
            Call ReleaseDC(m_hWnd_Destination, n_hDC)
        Else
            If Not m_Initialized Then
                Call HandleError(14)
            Else
                'Auf Plausibilität der Weite und Höhe der Quell-DIB testen
                If w_Src = 0 Then
                    w = m_Width
                Else
                    w = w_Src
                    If w > m_Width Then w = m_Width
                End If
                
                
                If h_Src = 0 Then
                    h = m_Height
                Else
                    h = h_Src
                    If h > m_Height Then h = m_Height
                End If
                
                If w > 0 And h > 0 Then
                    'Auf Plausibilität der Auschnittsabmaße der Quell-DIB testen
                    If x1_Src > m_Width - 1 Then
                        Call HandleError(8)
                    ElseIf y1_Src > m_Height - 1 Then
                        Call HandleError(9)
                    ElseIf x1_Src + w > m_Width Then
                        Call HandleError(10)
                    ElseIf y1_Src + h > m_Height Then
                        Call HandleError(11)
                    Else
                        If m_hWnd_Destination <> 0 Then
                            'Ziel-DC festnageln
                            n_hDC = GetDC(m_hWnd_Destination)
                            If n_hDC <> 0 Then
                            
                                'ByteArray auf DC ausgeben
                                Call SetDIBitsToDevice(n_hDC, x1_Dst, y1_Dst, w, h, x1_Src, y1_Src, 0, m_bi24BitInfo.bmiHeader.biHeight, _
                                                       bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS)
                            
                                'Ziel-DC freigeben
                                Result = ReleaseDC(m_hWnd_Destination, n_hDC)
                                
                                If Result = 0 Then
                                    Call HandleError(2)
                                Else
                                    BitBlit = True
                                End If
                            Else
                                Call HandleError(2)
                            End If
                        Else
                            Call HandleError(1)
                        End If
                    End If
                Else
                    If h < 1 Then
                        Call HandleError(6)
                    Else
                        Call HandleError(7)
                    End If
                End If
            End If
        End If
    
        Exit Function
    
Err_BitBlit:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.BitBlit")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'Bei Entladen der Klasse Speicher-DC wieder freigeben und Bitmap löschen
Private Function TerminateDIB() As Boolean
    On Error GoTo Err_TerminateDIB
    
    If m_Initialized Then
        Call DeleteDC(m_iDC)
        Call DeleteObject(m_iBitmap)
        m_Initialized = False
        TerminateDIB = True
    End If
        
    Exit Function
    
Err_TerminateDIB:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.TerminateDIB")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'DIB zurücksetzen
Public Function Release() As Boolean
    On Error GoTo Err_Release
    
    Release = TerminateDIB
    
    Exit Function
    
Err_Release:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.Release")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'Byte-Array dimensionieren
Public Function SizeArray(bArray() As Byte, Optional Mode As DimConstants = dibNew) As Boolean
    On Error GoTo Err_SizeArray
    
    Dim uX As Long, uY As Long, w As Long
        
        If m_Width < 1 Then
            Call HandleError(12)
        ElseIf m_Height < 1 Then
            Call HandleError(13)
        Else
            Call GetArrayDimensions(uX, uY)
            
            m_LastUx = uX
            m_LastUy = uY
            
            If Mode = dibNew Then
                'Modus = neu, alte Daten werden gelöscht
                ReDim bArray(0 To uX, 0 To uY)
                SizeArray = True
            ElseIf Mode = dibPreserve Then
                'Modus = Bestand, alte Daten bleiben erhalten
                ReDim Preserve bArray(0 To uX, 0 To uY)
                SizeArray = True
            Else
                Call HandleError(5)
            End If
        End If
    
        Exit Function
    
Err_SizeArray:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.SizeArray")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'Array-Grenzen berechnen
Private Sub GetArrayDimensions(uX As Long, uY As Long)
    On Error GoTo Err_GetArrayDimensions
    
    Dim w As Long
    
        'Array muß durch vier teilbar sein
        If m_Width Mod 4 <> 0 Then
            w = Int(m_Width / 4) * 4 + 4
        Else
            w = m_Width
        End If
        
        uX = w * 3 - 1
        uY = m_Height - 1
    
        Exit Sub
    
Err_GetArrayDimensions:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.GetArrayDimensions")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'Klassen-Initialisierung
Private Sub Class_Initialize()
    On Error GoTo Err_Class_Initialize
    
    m_Initialized = False
    m_Language = dibGerman
    m_RunTime = dibSecure
    
    Exit Sub
    
Err_Class_Initialize:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.Class_Initialize")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'Klassen-Terminierung
Private Sub Class_Terminate()
    On Error GoTo Err_Class_Terminate
    
    Call TerminateDIB
    
    Exit Sub
    
Err_Class_Terminate:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.Class_Terminate")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'Fehlermeldungen verwalten
Private Sub HandleError(ErrNum As Long)
    On Error GoTo Err_HandleError
    
    Dim Text As String, lFlag As Boolean
        
        If m_Language = dibGerman Then
        
            Select Case ErrNum
                Case 1: Text = "Property 'hWndDestination' ist ungültig oder nicht festgelegt!"
                Case 2: Text = "Fehler beim Ermitteln des Destination-DCs!"
                Case 3: Text = "Fehler beim Freigeben des Destination-DCs!"
                Case 4: Text = "DIB ist bereits initialisiert!"
                Case 5: Text = "Unbekannter Dimensionierungs-Modus!"
                Case 6: Text = "Ungültige Höhenangabe!"
                Case 7: Text = "Ungültige Weitenangabe!"
                Case 8: Text = "x1 ausserhalb des gültigen Bereichs!"
                Case 9: Text = "y1 ausserhalb des gültigen Bereichs!"
                Case 10: Text = "Weite außerhalb des gültigen Bereichs!"
                Case 11: Text = "Höhe außerhalb des gültigen Bereichs!"
                Case 12: Text = "Property 'Width' ist ungültig oder nicht festgelegt!"
                Case 13: Text = "Property 'Height' ist ungültig oder nicht festgelegt!"
                Case 14: Text = "DIB ist noch nicht initialisiert!"
                Case 15: Text = "Property 'hWndSource' ist ungültig oder nicht festgelegt"
                Case 16: Text = "Fehler beim Ermitteln des Source-DCs"
                Case 17: Text = "Fehler beim Freigeben des Source-DCs"
            End Select
            
            lFlag = True
        ElseIf m_Language = dibEnglish Then
        
            Select Case ErrNum
                Case 1: Text = "Property 'hWndDestination' has an invalid or no value!"
                Case 2: Text = "Error while releasing destination-DC!"
                Case 3: Text = "Error while releasing destination-DC!"
                Case 4: Text = "DIB is initialized yet!"
                Case 5: Text = "Unkown mode for dimension!"
                Case 6: Text = "Illegal Quantity for 'Width'!"
                Case 7: Text = "Illegal Quantity for 'Height'!"
                Case 8: Text = "Illegal Quantity for 'x1'!"
                Case 9: Text = "Illegal Quantity for 'y1'!"
                Case 10: Text = "Illegal Quantity for 'Width'!"
                Case 11: Text = "Illegal Quantity for 'Height'!"
                Case 12: Text = "Property 'Width' has an invalid or no value!"
                Case 13: Text = "Property 'Height' has an invalid or no value!"
                Case 14: Text = "DIB is not initialized!"
                Case 15: Text = "Property 'hWndSource' has an invalid or no value!"
                Case 16: Text = "Error while getting source-DC"
                Case 17: Text = "Error while releasing source-DC"
            End Select
            
            lFlag = True
        End If
        
        If lFlag And (ErrNum > 0) Then
            RaiseEvent Error(ErrNum, Text)
        End If
    
        Exit Sub
    
Err_HandleError:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.HandleError")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'Größe des Arrays in x-Richtung
Public Property Get SizeX() As Long
    On Error GoTo Err_SizeX
    
    SizeX = m_LastUx
    
    Exit Property
    
Err_SizeX:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.SizeX")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'Größe des Arrays in y-Richtung
Public Property Get SizeY() As Long
    On Error GoTo Err_SizeY
    
    SizeY = m_LastUy
    
    Exit Property
    
Err_SizeY:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.SizeY")
        Case vbAbort: Exit Property
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Property

'x-Koordinaten auf den Gültigkeitsbreich stutzen
Public Function ScaleX(CoordinateX As Long) As Long
    On Error GoTo Err_ScaleX
    
    Dim x As Long
        
        x = CoordinateX
        If x < 0 Then
            x = 0
        ElseIf x > m_LastUx Then
            x = m_LastUx
        End If
    
        ScaleX = x
    
        Exit Function
    
Err_ScaleX:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.ScaleX")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'y-Koordinaten auf den Gültigkeitsbreich stutzen
Public Function ScaleY(CoordinateY As Long) As Long
    On Error GoTo Err_ScaleY
    
    Dim y As Long
    
        y = m_LastUx - CoordinateY
        If y < 0 Then
            y = 0
        ElseIf y > m_LastUy Then
            y = m_LastUy
        End If
    
        ScaleY = y
    
        Exit Function
    
Err_ScaleY:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.ScaleY")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'OLE-Farbe in Long-RGB-Wert wandeln
Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
    On Error GoTo Err_TranslateColor
    
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
    
    Exit Function
    
Err_TranslateColor:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.TranslateColor")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

'Long RGB-Wert in seine Rot, Grün und Blau-Anteile zerlegen
Public Sub RGBsplit(ByVal Col, R As Byte, G As Byte, B As Byte)
    On Error GoTo Err_RGBsplit
    
    B = (Col And 16711680) / 65536
    G = (Col And 65280) / 256
    R = Col And 255
    
    Exit Sub
    
Err_RGBsplit:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.RGBsplit")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'Rot, Grün und Blau-Anteil an x,y setzen, keine Sicherheitsabfrage!
Public Sub Poke(ByRef bArray() As Byte, ByVal x As Long, ByVal y As Long, R As Byte, G As Byte, B As Byte)
    On Error GoTo Err_Poke
    
    If x >= 0 And y >= 0 Then
        x = x * 3
        If x + 1 < UBound(bArray, 1) And y <= UBound(bArray, 2) Then
            bArray(x + 0, y) = B
            bArray(x + 1, y) = G
            bArray(x + 2, y) = R
        End If
    End If
    
    Exit Sub
    
Err_Poke:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.Poke")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub

'RGB-Long-Wert an x,y setzen, mit Sicherheitsabfrage
Public Sub PokeEx(ByRef bArray() As Byte, ByVal x As Long, ByVal y As Long, Color As Long)
    On Error GoTo Err_PokeEx
    
    Dim R As Byte, G As Byte, B As Byte
    
        If x >= 0 And y >= 0 Then
            x = x * 3
            If x + 1 < UBound(bArray, 1) And y <= UBound(bArray, 2) Then
                Call RGBsplit(Color, R, G, B)
                bArray(x + 0, y) = B
                bArray(x + 1, y) = G
                bArray(x + 2, y) = R
            End If
        End If
    
        Exit Sub
    
Err_PokeEx:

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "tdSplitter.clsDIB.PokeEx")
        Case vbAbort: Exit Sub
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Sub
