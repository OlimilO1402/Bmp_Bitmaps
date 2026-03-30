VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DIBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Error(ErrorNumber As Long, ErrorText As String)

Private Declare Function OleTranslateColor Lib "olepro32" (ByVal OLE_COLOR As Long, ByVal HPALETTE As LongPtr, pccolorref As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As LongPtr, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As LongPtr, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal ahDC As LongPtr, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As LongPtr, ByVal DestX As Long, ByVal DestY As Long, ByVal DestWidth As Long, ByVal DestHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal wMsg  As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long

Private Const CLR_INVALID  As Long = -1
Private Const WM_SETREDRAW As Long = &HB
                     
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

Private Type RGBQUAD
    rgbBlue         As Byte
    rgbGreen        As Byte
    rgbRed          As Byte
    rgbReserved     As Byte
End Type

Private Type BITMAPINFO
    bmiHeader       As BITMAPINFOHEADER
    bmiColors       As RGBQUAD
End Type

Private Const BI_RGB              As Long = 0&
Private Const DIB_RGB_COLORS      As Long = 0&
Private Const MCI_NOTIFY&         As Long = &H3B9
Private Const MCI_NOTIFY_SUCCESS& As Long = &H1
Private Const GWL_WNDPROC         As Long = -4
Private Const SWP_SHOWWINDOW      As Long = &H40
Private Const SWP_HIDEWINDOW      As Long = &H80

Private m_hBitmap     As LongPtr
Private m_hDC         As LongPtr
Private m_bi24BitInfo As BITMAPINFO

Private m_Width       As Long 'why not use: ' m_bi24BitInfo.bmiHeader.biWidth  'because it must bei div.b.4
Private m_Height      As Long 'why not use: ' m_bi24BitInfo.bmiHeader.biHeight 'because it must bei div.b.4
Private m_RunTime     As RunTimeConstants

Private m_hWnd_Destination As LongPtr
Private m_hWnd_Source As LongPtr
Private m_Language    As LanguageConstants
Private m_Initialized As Boolean
Private m_RedrawState As Boolean

Private m_LastUx As Long
Private m_LastUy As Long

'Klassen-Initialisierung
Private Sub Class_Initialize()
    m_Initialized = False
    m_Language = dibGerman
    m_RunTime = dibSecure
End Sub

'Klassen-Terminierung
Private Sub Class_Terminate()
Try: On Error GoTo Catch    
    TerminateDIB    
    Exit Sub    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.Class_Terminate")
	Case vbAbort:  Exit Sub
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Sub

'Bei Entladen der Klasse Speicher-DC wieder freigeben und Bitmap löschen
Private Function TerminateDIB() As Boolean
Try: On Error GoTo Catch    
    If m_Initialized Then
        DeleteDC m_hDC
        DeleteObject m_hBitmap
        m_Initialized = False
        TerminateDIB = True
    End If        
    Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.TerminateDIB")
	Case vbAbort:  Exit Function
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'DIB zurücksetzen
Public Function Release() As Boolean
Try: On Error GoTo Catch    
    Release = TerminateDIB    
    Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.Release")
	Case vbAbort:  Exit Function
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'Weite der DIB lesen oder festlegen
Public Property Get Width() As Long
    Width = m_Width
End Property
Public Property Let Width(ByVal Value  As Long)
    m_Width = Value
End Property

'Höhe der DIB lesen oder festlegen
Public Property Get Height() As Long
    Height = m_Height
End Property
Public Property Let Height(ByVal Value As Long)
    m_Height = Value
End Property

'hWnd des Ziel-DCs auslesen oder festlegen
Public Property Get hWndDestination() As Long
    hWndDestination = m_hWnd_Destination
End Property
Public Property Let hWndDestination(ByVal Value As Long)
    m_hWnd_Destination = Value
End Property

'hWnd des Quell-DCs auslesen oder festlegen
Public Property Get hWndSource() As Long
    hWndSource = m_hWnd_Source
End Property
Public Property Let hWndSource(ByVal Value As Long)
    m_hWnd_Source = Value
End Property

'RuntimeMode auslesen oder festlegen (tdSpeed und tdSecure)
Public Property Get RuntimeMode() As RunTimeConstants
    RuntimeMode = m_RunTime
End Property
Public Property Let RuntimeMode(ByVal Value As RunTimeConstants)
    m_RunTime = Value
End Property

'Sprachauswahl auslesen oder festlegen
Public Property Get Language() As LanguageConstants
    Language = m_Language
End Property
Public Property Let Language(ByVal Value As LanguageConstants)
    m_RunTime = Value
End Property


'Lock-Windows-Update auslesen bzw ent- oder verriegeln
Public Property Get Redraw() As Boolean
    Redraw = m_RedrawState
End Property
Public Property Let Redraw(ByVal Value As Boolean)
Try: On Error GoTo Catch    
    If m_hWnd_Destination <> 0 Then
        SendMessage m_hWnd_Destination, WM_SETREDRAW, Abs(Value), 0&
        m_RedrawState = State
    Else
        HandleError 1
    End If    
    Exit Property
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.SetRedraw")
    Case vbAbort:  Exit Property
    Case vbRetry:  Resume
    Case vbIgnore: Resume Next
    End Select
End Property

'Abfragemöglichkeit für in der WindowsMain stehende Nachrichten mit Filter
Public Function MessagesQueued(Messages As QueueStatusConstants, Mode As QueueFilterConstants) As Long
Try: On Error GoTo Catch    
    Dim x  As Long:  x = GetQueueStatus(Messages)
	Dim	Hi As Long: Hi = (x And &HFF00) / &HF00
	Dim Lo As Long: Lo = x And &HFF
    If Mode = qsAll Then
        MessagesQueued = Hi
    Else
        MessagesQueued = Lo
    End If    
    Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.MessagesQueued")
    Case vbAbort:  Exit Function
    Case vbRetry:  Resume
    Case vbIgnore: Resume Next
    End Select
End Function


'Schaun ob die DIB bereits initialisiert ist
Public Property Get IsCreated() As Boolean
    IsCreated = m_Initialized
End Property

'DIB initialisieren
Public Function Create() As Boolean
Try: On Error GoTo Catch
    
    Dim bWidth  As Long
	Dim	bHeight As Long
	Dim w As Long
    Dim n_hDC As LongPtr
	Dim	Result As Long
  
	If m_Initialized Then
		'DIB ist noch nicht intialisiert
		HandleError 4
	Else
		If m_hWnd_Destination <> 0 Then
			'DC des Zielbildes festnageln
			n_hDC = GetDC(m_hWnd_Destination)
			If n_hDC <> 0 Then
				
				'Weite der DIB muß intern durch vier teilbar sein
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
				m_hDC = CreateCompatibleDC(n_hDC)
				
				'DIB erstellen
				m_hBitmap = CreateDIBSection(m_hDC, m_bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
		
				'DIB in den DC wählen
				SelectObject m_hDC, m_hBitmap
		
				'DC des Zielbildes wieder freigeben
				Result = ReleaseDC(m_hWnd_Destination, n_hDC)
				If Result <> 0 Then
					m_Initialized = True
					Create = True
				Else
					HandleError 3 
				End If

			Else
				HandleError 2
			End If
		Else
			HandleError 1
		End If
	End If
	Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.Create")
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
                                      
    Dim w As Long, h As Long, n_hDC As LongPtr, Result As Long, x As Long, y As Long
    
        If m_RunTime = dibSpeed Then
            'Quell-DC festnageln
            n_hDC = GetDC(m_hWnd_Source)
            
            'Quellbild in den Speicher-DC kopieren, ROPs anwenden
            Call BitBlt(m_hDC, x1_Dst, y1_Dst, w_Dst, h_Dst, n_hDC, x1_Src, y1_Src, RasterOperation)
                                      
            'Bytes aus dem Speicher-DC in das Array holen
            Call GetDIBits(n_hDC, m_hBitmap, 0, m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), _
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
                                BitBlt m_hDC, x1_Dst, y1_Dst, w, h, n_hDC, x1_Src, y1_Src, RasterOperation
                                
                                'Bytes aus dem Speicher-DC in das Array holen
                                GetDIBits n_hDC, m_hBitmap, 0, m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS
                                               
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

    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.CopyBitmapToArray")
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
                
Try: On Error GoTo Catch
                    
    Dim h As Long, w As Long, n_hDC As LongPtr, Result As Long
        
	If m_RunTime = dibSpeed Then
		'Ziel-DC festnageln
		n_hDC = GetDC(m_hWnd_Destination)
		
		'ByteArray auf DC ausgeben
		StretchDIBits(n_hDC, x1_Dst, y1_Dst, w_Dst, h_Dst, x1_Src, y1_Src, w_Src, h_Src, bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS, RasterOperation)
		
		'Ziel-DC wieder freigeben
		ReleaseDC m_hWnd_Destination, n_hDC
	Else
	
		If Not m_Initialized Then
			HandleError 14
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
					HandleError 8
				ElseIf y1_Src > m_Height - 1 Then
					HandleError 9
				ElseIf x1_Src + w > m_Width Then
					HandleError 10
				ElseIf y1_Src + h > m_Height Then
					HandleError 11
				Else
					If m_hWnd_Destination <> 0 Then
						'Ziel-DC festnageln
						n_hDC = GetDC(m_hWnd_Destination)
						If n_hDC <> 0 Then
						
							If w_Dst = 0 Then w_Dst = w
							If h_Dst = 0 Then h_Dst = h
															
							'ByteArray auf DC ausgeben
							StretchDIBits(n_hDC, x1_Dst, y1_Dst, w_Dst, h_Dst, x1_Src, y1_Src, w, h, bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS, RasterOperation)
						
							'Ziel-DC wieder freigeben
							Result = ReleaseDC(m_hWnd_Destination, n_hDC)
							
							If Result = 0 Then
								HandleError 3
							Else
								StretchBitBlit = True
							End If
						Else
							HandleError 2
						End If
					Else
						HandleError 1
					End If
				End If
			Else
				If h < 1 Then
					HandleError 6
				Else
					HandleError 7
				End If
			End If
		End If
	End If

	Exit Function
    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.StretchBitBlit")
        Case vbAbort: Exit Function
        Case vbRetry: Resume
        Case vbIgnore: Resume Next
    End Select
End Function

Public Function BitBlit(bArray() As Byte, _
                        Optional x1_Src As Long, Optional y1_Src As Long, _
                        Optional w_Src As Long, Optional h_Src As Long, _
                        Optional x1_Dst As Long, Optional y1_Dst As Long) As Boolean
                

Try: On Error GoTo Catch
                                        
    Dim h As Long, w As Long
	Dim	n_hDC As LongPtr, Result As Long
        
	If m_RunTime = dibSpeed Then
		'Ziel-DC festnageln
		n_hDC = GetDC(m_hWnd_Destination)
		
		'ByteArray auf DC ausgeben
		SetDIBitsToDevice n_hDC, x1_Dst, y1_Dst, w_Src, h_Src, x1_Src, y1_Src, 0, m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS
							   
		'Ziel-DC freigeben
		ReleaseDC m_hWnd_Destination, n_hDC
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
					HandleError 8
				ElseIf y1_Src > m_Height - 1 Then
					HandleError 9
				ElseIf x1_Src + w > m_Width Then
					HandleError 10
				ElseIf y1_Src + h > m_Height Then
					HandleError 11
				Else
					If m_hWnd_Destination <> 0 Then
						'Ziel-DC festnageln
						n_hDC = GetDC(m_hWnd_Destination)
						If n_hDC <> 0 Then
						
							'ByteArray auf DC ausgeben
							SetDIBitsToDevice n_hDC, x1_Dst, y1_Dst, w, h, x1_Src, y1_Src, 0, m_bi24BitInfo.bmiHeader.biHeight, bArray(0, 0), m_bi24BitInfo, DIB_RGB_COLORS
						
							'Ziel-DC freigeben
							Result = ReleaseDC(m_hWnd_Destination, n_hDC)
							
							If Result = 0 Then
								HandleError 2
							Else
								BitBlit = True
							End If
						Else
							HandleError 2
						End If
					Else
						HandleError 1
					End If
				End If
			Else
				If h < 1 Then
					HandleError 6
				Else
					HandleError 7
				End If
			End If
		End If
	End If

	Exit Function
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.BitBlit")
    Case vbAbort:  Exit Function
    Case vbRetry:  Resume
    Case vbIgnore: Resume Next
    End Select
End Function

'Byte-Array dimensionieren
Public Function SizeArray(bArray() As Byte, Optional Mode As DimConstants = dibNew) As Boolean
Try: On Error GoTo Catch
    
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
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.SizeArray")
	Case vbAbort: Exit Function
	Case vbRetry: Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'Array-Grenzen berechnen
Private Sub GetArrayDimensions(ByRef uX_out As Long, uY_out As Long)
Try: On Error GoTo Catch
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
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.GetArrayDimensions")
    Case vbAbort:  Exit Sub
    Case vbRetry:  Resume
    Case vbIgnore: Resume Next
    End Select
End Sub

'Fehlermeldungen verwalten
Private Sub HandleError(ByVal e As Long)
Try: On Error GoTo Catch    
    Dim s As String, lFlag As Boolean        
	If m_Language = dibGerman Then	
		Select Case e
		Case 1:  s = "Property 'hWndDestination' ist ungültig oder nicht festgelegt!"
		Case 2:  s = "Fehler beim Ermitteln des Destination-DCs!"
		Case 3:  s = "Fehler beim Freigeben des Destination-DCs!"
		Case 4:  s = "DIB ist bereits initialisiert!"
		Case 5:  s = "Unbekannter Dimensionierungs-Modus!"
		Case 6:  s = "Ungültige Höhenangabe!"
		Case 7:  s = "Ungültige Weitenangabe!"
		Case 8:  s = "x1 ausserhalb des gültigen Bereichs!"
		Case 9:  s = "y1 ausserhalb des gültigen Bereichs!"
		Case 10: s = "Weite außerhalb des gültigen Bereichs!"
		Case 11: s = "Höhe außerhalb des gültigen Bereichs!"
		Case 12: s = "Property 'Width' ist ungültig oder nicht festgelegt!"
		Case 13: s = "Property 'Height' ist ungültig oder nicht festgelegt!"
		Case 14: s = "DIB ist noch nicht initialisiert!"
		Case 15: s = "Property 'hWndSource' ist ungültig oder nicht festgelegt"
		Case 16: s = "Fehler beim Ermitteln des Source-DCs"
		Case 17: s = "Fehler beim Freigeben des Source-DCs"
		End Select
		lFlag = True
	ElseIf m_Language = dibEnglish Then
		Select Case e
		Case 1:  s = "Property 'hWndDestination' has an invalid or no value!"
		Case 2:  s = "Error while releasing destination-DC!"
		Case 3:  s = "Error while releasing destination-DC!"
		Case 4:  s = "DIB is initialized yet!"
		Case 5:  s = "Unkown mode for dimension!"
		Case 6:  s = "Illegal Quantity for 'Width'!"
		Case 7:  s = "Illegal Quantity for 'Height'!"
		Case 8:  s = "Illegal Quantity for 'x1'!"
		Case 9:  s = "Illegal Quantity for 'y1'!"
		Case 10: s = "Illegal Quantity for 'Width'!"
		Case 11: s = "Illegal Quantity for 'Height'!"
		Case 12: s = "Property 'Width' has an invalid or no value!"
		Case 13: s = "Property 'Height' has an invalid or no value!"
		Case 14: s = "DIB is not initialized!"
		Case 15: s = "Property 'hWndSource' has an invalid or no value!"
		Case 16: s = "Error while getting source-DC"
		Case 17: s = "Error while releasing source-DC"
		End Select		
		lFlag = True
	End If	
	If lFlag And (e > 0) Then
		RaiseEvent Error(e, s)
	End If
	Exit Sub    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.HandleError")
	Case vbAbort:  Exit Sub
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Sub

'Größe des Arrays in x-Richtung
Public Property Get SizeX() As Long
    SizeX = m_LastUx
End Property

'Größe des Arrays in y-Richtung
Public Property Get SizeY() As Long
    SizeY = m_LastUy
End Property

'x-Koordinaten auf den Gültigkeitsbreich stutzen
Public Function ScaleX(ByVal CoordinateX As Long) As Long
Try: On Error GoTo Catch    
    Dim x As Long: x = CoordinateX
	If x < 0 Then
		x = 0
	ElseIf x > m_LastUx Then
		x = m_LastUx
	End If
	ScaleX = x
	Exit Function
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.ScaleX")
	Case vbAbort:  Exit Function
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'y-Koordinaten auf den Gültigkeitsbreich stutzen
Public Function ScaleY(CoordinateY As Long) As Long
Try: On Error GoTo Catch    
    Dim y As Long: y = m_LastUx - CoordinateY
	If y < 0 Then
		y = 0
	ElseIf y > m_LastUy Then
		y = m_LastUy
	End If
	ScaleY = y
	Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.ScaleY")
	Case vbAbort:  Exit Function
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'OLE-Farbe in Long-RGB-Wert wandeln
Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
Try: On Error GoTo Catch    
    If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID    
    Exit Function    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.TranslateColor")
	Case vbAbort:  Exit Function
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Function

'Long RGB-Wert in seine Rot, Grün und Blau-Anteile zerlegen
Public Sub RGBsplit(ByVal Col, R As Byte, G As Byte, B As Byte)
Try: On Error GoTo Catch    
    B = (Col And 16711680) / 65536
    G = (Col And 65280) / 256
    R = Col And 255    
    Exit Sub
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.RGBsplit")
	Case vbAbort:  Exit Sub
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Sub

'Rot, Grün und Blau-Anteil an x,y setzen, keine Sicherheitsabfrage!
Public Sub Poke(ByRef bArray() As Byte, ByVal x As Long, ByVal y As Long, R As Byte, G As Byte, B As Byte)
Try: On Error GoTo Catch    
    If x >= 0 And y >= 0 Then
        x = x * 3
        If x + 1 < UBound(bArray, 1) And y <= UBound(bArray, 2) Then
            bArray(x + 0, y) = B
            bArray(x + 1, y) = G
            bArray(x + 2, y) = R
        End If
    End If    
    Exit Sub    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.Poke")
	Case vbAbort:  Exit Sub
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Sub

'RGB-Long-Wert an x,y setzen, mit Sicherheitsabfrage
Public Sub PokeEx(ByRef bArray() As Byte, ByVal x As Long, ByVal y As Long, Color As Long)
Try: On Error GoTo Catch
	If x >= 0 And y >= 0 Then
		x = x * 3
		If x + 1 < UBound(bArray, 1) And y <= UBound(bArray, 2) Then
			Dim R As Byte, G As Byte, B As Byte    
			RGBsplit Color, R, G, B
			bArray(x + 0, y) = B
			bArray(x + 1, y) = G
			bArray(x + 2, y) = R
		End If
	End If
	Exit Sub    
Catch:
    Select Case MsgBox(Err.Description, vbExclamation + vbAbortRetryIgnore, "DIBitmap.PokeEx")
	Case vbAbort:  Exit Sub
	Case vbRetry:  Resume
	Case vbIgnore: Resume Next
    End Select
End Sub
