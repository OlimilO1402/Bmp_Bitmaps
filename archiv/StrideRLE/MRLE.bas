Attribute VB_Name = "MRLE"
Option Explicit
'https://learn.microsoft.com/en-us/windows/win32/gdi/bitmap-compression
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pDst As Any, ByVal bytLength As Long)
Private Declare Sub RtlFillMemory Lib "kernel32" (pDst As Any, ByVal bytLength As Long, ByVal fill As Byte)

Public Function RLE8_Decode(ByVal W As Long, ByVal H As Long, buffer_in() As Byte, buffer_out() As Byte) As Boolean
Try: On Error GoTo Catch
    Dim Stride As Long: Stride = CalcStride(W, 8)
    'Dim Stride As Long: Stride = GetStride(W, 8)
    Dim Size_of_buffer_out As Long: Size_of_buffer_out = Stride * H
    Dim i As Long, ui As Long: ui = UBound(buffer_in)
    Dim o As Long, uo As Long: uo = Size_of_buffer_out - 1: ReDim buffer_out(0 To uo)
    Dim d As Byte: d = Stride - W
    Dim Byte1 As Byte, Byte2 As Byte, size As Byte, fill As Byte
    Dim dRi As Byte, dUp As Byte
    Do While i <= ui
        'read 2 bytes
        Byte1 = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
        Byte2 = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
        Select Case Byte2
        Case 0: 'End of line
                'Striding-Padbytes in Ausgabe einfügen, bzw einfach den Ausgabe-Index weitersetzen
                o = o + d
        Case 1: 'End of Bitmap
                Exit Do
        Case 2: 'delta mode
            dRi = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
            dUp = buffer_in(i): i = i + 1: If ui <= i Then Exit Do
            'jump to the position
            'todo
        Case Else
            If Byte1 = 0 Then
                'Absolute mode
                size = Byte2
                RtlMoveMemory buffer_out(o), buffer_in(i), size
                i = i + size: If ui <= i Then Exit Do
                o = o + size: If uo <= o Then Exit Do
                If IsOdd(size) Then i = i + 1
                'do the padding
                '
            Else
                'Encoded mode
                'now fill the outputbuffer with the same pixels
                size = Byte1
                fill = Byte2
                RtlFillMemory buffer_out(o), size, fill
                o = o + size
            End If
        End Select
    Loop
    RLE8_Decode = True
    Exit Function
Catch:
    MsgBox Err.Description
End Function
Public Function IsOdd(ByVal n As Byte) As Boolean
    'gibt zurück ob die Zahl n ungerade ist
    IsOdd = n Mod 2
End Function
Public Function IsEven(ByVal n As Byte) As Boolean
    'gibt zurück ob die Zahl n gerade ist
    IsEven = (n Mod 2) = 0
End Function

Public Property Get CalcStride(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
    'calculates the number of bytes in one horizontal line of pixels,
    'including the number of pad-bytes for the 4-aligned result
    CalcStride = W * BitsPerPixel / 8
    Dim m As Long: m = CalcStride Mod 4: If m > 0 Then m = 4 - m
    CalcStride = CalcStride + m
End Property

Private Function GetStride(ByVal W As Long, ByVal BitsPerPixel As Byte) As Long
    Dim bypp As Long: bypp = BitsPerPixel / 8
    Select Case m_bpp
    Case 8:        GetStride = (W + 3) And Not 3
    Case 16, 24:   GetStride = ((W * bypp) + bypp) And Not bypp
    Case 32:       GetStride = W * bypp
    End Select
End Function

