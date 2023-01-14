Attribute VB_Name = "MNew"
Option Explicit
#If False Then
    Bitmap 'otherwise it turns to start with lowercase letter "bitmap"
#End If

Public Function Bitmap(aPFN As String) As Bitmap
    Set Bitmap = New Bitmap: Bitmap.New_ aPFN
End Function

Public Function BitmapWH(ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As EPixelFormat) As Bitmap
    Set BitmapWH = New Bitmap: BitmapWH.NewWH Width, Height, PixelFormat
End Function

