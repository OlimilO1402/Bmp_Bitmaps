Attribute VB_Name = "MNew"
Option Explicit
#If False Then
    Bitmap
#End If

Public Function Bitmap(ByVal aPFN As String) As Bitmap
    Set Bitmap = New Bitmap: Bitmap.New_ aPFN
End Function

Public Function BitmapWH(ByVal Width As Long, ByVal Height As Long, ByVal PixelFormat As EPixelFormat) As Bitmap
    Set BitmapWH = New Bitmap: BitmapWH.NewWH Width, Height, PixelFormat
End Function

Public Function BitmapSP(aStdPicture As StdPicture, ByVal aPFN As String) As Bitmap
    Set BitmapSP = New Bitmap: BitmapSP.NewSP aStdPicture, aPFN
End Function

Public Function PictureBoxZoom(Window As Form, Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Window, Canvas, aImage
End Function

Public Function ColorSelector(aTimer As Timer, aButton As CommandButton, aColorView As PictureBox, aLabel As Label) As ColorSelector
    Set ColorSelector = New ColorSelector: ColorSelector.New_ aTimer, aButton, aColorView, aLabel
End Function

Public Function ScannerTwain(Owner As Form) As ScannerTwain
    Set ScannerTwain = New ScannerTwain: ScannerTwain.New_ Owner
End Function

