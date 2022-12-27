Attribute VB_Name = "MNew"
Option Explicit

Public Function Bitmap(aPFN As String) As Bitmap
    Set Bitmap = New Bitmap: Bitmap.New_ aPFN
End Function

