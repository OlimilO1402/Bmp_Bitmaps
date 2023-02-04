Attribute VB_Name = "MMain"
Option Explicit
Private m_Filter As String

Sub Main()
    m_Filter = "Windows & OS/2 Bitmap [*.bmp]|*.bmp|Portable Network Graphic [*.png]|*.png|Jpeg [*jpg]|*.jpg|Graphics Interchange Format [*.gif]|*.gif|All files (*.*)|*.*"""
    FMain.Show
End Sub
'
Public Function GetOpenFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New OpenFileDialog
    GetOpenFileName = GetFileName(Owner, FD, PFN)
'    With FD
'        .Filter = m_Filter
'        '.InitialDirectory
'        If .ShowDialog(FMain) = vbCancel Then Exit Function
'        GetOpenFileName = .FileName
'    End With
End Function
'
Public Function GetSaveFileName(Owner As Form, Optional ByVal PFN As String) As String
    Dim FD As New SaveFileDialog
    GetSaveFileName = GetFileName(Owner, FD, PFN)
'    GetSaveFileName
'    With FD
'        .Filter = m_Filter
'        '.InitialDirectory
'        If .ShowDialog(FMain) = vbCancel Then Exit Function
'        GetSaveFileName = .FileName
'    End With
End Function

Public Function GetFileName(Owner As Form, FD, Optional ByVal PFN As String) As String
    If FD Is Nothing Then Exit Function
    With FD
        .Filter = m_Filter
        .FileName = PFN
        '.InitialDirectory
        If .ShowDialog(Owner) = vbCancel Then Exit Function
        GetFileName = .FileName
    End With
End Function

