Attribute VB_Name = "Módulo2"
Function QrCode(codetext As String) As String

Dim URL As String, MyCell As Range

Set MyCell = Application.Caller
URL = "https://api.qrserver.com/v1/create-qr-code/?size=100x100&data=" & codetext

 
On Error Resume Next
  ActiveSheet.Pictures("QR_" & MyCell.Address(False, False)).Delete
On Error GoTo 0
ActiveSheet.Pictures.Insert(URL).Select
With Selection.ShapeRange(1)
 '.PictureFormat.CropLeft = 10
 '.PictureFormat.CropRight = 10
 '.PictureFormat.CropTop = 10
 '.PictureFormat.CropBottom = 10
 .Name = "QR_" & MyCell.Address(False, False)
 
 .Left = MyCell.Left + 2
 .Top = MyCell.Top + 2
End With
QrCode = ""
Selection.ShapeRange(1).Name = "NOMQR"
End Function






