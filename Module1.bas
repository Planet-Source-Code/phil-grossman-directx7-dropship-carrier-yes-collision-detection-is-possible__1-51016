Attribute VB_Name = "Module1"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Global AMB
Public Function JPEG2BMP(Filename As String, LoadPB As PictureBox, SavePB As PictureBox) As Boolean
On Error GoTo FileMuffUp

LoadPB = LoadPicture(Filename & ".jpg")
SavePB = LoadPB
SavePicture SavePB.Picture, Filename & ".bmp"
JPEG2BMP = True
Exit Function

FileMuffUp:
JPEG2BMP = False
End Function
