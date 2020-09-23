VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type Bitmap
    bmType        As Long
    bmWidth   As Long
    bmHeight  As Long
    bmWidthBytes As Long
    bmPlanes  As Integer
    bmBitsPixel As Integer
    bmBits    As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private TargetBits() As Byte
Private SourceBits() As Byte
Private SINGLEb() As Single


Private hBmp  As Bitmap

Private iRET  As Long


Public Sub InitTarget(pBoxPicturehandle As Long)

    iRET = GetObject(pBoxPicturehandle, Len(hBmp), hBmp)

    ReDim TargetBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte

End Sub
Public Sub GetBits(pBoxPicturehandle As Long)
Dim iRET      As Long
    'Get the bitmap header
    iRET = GetObject(pBoxPicturehandle, Len(hBmp), hBmp)
    'Resize to hold image data


    ReDim SourceBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into SOURCEbits array
    iRET = GetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, SourceBits(0, 0, 0))

    '    ReDim TargetBits(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
    'Get the image data and store into TargetBits array
    '  iRet = GetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, TargetBits(0, 0, 0))


    ReDim SINGLEb(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Single

End Sub

Public Sub SetBits(pBoxPicturehandle As Long)
Dim iRET      As Long
    'Set the new image data back onto pBox
    iRET = SetBitmapBits(pBoxPicturehandle, hBmp.bmWidthBytes * hBmp.bmHeight, TargetBits(0, 0, 0))
    'Erase TargetBits because we finished with it now
    'Erase TargetBits
End Sub



Public Function SetBezierBits(X, Y, ByVal R As Single, ByVal G As Single, ByVal B As Single)
Dim R2
Dim G2
Dim B2
Dim X2        As Integer
Dim Y2        As Integer

    R = R * 0.05
    G = G * 0.05
    B = B * 0.05


    R2 = R * 0.7
    G2 = G * 0.7
    B2 = B * 0.7


    SINGLEb(2, X, Y) = SINGLEb(2, X, Y) + R
    SINGLEb(1, X, Y) = SINGLEb(1, X, Y) + G
    SINGLEb(0, X, Y) = SINGLEb(0, X, Y) + B



    X2 = X - 1
    Y2 = Y
    SINGLEb(2, X2, Y2) = SINGLEb(2, X2, Y2) + R2
    SINGLEb(1, X2, Y2) = SINGLEb(1, X2, Y2) + G2
    SINGLEb(0, X2, Y2) = SINGLEb(0, X2, Y2) + B2
    X2 = X + 1
    Y2 = Y
    SINGLEb(2, X2, Y2) = SINGLEb(2, X2, Y2) + R2
    SINGLEb(1, X2, Y2) = SINGLEb(1, X2, Y2) + G2
    SINGLEb(0, X2, Y2) = SINGLEb(0, X2, Y2) + B2
    X2 = X
    Y2 = Y - 1
    SINGLEb(2, X2, Y2) = SINGLEb(2, X2, Y2) + R2
    SINGLEb(1, X2, Y2) = SINGLEb(1, X2, Y2) + G2
    SINGLEb(0, X2, Y2) = SINGLEb(0, X2, Y2) + B2
    X2 = X
    Y2 = Y + 1
    SINGLEb(2, X2, Y2) = SINGLEb(2, X2, Y2) + R2
    SINGLEb(1, X2, Y2) = SINGLEb(1, X2, Y2) + G2
    SINGLEb(0, X2, Y2) = SINGLEb(0, X2, Y2) + B2

skip:

End Function


Public Sub PaintBezierImage()
Dim R
Dim G
Dim B
Dim bR        As Byte
Dim bG        As Byte
Dim bB        As Byte
Dim X         As Long
Dim Y         As Long


    For X = 0 To hBmp.bmWidth - 1
        For Y = 0 To hBmp.bmHeight - 1
            SINGLEb(2, X, Y) = SINGLEb(2, X, Y) * 0.999    '0.9987
            SINGLEb(1, X, Y) = SINGLEb(1, X, Y) * 0.999    '0.9987
            SINGLEb(0, X, Y) = SINGLEb(0, X, Y) * 0.999    '0.9987

            'If SINGLEb(2, X, Y) > 255 Then SINGLEb(2, X, Y) = 255
            'If SINGLEb(1, X, Y) > 255 Then SINGLEb(1, X, Y) = 255
            'If SINGLEb(0, X, Y) > 255 Then SINGLEb(0, X, Y) = 255

            If SINGLEb(2, X, Y) < 255 Then
                bR = CByte(SINGLEb(2, X, Y))
            Else
                bR = 255
            End If

            If SINGLEb(1, X, Y) < 255 Then
                bG = CByte(SINGLEb(1, X, Y))
            Else
                bG = 255
            End If

            If SINGLEb(0, X, Y) < 255 Then
                bB = CByte(SINGLEb(0, X, Y))
            Else
                bB = 255
            End If


            TargetBits(2, X, Y) = Not (bR)
            TargetBits(1, X, Y) = Not (bG)
            TargetBits(0, X, Y) = Not (bB)
        Next
    Next



End Sub

Public Sub ClearBezierBits()
'ReDim SINGLEb(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Single


End Sub

