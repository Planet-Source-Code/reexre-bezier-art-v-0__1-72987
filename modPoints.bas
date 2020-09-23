Attribute VB_Name = "modPoints"
Public Type mPoint
    X             As Double
    Y         As Double
    vX        As Double
    vY        As Double
End Type



Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long


Public Function QuadDistance(dx, dy) As Single
    QuadDistance = (dx * dx + dy * dy)
End Function


Public Function Force(Qdist) As Single

    If Qdist <> 0 Then
        Force = 20 / (Qdist)    '25
        'Force = 100 / (Qdist)
    Else
        Force = 1
    End If


End Function





Public Sub Long2RGB(RGBcol As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer)

    If RGBcol > 0 Then
        R = RGBcol And &HFF    ' set red
        G = (RGBcol And &H100FF00) / &H100    ' set green
        B = (RGBcol And &HFF0000) / &H10000    ' set blue
    Else
        R = 0: G = 0: B = 0
    End If

End Sub
