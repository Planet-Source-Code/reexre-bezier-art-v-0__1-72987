VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Code by Cyborg (contact at vbforums.com)
' 30/3 - 2009

Option Explicit

Private Type Point2D
    X             As Double
    Y         As Double
End Type

Private Points() As Point2D

Private FX    As New clsFX

Public Sub InitTarget(PicImageHandle As Long)
    FX.InitTarget PicImageHandle
    FX.GetBits PicImageHandle

End Sub

Public Sub InitCurve(numVerts As Long)
ReDim Points(numVerts - 1)  'set size of array
End Sub

Public Sub RemoveCurve()
Erase Points    'free mem
End Sub

'find a point in close proximity to a coordinate
Public Function FindPoint(ByVal X As Double, ByVal Y As Double, Precision As Double) As Long
Dim i         As Long

    For i = 0 To UBound(Points)
        With Points(i)
            If X > .X - Precision And X < .X + Precision And Y > .Y - Precision And Y < .Y + Precision Then
                FindPoint = i
                Exit Function
            End If
        End With
    Next
    FindPoint = -1
End Function

'return values
Public Function GetPointX(i As Long) As Double
    GetPointX = Points(i).X
End Function

Public Function GetPointY(i As Long) As Double
    GetPointY = Points(i).Y
End Function

Public Function GetNumVerts() As Long
    GetNumVerts = UBound(Points) + 1
End Function

Public Sub SetPointCoords(i As Long, ByVal X As Double, ByVal Y As Double)
    Points(i).X = X
    Points(i).Y = Y
End Sub

'returns x and y coords for the curve
Public Sub GetBezierPoint(t As Double, ByRef X As Double, ByRef Y As Double)
Dim Point     As Point2D
    Point = TraceCurve(Points, UBound(Points) + 1, t)
    X = Point.X
    Y = Point.Y
End Sub

'recursive function to trace out the curve.
'for each call it connects the lines from the previous call at distance t from each point.
'each call will produce one less point(or line), and it will keep calling itself until there's only one point left.
'the last point will finally be returned
Private Function TraceCurve(P() As Point2D, ByVal numPoints As Long, t As Double) As Point2D
Dim i         As Long
Dim newP()    As Point2D    'holds the new points
    ReDim newP(numPoints - 2)

    For i = 0 To numPoints - 2
        'get a new point at fraction t between 2 other points
        newP(i).X = t * (P(i + 1).X - P(i).X) + P(i).X
        newP(i).Y = t * (P(i + 1).Y - P(i).Y) + P(i).Y
    Next

    If numPoints > 2 Then
        'send the new points to TraceCurve
        TraceCurve = TraceCurve(newP, numPoints - 1, t)
    ElseIf numPoints = 2 Then
        'if the function is at it's last stop i.e. there is only one new point.
        'return that point instead of calling TraceCurve again
        TraceCurve = newP(0)
    End If

    Erase newP    'free mem
End Function





Public Sub RenderOLD(ByRef dPIC As PictureBox, Iterations As Long, Color As Long)

Dim R1        As Integer
Dim g1        As Integer
Dim b1        As Integer
Dim R2        As Integer
Dim G2        As Integer
Dim B2        As Integer
Dim C         As Long


    Long2RGB Color, R1, b1, g1

    Dim t     As Double
    Dim X     As Double
    Dim Y     As Double
    Dim i     As Long

    'dPIC.Cls
    'draw lines
    'If chkLines.Value Then
    '    For i = 0 To GetNumVerts - 2
    '        dPIC.Line (GetPointX(i), GetPointY(i))-(GetPointX(i + 1), GetPointY(i + 1)), vbBlue
    '    Next
    'End If

    For t = 0 To 1 Step 1 / Iterations    '1000    '1000 iterations (increase for a smoother curve)
        GetBezierPoint t, X, Y    'find the positions for the curve
        'dPIC.PSet (x, y), Color 'draw curve

        Long2RGB GetPixel(dPIC.hdc, X - 5, Y - 7), R2, G2, B2

        R2 = (R1 + R2) * 0.5
        G2 = (g1 + G2) * 0.5
        B2 = (b1 + B2) * 0.5

        SetPixel dPIC.hdc, X, Y, RGB(R2, G2, B2)    ' Color

    Next

    'draw handles
    'If chkPoints.Value Then
    '    For i = 0 To GetNumVerts - 1
    '        dPIC.Circle (Points(i).x, Points(i).y), 3, vbBlue
    '    Next
    'End If



    dPIC.Refresh

End Sub






















Public Sub Render(ByRef dPIC As PictureBox, Iterations As Long, R As Single, G As Single, B As Single)
Dim Color     As Long

Dim R1        As Integer
Dim g1        As Integer
Dim b1        As Integer
Dim R2        As Integer
Dim G2        As Integer
Dim B2        As Integer
Dim C         As Long


Dim t         As Double
Dim X         As Double
Dim Y         As Double
Dim i         As Long

Dim PrevX     As Double
Dim PrevY     As Double

    Color = RGB(R, G, B)

    'dPIC.Cls
    'draw lines
    'If chkLines.Value Then
    '    For i = 0 To GetNumVerts - 2
    '        dPIC.Line (GetPointX(i), GetPointY(i))-(GetPointX(i + 1), GetPointY(i + 1)), vbBlue
    '    Next
    'End If

    '**********
    'FX.GetBits dPIC.Image.Handle
    FX.ClearBezierBits

    '***********
    PrevX = 0
    PrevY = 0

    For t = 0 To 1 Step 1 / Iterations    '1000    '1000 iterations (increase for a smoother curve)
        GetBezierPoint t, X, Y    'find the positions for the curve
        'dPIC.PSet (x, y), Color 'draw curve

        'Long2RGB GetPixel(dPIC.hdc, x - 5, y - 7), R2, g2, b2

        'R2 = (R1 + R2) * 0.5
        'g2 = (g1 + g2) * 0.5
        'b2 = (b1 + b2) * 0.5

        ''SetPixel dPIC.hdc, X, Y, Color 'RGB(R2, g2, b2) ' Color
        ''SetPixel dPIC.hdc, X + 1, Y, Color 'RGB(R2, g2, b2) ' Color
        ''SetPixel dPIC.hdc, X, Y + 1, Color 'RGB(R2, g2, b2) ' Color
        ''SetPixel dPIC.hdc, X + 1, Y + 1, Color 'RGB(R2, g2, b2) ' Color
        'Stop

        If PrevX <> X \ 1 Or PrevY <> Y \ 1 Then
            FX.SetBezierBits X \ 1, Y \ 1, R, G, B
            PrevX = X \ 1
            PrevY = Y \ 1
        End If

    Next
    '    Stop

    FX.PaintBezierImage
    FX.SetBits dPIC.Image.Handle

    'draw handles
    'If chkPoints.Value Then
    '    For i = 0 To GetNumVerts - 1
    '        dPIC.Circle (Points(i).x, Points(i).y), 3, vbBlue
    '    Next
    'End If

    'dPIC.Refresh

    'If Rnd < 0.002 Then
    '    FX.GetBits dPIC.Image.Handle
    '    FX.BLUR
    '    FX.SetBits dPIC.Image.Handle
    'End If

    '*********
    'dPIC.Refresh
    '*********
End Sub

