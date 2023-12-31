''
'This function is designed to determine if three points (x1, y1), (x, y), and (x2, y2) are collinear, which means they lie on the same straight line.
'It first calculates the distances dist1, dist2, and totalDist using the distanceXY function.
    'dist1 is the distance between (x1, y1) and (x, y).
    'dist2 is the distance between (x, y) and (x2, y2).
    'totalDist is the distance between (x1, y1) and (x2, y2).
'It then checks if the sum of dist1 and dist2 is equal to totalDist. 
'If they are equal, the points are collinear, and the function returns True. Otherwise, it returns False.
''
Public Function IsCollinear(x1 As Double, y1 As Double, x As Double, y As Double, x2 As Double, y2 As Double) As Boolean
    Dim dist1 As Double, dist2 As Double, totalDist As Double
    dist1 = distanceXY(x1, y1, x, y)
    dist2 = distanceXY(x, y, x2, y2)
    totalDist = distanceXY(x1, y1, x2, y2)
    
    If (totalDist = dist1 + dist2) Then
        isCollinear = True
    Else
        isCollinear = False
    End If
End Function

''
'This function calculates the Euclidean distance between two points (x1, y1) and (x2, y2) in a 2D plane.
''
Public Function distanceXY(x1 As Double, y1 As Double, x2 As Double, y2 As Double) As Double
    distanceXY = Round(Sqr((x2 - x1) ^ 2 + (y2 - y1) ^ 2), 2)
End Function

'Returns the area of polygon formed by the given points.
'This function calculates the area of a polygon defined by a set of points (X(), Y()). 
'It uses a loop to calculate the signed area of trapezoid-like sections formed by the polygon's edges and vertices, 
'and then it returns the absolute value of the total area divided by 2.
''
Private Function Area(ByRef X() As Double, ByRef Y() As Double) As Double
    Dim i As Short
    Dim g As Double
    Dim n As Short
    Dim h As Double

    n = UBound(X) - LBound(X)
    g = 0.0#
    For i = 0 To n
        If i <> n Then
            h = (X(i + 1) - X(i)) * (Y(i) + Y(i + 1))
        Else
            h = (X(0) - X(n)) * (Y(0) + Y(n))
        End If
        g = g + h
    Next i
    Area = System.Math.Abs(g) / 2
End Function
