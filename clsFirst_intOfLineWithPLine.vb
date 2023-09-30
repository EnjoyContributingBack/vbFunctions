Option Strict Off
Option Explicit On

Public Class clsintOfLineWithPLine
    'by kabindra.
    'Date 2005-01-11
    Public Structure XY
        Dim X As Double
        Dim Y As Double
    End Structure
    Private nullValue As Single = -1.123456F

    'A line-> (x1,y1)-(x2,y2)
    'Polyline is array of (pd,Prl) and npd is number of points in the array.
    'scan direction is direction of scaning the polyline from end or start.
    'Intersection point returned -> (Xint,Yint)

    'Routines used for finding intersection points.
    '------------------------------------------------
    Public Sub GetIntPoint(ByRef Xint As Double, ByRef Yint As Double, ByRef x1 As Double, ByRef Y1 As Double, ByRef x2 As Double, ByRef Y2 As Double, ByRef pd() As Double, ByRef Prl() As Double, ByRef nPd As Integer, ByRef scanDir As Short)

        Dim Pt1 As XY = Nothing, PtXY As XY = Nothing, Pt2 As XY = Nothing

        Pt1.X = x1 : Pt1.Y = Y1
        Pt2.X = x2 : Pt2.Y = Y2

        PtXY = IntXY(pd, Prl, nPd, scanDir, Pt1, Pt2)

        Xint = PtXY.X
        Yint = PtXY.Y
    End Sub

    Private Function IntXY(ByRef pd() As Double, ByRef Prl() As Double, ByRef nP As Integer, ByRef scanDir As Short, ByRef Pt1 As XY, ByRef Pt2 As XY) As XY

        Dim Pt4 As XY = Nothing, Pt3 As XY = Nothing, tIntXy As XY = Nothing
        Dim EndI, i, StI, StepI As Integer

        If scanDir = -1 Then
            StI = nP : EndI = 1 : StepI = -1
        Else
            StI = 0 : EndI = nP - 1 : StepI = 1
        End If

        For i = StI To EndI Step StepI
            Pt3.X = pd(i)
            Pt3.Y = Prl(i)
            Pt4.X = pd(i + StepI)
            Pt4.Y = Prl(i + StepI)
            tIntXy = Intersection_Point(Pt1, Pt2, Pt3, Pt4)
            If tIntXy.X <> nullValue And tIntXy.Y <> nullValue Then
                Exit For
            End If
        Next i

        IntXY = tIntXy
    End Function

    'Gives intersection points for the lines joining given points.
    Public Function Intersection_Point(ByRef Pt1 As XY, ByRef Pt2 As XY, ByRef Pt3 As XY, ByRef Pt4 As XY) As XY
        Dim m1, m2 As Double
        Dim IntXY As XY = Nothing
        Const smallValue As Double = 0.000001

        If Pt3.X = Pt4.X Then Pt3.X = Pt4.X + smallValue
        m2 = (Pt3.Y - Pt4.Y) / (Pt3.X - Pt4.X)
        If Pt1.X = Pt2.X Then
            IntXY.X = Pt1.X
            IntXY.Y = m2 * (IntXY.X - Pt3.X) + Pt3.Y
            GoTo Last
        Else
            m1 = (Pt1.Y - Pt2.Y) / (Pt1.X - Pt2.X)
        End If

        If (m2 - m1) = 0 Then m2 = m1 - smallValue
        If System.Math.Abs(Pt1.X - Pt2.X) <> smallValue Then
            IntXY.X = (Pt3.Y - Pt1.Y + m1 * Pt1.X - m2 * Pt3.X) / (m1 - m2)
            If System.Math.Abs(Pt3.Y - Pt4.Y) >= smallValue Then
                IntXY.Y = m1 * (IntXY.X - Pt1.X) + Pt1.Y
            Else
                IntXY.Y = Pt3.Y
            End If
        Else
            IntXY.X = Pt1.X
            If System.Math.Abs(Pt3.Y - Pt4.Y) > smallValue Then
                IntXY.Y = m2 * (IntXY.X - Pt3.X) + Pt3.Y
            Else
                IntXY.Y = Pt3.Y
            End If
        End If
Last:
        If System.Math.Round(Distance(Pt1, IntXY) + Distance(IntXY, Pt2), 3) = System.Math.Round(Distance(Pt1, Pt2), 3) And System.Math.Round(Distance(Pt3, IntXY) + Distance(IntXY, Pt4), 3) = System.Math.Round(Distance(Pt3, Pt4), 3) Then
            Intersection_Point = IntXY
        Else
            Intersection_Point.X = nullValue : Intersection_Point.Y = nullValue
        End If
    End Function

    Private Function Distance(ByRef A As XY, ByRef b As XY) As Double
        Distance = ((A.X - b.X) ^ 2 + (A.Y - b.Y) ^ 2) ^ 0.5
    End Function
    '--------------------------------------------------------
End Class