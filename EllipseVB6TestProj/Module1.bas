Attribute VB_Name = "Module1"
Option Explicit
Const pi = 3.14159265358979
Public Type Complex
    re As Double
    im As Double
End Type

Dim c1 As Complex
Dim c2 As Complex
Dim y1 As Complex
Dim y2 As Complex
Dim y3 As Complex

Dim sqrt1 As Complex
Dim sqrt2 As Complex
Dim sqrt3 As Complex

Dim aInit As Double
Dim bInit As Double
Dim cInit As Double
Dim dInit As Double
Dim eInit As Double
Dim a, b, c, d, p, q, m, n, o, pN, qN, r As Double
Dim realSolutions(3) As Double
Dim z1 As Double
Dim z2 As Double
Dim z3 As Double
Dim z4 As Double
Dim zz1 As Complex
Dim zz2 As Complex
Dim zz3 As Complex
Dim zz4 As Complex
Dim tolerance As Double
    
Dim sum1 As Complex
Dim u, v, w As Double

Dim x1 As Complex
Dim x2 As Complex
Dim x3 As Complex

Dim xx1 As Double
Dim xx2 As Double

Dim rr As Double
Dim tt As Double
Public Function GetPoints(ByVal xoK As Double, ByVal yoK As Double, ByVal radius As Double, ByVal xoE As Double, ByVal yoE As Double, ByVal ellA As Double, ByVal ellB As Double) As Double()

    aInit = Pow(ellA, 2) * (Pow(yoK - yoE, 2) - Pow(ellB, 2)) + Pow(ellB, 2) * Pow((xoK - xoE) - radius, 2)
    bInit = 4 * Pow(ellA, 2) * radius * (yoK - yoE)
    cInit = 2 * (Pow(ellA, 2) * (Pow(yoK - yoE, 2) - Pow(ellB, 2) + 2 * Pow(radius, 2)) + Pow(ellB, 2) * (Pow(xoK - xoE, 2) - Pow(radius, 2)))
    dInit = 4 * Pow(ellA, 2) * radius * (yoK - yoE)
    eInit = Pow(ellA, 2) * (Pow(yoK - yoE, 2) - Pow(ellB, 2)) + Pow(ellB, 2) * Pow((xoK - xoE) + radius, 2)
    a = bInit / aInit
    b = cInit / aInit
    c = dInit / aInit
    d = eInit / aInit
    p = b - 3 * Pow(a, 2) / 8
    q = Pow(a, 3) / 8 - a * b / 2 + c
    r = -(3 * Pow(a, 4) - 16 * Pow(a, 2) * b + 64 * a * c - 256 * d) / 256#
    m = -2 * p
    n = Pow(p, 2) - 4 * r
    o = Pow(q, 2)
    pN = n - Pow(m, 2) / 3
    qN = 2 * Pow(m, 3) / 27 - m * n / 3 + o
    rr = Pow(qN / 2, 2) + Pow(pN / 3, 3)
    tolerance = 0.0000001



    If (Abs(a) < tolerance) And (Abs(c) < tolerance) Then
        xx1 = -b / 2 + Sqr(Pow(-b / 2, 2) - d)
        xx2 = -b / 2 - Sqr(Pow(-b / 2, 2) - d)
        z1 = Sqr(xx1)
        z2 = Sqr(xx2)
        z3 = -z1
        z4 = -z2
        realSolutions(0) = z1
        realSolutions(1) = z2
        realSolutions(2) = z3
        realSolutions(3) = z4
    Else


        If rr >= 0 Then
            tt = Sqr(rr)
            c1.re = (-1# * qN / 2) + tt
            c2.re = (-1# * qN / 2) - tt

            
            Dim signC1, signC2 As Double
            signC1 = 1#
            signC2 = 1#
            If c1.re < 0 Then
                signC1 = -1#
            End If
            
            If c2.re < 0 Then
                signC2 = -1#
            End If

            u = signC1 * Magnitude(ComplexPow(c1, 1# / 3))
            v = signC2 * Magnitude(ComplexPow(c2, 1# / 3))
            
            y1.re = u + v
            y1.im = 0
            y2.re = -(u + v) / 2
            y2.im = -((u - v) / 2) * Sqr(3)
            y3.re = -(u + v) / 2
            y3.im = ((u - v) / 2) * Sqr(3)
        Else
            u = -1# * Pow(pN / 3, 3)
            u = Pow(u, 1# / 2)
            w = ArcCos(-qN / (2 * u))
            y1.re = 2 * Pow(u, 1# / 3) * Cos(w / 3)
            y1.im = 0
            y2.re = 2 * Pow(u, 1# / 3) * Cos(w / 3 + 2 * pi / 3)
            y2.im = 0
            y3.re = 2 * Pow(u, 1# / 3) * Cos(w / 3 + 4 * pi / 3)
            y3.im = 0
        End If

        x1.re = (y1.re - (m / 3))
        x1.im = y1.im
        x2.re = (y2.re - (m / 3))
        x2.im = y2.im
        x3.re = (y3.re - (m / 3))
        x3.im = y3.im

        sqrt1 = ComplexPow(Negate(x1), 1# / 2)
        sqrt2 = ComplexPow(Negate(x2), 1# / 2)
        sqrt3 = ComplexPow(Negate(x3), 1# / 2)
        sum1 = MultiplyComplex(MultiplyComplex(sqrt1, sqrt2), sqrt3)

        If Not (Abs(sum1.re + q) <= 0.00001) Then
            sqrt1 = Negate(sqrt1)
            sqrt2 = Negate(sqrt2)
            sqrt3 = Negate(sqrt3)
        End If

        zz1 = MultiplyComplexWithDouble(AddComplex(sqrt1, AddComplex(sqrt2, sqrt3)), 0.5)
        zz2 = MultiplyComplexWithDouble(AddComplex(sqrt1, Negate(AddComplex(sqrt2, sqrt3))), 0.5)
        zz3 = MultiplyComplexWithDouble(AddComplex(Negate(sqrt1), AddComplex(sqrt2, Negate(sqrt3))), 0.5)
        zz4 = MultiplyComplexWithDouble(AddComplex(Negate(sqrt1), AddComplex(Negate(sqrt2), sqrt3)), 0.5)
        zz1.re = zz1.re - a / 4
        zz2.re = zz2.re - a / 4
        zz3.re = zz3.re - a / 4
        zz4.re = zz4.re - a / 4
        Dim imaginarySolutions(3) As Complex
        imaginarySolutions(0) = zz1
        imaginarySolutions(1) = zz2
        imaginarySolutions(2) = zz3
        imaginarySolutions(3) = zz4

        Dim index As Integer
        For index = 0 To 3
            If Abs(imaginarySolutions(index).im) < 0.0001 Then
                realSolutions(index) = imaginarySolutions(index).re
            Else
                realSolutions(index) = -10000#
            End If
        Next index
    End If

    Dim sols(7) As Double
    Dim foo As Double
    For index = 0 To 3
 
        Dim zPow2, x, y As Double
        foo = realSolutions(index)
        If foo > -10000# Then
            zPow2 = Pow(foo, 2)
            sols(2 * index) = xoK + radius * (1 - zPow2) / (1# + zPow2)
            sols(2 * index + 1) = yoK + radius * (2 * foo) / (1 + zPow2)
        End If
    Next index
    GetPoints = sols
End Function

Public Function MultiplyComplex(a As Complex, b As Complex) As Complex

    MultiplyComplex.re = a.re * b.re - a.im * b.im
    MultiplyComplex.im = a.re * b.im + a.im * b.re

End Function

Public Function AddComplex(a As Complex, b As Complex) As Complex

    AddComplex.re = a.re + b.re
    AddComplex.im = a.im + b.im

End Function
Private Function GetSolutionPoint(ByVal z As Double, ByVal r As Double, ByVal xoK As Double, ByVal yoK As Double) As Double()
    Dim zPow2, x, y As Double
    zPow2 = Pow(z, 2)
    x = xoK + r * (1 - zPow2) / (1# + zPow2)
    y = yoK + r * (2 * z) / (1 + zPow2)
    GetSolutionPoint(0) = x
    GetSolutionPoint(1) = y
End Function



Public Function MultiplyComplexWithDouble(ByRef a As Complex, ByRef b As Double) As Complex

MultiplyComplexWithDouble.re = a.re * b
MultiplyComplexWithDouble.im = a.im * b

End Function


Public Function Pow(ByVal v As Double, ByVal e As Double) As Double
    Pow = v ^ e
End Function

Public Function Negate(ByRef v As Complex) As Complex
    Negate.re = -v.re
    Negate.im = -v.im
End Function
Public Function Pow2Complex(ByRef a As Complex) As Complex

    Pow2Complex.re = a.re * a.re - a.im * a.im
    Pow2Coomplex.im = 2 * a.re * a.im

End Function

Public Function Conjugate(ByRef a As Complex) As Complex

Conjugate.re = a.re
Conjugate.im = -a.im

End Function

Public Function ComplexPow2(ByRef v As Complex) As Complex
    ComplexPow2.re = v.re * v.re - v.im * v.im
    ComplexPow2.im = 2 * v.re * v.im
End Function

' // Exponentiation of a complex number
Public Function ComplexPow(ByRef Op As Complex, ByRef e As Double) As Complex
    Dim md, ar As Double
    Dim foo As Complex
    md = cxMod(Op)
    ar = cxArg(Op)
    md = md ^ e
    ar = ar * e
    foo.re = md * Cos(ar)
    foo.im = md * Sin(ar)
    ComplexPow = foo
End Function

Public Function cxArg(ByRef Op As Complex) As Double
    If Op.im = 0 Then
        If Op.re >= 0 Then cxArg = 0 Else cxArg = pi
    ElseIf Op.re = 0 Then
        If Op.im >= 0 Then cxArg = pi / 2 Else cxArg = -pi / 2
    Else
        If Op.re > 0 Then
            cxArg = Atn(Op.im / Op.re)
        ElseIf Op.re < 0 And Op.im > 0 Then
            cxArg = pi + Atn(Op.im / Op.re)
        ElseIf Op.re < 0 And Op.im < 0 Then
            cxArg = -pi + Atn(Op.im / Op.re)
        End If
    End If
End Function
Public Function cxMod(ByRef Op As Complex) As Double
    Dim R2 As Double, i2 As Double
    R2 = Op.re * Op.re: i2 = Op.im * Op.im
    cxMod = Sqr(R2 + i2)
End Function
Public Function Magnitude(ByRef v As Complex) As Double
    Dim mag As Double
    mag = Sqr(Pow(v.re, 2) + Pow(v.im, 2))
    Magnitude = mag
End Function


Public Function ImaginaryOne() As Complex
    Dim foo As Complex
    foo.re = 0
    foo.im = 1
    ImaginaryOne = foo
End Function




Public Function ArcCos(x As Double) As Double
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function
