'Option Explicit
Dim tera As Double, giga As Double, mega As Double, kilo As Double, milli As Double, micro As Double, nano As Double, pico As Double, femt As Double


Dim iprmm
Dim x() As Double, y() As Double, func() As Double
Dim prm() As Double, prml() As Double, prmh() As Double
Dim vexsimplex() As Double, ressimplex() As Double, stepsimplex() As Double, tolsimplex() As Double
Dim prmbest() As Double




Sub Main()
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
Dim prmbest() As Double
Dim itermax As Integer


Sheet1.Select
Range("A1").Select

tera = 1000000000000#
giga = 1000000000#
mega = 1000000#
kilo = 1000#
milli = 0.001
micro = 0.000001
nano = 0.000001
pico = 0.000000000001
femt = 0.000000000000001



itrmax = 3000
itr = 0
kbest0 = 1E+120

'*** Read Measured Data ***
idatan = 51
ReDim x(idatan) As Double, y(idatan) As Double, func(idatan) As Double
For i = 1 To idatan
    x(i) = (CDbl(i) - 1#) / 10#
    ws.Cells(i, 1) = x(i)
    y(i) = 0.3 * x(i) - 1.7 * Exp(-3# * x(i)) + (1.4 + 1.3 * Exp(-1# * x(i))) + 0.3 * Sin(4.5 * x(i)) + Rnd(2) * 0.3
    ws.Cells(i, 2) = y(i)
Next i


Do While itr < itrmax

    DoEvents

'*** prms ***
iprmn = 8
ReDim prm(iprmn) As Double, prml(iprmn) As Double, prmh(iprmn) As Double
ReDim vexsimplex(iprmn + 1, iprmn) As Double, ressimplex(iprmn + 1) As Double, stepsimplex(iprmn) As Double, tolsimplex(iprmn) As Double
ReDim prmbest(iprmn) As Double

prml(1) = -100#: prmh(1) = 200#: prm(1) = (prml(1) + prmh(1)) / 2#
prml(2) = -20#: prmh(2) = 20#: prm(2) = (prml(2) + prmh(2)) / 2#
prml(3) = -50#: prmh(3) = 40#: prm(3) = (prml(3) + prmh(3)) / 2#
prml(4) = -30#: prmh(4) = 20#: prm(4) = (prml(4) + prmh(4)) / 2#
prml(5) = -20#: prmh(5) = 10#: prm(5) = (prml(5) + prmh(5)) / 2#
prml(6) = -10#: prmh(6) = 30#: prm(6) = (prml(6) + prmh(6)) / 2#
prml(7) = -60#: prmh(7) = 100#: prm(7) = (prml(7) + prmh(7)) / 2#
prml(8) = -10#: prmh(8) = 30#: prm(8) = (prml(8) + prmh(8)) / 2#


For j = 1 To iprmn
    vexsimplex(1, j) = prm(j)
    stepsimplex(j) = 10#
    tolsimplex(j) = 1# * femt
Next j


itermax = 1000
    kbest = Simplex(idatan, iprmn, vexsimplex(), stepsimplex(), ressimplex(), tolsimplex(), prm(), itermax)
    If kbest0 > kbest Then
        For i = 1 To iprmn
            prmbest(i) = prm(i)
        ws.Cells(i, 9) = prmbest(i)
        Next i
        kbest0 = kbest
    End If

    'Debug.Print "Simplex="; kbest
'    For j = 1 To iprmn
'        ws.Cells(j, 5) = j
'        ws.Cells(j, 6) = prm(j)
'    Next j
    'Debug.Print "res="; ressimplex(kbest) * 100#


'For i = 1 To idatan
'    ws.Cells(i, 7) = x(i)
'    ws.Cells(i, 8) = Abs(fx(x(i), iprmn, prm()))
'Next i

itr = itr + 1
Loop

For j = 1 To iprmn
    prmbest(j) = ws.Cells(j, 9)
Next j
For i = 1 To idatan
    ws.Cells(i, 7) = x(i)
    ws.Cells(i, 8) = Abs(fx(x(i), iprmn, prmbest()))
Next i
            
End Sub
Function fx(x, iprmn, prm() As Double)
'        px = 0
'        For j = 1 To iprmn
'            px = px + prm(j) * x ^ (j - 1)
'        Next j
        fx = prm(1) * x - prm(2) * Exp(prm(3) * x) + (prm(4) + prm(8) * Exp(prm(5) * x)) + prm(6) * Sin(prm(7) * x)
    'fx = px
End Function
Sub EvalFunction(idatan, iprmn, prm() As Double)
    For i = 1 To idatan
        'px = 0#
        'For j = 1 To iprmn - 1
            'func(i) = prm(1) + prm(2) * Exp(prm(3) * x(i)) + prm(4) * Exp(prm(5) * x(i)) - y(i)
        '    px = px + prm(j + 1) * Exp(prm(j) * x(i))
        'Next j
        func(i) = fx(x(i), iprmn, prm()) - y(i)
    Next i
End Sub
Function Simplex(idatan, iprmn, vexsimplex() As Double, stepsimplex() As Double, ressimplex() As Double, tolsimplex() As Double, prm() As Double, itermax As Integer)
    Dim ws As Worksheet
    Set ws = Worksheets("Sheet1")
    Dim p As Double
    
    vars = Sqr(0.5)
    v1 = vars * (Sqr(iprmn + 1#) - 1#) / CDbl(iprmn)
    For j = 1 To iprmn
        v2 = vexsimplex(1, j) + v1 * stepsimplex(j)
        For k = 2 To iprmn + 1#
            vexsimplex(k, j) = v2
        Next k
        vexsimplex(j + 1, j) = v2 + vars * stepsimplex(j)
    Next j

    For k = 1 To iprmn + 1
        For j = 1 To iprmn
            prm(j) = compare(vexsimplex(k, j), prml(j), prmh(j))
        Next j
        Call EvalFunction(idatan, iprmn, prm())
        resvar = SumsqSimplex(idatan)
        ressimplex(k) = resvar
    Next k
    Iteration = 0

ressimplex(kbest) = 1000#
imod = 1
    Do
        Iteration = Iteration + 1
        
        If Iteration Mod 50 = 0 Then
        
            ws.Cells(imod, 3) = imod
            ws.Cells(imod, 4) = ressimplex(kbest)
            imod = imod + 1
        End If
        
        If Iteration > itermax Then
            Simplex = kbest
            Exit Function
        End If


        kworst = 1
        kbest = 1

        For k = 2 To iprmn + 1
            If ressimplex(k) > ressimplex(kworst) Then kworst = k
            If ressimplex(k) < ressimplex(kbest) Then kbest = k
        Next k

        flag = 0
        For j = 1 To iprmn
            vmax = vexsimplex(1, j)
            vmin = vmax
            For k = 2 To iprmn + 1
                v = vexsimplex(k, j)
                If v > vmax Then vmax = v
                If v < vmin Then vmin = v
            Next k
            If (vmax - vmin) > tolsimplex(j) Then flag = 1
        Next j
        If flag = 0 Then
            Simplex = kbest
            Exit Function
            
        End If

        For j = 1 To iprmn
            vsum = 0
            For k = 1 To iprmn + 1
                If k <> kworst Then vsum = vsum + vexsimplex(k, j)
            Next k
            prm(j) = compare(2# * vsum / CDbl(iprmn) - vexsimplex(kworst, j), prml(j), prmh(j))
        Next j

        Call EvalFunction(idatan, iprmn, prm())
        resvar = SumsqSimplex(idatan)

        Select Case resvar
            Case Is < ressimplex(kbest)
                For j = 1 To iprmn
                    p = 1.5 * prm(j) - 0.5 * vexsimplex(kworst, j)
                    vexsimplex(kworst, j) = compare(prm(j), prml(j), prmh(j))
                    prm(j) = compare(p, prml(j), prmh(j))
                Next j
                ressimplex(kworst) = resvar
                Call EvalFunction(idatan, iprmn, prm())
                resvar = SumsqSimplex(idatan)
                If resvar < ressimplex(kworst) Then
                    For j = 1 To iprmn
                        vexsimplex(kworst, j) = prm(j)
                    Next j
                    ressimplex(kworst) = resvar
                End If
            Case ressimplex(kbest) To ressimplex(kworst)
                For j = 1 To iprmn
                    vexsimplex(kworst, j) = prm(j)
                Next j
                ressimplex(kworst) = resvar
            Case Is > ressimplex(kworst)
                For j = 1 To iprmn
                    prm(j) = compare(0.75 * vexsimplex(kworst, j) + 0.25 * prm(j), prml(j), prmh(j))
                Next j
                Call EvalFunction(idatan, iprmn, prm())
                resvar = SumsqSimplex(idatan)
                If resvar >= ressimplex(kworst) Then
                    For k = 1 To iprmn + 1
                        If k <> kbest Then
                            For j = 1 To iprmn
                                prm(j) = compare(0.5 * (vexsimplex(k, j) + vexsimplex(kbest, j)), prml(j), prmh(j))
                                vexsimplex(k, j) = prm(j)
                            Next j
                            Call EvalFunction(idatan, iprmn, prm())
                            resvar = SumsqSimplex(idatan)
                            ressimplex(k) = resvar
                        End If
                    Next k
                Else
                    For j = 1 To iprmn
                        vexsimplex(kworst, j) = prm(j)
                    Next j
                    ressimplex(kworst) = resvar
                End If

                Case Else

            End Select
    Loop

    Simplex = kbest

End Function

Function SumsqSimplex(idatan)
    resvar = 0#
    For i = 1 To idatan
        resvar = resvar + func(i) * func(i)
    Next i
    SumsqSimplex = resvar
End Function


Function compare(x As Double, low As Double, high As Double) As Double

    If x > high Then
        compare = high
    ElseIf x < low Then
        compare = low
    Else
        compare = x
    End If
    
End Function
