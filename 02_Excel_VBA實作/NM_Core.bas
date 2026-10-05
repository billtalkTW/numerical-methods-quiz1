Attribute VB_Name = "Module1"
Option Explicit

' Numerical Methods quiz practice. All arrays from Range are 1-based.
Public Function NM_E1(ByVal x As Double) As Variant
    If x <= 1 Then
        NM_E1 = CVErr(xlErrNum)
    Else
        NM_E1 = Log(x - 1) + Cos(x - 1)
    End If
End Function

Public Function NM_Lagrange(ByVal data As Range, ByVal x As Double) As Variant
    Dim a As Variant, n As Long, i As Long, j As Long, m As Double, total As Double
    On Error GoTo BadInput
    a = data.Value2
    n = UBound(a, 1)
    If UBound(a, 2) <> 2 Or n < 2 Then GoTo BadInput
    For i = 1 To n
        If Not IsNumeric(a(i, 1)) Or Not IsNumeric(a(i, 2)) Then GoTo BadInput
        m = 1
        For j = 1 To n
            If i <> j Then
                If a(i, 1) = a(j, 1) Then GoTo BadInput
                m = m * (x - a(j, 1)) / (a(i, 1) - a(j, 1))
            End If
        Next j
        total = total + m * a(i, 2)
    Next i
    NM_Lagrange = total
    Exit Function
BadInput:
    NM_Lagrange = CVErr(xlErrValue)
End Function

Public Function NM_NewtonCoeff(ByVal data As Range, ByVal k As Long) As Variant
    Dim a As Variant, c() As Double, n As Long, i As Long, j As Long
    On Error GoTo BadInput
    a = data.Value2: n = UBound(a, 1)
    If UBound(a, 2) <> 2 Or k < 0 Or k >= n Then GoTo BadInput
    ReDim c(1 To n)
    For i = 1 To n
        c(i) = CDbl(a(i, 2))
    Next i
    For j = 1 To k
        For i = n To j + 1 Step -1
            If a(i, 1) = a(i - j, 1) Then GoTo BadInput
            c(i) = (c(i) - c(i - 1)) / (a(i, 1) - a(i - j, 1))
        Next i
    Next j
    NM_NewtonCoeff = c(k + 1)
    Exit Function
BadInput:
    NM_NewtonCoeff = CVErr(xlErrValue)
End Function

Public Function NM_NewtonPoly(ByVal data As Range, ByVal x As Double) As Variant
    Dim a As Variant, n As Long, k As Long, prod As Double, total As Double, c As Variant
    On Error GoTo BadInput
    a = data.Value2: n = UBound(a, 1): prod = 1
    For k = 0 To n - 1
        c = NM_NewtonCoeff(data, k)
        If IsError(c) Then GoTo BadInput
        total = total + c * prod
        prod = prod * (x - a(k + 1, 1))
    Next k
    NM_NewtonPoly = total
    Exit Function
BadInput:
    NM_NewtonPoly = CVErr(xlErrValue)
End Function

Public Function NM_D1(ByVal data As Range, ByVal x As Double, ByVal h As Double) As Variant
    On Error GoTo BadInput
    If h <= 0 Then GoTo BadInput
    NM_D1 = (NM_Lagrange(data, x + h) - NM_Lagrange(data, x - h)) / (2 * h)
    Exit Function
BadInput:
    NM_D1 = CVErr(xlErrValue)
End Function

Public Function NM_D2(ByVal data As Range, ByVal x As Double, ByVal h As Double) As Variant
    On Error GoTo BadInput
    If h <= 0 Then GoTo BadInput
    NM_D2 = (NM_Lagrange(data, x + h) - 2 * NM_Lagrange(data, x) + NM_Lagrange(data, x - h)) / h ^ 2
    Exit Function
BadInput:
    NM_D2 = CVErr(xlErrValue)
End Function

Public Function NM_ExpTaylor(ByVal x As Double, ByVal degree As Long) As Variant
    Dim k As Long, term As Double, total As Double
    If degree < 0 Or degree > 170 Then NM_ExpTaylor = CVErr(xlErrNum): Exit Function
    term = 1: total = 1
    For k = 1 To degree
        term = term * x / k: total = total + term
    Next k
    NM_ExpTaylor = total
End Function

Public Function NM_SinTaylor(ByVal x As Double, ByVal degree As Long) As Variant
    Dim k As Long, term As Double, total As Double
    If degree < 0 Or degree > 170 Then NM_SinTaylor = CVErr(xlErrNum): Exit Function
    If degree = 0 Then NM_SinTaylor = 0: Exit Function
    term = x: total = x
    For k = 3 To degree Step 2
        term = -term * x * x / (k * (k - 1)): total = total + term
    Next k
    NM_SinTaylor = total
End Function

Public Function NM_AddUp(ByVal x As Double, ByVal n As Long) As Variant
    Dim i As Long, total As Double
    If n < 0 Or n > 1000000 Then NM_AddUp = CVErr(xlErrNum): Exit Function
    For i = 1 To n: total = total + x: Next i
    NM_AddUp = total
End Function

Private Function NM_Eval(ByVal fn As Long, ByVal x As Double, ByVal data As Range, ByVal h As Double) As Variant
    On Error GoTo BadInput
    Select Case fn
        Case 1: NM_Eval = NM_E1(x)
        Case 2: NM_Eval = Cos(x) - x
        Case 3: NM_Eval = NM_Lagrange(data, x)
        Case 4: NM_Eval = NM_D1(data, x, h)
        Case 5: NM_Eval = NM_D2(data, x, h)
        Case Else: GoTo BadInput
    End Select
    Exit Function
BadInput:
    NM_Eval = CVErr(xlErrValue)
End Function

Private Function NM_Derivative(ByVal fn As Long, ByVal x As Double, ByVal data As Range, ByVal h As Double) As Variant
    Dim delta As Double
    On Error GoTo BadInput
    Select Case fn
        Case 1
            If x <= 1 Then GoTo BadInput
            NM_Derivative = 1 / (x - 1) - Sin(x - 1)
        Case 2: NM_Derivative = -Sin(x) - 1
        Case 3: NM_Derivative = NM_D1(data, x, h)
        Case 4: NM_Derivative = NM_D2(data, x, h)
        Case Else
            delta = h
            NM_Derivative = (NM_Eval(fn, x + delta, data, h) - NM_Eval(fn, x - delta, data, h)) / (2 * delta)
    End Select
    Exit Function
BadInput:
    NM_Derivative = CVErr(xlErrNum)
End Function

' fn: 1=log(x-1)+cos(x-1); 2=cos(x)-x; 3=P; 4=D1(P); 5=D2(P).
' For Newton a is x0 and b is ignored. k coefficients start at zero.
Public Function NM_Root(ByVal method As String, ByVal fn As Long, ByVal a As Double, ByVal b As Double, ByVal tol As Double, ByVal maxIter As Long, Optional ByVal data As Range, Optional ByVal h As Double = 0.001) As Variant
    Dim status As String
    NM_Root = NM_Solve(method, fn, a, b, tol, maxIter, data, h, Nothing, status)
End Function

Private Function NM_Solve(ByVal method As String, ByVal fn As Long, ByVal a As Double, ByVal b As Double, ByVal tol As Double, ByVal maxIter As Long, ByVal data As Range, ByVal h As Double, ByVal logTop As Range, ByRef status As String) As Variant
    Dim fa As Variant, fb As Variant, fc As Variant, df As Variant
    Dim c As Double, last As Double, change As Double, den As Double
    Dim i As Long, bracket As Boolean
    On Error GoTo Invalid
    method = LCase$(method)
    If tol <= 0 Or maxIter < 1 Or maxIter > 100000 Or h <= 0 Then GoTo Invalid
    bracket = (method = "bisection" Or method = "falseposition")
    If Not bracket And method <> "newton" And method <> "secant" Then GoTo Invalid
    fa = NM_Eval(fn, a, data, h)
    If IsError(fa) Then status = "Invalid initial point / domain": GoTo Failed
    If fa = 0 Then NM_Solve = a: status = "Initial point is a root": Exit Function
    If method <> "newton" Then
        fb = NM_Eval(fn, b, data, h)
        If IsError(fb) Then status = "Invalid second point / domain": GoTo Failed
        If fb = 0 Then NM_Solve = b: status = "Endpoint is a root": Exit Function
        If bracket And (a >= b Or Sgn(fa) = Sgn(fb)) Then status = "Need a<b and opposite signs": GoTo Failed
    End If
    last = a
    For i = 1 To maxIter
        Select Case method
            Case "bisection": c = a + (b - a) / 2
            Case "falseposition"
                den = fb - fa
                If Abs(den) < 0.000000000000001 Then status = "Near-zero denominator": GoTo Failed
                c = a - fa * (b - a) / den
            Case "newton"
                df = NM_Derivative(fn, a, data, h)
                If IsError(df) Then status = "Invalid derivative": GoTo Failed
                If Abs(df) < 0.00000000000001 Then status = "Near-zero derivative": GoTo Failed
                c = a - fa / df
            Case "secant"
                den = fb - fa
                If Abs(den) < 0.000000000000001 Then status = "Near-zero denominator": GoTo Failed
                c = b - fb * (b - a) / den
        End Select
        If fn = 1 And c <= 1 Then
            status = "STOP: candidate x<=1 (before Log)"
            If Not logTop Is Nothing Then
                logTop.Offset(i, 0).Value2 = i
                logTop.Offset(i, 3).Value2 = c
                logTop.Offset(i, 6).Value2 = status
            End If
            GoTo Failed
        End If
        fc = NM_Eval(fn, c, data, h)
        If IsError(fc) Then status = "Invalid function value": GoTo Failed
        If method = "bisection" Then change = Abs(b - a) / 2 Else change = Abs(c - last)
        If Not logTop Is Nothing Then
            logTop.Offset(i, 0).Value2 = i
            logTop.Offset(i, 1).Value2 = a
            logTop.Offset(i, 2).Value2 = b
            logTop.Offset(i, 3).Value2 = c
            logTop.Offset(i, 4).Value2 = fc
            logTop.Offset(i, 5).Value2 = change
        End If
        If fc = 0 Or (change <= tol And Abs(fc) <= tol) Then
            status = "Converged": NM_Solve = c
            If Not logTop Is Nothing Then logTop.Offset(i, 6).Value2 = status
            Exit Function
        End If
        If bracket Then
            If Sgn(fa) <> Sgn(fc) Then b = c: fb = fc Else a = c: fa = fc
        ElseIf method = "newton" Then
            a = c: fa = fc
        Else
            a = b: fa = fb: b = c: fb = fc
        End If
        last = c
    Next i
    status = "Max iterations reached; not certified converged"
Failed:
    NM_Solve = CVErr(xlErrNum): Exit Function
Invalid:
    status = "Invalid input / arithmetic error": NM_Solve = CVErr(xlErrValue)
End Function

Public Sub NM_RunE1Logs()
    Dim ws As Worksheet, target As Worksheet, status As String, answer As Variant
    Dim methods As Variant, j As Long, a As Double, b As Double
    Set ws = ThisWorkbook.Worksheets("E1_¨D®Ú")
    Set target = ThisWorkbook.Worksheets("E1_­¡¥N")
    target.Range("A5:AO110").ClearContents
    methods = Array("Bisection", "FalsePosition", "Newton", "Secant", "Newton")
    For j = 0 To 4
        If j < 2 Then
            a = ws.Range("B5").Value2: b = ws.Range("C5").Value2
        ElseIf j = 2 Then
            a = ws.Range("B6").Value2: b = 0
        ElseIf j = 3 Then
            a = ws.Range("B7").Value2: b = ws.Range("C7").Value2
        Else
            a = ws.Range("B8").Value2: b = 0
        End If
        answer = NM_Solve(CStr(methods(j)), 1, a, b, ws.Range("B9").Value2, ws.Range("C9").Value2, Nothing, 0.001, target.Cells(4, 1 + j * 8), status)
        target.Cells(2, 1 + j * 8).Value2 = CStr(methods(j)) & " : " & status
    Next j
End Sub
