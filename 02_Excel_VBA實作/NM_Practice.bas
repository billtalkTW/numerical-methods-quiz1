Attribute VB_Name = "Module1"
Option Explicit

' Scaffold only. Fill each TODO, then test against the demonstration workbook.
Public Function NM_TryLagrange(ByVal data As Range, ByVal x As Double) As Variant
    Dim a As Variant, i As Long, j As Long, n As Long
    Dim product As Double, total As Double
    a = data.Value2
    n = UBound(a, 1)
    ' TODO: Outer loop i; product=1; inner loop j skipping i.
    ' TODO: Reject duplicate x nodes; accumulate y(i)*basis(i).
    NM_TryLagrange = CVErr(xlErrNA)
End Function

Public Function NM_TryD1(ByVal data As Range, ByVal x As Double, ByVal h As Double) As Variant
    ' TODO: Validate h, evaluate x+h and x-h, divide the difference.
    NM_TryD1 = CVErr(xlErrNA)
End Function

Public Function NM_TryNewton(ByVal x0 As Double, ByVal tol As Double, ByVal maxIter As Long) As Variant
    Dim x1 As Double, fx As Double, df As Double, k As Long
    ' TODO: Validate x0>1 before Log; use f=Log(x-1)+Cos(x-1).
    ' TODO: Guard derivative, candidate domain, convergence and maxIter.
    NM_TryNewton = CVErr(xlErrNA)
End Function

Public Sub NM_PracticeGuide()
    MsgBox "Fill the TODO blocks in Module1. Do not treat #N/A as an answer.", vbInformation, "NM practice"
End Sub
