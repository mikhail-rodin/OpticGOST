Attribute VB_Name = "tools"
Option Explicit
Public Function BoolToInt(b As Boolean) As Integer
'in this ancient language boolean true is stored as 16 bit 111...1, which has a -1 int value.
'VBA, I don't have any patience for your senile nonsense
'thus we need a custom CInt
    If b Then BoolToInt = 1 Else BoolToInt = 0
End Function

'quicksort algorithm courtesy of Jorge Ferreira
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub
Public Sub swap(ByRef arr() As Variant, i As Integer, j As Integer)
    Dim temp As Variant
    temp = arr(i)
    arr(i) = arr(j)
    arr(j) = temp
End Sub
Public Function angDeg(degrees As Double) As Integer
    angDeg = Fix(degrees)
End Function
Public Function angMin(degrees As Double) As Integer
    Dim deg As Double
    deg = Fix(degrees)
    angMin = Int((Abs(degrees - deg)) * 60)
End Function
Public Function angSec(degrees As Double) As Integer
    Dim deg As Double
    deg = Fix(degrees)
    angSec = Int(((Abs(degrees - deg)) * 60 - Int((Abs(degrees - deg)) * 60)) * 60)
End Function
Public Function angFractSec(degrees As Double) As Double
    Dim deg As Double
    deg = Fix(degrees)
    angFractSec = ((Abs(degrees - deg)) * 60 - Int((Abs(degrees - deg)) * 60)) * 60
End Function
Public Function degMinSec(degrees As Double) As String
    Dim deg As Integer, min As Integer, sec As Double
    deg = angDeg(degrees)
    min = angMin(degrees)
    sec = angFractSec(degrees)
    If deg = 0 Then
        If min = 0 Then
            degMinSec = Replace(CStr(Round(sec, 2)), ".", ",") & "''"
        Else
            degMinSec = CStr(min) & "'" & CStr(Int(sec)) & "''"
        End If
    Else
        degMinSec = CStr(deg) & ChrW(176) & CStr(min) & "'" & CStr(Int(sec)) & "''"
    End If
End Function
Public Function degMin(degrees As Double) As String
    Dim deg As Integer, min As Integer
    deg = angDeg(degrees)
    min = angMin(degrees)
    If deg = 0 Then
        degMin = CStr(min) & "'"
    Else
        degMin = CStr(deg) & ChrW(176) & CStr(min) & "'"
    End If
End Function
Public Function rad(deg As Double) As Double
    Const Pi As Double = 3.1415927
    rad = deg * Pi / 180
End Function
Public Function deg(rad As Double) As Double
    Const Pi As Double = 3.1415927
    deg = rad * 180 / Pi
End Function
Public Function ArcCos(A As Double) As Double 'in radians
  'Inverse Cosine
    On Error Resume Next
        If A = 1 Then
            ArcCos = 0
            Exit Function
        End If
        ArcCos = Atn(-A / Sqr(-A * A + 1)) + 2 * Atn(1)
    On Error GoTo 0
End Function
Public Function ArcSin(ByVal x As Double) As Double 'in radians
    If x = 1 Then
        ArcSin = 0
        Exit Function
    Else
        ArcSin = Atn(x / Sqr(-x * x + 1))
    End If
End Function
