Attribute VB_Name = "tools"
Option Explicit

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
Public Function degMinSec(degrees As Double) As String
    degMinSec = CStr(angDeg(degrees)) & ChrW(176) & CStr(angMin(degrees)) & "'" & CStr(angSec(degrees)) & "''"
End Function
Public Function Rad(deg As Double) As Double
    Const Pi As Double = 3.1415927
    Rad = deg * Pi / 180
End Function
