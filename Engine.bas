Attribute VB_Name = "Engine"
Public Type RECT
    x As Long
    y As Long
    w As Long
    h As Long
End Type

Public ShapeNum As Byte
'Public Shapes(100) As Shape

Public Function Max(Num1 As Long, Num2 As Long) As Long
    If Num1 > Num2 Then
        Max = Num1
    Else
        Max = Num2
    End If
End Function

Public Function Min(Num1 As Long, Num2 As Long) As Long
    If Num1 < Num2 Then
        Min = Num1
    Else
        Min = Num2
    End If
End Function

Public Function checkShape(ByVal x As Long, ByVal y As Long) As Byte
Dim i As Shape
    For Each i In Form1.Shape
        With i
            If x > .Left And x < (.Left + .Width) And _
               y > .Top And y < (.Top + .Height) Then
                checkShape = i.Index
            End If
        End With
    Next
End Function
