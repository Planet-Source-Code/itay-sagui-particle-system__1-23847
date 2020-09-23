Attribute VB_Name = "Module1"
Public Const PiRad = 3.14 / 180

Public TheHDC As Long
Public QuitRender As Boolean
Public Speed As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public ShapeNum As Byte

Public Type RGBCol
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Type PointAPI
    x As Long
    Y As Long
End Type

Public Type AttracPoint
    XLoc As Long
    YLoc As Long
    ZLoc As Long
    Mass As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public BoundingRect As RECT
Public PartsImages() As Picture

Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)


'List of parameters the user should be able to change:
'Xloc, YLoc, ZLoc
'Xvel (Not in Circular,Conus,Cylinder,Snow,Spiral,Tornado)
'Yvel (Not in Circular,Spiral)
'Zvel (Not in Circular,Conus,Cylinder,Snow,Spiral,Tornado)
'XAcl, YAcl, ZAcl
'Lifespan (In Circular, Spiral should be the same for all
'          particles
'Life (I'm not sure about this one.)
'Wind (Don't affect Circular, Conus,Cylinder,Spiral,Tornado)
'SizeType (1-Expanding,2-Shrinking,3-Accoring2Z)
'PartType (1-SimpleBrush,2-Images)
'Colors

Public Sub Main()
    Randomize Timer
    Form1.Show
    TheHDC = Form1.Picture1.hDC
    Form1.cboPS.ListIndex = 0
    Form1.Slider1_Change
End Sub

Public Sub HandleError(ByVal ModuleName As String, ByVal ProcedureName As String, ByVal ErrNumber As Long, ByVal ErrDescription As String)
    MsgBox "Module: " & ModuleName & vbCrLf & _
    "Procedure: " & ProcedureName & vbCrLf & vbCrLf & _
    "Err Number: " & ErrNumber & vbCrLf & _
    "Err Description: " & ErrDescription, vbCritical
End Sub

Public Sub ShutDown()
    QuitRender = True
    DoEvents
    End
End Sub

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

Public Function checkShape(ByVal x As Long, ByVal Y As Long) As Byte
Dim i As Shape
    For Each i In Form1.Shape
        With i
            If x > .Left And x < (.Left + .Width) And _
               Y > .Top And Y < (.Top + .Height) Then
                checkShape = i.Index
            End If
        End With
    Next
End Function

