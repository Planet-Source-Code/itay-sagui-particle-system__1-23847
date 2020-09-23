VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Circular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Circular"
Private Parts() As Particle

Public Sub InitCircular(Lft As Long, Tp As Long, _
        Wdth As Long, Hght As Long, Optional Num _
        As Long = 50, Optional Steps As Long = 50, _
        Optional MxSz As Integer = 3, Optional SzTp _
        As Byte = 0, Optional Sz = 0, Optional PartTp _
        As Byte = 1, Optional FirstImg As Byte = 0)
    On Error GoTo Err_Init

    Dim temp As Double
    Dim XFor As Long, YFor As Long
    Dim XLo As Long, YLo As Long
    Dim i As Integer
    
    temp = 360 / Num
    
    XFor = (Wdth / 2) / Steps
    YFor = (Hght / 2) / Steps
    XLo = Lft + (Wdth / 2)
    YLo = Tp + (Hght / 2)
    
    ReDim Parts(1 To Num)
    Dim K As Long
    For i = 1 To Num
        Set Parts(i) = New Particle
        Parts(i).InitParticle XLo, YLo, 0, Steps, 0, _
            XFor * Cos((i * temp) * PiRad), YFor * _
            Sin((i * temp) * PiRad), , , , , , MxSz, _
            Sz, SzTp, PartTp, FirstImg
        With Parts(i)
            For K = 250 To 0 Step -5
                .AddColor K, K, K
            Next
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "InitCircular", Err.Number, Err.Description
End Sub


Public Sub RandomCircular(XL As Long, YL As Long, _
        ZL As Long, Optional Num As Long = 50, _
        Optional XFor As Long = 5, Optional YFor _
        As Long = 5, Optional Tilt As Integer = 0)
    On Error GoTo Err_Init
    
    Dim temp As Double
    Dim t As Double
    Dim fx As Double
    Dim fy As Double
    Dim Center As Particle
    Dim i As Long, K As Long

    temp = 360 / Num
    t = Tilt * PiRad
    Set Center = New Particle
    Center.XLoc = XL
    Center.YLoc = YL
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = XL
            .YLoc = YL
            .ZLoc = ZL
            fx = XFor * Cos((i * temp) * PiRad)
            fy = YFor * Sin((i * temp) * PiRad)
            .XVel = fx
            .YVel = fy
            .PartType = 1
            .SizeType = 1
            .Life = 0
            .LifeSpan = XFor + YFor * 5
            For K = 250 To 0 Step -5
                .AddColor K, K, K
            Next
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomCircular", Err.Number, Err.Description
End Sub

Public Function MoveParticles() As Boolean
    On Error GoTo Err_Init
    Dim i As Long
    Dim SomeLeft As Boolean
    SomeLeft = False
    For i = 1 To UBound(Parts)
        Parts(i).MoveParticle
        If Parts(i).ParticleDead = False Then
            SomeLeft = True
        End If
    Next
    MoveParticles = SomeLeft
    Exit Function

Err_Init:
    HandleError CurrentModule, "MoveParticles", Err.Number, Err.Description
End Function

Public Sub DrawParticles()
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        Parts(i).DrawParticle
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawParticles", Err.Number, Err.Description
End Sub

Public Sub RemoveColor(Num As Long)
    On Error GoTo Err_Init
    Dim i As Long
    For i = Num To UBound(Parts)
        With Parts(i)
            .RemoveColor Num
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RemoveColor", Err.Number, Err.Description
End Sub

Public Sub RemoveLastColor()
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        With Parts(i)
            .RemoveLastColor
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RemoveLastColor", Err.Number, Err.Description
End Sub

Public Sub AddColor(Red As Long, Green As Long, Blue As Long)
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        With Parts(i)
            .AddColor Red, Green, Blue
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "AddColor", Err.Number, Err.Description
End Sub

Public Sub MoveAll(x As Long, Y As Long, Z As Long)
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        With Parts(i)
            .XLoc = .XLoc + x
            .YLoc = .YLoc + Y
            .ZLoc = .ZLoc + Z
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "MoveAll", Err.Number, Err.Description
End Sub

Public Sub Render()
    On Error GoTo Err_Init
    Dim NextTime As Long
    Do
        If QuitRender = True Then
            Exit Do
        End If
        If timeGetTime < NextTime Then
            DoEvents
        Else
            NextTime = timeGetTime + Speed
            If MoveParticles = True Then
                DrawParticles
            Else
                Exit Do
            End If
        End If
    Loop
    Exit Sub

Err_Init:
    HandleError CurrentModule, "Render", Err.Number, Err.Description
End Sub
