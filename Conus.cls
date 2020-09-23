VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Conus"
Private mvarDist As Double
Private mvarMaxDist As Double
Private Parts() As Particle

Public Sub RandomConus(XL As Long, YL As Long, _
        ZL As Long, Optional Num As Long = 50, _
        Optional Dist As Double = 0, Optional _
        Speed As Double = 0.1)
    On Error GoTo Err_Init
    Dim i As Long
    mvarDist = Dist
    ReDim Parts(Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .YLoc = YL
            .ZLoc = ZL
            .XLoc = XL - Dist + (Rnd * 2 * Dist)
            .ZLoc = ZL - Dist + (Rnd * 2 * Dist)
            .YVel = -(Rnd * 0.4) - Speed
            .PartType = 1
            .SizeType = 2
            .Size = 1
            .MaxSize = 3
            .Life = -(100 / Num) * i
            .LifeSpan = 2000
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomConus", Err.Number, Err.Description
End Sub

Public Function MoveParticles() As Boolean
    On Error GoTo Err_Init
    Dim SomeLeft As Boolean
    Dim i As Long
    SomeLeft = False
    For i = 1 To UBound(Parts)
        Parts(i).XVel = mvarDist * Cos(Parts(i).Life / 10)
        Parts(i).MoveParticle
        If Parts(i).ParticleDead = False Then
            SomeLeft = True
        End If
    Next
    mvarDist = mvarDist + 0.04
    MoveParticles = SomeLeft
    Exit Function

Err_Init:
    HandleError CurrentModule, "MoveParticles", Err.Number, Err.Description
End Function

Public Sub MoveParticles3D()
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        Parts(i).XVel = mvarMaxDist * Cos(Parts(i).Life / 10)
        Parts(i).YVel = mvarMaxDist * Cos(Parts(i).Life / 10)
        Parts(i).MoveParticle
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "MoveParticles3D", Err.Number, Err.Description
End Sub

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

Public Sub DrawParticles3D()
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        Parts(i).DrawParticle3D Form1.hDC
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawParticles3D", Err.Number, Err.Description
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

Public Sub Render(pic As PictureBox)
    On Error GoTo Err_Init
    Dim NextTime As Long
    Do
        'Check for user stop
        DoEvents
        If QuitRender = True Then
            Exit Do
        End If
        'Wait the minimum amount of time
        If timeGetTime < NextTime Then
            DoEvents
        Else
            'Draw the remaining live particles
            pic.Cls
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

