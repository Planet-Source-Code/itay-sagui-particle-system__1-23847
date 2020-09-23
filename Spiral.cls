VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spiral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Spiral"
Private Parts() As Particle

Public Sub RandomSpiral(XL As Long, YL As Long, _
        ZL As Long, Optional Num As Long = 50, _
        Optional XFor As Double = 5, Optional YFor As Double = 5)
    On Error GoTo Err_Init
    Dim i As Long, k As Long
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = XL
            .YLoc = YL
            .ZLoc = ZL
            .XVel = XFor * Cos((i * 360 / Num) * 3.14 / 180)
            .YVel = YFor * Sin((i * 360 / Num) * 3.14 / 180)
            .Life = -Int(i / (XFor + YFor))
            .LifeSpan = XFor + YFor * 5
            .PartType = 1
            .SizeType = 1
            For k = 250 To 0 Step -5
                .AddColor k, k, k
            Next
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomSpiral", Err.Number, Err.Description
End Sub

Public Function MoveParticles() As Boolean
    On Error GoTo Err_Init
    Dim SomeLeft As Boolean
    Dim i As Long
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

Public Sub MoveAll(X As Long, Y As Long, Z As Long)
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        With Parts(i)
            .XLoc = .XLoc + X
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

