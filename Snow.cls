VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Snow"
Private Parts() As Particle

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

Public Sub RandomSnow(XLs As Long, XLe As Long, _
        ZL As Long, Optional Num As Long = 500, _
        Optional Wnd As Double = 0)
    On Error GoTo Err_Init
    Dim i As Long, X As Long
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = (Rnd * (XLe - XLs)) + XLs
            .YLoc = 0
            .YVel = 1
            .XVel = Wnd
            .PartType = 1
            .SizeType = 0
            .Size = 1
            For X = 250 To 0 Step -5
                .AddColor 0, 0, 255
            Next
            .LifeSpan = 500
            .Life = -i \ 4
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomSnow", Err.Number, Err.Description
End Sub

Public Function MoveParticles() As Boolean
    On Error GoTo Err_Init
    Dim SomeLeft As Boolean
    Dim i As Long
    SomeLeft = False
    For i = 1 To UBound(Parts)
        Parts(i).XVel = (Rnd * 2) - 1 'Parts(i).XVel + (Rnd * 0.2) - 0.1
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
    Dim i As Long
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


