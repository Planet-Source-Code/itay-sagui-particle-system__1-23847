VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tornado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Tornado"
'local variable(s) to hold property value(s)
Private mvarXLoc As Long
Private Parts() As Particle

Public Sub RandomTornado(XL As Long, YL As Long, _
        ZL As Long, MaxDist As Long, Optional _
        Num As Long = 100, Optional YVel As Double = -1)
    On Error GoTo Err_Init
    Dim i As Long, X As Long
    mvarXLoc = XL
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = XL
            .XAcl = (Rnd * 0.03) - 0.01
            .YLoc = YL
            .ZLoc = ZL
            .YVel = YVel
            .PartType = 1
            .SizeType = 1
            For X = 250 To 0 Step -5
                .AddColor 0, 0, 255
            Next
            .AddColor 0, 0, 255
            .LifeSpan = 500
            .Life = -i \ 2
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomTornado", Err.Number, Err.Description
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

Public Function MoveParticles() As Boolean
    On Error GoTo Err_Init
    Dim SomeLeft As Boolean
    Dim i As Long
    SomeLeft = False
    For i = 1 To UBound(Parts)
        If Parts(i).Life >= 0 Then
            If Parts(i).XLoc > mvarXLoc Then
                Parts(i).XAcl = Parts(i).XAcl - 0.0005
            Else
                Parts(i).XAcl = Parts(i).XAcl + 0.0005
            End If
        End If
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

Public Sub Render()
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

