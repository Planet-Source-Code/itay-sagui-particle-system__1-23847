VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Rain"
'local variable(s) to hold property value(s)
Private Parts() As Particle
Private mvarRegenerate As Boolean
Private mvarXLs As Long, mvarXLe As Long
Private mvarZLs As Long, mvarZLe As Long
Private mvarYL As Long
Private mvarWind As Double, mvarGrav As Double

Public Property Let Regenerate(ByVal vData As Boolean)
    mvarRegenerate = vData
End Property

Public Property Get Regenerate() As Boolean
Attribute Regenerate.VB_UserMemId = 0
    Regenerate = mvarRegenerate
End Property

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

Public Sub RandomRain(XLs As Long, XLe As Long, _
        ZLs As Long, ZLe As Long, Optional YL As _
        Long = 0, Optional Regenerate As Boolean = False, _
        Optional Num As Long = 500, Optional _
        Wnd As Double = 0, Optional Grav As Double = 0.1)
    On Error GoTo Err_Init
    Dim i As Long, x As Long
    
    mvarXLs = XLs
    mvarXLe = XLe
    mvarZLs = ZLs
    mvarZLe = ZLe
    mvarYL = YL
    mvarWind = Wnd
    mvarGrav = Grav
    mvarRegenerate = Regenerate
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = (Rnd * (XLe - XLs)) + XLs
            .YLoc = YL
            .ZLoc = (Rnd * (ZLe - ZLs)) + ZLs
            .YVel = 2
            .Wind = Wnd
            .YAcl = Grav
            .PartType = 1
            .SizeType = 0
            .Size = 0
            For x = 250 To 0 Step -5
                .AddColor 0, 0, 255
            Next
            .LifeSpan = 50
            .Life = -i
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomRain", Err.Number, Err.Description
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
        Else
            If Regenerate Then
                With Parts(i)
                    .ParticleDead = False
                    .XLoc = (Rnd * (mvarXLe - mvarXLs)) + mvarXLs
                    .YLoc = mvarYL
                    .ZLoc = (Rnd * (mvarZLe - mvarZLs)) + mvarZLs
                    .YVel = 2
                    .Wind = mvarWind
                    .YAcl = mvarGrav
                    .PartType = 1
                    .SizeType = 0
                    .Size = 0
                    .LifeSpan = 50
                    .Life = 0
                End With
            End If
        End If
    Next
    MoveParticles = SomeLeft
    Exit Function

Err_Init:
    HandleError CurrentModule, "MoveParticles", Err.Number, Err.Description
End Function

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

