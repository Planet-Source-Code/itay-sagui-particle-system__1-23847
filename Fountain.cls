VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fountain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Fountain"
Private Parts() As Particle
Private mvarRegenerate As Boolean
Private mvarWnd As Double
Private mvarGrav As Double
Private mvarForce As Integer
Private mvarXL As Long, mvarYL As Long, mvarZL As Long

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

Public Sub DrawParticles3D(hDC As Long)
    On Error GoTo Err_Init
    Dim i As Long
    For i = 1 To UBound(Parts)
        Parts(i).DrawParticle3D hDC
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawParticles3D", Err.Number, Err.Description
End Sub

Public Sub InitFountain(Lft As Long, Tp As Long, _
        Wdth As Long, Hght As Long, Optional Num _
        As Long = 50, Optional Steps As Long = 50, _
        Optional Wnd As Double = 0, Optional Regenerat _
        As Boolean = False, Optional MxSz As Integer = 3, _
        Optional SzTp As Byte = 0, Optional Sz = 0, _
        Optional PartTp As Byte = 1, Optional FirstImg _
        As Byte = 0)
    On Error GoTo Err_Init

    Dim temp As Double
    Dim XFor As Long, YFor As Long
    Dim XLo As Long, YLo As Long
    Dim i As Integer
    Dim Check As Long
    XLo = Lft + (Wdth / 2)
    YLo = Tp + Hght
    
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        Do
            XFor = (Rnd * 0.3) - 0.1
            Check = XFor * Steps
        Loop Until Abs(Check) < Wdth / 2
        Do
            YFor = (Hght * 5) / Steps
            Check = YFor * Steps / 2
        Loop Until Abs(Check) < Hght
        
        Parts(i).InitParticle XLo, YLo, 0, Steps, -i, _
            XFor, YFor, , , 0.1, , Wnd, MxSz, _
            Sz, SzTp, PartTp, FirstImg
        With Parts(i)
            Dim K As Long
            For K = 250 To 0 Step -5
                .AddColor K, K, K
            Next
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "InitFountain", Err.Number, Err.Description
End Sub


Public Sub RandomFountain(XL As Long, YL As Long, _
        ZL As Long, Optional Force As Integer = 5, _
        Optional Regenerate As Boolean = False, _
        Optional Num As Long = 50, Optional Wnd As _
        Double = 0, Optional Grav As Double = 0.1)
    On Error GoTo Err_Init
    Dim i As Long
    mvarXL = XL
    mvarYL = YL
    mvarZL = ZL
    mvarWnd = Wnd
    mvarGrav = Grav
    mvarForce = Force
    mvarRegenerate = Regenerate
    ReDim Parts(1 To Num)
    For i = 1 To Num
        Set Parts(i) = New Particle
        With Parts(i)
            .XLoc = XL
            .YLoc = YL
            .ZLoc = ZL
            .XVel = (Rnd * 0.3) - 0.1
            .YVel = -Rnd * Force
            .YAcl = Grav
            .Wind = Wnd
            .LifeSpan = 200
            .PartType = 1
            .SizeType = 0
            .Size = 0
            .Life = -i
        End With
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomFountain", Err.Number, Err.Description
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
            If mvarRegenerate Then
                With Parts(i)
                    .ParticleDead = False
                    .LifeSpan = 200
                    .XLoc = mvarXL
                    .YLoc = mvarYL
                    .ZLoc = mvarZL
                    .XVel = (Rnd * 0.3) - 0.1
                    .YVel = -Rnd * mvarForce
                    .YAcl = mvarGrav
                    .Wind = mvarWnd
                    .PartType = 1
                    .SizeType = 0
                    .Size = 0
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

