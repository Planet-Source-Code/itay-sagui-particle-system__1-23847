VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Particle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Const CurrentModule As String = "Particle"

'local variable(s) to hold property value(s)
Private mvarXLoc As Double
Private mvarYLoc As Double
Private mvarZLoc As Double
Private mvarXVel As Double
Private mvarYVel As Double
Private mvarZVel As Double
Private mvarXAcl As Double
Private mvarYAcl As Double
Private mvarZAcl As Double
Private mvarLifeSpan As Long
Private mvarLife As Long
Private mvarWind As Double
Private mvarParticleDead As Boolean
Private mvarSize As Double
Private mvarMaxSize As Integer
Private mvarSizeType As Long
Private mvarPartType As Long
Private mvarFirstImage As Long
Private Colors() As RGBCol


Public Property Let FirstImage(ByVal vData As Long)
    mvarFirstImage = vData
End Property


Public Property Get FirstImage() As Long
Attribute FirstImage.VB_UserMemId = 0
    FirstImage = mvarFirstImage
End Property


Public Property Let MaxSize(ByVal vData As Integer)
    mvarMaxSize = vData
End Property

Public Property Get MaxSize() As Integer
    MaxSize = mvarMaxSize
End Property

Public Function OutOfRange() As Boolean
    On Error GoTo Err_Init
    Dim x As Long, Y As Long
    x = LocationX2D  'Calculate these once and store the values to save
    Y = LocationY2D 'computing time.
    With BoundingRect
        If x > .Right Or x < .Left Or _
        Y < .Top Or Y > .Bottom Then
            OutOfRange = True
        Else
            OutOfRange = False
        End If
    End With
    Exit Function

Err_Init:
    HandleError CurrentModule, "OutOfRange", Err.Number, Err.Description
End Function


Public Sub InitParticle(ByVal XL As Long, ByVal YL As Long, _
        ByVal ZL As Long, LifeSp As Long, Optional Lf As Long = 0, _
        Optional ByVal Xv As Double = 0, Optional ByVal YV As _
        Double = 0, Optional ZV As Double = 0, Optional XA _
        As Double = 0, Optional YA As Double = 0, Optional _
        ZA As Double = 0, Optional Wnd As Double = 0, _
        Optional ByVal MxSz As Byte = 1, Optional ByVal Sz As _
        Integer = 0, Optional SzTp As Byte = 0, Optional _
        PartTp As Byte = 1, Optional FirstImg As Byte = 0)
    
    On Error GoTo Err_Init
    
    XLoc = XL: YLoc = YL: ZLoc = ZL
    LifeSpan = LifeSp: Life = Lf
    XVel = Xv: YVel = YV: ZVel = ZV
    XAcl = XA: YAcl = YA: ZAcl = ZA
    Wind = Wnd
    MaxSize = MxSz: Size = Sz: SizeType = SzTp
    PartType = PartTp: FirstImage = FirstImg
    Exit Sub

Err_Init:
    HandleError CurrentModule, "InitParticle", Err.Number, Err.Description
End Sub


Public Sub RandomParticle(ByVal XL As Long, ByVal YL As Long, _
        ByVal ZL As Long, Optional ByVal Wnd As Double = 0, _
        Optional ByVal Gravity As Double = 0.1, Optional _
        Force As Long = 9, Optional _
        ByVal MaxSz As Integer = 3, Optional PartTp As _
        Long = 1, Optional SizeTp As Long = 1, Optional _
        FirstImg As Long = 0)
    Dim x As Long
    On Error GoTo Err_Init
    ParticleDead = False
    MaxSize = MaxSz
    XLoc = XL
    YLoc = YL
    ZLoc = ZL
    XVel = (Rnd * Force) - (Force \ 2)
    YVel = (Rnd * Force) - (Force \ 2)
    ZVel = (Rnd * Force) - (Force \ 2)
    YAcl = Gravity
    Wind = Wnd
    Size = 0
    SizeType = SizeTp
    PartType = PartTp
    FirstImage = FirstImg
    ReDim Colors(0)
    For x = 250 To 0 Step -5
        AddColor x, 0, 0
    Next
    LifeSpan = Int(Rnd * 30) + 20
    Life = 0
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RandomParticle", Err.Number, Err.Description
End Sub


Public Sub MoveTo(x As Long, Y As Long, Z As Long)
    XLoc = x
    YLoc = Y
    ZLoc = Z
End Sub


Public Property Let ParticleDead(ByVal vData As Boolean)
    mvarParticleDead = vData
End Property


Public Property Get ParticleDead() As Boolean
    ParticleDead = mvarParticleDead
End Property


Public Property Let Size(ByVal vData As Double)
    mvarSize = vData
End Property


Public Property Get Size() As Double
    Size = mvarSize
End Property


Public Property Let PartType(ByVal vData As Long)
    mvarPartType = vData
End Property


Public Property Get PartType() As Long
    PartType = mvarPartType
End Property


Public Property Let SizeType(ByVal vData As Long)
    mvarSizeType = vData
End Property


Public Property Get SizeType() As Long
    SizeType = mvarSizeType
End Property


Public Function WindHorzForce() As Variant
Attribute WindHorzForce.VB_Description = "Calculates the Wind's horizontal force."
    If Wind <> 0 Then
        WindHorzForce = (Wind - XVel) * 0.1
    Else
        WindHorzForce = 0
    End If
End Function


Public Property Let Wind(ByVal vData As Double)
Attribute Wind.VB_Description = "Wind velocity."
    mvarWind = vData
End Property


Public Property Get Wind() As Double
    Wind = mvarWind
End Property


Public Property Let Life(ByVal vData As Long)
    mvarLife = vData
End Property


Public Property Get Life() As Long
    Life = mvarLife
End Property


Public Sub DrawPart(ByVal hDC As Long, XL As Long, YL As Long, Col As Long, Size As Long)
    On Error GoTo Err_Init
    Dim i As Long, j As Long
    For i = XL To XL + Size
        For j = YL To YL + Size
            SetPixel hDC, i, j, Col
        Next
    Next
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawPart", Err.Number, Err.Description
End Sub

Public Sub DrawParticle()
    On Error GoTo Err_Init
    Dim i As Long
    If Life > 0 Then
        Select Case PartType
        Case 1
            i = GetColor
            DrawPart TheHDC, XLoc, YLoc, RGB(Colors(i).Red, Colors(i).Green, Colors(i).Blue), Size
        Case 2
            i = FirstImage + Int(Size)
        End Select
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawParticle", Err.Number, Err.Description
End Sub

Public Sub DrawParticle3D(ByVal hDC As Long)
    On Error GoTo Err_Init
    Dim t As PointAPI
    Dim i As Long
    If Life > 0 Then
        i = GetColor
        SetPixel hDC, LocationX2D(), LocationY2D(), RGB(Colors(i).Red, Colors(i).Green, Colors(i).Blue)
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "DrawParticle3D", Err.Number, Err.Description
End Sub


Public Sub MoveParticle()
Attribute MoveParticle.VB_Description = "Move particle to its new location."
    On Error GoTo Err_Init
    If Life < LifeSpan And LifeSpan > 0 Then
        If OutOfRange = True Then
            'The fastest code is code that never executes.
            'Kill off non-visible particles so they don't slow down the rest.
            ParticleDead = True
        Else
            If Life >= 0 Then
                XLoc = XLoc + XVel
                XVel = XVel + XAcl + WindHorzForce
                YLoc = YLoc + YVel
                YVel = YVel + YAcl
                ZLoc = ZLoc + ZVel
                ZVel = ZVel + ZAcl
            End If
            Life = Life + 1
            Select Case mvarSizeType
            Case 1
                Size = (MaxSize * Life / LifeSpan)
            Case 2
                Size = MaxSize - (MaxSize * Life / LifeSpan)
            Case 3
                Size = (MaxSize * ZLoc / 256)
            End Select
        End If
    Else
        ParticleDead = True
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "MoveParticle", Err.Number, Err.Description
    Resume Next
End Sub


Public Property Let LifeSpan(ByVal vData As Long)
Attribute LifeSpan.VB_Description = "Life span of the particle. This will affect the color of the particle."
    mvarLifeSpan = vData
End Property


Public Property Get LifeSpan() As Long
    LifeSpan = mvarLifeSpan
End Property


Public Property Let ZAcl(ByVal vData As Double)
Attribute ZAcl.VB_Description = "The Z-Axis acceleration of the particle."
    mvarZAcl = vData
End Property


Public Property Get ZAcl() As Double
    ZAcl = mvarZAcl
End Property


Public Property Let YAcl(ByVal vData As Double)
Attribute YAcl.VB_Description = "The Y-Axis acceleration of the particle (The gravity)."
    mvarYAcl = vData
End Property


Public Property Get YAcl() As Double
    YAcl = mvarYAcl
End Property


Public Property Let XAcl(ByVal vData As Double)
Attribute XAcl.VB_Description = "The X-Axis acceleration of the particle."
    mvarXAcl = vData
End Property


Public Property Get XAcl() As Double
    XAcl = mvarXAcl
End Property


Public Sub SetGravity(Optional Force As Double = 9.8)
Attribute SetGravity.VB_Description = "Set the gravity that acts on the particle. The same as setting the YAcl"
    mvarYAcl = Force
End Sub

Public Property Let ZVel(ByVal vData As Double)
Attribute ZVel.VB_Description = "The Z-Axis velocity of the particle."
    mvarZVel = vData
End Property


Public Property Get ZVel() As Double
    ZVel = mvarZVel
End Property


Public Property Let YVel(ByVal vData As Double)
Attribute YVel.VB_Description = "The Y-Axis velocity of the particle."
    mvarYVel = vData
End Property


Public Property Get YVel() As Double
    YVel = mvarYVel
End Property


Public Property Let XVel(ByVal vData As Double)
Attribute XVel.VB_Description = "The X-Axis velocity of the particle."
    mvarXVel = vData
End Property


Public Property Get XVel() As Double
    XVel = mvarXVel
End Property


Public Property Let ZLoc(ByVal vData As Double)
Attribute ZLoc.VB_Description = "Location on the Z-Axis of the particle."
    mvarZLoc = vData
End Property


Public Property Get ZLoc() As Double
    ZLoc = mvarZLoc
End Property


Public Property Let YLoc(ByVal vData As Double)
Attribute YLoc.VB_Description = "Location on the Y-Axis of the particle."
    mvarYLoc = vData
End Property


Public Property Get YLoc() As Double
    YLoc = mvarYLoc
End Property


Public Property Let XLoc(ByVal vData As Double)
Attribute XLoc.VB_Description = "Location on the X-Axis of the particle."
    mvarXLoc = vData
End Property


Public Property Get XLoc() As Double
    XLoc = mvarXLoc
End Property


Private Sub Class_Initialize()
    ReDim Colors(0)
    ParticleDead = False
End Sub


Public Sub AddColor(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long)
Attribute AddColor.VB_Description = "Add a new color to the Colors array."
    On Error GoTo Err_Init
    Dim UB As Long
    UB = UBound(Colors) + 1
    ReDim Preserve Colors(UB)
    With Colors(UB)
        .Red = Red
        .Green = Green
        .Blue = Blue
    End With
    Exit Sub

Err_Init:
    HandleError CurrentModule, "AddColor", Err.Number, Err.Description
End Sub


Public Sub RemoveLastColor()
Attribute RemoveLastColor.VB_Description = "Remove the last color from the Colors array."
    ReDim Preserve Colors(UBound(Colors) - 1)
End Sub


Public Sub RemoveColor(ByVal Num As Single)
Attribute RemoveColor.VB_Description = "Remove a color from the Colors array."
    On Error GoTo Err_Init
    Dim i As Long
    For i = Num To UBound(Colors)
        With Colors(i)
            .Red = Colors(i + 1).Red
            .Green = Colors(i + 1).Green
            .Blue = Colors(i + 1).Blue
        End With
    Next
    ReDim Preserve Colors(UBound(Colors) - 1)
    Exit Sub

Err_Init:
    HandleError CurrentModule, "RemoveColor", Err.Number, Err.Description
End Sub

Public Function GetColor() As Long
Attribute GetColor.VB_Description = "Get the color number according to the LifeSpan."
    On Error GoTo Err_Init
    GetColor = Int(Int(Life * 100 / LifeSpan) * UBound(Colors) / 100)
    Exit Function

Err_Init:
    HandleError CurrentModule, "GetColor", Err.Number, Err.Description
End Function

Public Function LocationX2D() As Long
    On Error GoTo Err_Init
    LocationX2D = 200 * XLoc / (200 + ZLoc)
    Exit Function

Err_Init:
    HandleError CurrentModule, "LocationX2D", Err.Number, Err.Description
End Function

Public Function LocationY2D() As Long
    On Error GoTo Err_Init
    LocationY2D = 200 * YLoc / (200 + ZLoc)
    Exit Function

Err_Init:
    HandleError CurrentModule, "LocationY2D", Err.Number, Err.Description
End Function
