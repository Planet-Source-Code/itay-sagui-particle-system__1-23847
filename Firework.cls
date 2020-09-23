VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Firework"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarNum As Long 'local copy
Private Fuse As Particle
Private Explose As Explosion

Public Sub RandomFirework(XL As Long, YL As Long, _
        ZL As Long, Optional Num As Long = 50, _
        Optional Wnd As Double = 0, Optional Grav As _
        Double = 0.1)
    Set Fuse = New Particle
    Fuse.XLoc = XL
    Fuse.YLoc = YL
    Fuse.ZLoc = ZL
    Fuse.XVel = (Rnd * 6) - 4
    Fuse.YVel = -(Rnd * 4) - 4
    Fuse.ZVel = (Rnd * 6) - 4
    Fuse.Wind = Wnd
    Fuse.YAcl = Grav
    Fuse.PartType = 1
    Fuse.SizeType = 0
    Fuse.Size = 1
    Fuse.LifeSpan = 50
    mvarNum = Num
End Sub

Public Sub MoveParticle()
    If Fuse.ParticleDead Then
        Explose.MoveParticles
    Else
        Fuse.MoveParticle
        If Fuse.ParticleDead Then
            Set Explose = New Explosion
            Explose.RandomExplosion Fuse.XLoc, Fuse.YLoc, _
                Fuse.ZLoc, mvarNum, Fuse.Wind
        End If
    End If
End Sub

Public Sub DrawParticle(hDC As Long)
    If Fuse.ParticleDead Then
        Explose.DrawParticles hDC
    Else
        Fuse.DrawParticle hDC
    End If
End Sub

Public Sub DrawParticle3D(hDC As Long)
    If Fuse.ParticleDead Then
        Explose.DrawParticles3D hDC
    Else
        Fuse.DrawParticle3D hDC
    End If
End Sub

Public Sub RemoveColor(Num As Byte)
    For i = Num To UBound(Parts)
        With Parts(i)
            .RemoveColor Num
        End With
    Next
End Sub

Public Sub RemoveLastColor()
    For i = 1 To UBound(Parts)
        With Parts(i)
            .RemoveLastColor
        End With
    Next
End Sub

Public Sub AddColor(Red As Byte, Green As Byte, Blue As Byte)
    For i = 1 To UBound(Parts)
        With Parts(i)
            .AddColor Red, Green, Blue
        End With
    Next
End Sub

Public Sub MoveAll(x As Long, y As Long, Z As Long)
    For i = 1 To UBound(Parts)
        With Parts(i)
            .XLoc = .XLoc + x
            .YLoc = .YLoc + y
            .ZLoc = .ZLoc + Z
        End With
    Next
End Sub

