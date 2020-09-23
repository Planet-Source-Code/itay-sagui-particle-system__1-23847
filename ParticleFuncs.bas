Attribute VB_Name = "ParticleFuncs"
Public Function Distance2D(Part1 As Particle, Part2 As Particle) As Long
    Distance2D = Sqr((Part2.XLoc - Part1.XLoc) ^ 2 + (Part2.YLoc - Part1.YLoc) ^ 2)
End Function

Public Function Distance3D(Part1 As Particle, Part2 As Particle) As Long
    Distance3D = Sqr(DistanceX(Part1, Part2) ^ 2 + _
                      DistanceY(Part1, Part2) ^ 2 + _
                      DistanceZ(Part1, Part2) ^ 2)
End Function

Public Function DistanceX(Part1 As Particle, Part2 As Particle) As Long
    DistanceX = Part2.XLoc - Part1.XLoc
End Function
Public Function DistanceY(Part1 As Particle, Part2 As Particle) As Long
    DistanceY = Part2.YLoc - Part1.YLoc
End Function
Public Function DistanceZ(Part1 As Particle, Part2 As Particle) As Long
    DistanceZ = Part2.ZLoc - Part1.ZLoc
End Function

