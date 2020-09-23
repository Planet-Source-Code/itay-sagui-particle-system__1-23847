Attribute VB_Name = "Module2"
Public Function AddVectors(V1 As Vector, V2 As Vector) As Vector
    With V1.StartP
        AddVectors.StartP.XLoc = .XLoc + V2.StartP.XLoc
        AddVectors.StartP.XLoc = .XLoc + V2.StartP.XLoc
        AddVectors.StartP.XLoc = .XLoc + V2.StartP.XLoc
    End With
End Function

Public Function SubstractVectors(V1 As Vector, V2 As Vector) As Vector
    With V1.StartP
        SubstractVectors.StartP.XLoc = .XLoc - V2.StartP.XLoc
        SubstractVectors.StartP.XLoc = .XLoc - V2.StartP.XLoc
        SubstractVectors.StartP.XLoc = .XLoc - V2.StartP.XLoc
    End With
End Function

Public Sub ProductVectors(V1 As Vector, V2 As Vector, ByRef V3 As Vector)
    V3.MoveStart V1.StartP.XLoc, V1.StartP.YLoc, V1.StartP.ZLoc
    V3.EndP.XLoc = V1.StartP.XLoc + (V1.EndP.YLoc * V2.EndP.ZLoc - V2.EndP.YLoc * V1.EndP.ZLoc)
    V3.EndP.YLoc = V1.StartP.XLoc + (V1.EndP.ZLoc * V2.EndP.XLoc - V2.EndP.ZLoc * V1.EndP.XLoc)
    V3.EndP.ZLoc = V1.StartP.XLoc + (V1.EndP.XLoc * V2.EndP.YLoc - V2.EndP.XLoc * V1.EndP.YLoc)
End Sub

Public Function DotVectors(V1 As Vector, V2 As Vector) As Variant
    Ax = DistanceX(V1.StartP, V1.EndP)
    Ay = DistanceY(V1.StartP, V1.EndP)
    Az = DistanceZ(V1.StartP, V1.EndP)
    Bx = DistanceX(V2.StartP, V2.EndP)
    By = DistanceY(V2.StartP, V2.EndP)
    Bz = DistanceZ(V2.StartP, V2.EndP)

    DotVectors = Ax * Bx + Ay * By + Az * Bz
 
End Function
