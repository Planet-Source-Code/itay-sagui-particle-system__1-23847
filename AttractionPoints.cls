VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AttractionPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private AttPs() As AttracPoint
Public APnum As Long

Public Sub RemovePoint(Num As Long)
    For i = Num To UBound(AttPs)
        With AttPs(i)
            .XLoc = AttPs(i + 1).XLoc
            .YLoc = AttPs(i + 1).YLoc
            .ZLoc = AttPs(i + 1).ZLoc
            .Mass = AttPs(i + 1).Mass
        End With
    Next
    ReDim Preserve AttPs(UBound(AttPs) - 1)
    APnum = UBound(AttPs)
End Sub

Public Sub RemoveLastPoint()
    ReDim Preserve AttPs(UBound(AttPs) - 1)
    APnum = UBound(AttPs)
End Sub

Public Sub AddPoint(XLoc As Long, YLoc As Long, ZLoc As Long, Mass As Long)
    UB = UBound(AttPs) + 1
    ReDim Preserve AttPs(UB)
    With AttPs(UB)
        .XLoc = XLoc
        .YLoc = YLoc
        .ZLoc = ZLoc
        .Mass = Mass
    End With
    APnum = UBound(AttPs)
End Sub

Public Sub Attract(Part As Particle, ByVal Att As Long)
Dim Xv As Double
Dim Xy As Double
Dim Distance As Double
    
    Xv = AttPs(Att).XLoc - Part.XLoc
    Yv = AttPs(Att).YLoc - Part.YLoc
    Distance = Sqr((Xv ^ 2) + (Yv ^ 2))
    Xv = Xv / Distance
    Yv = Yv / Distance
    Force = (1 * AttPs(Att).Mass) / (Distance ^ 2)
    Xv = Xv * Force / (AttPs(Att).Mass + 1)
    Yv = Yv * Force / (AttPs(Att).Mass + 1)
    Part.XVel = Part.XVel + Xv
    Part.YVel = Part.YVel + Yv
End Sub

Public Sub MovePoint(ByVal Att As Long, XL As Long, YL As Long, ZL As Long)
    With AttPs(Att)
        .XLoc = XL
        .YLoc = YL
        .ZLoc = ZL
    End With
End Sub

Private Sub Class_Initialize()
    ReDim AttPs(0)
End Sub
