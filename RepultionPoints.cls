VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RepultionPoints"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private RepPs() As AttracPoint
Public RPnum As Long

Public Sub RemovePoint(Num As Long)
    For i = Num To UBound(RepPs)
        With RepPs(i)
            .XLoc = RepPs(i + 1).XLoc
            .YLoc = RepPs(i + 1).YLoc
            .ZLoc = RepPs(i + 1).ZLoc
            .Mass = RepPs(i + 1).Mass
        End With
    Next
    ReDim Preserve RepPs(UBound(RepPs) - 1)
    RPnum = UBound(RepPs)
End Sub

Public Sub RemoveLastPoint()
    ReDim Preserve RepPs(UBound(RepPs) - 1)
    RPnum = UBound(RepPs)
End Sub

Public Sub AddPoint(XLoc As Long, YLoc As Long, ZLoc As Long, Mass As Long)
    UB = UBound(RepPs) + 1
    ReDim Preserve RepPs(UB)
    With RepPs(UB)
        .XLoc = XLoc
        .YLoc = YLoc
        .ZLoc = ZLoc
        .Mass = Mass
    End With
    RPnum = UBound(RepPs)
End Sub

Public Sub Repulse(Part As Particle, ByVal Rep As Single)
Dim Xv As Double
Dim Xy As Double
Dim Distance As Double
    
    Xv = RepPs(Rep).XLoc - Part.XLoc
    Yv = RepPs(Rep).YLoc - Part.YLoc
    Distance = Sqr((Xv ^ 2) + (Yv ^ 2))
    Xv = Xv / Distance
    Yv = Yv / Distance
    Force = (1 * RepPs(Rep).Mass) / (Distance ^ 2)
    Xv = Xv * Force / (RepPs(Rep).Mass + 1)
    Yv = Yv * Force / (RepPs(Rep).Mass + 1)
    Part.XVel = Part.XVel - Xv
    Part.YVel = Part.YVel - Yv
End Sub

Public Sub MovePoint(ByVal Rep As Long, XL As Long, YL As Long, ZL As Long)
    With RepPs(Rep)
        .XLoc = XL
        .YLoc = YL
        .ZLoc = ZL
    End With
End Sub

Private Sub Class_Initialize()
    ReDim RepPs(0)
End Sub
