VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarEndP As Particle
Private mvarStartP As Particle

Public Sub MoveEnd(x As Variant, Y As Variant, Z As Variant)
    With EndP
        .XLoc = x
        .YLoc = Y
        .ZLoc = Z
    End With
End Sub

Public Sub MoveStart(x As Variant, Y As Variant, Z As Variant)
    With StartP
        .XLoc = x
        .YLoc = Y
        .ZLoc = Z
    End With
End Sub


Public Function Length() As Variant
    Length = ParticleFuncs.Distance3D(StartP, EndP)
End Function

Public Property Set StartP(ByVal vData As Particle)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.StartP = Form1
    Set mvarStartP = vData
End Property


Public Property Get StartP() As Particle
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.StartP
    Set StartP = mvarStartP
End Property



Public Property Set EndP(ByVal vData As Particle)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.EndP = Form1
    Set mvarEndP = vData
End Property


Public Property Get EndP() As Particle
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EndP
    Set EndP = mvarEndP
End Property



Private Sub Class_Initialize()
    Set mvarStartP = New Particle
    Set mvarEndP = New Particle
End Sub
