VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "EZEffects"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSettings 
      Caption         =   "Settings"
      Height          =   6780
      Left            =   6375
      TabIndex        =   2
      Top             =   0
      Width           =   3060
      Begin VB.CommandButton Command1 
         Caption         =   "Render2"
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   5640
         Width           =   1215
      End
      Begin VB.TextBox txtSP 
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1185
         Width           =   1695
      End
      Begin VB.CommandButton cmdRender 
         Caption         =   "Render"
         Height          =   495
         Left            =   975
         TabIndex        =   7
         Top             =   6090
         Width           =   1215
      End
      Begin VB.ComboBox cboPS 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   720
         List            =   "Form1.frx":0022
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   255
         Width           =   2220
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   660
         TabIndex        =   5
         Top             =   810
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   20
         SmallChange     =   10
         Max             =   100
         SelStart        =   20
         TickFrequency   =   10
         Value           =   20
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   660
         TabIndex        =   13
         Top             =   2280
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Min             =   -50
         Max             =   50
         TickFrequency   =   10
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   255
         Left            =   660
         TabIndex        =   14
         Top             =   2760
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   90
         TickFrequency   =   10
         Value           =   90
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   840
         TabIndex        =   11
         Top             =   1620
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         OLEDropMode     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   200
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin MSComctlLib.Slider Slider5 
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   1920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         OLEDropMode     =   1
         LargeChange     =   10
         Min             =   1
         Max             =   200
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Force:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label5 
         Caption         =   "Wind:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Num. Of Particles:"
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Starting Point:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   1245
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Speed:"
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Effect:"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   315
         Width           =   525
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6840
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      SimpleText      =   "Welcome to EZEffects, by Itay Sagui!"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14764
            Text            =   "Welcome to EZEffects, by Itay Sagui!"
            TextSave        =   "Welcome to EZEffects, by Itay Sagui!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1429
            MinWidth        =   1429
            Text            =   "000, 000"
            TextSave        =   "000, 000"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Height          =   6675
      Left            =   45
      ScaleHeight     =   441
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   6255
      Begin VB.Shape Shape 
         Height          =   375
         Index           =   0
         Left            =   600
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Shape Guide 
         Height          =   975
         Left            =   2760
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const CurrentModule As String = "Form1"
Dim StartX As Long, StartY As Long
Dim ChangeSP As Boolean
Public Act As String
Dim XX As Long
Dim YY As Long
Dim Effects()


Private Sub cboPS_Click()
    Dim i As Long, s As String
    Act = cboPS.Text
    i = cboPS.ListIndex
    s = LCase(cboPS.List(i))
    txtSP.Visible = False
    Slider2.Max = 200
    Slider2.TickFrequency = 10
    Slider3.Visible = False
    Slider3.Value = 0
    Slider5.Visible = False
    
    Label3.Visible = False
    Label5.Visible = False
    If s = "spiral" Then
        Slider2.Value = 50
        Slider1.Value = 0
        txtSP.Visible = True
        Label3.Visible = True
    ElseIf s = "explosion" Then
        Slider3.Visible = True
        Slider2.Value = 50
        Slider1.Value = 95
        txtSP.Visible = True
        Label3.Visible = True
        Label5.Visible = True
    ElseIf s = "circular" Then
        Slider2.Value = 40
        Slider1.Value = 75
        txtSP.Visible = True
        Label3.Visible = True
    ElseIf s = "cylinder" Then
        Slider2.Value = 200
        Slider1.Value = 80
        txtSP.Visible = True
        Label3.Visible = True
    ElseIf s = "conus" Then
        Slider2.Value = 200
        Slider1.Value = 80
        txtSP.Visible = True
        Label3.Visible = True
    ElseIf s = "firework" Then
        Slider5.Visible = True
        Slider5.Value = 3
        Slider3.Visible = True
        Slider2.Max = 20
        Slider2.TickFrequency = 2
        Slider2.Value = 5
        Slider1.Value = 85
        txtSP.Visible = True
        Label3.Visible = True
        Label5.Visible = True
    ElseIf s = "rain" Then
        Slider3.Visible = True
        Slider2.Max = 500
        Slider2.TickFrequency = 25
        Slider2.Value = 500
        Slider1.Value = 20
        Label5.Visible = True
    ElseIf s = "snow" Then
        Slider3.Visible = True
        Slider2.Max = 400
        Slider2.TickFrequency = 20
        Slider2.Value = 400
        Slider1.Value = 80
        Label5.Visible = True
    ElseIf s = "tornado" Then
        Slider2.Value = 200
        Slider1.Value = 95
        txtSP.Visible = True
        Label3.Visible = True
    ElseIf s = "fountain" Then
        Slider3.Visible = True
        Slider2.Value = 200
        Slider1.Value = 65
        txtSP.Visible = True
        Label3.Visible = True
        Label5.Visible = True
    End If
End Sub

Private Sub cmdRender_Click()
    On Error GoTo Err_Init
    Dim i As Long, s As String, c As Object, c2 As Object
    QuitRender = True
    DoEvents
    QuitRender = False
    i = cboPS.ListIndex
    s = LCase(cboPS.List(i))
    Picture1.Cls
    If s = "explosion" Then
        Set c = New Explosion
        c.RandomExplosion StartX, StartY, 200, Slider2.Value, Slider3.Value / 10, , Slider4.Value / 10, True
        c.Render
    ElseIf s = "circular" Then
        Set c = New Circular
        c.RandomCircular StartX, StartY, 200, Slider2.Value, 10, 2, 60
        c.Render
    ElseIf s = "cylinder" Then
        Set c = New Cylinder
        c.RandomCylinder StartX, StartY, 200, Slider2.Value
        c.Render Picture1
    ElseIf s = "firework" Then
        'Set up first stage 'fuse' firework.
        'When it dies, it will automatically generate an 'explosion' firework
        'at the ending x,y,z location.
        Set c = New Fuse
        c.RandomFuse StartX, StartY, 200, Slider5.Value, Slider2.Value, Slider3.Value / 10
        c.Render
    ElseIf s = "conus" Then
        Set c = New Conus
        c.RandomConus StartX, StartY, 200, Slider2.Value
        c.Render Picture1
    ElseIf s = "rain" Then
        Set c = New Rain
        With BoundingRect
            c.RandomRain .Left, .Right, .Top, .Bottom, 0, False, Slider2.Value, Slider3.Value / 10
        End With
        c.Render Picture1
    ElseIf s = "snow" Then
        Set c = New Snow
        With BoundingRect
            c.RandomSnow .Left, .Right, 100, Slider2.Value, Slider3.Value / 10
        End With
        c.Render Picture1
    ElseIf s = "spiral" Then
        Set c = New Spiral
        c.RandomSpiral StartX, StartY, 200, Slider2.Value, 5, 3
        c.Render Picture1
    ElseIf s = "tornado" Then
        Set c = New Tornado
        c.RandomTornado StartX, StartY, 200, 1, Slider2.Value
        c.Render
    ElseIf s = "fountain" Then
        Set c = New Fountain
        c.RandomFountain StartX, StartY, 100, , False, Slider2.Value, Slider3.Value / 10
        c.Render Picture1
    End If
    Exit Sub

Err_Init:
    HandleError CurrentModule, "cmdRender_Click", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub Command1_Click()
Dim j As Long
Dim i As Long
    For j = 1 To 100
        For i = LBound(Effects) + 1 To UBound(Effects)
            If IsObject(Effects(i)) Then
                Effects(i).MoveParticles
                Effects(i).DrawParticles
            End If
        Next
    Next
End Sub

Private Sub Form_Load()
    SetBoundingRect
    StartX = 200
    StartY = 200
    txtSP = StartX & ", " & StartY
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized And Visible = True Then
        Picture1.Move 0, 7, ScaleWidth - fraSettings.Width, ScaleHeight - StatusBar1.Height - 7
        fraSettings.Move ScaleWidth - fraSettings.Width, 0, fraSettings.Width, ScaleHeight - StatusBar1.Height
        cmdRender.Move ScaleX(((fraSettings.Width - ScaleX(cmdRender.Width, vbTwips, vbPixels)) / 2), vbPixels, vbTwips), ScaleY((fraSettings.Height - ScaleY(cmdRender.Height, vbTwips, vbPixels) - 10), vbPixels, vbTwips)
    End If
    SetBoundingRect
End Sub

Private Sub SetBoundingRect()
    With BoundingRect
        .Left = 0
        .Top = 0
        .Right = Picture1.ScaleWidth
        .Bottom = Picture1.ScaleHeight
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    QuitRender = True
End Sub

Private Sub mnuFileExit_Click()
    ShutDown
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    XX = X
    YY = Y
    Select Case Act
    Case "Explosion": Guide.Shape = 3
    Case "Circular": Guide.Shape = 2
    Case "Spiral": Guide.Shape = 2
    Case "Cylinder": Guide.Shape = 0
    Case Else: Guide.Shape = 4
    End Select
    Guide.Left = IIf(X > XX, XX, X)
    Guide.Top = IIf(Y > YY, YY, Y)
    Guide.Width = Abs(X - XX)
    Guide.Height = Abs(Y - YY)
    Guide.Visible = True
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar1.Panels(2).Text = Format(X, "000") & ", " & Format(Y, "000")
    Guide.Left = IIf(X > XX, XX, X)
    Guide.Top = IIf(Y > YY, YY, Y)
    Guide.Width = Abs(X - XX)
    Guide.Height = Abs(Y - YY)
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'StartX = X
    'StartY = Y
    'txtSP = StartX & ", " & StartY

Dim x1 As Long, y1 As Long
Dim x2 As Long, y2 As Long
Dim temp As Long
    
    Guide.Visible = False
    Select Case Act
    Case "Delete"
        temp = checkShape(X, Y)
        If temp <> 0 Then
            Unload Shape(temp)
            ShapeNum = ShapeNum - 1
        End If
    
    Case Else
        ShapeNum = ShapeNum + 1
        Load Shape(ShapeNum)
        ReDim Preserve Effects(ShapeNum)
        With Shape(ShapeNum)
            .Shape = Guide.Shape
            .Left = Guide.Left
            .Top = Guide.Top
            .Width = Guide.Width
            .Height = Guide.Height
            .Tag = Act
            .Visible = True
            If Act = "Explosion" Then
                Set Effects(ShapeNum) = New Explosion
                Effects(ShapeNum).RandomExplosion .Left + (.Width / 2), .Top + (.Height / 2), 200, Slider2.Value, Slider3.Value / 10, , Slider4.Value / 10, True
            End If
        End With
    End Select
End Sub

Public Sub Slider1_Change()
    Speed = Slider1.Max - Slider1.Value
End Sub

