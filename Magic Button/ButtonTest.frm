VERSION 5.00
Object = "{C79217B8-04F5-4DBA-A319-C997A90CE9D3}#18.0#0"; "MagicButton.ocx"
Begin VB.Form ButtonTest 
   Caption         =   "Magic Button Test Form"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   Picture         =   "ButtonTest.frx":0000
   ScaleHeight     =   6270
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin MagicButton.MButton MButton3 
      Height          =   855
      Left            =   8400
      TabIndex        =   10
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      ShapeFormat     =   11
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   192
      Depth           =   55
      PictureMaskColor=   16711935
      AmbientLight    =   1.1
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1215
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2143
      ShapeFormat     =   0
      Caption         =   "&Click Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1485
      Left            =   3960
      Picture         =   "ButtonTest.frx":21BAF
      ScaleHeight     =   1425
      ScaleWidth      =   1845
      TabIndex        =   0
      Top             =   3960
      Width           =   1905
      Begin MagicButton.MButton MButton2 
         Height          =   1215
         Left            =   0
         TabIndex        =   8
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2143
         ShapeFormat     =   0
         Caption         =   "&Hiding a boat"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   14737632
         PictureMaskColor=   16711935
      End
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1575
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2778
      ShapeFormat     =   0
      Caption         =   "&Me Also"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1335
      Index           =   2
      Left            =   6480
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      ShapeFormat     =   0
      Caption         =   "&And Me!"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1095
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      ShapeFormat     =   0
      Caption         =   "&This Way"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1335
      Index           =   4
      Left            =   7320
      TabIndex        =   5
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2355
      ShapeFormat     =   0
      Caption         =   "S&wiming"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1095
      Index           =   5
      Left            =   7560
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1931
      ShapeFormat     =   0
      Caption         =   "&Surprise"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1695
      Index           =   6
      Left            =   1800
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2990
      ShapeFormat     =   0
      Caption         =   "This one is Craz&y"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
   Begin MagicButton.MButton MButton1 
      Height          =   1575
      Index           =   7
      Left            =   5280
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2778
      ShapeFormat     =   0
      Caption         =   "T&ry Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632319
      PictureMaskColor=   16711935
   End
End
Attribute VB_Name = "ButtonTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MButton1_Click(index As Integer)
    Select Case index
        Case 0
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = EllipseShape
            Set MButton1(index).Texture = LoadPicture(App.Path + "\images1.jpg")
            MButton1(index).BackgroundMode = TextureBG
            MButton1(index).Font.Name = "Arial"
            MButton1(index).Font.Size = 16
            MButton1(index).Font.Bold = True
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "OH YES"
            MButton1(index).AutoRedraw = True
        Case 1
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = Starshape
            MButton1(index).Depth = 25
            Set MButton1(index).Texture = LoadPicture(App.Path + "\images3.jpg")
            MButton1(index).BackgroundMode = TextureBG
            MButton1(index).Font.Bold = True
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "Magic"
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 10
            MButton1(index).CaptionRotation = 310
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 20
            MButton1(index).CaptionRotation = 225
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 30
            MButton1(index).CaptionRotation = 144
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 40
            MButton1(index).CaptionRotation = 177
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 50
            MButton1(index).CaptionRotation = 111
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 60
            MButton1(index).CaptionRotation = 44
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 70
            MButton1(index).CaptionRotation = 99
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeRotation = 80
            MButton1(index).CaptionRotation = 10
            MButton1(index).AutoRedraw = True
        Case 2
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = PolygoneShape
            MButton1(index).PolygonSides = 8
            Set MButton1(index).Texture = LoadPicture(App.Path + "\images2.jpg")
            MButton1(index).BackgroundMode = TextureBG
            MButton1(index).Font.Name = "Arial"
            MButton1(index).Font.Size = 8
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "It's Really Magic"
            Set MButton1(index).Picture = LoadPicture(App.Path + "\picture.bmp")
            MButton1(index).PictureMaskColor = vbWhite
            MButton1(index).PictureTransparency = True
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).PictureAlignment = BottomPicture
            DoEvents
            MButton1(index).PictureAlignment = leftPicture
            DoEvents
            MButton1(index).PictureAlignment = topPicture
            DoEvents
            MButton1(index).PictureAlignment = RightPicture
        Case 3
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = ArrowShape
            Set MButton1(index).Texture = LoadPicture(App.Path + "\images5.jpg")
            MButton1(index).BackgroundMode = TextureBG
            MButton1(index).Font.Name = "Arial"
            MButton1(index).Font.Size = 15
            MButton1(index).ForeColor = vbCyan
            MButton1(index).Caption = "Ok?"
            MButton1(index).AutoRedraw = True
        Case 4
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = Starshape
            Set MButton1(index).Texture = LoadPicture(App.Path + "\images4.jpg")
            MButton1(index).BackgroundMode = TextureBG
            MButton1(index).Font.Bold = True
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "Not Bad"
            MButton1(index).Depth = 25
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).Depth = 10
            DoEvents
            MButton1(index).Depth = 25
            DoEvents
            MButton1(index).Depth = 10
            DoEvents
            MButton1(index).Depth = 25
            DoEvents
            MButton1(index).ShapeFormat = PieShape
            DoEvents
            MButton1(index).ShapeFormat = ArrowHeadShape
            DoEvents
            MButton1(index).ShapeFormat = LShape
        Case 5
            MButton1(index).AutoRedraw = False
            MButton1(index).BackgroundMode = TransparentBG
            MButton1(index).ShapeFormat = ArrowHeadShape
            MButton1(index).Font.Bold = True
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "I am Here"
            MButton1(index).AutoRedraw = True
            DoEvents
            MButton1(index).BackColor = RGB(255, 200, 200)
            DoEvents
            MButton1(index).BackColor = RGB(128, 255, 200)
            DoEvents
            MButton1(index).BackColor = RGB(255, 200, 255)
            DoEvents
            MButton1(index).BackColor = RGB(255, 255, 200)
            DoEvents
            MButton1(index).BackColor = RGB(200, 255, 255)
        Case 6
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = MaskShape
            Set MButton1(index).ShapeMaskPicture = LoadPicture(App.Path + "\button mask.bmp")
            MButton1(index).BackgroundMode = RectangularGradientBG
            MButton1(index).BackColor = RGB(255, 64, 64)
            MButton1(index).BackendColor = RGB(255, 255, 64)
            MButton1(index).GradientAngle = 90
            MButton1(index).Font.Name = "Arial"
            MButton1(index).Font.Bold = True
            MButton1(index).Font.Size = 12
            MButton1(index).ForeColor = vbGreen
            MButton1(index).Caption = "CRAZY!!!"
            MButton1(index).AutoRedraw = True
        Case 7
            MButton1(index).AutoRedraw = False
            MButton1(index).ShapeFormat = PieShape
            MButton1(index).PieEndAngle = 180
            MButton1(index).PiestartAngle = 270
            MButton1(index).CaptionDeltay = -10
            MButton1(index).CaptionRotation = 45
            MButton1(index).BackgroundMode = RadialGradientBG
            MButton1(index).BackColor = RGB(195, 195, 255)
            MButton1(index).BackendColor = RGB(255, 0, 255)
            
            MButton1(index).Font.Name = "Arial"
            MButton1(index).Font.Bold = True
            MButton1(index).Font.Size = 12
            MButton1(index).ForeColor = vbYellow
            MButton1(index).Caption = "MAD"
            MButton1(index).AutoRedraw = True
        End Select

End Sub

Private Sub MButton2_Click()
    Dim i As Integer
    MButton2.AutoRedraw = False
    MButton2.ShapeFormat = SlantShape
    MButton2.BackgroundMode = TransparentBG
    MButton2.BackColor = RGB(255, 255, 255)
    MButton2.Caption = "no more"
    MButton2.AutoRedraw = True
    DoEvents
    For i = 30 To 360 Step 30
        MButton2.CaptionRotation = i
    Next i
End Sub

Private Sub MButton3_Click()
    End
End Sub
