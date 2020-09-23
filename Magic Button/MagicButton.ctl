VERSION 5.00
Begin VB.UserControl MButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ForeColor       =   &H00FFFFFF&
   PropertyPages   =   "MagicButton.ctx":0000
   ScaleHeight     =   68
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   154
   ToolboxBitmap   =   "MagicButton.ctx":000F
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   480
   End
End
Attribute VB_Name = "MButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===========================================================================
' Magic Button ActiveX by frk (frk@paris.com)
'
' Feel free to use this control in your non-commercial application
' if you keep the credits to the author
'
' Thanks to Carles P.V. for his gradient algorithm
' and to Andrew (Booda) Stickney for his wormhole gradient
'
' If you want to use it in a commercial application, contact me first
'
' To Do List
'
'
'===========================================================================

'===========================================================================
'
' Windows API and Type Delarations
'
'===========================================================================

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyfillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointApi) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As Any, ByVal nCount As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointApi) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDCDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hDCSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long

' Constants used for the font creation
Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10


Private Const LF_FACESIZE = 32
Private Const FW_BOLD = 700
Private Const FW_NORMAL = 400
Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const CLIP_LH_ANGLES = 16
Private Const ANTIALIASED_QUALITY = 4
Private Const DEFAULT_PITCH = 0
Private Const FF_DONTCARE = 0

' constant used to merge region when calculting transparency
Private Const RGN_OR = 2

' constant used to creatge a solid Pen
Private Const PS_SOLID = 0

' constant used when calling bitblt function
Private Const SRCCOPY = &HCC0020

' Constant to set the font drawing mode
Private Const TRANSPARENT = 1

' Structure definition used when creating font
Private Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName As String * LF_FACESIZE
End Type

' Structure used when addressing a specific point
Private Type PointApi
   X As Long
   Y As Long
End Type

' Rectangle structure
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' Bitmap Info Header Structure
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

' Bitmap Structure, used to create the images in memory
Private Type BITMAPINFO
    Header As BITMAPINFOHEADER
    Bytes() As Byte
End Type


'===========================================================================
'
' Enum and Variables Declaration
'
'===========================================================================

' The different possible shapes
Public Enum ShapeFormat_Enum
    RectangleShape
    RoundRectangleShape
    EllipseShape
    PolygoneShape
    Starshape
    ArrowShape
    MaskShape
    PieShape
    LShape
    SlantShape
    ArrowHeadShape
    CrossShape
End Enum

' The caption alignment possibilities
Public Enum CaptionAlignment_Enum
    RightCaption
    LeftCaption
    CenterCaption
End Enum

'The Different BackgroundMode
Public Enum BackgroundMode_Enum
    SolidBG
    TransparentBG
    TextureBG
    LinearGradientBG
    RadialGradientBG
    RectangularGradientBG
End Enum

' The 2 possible emboss ways
Private Enum Emboss_Enum
    EmbossUp
    EmbossDown
End Enum

' All the button States
Private Enum ButtonState_Enum
    ButtonisUp
    ButtonIsDown
    MouseIsOver
    ButtonIsDisabled
End Enum

' the different picture alignment
Public Enum PictureAligment_Enum
    CenterPicture
    RightPicture
    leftPicture
    topPicture
    BottomPicture
End Enum

' Private Type, used to draw the arrow
Private Type EllipticCOORD
    RX As Single
    RY As Single
    Angle As Single
End Type

'Private Type, used to store parts of the drawings
'as if few cases we use only one color, we don't need to have
'in memory bytes we never use
Private Type LIGHTBITMAP
    Bytes() As Byte
End Type

' Few constants
Private Const DefaultRoundCornerWidth = 40
Private Const DefaultTopBorder = 18
Private Const DefaultLeftBorder = 18
Private Const DefaultRightBorder = 15
Private Const DefaultBottomBorder = 15
Private Const Thickness = 3
Private Const DeltaShift = 1

'Global variables for all the properties
Private CurrentRoundCornerWidth As Integer
Private CurrentShapeFormat As ShapeFormat_Enum
Private CurrentTexture As StdPicture
Private CurrentPicture As StdPicture
Private CurrentCaption As String
Private CurrentAlignment As CaptionAlignment_Enum
Private CurrentForeColor As Long
Private CurrentFont As Font
Private CurrentBackColor As Long
Private CurrentPolygonSides As Integer
Private CurrentShapeRotation As Integer
Private CurrentEnabled As Boolean
Private CurrentDepth As Integer
Private CurrentAutoRedraw As Boolean
Private CurrentPictureAlignment As PictureAligment_Enum
Private CurrentPictureMaskColor As Long
Private CurrentPictureTransparency As Boolean
Private CurrentPictureDeltaX As Integer
Private CurrentPictureDeltaY As Integer
Private CurrentShapeMaskPicture As StdPicture
Private CurrentCaptionRotation As Integer
Private CurrentPieStartAngle As Integer
Private CurrentPieEndAngle As Integer
Private CurrentCaptionDeltaX As Integer
Private CurrentCaptionDeltaY As Integer
Private CurrentSlantStart As Integer
Private CurrentSlantEnd As Integer
Private CurrentAmbientLight As Single
Private CurrentGradientAngle As Integer
Private CurrentBackEndColor As Long
Private CurrentBackgroundMode As BackgroundMode_Enum


' Global Variables
Private DesignWidth As Long 'Used to store the drawing with
Private DesignHeight As Long 'Used to store the drawing height
Private ButtonState As ButtonState_Enum ' How is the button?
Private PI As Double 'Guess What?
Private CurrentObjectPosition As PointApi 'Trick used when button transparency is set to true
Private ArrowPoints(7) As EllipticCOORD 'Used for the arrow

' All the buffers containing the images in memory
Private DesignBuffer As BITMAPINFO
Private TextureBuffer As BITMAPINFO
Private ButtonUpBuffer As BITMAPINFO
Private ButtonDownBuffer As BITMAPINFO
Private ButtonOverBuffer As BITMAPINFO
Private ButtonDisabledBuffer As BITMAPINFO

'and the buffers containing only one color
Private BlurBuffer As LIGHTBITMAP
Private EmbossBlurBuffer As LIGHTBITMAP
Private EmbossBuffer As LIGHTBITMAP

'===========================================================================
'
' Event manage by the control
'
'===========================================================================

Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    UserControl_Click
End Sub

Private Sub UserControl_Click()
    If ButtonState = ButtonIsDisabled Then Exit Sub
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    If ButtonState = ButtonIsDisabled Then Exit Sub
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ButtonState = ButtonIsDisabled Then Exit Sub
    RaiseEvent MouseDown(Button, Shift, ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
    'When mouse in down the button state changes
    If Button = 1 Then
        ButtonState = ButtonIsDown
        DisplayButton ButtonState
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ButtonState = ButtonIsDisabled Then Exit Sub
    Dim TestX As Single
    Dim TestY As Single
    Dim CursorPosition As PointApi
    
    RaiseEvent MouseMove(Button, Shift, ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
    
    ' Get the cursor position and convet it to usercontrol relative coordinates
    GetCursorPos CursorPosition
    ScreenToClient UserControl.hwnd, CursorPosition

    TestX = CursorPosition.X
    TestY = CursorPosition.Y
    
    'The SetCapture function sets the mouse capture to the specified window belonging to the current thread. SetCapture captures mouse input either when the mouse is over the capturing window, or when the mouse button was pressed while the mouse was over the capturing window and the button is still down. Only one window at a time can capture the mouse.
    If ButtonState = ButtonisUp Then
        If TestX >= 0 And TestY >= 0 And TestX <= UserControl.ScaleWidth And TestY <= UserControl.ScaleHeight Then
            If DesignBuffer.Bytes(0, TestX, TestY) > 0 Or DesignBuffer.Bytes(1, TestX, TestY) > 0 Or DesignBuffer.Bytes(1, TestX, TestY) > 0 Then
                SetCapture UserControl.hwnd
            End If
        End If
    End If
   

    ' if the point is ouside the rectangle we release the mouse capture and set the button state to up
    If TestX < 0 Or TestY < 0 Or TestX > UserControl.ScaleWidth Or TestY > UserControl.ScaleHeight Then
        ReleaseCapture
        ButtonState = ButtonisUp
        DisplayButton ButtonState
    Else
        'Check if the mouse is over a visible part of the design
        If DesignBuffer.Bytes(0, TestX, TestY) = 0 And DesignBuffer.Bytes(1, TestX, TestY) = 0 And DesignBuffer.Bytes(1, TestX, TestY) = 0 Then
            If ButtonState <> ButtonisUp Then
                ReleaseCapture
                ButtonState = ButtonisUp
                DisplayButton ButtonState
            End If
        Else
            ' else if the button state is up we set it to over (else it is already over or pressed)
            If ButtonState = ButtonisUp Then
                ButtonState = MouseIsOver
                DisplayButton ButtonState
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If ButtonState = ButtonIsDisabled Then Exit Sub
    RaiseEvent MouseUp(Button, Shift, ScaleX(X, vbTwips, vbContainerPosition), ScaleY(Y, vbTwips, vbContainerPosition))
    'Change state back to up when releasing the mouse button
    If Button = 1 Then
        ButtonState = ButtonisUp
        DisplayButton ButtonState
    End If
End Sub

'===========================================================================
'
' Events sent by the parent form / design
'
'===========================================================================

' used to show the about box when called from the properties
Public Sub showAbout()
Attribute showAbout.VB_UserMemId = -552
   frmAbout.Show vbModal
End Sub

'This is a dirty trick from  [crenaud76 on www.vbfrance.com] thanks to him
Private Sub Timer1_Timer()
    Dim NewPosition As PointApi
    If Ambient.UserMode = False And CurrentBackgroundMode = TransparentBG Then
        NewPosition.X = ScaleX(UserControl.Extender.Left, Parent.ScaleMode, vbPixels)
        NewPosition.Y = ScaleY(UserControl.Extender.Top, Parent.ScaleMode, vbPixels)
        If NewPosition.X <> CurrentObjectPosition.X Or NewPosition.Y <> CurrentObjectPosition.Y Then
            Refresh
            CurrentObjectPosition = NewPosition
        End If
    Else
        Timer1.Enabled = False
    End If
End Sub

'Called when the control is intialized
Private Sub UserControl_Initialize()
    ButtonState = ButtonisUp
    PI = 4 * Atn(1)
    Set CurrentTexture = Nothing
    Set CurrentPicture = Nothing
    CurrentRoundCornerWidth = DefaultRoundCornerWidth
    CurrentShapeFormat = RoundRectangleShape
    CurrentCaption = UserControl.Name
    Set CurrentFont = UserControl.Font
    CurrentAlignment = CenterCaption
    CurrentForeColor = vbBlack
    CurrentBackColor = &HFF8080
    CurrentPolygonSides = 5
    CurrentShapeRotation = 0
    CurrentEnabled = True
    CurrentDepth = 50
    CurrentPictureAlignment = CenterPicture
    CurrentPictureMaskColor = vbMagenta
    CurrentAutoRedraw = False
    CurrentPictureTransparency = False
    CurrentBackgroundMode = SolidBG
    CurrentPieStartAngle = 0
    CurrentPieEndAngle = 270
    CurrentCaptionDeltaX = 0
    CurrentCaptionDeltaY = 0
    CurrentSlantStart = 10
    CurrentSlantEnd = 10
    CurrentPictureDeltaX = 0
    CurrentPictureDeltaY = 0
    CurrentAmbientLight = 1
    Set CurrentShapeMaskPicture = Nothing
    CurrentCaptionRotation = 0
    CurrentGradientAngle = 0
    CurrentBackEndColor = vbWhite
    Timer1.Enabled = False
End Sub

' called when initializing the properties
' set the autoredraw to false to avoid multiple refresh
Private Sub UserControl_InitProperties()
    CurrentAutoRedraw = True
End Sub

' Called when the control is resized
Private Sub UserControl_Resize()
    UserControl.ScaleMode = vbPixels
    If CurrentAutoRedraw = True Then
        Refresh
    End If
End Sub

'===========================================================================
'
' User Control Properties
'
'===========================================================================

' Refrech function, called from every where
Public Sub Refresh()
Attribute Refresh.VB_UserMemId = 0
    'Design the button
    DesignButton
    ' and display it regarding the current state
    DisplayButton ButtonState
End Sub

'no need to comment this one :)
Private Sub UserControl_ReadProperties(PBag As PropertyBag)
    Set CurrentTexture = PBag.ReadProperty("Texture", Nothing)
    CurrentBackColor = PBag.ReadProperty("BackColor", &HFF8080)
    CurrentShapeFormat = PBag.ReadProperty("ShapeFormat", RoundRectangleShape)
    Set CurrentShapeMaskPicture = PBag.ReadProperty("ShapeMaskPicture", Nothing)
    CurrentRoundCornerWidth = PBag.ReadProperty("RoundCornerWidth", DefaultRoundCornerWidth)
    CurrentPolygonSides = PBag.ReadProperty("PolygonSides", 5)
    CurrentShapeRotation = PBag.ReadProperty("ShapeRotation", 0)
    CurrentDepth = PBag.ReadProperty("Depth", 50)
    CurrentCaption = PBag.ReadProperty("Caption", UserControl.Name)
    CurrentAlignment = PBag.ReadProperty("Alignment", CenterCaption)
    CurrentForeColor = PBag.ReadProperty("ForeColor", vbBlack)
    Set CurrentFont = PBag.ReadProperty("Font", UserControl.Font)
    CurrentCaptionRotation = PBag.ReadProperty("CaptionRotation", 0)
    CurrentPictureAlignment = PBag.ReadProperty("PictureAlignment", CenterPicture)
    Set CurrentPicture = PBag.ReadProperty("Picture", Nothing)
    CurrentPictureTransparency = PBag.ReadProperty("PictureTransparency", False)
    CurrentPictureMaskColor = PBag.ReadProperty("PictureMaskColor", vbBlack)
    CurrentEnabled = PBag.ReadProperty("Enabled", True)
    CurrentAutoRedraw = PBag.ReadProperty("AutoRedraw", True)
    CurrentPieStartAngle = PBag.ReadProperty("PieStartAngle", 0)
    CurrentPieEndAngle = PBag.ReadProperty("PieEndAngle", 270)
    CurrentCaptionDeltaX = PBag.ReadProperty("CaptionDeltaX", 0)
    CurrentCaptionDeltaY = PBag.ReadProperty("CaptionDeltaY", 0)
    CurrentPictureDeltaX = PBag.ReadProperty("PictureDeltaX", 0)
    CurrentPictureDeltaY = PBag.ReadProperty("PictureDeltaY", 0)
    CurrentSlantStart = PBag.ReadProperty("SlantStart", 10)
    CurrentSlantEnd = PBag.ReadProperty("SlantEnd", 10)
    CurrentAmbientLight = PBag.ReadProperty("AmbientLight", 1)
    CurrentBackgroundMode = PBag.ReadProperty("BackgroundMode", SolidBG)
    CurrentBackEndColor = PBag.ReadProperty("BackEndColor", vbWhite)
    CurrentGradientAngle = PBag.ReadProperty("GradientAngle", 0)
    'Check just if we are disabled or not
    If CurrentEnabled = False Then
        ButtonState = ButtonIsDisabled
    Else
        ButtonState = ButtonisUp
    End If
    
    'Remove transparency if container does not have a picture
    If CurrentBackgroundMode = TransparentBG Then
        On Error Resume Next 'to avoid errors when parent is not available
        If UserControl.Extender.Container.Picture.Handle = 0 Then
            CurrentBackgroundMode = SolidBG
        End If
        On Error GoTo 0
    End If
    
    'Still the same dirty trick: the timer will run in background every second
    ' and will check the position of the control
    If CurrentBackgroundMode = TransparentBG Then
        Timer1.Enabled = True
    End If
    
    'We have set all the properties let's draw
    CurrentAutoRedraw = True
    Refresh
End Sub

' really need a comment?
Private Sub UserControl_WriteProperties(PBag As PropertyBag)
    PBag.WriteProperty "Picture", CurrentPicture, Nothing
    PBag.WriteProperty "Texture", CurrentTexture, Nothing
    PBag.WriteProperty "ShapeMaskPicture", CurrentShapeMaskPicture, Nothing
    PBag.WriteProperty "RoundCornerWidth", CurrentRoundCornerWidth, DefaultRoundCornerWidth
    PBag.WriteProperty "ShapeFormat", CurrentShapeFormat, RoundRectangleShape
    PBag.WriteProperty "Caption", CurrentCaption, UserControl.Name
    PBag.WriteProperty "Alignment", CurrentAlignment, CenterCaption
    PBag.WriteProperty "ForeColor", CurrentForeColor, vbBlack
    PBag.WriteProperty "Font", CurrentFont, Nothing
    PBag.WriteProperty "BackColor", CurrentBackColor, &HFF8080
    PBag.WriteProperty "PolygonSides", CurrentPolygonSides, 5
    PBag.WriteProperty "ShapeRotation", CurrentShapeRotation, 0
    PBag.WriteProperty "Enabled", CurrentEnabled, True
    PBag.WriteProperty "Depth", CurrentDepth, 50
    PBag.WriteProperty "AutoRedraw", CurrentAutoRedraw, True
    PBag.WriteProperty "PictureAlignment", CurrentPictureAlignment, CenterPicture
    PBag.WriteProperty "PictureMaskColor", CurrentPictureMaskColor, vbBlack
    PBag.WriteProperty "PictureTransparency", CurrentPictureTransparency, False
    PBag.WriteProperty "CaptionRotation", CurrentCaptionRotation, 0
    PBag.WriteProperty "PieStartAngle", CurrentPieStartAngle, 0
    PBag.WriteProperty "PieEndAngle", CurrentPieEndAngle, 270
    PBag.WriteProperty "CaptionDeltaX", CurrentCaptionDeltaX, 0
    PBag.WriteProperty "CaptionDeltaY", CurrentCaptionDeltaY, 0
    PBag.WriteProperty "PictureDeltaX", CurrentPictureDeltaX, 0
    PBag.WriteProperty "PictureDeltaY", CurrentPictureDeltaY, 0
    PBag.WriteProperty "SlantStart", CurrentSlantStart, 10
    PBag.WriteProperty "SlantEnd", CurrentSlantEnd, 10
    PBag.WriteProperty "AmbientLight", CurrentAmbientLight, 1
    PBag.WriteProperty "BackgroundMode", CurrentBackgroundMode, SolidBG
    PBag.WriteProperty "BackEndColor", CurrentBackEndColor, vbWhite
    PBag.WriteProperty "GradientAngle", CurrentGradientAngle, 0
End Sub

'===========================================================================
'
' All the following methods are called when setting a property
' mainly every time a property is changed we call the refresh function
' autoredraw flag is tested in the design method and in the display method
'
'===========================================================================

Public Property Get Texture() As StdPicture
Attribute Texture.VB_Description = "Returns/Sets the texture image"
    Set Texture = CurrentTexture
End Property

Public Property Set Texture(NewValue As StdPicture)
    Set CurrentTexture = NewValue
    Refresh
    PropertyChanged "Texture"
End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/Sets the picture"
    Set Picture = CurrentPicture
End Property

Public Property Set Picture(NewValue As StdPicture)
    Set CurrentPicture = NewValue
    Refresh
    PropertyChanged "Picture"
End Property

Public Property Get ShapeMaskPicture() As StdPicture
Attribute ShapeMaskPicture.VB_Description = "Returns/Sets the picture is the shape format is set to mask"
    Set ShapeMaskPicture = CurrentShapeMaskPicture
End Property

Public Property Set ShapeMaskPicture(NewValue As StdPicture)
    Set CurrentShapeMaskPicture = NewValue
    Refresh
    PropertyChanged "ShapeMaskPicture"
End Property

Public Property Get RoundCornerWidth() As Integer
Attribute RoundCornerWidth.VB_Description = "Returns/Sets the width of the rounded corners in pixels"
    RoundCornerWidth = CurrentRoundCornerWidth
End Property

Public Property Let RoundCornerWidth(NewValue As Integer)
    CurrentRoundCornerWidth = NewValue
    Refresh
    PropertyChanged "RoundCornerWidth"
End Property

Public Property Get ShapeFormat() As ShapeFormat_Enum
Attribute ShapeFormat.VB_Description = "Returns/Sets the shape format"
    ShapeFormat = CurrentShapeFormat
End Property

Public Property Let ShapeFormat(NewValue As ShapeFormat_Enum)
    CurrentShapeFormat = NewValue
    Refresh
    PropertyChanged "CurrentShapeFormat"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/Sets the button caption"
    Caption = CurrentCaption
End Property

Public Property Let Caption(NewValue As String)
    CurrentCaption = NewValue
    Refresh
    PropertyChanged "Caption"
End Property

Public Property Get Alignment() As CaptionAlignment_Enum
Attribute Alignment.VB_Description = "Returns/Sets Caption Alignment"
    Alignment = CurrentAlignment
End Property

Public Property Let Alignment(NewValue As CaptionAlignment_Enum)
    CurrentAlignment = NewValue
    Refresh
    PropertyChanged "Alignment"
End Property

Public Property Get PictureAlignment() As PictureAligment_Enum
Attribute PictureAlignment.VB_Description = "Returns/Sets the picture alignment"
    PictureAlignment = CurrentPictureAlignment
End Property

Public Property Let PictureAlignment(NewValue As PictureAligment_Enum)
    CurrentPictureAlignment = NewValue
    Refresh
    PropertyChanged "PictureAlignment"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/Sets the caption color"
    ForeColor = CurrentForeColor
End Property

Public Property Let ForeColor(NewValue As OLE_COLOR)
    CurrentForeColor = NewValue
    Refresh
    PropertyChanged "ForeColor"
End Property

Public Property Get PictureMaskColor() As OLE_COLOR
Attribute PictureMaskColor.VB_Description = "Returns/Sets the picture mask color, used if transparency is set to true"
    PictureMaskColor = CurrentPictureMaskColor
End Property

Public Property Let PictureMaskColor(NewValue As OLE_COLOR)
    CurrentPictureMaskColor = NewValue
    Refresh
    PropertyChanged "PictureMaskColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/Sets the caption font"
    Set Font = CurrentFont
End Property

Public Property Set Font(NewValue As Font)
    CurrentFont.Name = NewValue.Name
    CurrentFont.Size = NewValue.Size
    CurrentFont.Bold = NewValue.Bold
    CurrentFont.Italic = NewValue.Italic
    CurrentFont.Underline = NewValue.Underline
    CurrentFont.Strikethrough = NewValue.Strikethrough
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = CurrentBackColor
End Property

Public Property Let BackColor(NewValue As OLE_COLOR)
    CurrentBackColor = NewValue
    Refresh
    PropertyChanged "BackColor"
End Property

Public Property Get BackEndColor() As OLE_COLOR
    BackEndColor = CurrentBackEndColor
End Property

Public Property Let BackEndColor(NewValue As OLE_COLOR)
    CurrentBackEndColor = NewValue
    Refresh
    PropertyChanged "BackEndColor"
End Property

Public Property Get PolygonSides() As Integer
    PolygonSides = CurrentPolygonSides
End Property

Public Property Let PolygonSides(NewValue As Integer)
    CurrentPolygonSides = NewValue
    Refresh
    PropertyChanged "PolygonSides"
End Property

Public Property Get GradientAngle() As Integer
    GradientAngle = CurrentGradientAngle
End Property

Public Property Let GradientAngle(NewValue As Integer)
    CurrentGradientAngle = NewValue
    Refresh
    PropertyChanged "GradientAngle"
End Property

Public Property Get AmbientLight() As Single
Attribute AmbientLight.VB_Description = "Returns/Sets the light corrector factor"
    AmbientLight = CurrentAmbientLight
End Property

Public Property Let AmbientLight(NewValue As Single)
    CurrentAmbientLight = NewValue
    Refresh
    PropertyChanged "AmbientLight"
End Property

Public Property Get ShapeRotation() As Integer
Attribute ShapeRotation.VB_Description = "Returns/Sets the shape rotation, used for the polygon shape, the star shape, the cross shape and the arrow shape."
    ShapeRotation = CurrentShapeRotation
End Property

Public Property Let ShapeRotation(NewValue As Integer)
    CurrentShapeRotation = NewValue
    Refresh
    PropertyChanged "ShapeRotation"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/Sets if the control is enabled or not"
    Enabled = CurrentEnabled
End Property

Public Property Let Enabled(NewValue As Boolean)
    CurrentEnabled = NewValue
    If CurrentEnabled = False Then
        ButtonState = ButtonIsDisabled
    Else
        ButtonState = ButtonisUp
    End If
    DisplayButton ButtonState
    PropertyChanged "Enabled"
End Property

Public Property Get PictureTransparency() As Boolean
Attribute PictureTransparency.VB_Description = "Returns/Sets the transparency flag"
    PictureTransparency = CurrentPictureTransparency
End Property

Public Property Let PictureTransparency(NewValue As Boolean)
    CurrentPictureTransparency = NewValue
    Refresh
    PropertyChanged "PictureTransparency"
End Property

Public Property Get Depth() As Integer
Attribute Depth.VB_Description = "Returns/Sets the depth in percent, used for the star shape, Cross Shape and the L shape"
    Depth = CurrentDepth
End Property

Public Property Let Depth(NewValue As Integer)
    CurrentDepth = NewValue
    Refresh
    PropertyChanged "Depth"
End Property

Public Property Get CaptionRotation() As Integer
Attribute CaptionRotation.VB_Description = "Returns/Sets the caption rotation in degree"
    CaptionRotation = CurrentCaptionRotation
End Property

Public Property Let CaptionRotation(NewValue As Integer)
    CurrentCaptionRotation = NewValue
    Refresh
    PropertyChanged "CaptionRotation"
End Property

Public Property Get PieStartAngle() As Integer
Attribute PieStartAngle.VB_Description = "Returns/Sets the pie start angle in degree"
    PieStartAngle = CurrentPieStartAngle
End Property

Public Property Let PieStartAngle(NewValue As Integer)
    CurrentPieStartAngle = NewValue
    Refresh
    PropertyChanged "PieStartAngle"
End Property

Public Property Get PieEndAngle() As Integer
Attribute PieEndAngle.VB_Description = "Returns/Sets the pie end angle in degree"
    PieEndAngle = CurrentPieEndAngle
End Property

Public Property Let PieEndAngle(NewValue As Integer)
    CurrentPieEndAngle = NewValue
    Refresh
    PropertyChanged "PieEndAngle"
End Property

Public Property Get SlantStart() As Integer
Attribute SlantStart.VB_Description = "Returns/Sets the slant in pixels for the left side of the shape, used in slant shape and arrow head shape"
    SlantStart = CurrentSlantStart
End Property

Public Property Let SlantStart(NewValue As Integer)
    CurrentSlantStart = NewValue
    Refresh
    PropertyChanged "SlantStart"
End Property

Public Property Get SlantEnd() As Integer
Attribute SlantEnd.VB_Description = "Returns/Sets the slant in pixels for the right side of the shape, used in slant shape and arrow head shape"
    SlantEnd = CurrentSlantEnd
End Property

Public Property Let SlantEnd(NewValue As Integer)
    CurrentSlantEnd = NewValue
    Refresh
    PropertyChanged "SlantEnd"
End Property

Public Property Get CaptionDeltaX() As Integer
Attribute CaptionDeltaX.VB_Description = "Returns/Sets the caption horizontal delta in pixels"
    CaptionDeltaX = CurrentCaptionDeltaX
End Property

Public Property Let CaptionDeltaX(NewValue As Integer)
    CurrentCaptionDeltaX = NewValue
    Refresh
    PropertyChanged "CaptionDeltaX"
End Property

Public Property Get CaptionDeltaY() As Integer
Attribute CaptionDeltaY.VB_Description = "Returns/Sets the caption vertical delta in pixels"
    CaptionDeltaY = CurrentCaptionDeltaY
End Property

Public Property Let CaptionDeltaY(NewValue As Integer)
    CurrentCaptionDeltaY = NewValue
    Refresh
    PropertyChanged "CaptionDeltaY"
End Property

Public Property Get PictureDeltaX() As Integer
Attribute PictureDeltaX.VB_Description = "Returns/Sets picture horizontal delta in pixels"
    PictureDeltaX = CurrentPictureDeltaX
End Property

Public Property Let PictureDeltaX(NewValue As Integer)
    CurrentPictureDeltaX = NewValue
    Refresh
    PropertyChanged "PictureDeltaX"
End Property

Public Property Get PictureDeltaY() As Integer
Attribute PictureDeltaY.VB_Description = "Returns/Sets the picture vertical delta in pixels"
    PictureDeltaY = CurrentPictureDeltaY
End Property

Public Property Let PictureDeltaY(NewValue As Integer)
    CurrentPictureDeltaY = NewValue
    Refresh
    PropertyChanged "PictureDeltaY"
End Property

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/Sets AutoRedraw when changing a parameter"
    AutoRedraw = CurrentAutoRedraw
End Property

Public Property Let AutoRedraw(NewValue As Boolean)
    CurrentAutoRedraw = NewValue
    If CurrentEnabled = True Then
        Refresh
    End If
    PropertyChanged "AutoRedraw"
End Property

Public Property Get BackgroundMode() As BackgroundMode_Enum
Attribute BackgroundMode.VB_Description = "Returns/Sets the transparency flag, set the back color to white to have a real transparency."
    BackgroundMode = CurrentBackgroundMode
End Property

Public Property Let BackgroundMode(NewValue As BackgroundMode_Enum)
    CurrentBackgroundMode = NewValue
    'Remove transparency if container does not have a picture
    If CurrentBackgroundMode = TransparentBG Then
        On Error Resume Next 'to avoid errors when parent is not available
        If UserControl.Extender.Container.Picture.Handle = 0 Then
            CurrentBackgroundMode = SolidBG
        End If
        On Error GoTo 0
        Timer1.Enabled = True
    End If
    Refresh
    PropertyChanged "BackgroundMode"
End Property

'===========================================================================
'
' The following Functions and Methods concern all the drawing
'
'===========================================================================

'Merge the texture and the embossed design to create the button
Private Sub Texturize(Destination As BITMAPINFO, Delta As Integer)
    Dim iX As Integer
    Dim iY As Integer
    Dim GreyLevel As Long
    Dim Color As Long
    Dim Light As Single
    Dim Base As Integer
    Dim tR As Integer
    Dim tG As Integer
    Dim tB As Integer
    
    'Read all pixels loops
    For iX = 0 To DesignWidth - 1
        For iY = 0 To DesignHeight - 1
            
            ' check the light
            GreyLevel = EmbossBlurBuffer.Bytes(iX, iY)
            Light = (GreyLevel / 128) * CurrentAmbientLight
            Base = (GreyLevel - 128) * CurrentAmbientLight
            
            ' compute the new pixel color and check the bounds
            tB = (Light * TextureBuffer.Bytes(0, iX + Delta, iY + Delta)) + Base
            tB = IIf(tB < 0, 0, IIf(tB > 255, 255, tB))
            
            tG = (Light * TextureBuffer.Bytes(1, iX + Delta, iY + Delta)) + Base
            tG = IIf(tG < 0, 0, IIf(tG > 255, 255, tG))
            
            tR = (Light * TextureBuffer.Bytes(2, iX + Delta, iY + Delta)) + Base
            tR = IIf(tR < 0, 0, IIf(tR > 255, 255, tR))
            
            'Store in the destination image the new pixel color
            Destination.Bytes(0, iX, iY) = tB
            Destination.Bytes(1, iX, iY) = tG
            Destination.Bytes(2, iX, iY) = tR
        Next iY
    Next iX
End Sub

' Base fonction where the shape of the button is either designed
' either imported from a bitmap mask
Private Sub DrawDesignShape()
    Dim PolygonCoord() As PointApi
    Dim ReturnValue As Long
    Dim i As Double
    Dim PolygonPoint As Integer
    Dim CurrentRotation As Double
    Dim StarDelta As Integer
    Dim WorkingDC As Long
    Dim WorkingBitmap As Long
    Dim WorkingBitmapObject As Long
    Dim PaintBrush As Long
    Dim OldPaintBrush As Long
    Dim NewPen As Long
    Dim OldPen As Long
    Dim MaskDC As Long
    Dim WorkingMaskObject As Long
    Dim Tw As Long
    Dim Th As Long
    Dim Angle As Single

    'Create a DC and a bitmap to draw in
    WorkingDC = CreateCompatibleDC(UserControl.hdc)
    WorkingBitmap = CreateCompatibleBitmap(UserControl.hdc, DesignWidth - 1, DesignHeight - 1)
    WorkingBitmapObject = SelectObject(WorkingDC, WorkingBitmap)
    'Create the brush (white on black)
    PaintBrush = CreateSolidBrush(&HFFFFFF)
    OldPaintBrush = SelectObject(WorkingDC, PaintBrush)
    NewPen = CreatePen(PS_SOLID, 1, &HFFFFFF)
    OldPen = SelectObject(WorkingDC, NewPen)
       
    'Draw the shapes
    Select Case CurrentShapeFormat
        Case EllipseShape
            ReturnValue = Ellipse(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
        Case PolygoneShape
            ReDim PolygonCoord(CurrentPolygonSides)
            PolygonPoint = 0
            CurrentRotation = ((2 * PI) / 360) * CurrentShapeRotation
            For i = 0 To 2 * PI Step (2 * PI) / CurrentPolygonSides
                PolygonCoord(PolygonPoint).X = Cos(i + CurrentRotation) * ((DesignWidth - Thickness - 2 - Thickness - 2) / 2) + ((DesignWidth) / 2)
                PolygonCoord(PolygonPoint).Y = Sin(i + CurrentRotation) * ((DesignHeight - Thickness - 2 - Thickness - 2) / 2) + ((DesignHeight) / 2)
                PolygonPoint = PolygonPoint + 1
            Next i
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), CurrentPolygonSides + 1)
        Case Starshape
            ReDim PolygonCoord(CurrentPolygonSides * 2)
            PolygonPoint = 0
            CurrentRotation = ((2 * PI) / 360) * CurrentShapeRotation
            For i = 0 To 2 * PI Step (2 * PI) / (CurrentPolygonSides * 2)
                If (PolygonPoint Mod 2) = 1 Then
                    StarDelta = CurrentDepth
                Else
                    StarDelta = 0
                End If
                PolygonCoord(PolygonPoint).X = Cos(i + CurrentRotation) * (((DesignWidth - Thickness - 2 - Thickness - 2) / 2) * (1 - StarDelta / 100)) + ((DesignWidth) / 2)
                PolygonCoord(PolygonPoint).Y = Sin(i + CurrentRotation) * (((DesignHeight - Thickness - 2 - Thickness - 2) / 2) * (1 - StarDelta / 100)) + ((DesignHeight) / 2)
                PolygonPoint = PolygonPoint + 1
            Next i
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), (CurrentPolygonSides * 2))
        Case RectangleShape
            ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
        Case RoundRectangleShape
            ReturnValue = RoundRect(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2, CurrentRoundCornerWidth, CurrentRoundCornerWidth)
        Case MaskShape
            If CurrentShapeMaskPicture Is Nothing Then
                ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
            Else
                MaskDC = CreateCompatibleDC(UserControl.hdc)
                WorkingMaskObject = SelectObject(MaskDC, CurrentShapeMaskPicture.Handle)
                Tw = ScaleX(CurrentShapeMaskPicture.Width, vbHimetric, vbPixels)
                Th = ScaleY(CurrentShapeMaskPicture.Height, vbHimetric, vbPixels)
                BitBlt WorkingDC, 0, 0, Tw, Th, MaskDC, 0, 0, SRCCOPY
                SelectObject MaskDC, WorkingMaskObject
                DeleteDC MaskDC
            End If
        Case ArrowShape
            FillArrowArray
            ReDim PolygonCoord(7)
            For i = 0 To 7
                Angle = (2 * PI) * (CurrentShapeRotation + ArrowPoints(i).Angle) / 360
                PolygonCoord(i).X = (Cos(Angle) * ArrowPoints(i).RX) + (DesignWidth / 2)
                PolygonCoord(i).Y = (Sin(Angle) * ArrowPoints(i).RY) + (DesignHeight / 2)
            Next i
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), 8)
        Case PieShape
            ReDim PolygonCoord(0 To 1)
            Angle = (2 * PI) * (CurrentShapeRotation + 360 - CurrentPieStartAngle) / 360
            PolygonCoord(0).X = (Cos(Angle) * DesignWidth) + (DesignWidth / 2)
            PolygonCoord(0).Y = (Sin(Angle) * DesignHeight) + (DesignHeight / 2)
            Angle = (2 * PI) * (CurrentShapeRotation + 360 - CurrentPieEndAngle) / 360
            PolygonCoord(1).X = (Cos(Angle) * DesignWidth) + (DesignWidth / 2)
            PolygonCoord(1).Y = (Sin(Angle) * DesignHeight) + (DesignHeight / 2)
            ReturnValue = Pie(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2, PolygonCoord(0).X, PolygonCoord(0).Y, PolygonCoord(1).X, PolygonCoord(1).Y)
        Case LShape
            Select Case CurrentShapeRotation
                Case 0 To 89
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth * (100 - CurrentDepth) / 100, DesignHeight - Thickness - 2)
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, DesignHeight * (CurrentDepth) / 100, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
                Case 90 To 179
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth * (100 - CurrentDepth) / 100, DesignHeight - Thickness - 2)
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight * (100 - CurrentDepth) / 100)
                Case 180 To 269
                    ReturnValue = Rectangle(WorkingDC, DesignWidth * (CurrentDepth) / 100, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight * (100 - CurrentDepth) / 100)
                Case Else
                    ReturnValue = Rectangle(WorkingDC, DesignWidth * (CurrentDepth) / 100, Thickness + 2, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
                    ReturnValue = Rectangle(WorkingDC, Thickness + 2, DesignHeight * (CurrentDepth) / 100, DesignWidth - Thickness - 2, DesignHeight - Thickness - 2)
            End Select
        Case SlantShape
            ReDim PolygonCoord(0 To 3)
            PolygonCoord(0).X = Thickness + 2 + IIf(CurrentSlantStart > 0, CurrentSlantStart, 0)
            PolygonCoord(0).Y = Thickness + 2
            PolygonCoord(1).X = DesignWidth - Thickness - 2 + IIf(CurrentSlantEnd < 0, CurrentSlantEnd, 0)
            PolygonCoord(1).Y = Thickness + 2
            PolygonCoord(2).X = DesignWidth - Thickness - 2 - IIf(CurrentSlantEnd > 0, CurrentSlantEnd, 0)
            PolygonCoord(2).Y = DesignHeight - Thickness - 2
            PolygonCoord(3).X = Thickness + 2 - IIf(CurrentSlantStart < 0, CurrentSlantStart, 0)
            PolygonCoord(3).Y = DesignHeight - Thickness - 2
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), 4)
        Case ArrowHeadShape
            ReDim PolygonCoord(0 To 5)
            PolygonCoord(0).X = Thickness + 2 + IIf(CurrentSlantStart > 0, CurrentSlantStart, 0)
            PolygonCoord(0).Y = Thickness + 2
            PolygonCoord(1).X = DesignWidth - Thickness - 2 + IIf(CurrentSlantEnd < 0, CurrentSlantEnd, 0)
            PolygonCoord(1).Y = Thickness + 2
            PolygonCoord(2).X = DesignWidth - Thickness - 2 - IIf(CurrentSlantEnd > 0, CurrentSlantEnd, 0)
            PolygonCoord(2).Y = DesignHeight / 2
            PolygonCoord(3).X = DesignWidth - Thickness - 2 + IIf(CurrentSlantEnd < 0, CurrentSlantEnd, 0)
            PolygonCoord(3).Y = DesignHeight - Thickness - 2
            PolygonCoord(4).X = Thickness + 2 + IIf(CurrentSlantStart > 0, CurrentSlantStart, 0)
            PolygonCoord(4).Y = DesignHeight - Thickness - 2
            PolygonCoord(5).X = Thickness + 2 - IIf(CurrentSlantStart < 0, CurrentSlantStart, 0)
            PolygonCoord(5).Y = DesignHeight / 2
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), 6)
        Case CrossShape
            ReDim PolygonCoord(0 To 12)
            PolygonPoint = 0
            CurrentRotation = ((2 * PI) / 360) * CurrentShapeRotation
            For i = 0 To 2 * PI Step (2 * PI) / 12
                If (PolygonPoint Mod 3) = 0 Then
                    StarDelta = CurrentDepth
                Else
                    StarDelta = 0
                End If
                PolygonCoord(PolygonPoint).X = Cos(i + CurrentRotation) * (((DesignWidth - Thickness - 2 - Thickness - 2) / 2) * (1 - StarDelta / 100)) + ((DesignWidth) / 2)
                PolygonCoord(PolygonPoint).Y = Sin(i + CurrentRotation) * (((DesignHeight - Thickness - 2 - Thickness - 2) / 2) * (1 - StarDelta / 100)) + ((DesignHeight) / 2)
                PolygonPoint = PolygonPoint + 1
            Next i
            ReturnValue = Polygon(WorkingDC, PolygonCoord(0), 12)
            
    End Select
        
    'get the pixels from the drawing bitmap and store them in a buffer
    GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, DesignBuffer.Bytes(0, 0, 0), DesignBuffer, 0&
    
    'then delete the unused resources
    SelectObject WorkingDC, OldPaintBrush
    DeleteObject PaintBrush
    SelectObject WorkingDC, OldPen
    DeleteObject NewPen
    SelectObject WorkingDC, WorkingBitmapObject
    DeleteObject WorkingBitmap
    DeleteDC WorkingDC
End Sub

' Stupid method to create the design of an arrow
Private Sub FillArrowArray()
    Dim RadiusX As Single
    Dim RadiusY As Single
    RadiusX = (DesignWidth / 2) - Thickness - 2
    RadiusY = (DesignHeight / 2) - Thickness - 2
    ArrowPoints(0).RX = RadiusX
    ArrowPoints(0).RY = RadiusY
    ArrowPoints(0).Angle = 0
    ArrowPoints(1).RX = RadiusX
    ArrowPoints(1).RY = RadiusY
    ArrowPoints(1).Angle = 300
    ArrowPoints(2).RX = RadiusX * 0.71
    ArrowPoints(2).RY = RadiusY * 0.71
    ArrowPoints(2).Angle = 315
    ArrowPoints(3).RX = RadiusX
    ArrowPoints(3).RY = RadiusY
    ArrowPoints(3).Angle = 210
    ArrowPoints(4).RX = RadiusX
    ArrowPoints(4).RY = RadiusY
    ArrowPoints(4).Angle = 150
    ArrowPoints(5).RX = RadiusX * 0.71
    ArrowPoints(5).RY = RadiusY * 0.71
    ArrowPoints(5).Angle = 45
    ArrowPoints(6).RX = RadiusX
    ArrowPoints(6).RY = RadiusY
    ArrowPoints(6).Angle = 60
    ArrowPoints(7).RX = RadiusX
    ArrowPoints(7).RY = RadiusY
    ArrowPoints(7).Angle = 0
End Sub

'Cut the object to the shape, to get a better shape instead of re-using the shaping
' an analyze the blur image is used as a mask
Private Sub ShapeControl(DestHwnd As Long, SourceBuffer As LIGHTBITMAP)
    Dim iX As Integer
    Dim iY As Integer
    Dim i As Long
    Dim StartiY As Long
    Dim Region As Long
    Dim RegionTemp As Long
    Dim Result As Long
    
    'read the pixels from left to right
    For iX = 0 To DesignWidth - 1
        StartiY = -1
        ' and from top to bottom
        For iY = 0 To DesignHeight - 1
            ' The pixel is not black, so if we are not in an array start it
            If SourceBuffer.Bytes(iX, iY) > 0 Then
                If StartiY = -1 Then
                    StartiY = iY
                End If
            Else
                'The pixel is black, check if an array has been stated
                If StartiY > -1 Then
                    'it's the case so we have to define a region
                    'First time, so create it
                    If Region = 0 Then
                        Region = CreateRectRgn(iX, StartiY + 1, iX + 1, iY + 1)
                    Else
                        ' next time create a region, combine it to the first, and delete the object
                        RegionTemp = CreateRectRgn(iX, StartiY + 1, iX + 1, iY + 1)
                        Call CombineRgn(Region, Region, RegionTemp, RGN_OR)
                        DeleteObject RegionTemp
                    End If
                    StartiY = -1
                End If
            End If
        Next iY
        ' at the end of the loop do the same test
        If StartiY > -1 Then
            If Region = 0 Then
                Region = CreateRectRgn(iX, StartiY + 1, iX + 1, iY + 1)
            Else
                RegionTemp = CreateRectRgn(iX, StartiY + 1, iX + 1, iY + 1)
                Call CombineRgn(Region, Region, RegionTemp, RGN_OR)
                DeleteObject RegionTemp
            End If
        End If
    Next iX
    'and set the region of the control
    Result = SetWindowRgn(DestHwnd, Region, True)
    DeleteObject Region
End Sub

'Apply a Blur more filter to the base shape to create the gradient level
'the blur more filter is performed by doing a simple average on 25 pixel around
'the designated one.
'we use only the blue value, as the original picture is black an dwhite
'so we can set the destination pixel to a gray value
Private Sub ApplyBlurOnDesign()
    Dim iX As Integer
    Dim iY As Integer
    Dim bX As Integer
    Dim bY As Integer
    Dim Value As Long
    For iX = 2 To DesignWidth - 3
        For iY = 2 To DesignHeight - 3
            Value = 0
            For bX = iX - 2 To iX + 2
                For bY = iY - 2 To iY + 2
                    Value = Value + DesignBuffer.Bytes(0, bX, bY)
                Next bY
            Next bX
            Value = Value / 25
            BlurBuffer.Bytes(iX, iY) = Value
        Next iY
    Next iX
End Sub

'Apply a Blur Light filter to the embossed picture to smooth the borders
'same as above, except that the average is done on only 9 pixels
Private Sub ApplyBlurOnEmboss()
    Dim iX As Integer
    Dim iY As Integer
    Dim bX As Integer
    Dim bY As Integer
    Dim Value As Long
    For iX = 1 To DesignWidth - 2
        For iY = 1 To DesignHeight - 2
            Value = 0
            For bX = iX - 1 To iX + 1
                For bY = iY - 1 To iY + 1
                    Value = Value + EmbossBuffer.Bytes(bX, bY)
                Next bY
            Next bX
            Value = Value / 9
            EmbossBlurBuffer.Bytes(iX, iY) = Value
        Next iY
    Next iX
End Sub

'Apply Emboss filter regarding the direction
'(Down = read Left to Right & Top to Bottom,
' UP = Read Right to Left & Bottom to Top)
Private Sub EmbossEffect(Direction As Emboss_Enum, LightLevel As Single)
    Dim iX As Integer
    Dim iY As Integer
    Dim Value As Long
    Select Case Direction
        Case EmbossDown
            For iX = 0 To DesignWidth - 2
                For iY = 0 To DesignHeight - 2
                    Value = Abs(BlurBuffer.Bytes(iX, iY) - BlurBuffer.Bytes(iX + 1, iY + 1) * LightLevel + 128)
                    Value = IIf(Value < 255, Value, 255)
                    EmbossBuffer.Bytes(iX, iY) = Value
                Next iY
            Next iX
        Case EmbossUp
            For iX = DesignWidth - 1 To 1 Step -1
                For iY = DesignHeight - 1 To 1 Step -1
                    Value = Abs(BlurBuffer.Bytes(iX, iY) - BlurBuffer.Bytes(iX - 1, iY - 1) * LightLevel + 128)
                    Value = IIf(Value < 255, Value, 255)
                    EmbossBuffer.Bytes(iX, iY) = Value
                Next iY
            Next iX
    End Select
End Sub

'Convert the final Up Picture to grayscale to have a picture for the disabled state
'Greyscale is computed by doing the average of the sum of the different colors
Private Sub GrayScaleForDisabled()
    Dim iX As Integer
    Dim iY As Integer
    Dim i As Integer
    Dim Value As Long
    For iX = 0 To DesignWidth - 2
        For iY = 0 To DesignHeight - 2
            Value = 0
            For i = 0 To 2
                Value = Value + ButtonUpBuffer.Bytes(i, iX, iY)
            Next i
            Value = Value / 3
            ButtonDisabledBuffer.Bytes(0, iX, iY) = Value
            ButtonDisabledBuffer.Bytes(1, iX, iY) = Value
            ButtonDisabledBuffer.Bytes(2, iX, iY) = Value
        Next iY
    Next iX
End Sub

'Write the text and the picture (if one) in the button buffers
' As the drawtext API is crap when trying to rotate the text
' We output the text in a new buffer, then we merge the texture
' and the text (with the color) to the destination buffer
' and to avoid distortion we apply anti-aliasing
' the orginal anti-aliasing algorythm is taken
' from [Anti-aliasing Demo by Robert Rayment] found on PSC
' I change it to merge the 2 picture and to avoid the
' anti aliasing of the texture when not needed
Private Sub DrawTextAndPictureOnTexture()
    Dim TextTPHeight As Integer
    Dim TextDestRect As RECT
    Dim TestRect As RECT
    Dim DrawTextFlags As Long
    Dim TextFont As LOGFONT
    Dim WorkingDC As Long
    Dim WorkingBitmap As Long
    Dim WorkingBitmapObject As Long
    Dim FontHandle As Long
    Dim OldFontHandle As Long
    Dim DestX As Single
    Dim DestY As Single
    Dim PictureW As Single
    Dim PictureH As Single
    Dim PictureDC As Long
    Dim PictureObject As Long
    Dim SaveMode As Long
    Dim SaveColor As Long
    Dim TempBuffer As BITMAPINFO
    Dim ZCos As Double
    Dim ZSin As Double
    Dim CenterX As Integer
    Dim CenterY As Integer
    Dim iX As Integer
    Dim iY As Integer
    Dim SourceX As Single
    Dim SourceY As Single
    Dim SourceXint As Integer
    Dim SourceYint As Integer
    Dim WeightX As Single
    Dim WeightY As Single
    Dim Flag As Boolean
    Dim BMask As Long
    Dim GMask As Long
    Dim RMask As Long
    Dim pRed(0 To 1, 0 To 1) As Integer
    Dim pGreen(0 To 1, 0 To 1) As Integer
    Dim pBlue(0 To 1, 0 To 1) As Integer
    Dim jX As Integer
    Dim jY As Integer
    Dim PaintBrush As Long
    Dim OldPaintBrush As Long
    Dim NewAccessKey As String * 1

    'Create a DC to write on
    WorkingDC = CreateCompatibleDC(UserControl.hdc)
    WorkingBitmap = CreateCompatibleBitmap(UserControl.hdc, DesignWidth + DeltaShift, DesignHeight + DeltaShift)
    WorkingBitmapObject = SelectObject(WorkingDC, WorkingBitmap)
    
    
    PaintBrush = CreateSolidBrush(CurrentPictureMaskColor)
    OldPaintBrush = SelectObject(WorkingDC, PaintBrush)
    Rectangle WorkingDC, 0, 0, DesignWidth + DeltaShift + 1, DesignHeight + DeltaShift + 1
    SelectObject WorkingDC, OldPaintBrush
    DeleteObject PaintBrush
    
    'and copy the texture to it if the CaptionRotation is null we don't need to
    'compute the text rotation, so let's simplify the process
    If CurrentCaptionRotation = 0 Then
        SetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TextureBuffer.Bytes(0, 0, 0), TextureBuffer, 0&
    End If
    
    If Not CurrentPicture Is Nothing Then
        PictureW = ScaleX(CurrentPicture.Width, vbHimetric, vbPixels)
        PictureH = ScaleY(CurrentPicture.Height, vbHimetric, vbPixels)
        Select Case CurrentPictureAlignment
            Case CenterPicture
                DestX = ((DesignWidth - PictureW) / 2)
                DestY = ((DesignHeight - PictureH) / 2)
            Case leftPicture
                DestX = DefaultLeftBorder
                DestY = ((DesignHeight - PictureH) / 2)
            Case RightPicture
                DestX = DesignWidth - DefaultRightBorder - PictureW
                DestY = ((DesignHeight - PictureH) / 2)
            Case topPicture
                DestX = ((DesignWidth - PictureW) / 2)
                DestY = DefaultTopBorder
            Case BottomPicture
                DestX = ((DesignWidth - PictureW) / 2)
                DestY = DesignHeight - DefaultTopBorder - PictureH
        End Select
        DestX = DestX + CurrentPictureDeltaX
        DestY = DestY + CurrentPictureDeltaY
        PictureDC = CreateCompatibleDC(UserControl.hdc)
        PictureObject = SelectObject(PictureDC, CurrentPicture.Handle)
        If CurrentPictureTransparency Then
            'if the mask color is set we use this function which has a memory leak in older version of windows
            TransparentBlt WorkingDC, DestX, DestY, PictureW, PictureH, PictureDC, 0, 0, PictureW, PictureH, CurrentPictureMaskColor
        Else
            BitBlt WorkingDC, DestX, DestY, PictureW, PictureH, PictureDC, 0, 0, SRCCOPY
        End If
        SelectObject PictureDC, PictureObject
        DeleteDC PictureDC
    End If
    'Write the text
    DrawTextFlags = SetFontParameters(TextFont)
    SetBkMode WorkingDC, TRANSPARENT
    SetTextColor WorkingDC, CurrentForeColor
    'Create first the font
    FontHandle = CreateFontIndirect(TextFont)
    OldFontHandle = SelectObject(WorkingDC, FontHandle)
    'Compute the size of the rectangle to write the text in
    TextDestRect = GetTextRectangle
    TestRect = TextDestRect
    TextTPHeight = DrawText(WorkingDC, CurrentCaption, Len(CurrentCaption), TestRect, DT_CALCRECT Or DrawTextFlags)
    'and align the rectangle to the middle of the button
    TextDestRect.Top = TextDestRect.Top + (((TextDestRect.Bottom - TextDestRect.Top) - TextTPHeight) / 2) + CurrentCaptionDeltaY
    TextDestRect.Bottom = TextDestRect.Top + TextTPHeight '+ CurrentCaptionDeltaY
    TextDestRect.Left = TextDestRect.Left + CurrentCaptionDeltaX
    TextDestRect.Right = TextDestRect.Right + CurrentCaptionDeltaX
    DrawText WorkingDC, CurrentCaption, Len(CurrentCaption), TextDestRect, DrawTextFlags
    
    'Text must be rotated
    If CurrentCaptionRotation > 0 Then
        InitializeBitmapInfoHeader TempBuffer, DesignWidth + DeltaShift, DesignHeight + DeltaShift
        GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TempBuffer.Bytes(0, 0, 0), TempBuffer, 0&
        
        ZCos = Cos(PI / 180 * CurrentCaptionRotation)
        ZSin = Sin(PI / 180 * CurrentCaptionRotation)
        CenterX = DesignWidth / 2
        CenterY = DesignHeight / 2
        RMask = CurrentForeColor Mod 256
        GMask = Int((CurrentForeColor / 256)) Mod 256
        BMask = Int(CurrentForeColor / 65536)
    
        For iY = 0 To DesignHeight
            For iX = 0 To DesignWidth
                'Compute the position of the Source Pixel going to the
                'Current pixel once rotated
                'This value is rarely an integer, this gives a weight
                'for the close pixel use for anti aliasing
                SourceX = CenterX + (iX - CenterX) * ZCos + (iY - CenterY) * ZSin
                SourceY = CenterY + (iY - CenterY) * ZCos - (iX - CenterX) * ZSin
                SourceXint = Int(SourceX)
                SourceYint = Int(SourceY)
                'Check if we are in the bound of the rectangle when rotated
                If (SourceXint > 0 And SourceXint < DesignWidth - 1) And (SourceYint > 0 And SourceYint < DesignHeight - 1) Then
                    WeightX = SourceX - SourceXint
                    WeightY = SourceY - SourceYint
                    Flag = False
                    'if source pixel is in the text get the text color
                    ' else get the pixel in the texture
                    For jX = 0 To 1
                        For jY = 0 To 1
                            If RGB(TempBuffer.Bytes(2, SourceXint + jX, SourceYint + jY), TempBuffer.Bytes(1, SourceXint + jX, SourceYint + jY), TempBuffer.Bytes(0, SourceXint + jX, SourceYint + jY)) <> CurrentPictureMaskColor Then
                                pBlue(jX, jY) = TempBuffer.Bytes(0, SourceXint + jX, SourceYint + jY)
                                pGreen(jX, jY) = TempBuffer.Bytes(1, SourceXint + jX, SourceYint + jY)
                                pRed(jX, jY) = TempBuffer.Bytes(2, SourceXint + jX, SourceYint + jY)
                                'we have at least one pixel in the four coming from the text
                                Flag = True
                            Else
                                pBlue(jX, jY) = TextureBuffer.Bytes(0, iX + jX, iY + jY)
                                pGreen(jX, jY) = TextureBuffer.Bytes(1, iX + jX, iY + jY)
                                pRed(jX, jY) = TextureBuffer.Bytes(2, iX + jX, iY + jY)
                            End If
                        Next jY
                    Next jX
                    'we had as least one pixel so put the color in the texture box
                    If Flag Then
                        TextureBuffer.Bytes(0, iX, iY) = (1 - WeightY) * (((1 - WeightX) * pBlue(0, 0)) + (WeightX * pBlue(1, 0))) + WeightY * (((1 - WeightX) * pBlue(0, 1)) + (WeightX * pBlue(1, 1)))
                        TextureBuffer.Bytes(1, iX, iY) = (1 - WeightY) * (((1 - WeightX) * pGreen(0, 0)) + (WeightX * pGreen(1, 0))) + WeightY * (((1 - WeightX) * pGreen(0, 1)) + (WeightX * pGreen(1, 1)))
                        TextureBuffer.Bytes(2, iX, iY) = (1 - WeightY) * (((1 - WeightX) * pRed(0, 0)) + (WeightX * pRed(1, 0))) + WeightY * (((1 - WeightX) * pRed(0, 1)) + (WeightX * pRed(1, 1)))
                    End If
                End If
            Next iX
        Next iY
    Else
        GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TextureBuffer.Bytes(0, 0, 0), TextureBuffer, 0&
    End If
    SelectObject WorkingDC, OldFontHandle
    DeleteObject FontHandle
    SelectObject WorkingDC, WorkingBitmapObject
    DeleteObject WorkingBitmap
    DeleteDC WorkingDC
    
    If InStr(CurrentCaption, "&") Then
        NewAccessKey = Mid$(CurrentCaption, InStr(CurrentCaption, "&") + 1, 1)
        If NewAccessKey <> "&" Then UserControl.AccessKeys = NewAccessKey
    End If
End Sub

'Set The Font parameters
Private Function SetFontParameters(ThisFont As LOGFONT) As Long
    Dim ReturnFlag As Long
    ThisFont.lfHeight = (CurrentFont.Size * -20) / Screen.TwipsPerPixelY
    ThisFont.lfItalic = CurrentFont.Italic
    ThisFont.lfStrikeOut = CurrentFont.Strikethrough
    ThisFont.lfUnderline = CurrentFont.Underline
    ThisFont.lfCharSet = DEFAULT_CHARSET
    ThisFont.lfClipPrecision = CLIP_LH_ANGLES 'CLIP_DEFAULT_PRECIS
    ThisFont.lfEscapement = 0
    ThisFont.lfFaceName = Left$(CurrentFont.Name & String$(32, 0), 32)
    ThisFont.lfOrientation = 0
    ThisFont.lfOutPrecision = OUT_DEFAULT_PRECIS
    ThisFont.lfPitchAndFamily = DEFAULT_PITCH Or FF_DONTCARE
    ThisFont.lfQuality = ANTIALIASED_QUALITY
    ThisFont.lfWeight = IIf(CurrentFont.Bold, FW_BOLD, FW_NORMAL)
    ThisFont.lfWidth = 0
    Select Case CurrentAlignment
        Case LeftCaption
            ReturnFlag = DT_LEFT Or DT_WORDBREAK
        Case RightCaption
            ReturnFlag = DT_RIGHT Or DT_WORDBREAK
        Case CenterCaption
            ReturnFlag = DT_CENTER Or DT_WORDBREAK
    End Select
    SetFontParameters = ReturnFlag
End Function

'Define The Rectangle to output the caption
'regarding the place of the pisture (if one)
Private Function GetTextRectangle() As RECT
    Dim PictureW As Single
    Dim PictureH As Single
    If Not CurrentPicture Is Nothing Then
        PictureW = ScaleX(CurrentPicture.Width, vbHimetric, vbPixels)
        PictureH = ScaleY(CurrentPicture.Height, vbHimetric, vbPixels)
        Select Case CurrentPictureAlignment
            Case CenterPicture
                GetTextRectangle.Left = DefaultLeftBorder
                GetTextRectangle.Right = DesignWidth - DefaultRightBorder
                GetTextRectangle.Top = DefaultTopBorder
                GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder
            Case leftPicture
                GetTextRectangle.Left = DefaultLeftBorder + PictureH
                GetTextRectangle.Right = DesignWidth - DefaultRightBorder
                GetTextRectangle.Top = DefaultTopBorder
                GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder
            Case RightPicture
                GetTextRectangle.Left = DefaultLeftBorder
                GetTextRectangle.Right = DesignWidth - DefaultRightBorder - PictureW
                GetTextRectangle.Top = DefaultTopBorder
                GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder
            Case topPicture
                GetTextRectangle.Left = DefaultLeftBorder
                GetTextRectangle.Right = DesignWidth - DefaultRightBorder
                GetTextRectangle.Top = DefaultTopBorder + PictureH
                GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder
            Case BottomPicture
                GetTextRectangle.Left = DefaultLeftBorder
                GetTextRectangle.Right = DesignWidth - DefaultRightBorder
                GetTextRectangle.Top = DefaultTopBorder
                GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder - PictureH
        End Select
    Else
        GetTextRectangle.Left = DefaultLeftBorder
        GetTextRectangle.Right = DesignWidth - DefaultRightBorder
        GetTextRectangle.Top = DefaultTopBorder
        GetTextRectangle.Bottom = DesignHeight - DefaultBottomBorder
    End If
End Function

'Refresh the texture regarding either the current picture or the color
' in case of a picture the picture box is filled up with tiles
Private Sub RefreshTextureBox()
    Dim iX As Integer
    Dim iY As Integer
    Dim WorkingDC As Long
    Dim WorkingBitmap As Long
    Dim WorkingBitmapObject As Long
    Dim TextureDC As Long
    Dim WorkingTexture As Long
    Dim WorkingTextureObject As Long
    Dim Th As Single
    Dim Tw As Single
    Dim RealColor As Long
    Dim LeftInParent As Long
    Dim TopInParent As Long
    Dim BMask As Long
    Dim GMask As Long
    Dim RMask As Long
    
    If CurrentBackgroundMode = TransparentBG Then
        On Error Resume Next 'to avoid errors when parent is not available
        If UserControl.Extender.Container.Picture.Handle = 0 Then
            CurrentBackgroundMode = SolidBG
        End If
        On Error GoTo 0
    End If
    
    If CurrentBackColor < 0 Then
        RealColor = GetSysColor(CurrentBackColor And &HFF)
    Else
        RealColor = CurrentBackColor
    End If
    
    If CurrentBackgroundMode = TransparentBG Then
        WorkingDC = CreateCompatibleDC(UserControl.hdc)
        WorkingBitmap = CreateCompatibleBitmap(UserControl.hdc, DesignWidth + DeltaShift - 1, DesignHeight + DeltaShift - 1)
        WorkingBitmapObject = SelectObject(WorkingDC, WorkingBitmap)
        TextureDC = CreateCompatibleDC(UserControl.hdc)
        WorkingTextureObject = SelectObject(TextureDC, UserControl.Extender.Container.Picture.Handle)
        LeftInParent = ScaleX(UserControl.Extender.Left, Parent.ScaleMode, vbPixels)
        TopInParent = ScaleY(UserControl.Extender.Top, Parent.ScaleMode, vbPixels)
        BitBlt WorkingDC, iX, iY, DesignWidth, DesignHeight, TextureDC, LeftInParent, TopInParent, SRCCOPY
        GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TextureBuffer.Bytes(0, 0, 0), TextureBuffer, 0&
        SelectObject WorkingDC, WorkingBitmapObject
        DeleteObject WorkingBitmap
        DeleteDC WorkingDC
        SelectObject TextureDC, WorkingTextureObject
        DeleteDC TextureDC
        RMask = RealColor Mod 256
        GMask = Int((RealColor / 256)) Mod 256
        BMask = Int(RealColor / 65536)
        For iX = 0 To DesignWidth - 1
            For iY = 0 To DesignHeight - 1
                TextureBuffer.Bytes(0, iX, iY) = TextureBuffer.Bytes(0, iX, iY) And BMask
                TextureBuffer.Bytes(1, iX, iY) = TextureBuffer.Bytes(1, iX, iY) And GMask
                TextureBuffer.Bytes(2, iX, iY) = TextureBuffer.Bytes(2, iX, iY) And RMask
            Next iY
        Next iX
    ElseIf CurrentBackgroundMode = TextureBG And Not (CurrentTexture Is Nothing) Then
        WorkingDC = CreateCompatibleDC(UserControl.hdc)
        WorkingBitmap = CreateCompatibleBitmap(UserControl.hdc, DesignWidth + DeltaShift - 1, DesignHeight + DeltaShift - 1)
        WorkingBitmapObject = SelectObject(WorkingDC, WorkingBitmap)
        TextureDC = CreateCompatibleDC(UserControl.hdc)
        WorkingTextureObject = SelectObject(TextureDC, CurrentTexture.Handle)
        Tw = ScaleX(CurrentTexture.Width, vbHimetric, vbPixels)
        Th = ScaleY(CurrentTexture.Height, vbHimetric, vbPixels)
        For iX = 0 To DesignWidth + DeltaShift - 1 Step Tw
            For iY = 0 To DesignHeight + DeltaShift - 1 Step Th
                BitBlt WorkingDC, iX, iY, Tw, Th, TextureDC, 0, 0, SRCCOPY
            Next iY
        Next iX
        GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TextureBuffer.Bytes(0, 0, 0), TextureBuffer, 0&
        SelectObject WorkingDC, WorkingBitmapObject
        DeleteObject WorkingBitmap
        DeleteDC WorkingDC
        SelectObject TextureDC, WorkingTextureObject
        DeleteDC TextureDC
    ElseIf CurrentBackgroundMode = LinearGradientBG Then
        LinearGradientFill
    ElseIf CurrentBackgroundMode = RadialGradientBG Then
        RadialGradientFill
    ElseIf CurrentBackgroundMode = RectangularGradientBG Then
        RectangleGradientFill
    Else
        RMask = RealColor Mod 256
        GMask = Int((RealColor / 256)) Mod 256
        BMask = Int(RealColor / 65536)
        For iX = 0 To DesignWidth - 1
            For iY = 0 To DesignHeight - 1
                TextureBuffer.Bytes(0, iX, iY) = BMask
                TextureBuffer.Bytes(1, iX, iY) = GMask
                TextureBuffer.Bytes(2, iX, iY) = RMask
            Next iY
        Next iX
    End If
End Sub

' Thanks to Carles P.V. for his gradient algorithm
' this sub fill the texture buffer with a gradient fill
Private Sub LinearGradientFill()
    Dim StartBlue As Integer
    Dim StartRed As Integer
    Dim StartGreen As Integer
    Dim EndBlue As Integer
    Dim EndRed As Integer
    Dim EndGreen As Integer
    Dim Angle As Double
    Dim iX As Integer
    Dim iY As Integer
    Dim StartP As PointApi
    Dim EndP As PointApi
    Dim Dg As Double
    Dim Xc As Double
    Dim Yc As Double
    Dim YOut As Integer
    Dim XOut As Integer
    Dim c1 As Double
    Dim c2 As Double
    Dim RealColor As Long
    Dim NewGreen As Integer
    Dim NewBlue As Integer
    Dim NewRed As Integer
    
    If CurrentBackColor < 0 Then
        RealColor = GetSysColor(CurrentBackColor And &HFF)
    Else
        RealColor = CurrentBackColor
    End If
    StartBlue = Int(RealColor / 65536)
    StartGreen = Int((RealColor / 256)) Mod 256
    StartRed = RealColor Mod 256
    
    If CurrentBackEndColor < 0 Then
        RealColor = GetSysColor(CurrentBackEndColor And &HFF)
    Else
        RealColor = CurrentBackEndColor
    End If
    EndBlue = Int(RealColor / 65536)
    EndGreen = Int((RealColor / 256)) Mod 256
    EndRed = RealColor Mod 256
       
    Angle = CurrentGradientAngle * (PI / 180)
    
    EndP.X = Int(Cos(Angle) * (DesignWidth / 2) + (DesignWidth / 2))
    StartP.X = Int(Cos(Angle + PI) * (DesignWidth / 2) + (DesignWidth / 2))
    EndP.Y = Int(Sin(Angle) * (DesignHeight / 2) + (DesignHeight / 2))
    StartP.Y = Int(Sin(Angle + PI) * (DesignHeight / 2) + (DesignHeight / 2))
       
    Dg = Sqr((EndP.X - StartP.X) ^ 2 + (EndP.Y - StartP.Y) ^ 2)
    Xc = Cos(Angle) / Dg
    Yc = Sin(Angle) / Dg
        
    For iY = -StartP.Y To DesignHeight - 1 - StartP.Y
        YOut = iY + StartP.Y
        For iX = -StartP.X To DesignWidth - 1 - StartP.X
            XOut = iX + StartP.X
            c1 = iX * Xc + iY * Yc
            c1 = IIf(c1 > 1, c1, IIf(c1 < 0, 0, c1))
            c2 = 1 - c1
            NewBlue = (c1 * EndBlue + c2 * StartBlue)
            NewGreen = (c1 * EndGreen + c2 * StartGreen)
            NewRed = (c1 * EndRed + c2 * StartRed)
            TextureBuffer.Bytes(0, XOut, YOut) = IIf(NewBlue > 255, 255, NewBlue)
            TextureBuffer.Bytes(1, XOut, YOut) = IIf(NewGreen > 255, 255, NewGreen)
            TextureBuffer.Bytes(2, XOut, YOut) = IIf(NewRed > 255, 255, NewRed)
        Next iX
    Next iY
End Sub

'This sub fills the texture with a circular gradiant
Public Sub RadialGradientFill()
    Dim StartBlue As Integer
    Dim StartRed As Integer
    Dim StartGreen As Integer
    Dim EndBlue As Integer
    Dim EndRed As Integer
    Dim EndGreen As Integer
    Dim Angle As Double
    Dim iX As Integer
    Dim iY As Integer
    Dim StartP As PointApi
    Dim EndP As PointApi
    Dim Dg As Double
    Dim Xc As Double
    Dim Yc As Double
    Dim YOut As Integer
    Dim XOut As Integer
    Dim c1 As Double
    Dim c2 As Double
    Dim RealColor As Long
    Dim NewGreen As Integer
    Dim NewBlue As Integer
    Dim NewRed As Integer
    
    If CurrentBackColor < 0 Then
        RealColor = GetSysColor(CurrentBackColor And &HFF)
    Else
        RealColor = CurrentBackColor
    End If
    StartBlue = Int(RealColor / 65536)
    StartGreen = Int((RealColor / 256)) Mod 256
    StartRed = RealColor Mod 256
    
    If CurrentBackEndColor < 0 Then
        RealColor = GetSysColor(CurrentBackEndColor And &HFF)
    Else
        RealColor = CurrentBackEndColor
    End If
    EndBlue = Int(RealColor / 65536)
    EndGreen = Int((RealColor / 256)) Mod 256
    EndRed = RealColor Mod 256
       
    EndP.X = DesignWidth
    StartP.X = DesignWidth / 2
    EndP.Y = DesignHeight
    StartP.Y = DesignHeight / 2
       
       
    Dg = Sqr(StartP.X ^ 2 + StartP.Y ^ 2)
    For iY = -StartP.Y To DesignHeight - 1 - StartP.Y
        YOut = iY + StartP.Y
        For iX = -StartP.X To DesignWidth - 1 - StartP.X
            XOut = iX + StartP.X
                 
            c1 = Sqr(iX * iX + iY * iY) / Dg
            If (c1 > 1) Then c1 = 1
            c2 = 1 - c1
            
            NewBlue = (c1 * EndBlue + c2 * StartBlue)
            NewGreen = (c1 * EndGreen + c2 * StartGreen)
            NewRed = (c1 * EndRed + c2 * StartRed)
            TextureBuffer.Bytes(0, XOut, YOut) = IIf(NewBlue > 255, 255, NewBlue)
            TextureBuffer.Bytes(1, XOut, YOut) = IIf(NewGreen > 255, 255, NewGreen)
            TextureBuffer.Bytes(2, XOut, YOut) = IIf(NewRed > 255, 255, NewRed)
        Next iX
    Next iY
End Sub

Private Sub RectangleGradientFill()
    Dim StartBlue As Integer
    Dim StartRed As Integer
    Dim StartGreen As Integer
    Dim EndBlue As Integer
    Dim EndRed As Integer
    Dim EndGreen As Integer
    Dim IncGreen As Single
    Dim IncBlue As Single
    Dim IncRed As Single
    Dim Steps As Integer
    Dim CenterX As Integer
    Dim CenterY As Integer
    Dim XOffset As Single
    Dim YOffset As Single
    Dim RotateOffset As Single
    Dim i As Integer
    Dim j As Single
    Dim WorkingDC As Long
    Dim WorkingBitmap As Long
    Dim WorkingBitmapObject As Long
    Dim PaintBrush As Long
    Dim OldPaintBrush As Long
    Dim NewPen As Long
    Dim OldPen As Long
    Dim PolygonCoord(0 To 4) As PointApi
    Dim NewColor As Long
    Dim ReturnValue As Long
    Dim CurrentRotation As Single
    Dim RealColor As Long
    
    Dim DeltaX As Integer
    Dim DeltaY As Integer
    
    WorkingDC = CreateCompatibleDC(UserControl.hdc)
    WorkingBitmap = CreateCompatibleBitmap(UserControl.hdc, DesignWidth + DeltaShift - 1, DesignHeight + DeltaShift - 1)
    WorkingBitmapObject = SelectObject(WorkingDC, WorkingBitmap)
    
    If CurrentBackColor < 0 Then
        RealColor = GetSysColor(CurrentBackColor And &HFF)
    Else
        RealColor = CurrentBackColor
    End If
    StartBlue = Int(RealColor / 65536)
    StartGreen = Int((RealColor / 256)) Mod 256
    StartRed = RealColor Mod 256
    
    If CurrentBackEndColor < 0 Then
        RealColor = GetSysColor(CurrentBackEndColor And &HFF)
    Else
        RealColor = CurrentBackEndColor
    End If
    EndBlue = Int(RealColor / 65536)
    EndGreen = Int((RealColor / 256)) Mod 256
    EndRed = RealColor Mod 256
    
    CenterX = DesignWidth / 2
    CenterY = DesignHeight / 2
    Steps = IIf(CenterX < CenterY, CenterY, CenterX) + 1
    
    IncBlue = (EndBlue - StartBlue) / Steps
    IncGreen = (EndGreen - StartGreen) / Steps
    IncRed = (EndRed - StartRed) / Steps
    
    XOffset = CenterX / Steps
    YOffset = CenterY / Steps
    RotateOffset = CurrentGradientAngle / Steps

    For i = 0 To Steps
        CurrentRotation = ((2 * PI) / 360) * ((RotateOffset * i))
        PolygonCoord(0).X = CenterX + RotateX(-(CenterX - (i * XOffset)), -(CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(0).Y = CenterY + RotateY(-(CenterX - (i * XOffset)), -(CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(1).X = CenterX + RotateX((CenterX - (i * XOffset)), -(CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(1).Y = CenterY + RotateY((CenterX - (i * XOffset)), -(CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(2).X = CenterX + RotateX((CenterX - (i * XOffset)), (CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(2).Y = CenterY + RotateY((CenterX - (i * XOffset)), (CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(3).X = CenterX + RotateX(-(CenterX - (i * XOffset)), (CenterY - (i * YOffset)), CurrentRotation)
        PolygonCoord(3).Y = CenterY + RotateY(-(CenterX - (i * XOffset)), (CenterY - (i * YOffset)), CurrentRotation)
        
        NewColor = RGB(StartRed + (IncRed * i), StartGreen + (IncGreen * i), StartBlue + (IncBlue * i))
        PaintBrush = CreateSolidBrush(NewColor)
        OldPaintBrush = SelectObject(WorkingDC, PaintBrush)
        NewPen = CreatePen(PS_SOLID, 1, NewColor)
        OldPen = SelectObject(WorkingDC, NewPen)
        
        ReturnValue = Polygon(WorkingDC, PolygonCoord(0), 4)
        
        SelectObject WorkingDC, OldPen
        DeleteObject NewPen
        SelectObject WorkingDC, OldPaintBrush
        DeleteObject PaintBrush

    Next i

    GetDIBits WorkingDC, WorkingBitmap, 0, DesignHeight, TextureBuffer.Bytes(0, 0, 0), TextureBuffer, 0&
    SelectObject WorkingDC, WorkingBitmapObject
    DeleteObject WorkingBitmap
    DeleteDC WorkingDC

End Sub

'Basic Fonctions for point Rotation
Private Function RotateX(XRadius As Single, YRadius As Single, Angle As Single) As Double
    RotateX = XRadius * Cos(Angle) - YRadius * Sin(Angle)
End Function

Private Function RotateY(XRadius As Single, YRadius As Single, Angle As Single) As Double
    RotateY = YRadius * Cos(Angle) + XRadius * Sin(Angle)
End Function


'===========================================================================
'
' The following Functions and Methods are called to redesign the button
' and produce the pictures for the different button states
'
'===========================================================================

' Size the Tables and perform the pictures drawing
Private Sub DesignButton()
    If CurrentAutoRedraw = False Then Exit Sub
    'to be sure we work in pixel (twips is crap)
    UserControl.ScaleMode = vbPixels

    'Resize the tables only if the size has changed
    If DesignWidth <> UserControl.ScaleWidth Or DesignHeight <> UserControl.ScaleHeight Then
        DesignWidth = UserControl.ScaleWidth
        DesignHeight = UserControl.ScaleHeight
        InitializeBitmapInfoHeader DesignBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeBitmapInfoHeader TextureBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeBitmapInfoHeader ButtonUpBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeBitmapInfoHeader ButtonDownBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeBitmapInfoHeader ButtonOverBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeBitmapInfoHeader ButtonDisabledBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeLightBitmap BlurBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeLightBitmap EmbossBlurBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
        InitializeLightBitmap EmbossBuffer, DesignWidth + DeltaShift + DeltaShift, DesignHeight + DeltaShift + DeltaShift
    End If

    'Create the texture from either the color or the specified picture
    RefreshTextureBox
    'Write the text on the texture (and the picture if one)
    DrawTextAndPictureOnTexture
    'Draw the basic shape of the button
    DrawDesignShape
    'first a strong blur (this gives a gradient between the black and the white)
    ApplyBlurOnDesign
    
    'emboss the blured image to have the up or down effect
    EmbossEffect EmbossDown, 0.9
    'a small blur to have a smooth bevel effect
    ApplyBlurOnEmboss
    'and merge the texture(with the text and picture) with this bevel map
    Texturize ButtonDownBuffer, 0

    'do the same for the up button
    EmbossEffect EmbossUp, 1
    ApplyBlurOnEmboss
    Texturize ButtonUpBuffer, DeltaShift

    'do the same for the over button (light is set to .9 to have a light effect when passing over the button)
    EmbossEffect EmbossUp, 0.9
    ApplyBlurOnEmboss
    Texturize ButtonOverBuffer, DeltaShift
 
    'simply convert the upbutton to greyscale for the disabled picture
    GrayScaleForDisabled
    
    ' and finally cut the control
    'we sue the blur buffer which looks like a mask
    ShapeControl UserControl.hwnd, BlurBuffer

End Sub

'Display the button picture regarding the State
Private Sub DisplayButton(State As ButtonState_Enum)
    If CurrentAutoRedraw = False Then Exit Sub
    UserControl.Cls
    Select Case State
        Case ButtonIsDown
            SetDIBits UserControl.hdc, UserControl.Image.Handle, 0, DesignHeight, ButtonDownBuffer.Bytes(0, 0, 0), ButtonDownBuffer, 0&
        Case ButtonisUp
            SetDIBits UserControl.hdc, UserControl.Image.Handle, 0, DesignHeight, ButtonUpBuffer.Bytes(0, 0, 0), ButtonUpBuffer, 0&
        Case MouseIsOver
            SetDIBits UserControl.hdc, UserControl.Image.Handle, 0, DesignHeight, ButtonOverBuffer.Bytes(0, 0, 0), ButtonOverBuffer, 0&
        Case ButtonIsDisabled
            SetDIBits UserControl.hdc, UserControl.Image.Handle, 0, DesignHeight, ButtonDisabledBuffer.Bytes(0, 0, 0), ButtonDisabledBuffer, 0&
    End Select
    UserControl.Refresh
End Sub

'Method to initalize the different bitmap info used to store the pixels and images
Private Sub InitializeBitmapInfoHeader(Buffer As BITMAPINFO, SizeWidth As Long, SizeHeight As Long)
    ReDim Buffer.Bytes(3, SizeWidth - 1, SizeHeight - 1)
    Buffer.Header.biSize = 40
    Buffer.Header.biWidth = SizeWidth
    Buffer.Header.biHeight = -SizeHeight
    Buffer.Header.biPlanes = 1
    Buffer.Header.biBitCount = 32
    Buffer.Header.biSizeImage = 3 * SizeWidth * SizeHeight
End Sub

'Method to initalize the different bytes table used to store only one color
Private Sub InitializeLightBitmap(Buffer As LIGHTBITMAP, SizeWidth As Long, SizeHeight As Long)
    ReDim Buffer.Bytes(SizeWidth - 1, SizeHeight - 1)
End Sub




