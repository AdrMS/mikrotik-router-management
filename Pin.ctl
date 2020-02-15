VERSION 5.00
Begin VB.UserControl Pin 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   MaskColor       =   &H00C0C0C0&
   MaskPicture     =   "Pin.ctx":0000
   MouseIcon       =   "Pin.ctx":008C
   Picture         =   "Pin.ctx":02AA
   ScaleHeight     =   480
   ScaleWidth      =   480
End
Attribute VB_Name = "Pin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'We rely on mask transparency, so Picture must be our image and MaskPicture must
'be a "cutout" image with transparent areas = BackColor.

Private Const DEFAULT_COLOR = vbRed 'This must match the color of our background Picture
                                    'bitmap's areas to be recolored when Color is assigned
                                    'a new value.

Private Const DIB_RGB_COLORS As Long = 0&
Private Const BI_RGB As Long = 0&

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
'
'Private Type RGBQUAD
'    rgbBlue As Byte
'    rgbGreen As Byte
'    rgbRed As Byte
'    rgbReserved As Byte
'End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    'bmiColors(255) As RGBQUAD 'Not used here, we work with 24-bit color.
End Type

Private Declare Function GetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef Bits As Byte, _
    ByRef BI As BITMAPINFO, _
    ByVal wUsage As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetDIBits Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal hBitmap As Long, _
    ByVal nStartScan As Long, _
    ByVal nNumScans As Long, _
    ByRef Bits As Byte, _
    ByRef BI As BITMAPINFO, _
    ByVal wUsage As Long) As Long

Private ColorR As Byte
Private ColorG As Byte
Private ColorB As Byte

Private mColor As Long

Public Event Click()

Public Property Get Color() As OLE_COLOR
    Color = mColor
End Property

Public Property Let Color(ByVal RHS As OLE_COLOR)
    Dim RGBQUAD As Long
    Dim NewR As Byte
    Dim NewG As Byte
    Dim NewB As Byte
    Dim WidthPx As Long
    Dim HeightPx As Long
    Dim bmi As BITMAPINFO
    Dim Stride As Long
    Dim Triples() As Byte
    Dim Line As Long
    Dim Pixel As Long
    
    mColor = RHS
    If mColor And &H80000000 Then
        RGBQUAD = GetSysColor(mColor And &HFFFF&)
    Else
        RGBQUAD = mColor
    End If
    NewR = RGBQUAD And &HFF&
    NewG = (RGBQUAD And &HFF00&) \ &H100&
    NewB = RGBQUAD \ &H10000

    With UserControl
        WidthPx = ScaleX(.Picture.Width, vbHimetric, vbPixels)
        HeightPx = ScaleY(.Picture.Height, vbHimetric, vbPixels)
        With bmi.bmiHeader
            .biSize = Len(bmi.bmiHeader)
            .biWidth = WidthPx
            .biHeight = HeightPx
            .biPlanes = 1
            .biCompression = BI_RGB
            .biBitCount = 24
        End With
        Stride = ((3 * WidthPx + 3) \ 4) * 4
        ReDim Triples(Stride * HeightPx - 1)
        .AutoRedraw = True
        GetDIBits .hDC, .Image.Handle, 0, HeightPx, Triples(0), bmi, DIB_RGB_COLORS
        For Line = 0 To (HeightPx - 1) * Stride Step Stride
            For Pixel = Line To Line + (WidthPx - 1) * 3 Step 3
                If Triples(Pixel + 2) = ColorR And _
                   Triples(Pixel + 1) = ColorG And _
                   Triples(Pixel) = ColorB Then
                        Triples(Pixel + 2) = NewR
                        Triples(Pixel + 1) = NewG
                        Triples(Pixel) = NewB
                End If
            Next
        Next
        SetDIBits .hDC, .Image.Handle, 0, HeightPx, Triples(0), bmi, DIB_RGB_COLORS
        .AutoRedraw = False
    End With

    ColorR = NewR
    ColorG = NewG
    ColorB = NewB

    PropertyChanged "Color"
End Property

Private Sub SetUpDefaultColor()
    Dim RGBQUAD As Long
    'These must match the parts of the original Picture bitmap that we recolorize via our
    'Color property, i.e. they are the default Color:
    mColor = DEFAULT_COLOR
    If mColor And &H80000000 Then
        RGBQUAD = GetSysColor(mColor And &HFFFF&)
    Else
        RGBQUAD = mColor
    End If
    ColorR = RGBQUAD And &HFF&
    ColorG = (RGBQUAD And &HFF00&) \ &H100&
    ColorB = RGBQUAD \ &H10000
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_InitProperties()
    SetUpDefaultColor
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim NewText As String
    
    If Button = vbLeftButton Then
        MousePointer = vbCustom
    ElseIf Button = vbRightButton Then
        With Extender
            'This could be fancier, but we'll just use InputBox() for this:
            NewText = InputBox("Edit toolltip text", _
                               , _
                               .ToolTipText, _
                               .Parent.Left + X, _
                               .Parent.Top + Y)
            If StrPtr(NewText) <> 0 Then .ToolTipText = NewText
        End With
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        With Extender
            .Move .Left + X - .Width / 2, .Top + Y - .Height / 2
        End With
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbDefault
End Sub

Private Sub UserControl_Paint()
    Size 480, 480
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SetUpDefaultColor 'We need to do this here to assign the initial ColorR, ColorG, and ColorB values.
    With PropBag
        Color = .ReadProperty("Color", DEFAULT_COLOR)
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Color", Color, DEFAULT_COLOR
    End With
End Sub
