VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct2D perspective transform effect"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldAngle 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   767
      _Version        =   327682
      Max             =   3600
      TickFrequency   =   100
   End
   Begin VB.PictureBox picResult 
      BorderStyle     =   0  'None
      Height          =   5880
      Left            =   60
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   0
      Top             =   60
      Width           =   7830
   End
   Begin ComctlLib.Slider sldOffset 
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   6480
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   767
      _Version        =   327682
      Min             =   -1000
      Max             =   1000
      TickFrequency   =   100
   End
   Begin ComctlLib.Slider sldDepth 
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   6960
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   767
      _Version        =   327682
      Min             =   1
      Max             =   10000
      SelStart        =   1000
      TickFrequency   =   100
      Value           =   1000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Direct2D 3d perspective effect demo by The trick
' //

Option Explicit

Private Const GENERIC_READ As Long = &H80000000

Private m_cFactory      As ID2D1Factory
Private m_cWICFactory   As WICImagingFactory
Private m_cRenderTarget As ID2D1DeviceContext
Private m_cEffect       As ID2D1Effect

Private Sub Form_Load()
    Dim cBitmap     As ID2D1Bitmap
    Dim cHwndTgt    As ID2D1HwndRenderTarget

    ' // Create a factory
    Set m_cFactory = D2D1.CreateFactory()
    
    ' // Create render target
    Set m_cRenderTarget = m_cFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties(D2D1.PixelFormat()), _
                                                            D2D1.HwndRenderTargetProperties(picResult.hWnd, D2D1.SizeU()))
    
    Set cHwndTgt = m_cRenderTarget
    
    cHwndTgt.Resize D2D1.SizeU(picResult.ScaleWidth, picResult.ScaleHeight)

    Set cBitmap = LoadImage(App.Path & "\image_01.jpg", m_cRenderTarget)

    ' // Create blur effect
    Set m_cEffect = m_cRenderTarget.CreateEffect(CLSID_D2D13DPerspectiveTransform)

    m_cEffect.SetInput 0, cBitmap
    m_cEffect.SetValue D2D1_3DPERSPECTIVETRANSFORM_PROP_PERSPECTIVE_ORIGIN, D2D1_PROPERTY_TYPE_UNKNOWN, _
                       D2D1.Vector2F(0, cBitmap.GetSize.Height / 2), 8
    m_cEffect.SetValue D2D1_3DPERSPECTIVETRANSFORM_PROP_ROTATION_ORIGIN, D2D1_PROPERTY_TYPE_UNKNOWN, _
                       D2D1.Vector3F(cBitmap.GetSize.Width / 2, 0, 0), 12

    sldAngle_Change
    
End Sub

Private Function LoadImage( _
                 ByRef sPath As String, _
                 ByVal cRenderTarget As ID2D1RenderTarget) As ID2D1Bitmap
    Dim cDecoder        As IWICBitmapDecoder
    Dim cFrameDecode    As IWICBitmapFrameDecode
    Dim cConverter      As IWICFormatConverter
    
    If m_cWICFactory Is Nothing Then
        Set m_cWICFactory = New WICImagingFactory
    End If
    
    Set cDecoder = m_cWICFactory.CreateDecoderFromFilename(sPath, ByVal 0&, GENERIC_READ, WICDecodeMetadataCacheOnLoad)
    Set cFrameDecode = cDecoder.GetFrame(0)
    Set cConverter = m_cWICFactory.CreateFormatConverter()
    
    cConverter.Initialize cFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteTypeMedianCut
    
    Set LoadImage = cRenderTarget.CreateBitmapFromWicBitmap(ByVal cConverter, ByVal 0&)
    
End Function

Private Sub picResult_Paint()

    m_cRenderTarget.BeginDraw
    
    m_cRenderTarget.Clear D2D1.ColorF(Indigo)

    m_cRenderTarget.DrawImage m_cEffect, ByVal 0&, ByVal 0&
    
    m_cRenderTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub sldAngle_Change()
    m_cEffect.SetValue D2D1_3DPERSPECTIVETRANSFORM_PROP_ROTATION, D2D1_PROPERTY_TYPE_UNKNOWN, D2D1.Vector3F(0, sldAngle.Value / 10, 0), 12
    picResult_Paint
End Sub

Private Sub sldAngle_Scroll()
    sldAngle_Change
End Sub

Private Sub sldOffset_Change()
    m_cEffect.SetValue D2D1_3DPERSPECTIVETRANSFORM_PROP_GLOBAL_OFFSET, D2D1_PROPERTY_TYPE_UNKNOWN, _
                       D2D1.Vector3F(-sldOffset.Value / 4, 0, sldOffset.Value), 12
    picResult_Paint
End Sub

Private Sub sldOffset_Scroll()
    sldOffset_Change
End Sub

Private Sub sldDepth_Change()
    m_cEffect.SetValue D2D1_3DPERSPECTIVETRANSFORM_PROP_DEPTH, D2D1_PROPERTY_TYPE_FLOAT, CSng(sldDepth), 4
    picResult_Paint
End Sub

Private Sub sldDepth_Scroll()
    sldDepth_Change
End Sub


