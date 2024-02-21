VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct2D blur effect"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldValue 
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   6000
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   767
      _Version        =   327682
      Max             =   200
      TickFrequency   =   10
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Direct2D blur demo by The trick
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
    Set m_cEffect = m_cRenderTarget.CreateEffect(CLSID_D2D1GaussianBlur)

    m_cEffect.SetInput 0, cBitmap
    m_cEffect.SetValue D2D1_GAUSSIANBLUR_PROP_BORDER_MODE, D2D1_PROPERTY_TYPE_UNKNOWN, D2D1_BORDER_MODE_HARD, 4
    
    sldValue_Change
    
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

Private Sub sldValue_Change()
    m_cEffect.SetValue D2D1_GAUSSIANBLUR_PROP_STANDARD_DEVIATION, D2D1_PROPERTY_TYPE_FLOAT, CSng(sldValue.Value / 10), 4
    picResult_Paint
End Sub

Private Sub sldValue_Scroll()
    sldValue_Change
End Sub
