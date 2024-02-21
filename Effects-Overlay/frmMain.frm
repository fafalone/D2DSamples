VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct2D blending modes"
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
   Begin VB.ComboBox cboMode 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   60
      List            =   "frmMain.frx":0052
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   6000
      Width           =   7875
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
' // Direct2D blending modes demo by The trick
' //

Option Explicit

Private Const GENERIC_READ As Long = &H80000000

Private m_cFactory      As ID2D1Factory
Private m_cWICFactory   As WICImagingFactory
Private m_cRenderTarget As ID2D1DeviceContext
Private m_cEffect       As ID2D1Effect

Private Sub cboMode_Click()
    m_cEffect.SetValue D2D1_BLEND_PROP_MODE, D2D1_PROPERTY_TYPE_UNKNOWN, CLng(cboMode.ListIndex), 4
    picResult.Refresh
End Sub

Private Sub Form_Load()
    Dim cBitmaps()  As ID2D1Bitmap
    Dim cHwndTgt    As ID2D1HwndRenderTarget

    ' // Create a factory
    Set m_cFactory = D2D1.CreateFactory()

    ' // Create render target
    Set m_cRenderTarget = m_cFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties(D2D1.PixelFormat()), _
                                                            D2D1.HwndRenderTargetProperties(picResult.hWnd, D2D1.SizeU()))
    
    Set cHwndTgt = m_cRenderTarget
    
    cHwndTgt.Resize D2D1.SizeU(picResult.ScaleWidth, picResult.ScaleHeight)
    
    ReDim cBitmaps(1)
    
    Set cBitmaps(0) = LoadImage(App.Path & "\image_01.jpg", m_cRenderTarget)
    Set cBitmaps(1) = LoadImage(App.Path & "\image_02.png", m_cRenderTarget)

    ' // Create blur effect
    Set m_cEffect = m_cRenderTarget.CreateEffect(CLSID_D2D1Blend)

    m_cEffect.SetInput 0, cBitmaps(0)
    m_cEffect.SetInput 1, cBitmaps(1)
    
    cboMode.ListIndex = 0
    
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
