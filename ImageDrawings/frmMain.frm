VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image drawing"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2760
      Top             =   4800
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // Direct2D basic drawings image demo by The trick
' //

Option Explicit

Private Const GENERIC_READ As Long = &H80000000

Dim cFactory        As ID2D1Factory
Dim cWICFactory     As WICImagingFactory
Dim cRenderTarget   As ID2D1HwndRenderTarget
Dim cBitmap         As ID2D1Bitmap
Dim cMemBitmap      As ID2D1Bitmap

Private Sub Form_Load()
    Dim cDecoder        As IWICBitmapDecoder
    Dim cFrameDecode    As IWICBitmapFrameDecode
    Dim cConverter      As IWICFormatConverter
    Dim cMemTarget      As ID2D1BitmapRenderTarget
    Dim cBrush          As ID2D1SolidColorBrush
    Dim tPos            As D2D1_POINT_2F
    
    ' // Load bitmap
    Set cWICFactory = New WICImagingFactory
    Set cDecoder = cWICFactory.CreateDecoderFromFilename(App.Path & "/image.png", ByVal 0&, GENERIC_READ, WICDecodeMetadataCacheOnLoad)
    Set cFrameDecode = cDecoder.GetFrame(0)
    Set cConverter = cWICFactory.CreateFormatConverter()
    
    cConverter.Initialize cFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteTypeMedianCut
    
    ' // Create a factory
    Set cFactory = D2D1.CreateFactory()

    ' // Create render target
    Set cRenderTarget = cFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties(D2D1.PixelFormat()), _
                                                        D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU()))

    Set cBitmap = cRenderTarget.CreateBitmapFromWicBitmap(ByVal cConverter, ByVal 0&)
    
    ' // Create bitmap from memory
    ' // We'll draw into this bitmap using Direct2D
    Set cMemTarget = cRenderTarget.CreateCompatibleRenderTarget(D2D1.SizeF(100, 100), ByVal 0&, _
                      ByVal 0&, D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE)
                     
    ' // Draw circles
    Set cBrush = cMemTarget.CreateSolidColorBrush(D2D1.ColorF(Brown), ByVal 0&)

    cMemTarget.BeginDraw
    
    For tPos.x = 0 To 100 Step 20
        For tPos.y = 0 To 100 Step 20
            
            cBrush.SetColor D2D1.ColorF(Rnd * &H1000000, Rnd)
            cMemTarget.DrawEllipse D2D1.Ellipse(tPos, 10, 10), cBrush, Rnd * 4 + 1
            
        Next
    Next
    
    cMemTarget.EndDraw ByVal 0&, ByVal 0&
    
    Set cMemBitmap = cMemTarget.GetBitmap()
    
End Sub

Private Sub Form_Resize()
    ' // Drawing area has been changed
    cRenderTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
End Sub

Private Sub Timer1_Timer()
    Static fPhase As Single
    Dim tSize As D2D1_SIZE_F
    
    fPhase = fPhase + 1
    
    tSize = cBitmap.GetSize
    
    cRenderTarget.BeginDraw
    
    cRenderTarget.Clear D2D1.ColorF(Indigo)
    
    ' // Draw using different interpolations
    
    ' // Rotate image around center and move it on 100x100 pixels
    cRenderTarget.SetTransform D2D1.Matrix3x2F_SetProduct( _
                               D2D1.Matrix3x2F_Rotation2(fPhase, tSize.Width / 2, tSize.Height / 2), _
                               D2D1.Matrix3x2F_Translation2(30, 30))
    
    ' // Rough
    cRenderTarget.DrawBitmap cBitmap, ByVal 0&, , D2D1_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR, ByVal 0&
     
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_SetProduct( _
                               D2D1.Matrix3x2F_Rotation2(fPhase, tSize.Width / 2, tSize.Height / 2), _
                               D2D1.Matrix3x2F_Translation2(150, 30))
                               
    ' // Linear
    cRenderTarget.DrawBitmap cBitmap, ByVal 0&, , D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, ByVal 0&
    
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Identity
                               
    ' // Draw part
    cRenderTarget.DrawBitmap cBitmap, D2D1.RectF(30, 155, 120, 220), , D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, _
                              D2D1.RectF(25, 25, 50, 50)
    
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(150, 155)

    ' // Draw mem-bitmap
    cRenderTarget.DrawBitmap cMemBitmap, ByVal 0&, , D2D1_BITMAP_INTERPOLATION_MODE_NEAREST_NEIGHBOR, ByVal 0&
    
    ' // Stretch bitmap
    cRenderTarget.SetTransform D2D1.Matrix3x2F_SetProduct( _
                               D2D1.Matrix3x2F_Scale2(2.3, 2, D2D1.Point2F(0, 0)), _
                               D2D1.Matrix3x2F_Translation2(250, 30))
    
    cRenderTarget.DrawBitmap cMemBitmap, ByVal 0&, , D2D1_BITMAP_INTERPOLATION_MODE_LINEAR, ByVal 0&
    
    cRenderTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub
