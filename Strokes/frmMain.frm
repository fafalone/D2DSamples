VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Strokes demo"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   370
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // Stroke demo
' // Background sprite from Battletoads & Double Dragon (NES)

Option Explicit

Dim cFactory        As ID2D1Factory
Dim cRenderTarget   As ID2D1HwndRenderTarget
Dim cSolidBrush     As ID2D1SolidColorBrush

Private Sub Form_Load()
    
    ' // Create a factory
    Set cFactory = D2D1.CreateFactory()
    
    ' // Create render target
    Set cRenderTarget = cFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties(D2D1.PixelFormat()), _
                                                        D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU()))
                                                        
    cRenderTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
    
End Sub


Private Sub Form_Paint()
    Dim cBrush  As ID2D1SolidColorBrush
    Dim cStroke As ID2D1StrokeStyle
    Dim fDash() As Single
    
    cRenderTarget.BeginDraw
    
    Set cBrush = cRenderTarget.CreateSolidColorBrush(D2D1.ColorF(Wheat), ByVal 0&)
    
    FillBackground
    
    ' // Default stroke
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(10, 50)
    DrawLines cBrush, Nothing
    
    ' // Dash rounded begin and end
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_ROUND, _
                    D2D1_CAP_STYLE_ROUND, D2D1_CAP_STYLE_FLAT, D2D1_LINE_JOIN_MITER, , _
                    D2D1_DASH_STYLE_DASH), ByVal 0&, 0)
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(120, 50)
    DrawLines cBrush, cStroke
    
    ' // Dash rounded all
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_ROUND, _
                    D2D1_CAP_STYLE_ROUND, D2D1_CAP_STYLE_ROUND, D2D1_LINE_JOIN_MITER, , _
                    D2D1_DASH_STYLE_DASH), ByVal 0&, 0)
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(230, 50)
    DrawLines cBrush, cStroke
    
    ' // Dash triangled caps
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_ROUND, _
                    D2D1_CAP_STYLE_ROUND, D2D1_CAP_STYLE_TRIANGLE, D2D1_LINE_JOIN_MITER, , _
                    D2D1_DASH_STYLE_DASH), ByVal 0&, 0)
    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(340, 50)
    DrawLines cBrush, cStroke
    
    ' // Dash beveled joins
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_ROUND, _
                    D2D1_CAP_STYLE_ROUND, D2D1_CAP_STYLE_FLAT, D2D1_LINE_JOIN_BEVEL, , _
                    D2D1_DASH_STYLE_SOLID), ByVal 0&, 0)
                    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(10, 150)
    DrawLines cBrush, cStroke
    
    ' // Dash rounded joins
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_ROUND, _
                    D2D1_CAP_STYLE_ROUND, D2D1_CAP_STYLE_FLAT, D2D1_LINE_JOIN_ROUND, , _
                    D2D1_DASH_STYLE_SOLID), ByVal 0&, 0)
                    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(120, 150)
    DrawLines cBrush, cStroke
    
    ' // Custom dashes
    ReDim fDash(5)
    
    fDash(0) = 2    ' // Dash length
    fDash(1) = 1    ' // Space length
    fDash(2) = 0.5  ' // Dash length
    fDash(3) = 0.5  ' // ...
    fDash(4) = 1
    fDash(5) = 1
    
    Set cStroke = cFactory.CreateStrokeStyle(D2D1.StrokeStyleProperties(D2D1_CAP_STYLE_FLAT, _
                    D2D1_CAP_STYLE_FLAT, D2D1_CAP_STYLE_FLAT, D2D1_LINE_JOIN_MITER, , _
                    D2D1_DASH_STYLE_CUSTOM), fDash(0), UBound(fDash) + 1)
                    
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Translation2(230, 150)
    DrawLines cBrush, cStroke
    
    cRenderTarget.EndDraw ByVal 0&, ByVal 0&

    
End Sub

Private Sub DrawLines( _
            ByVal cBrush As ID2D1Brush, _
            ByVal cStroke As ID2D1StrokeStyle)
    Dim cPath   As ID2D1PathGeometry
    Dim cSink   As ID2D1GeometrySink
    
    Set cPath = cFactory.CreatePathGeometry
    
    Set cSink = cPath.Open
    
    cSink.BeginFigure 0, 0, D2D1_FIGURE_BEGIN_HOLLOW
    
    cSink.AddLine 20, 30
    cSink.AddLine 30, -30
    cSink.AddLine 50, 40
    cSink.AddLine 60, -20
    cSink.AddLine 80, 0
    
    cSink.EndFigure D2D1_FIGURE_END_OPEN
    
    cSink.Close
    
    cRenderTarget.DrawGeometry cPath, cBrush, 7, cStroke

End Sub

Private Sub FillBackground()
    Dim cBitmap     As ID2D1Bitmap
    Dim cBrush      As ID2D1Brush
    Dim bPixels()   As Byte
    
    bPixels = LoadResData(101, "CUSTOM")
    
    ' // Create bitmap based on pixels data
    Set cBitmap = cRenderTarget.CreateBitmap(36, 32, bPixels(0), 36 * 4, D2D1.BitmapProperties(D2D1.PixelFormat( _
                 DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE), 96, 96))
    
    Set cBrush = cRenderTarget.CreateBitmapBrush(cBitmap, D2D1.BitmapBrushProperties(D2D1_EXTEND_MODE_WRAP, _
                            D2D1_EXTEND_MODE_WRAP, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR), D2D1.BrushProperties( _
                            0.75, D2D1.Matrix3x2F_Rotation(45, D2D1.Point2F(0, 0))))
                           
    cRenderTarget.Clear D2D1.ColorF(SaddleBrown)
    cRenderTarget.FillRectangle D2D1.RectF(0, 0, cRenderTarget.GetSize.Width, cRenderTarget.GetSize.Height), cBrush
    
End Sub


