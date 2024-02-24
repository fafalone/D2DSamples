VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clip demo"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   397
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // Clipping demo

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
    Dim cPath   As ID2D1PathGeometry
    Dim cSink   As ID2D1GeometrySink
    Dim cBrush  As ID2D1SolidColorBrush
    Dim cLayer  As ID2D1Layer
    
    cRenderTarget.BeginDraw
    
    FillBackground
    
    Set cBrush = cRenderTarget.CreateSolidColorBrush(D2D1.ColorF(MintCream), ByVal 0&)
    
    Set cPath = cFactory.CreatePathGeometry
    
    Set cSink = cPath.Open
    
    cSink.BeginFigure 100, 50, D2D1_FIGURE_BEGIN_FILLED
    
    cSink.AddLine 200, 300
    cSink.AddArc D2D1.ArcSegment(D2D1.Point2F(250, 200), D2D1.SizeF(50, 30), 0, _
                 D2D1_SWEEP_DIRECTION_COUNTER_CLOCKWISE, D2D1_ARC_SIZE_SMALL)
    cSink.AddQuadraticBezier D2D1.QuadraticBezierSegment(D2D1.Point2F(0, 150), D2D1.Point2F(200, 100))
    cSink.AddBezier D2D1.BezierSegment(D2D1.Point2F(300, 0), D2D1.Point2F(150, 0), D2D1.Point2F(10, 20))
    cSink.EndFigure D2D1_FIGURE_END_CLOSED
    
    cSink.Close
    
    cRenderTarget.DrawGeometry cPath, cBrush, 3
    
    Set cLayer = cRenderTarget.CreateLayer(ByVal 0&)
    
    cRenderTarget.PushLayer D2D1.LayerParameters(D2D1.RectF_InfiniteRect, cPath, D2D1_ANTIALIAS_MODE_PER_PRIMITIVE, _
                            D2D1.Matrix3x2F_Identity), cLayer
    
    ' // Draw ellipse
    cBrush.SetColor D2D1.ColorF(MediumSpringGreen)
    
    cRenderTarget.FillEllipse D2D1.Ellipse(D2D1.Point2F(150, 150), 100, 100), cBrush
    
    cRenderTarget.PopLayer
    
    cRenderTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub FillBackground()
    Dim cMemSurf    As ID2D1BitmapRenderTarget
    Dim cBrush      As ID2D1Brush
    
    Set cMemSurf = cRenderTarget.CreateCompatibleRenderTarget(D2D1.SizeF(10, 10), ByVal 0&, ByVal 0&, _
                                D2D1_COMPATIBLE_RENDER_TARGET_OPTIONS_NONE)
    Set cBrush = cRenderTarget.CreateSolidColorBrush(D2D1.ColorF(MintCream), ByVal 0&)
    
    cMemSurf.BeginDraw
    
    cMemSurf.DrawLine 2, 2, 8, 8, cBrush, 5
    
    cMemSurf.EndDraw ByVal 0&, ByVal 0&
    
    Set cBrush = cMemSurf.CreateBitmapBrush(cMemSurf.GetBitmap, D2D1.BitmapBrushProperties(D2D1_EXTEND_MODE_MIRROR, _
                            D2D1_EXTEND_MODE_MIRROR, D2D1_BITMAP_INTERPOLATION_MODE_LINEAR), D2D1.BrushProperties( _
                            0.15, D2D1.Matrix3x2F_Rotation(45, D2D1.Point2F(0, 0))))
                           
    cRenderTarget.Clear D2D1.ColorF(MediumBlue)
    cRenderTarget.FillRectangle D2D1.RectF(0, 0, cRenderTarget.GetSize.Width, cRenderTarget.GetSize.Height), cBrush
    
End Sub
