VERSION 5.00
Begin VB.Form frmBasicDrawing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct2D Basic drawing by The trick"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBasicDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' //
' // Direct2D basic drawings demo by The trick
' //

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
                                                        
    ' // Create solid brush for strokes
    Set cSolidBrush = cRenderTarget.CreateSolidColorBrush(D2D1.ColorF(D2D1_COLORS.Red), ByVal 0&)
    
End Sub

Private Sub Form_Paint()
 
    cRenderTarget.BeginDraw
    
    ' // Set color
    cSolidBrush.SetColor D2D1.ColorF(D2D1_COLORS.BlueViolet)
    
    ' // Reset transform
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Identity()
    
    ' // Clear background
    cRenderTarget.Clear D2D1.ColorF(D2D1_COLORS.AliceBlue)
    
    ' // Draw simple rectangle
    cRenderTarget.DrawRectangle D2D1.RectF(5, 5, 90, 140), cSolidBrush, 1
    
    ' // Rotate system by 60 degrees and offset by 250 pix
    cRenderTarget.SetTransform D2D1.Matrix3x2F_SetProduct( _
                               D2D1.Matrix3x2F_Rotation2(60, 0, 0), _
                               D2D1.Matrix3x2F_Translation2(250, 0))
    
    ' // Draw simple rectangle
    cRenderTarget.FillRectangle D2D1.RectF(5, 5, 90, 140), cSolidBrush
    
    ' // Reset transform
    cRenderTarget.SetTransform D2D1.Matrix3x2F_Identity()
    
    ' // Change brush color
    cSolidBrush.SetColor D2D1.ColorF(D2D1_COLORS.ForestGreen)
    
    ' // Draw ellipse by 5 pix stroke
    cRenderTarget.DrawEllipse D2D1.Ellipse(D2D1.Point2F(100, 250), 90, 50), cSolidBrush, 5
    
    ' // Change brush color
    cSolidBrush.SetColor D2D1.ColorF(D2D1_COLORS.Coral)
    
    ' // Change brush opacity
    cSolidBrush.SetOpacity 0.5
    
    ' // Fill inside
    cRenderTarget.FillEllipse D2D1.Ellipse(D2D1.Point2F(100, 250), 90, 50), cSolidBrush
    
    cSolidBrush.SetOpacity 1
    
    ' // Draw rounded rect
    cRenderTarget.DrawRoundedRectangle D2D1.RoundedRect(D2D1.RectF(350, 5, 450, 200), 30, 50), cSolidBrush, 2
    
    ' // Fill rounded rect
    cRenderTarget.FillRoundedRectangle D2D1.RoundedRect(D2D1.RectF(250, 250, 450, 300), 30, 50), cSolidBrush
    
    ' // Draw line
    cRenderTarget.DrawLine 5, 20, 300, 300, cSolidBrush, 4
    
    cRenderTarget.EndDraw ByVal 0&, ByVal 0&

End Sub

Private Sub Form_Resize()

    ' // Drawing area has been changed
    cRenderTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
    
End Sub
