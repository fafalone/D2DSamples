VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectWrite"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // DirectWrite basic drawing example

Option Explicit

Dim cD2dFactory As ID2D1Factory
Dim cDWFactory  As IDWriteFactory
Dim cTarget     As ID2D1HwndRenderTarget
Dim cTextFormat As IDWriteTextFormat
Dim cBrush      As ID2D1Brush

Private Sub Form_Load()
    
    Set cD2dFactory = D2D1.CreateFactory
    
    Set cTarget = cD2dFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties( _
                                    D2D1.PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM)), _
                                    D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU))
    
    Set cDWFactory = DW.CreateFactory(DWRITE_FACTORY_TYPE_SHARED)
    
    ' // Create text format
    Set cTextFormat = cDWFactory.CreateTextFormat("Arial", Nothing, DWRITE_FONT_WEIGHT_NORMAL, _
                                    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, _
                                    19# * 96# / 72#, "en-US")
                                    
    ' // Color brush
    Set cBrush = cTarget.CreateSolidColorBrush(D2D1.ColorF(Red), ByVal 0&)
                                    
End Sub

Private Sub Form_Paint()
    Dim sText   As String
    
    sText = "Hello World!"
    
    cTarget.BeginDraw
    
    cTarget.Clear D2D1.ColorF(Ivory)
    
    cTarget.DrawText sText, Len(sText), ByVal cTextFormat, D2D1.RectF(20, 20, 220, 120), cBrush

    cTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub Form_Resize()
    cTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
End Sub
