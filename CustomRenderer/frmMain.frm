VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Custom renderer and effects"
   ClientHeight    =   6285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cD2dFactory As ID2D1Factory
Dim cTarget     As ID2D1HwndRenderTarget
Dim cDWFactory  As IDWriteFactory
Dim cTextFormat As IDWriteTextFormat
Dim cTextlayout As IDWriteTextLayout
Dim sText       As String
Dim cCustRender As CRenderer

Private Sub Form_Load()
    
    sText = "Custom renderer example Visual Basic 6"
    
    Set cD2dFactory = D2D1.CreateFactory
    
    Set cTarget = cD2dFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties( _
                                    D2D1.PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM)), _
                                    D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU))
                                    
    Set cDWFactory = DW.CreateFactory(DWRITE_FACTORY_TYPE_SHARED)
    Set cTextFormat = cDWFactory.CreateTextFormat("Arial", Nothing, DWRITE_FONT_WEIGHT_NORMAL, _
                                    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, _
                                    32# * 96# / 72#, "en-US")
                                    
    Set cTextlayout = cDWFactory.CreateTextLayout(sText, Len(sText), cTextFormat, Me.ScaleWidth, Me.ScaleHeight)
    
    cTextlayout.SetFontFamilyName "Courier New", 0, 6
    cTextlayout.SetFontFamilyName "Impact", 7, 8
    
    cTextlayout.SetFontStyle DWRITE_FONT_STYLE_ITALIC, 24, 6
    cTextlayout.SetFontWeight DWRITE_FONT_WEIGHT_BOLD, 31, 5
    cTextlayout.SetParagraphAlignment DWRITE_PARAGRAPH_ALIGNMENT_CENTER
    cTextlayout.SetTextAlignment DWRITE_TEXT_ALIGNMENT_CENTER
    
    cTextlayout.SetStrikethrough 1, 0, 6
    cTextlayout.SetUnderline 1, 31, 5
    
    Set cCustRender = New CRenderer
    
    cCustRender.Initialize cTarget, cD2dFactory
                           
    ' // Create effect
    Dim cEffect As CColorEffect
    
    Set cEffect = New CColorEffect
    
    cEffect.Color = D2D1_COLORS.Green
    
    cTextlayout.SetDrawingEffect cEffect, 0, 6
                        
    Set cEffect = New CColorEffect
    
    cEffect.Color = D2D1_COLORS.HotPink
    
    cTextlayout.SetDrawingEffect cEffect, 24, 6
    
End Sub

Private Sub Form_Paint()
    
    cTarget.BeginDraw
    
    cTarget.Clear D2D1.ColorF(Ivory)
    
    cTextlayout.Draw ByVal 0&, cCustRender, 10, 10
    
    cTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub Form_Resize()
    
    cTextlayout.SetMaxWidth Me.ScaleWidth
    cTextlayout.SetMaxHeight Me.ScaleHeight
    
    cTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
    
    Form_Paint
    
End Sub
