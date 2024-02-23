VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "DirectWrite custom font demo"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' // DirectWrite custom font example
' // This example contains two fonts in resources
' // Application loads that fonts and draw text using that fonts
' // By The trick 2018

Option Explicit

Dim cD2dFactory     As ID2D1Factory
Dim cDWFactory      As IDWriteFactory
Dim cTarget         As ID2D1HwndRenderTarget
Dim cTextFormat1    As IDWriteTextFormat
Dim cTextFormat2    As IDWriteTextFormat
Dim cBrush          As ID2D1Brush
Dim cFontLoader     As CFontCollectionLoader

Private Sub Form_Load()
    Dim cFontCollection As IDWriteFontCollection
    
    Set cD2dFactory = D2D1.CreateFactory
    
    Set cTarget = cD2dFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties( _
                                    D2D1.PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM)), _
                                    D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU))
    
    Set cDWFactory = DW.CreateFactory(DWRITE_FACTORY_TYPE_SHARED)
    
    ' // Create font collection
    Set cFontLoader = New CFontCollectionLoader
    
    ' // Register our font loader
    cDWFactory.RegisterFontCollectionLoader cFontLoader
    
    Set cFontCollection = cDWFactory.CreateCustomFontCollection(cFontLoader, 0, 1)

    ' // Create text formats using custom fonts
    Set cTextFormat1 = cDWFactory.CreateTextFormat("pix 8BitFont_8pix", cFontCollection, DWRITE_FONT_WEIGHT_NORMAL, _
                                    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, _
                                    16, "en-us")
    Set cTextFormat2 = cDWFactory.CreateTextFormat("pix 8 8pt", cFontCollection, DWRITE_FONT_WEIGHT_NORMAL, _
                                    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, _
                                    24, "en-us")
    ' // Color brush
    Set cBrush = cTarget.CreateSolidColorBrush(D2D1.ColorF(SkyBlue), ByVal 0&)
    
    cTextFormat1.SetTextAlignment DWRITE_TEXT_ALIGNMENT_CENTER
    cTextFormat2.SetTextAlignment DWRITE_TEXT_ALIGNMENT_CENTER
    
End Sub

Private Sub Form_Paint()
    Dim sText   As String
    
    sText = "Custom 8 bit FONT!"
    
    cTarget.BeginDraw
    
    cTarget.Clear D2D1.ColorF(Navy)
    
    cTarget.DrawText sText, Len(sText), ByVal cTextFormat1, D2D1.RectF(0, 20, ScaleWidth, 120), cBrush
    
    sText = "C u s t o m  8  b i t  F O N T !"
    
    cTarget.DrawText sText, Len(sText), ByVal cTextFormat2, D2D1.RectF(0, 80, ScaleWidth, 120), cBrush
    
    cTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub Form_Resize()
    cTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
    Form_Paint
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cDWFactory.UnregisterFontCollectionLoader cFontLoader
End Sub
