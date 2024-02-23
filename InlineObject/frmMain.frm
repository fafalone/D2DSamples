VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Inline object test"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   482
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' // DirectWrite inline object implementation example

Option Explicit

Dim cD2dFactory As ID2D1Factory
Dim cDWFactory  As IDWriteFactory
Dim cTarget     As ID2D1HwndRenderTarget
Dim cTextFormat As IDWriteTextFormat
Dim cTextlayout As IDWriteTextLayout
Dim cBrush      As ID2D1Brush

Private sText   As String

Private Sub Form_Load()
    Dim cSmile  As CInlineObject    ' // Object inside line
    
    sText = "DirectWrite _ Visual Basic 6 example"
    
    Set cD2dFactory = D2D1.CreateFactory
    
    Set cTarget = cD2dFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties( _
                                    D2D1.PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM)), _
                                    D2D1.HwndRenderTargetProperties(Me.hWnd, D2D1.SizeU))
    
    Set cDWFactory = DW.CreateFactory(DWRITE_FACTORY_TYPE_SHARED)
    Set cTextFormat = cDWFactory.CreateTextFormat("Arial", Nothing, DWRITE_FONT_WEIGHT_NORMAL, _
                                    DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL, _
                                    19# * 96# / 72#, "en-US")
                                    
    Set cTextlayout = cDWFactory.CreateTextLayout(sText, Len(sText), cTextFormat, Me.ScaleWidth, Me.ScaleHeight)
    
    ' // Set different options
    cTextlayout.SetFontWeight DWRITE_FONT_WEIGHT_BOLD, 14, 6
 
    cTextlayout.SetFontSize 11 * 96 / 72, 21, 5

    cTextlayout.SetFontStyle DWRITE_FONT_STYLE_ITALIC, 0, 6
    
    cTextlayout.SetStrikethrough 1, 29, 7
    
    cTextlayout.SetTextAlignment DWRITE_TEXT_ALIGNMENT_CENTER
    cTextlayout.SetParagraphAlignment DWRITE_PARAGRAPH_ALIGNMENT_CENTER
    
    cTextlayout.SetUnderline 1, 6, 5
    
    Set cBrush = cTarget.CreateSolidColorBrush(D2D1.ColorF(Red), ByVal 0&)
                                                                    
    ' // Add inline object
    Set cSmile = New CInlineObject
    
    cSmile.Initialize cTarget, App.Path & "\icon.png"
    
    cTextlayout.SetInlineObject cSmile, 12, 1
    
End Sub

Private Sub Form_Paint()
    
    cTarget.BeginDraw
    
    cTarget.Clear D2D1.ColorF(Ivory)
    
    cTarget.DrawTextLayout 0, 0, ByVal cTextlayout, cBrush
    
    cTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub Form_Resize()
    
    cTextlayout.SetMaxWidth Me.ScaleWidth
    cTextlayout.SetMaxHeight Me.ScaleHeight
    
    cTarget.Resize D2D1.SizeU(Me.ScaleWidth, Me.ScaleHeight)
    
    Form_Paint
    
End Sub


