VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function vbaObjSetAddref Lib "MSVBVM60.DLL" _
                         Alias "__vbaObjSetAddref" ( _
                         ByRef dstObject As Any, _
                         ByRef srcObjPtr As Any) As Long
                         
Private Const GENERIC_READ As Long = &H80000000

Implements IDWriteTextRendererImp

Dim mcSolidBrush    As ID2D1SolidColorBrush
Dim mcBitmapBrush   As ID2D1BitmapBrush
Dim mcRenderTarget  As ID2D1RenderTarget
Dim mcD2DFactory    As ID2D1Factory

Public Sub Initialize( _
           ByVal cRenderTarget As ID2D1RenderTarget, _
           ByVal cD2dFactory As ID2D1Factory)
    Dim cFactory        As WICImagingFactory
    Dim cDecoder        As IWICBitmapDecoder
    Dim cFrameDecode    As IWICBitmapFrameDecode
    Dim cConverter      As IWICFormatConverter
    Dim cBitmap         As ID2D1Bitmap
    
    Set cFactory = New WICImagingFactory
    Set cDecoder = cFactory.CreateDecoderFromFilename(App.Path & "\texture.jpg", ByVal 0&, _
                                                      GENERIC_READ, WICDecodeMetadataCacheOnLoad)
    Set cFrameDecode = cDecoder.GetFrame(0)
    Set cConverter = cFactory.CreateFormatConverter()
    
    cConverter.Initialize cFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, _
                          Nothing, 0, WICBitmapPaletteTypeMedianCut
                            
    Set cBitmap = cRenderTarget.CreateBitmapFromWicBitmap(ByVal cConverter, ByVal 0&)
    
    Set mcBitmapBrush = cRenderTarget.CreateBitmapBrush(cBitmap, D2D1.BitmapBrushProperties(D2D1_EXTEND_MODE_WRAP, _
                                                        D2D1_EXTEND_MODE_WRAP), ByVal 0&)
    Set mcSolidBrush = cRenderTarget.CreateSolidColorBrush(D2D1.ColorF(Black), ByVal 0&)

    Set mcRenderTarget = cRenderTarget
    Set mcD2DFactory = cD2dFactory
    
End Sub

Private Sub IDWriteTextRendererImp_DrawGlyphRun( _
            ByVal pClientDrawingContext As Long, _
            ByVal baselineOriginX As Single, _
            ByVal baselineOriginY As Single, _
            ByVal measuringMode As D2DVBLib.DWRITE_MEASURING_MODE, _
            glyphRun As DWVBLib.DWRITE_GLYPH_RUN, _
            glyphRunDescription As DWVBLib.DWRITE_GLYPH_RUN_DESCRIPTION, _
            ByVal pClientDrawingEffect As Long)
    Dim cPath       As ID2D1PathGeometry
    Dim cSink       As ID2D1GeometrySink
    Dim cSimpSink   As ID2D1SimplifiedGeometrySink
    Dim cTransform  As ID2D1TransformedGeometry
    
    UpdateEffect pClientDrawingEffect
    
    Set cPath = mcD2DFactory.CreatePathGeometry
    Set cSink = cPath.Open
    
    ' // Because there is the bug on some OS (ID2D1SimplifiedGeometrySink isn't supported by QI call)
    ' // we need to cast manually
    vbaObjSetAddref cSimpSink, ByVal ObjPtr(cSink)
    
    glyphRun.fontFace.GetGlyphRunOutline glyphRun.fontEmSize, ByVal glyphRun.pglyphIndices, ByVal glyphRun.pglyphAdvances, _
                                         ByVal glyphRun.pglyphOffsets, glyphRun.glyphCount, glyphRun.isSideways, _
                                         glyphRun.bidiLevel Mod 2, cSimpSink
    
    cSink.Close
    
    Set cTransform = mcD2DFactory.CreateTransformedGeometry(cPath, D2D1.Matrix3x2F(1, 0, 0, 1, baselineOriginX, baselineOriginY))
    
    mcRenderTarget.DrawGeometry cTransform, mcSolidBrush, 4
    mcRenderTarget.FillGeometry cTransform, mcBitmapBrush
    
End Sub

Private Sub IDWriteTextRendererImp_DrawInlineObject( _
            ByVal pClientDrawingContext As Long, _
            ByVal originX As Single, _
            ByVal originY As Single, _
            ByVal pInlineObject As Long, _
            ByVal isSideways As Long, _
            ByVal isRightToLeft As Long, _
            ByVal pClientDrawingEffect As Long)
    ' // Unimplemented
End Sub

Private Sub IDWriteTextRendererImp_DrawStrikethrough( _
            ByVal pClientDrawingContext As Long, _
            ByVal baselineOriginX As Single, _
            ByVal baselineOriginY As Single, _
            strikethrough As DWVBLib.DWRITE_STRIKETHROUGH, _
            ByVal pClientDrawingEffect As Long)
    Dim cRectangle  As ID2D1RectangleGeometry
    Dim cTransform  As ID2D1TransformedGeometry
    
    UpdateEffect pClientDrawingEffect
    
    Set cRectangle = mcD2DFactory.CreateRectangleGeometry(D2D1.RectF(0, strikethrough.offset, strikethrough.Width, _
                                                          strikethrough.offset + strikethrough.thickness))
    Set cTransform = mcD2DFactory.CreateTransformedGeometry(cRectangle, D2D1.Matrix3x2F(1, 0, 0, 1, _
                                                          baselineOriginX, baselineOriginY))
                                                          
    mcRenderTarget.DrawGeometry cTransform, mcSolidBrush, 4
    mcRenderTarget.FillGeometry cTransform, mcBitmapBrush
    
End Sub

Private Sub IDWriteTextRendererImp_DrawUnderline( _
            ByVal pClientDrawingContext As Long, _
            ByVal baselineOriginX As Single, _
            ByVal baselineOriginY As Single, _
            underline As DWVBLib.DWRITE_UNDERLINE, _
            ByVal pClientDrawingEffect As Long)
    Dim cRectangle  As ID2D1RectangleGeometry
    Dim cTransform  As ID2D1TransformedGeometry
    
    UpdateEffect pClientDrawingEffect
    
    Set cRectangle = mcD2DFactory.CreateRectangleGeometry(D2D1.RectF(0, underline.offset, underline.Width, _
                                                          underline.offset + underline.thickness))
    Set cTransform = mcD2DFactory.CreateTransformedGeometry(cRectangle, D2D1.Matrix3x2F(1, 0, 0, 1, _
                                                          baselineOriginX, baselineOriginY))
                                                          
    mcRenderTarget.DrawGeometry cTransform, mcSolidBrush, 4
    mcRenderTarget.FillGeometry cTransform, mcBitmapBrush
    
End Sub

Private Function IDWriteTextRendererImp_GetCurrentTransform( _
                 ByVal pClientDrawingContext As Long) As DWVBLib.DWRITE_MATRIX
    MoveMemory IDWriteTextRendererImp_GetCurrentTransform, mcRenderTarget.GetTransform, _
                Len(IDWriteTextRendererImp_GetCurrentTransform)
End Function

Private Function IDWriteTextRendererImp_GetPixelsPerDip( _
                 ByVal pClientDrawingContext As Long) As Single
    Dim fDpiX   As Single, fDpiY    As Single
    
    mcRenderTarget.GetDpi fDpiX, fDpiY
    
    IDWriteTextRendererImp_GetPixelsPerDip = fDpiX / 96
    
End Function

Private Function IDWriteTextRendererImp_IsPixelSnappingDisabled( _
                 ByVal pClientDrawingContext As Long) As Long
    IDWriteTextRendererImp_IsPixelSnappingDisabled = 0
End Function

Private Sub UpdateEffect( _
            ByVal pEffect As Long)
    Dim cEffect As CColorEffect
    
    vbaObjSetAddref cEffect, ByVal pEffect
    
    If cEffect Is Nothing Then
        ' // No effect set default color
        mcSolidBrush.SetColor D2D1.ColorF(Black)
    Else
        mcSolidBrush.SetColor D2D1.ColorF(cEffect.Color)
    End If
    
End Sub

