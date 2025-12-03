VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInlineObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const GENERIC_READ As Long = &H80000000

Implements IDWriteInlineObject

Private mcRenderTarget  As ID2D1RenderTarget
Private mcBitmap        As ID2D1Bitmap

Public Sub Initialize( _
           ByVal cRenderTarget As ID2D1RenderTarget, _
           ByVal sFileName As String)
    Dim cFactory        As WICImagingFactory
    Dim cDecoder        As IWICBitmapDecoder
    Dim cFrameDecode    As IWICBitmapFrameDecode
    Dim cConverter      As IWICFormatConverter

    Set cFactory = New WICImagingFactory
    Set cDecoder = cFactory.CreateDecoderFromFilename(StrPtr(sFileName), ByVal vbNullPtr, GENERIC_READ, WICDecodeMetadataCacheOnLoad)
    Set cFrameDecode = cDecoder.GetFrame(0)
    Set cConverter = cFactory.CreateFormatConverter()
    
    cConverter.Initialize cFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, _
                            Nothing, 0, WICBitmapPaletteTypeMedianCut
                            
    Set mcBitmap = cRenderTarget.CreateBitmapFromWicBitmap(ByVal cConverter, ByVal vbNullPtr)
    
    Set mcRenderTarget = cRenderTarget
    
End Sub

Private Sub IDWriteInlineObject_Draw( _
            ByVal pClientDrawingContext As LongPtr, _
            ByVal renderer As IDWriteTextRenderer, _
            ByVal originX As Single, _
            ByVal originY As Single, _
            ByVal isSideways As BOOL, _
            ByVal isRightToLeft As BOOL, _
            ByVal pClientDrawingEffect As IUnknown)
   
   mcRenderTarget.DrawBitmap mcBitmap, D2D1.RectF(originX, originY, originX + mcBitmap.GetSize.width, _
                             originY + mcBitmap.GetSize.Height), 1, 1, ByVal vbNullPtr
    
End Sub

Private Sub IDWriteInlineObject_GetBreakConditions( _
            breakConditionBefore As DWRITE_BREAK_CONDITION, _
            breakConditionAfter As DWRITE_BREAK_CONDITION)
            
    breakConditionBefore = DWRITE_BREAK_CONDITION_NEUTRAL
    breakConditionAfter = DWRITE_BREAK_CONDITION_NEUTRAL
    
End Sub

Private Function IDWriteInlineObject_GetMetrics() As DWRITE_INLINE_OBJECT_METRICS
    IDWriteInlineObject_GetMetrics = DW.InlineObjectMetrics(mcBitmap.GetSize.width, _
                            mcBitmap.GetSize.Height, mcBitmap.GetSize.Height - 4, False)
End Function

Private Function IDWriteInlineObject_GetOverhangMetrics() As DWRITE_OVERHANG_METRICS
    IDWriteInlineObject_GetOverhangMetrics = DW.OverhangMetrics(0, 0, 0, 0)
End Function

