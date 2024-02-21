Attribute VB_Name = "DW"
' //
' // Module DW.bas - Helpes functions for DirectWrite
' // By The trick 2018 (c)
' //

Option Explicit

Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
Private Declare Function GetMem4 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
                         
Public Function CreateFactory( _
                ByVal eFactoryType As DWRITE_FACTORY_TYPE) As IDWriteFactory
    Set CreateFactory = DWriteCreateFactory(eFactoryType, IID_IDWriteFactory)
End Function
                         
Public Function IID_IDWriteFactory() As UUID
    GetMem8 543017151384793.4554@, IID_IDWriteFactory
    GetMem8 524995195940339.1138@, ByVal VarPtr(IID_IDWriteFactory) + 8
End Function

Public Function ClusterMetric( _
                Optional ByVal eFlags As DWRITE_CLUSTER_METRICS_FLAGS, _
                Optional ByVal lLength As Long, _
                Optional ByVal fWidth As Single) As DWRITE_CLUSTER_METRICS
    
    With ClusterMetric
    
    .flags = eFlags
    GetMem4 lLength, .length
    .Width = fWidth
    
    End With
    
End Function

Public Function FontFeature( _
                Optional ByVal eTag As DWRITE_FONT_FEATURE_TAG, _
                Optional ByVal lParameter As Long) As DWRITE_FONT_FEATURE
    
    With FontFeature
    
    .nameTag = eTag
    .parameter = lParameter
    
    End With
        
End Function

Public Function TextRange( _
                Optional ByVal lStartPosition As Long, _
                Optional ByVal lLength As Long) As DWRITE_TEXT_RANGE
    
    With TextRange
    
    .startPosition = lStartPosition
    .length = lLength
    
    End With
        
End Function

Public Function FontMetrics( _
                Optional ByVal lDesignUnitsPerEm As Long, _
                Optional ByVal lAscent As Long, _
                Optional ByVal lDescent As Long, _
                Optional ByVal iLineGap As Integer, _
                Optional ByVal lCapHeight As Long, _
                Optional ByVal lXHeight As Long, _
                Optional ByVal iUnderlinePosition As Integer, _
                Optional ByVal lUnderlineThickness As Long, _
                Optional ByVal iStrikethroughPosition As Integer, _
                Optional ByVal lStrikethroughThickness As Long) As DWRITE_FONT_METRICS
    
    With FontMetrics
    
    GetMem4 lDesignUnitsPerEm, .designUnitsPerEm
    GetMem4 lAscent, .ascent
    GetMem4 lDescent, .descent
    GetMem4 lCapHeight, .capHeight
    GetMem4 lXHeight, .xHeight
    GetMem4 lUnderlineThickness, .underlineThickness
    GetMem4 lStrikethroughThickness, .strikethroughThickness
    
    .lineGap = iLineGap
    .underlinePosition = iUnderlinePosition
    .strikethroughPosition = iStrikethroughPosition

    End With

End Function

Public Function Trimming( _
                Optional ByVal eGranularity As DWRITE_TRIMMING_GRANULARITY, _
                Optional ByVal lDelimiter As Long, _
                Optional ByVal lDelimiterCount As Long) As DWRITE_TRIMMING

    With Trimming
    
    .granularity = eGranularity
    .delimiter = lDelimiter
    .delimiterCount = lDelimiterCount
    
    End With
    
End Function

Public Function InlineObjectMetrics( _
                ByVal fWidth As Single, _
                ByVal fHeight As Single, _
                ByVal fBaseline As Single, _
                ByVal bSupportsSideways As Boolean) As DWRITE_INLINE_OBJECT_METRICS

    With InlineObjectMetrics
    
    .Width = fWidth
    .Height = fHeight
    .baseline = fBaseline
    .supportsSideways = bSupportsSideways And 1
    
    End With
    
End Function

Public Function OverhangMetrics( _
                ByVal fLeft As Single, _
                ByVal fTop As Single, _
                ByVal fRight As Single, _
                ByVal fBottom As Single) As DWRITE_OVERHANG_METRICS

    With OverhangMetrics
    
    .Left = fLeft
    .Top = fTop
    .Right = fRight
    .bottom = fBottom
    
    End With
    
End Function
'
'Public Function GlyphMetrics( _
'                Optional ByVal lLeftSideBearing As Long, _
'                Optional ByVal lAdvanceWidth As Long, _
'                Optional ByVal lRightSideBearing As Long, _
'                Optional ByVal lTopSideBearing As Long, _
'                Optional ByVal lAdvanceHeight As Long, _
'                Optional ByVal lBottomSideBearing As Long, _
'                Optional ByVal lVerticalOriginY As Long) As DWRITE_GLYPH_METRICS
'
'    With GlyphMetrics
'
'    .leftSideBearing = lLeftSideBearing
'    .advanceWidth = lAdvanceWidth
'    .rightSideBearing = lRightSideBearing
'    .topSideBearing = lTopSideBearing
'    .advanceHeight = lAdvanceHeight
'    .bottomSideBearing = lBottomSideBearing
'    .verticalOriginY = lVerticalOriginY
'
'    End With
'
'End Function
'
'Public Function GlyphOffset( _
'                Optional ByVal fAdvanceOffset As Single, _
'                Optional ByVal fAscenderOffset As Single) As DWRITE_GLYPH_OFFSET
'
'    With GlyphOffset
'
'    .advanceOffset = fAdvanceOffset
'    .ascenderOffset = fAscenderOffset
'
'    End With
'
'End Function
'
'Public Function GlyphRun( _
'                Optional ByVal cFontFace As IDWriteFontFace, _
'                Optional ByVal fFontEmSize As Single, _
'                Optional ByVal lGlyphCount As Long, _
'                Optional ByVal pGlyphIndices As Long, _
'                Optional ByVal pGlyphAdvances As Long, _
'                Optional ByVal pGlyphOffsets As Long, _
'                Optional ByVal bIsSideways As Boolean, _
'                Optional ByVal lBidiLevel As Long) As DWRITE_GLYPH_RUN
'
'    With GlyphRun
'
'    Set .fontFace = cFontFace
'    .fontEmSize = fFontEmSize
'    .glyphCount = lGlyphCount
'    .pGlyphIndices = pGlyphIndices
'    .pGlyphAdvances = pGlyphAdvances
'    .pGlyphOffsets = pGlyphOffsets
'    .isSideways = bIsSideways And 1
'    .bidiLevel = lBidiLevel
'
'    End With
'
'End Function
'
'Public Function GlyphRunDescription( _
'                Optional ByRef sLocaleName As String, _
'                Optional ByRef sString As String, _
'                Optional ByVal lStringLength As Long, _
'                Optional ByVal pClusterMap As Long, _
'                Optional ByVal lTextPosition As Long) As DWRITE_GLYPH_RUN_DESCRIPTION
'
'    With GlyphRunDescription
'
'    .plocaleName = StrPtr(sLocaleName)
'    .pstring = StrPtr(sString)
'    .stringLength = lStringLength
'    .pClusterMap = pClusterMap
'    .textPosition = lTextPosition
'
'    End With
'
'End Function


