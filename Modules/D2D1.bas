Attribute VB_Name = "D2D1"

' //
' // Module D2D1.bas - Helpes functions for Direct2D
' // By The trick 2018-2023 (c)
' // Most part of functions is ported from D2D1helper.h, d2d1_1helper.h (Microsoft Corporation (c))
' //

Option Explicit

Private Declare Function GetMem8 Lib "msvbvm60" ( _
                         ByRef pSrc As Any, _
                         ByRef pDst As Any) As Long
                         
Public Const FloatMax   As Single = 3.402823466E+38

' // Create ID2D1Factory object
Public Function CreateFactory( _
                Optional ByVal eType As D2D1_FACTORY_TYPE = D2D1_FACTORY_TYPE_SINGLE_THREADED, _
                Optional ByVal eDebugOptions As D2D1_DEBUG_LEVEL = -1) As ID2D1Factory
    Dim tOptions    As D2D1_FACTORY_OPTIONS
    Dim pOptions    As Long
    Dim cIID(1)     As Currency
    
    If eDebugOptions <> -1 Then
    
        tOptions.debugLevel = eDebugOptions
        pOptions = VarPtr(tOptions)
        
    End If
    
    ' // IID_ID2D1Factory
    cIID(0) = 506948672004902.9703@
    cIID(1) = 53149071617564.8146@
    
    Set CreateFactory = D2D1CreateFactory(eType, cIID(0), ByVal pOptions)
    
End Function

Public Function Point2F( _
                ByVal fX As Single, _
                ByVal fY As Single) As D2D1_POINT_2F
                
    Point2F.x = fX
    Point2F.y = fY
    
End Function

Public Function Point2U( _
                ByVal lX As Long, _
                ByVal lY As Long) As D2D1_POINT_2U
                
    Point2U.x = lX
    Point2U.y = lY
    
End Function

Public Function SizeU( _
                Optional ByVal lWidth As Long, _
                Optional ByVal lHeight As Long) As D2D1_SIZE_U
                
    SizeU.Width = lWidth
    SizeU.Height = lHeight
    
End Function

Public Function SizeF( _
                Optional ByVal fWidth As Single, _
                Optional ByVal fHeight As Single) As D2D1_SIZE_F
                
    SizeF.Width = fWidth
    SizeF.Height = fHeight
    
End Function

Public Function RectF( _
                Optional ByVal fLeft As Single, _
                Optional ByVal fTop As Single, _
                Optional ByVal fRight As Single, _
                Optional ByVal fBottom As Single) As D2D1_RECT_F
    
    With RectF
    
    .Left = fLeft
    .Right = fRight
    .Top = fTop
    .bottom = fBottom
    
    End With
    
End Function

Public Function RectF_InfiniteRect() As D2D1_RECT_F
    
    With RectF_InfiniteRect
    
    .Left = -FloatMax
    .Right = FloatMax
    .Top = -FloatMax
    .bottom = FloatMax
    
    End With
    
End Function

Public Function RectU( _
                Optional ByVal lLeft As Long, _
                Optional ByVal lTop As Long, _
                Optional ByVal lRight As Long, _
                Optional ByVal lBottom As Long) As D2D1_RECT_U
    
    With RectU
    
    .Left = lLeft
    .Right = lRight
    .Top = lTop
    .bottom = lBottom
    
    End With
    
End Function

Public Function ArcSegment( _
                ByRef tPoint As D2D1_POINT_2F, _
                ByRef tSize As D2D1_SIZE_F, _
                ByVal fRotationAngle As Single, _
                ByVal eSweepDirection As D2D1_SWEEP_DIRECTION, _
                ByVal eArcSize As D2D1_ARC_SIZE) As D2D1_ARC_SEGMENT
    
    With ArcSegment
    
    .Point = tPoint
    .Size = tSize
    .rotationAngle = fRotationAngle
    .sweepDirection = eSweepDirection
    .arcSize = eArcSize
    
    End With
    
End Function

Public Function BezierSegment( _
                ByRef tPoint1 As D2D1_POINT_2F, _
                ByRef tPoint2 As D2D1_POINT_2F, _
                ByRef tPoint3 As D2D1_POINT_2F) As D2D1_BEZIER_SEGMENT

    With BezierSegment
    
    .point1 = tPoint1
    .point2 = tPoint2
    .point3 = tPoint3
    
    End With
    
End Function

Public Function Ellipse( _
                ByRef tCenter As D2D1_POINT_2F, _
                ByVal fRadiusX As Single, _
                ByVal fRadiusY As Single) As D2D1_ELLIPSE
    
    With Ellipse
    
    .Point = tCenter
    .radiusX = fRadiusX
    .radiusY = fRadiusY
    
    End With
    
End Function

Public Function RoundedRect( _
                ByRef tRect As D2D1_RECT_F, _
                ByVal fRadiusX As Single, _
                ByVal fRadiusY As Single) As D2D1_ROUNDED_RECT
    
    With RoundedRect
    
    .rect = tRect
    .radiusX = fRadiusX
    .radiusY = fRadiusY
    
    End With
    
End Function
    
Public Function BrushProperties( _
                ByVal fOpacity As Single, _
                ByRef tTransform As D2D1_MATRIX_3X2_F) As D2D1_BRUSH_PROPERTIES
    
    With BrushProperties
    
    .opacity = fOpacity
    .Transform = tTransform
    
    End With
    
End Function

Public Function GradientStop( _
                ByVal fPosition As Single, _
                ByRef tColor As D2D1_COLOR_F) As D2D1_GRADIENT_STOP
    
    With GradientStop
    
    .position = fPosition
    .Color = tColor
    
    End With
    
End Function

Public Function QuadraticBezierSegment( _
                ByRef tPoint1 As D2D1_POINT_2F, _
                ByRef tPoint2 As D2D1_POINT_2F) As D2D1_QUADRATIC_BEZIER_SEGMENT
    
    With QuadraticBezierSegment
    
    .point1 = tPoint1
    .point2 = tPoint2
    
    End With
    
End Function

Public Function StrokeStyleProperties( _
                Optional ByVal eStartCap As D2D1_CAP_STYLE = D2D1_CAP_STYLE_FLAT, _
                Optional ByVal eEndCap As D2D1_CAP_STYLE = D2D1_CAP_STYLE_FLAT, _
                Optional ByVal eDashCap As D2D1_CAP_STYLE = D2D1_CAP_STYLE_FLAT, _
                Optional ByVal eLineJoin As D2D1_LINE_JOIN = D2D1_LINE_JOIN_MITER, _
                Optional ByVal fMiterLimit As Single = 10!, _
                Optional ByVal eDashStyle As D2D1_DASH_STYLE = D2D1_DASH_STYLE_SOLID, _
                Optional ByVal fDashOffset As Single) As D2D1_STROKE_STYLE_PROPERTIES
                
    With StrokeStyleProperties
    
    .startCap = eStartCap
    .endCap = eEndCap
    .dashCap = eDashCap
    .lineJoin = eLineJoin
    .miterLimit = fMiterLimit
    .dashStyle = eDashStyle
    .dashOffset = fDashOffset
    
    End With
    
End Function

Public Function BitmapBrushProperties( _
                Optional ByVal eExtendModeX As D2D1_EXTEND_MODE = D2D1_EXTEND_MODE_CLAMP, _
                Optional ByVal eExtendModeY As D2D1_EXTEND_MODE = D2D1_EXTEND_MODE_CLAMP, _
                Optional ByVal eInterpolationMode As D2D1_BITMAP_INTERPOLATION_MODE = _
                D2D1_BITMAP_INTERPOLATION_MODE_LINEAR) As D2D1_BITMAP_BRUSH_PROPERTIES
                
    With BitmapBrushProperties
    
    .extendModeX = eExtendModeX
    .extendModeY = eExtendModeY
    .interpolationMode = eInterpolationMode
    
    End With
    
End Function

Public Function LinearGradientBrushProperties( _
                ByRef tStartPoint As D2D1_POINT_2F, _
                ByRef tEndPoint As D2D1_POINT_2F) As D2D1_LINEAR_GRADIENT_BRUSH_PROPERTIES
                
    With LinearGradientBrushProperties
    
    .startPoint = tStartPoint
    .endPoint = tEndPoint
    
    End With
    
End Function

Public Function RadialGradientBrushProperties( _
                ByRef tCenter As D2D1_POINT_2F, _
                ByRef tGradientOriginOffset As D2D1_POINT_2F, _
                ByVal fRadiusX As Single, _
                ByVal fRadiusY As Single) As D2D1_RADIAL_GRADIENT_BRUSH_PROPERTIES
                
    With RadialGradientBrushProperties
    
    .center = tCenter
    .gradientOriginOffset = tGradientOriginOffset
    .radiusX = fRadiusX
    .radiusY = fRadiusY
    
    End With
    
End Function

Public Function PixelFormat( _
                Optional ByVal eDxgiFormat As DXGI_FORMAT = DXGI_FORMAT_UNKNOWN, _
                Optional ByVal eAlphaMode As D2D1_ALPHA_MODE = D2D1_ALPHA_MODE_UNKNOWN) As D2D1_PIXEL_FORMAT
                
    PixelFormat.Format = eDxgiFormat
    PixelFormat.alphaMode = eAlphaMode
    
End Function

Public Function BitmapProperties( _
                ByRef tPixelFormat As D2D1_PIXEL_FORMAT, _
                ByVal fDpiX As Single, _
                ByVal fDpiY As Single) As D2D1_BITMAP_PROPERTIES
                
    With BitmapProperties
    
    .PixelFormat = tPixelFormat
    .dpiX = fDpiX
    .dpiY = fDpiY
    
    End With
    
End Function

Public Function RenderTargetProperties( _
                ByRef tPixelFormat As D2D1_PIXEL_FORMAT, _
                Optional ByVal eType As D2D1_RENDER_TARGET_TYPE = D2D1_RENDER_TARGET_TYPE_DEFAULT, _
                Optional ByVal fDpiX As Single, _
                Optional ByVal fDpiY As Single, _
                Optional ByVal eUsage As D2D1_RENDER_TARGET_USAGE = D2D1_RENDER_TARGET_USAGE_NONE, _
                Optional ByVal eMinLevel As D2D1_FEATURE_LEVEL = _
                D2D1_FEATURE_LEVEL_DEFAULT) As D2D1_RENDER_TARGET_PROPERTIES
    
    With RenderTargetProperties
    
    .Type = eType
    .PixelFormat = tPixelFormat
    .dpiX = fDpiX
    .dpiY = fDpiY
    .usage = eUsage
    .minLevel = eMinLevel
    
    End With
    
End Function

Public Function HwndRenderTargetProperties( _
                ByVal hWnd As Long, _
                ByRef tPixelSize As D2D1_SIZE_U, _
                Optional ByVal ePresentOptions As D2D1_PRESENT_OPTIONS = _
                D2D1_PRESENT_OPTIONS_NONE) As D2D1_HWND_RENDER_TARGET_PROPERTIES
    
    With HwndRenderTargetProperties
    
    .hWnd = hWnd
    .pixelSize = tPixelSize
    .presentOptions = ePresentOptions
    
    End With
    
End Function

Public Function LayerParameters( _
                ByRef tContentBounds As D2D1_RECT_F, _
                ByVal cGeometricMask As ID2D1Geometry, _
                ByVal eMaskAntialiasMode As D2D1_ANTIALIAS_MODE, _
                ByRef tMaskTransform As D2D1_MATRIX_3X2_F, _
                Optional ByVal fOpacity As Single = 1!, _
                Optional ByVal cOpacityBrush As ID2D1Brush, _
                Optional ByVal eLayerOptions As D2D1_LAYER_OPTIONS = D2D1_LAYER_OPTIONS_NONE) As D2D1_LAYER_PARAMETERS
    
    With LayerParameters
    
    .contentBounds = tContentBounds
    Set .geometricMask = cGeometricMask
    .maskAntialiasMode = eMaskAntialiasMode
    .maskTransform = tMaskTransform
    .opacity = fOpacity
    Set .opacityBrush = cOpacityBrush
    .layerOptions = eLayerOptions
    
    End With
    
End Function

Public Function DrawingStateDescription( _
                ByVal eAntialiasMode As D2D1_ANTIALIAS_MODE, _
                ByVal eTextAntialiasMode As D2D1_TEXT_ANTIALIAS_MODE, _
                ByVal cTag1 As Currency, _
                ByVal cTag2 As Currency, _
                ByRef tTransform As D2D1_MATRIX_3X2_F) As D2D1_DRAWING_STATE_DESCRIPTION
    
    With DrawingStateDescription
    
    .antialiasMode = eAntialiasMode
    .textAntialiasMode = eTextAntialiasMode
    .tag1 = cTag1
    .tag2 = cTag2
    .Transform = tTransform
    
    End With
    
End Function
    
Public Function ColorF( _
                ByVal eColor As D2D1_COLORS, _
                Optional ByVal fAlpha As Single = 1!) As D2D1_COLOR_F
    Dim bR As Byte
    Dim bB As Byte
    Dim bG As Byte
    
    bB = eColor And &HFF
    bG = (eColor \ &H100) And &HFF
    bR = (eColor \ &H10000) And &HFF
    
    With ColorF
    
    .r = bR / 255!
    .g = bG / 255!
    .b = bB / 255!
    .a = fAlpha
    
    End With
    
End Function

Public Function ColorF2( _
                ByVal fR As Single, _
                ByVal fG As Single, _
                ByVal fB As Single, _
                Optional ByVal fAlpha As Single = 1!) As D2D1_COLOR_F
    
    With ColorF2
    
    .r = fR
    .g = fG
    .b = fB
    .a = fAlpha
    
    End With
    
End Function

Public Function Matrix3x2F( _
                ByVal f_11 As Single, _
                ByVal f_12 As Single, _
                ByVal f_21 As Single, _
                ByVal f_22 As Single, _
                ByVal f_31 As Single, _
                ByVal f_32 As Single) As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F
    
    .m_11 = f_11
    .m_12 = f_12
    .m_21 = f_21
    .m_22 = f_22
    .m_31 = f_31
    .m_32 = f_32
    
    End With
    
End Function

Public Function Matrix3x2F_Identity() As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F_Identity
    
    .m_11 = 1
    .m_22 = 1
    
    End With
    
End Function

Public Function Matrix3x2F_Translation( _
                ByRef tSize As D2D1_SIZE_F) As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F_Translation
    
    .m_11 = 1!: .m_12 = 0!
    .m_21 = 0!: .m_22 = 1!
    .m_31 = tSize.Width: .m_32 = tSize.Height
    
    End With
    
End Function

Public Function Matrix3x2F_Translation2( _
                ByVal fWidth As Single, _
                ByVal fHeight As Single) As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F_Translation2
    
    .m_11 = 1!: .m_12 = 0!
    .m_21 = 0!: .m_22 = 1!
    .m_31 = fWidth: .m_32 = fHeight
    
    End With
    
End Function

Public Function Matrix3x2F_Scale( _
                ByRef tSize As D2D1_SIZE_F, _
                ByRef tCenter As D2D1_POINT_2F) As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F_Scale
    
    .m_11 = tSize.Width: .m_12 = 0!
    .m_21 = 0!: .m_22 = tSize.Height
    .m_31 = tCenter.x - tSize.Width * tCenter.x
    .m_32 = tCenter.y - tSize.Height * tCenter.y
    
    End With
    
End Function

Public Function Matrix3x2F_Scale2( _
                ByVal fWidth As Single, _
                ByVal fHeight As Single, _
                ByRef tCenter As D2D1_POINT_2F) As D2D1_MATRIX_3X2_F
    
    With Matrix3x2F_Scale2
    
    .m_11 = fWidth: .m_12 = 0!
    .m_21 = 0!: .m_22 = fHeight
    .m_31 = tCenter.x - fWidth * tCenter.x
    .m_32 = tCenter.y - fHeight * tCenter.y
    
    End With
    
End Function

Public Function Matrix3x2F_Rotation( _
                ByVal fAngle As Single, _
                ByRef tCenter As D2D1_POINT_2F) As D2D1_MATRIX_3X2_F
    D2D1MakeRotateMatrix fAngle, tCenter.x, tCenter.y, Matrix3x2F_Rotation
End Function

Public Function Matrix3x2F_Rotation2( _
                ByVal fAngle As Single, _
                ByVal fCenterX As Single, _
                ByVal fCentery As Single) As D2D1_MATRIX_3X2_F
    D2D1MakeRotateMatrix fAngle, fCenterX, fCentery, Matrix3x2F_Rotation2
End Function

Public Function Matrix3x2F_Skew( _
                ByVal fAngleX As Single, _
                ByVal fAngleY As Single, _
                ByRef tCenter As D2D1_POINT_2F) As D2D1_MATRIX_3X2_F
    D2D1MakeSkewMatrix fAngleX, fAngleY, tCenter.x, tCenter.y, Matrix3x2F_Skew
End Function

Public Function Matrix3x2F_Determinant( _
                ByRef tMtx As D2D1_MATRIX_3X2_F) As Single
    Matrix3x2F_Determinant = (tMtx.m_11 * tMtx.m_22) - (tMtx.m_12 * tMtx.m_21)
End Function

Public Function Matrix3x2F_IsInvertible( _
                ByRef tMtx As D2D1_MATRIX_3X2_F) As Boolean
    Matrix3x2F_IsInvertible = D2D1IsMatrixInvertible(tMtx)
End Function

Public Function Matrix3x2F_Invert( _
                ByRef tMtx As D2D1_MATRIX_3X2_F) As Boolean
    Matrix3x2F_Invert = D2D1InvertMatrix(tMtx)
End Function

Public Function Matrix3x2F_IsIdentity( _
                ByRef tMtx As D2D1_MATRIX_3X2_F) As Boolean
    Matrix3x2F_IsIdentity = (tMtx.m_11 = 1!) And (tMtx.m_12 = 0!) And _
                            (tMtx.m_21 = 0!) And (tMtx.m_22 = 1!) And _
                            (tMtx.m_31 = 0!) And (tMtx.m_32 = 0!)
End Function

Public Function Matrix3x2F_SetProduct( _
                ByRef tMtx1 As D2D1_MATRIX_3X2_F, _
                ByRef tMtx2 As D2D1_MATRIX_3X2_F) As D2D1_MATRIX_3X2_F
                
    With Matrix3x2F_SetProduct
    
    .m_11 = tMtx1.m_11 * tMtx2.m_11 + tMtx1.m_12 * tMtx2.m_21
    .m_12 = tMtx1.m_11 * tMtx2.m_12 + tMtx1.m_12 * tMtx2.m_22
    .m_21 = tMtx1.m_21 * tMtx2.m_11 + tMtx1.m_22 * tMtx2.m_21
    .m_22 = tMtx1.m_21 * tMtx2.m_12 + tMtx1.m_22 * tMtx2.m_22
    .m_31 = tMtx1.m_31 * tMtx2.m_11 + tMtx1.m_32 * tMtx2.m_21 + tMtx2.m_31
    .m_32 = tMtx1.m_31 * tMtx2.m_12 + tMtx1.m_32 * tMtx2.m_22 + tMtx2.m_32
    
    End With
    
End Function

Public Function Matrix3x2F_TransformPoint( _
                ByRef tMtx As D2D1_MATRIX_3X2_F, _
                ByRef tPoint As D2D1_POINT_2F) As D2D1_POINT_2F
    
    With Matrix3x2F_TransformPoint
    
    .x = tPoint.x * tMtx.m_11 + tPoint.y * tMtx.m_21 + tMtx.m_31
    .y = tPoint.x * tMtx.m_12 + tPoint.y * tMtx.m_22 + tMtx.m_32
    
    End With
    
End Function

Public Function Matrix4x3F( _
                ByVal f_11 As Single, _
                ByVal f_12 As Single, _
                ByVal f_13 As Single, _
                ByVal f_21 As Single, _
                ByVal f_22 As Single, _
                ByVal f_23 As Single, _
                ByVal f_31 As Single, _
                ByVal f_32 As Single, _
                ByVal f_33 As Single, _
                ByVal f_41 As Single, _
                ByVal f_42 As Single, _
                ByVal f_43 As Single) As D2D1_MATRIX_4X3_F
    
    With Matrix4x3F
    
    .m_11 = f_11
    .m_12 = f_12
    .m_13 = f_13
    .m_21 = f_21
    .m_22 = f_22
    .m_23 = f_23
    .m_31 = f_31
    .m_32 = f_32
    .m_33 = f_33
    .m_41 = f_41
    .m_42 = f_42
    .m_43 = f_43

    End With
    
End Function

Public Function Matrix4x3F_Identity() As D2D1_MATRIX_4X3_F
    
    With Matrix4x3F_Identity
    
    .m_11 = 1!
    .m_22 = 1!
    .m_33 = 1!
    
    End With
    
End Function

Public Function Matrix4x4F( _
                ByVal f_11 As Single, _
                ByVal f_12 As Single, _
                ByVal f_13 As Single, _
                ByVal f_14 As Single, _
                ByVal f_21 As Single, _
                ByVal f_22 As Single, _
                ByVal f_23 As Single, _
                ByVal f_24 As Single, _
                ByVal f_31 As Single, _
                ByVal f_32 As Single, _
                ByVal f_33 As Single, _
                ByVal f_34 As Single, _
                ByVal f_41 As Single, _
                ByVal f_42 As Single, _
                ByVal f_43 As Single, _
                ByVal f_44 As Single) As D2D1_MATRIX_4X4_F
    
    With Matrix4x4F
    
    .m_11 = f_11
    .m_12 = f_12
    .m_13 = f_13
    .m_14 = f_14
    .m_21 = f_21
    .m_22 = f_22
    .m_23 = f_23
    .m_24 = f_24
    .m_31 = f_31
    .m_32 = f_32
    .m_33 = f_33
    .m_34 = f_34
    .m_41 = f_41
    .m_42 = f_42
    .m_43 = f_43
    .m_44 = f_44
    
    End With
    
End Function

Public Function Matrix4x4F_Identity() As D2D1_MATRIX_4X4_F
    
    With Matrix4x4F_Identity
    
    .m_11 = 1!
    .m_22 = 1!
    .m_33 = 1!
    .m_44 = 1!
    
    End With
    
End Function

Public Function Matrix4x4F_Translation( _
                ByVal fX As Single, _
                ByVal fY As Single, _
                ByVal fZ As Single) As D2D1_MATRIX_4X4_F
    
    With Matrix4x4F_Translation
    
    .m_11 = 1!
    .m_22 = 1!
    .m_33 = 1!
    .m_44 = 1!
    .m_41 = fX
    .m_42 = fY
    .m_43 = fZ
    
    End With
    
End Function

Public Function Matrix4x4F_Scale( _
                ByVal fX As Single, _
                ByVal fY As Single, _
                ByVal fZ As Single) As D2D1_MATRIX_4X4_F
    
    With Matrix4x4F_Scale
    
    .m_11 = fX
    .m_22 = fY
    .m_33 = fZ
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_RotationX( _
                ByVal fDegreeX As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fSin            As Single
    Dim fCos            As Single
    
    fAngleInRadian = fDegreeX * (3.141593! / 180!)
    
    D2D1SinCos fAngleInRadian, fSin, fCos
    
    With Matrix4x4F_RotationX
    
    .m_11 = 1
    .m_22 = fCos
    .m_33 = fCos
    .m_23 = fSin
    .m_32 = -fSin
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_RotationY( _
                ByVal fDegreeY As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fSin            As Single
    Dim fCos            As Single
    
    fAngleInRadian = fDegreeY * (3.141593! / 180!)
    
    D2D1SinCos fAngleInRadian, fSin, fCos
    
    With Matrix4x4F_RotationY
    
    .m_11 = fCos
    .m_13 = -fSin
    .m_22 = 1!
    .m_31 = fSin
    .m_33 = fCos
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_RotationZ( _
                ByVal fDegreeZ As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fSin            As Single
    Dim fCos            As Single
    
    fAngleInRadian = fDegreeZ * (3.141593! / 180!)
    
    D2D1SinCos fAngleInRadian, fSin, fCos
    
    With Matrix4x4F_RotationZ
    
    .m_11 = fCos
    .m_12 = fSin
    .m_21 = -fSin
    .m_22 = fCos
    .m_33 = 1!
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_RotationArbitraryAxis( _
                ByVal fX As Single, _
                ByVal fY As Single, _
                ByVal fZ As Single, _
                ByVal fDegree As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fMagnitude      As Single
    Dim fSin            As Single
    Dim fCos            As Single
    Dim fInvCos         As Single
    
    fMagnitude = D2D1Vec3Length(fX, fY, fZ)
    
    fX = fX / fMagnitude
    fY = fY / fMagnitude
    fZ = fZ / fMagnitude
    
    fAngleInRadian = fDegree * (3.141593! / 180!)
    
    D2D1SinCos fAngleInRadian, fSin, fCos
    
    fInvCos = 1 - fCos
    
    With Matrix4x4F_RotationArbitraryAxis
    
    .m_11 = 1 + fInvCos * (fX * fX - 1)
    .m_12 = fZ * fSin + fInvCos * fX * fY
    .m_13 = -fY * fSin + fInvCos * fX * fZ
    .m_21 = -fZ * fSin + fInvCos * fY * fX
    .m_22 = 1 + fInvCos * (fY * fY - 1)
    .m_23 = fX * fSin + fInvCos * fY * fZ
    .m_31 = fY * fSin + fInvCos * fZ * fX
    .m_32 = -fX * fSin + fInvCos * fZ * fY
    .m_33 = 1 + fInvCos * (fZ * fZ - 1)
    .m_44 = 1
    
    End With
    
End Function

Public Function Matrix4x4F_SkewX( _
                ByVal fDegreeX As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fTan            As Single
    
    fAngleInRadian = fDegreeX * (3.141593! / 180!)
    fTan = D2D1Tan(fAngleInRadian)

    With Matrix4x4F_SkewX
    
    .m_11 = 1!
    .m_21 = fTan
    .m_22 = 1!
    .m_33 = 1!
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_SkewY( _
                ByVal fDegreeY As Single) As D2D1_MATRIX_4X4_F
    Dim fAngleInRadian  As Single
    Dim fTan            As Single
    
    fAngleInRadian = fDegreeY * (3.141593! / 180!)
    fTan = D2D1Tan(fAngleInRadian)

    With Matrix4x4F_SkewY
    
    .m_11 = 1!
    .m_12 = fTan
    .m_22 = 1!
    .m_33 = 1!
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_PerspectiveProjection( _
                ByVal fDepth As Single) As D2D1_MATRIX_4X4_F
    Dim fProj   As Single
    
    If fDepth > 0 Then
        fProj = -1 / fDepth
    End If
    
    With Matrix4x4F_PerspectiveProjection
    
    .m_11 = 1!
    .m_12 = 1!
    .m_22 = 1!
    .m_33 = 1!
    .m_34 = fProj
    .m_44 = 1!

    End With
    
End Function

Public Function Matrix4x4F_Determinant( _
                ByRef tMtx As D2D1_MATRIX_4X4_F) As Single
    Dim fMinor1 As Single
    Dim fMinor2 As Single
    Dim fMinor3 As Single
    Dim fMinor4 As Single
    
    With tMtx
    
    fMinor1 = .m_41 * (.m_12 * (.m_23 * .m_34 - .m_33 * .m_24) - _
              .m_13 * (.m_22 * .m_34 - .m_24 * .m_32) + _
              .m_14 * (.m_22 * .m_33 - .m_23 * .m_32))
    fMinor2 = .m_42 * (.m_11 * (.m_21 * .m_34 - .m_31 * .m_24) - _
              .m_13 * (.m_21 * .m_34 - .m_24 * .m_31) + _
              .m_14 * (.m_21 * .m_33 - .m_23 * .m_31))
    fMinor3 = .m_43 * (.m_11 * (.m_22 * .m_34 - .m_32 * .m_24) - _
              .m_12 * (.m_21 * .m_34 - .m_24 * .m_31) + _
              .m_14 * (.m_21 * .m_32 - .m_22 * .m_31))
    fMinor4 = .m_44 * (.m_11 * (.m_22 * .m_33 - .m_32 * .m_23) - _
              .m_12 * (.m_21 * .m_33 - .m_23 * .m_31) + _
              .m_13 * (.m_21 * .m_32 - .m_22 * .m_31))

    Matrix4x4F_Determinant = fMinor1 - fMinor2 + fMinor3 - fMinor4
            
    End With
    
End Function

Public Function Matrix4x4F_IsIdentity( _
                ByRef tMtx As D2D1_MATRIX_4X4_F) As Boolean
    Matrix4x4F_IsIdentity = (tMtx.m_11 = 1!) And (tMtx.m_12 = 0!) And (tMtx.m_13 = 0!) And (tMtx.m_14 = 0!) And _
                            (tMtx.m_21 = 0!) And (tMtx.m_22 = 1!) And (tMtx.m_23 = 0!) And (tMtx.m_24 = 0!) And _
                            (tMtx.m_31 = 0!) And (tMtx.m_32 = 0!) And (tMtx.m_33 = 1!) And (tMtx.m_34 = 0!) And _
                            (tMtx.m_41 = 0!) And (tMtx.m_42 = 0!) And (tMtx.m_43 = 1!) And (tMtx.m_44 = 1!)
End Function

Public Function Matrix4x4F_SetProduct( _
                ByRef tMtx1 As D2D1_MATRIX_4X4_F, _
                ByRef tMtx2 As D2D1_MATRIX_4X4_F) As D2D1_MATRIX_4X4_F
                
    With Matrix4x4F_SetProduct
    
    .m_11 = tMtx1.m_11 * tMtx2.m_11 + tMtx1.m_12 * tMtx2.m_21 + tMtx1.m_13 * tMtx2.m_31 + tMtx1.m_14 * tMtx2.m_41
    .m_12 = tMtx1.m_11 * tMtx2.m_12 + tMtx1.m_12 * tMtx2.m_22 + tMtx1.m_13 * tMtx2.m_32 + tMtx1.m_14 * tMtx2.m_42
    .m_13 = tMtx1.m_11 * tMtx2.m_13 + tMtx1.m_12 * tMtx2.m_23 + tMtx1.m_13 * tMtx2.m_33 + tMtx1.m_14 * tMtx2.m_43
    .m_14 = tMtx1.m_11 * tMtx2.m_14 + tMtx1.m_12 * tMtx2.m_24 + tMtx1.m_13 * tMtx2.m_34 + tMtx1.m_14 * tMtx2.m_44

    .m_21 = tMtx1.m_21 * tMtx2.m_11 + tMtx1.m_22 * tMtx2.m_21 + tMtx1.m_23 * tMtx2.m_31 + tMtx1.m_24 * tMtx2.m_41
    .m_22 = tMtx1.m_21 * tMtx2.m_12 + tMtx1.m_22 * tMtx2.m_22 + tMtx1.m_23 * tMtx2.m_32 + tMtx1.m_24 * tMtx2.m_42
    .m_23 = tMtx1.m_21 * tMtx2.m_13 + tMtx1.m_22 * tMtx2.m_23 + tMtx1.m_23 * tMtx2.m_33 + tMtx1.m_24 * tMtx2.m_43
    .m_24 = tMtx1.m_21 * tMtx2.m_14 + tMtx1.m_22 * tMtx2.m_24 + tMtx1.m_23 * tMtx2.m_34 + tMtx1.m_24 * tMtx2.m_44

    .m_31 = tMtx1.m_31 * tMtx2.m_11 + tMtx1.m_32 * tMtx2.m_21 + tMtx1.m_33 * tMtx2.m_31 + tMtx1.m_34 * tMtx2.m_41
    .m_32 = tMtx1.m_31 * tMtx2.m_12 + tMtx1.m_32 * tMtx2.m_22 + tMtx1.m_33 * tMtx2.m_32 + tMtx1.m_34 * tMtx2.m_42
    .m_33 = tMtx1.m_31 * tMtx2.m_13 + tMtx1.m_32 * tMtx2.m_23 + tMtx1.m_33 * tMtx2.m_33 + tMtx1.m_34 * tMtx2.m_43
    .m_34 = tMtx1.m_31 * tMtx2.m_14 + tMtx1.m_32 * tMtx2.m_24 + tMtx1.m_33 * tMtx2.m_34 + tMtx1.m_34 * tMtx2.m_44

    .m_41 = tMtx1.m_41 * tMtx2.m_11 + tMtx1.m_42 * tMtx2.m_21 + tMtx1.m_43 * tMtx2.m_31 + tMtx1.m_44 * tMtx2.m_41
    .m_42 = tMtx1.m_41 * tMtx2.m_12 + tMtx1.m_42 * tMtx2.m_22 + tMtx1.m_43 * tMtx2.m_32 + tMtx1.m_44 * tMtx2.m_42
    .m_43 = tMtx1.m_41 * tMtx2.m_13 + tMtx1.m_42 * tMtx2.m_23 + tMtx1.m_43 * tMtx2.m_33 + tMtx1.m_44 * tMtx2.m_43
    .m_44 = tMtx1.m_41 * tMtx2.m_14 + tMtx1.m_42 * tMtx2.m_24 + tMtx1.m_43 * tMtx2.m_34 + tMtx1.m_44 * tMtx2.m_44
        
    End With
    
End Function

Public Function Matrix5x4F( _
                ByVal f_11 As Single, _
                ByVal f_12 As Single, _
                ByVal f_13 As Single, _
                ByVal f_14 As Single, _
                ByVal f_21 As Single, _
                ByVal f_22 As Single, _
                ByVal f_23 As Single, _
                ByVal f_24 As Single, _
                ByVal f_31 As Single, _
                ByVal f_32 As Single, _
                ByVal f_33 As Single, _
                ByVal f_34 As Single, _
                ByVal f_41 As Single, _
                ByVal f_42 As Single, _
                ByVal f_43 As Single, _
                ByVal f_44 As Single, _
                ByVal f_51 As Single, _
                ByVal f_52 As Single, _
                ByVal f_53 As Single, _
                ByVal f_54 As Single) As D2D1_MATRIX_5X4_F
    
    With Matrix5x4F
    
    .m_11 = f_11
    .m_12 = f_12
    .m_13 = f_13
    .m_14 = f_14
    .m_21 = f_21
    .m_22 = f_22
    .m_23 = f_23
    .m_24 = f_24
    .m_31 = f_31
    .m_32 = f_32
    .m_33 = f_33
    .m_34 = f_34
    .m_41 = f_41
    .m_42 = f_42
    .m_43 = f_43
    .m_44 = f_44
    .m_51 = f_51
    .m_52 = f_52
    .m_53 = f_53
    .m_54 = f_54
    
    End With
    
End Function

Public Function Matrix5x4F_Identity() As D2D1_MATRIX_5X4_F
    
    With Matrix5x4F_Identity
    
    .m_11 = 1!
    .m_22 = 1!
    .m_33 = 1!
    .m_44 = 1!
    
    End With
    
End Function

Public Function ConvertColorSpace( _
                ByVal eSourceColorSpace As D2D1_COLOR_SPACE, _
                ByVal eDestinationColorSpace As D2D1_COLOR_SPACE, _
                ByRef tColor As D2D1_COLOR_F) As D2D1_COLOR_F
    ConvertColorSpace = D2D1ConvertColorSpace(eSourceColorSpace, eDestinationColorSpace, tColor)
End Function

' // TODO
' // DrawingStateDescription1
' // BitmapProperties1
' // LayerParameters1
' // StrokeStyleProperties1
' // ImageBrushProperties
' // BitmapBrushProperties1
' // PrintControlProperties
' // RenderingControls
' // EffectInputDescription
' // CreationProperties
' // Point2L
' // RectL

Public Sub SetDpiCompensatedEffectInput( _
           ByVal cContext As ID2D1DeviceContext, _
           ByVal cEffect As ID2D1Effect, _
           ByVal lInputIndex As Long, _
           ByVal cBitmap As ID2D1Bitmap, _
           Optional ByVal eInterpolationMode As D2D1_INTERPOLATION_MODE = D2D1_INTERPOLATION_MODE_LINEAR, _
           Optional ByVal eBorderMode As D2D1_BORDER_MODE = D2D1_BORDER_MODE_HARD)
    Dim cDpiEffect  As ID2D1Effect
    Dim tCLSID      As UUID
    Dim tDPI        As D2D1_POINT_2F
    
    If cBitmap Is Nothing Then
        cEffect.SetInput lInputIndex
        Exit Sub
    End If
    
    ' // CLSID_D2D1DpiCompensation
    GetMem8 511502141527783.9815@, tCLSID
    GetMem8 294592394174280.438@, ByVal VarPtr(tCLSID) + 8
    
    Set cDpiEffect = cContext.CreateEffect(tCLSID)
    
    cDpiEffect.SetInput 0, cBitmap
    cBitmap.GetDpi tDPI.x, tDPI.y
    cDpiEffect.SetValue D2D1_DPICOMPENSATION_PROP_INPUT_DPI, D2D1_PROPERTY_TYPE_UNKNOWN, tDPI, LenB(tDPI)
    cDpiEffect.SetValue D2D1_DPICOMPENSATION_PROP_INTERPOLATION_MODE, D2D1_PROPERTY_TYPE_UNKNOWN, eInterpolationMode, LenB(eInterpolationMode)
    cDpiEffect.SetValue D2D1_DPICOMPENSATION_PROP_BORDER_MODE, D2D1_PROPERTY_TYPE_UNKNOWN, eBorderMode, LenB(eBorderMode)
    cEffect.SetInput lInputIndex, cDpiEffect.GetOutput
    
End Sub

Public Function Vector2F( _
                Optional ByVal fX As Single, _
                Optional ByVal fY As Single) As D2D1_VECTOR_2F

    Vector2F.x = fX
    Vector2F.y = fY
    
End Function

Public Function Vector3F( _
                Optional ByVal fX As Single, _
                Optional ByVal fY As Single, _
                Optional ByVal fZ As Single) As D2D1_VECTOR_3F

    Vector3F.x = fX
    Vector3F.y = fY
    Vector3F.z = fZ
    
End Function

Public Function Vector4F( _
                Optional ByVal fX As Single, _
                Optional ByVal fY As Single, _
                Optional ByVal fZ As Single, _
                Optional ByVal fW As Single) As D2D1_VECTOR_4F

    Vector4F.x = fX
    Vector4F.y = fY
    Vector4F.z = fZ
    Vector4F.w = fW
    
End Function

