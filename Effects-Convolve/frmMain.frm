VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "comctl32.Ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direct2D convolve effect"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   478
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldDivisor 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   6120
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   661
      _Version        =   327682
      Min             =   1
      Max             =   100
      SelStart        =   10
      TickFrequency   =   10
      Value           =   10
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   8
      Left            =   1500
      TabIndex        =   9
      Text            =   "0"
      Top             =   6720
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   7
      Left            =   780
      TabIndex        =   8
      Text            =   "0"
      Top             =   6720
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   6
      Left            =   60
      TabIndex        =   7
      Text            =   "0"
      Top             =   6720
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   5
      Left            =   1500
      TabIndex        =   6
      Text            =   "0"
      Top             =   6360
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   4
      Left            =   780
      TabIndex        =   5
      Text            =   "1"
      Top             =   6360
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   3
      Left            =   60
      TabIndex        =   4
      Text            =   "0"
      Top             =   6360
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   2
      Left            =   1500
      TabIndex        =   3
      Text            =   "0"
      Top             =   6000
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   1
      Left            =   780
      TabIndex        =   2
      Text            =   "0"
      Top             =   6000
      Width           =   675
   End
   Begin VB.TextBox txtMatrix 
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Text            =   "0"
      Top             =   6000
      Width           =   675
   End
   Begin VB.PictureBox picResult 
      BorderStyle     =   0  'None
      Height          =   5880
      Left            =   60
      ScaleHeight     =   392
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   522
      TabIndex        =   0
      Top             =   60
      Width           =   7830
   End
   Begin ComctlLib.Slider sldBias 
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   6540
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   661
      _Version        =   327682
      Min             =   -100
      Max             =   100
      TickFrequency   =   10
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //
' // Direct2D convolve effect demo by The trick
' //

Option Explicit

Private Const GENERIC_READ As Long = &H80000000

Private m_cFactory      As ID2D1Factory
Private m_cWICFactory   As WICImagingFactory
Private m_cRenderTarget As ID2D1DeviceContext
Private m_cEffect       As ID2D1Effect

Private Sub Form_Load()
    Dim cBitmap     As ID2D1Bitmap
    Dim cHwndTgt    As ID2D1HwndRenderTarget

    ' // Create a factory
    Set m_cFactory = D2D1.CreateFactory()

    ' // Create render target
    Set m_cRenderTarget = m_cFactory.CreateHwndRenderTarget(D2D1.RenderTargetProperties(D2D1.PixelFormat()), _
                                                            D2D1.HwndRenderTargetProperties(picResult.hWnd, D2D1.SizeU()))
    
    Set cHwndTgt = m_cRenderTarget
    
    cHwndTgt.Resize D2D1.SizeU(picResult.ScaleWidth, picResult.ScaleHeight)

    Set cBitmap = LoadImage(App.Path & "\image_01.jpg", m_cRenderTarget)

    ' // Create blur effect
    Set m_cEffect = m_cRenderTarget.CreateEffect(CLSID_D2D1ConvolveMatrix)

    m_cEffect.SetInput 0, cBitmap

    Update
    
End Sub

Private Function LoadImage( _
                 ByRef sPath As String, _
                 ByVal cRenderTarget As ID2D1RenderTarget) As ID2D1Bitmap
    Dim cDecoder        As IWICBitmapDecoder
    Dim cFrameDecode    As IWICBitmapFrameDecode
    Dim cConverter      As IWICFormatConverter
    
    If m_cWICFactory Is Nothing Then
        Set m_cWICFactory = New WICImagingFactory
    End If
    
    Set cDecoder = m_cWICFactory.CreateDecoderFromFilename(sPath, ByVal 0&, GENERIC_READ, WICDecodeMetadataCacheOnLoad)
    Set cFrameDecode = cDecoder.GetFrame(0)
    Set cConverter = m_cWICFactory.CreateFormatConverter()
    
    cConverter.Initialize cFrameDecode, GUID_WICPixelFormat32bppPBGRA, WICBitmapDitherTypeNone, Nothing, 0, WICBitmapPaletteTypeMedianCut
    
    Set LoadImage = cRenderTarget.CreateBitmapFromWicBitmap(ByVal cConverter, ByVal 0&)
    
End Function

Private Sub picResult_Paint()

    m_cRenderTarget.BeginDraw
    
    m_cRenderTarget.Clear D2D1.ColorF(Indigo)

    m_cRenderTarget.DrawImage m_cEffect, ByVal 0&, ByVal 0&
    
    m_cRenderTarget.EndDraw ByVal 0&, ByVal 0&
    
End Sub

Private Sub Update()
    Dim fMatrix(2, 2)   As Single
    Dim lIndex          As Long
    
    For lIndex = 0 To 8
        
        If IsNumeric(txtMatrix(lIndex)) Then
            fMatrix(lIndex Mod 3, lIndex \ 3) = CDbl(txtMatrix(lIndex))
        Else
            fMatrix(lIndex Mod 3, lIndex \ 3) = 0
        End If
        
    Next
    
    m_cEffect.SetValue D2D1_CONVOLVEMATRIX_PROP_KERNEL_MATRIX, D2D1_PROPERTY_TYPE_UNKNOWN, fMatrix(0, 0), _
                       LenB(fMatrix(0, 0)) * 9
    m_cEffect.SetValue D2D1_CONVOLVEMATRIX_PROP_DIVISOR, D2D1_PROPERTY_TYPE_FLOAT, CSng(sldDivisor.Value / 10), 4
    m_cEffect.SetValue D2D1_CONVOLVEMATRIX_PROP_BIAS, D2D1_PROPERTY_TYPE_FLOAT, CSng(sldBias.Value / 100), 4
    
    picResult_Paint
    
End Sub

Private Sub sldDivisor_Change()
    Update
End Sub
Private Sub sldDivisor_Scroll()
    Update
End Sub
Private Sub sldBias_Change()
    Update
End Sub
Private Sub sldBias_Scroll()
    Update
End Sub

Private Sub txtMatrix_Change( _
            ByRef iIndex As Integer)
    Update
End Sub
